import { test, expect, type Download, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { assembleClassicPdf } from "../src/lib/media/fixtures/build";

// SPEC-107 — images and PDFs in the library: the full journey (AC-41..43).
// Drop a PDF (spec107-pages.pdf, the app's own buildPdf output) → it opens
// in the media viewer; its library row carries a PDF badge and its grid card
// reads 3 pages; the details panel is built from the structural parse; a
// dropped PNG's card shows a REAL thumbnail; the image viewer zooms with
// SPEC-063's vocabulary; bulk download returns the PDF byte-for-byte;
// trash/restore and the kind-grouped filter chips treat media as
// first-class.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const pdfFixture = join(fixturesDir, "spec107-pages.pdf");
const pngFixture = join(fixturesDir, "spec107-shot.png");

async function dataTransferWithBytes(page: Page, name: string, mime: string, bytes: Uint8Array) {
  return page.evaluateHandle(
    (item) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(item.base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], item.name, { type: item.mime }));
      return dt;
    },
    { name, mime, base64: Buffer.from(bytes).toString("base64") },
  );
}

async function dataTransferWithFile(page: Page, name: string, mime: string, path: string) {
  return dataTransferWithBytes(page, name, mime, readFileSync(path));
}

/** A one-page PDF of the given size, written by the same classic-xref
 * assembler the media fixtures use. */
function onePagePdf(widthPt: number, heightPt: number): Uint8Array {
  return assembleClassicPdf(
    "1.7",
    [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${widthPt} ${heightPt}] >>`,
    ],
    "/Root 1 0 R",
  );
}

/** The PDF placeholder's RENDERED paper, measured rather than read off its
 * style: the first page's real aspect ratio, inside the thumbnail, and sized
 * from it — the paper spans the thumbnail along whichever axis binds, so its
 * labels cannot be what sizes it. (An inline aspect-ratio once read right
 * while the labels' widths stretched the paper off it.) */
async function expectPaperAtRatio(thumbnail: Locator, ratio: number) {
  const paper = thumbnail.getByTestId("pdf-paper");
  await expect(paper).toBeVisible();
  const outer = (await thumbnail.boundingBox())!;
  const box = (await paper.boundingBox())!;
  expect(box.width / box.height).toBeCloseTo(ratio, 2);
  expect(box.x).toBeGreaterThanOrEqual(outer.x);
  expect(box.y).toBeGreaterThanOrEqual(outer.y);
  expect(box.x + box.width).toBeLessThanOrEqual(outer.x + outer.width);
  expect(box.y + box.height).toBeLessThanOrEqual(outer.y + outer.height);
  expect(Math.max(box.width / outer.width, box.height / outer.height)).toBeGreaterThan(0.75);
}

async function waitUntilAppReady(page: Page) {
  await expect(page.getByTestId("app-bar")).toBeVisible();
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

test("media documents are first-class: drop, view, thumbnail, zoom, bulk download, trash, filter (AC-41..43)", async ({ page }) => {
  await page.goto("/");
  await waitUntilAppReady(page);

  // Drop the 3-page PDF: every open path funnels through openAnyDocumentFile,
  // so the drop opens the media viewer at /media/:id.
  const pdfTransfer = await dataTransferWithFile(page, "spec107-pages.pdf", "application/pdf", pdfFixture);
  await page.dispatchEvent("body", "drop", { dataTransfer: pdfTransfer });
  await expect(page).toHaveURL(/\/media\/.+/);

  // The details panel reads 3 pages (from parsePdf, so it is there whether
  // or not the browser embed renders) and the embed element is present.
  await expect(page.getByTestId("detail-pages")).toHaveText("3 pages");
  await expect(page.getByTestId("detail-page-size")).toContainText("mixed");
  await expect(page.getByTestId("detail-version")).toHaveText("1.7");
  await expect(page.getByTestId("detail-info-author")).toHaveText("docu-edit factory");
  await expect(page.getByTestId("detail-reconstructed")).toHaveText("No");
  await expect(page.locator('object[type="application/pdf"]')).toBeAttached();
  await expect(page.locator("embed")).toBeAttached();

  // Close (SPEC-060) → the library shows a row with a PDF badge.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const pdfRow = rowByTitle(page, "spec107-pages");
  await expect(pdfRow.getByTestId("row-format")).toHaveText("PDF");

  // A 3-page card. Its preview is a real raster of the first page (pdf.js —
  // the dependency exception CLAUDE.md records for PDF first-page previews),
  // replacing the aspect-true paper placeholder once it is drawn.
  await page.getByTestId("view-grid").click();
  await expect(rowByTitle(page, "spec107-pages").getByTestId("row-counts")).toContainText("3 pages");
  const pdfThumb = page.getByTestId("media-thumbnail").first();
  await expect(pdfThumb).toBeVisible();
  expect(await pdfThumb.getAttribute("src")).toMatch(/^data:image\/png;base64,/);
  await expect(page.getByTestId("pdf-placeholder")).toHaveCount(0);

  // Open it from the card — a plain click previews in the quick-look pane,
  // a double-click opens — then close again.
  await rowByTitle(page, "spec107-pages").dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-pages")).toHaveText("3 pages");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // Drop the PNG; its card shows a REAL downscaled thumbnail of the image.
  // Scoped to the PNG's own card: the PDF card next to it re-renders its
  // cached raster under the same testid, so an unscoped media-thumbnail
  // locator matches whichever raster lands first — and trips strict mode
  // the moment both are up.
  const pngTransfer = await dataTransferWithFile(page, "spec107-shot.png", "image/png", pngFixture);
  await page.dispatchEvent("body", "drop", { dataTransfer: pngTransfer });
  await expect(page).toHaveURL(/\/media\/.+/);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const thumbnail = rowByTitle(page, "spec107-shot").getByTestId("media-thumbnail");
  await expect(thumbnail).toBeVisible();
  await expect(thumbnail).toHaveAttribute("src", /^data:image\//);

  // Open the image (double-click — the grid's plain click previews): it
  // starts fit-to-window, zooms to 100% and back to fit (SPEC-063 vocabulary).
  await rowByTitle(page, "spec107-shot").dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("media-image")).toBeVisible();
  await expect(page.getByTestId("media-zoom-fit")).toHaveAttribute("aria-pressed", "true");
  await page.getByTestId("media-zoom-actual").click();
  await expect(page.getByTestId("media-zoom-level")).toHaveText("100%");
  await expect(page.getByTestId("media-zoom-fit")).toHaveAttribute("aria-pressed", "false");
  await page.getByTestId("media-zoom-fit").click();
  await expect(page.getByTestId("media-zoom-fit")).toHaveAttribute("aria-pressed", "true");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // Select both and bulk-download: the PDF comes back byte-for-byte.
  const downloads: Download[] = [];
  page.on("download", (download) => downloads.push(download));
  await page.getByTestId("row-select").nth(0).check();
  await page.getByTestId("row-select").nth(1).check();
  await expect(page.getByTestId("selection-count")).toContainText("2");
  await page.getByTestId("bulk-download").click();
  await expect.poll(() => downloads.length, { timeout: 15_000 }).toBe(2);
  const pdfDownload = downloads.find((d) => d.suggestedFilename() === "spec107-pages.pdf");
  expect(pdfDownload).toBeTruthy();
  const downloadedPath = await pdfDownload!.path();
  expect(downloadedPath).toBeTruthy();
  const downloadedBytes = new Uint8Array(readFileSync(downloadedPath!));
  expect(equalBytes(downloadedBytes, new Uint8Array(readFileSync(pdfFixture)))).toBe(true);

  // Trash both (the selection survives the download), restore both.
  await page.getByTestId("bulk-trash").click();
  await expect(page.getByTestId("library-empty")).toBeVisible();
  await page.getByTestId("trash-link").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);
  await page.getByTestId("row-restore").first().click();
  await page.getByTestId("row-restore").click();
  await expect(page.getByTestId("trash-empty")).toBeVisible();
  await page.getByTestId("back-to-library").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);

  // The kind-grouped chips: "Images" shows the PNG row, "PDFs" the PDF row.
  await page.getByTestId("filter-images").click();
  await expect(page.getByTestId("row-title")).toHaveCount(1);
  await expect(page.getByTestId("row-title")).toHaveText("spec107-shot");
  await page.getByTestId("filter-images").click();
  await page.getByTestId("filter-pdf").click();
  await expect(page.getByTestId("row-title")).toHaveCount(1);
  await expect(page.getByTestId("row-title")).toHaveText("spec107-pages");
});

test("with no raster, a PDF card is paper at its first page's real aspect ratio (AC-29)", async ({ page }) => {
  // The paper placeholder shows wherever no pdf.js raster is available. A
  // host without Workers is one such place (pdfRasterSupported() is false),
  // so there the placeholder is the card's final state and can be measured.
  await page.addInitScript(() => {
    delete (window as { Worker?: unknown }).Worker;
  });
  await page.goto("/");
  await waitUntilAppReady(page);

  // Portrait (the fixture's US Letter first page), A4 landscape, and a strip
  // narrower than the labels it carries.
  const pdfs = [
    { title: "spec107-pages", bytes: readFileSync(pdfFixture), ratio: 612 / 792, pages: "3 pages", size: "612 × 792 pt" },
    { title: "spec107-landscape", bytes: onePagePdf(842, 595), ratio: 842 / 595, pages: "1 page", size: "842 × 595 pt" },
    { title: "spec107-strip", bytes: onePagePdf(300, 900), ratio: 300 / 900, pages: "1 page", size: "300 × 900 pt" },
  ];
  for (const pdf of pdfs) {
    const transfer = await dataTransferWithBytes(page, `${pdf.title}.pdf`, "application/pdf", pdf.bytes);
    await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
    await expect(page).toHaveURL(/\/media\/.+/);
    await page.getByTestId("close-document").click();
    await expect(page.getByTestId("library")).toBeVisible();
  }

  await page.getByTestId("view-grid").click();
  for (const pdf of pdfs) {
    const thumbnail = rowByTitle(page, pdf.title).getByTestId("row-thumbnail");
    await expect(thumbnail.getByTestId("pdf-placeholder")).toContainText(pdf.pages);
    await expect(thumbnail.getByTestId("pdf-placeholder")).toContainText(pdf.size);
    await expectPaperAtRatio(thumbnail, pdf.ratio);
  }
  await expect(page.getByTestId("media-thumbnail")).toHaveCount(0);

  // The quick-look pane sizes the same thumbnail differently; the paper keeps
  // the page's ratio there too.
  await rowByTitle(page, "spec107-strip").click();
  await expectPaperAtRatio(page.getByTestId("quick-look").getByTestId("row-thumbnail"), 300 / 900);
});
