import { test, expect, type Download, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-107 — images and PDFs in the library: the full AC-16 journey. Drop a
// PDF → it opens in the media viewer; its library row carries a PDF badge
// and its grid card an honest 3-page placeholder; the details panel is built
// from the structural parse; a dropped PNG's card shows a REAL thumbnail;
// the image viewer zooms with SPEC-063's vocabulary; bulk download returns
// the PDF byte-for-byte; trash/restore and the kind-grouped filter chips
// treat media as first-class.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const pdfFixture = join(fixturesDir, "spec107-pages.pdf");
const pngFixture = join(fixturesDir, "spec107-shot.png");

async function dataTransferWithFile(page: Page, name: string, mime: string, path: string) {
  return page.evaluateHandle(
    (item) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(item.base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], item.name, { type: item.mime }));
      return dt;
    },
    { name, mime, base64: readFileSync(path).toString("base64") },
  );
}

async function waitUntilAppReady(page: Page) {
  await expect(page.getByTestId("app-bar")).toBeVisible();
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

test("media documents are first-class: drop, view, thumbnail, zoom, bulk download, trash, filter (AC-16)", async ({ page }) => {
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
  await expect(page.locator('object[type="application/pdf"]')).toBeAttached();
  await expect(page.locator("embed")).toBeAttached();

  // Close (SPEC-060) → the library shows a row with a PDF badge.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const pdfRow = rowByTitle(page, "spec107-pages");
  await expect(pdfRow.getByTestId("row-format")).toHaveText("PDF");

  // The grid card is an honest paper-shaped placeholder: page count on it,
  // no rasterised preview.
  await page.getByTestId("view-grid").click();
  const placeholder = page.getByTestId("pdf-placeholder");
  await expect(placeholder).toBeVisible();
  await expect(placeholder).toContainText("3 pages");

  // Open it from the card, then close again.
  await rowByTitle(page, "spec107-pages").click();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-pages")).toHaveText("3 pages");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // Drop the PNG; its card shows a REAL downscaled thumbnail of the image.
  const pngTransfer = await dataTransferWithFile(page, "spec107-shot.png", "image/png", pngFixture);
  await page.dispatchEvent("body", "drop", { dataTransfer: pngTransfer });
  await expect(page).toHaveURL(/\/media\/.+/);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const thumbnail = page.getByTestId("media-thumbnail");
  await expect(thumbnail).toBeVisible();
  await expect(thumbnail).toHaveAttribute("src", /^data:image\//);

  // Open the image: it starts fit-to-window, zooms to 100% and back to fit
  // (SPEC-063 vocabulary).
  await rowByTitle(page, "spec107-shot").click();
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
  const downloadedBytes = readFileSync(downloadedPath!);
  expect(Buffer.compare(downloadedBytes, readFileSync(pdfFixture))).toBe(0);

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
