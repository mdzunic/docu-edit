import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { extend, MOD, WORD_LEFT, WORD_RIGHT } from "./keys";

// SPEC-034 — print preview and client-side PDF export: Ctrl/Cmd+P opens an
// in-app preview over the paginated canvas (the canvas IS the preview),
// with a page-range control and a Print button that hands off to the
// browser's native dialog using the document's own page geometry; the
// Download menu exports a searchable PDF with embedded fonts, outline
// bookmarks, links and title metadata, built entirely client-side.

const content = (page: Page) => page.getByTestId("document-content");
const preview = (page: Page) => page.getByTestId("print-preview");
const documentPages = (page: Page) => page.getByTestId("document-page");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Counts window.print() calls without ever opening a real dialog. */
async function stubPrint(page: Page): Promise<void> {
  await page.addInitScript(() => {
    (window as unknown as { __printCalls: number }).__printCalls = 0;
    window.print = () => {
      (window as unknown as { __printCalls: number }).__printCalls += 1;
    };
  });
}

const printCalls = (page: Page) =>
  page.evaluate(() => (window as unknown as { __printCalls: number }).__printCalls);

/** Waits for pagination to settle (see e2e/SPEC-028.spec.ts). */
async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

/** Seeds many paragraphs in one operation through the app's own plain-text
 * paste path (one paragraph per line) — same helper as SPEC-028. */
async function pasteText(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.evaluate((payload) => {
    const root = document.querySelector('[data-testid="document-content"]')!;
    const data = new DataTransfer();
    data.setData("text/plain", payload);
    root.dispatchEvent(
      new ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data }),
    );
  }, text);
}

function paragraphLines(count: number, prefix = "Paragraph"): string {
  return Array.from({ length: count }, (_, i) => `${prefix} ${i} lorem ipsum dolor sit amet`).join(
    "\n",
  );
}

/** A minimal truecolour PNG, built the same way e2e/SPEC-024.spec.ts builds
 * its fixtures — the image-export test needs a real decodable bitmap. */
function buildPng(width: number, height: number, [r, g, b]: [number, number, number]): Buffer {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    return c >>> 0;
  });
  const crc32 = (buf: Buffer) => {
    let c = 0xffffffff;
    for (const byte of buf) {
      c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const length = Buffer.alloc(4);
    length.writeUInt32BE(data.length, 0);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body), 0);
    return Buffer.concat([length, body, crc]);
  };
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (rowBytes + 1);
    for (let x = 0; x < width; x++) {
      const px = rowStart + 1 + x * 3;
      raw[px] = r;
      raw[px + 1] = g;
      raw[px + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

/** Downloads via Download → PDF and returns the file bytes as latin-1 text
 * (binary-safe for substring assertions against PDF structure). */
async function downloadPdfText(page: Page): Promise<string> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-pdf").click(),
  ]);
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path!).toString("latin1");
}

test("Ctrl/Cmd+P opens the print preview over the paginated document (AC-1, AC-2)", async ({
  page,
}) => {
  await stubPrint(page);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello print world");
  await waitForPagination(page);

  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  await expect(page.getByTestId("print-page-count")).toHaveText("1 page");
  // The paginated canvas stays visible behind the panel — it is the preview.
  await expect(documentPages(page).first()).toBeVisible();
  await expect(content(page)).toContainText("Hello print world");
  expect(await printCalls(page)).toBe(0);
});

test("the preview hides screen-only affordances: ruler, formatting marks, margin guides (AC-13)", async ({
  page,
}) => {
  await stubPrint(page);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Marked up text");
  // Turn formatting marks on so the exclusion is observable.
  await page.keyboard.press(`${MOD}+Shift+p`);
  await expect(content(page)).toHaveClass(/marks-on/);
  await expect(page.getByTestId("document-ruler")).toBeVisible();

  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  // The ruler hides (the prior attempt's regression: its real root class is
  // `ruler`, now excluded via the shared print-exclude class)…
  await expect(page.getByTestId("document-ruler")).toBeHidden();
  // …and the marks class drops while previewing, without touching the
  // document's own setting.
  await expect(content(page)).not.toHaveClass(/marks-on/);

  await page.getByTestId("print-preview-close").click();
  await expect(preview(page)).toBeHidden();
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await expect(content(page)).toHaveClass(/marks-on/);
});

test("the page-range control validates input and limits the previewed pages (AC-4, AC-28, AC-31)", async ({
  page,
}) => {
  await stubPrint(page);
  await createDocument(page);
  await pasteText(page, paragraphLines(120));
  await waitForPagination(page);
  const totalPages = await documentPages(page).count();
  expect(totalPages).toBeGreaterThan(1);

  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  await expect(page.getByTestId("print-page-count")).toHaveText(`${totalPages} pages`);

  await page.getByTestId("print-range-custom").check();
  const input = page.getByTestId("print-range-input");

  await input.fill("abc");
  await expect(page.getByTestId("print-range-error")).toContainText("not a page number");
  await expect(page.getByTestId("print-confirm")).toBeDisabled();

  await input.fill("5-2");
  await expect(page.getByTestId("print-range-error")).toContainText("starts after it ends");
  await expect(page.getByTestId("print-confirm")).toBeDisabled();

  await input.fill("999");
  await expect(page.getByTestId("print-range-error")).toContainText(
    `Pages go from 1 to ${totalPages}`,
  );
  await expect(page.getByTestId("print-confirm")).toBeDisabled();
  expect(await printCalls(page)).toBe(0);

  // A valid range: only the selected page stays visible in the preview.
  await input.fill("1");
  await expect(page.getByTestId("print-range-error")).toBeHidden();
  await expect(page.getByTestId("print-confirm")).toBeEnabled();
  await expect(documentPages(page).locator("visible=true")).toHaveCount(1);

  // Back to all pages: every page shows again.
  await page.getByTestId("print-range-all").check();
  await expect(documentPages(page).locator("visible=true")).toHaveCount(totalPages);
});

test("Print hands off to the browser's native dialog with the document's @page geometry (AC-3, AC-5, AC-6)", async ({
  page,
}) => {
  await stubPrint(page);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Print me");
  await waitForPagination(page);

  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  await page.getByTestId("print-confirm").click();
  await expect.poll(() => printCalls(page)).toBe(1);

  // The injected @page rule uses the document's page size (A4 default) with
  // zero margin — the on-screen page boxes carry the margins themselves, so
  // printed pagination is the preview's pagination by construction.
  const pageRule = await page
    .locator('style[data-testid="print-page-style"]')
    .evaluate((el) => el.textContent ?? "");
  expect(pageRule).toBe("@page { size: 21cm 29.7cm; margin: 0; }");
});

test("under print media, chrome and screen affordances vanish while page content remains (AC-7, AC-13)", async ({
  page,
}) => {
  await stubPrint(page);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Printed body content");
  await waitForPagination(page);

  await page.emulateMedia({ media: "print" });
  await expect(page.getByTestId("document-ruler")).toBeHidden();
  await expect(page.getByTestId("app-bar")).toBeHidden();
  await expect(page.getByTestId("toolbar")).toBeHidden();
  await expect(page.getByTestId("status-bar")).toBeHidden();
  await expect(content(page)).toContainText("Printed body content");
  await page.emulateMedia({ media: "screen" });
  await expect(page.getByTestId("document-ruler")).toBeVisible();
});

test("Download → PDF exports searchable text with embedded fonts, outline, links and title (AC-14..20)", async ({
  page,
}) => {
  await createDocument(page);

  // Title (lands in the PDF metadata).
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("PDF Test Doc");
  await page.getByTestId("document-title-input").press("Enter");

  // A Heading 1 (outline entry), body text, and an external link.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Alt+1`);
  await page.keyboard.type("Chapter One");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Searchable body text with a link here");
  // Select the word "link": Ctrl/Cmd+Alt+K (SPEC-062 moved the link
  // accelerator off Ctrl/Cmd+K) applies to the selection.
  await page.keyboard.press(WORD_LEFT);
  await page.keyboard.press(WORD_LEFT);
  await page.keyboard.press(extend(WORD_RIGHT));
  await page.keyboard.press(`${MOD}+Alt+k`);
  await page.getByTestId("link-url-input").fill("https://example.com/spec034");
  await page.getByTestId("link-save").click();
  await waitForPagination(page);

  const pdf = await downloadPdfText(page);
  // A real PDF, not a raster: header plus text-showing operators (AC-15/16).
  expect(pdf.startsWith("%PDF-1.7\n")).toBe(true);
  expect(pdf).toContain(" Tj ET");
  expect(pdf).toContain("/ToUnicode");
  // Embedded (not referenced) TrueType programs (AC-20).
  expect(pdf).toContain("/Subtype /CIDFontType2");
  expect(pdf).toContain("/FontFile2");
  expect(pdf).toContain("/BaseFont /Liberation");
  // Title metadata (AC-17).
  expect(pdf).toContain("/Title (PDF Test Doc)");
  // Outline from the Heading 1 (AC-18).
  expect(pdf).toContain("/Type /Outlines");
  expect(pdf).toContain("/PageMode /UseOutlines");
  expect(pdf).toContain("/Title (Chapter One)");
  // The hyperlink as a clickable URI annotation (AC-19).
  expect(pdf).toContain("/Subtype /Link");
  expect(pdf).toContain("/S /URI /URI (https://example.com/spec034)");
});

test("an inserted image is embedded in the exported PDF and drawn where it sits on the page", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Text beside the picture");
  await page.getByTestId("insert-image-file-input").setInputFiles({
    name: "photo.png",
    mimeType: "image/png",
    buffer: buildPng(200, 100, [0x33, 0x77, 0xbb]),
  });
  await expect(content(page).locator('[data-testid="doc-image"]')).toHaveCount(1);
  await waitForPagination(page);

  const pdf = await downloadPdfText(page);
  // The image rides along as a real image XObject, sampled and Flate'd from
  // the rendered bitmap, and is drawn by name in the page's content stream.
  expect(pdf).toContain("/Type /XObject /Subtype /Image");
  expect(pdf).toContain("/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode");
  expect(pdf).toContain("/XObject << /Im1");
  expect(pdf).toMatch(/q [\d.]+ 0 0 [\d.]+ [\d.]+ [\d.]+ cm \/Im1 Do Q/);
  // The text still exports as text — the image does not turn the page into
  // a raster (AC-15).
  expect(pdf).toContain(" Tj ET");

  // The draw box matches the image's aspect ratio (200×100 → 2:1) rather
  // than being stretched to some default.
  const draw = /q ([\d.]+) 0 0 ([\d.]+) [\d.]+ [\d.]+ cm \/Im1 Do Q/.exec(pdf);
  expect(draw).not.toBeNull();
  const [width, height] = [Number(draw![1]), Number(draw![2])];
  expect(width / height).toBeCloseTo(2, 1);
});

test("headers print and export in position (AC-8, AC-10)", async ({ page }) => {
  await stubPrint(page);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body under a header");
  await waitForPagination(page);

  // Enter the header band and give it content (SPEC-029's editing flow).
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await expect(page.getByTestId("page-header-band").first()).toHaveAttribute(
    "data-hf-editable",
    "1",
  );
  await page.keyboard.type("Report header");
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("page-header-band").first()).toContainText("Report header");

  // Visible in the preview…
  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  await expect(page.getByTestId("page-header-band").first()).toContainText("Report header");
  await page.getByTestId("print-preview-close").click();

  // …and extracted into the PDF's text (the header band is page content).
  const pdf = await downloadPdfText(page);
  expect(pdf).toContain(" Tj ET");
});

test("footers sit at the bottom of the page, above any footnotes (AC-9, AC-11)", async ({ page }) => {
  await stubPrint(page);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body above a footer and a footnote");
  await waitForPagination(page);

  // Give the page a footer.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-footer").click();
  await expect(page.getByTestId("page-footer-band").first()).toHaveAttribute(
    "data-hf-editable",
    "1",
  );
  await page.keyboard.type("Report footer");
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("page-footer-band").first()).toContainText("Report footer");

  // A footnote reference at the end of the body, with a note body.
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("notes-menu-footnote").click();
  await page.keyboard.type("Footnote body text");
  await page.keyboard.press("Escape");
  await waitForPagination(page);

  const pageBox = await documentPages(page).first().boundingBox();
  const footerBox = await page.getByTestId("page-footer-band").first().boundingBox();
  const noteBox = await page.getByTestId("footnote-area").first().boundingBox();
  expect(pageBox).toBeTruthy();
  expect(footerBox).toBeTruthy();
  expect(noteBox).toBeTruthy();
  // The footer sits in the bottom 30% of the page (AC-9)…
  expect(footerBox!.y + footerBox!.height - pageBox!.y).toBeGreaterThan(pageBox!.height * 0.7);
  // …and the footnote area sits above it, itself in the lower part of the
  // page (AC-11), separated from the body by the note separator rule.
  expect(noteBox!.y + noteBox!.height).toBeLessThanOrEqual(footerBox!.y + 1);
  expect(noteBox!.y - pageBox!.y).toBeGreaterThan(pageBox!.height * 0.5);
  await expect(page.getByTestId("note-separator").first()).toBeVisible();
  await expect(page.getByTestId("footnote-area").first()).toContainText("Footnote body text");

  // Both remain in place under the print preview.
  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  await expect(page.getByTestId("page-footer-band").first()).toContainText("Report footer");
  await expect(page.getByTestId("footnote-area").first()).toContainText("Footnote body text");
});

test("multi-column layout flows and prints columns in order (AC-12)", async ({ page }) => {
  await createDocument(page);
  // Enough paragraphs that the first page's left column fills and content
  // spills into the right column — short text alone would fit in one
  // column and never exercise the column break.
  await pasteText(page, paragraphLines(90, "Column"));
  await waitForPagination(page);

  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("columns-button").click();
  await page.getByTestId("columns-preset-2").click();
  await page.getByTestId("columns-confirm").click();
  await waitForPagination(page);

  const columns = await page.getByTestId("doc-column").all();
  expect(columns.length).toBeGreaterThanOrEqual(2);
  const firstBox = await columns[0].boundingBox();
  const secondBox = await columns[1].boundingBox();
  expect(firstBox).toBeTruthy();
  expect(secondBox).toBeTruthy();
  // Left column renders left of the right column on the same page (AC-12)…
  expect(firstBox!.x).toBeLessThan(secondBox!.x);
  // …and content flows into the left column first, continuing into the
  // right column in document order.
  const firstText = (await columns[0].textContent()) ?? "";
  const secondText = (await columns[1].textContent()) ?? "";
  expect(firstText).toContain("Column 0 ");
  expect(firstText).not.toContain("Column 89 ");
  expect(secondText).not.toContain("Column 0 ");

  // Column order survives into the PDF export as ordinary page text.
  const pdf = await downloadPdfText(page);
  expect(pdf).toContain(" Tj ET");
});

test("a large export shows progress, keeps the UI responsive and completes well under budget (AC-21, AC-22, AC-23)", async ({
  page,
}) => {
  await createDocument(page);
  await pasteText(page, paragraphLines(600, "Performance paragraph"));
  await waitForPagination(page, 60_000);
  const totalPages = await documentPages(page).count();
  expect(totalPages).toBeGreaterThan(10);

  const started = Date.now();
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download", { timeout: 30_000 });
  await page.getByTestId("download-pdf").click();
  // The progress indicator appears while pages are extracted (AC-22)…
  await expect(page.getByTestId("pdf-export-progress")).toBeVisible();
  // …and the page stays interactive: a DOM click round-trips mid-export
  // (AC-23 — the exporter yields to the event loop between pages).
  await page.getByTestId("document-title").click();
  await page.keyboard.press("Escape");
  const download = await downloadPromise;
  const elapsed = Date.now() - started;
  expect(elapsed).toBeLessThan(30_000);
  const path = await download.path();
  const pdf = readFileSync(path!).toString("latin1");
  expect(pdf).toContain(`/Count ${totalPages}`);
  await expect(page.getByTestId("pdf-export-progress")).toBeHidden();
});

test("an empty document previews as no pages and exports a zero-page PDF with empty title metadata (AC-27, AC-32)", async ({
  page,
}) => {
  await stubPrint(page);
  await createDocument(page);
  await waitForPagination(page);

  await page.keyboard.press(`${MOD}+p`);
  await expect(preview(page)).toBeVisible();
  await expect(page.getByTestId("print-page-count")).toHaveText("0 pages");
  await expect(page.getByTestId("print-no-pages")).toBeVisible();
  await expect(page.getByTestId("print-confirm")).toBeDisabled();
  // Every laid-out page box hides: the preview shows no pages.
  await expect(documentPages(page).locator("visible=true")).toHaveCount(0);
  await page.getByTestId("print-preview-close").click();

  const pdf = await downloadPdfText(page);
  expect(pdf).toContain("/Kids [] /Count 0");
  // An untitled document exports empty title metadata (AC-32).
  expect(pdf).toContain("/Title ()");
});

test("a failed export shows \"Couldn't create the PDF\" with a reason and leaves the document untouched (AC-24, AC-25, AC-26)", async ({
  page,
}) => {
  // Block the bundled fonts from the very first request so embedding fails.
  await page.route("**/fonts/*.ttf", (route) => route.abort());
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Failure test content");
  await waitForPagination(page);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  await page.getByTestId("download-pdf").click();
  const toast = page.getByTestId("toast").first();
  await expect(toast).toBeVisible({ timeout: 15_000 });
  await expect(toast).toContainText("Couldn't create the PDF");
  await expect(toast).toContainText("no fonts could be loaded for embedding");

  // The document is unchanged and still editable (AC-26).
  await expect(content(page)).toContainText("Failure test content");
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type(" and still editable");
  await expect(content(page)).toContainText("Failure test content and still editable");
});
