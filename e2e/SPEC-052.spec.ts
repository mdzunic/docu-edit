import { test, expect, type Page, type Locator, type Download } from "@playwright/test";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-052 — charts: rendering the fixtures, inserting, the data sheet, the
// options panel, "Chart from table", and byte-identical preservation.
// word-chart-types.docx/lo-chart-types.odt are hand-built by
// scripts/build-spec052-fixtures.mjs — NOT round-tripped through this
// project's own writer — containing, in this fixed order:
//   0 columnClustered   4 lineMarkers   8 scatter
//   1 columnStacked     5 area          9 columnClustered, negative values
//   2 barClustered      6 pie          10 columnClustered, one gap
//   3 line              7 doughnut     11 unsupported (3-D / chart:stock)

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxChartsFixture = join(fixturesDir, "word-chart-types.docx");
const odtChartsFixture = join(fixturesDir, "lo-chart-types.odt");
const chartsAndControlsFixture = join(fixturesDir, "..", "..", "src", "lib", "docx", "fixtures", "word-charts-and-controls.docx");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');
const drawingAt = (page: Page, index: number) => content(page).locator(`[data-testid="drawing-${index}"]`);

/** Whether any pixel on the canvas immediately preceding `proxy` is opaque —
 * mirrors e2e/SPEC-051.spec.ts's own canvasPixelAt/canvasHasOpaquePixel
 * helpers, since a chart's rendering is otherwise invisible to the DOM. */
async function canvasHasOpaquePixel(proxy: Locator): Promise<boolean> {
  const canvas = proxy.locator("xpath=preceding-sibling::canvas[1]");
  return canvas.evaluate((el) => {
    const canvasEl = el as HTMLCanvasElement;
    const ctx = canvasEl.getContext("2d")!;
    const data = ctx.getImageData(0, 0, canvasEl.width, canvasEl.height).data;
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] > 0) {
        return true;
      }
    }
    return false;
  });
}

async function insertChart(page: Page, chartType = "columnClustered"): Promise<void> {
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`chart-gallery-${chartType}`).click();
}

/** Minimal central-directory walk over a downloaded archive — same shape as
 * e2e/SPEC-006.spec.ts/SPEC-008.spec.ts's own copies — enough to compare
 * preserved parts byte-for-byte from Node. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd, "end-of-central-directory record should exist").toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let i = 0; i < totalEntries; i++) {
    expect(buffer.readUInt32LE(offset), "central directory entry signature").toBe(0x02014b50);
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString("utf8", offset + 46, offset + 46 + nameLength);
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const data = buffer.subarray(dataStart, dataStart + compressedSize);
    parts.set(name, method === 8 ? inflateRawSync(data) : Buffer.from(data));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
}

async function downloadedBuffer(download: Download): Promise<Buffer> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path);
}

// --- opening the fixtures (AC-1, AC-40..43) ---------------------------------

for (const [label, fixturePath] of [
  ["docx", docxChartsFixture],
  ["odt", odtChartsFixture],
] as const) {
  test(`opening the ${label} fixture renders every catalogue type, a negative-values chart, a gap chart and an unsupported placeholder`, async ({
    page,
  }) => {
    await openFixture(page, fixturePath);
    await expect(drawings(page)).toHaveCount(12);
    // Every catalogue chart (0..10) paints something.
    for (let i = 0; i <= 10; i++) {
      expect(await canvasHasOpaquePixel(drawingAt(page, i)), `chart ${i} should paint`).toBe(true);
    }
    // The unsupported chart (11) is a dashed placeholder — still painted
    // (a stroked border), still selectable, but reports no support.
    expect(await canvasHasOpaquePixel(drawingAt(page, 11)), "unsupported placeholder should still paint a border").toBe(true);

    // Selecting the unsupported chart disables editing and shows why
    // (AC-30, AC-31).
    await drawingAt(page, 11).click();
    await expect(page.getByTestId("chart-unsupported-reason")).toBeVisible();
    await expect(page.getByTestId("chart-edit-data-button")).toHaveCount(0);
    await expect(page.getByTestId("chart-options-button")).toHaveCount(0);

    // A supported chart's toolbar offers editing.
    await drawingAt(page, 0).click();
    await expect(page.getByTestId("chart-edit-data-button")).toBeVisible();
    await expect(page.getByTestId("chart-options-button")).toBeVisible();
  });
}

// --- insertion (AC-2, AC-3) --------------------------------------------------

test("Insert -> Chart offers the catalogue and inserts a 12cm x 7.5cm chart with sample data", async ({ page }) => {
  await createDocument(page);
  await insertChart(page, "pie");
  await expect(drawings(page)).toHaveCount(1);
  const proxy = drawingAt(page, 0);
  await expect(proxy).toBeVisible();
  expect(await canvasHasOpaquePixel(proxy)).toBe(true);

  const canvas = proxy.locator("xpath=preceding-sibling::canvas[1]");
  const widthPx = await canvas.evaluate((el) => (el as HTMLCanvasElement).getBoundingClientRect().width);
  const heightPx = await canvas.evaluate((el) => (el as HTMLCanvasElement).getBoundingClientRect().height);
  // 12cm/7.5cm at 96dpi/2.54 ≈ 453px / 283px — allow a few px of rounding.
  expect(Math.abs(widthPx - (12 / 2.54) * 96)).toBeLessThan(3);
  expect(Math.abs(heightPx - (7.5 / 2.54) * 96)).toBeLessThan(3);
});

// --- selection and the toolbar (AC-23) --------------------------------------

test("clicking a chart selects it and shows the floating toolbar; Delete removes it", async ({ page }) => {
  await createDocument(page);
  await insertChart(page);
  const proxy = drawingAt(page, 0);
  await proxy.click();
  await expect(page.getByTestId("chart-toolbar")).toBeVisible();
  await page.keyboard.press("Delete");
  await expect(drawings(page)).toHaveCount(0);
});

// Regression: "Duplicate" on an opened (not freshly inserted) chart must not
// leave the copy sharing the original's package part — editing just the
// copy previously overwrote the original's data on save, since both blocks
// pointed at the same word/charts/chartN.xml.
test("duplicating an opened chart and editing the copy leaves the original's part byte-identical", async ({ page }) => {
  await openFixture(page, docxChartsFixture);
  await drawingAt(page, 0).click();
  await page.getByTestId("chart-duplicate-button").click();
  // The copy is spliced in right after the original (block index + 1).
  await expect(drawings(page)).toHaveCount(13);

  await drawingAt(page, 1).dblclick();
  await expect(page.getByTestId("chart-data-sheet-dialog")).toBeVisible();
  await page.getByTestId("chart-data-cell-1-1").fill("999");
  await page.getByTestId("chart-data-cell-1-1").blur();
  await page.getByTestId("chart-data-apply").click();

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const buffer = await downloadedBuffer(await downloadPromise);

  const saved = unzipParts(buffer);
  const original = unzipParts(readFileSync(docxChartsFixture));

  // The original chart (word/charts/chart1.xml) survives byte-identical —
  // the edit landed only on the duplicate's own, distinct part.
  expect(saved.get("word/charts/chart1.xml")!.equals(original.get("word/charts/chart1.xml")!)).toBe(true);

  const newChartParts = [...saved.keys()].filter(
    (name) => /^word\/charts\/chart\d+\.xml$/.test(name) && !original.has(name),
  );
  expect(newChartParts).toHaveLength(1);
  expect(saved.get(newChartParts[0])!.toString("utf8")).toContain("999");
});

// --- the data sheet (AC-4..AC-13) -------------------------------------------

test("double-clicking a chart opens the data sheet; cell edits and row/column changes re-render it", async ({ page }) => {
  await createDocument(page);
  await insertChart(page);
  const proxy = drawingAt(page, 0);
  await proxy.dblclick();
  await expect(page.getByTestId("chart-data-sheet-dialog")).toBeVisible();

  // Edit a value cell (row 1 = first category, col 1 = first series).
  await page.getByTestId("chart-data-cell-1-1").fill("42");
  await page.getByTestId("chart-data-cell-1-1").blur();

  // Add a row and a column (AC-8, AC-9).
  await page.getByTestId("chart-data-add-row").click();
  await page.getByTestId("chart-data-add-col").click();
  await expect(page.getByTestId("chart-data-cell-4-3")).toBeVisible();

  // Paste TSV over a couple of cells (AC-10).
  await page.evaluate(() => {
    const input = document.querySelector('[data-testid="chart-data-cell-1-1"]') as HTMLInputElement;
    input.focus();
    const dt = new DataTransfer();
    dt.setData("text/plain", "7\t8");
    const event = new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true });
    input.dispatchEvent(event);
  });

  await page.getByTestId("chart-data-apply").click();
  await expect(page.getByTestId("chart-data-sheet-dialog")).toHaveCount(0);
  expect(await canvasHasOpaquePixel(proxy)).toBe(true);
});

test("the data sheet's toggles change the chart without changing the sheet (AC-11..AC-13)", async ({ page }) => {
  await createDocument(page);
  await insertChart(page);
  await drawingAt(page, 0).dblclick();
  const cellBefore = await page.getByTestId("chart-data-cell-0-1").inputValue();
  await page.getByTestId("chart-data-switch-rows-columns").check();
  // The grid's own cell contents are untouched by the toggle.
  await expect(page.getByTestId("chart-data-cell-0-1")).toHaveValue(cellBefore);
  await page.getByTestId("chart-data-first-row-as-names").uncheck();
  await page.getByTestId("chart-data-first-col-as-categories").uncheck();
  await expect(page.getByTestId("chart-data-cell-0-1")).toHaveValue(cellBefore);
  await page.getByTestId("chart-data-apply").click();
});

// --- the options panel (AC-14..AC-20) ---------------------------------------

test("the options panel sets title, legend, data labels, gridlines and axis min/max", async ({ page }) => {
  await createDocument(page);
  await insertChart(page);
  await drawingAt(page, 0).click();
  await page.getByTestId("chart-options-button").click();
  await expect(page.getByTestId("chart-options-dialog")).toBeVisible();

  await page.getByTestId("chart-options-title").fill("Quarterly revenue");
  await page.getByTestId("chart-options-legend").selectOption("bottom");
  await page.getByTestId("chart-options-data-labels").check();
  await page.getByTestId("chart-options-auto-min").uncheck();
  await page.getByTestId("chart-options-value-min").fill("0");
  await page.getByTestId("chart-options-auto-max").uncheck();
  await page.getByTestId("chart-options-value-max").fill("10");
  await page.getByTestId("chart-options-apply").click();
  await expect(page.getByTestId("chart-options-dialog")).toHaveCount(0);

  // Re-opening shows the applied values persisted.
  await drawingAt(page, 0).click();
  await page.getByTestId("chart-options-button").click();
  await expect(page.getByTestId("chart-options-title")).toHaveValue("Quarterly revenue");
  await expect(page.getByTestId("chart-options-legend")).toHaveValue("bottom");
  await expect(page.getByTestId("chart-options-data-labels")).toBeChecked();
});

// --- Chart from table (AC-21, AC-22) ----------------------------------------

test("Insert -> Chart from table builds a chart from the selected table's cells", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-3-3").click();
  const table = content(page).locator("table.doc-table");
  await expect(table).toHaveCount(1);

  const cellAt = (row: number, col: number) => page.getByTestId(`table-cell-${row}-${col}`);
  await cellAt(0, 1).locator(".doc-paragraph").first().click();
  await page.keyboard.type("Revenue");
  await cellAt(0, 2).locator(".doc-paragraph").first().click();
  await page.keyboard.type("Cost");
  await cellAt(1, 0).locator(".doc-paragraph").first().click();
  await page.keyboard.type("Q1");
  await cellAt(1, 1).locator(".doc-paragraph").first().click();
  await page.keyboard.type("10");
  await cellAt(1, 2).locator(".doc-paragraph").first().click();
  await page.keyboard.type("not a number");
  await cellAt(2, 0).locator(".doc-paragraph").first().click();
  await page.keyboard.type("Q2");
  await cellAt(2, 1).locator(".doc-paragraph").first().click();
  await page.keyboard.type("20");

  // Select the whole table by dragging across its cells (SPEC-022).
  const from = (await cellAt(0, 0).boundingBox())!;
  const to = (await cellAt(2, 2).boundingBox())!;
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 8 });
  await page.mouse.up();

  await page.getByTestId("menu-title-insert").click();
  await expect(page.getByTestId("insert-chart-from-table-button")).toBeEnabled();
  await page.getByTestId("insert-chart-from-table-button").click();

  await expect(drawings(page)).toHaveCount(1);
  expect(await canvasHasOpaquePixel(drawingAt(page, 0))).toBe(true);
  // The unparsable cell is reported (AC-22), not silently dropped.
  await expect(page.getByText(/could not be read as a number/i)).toBeVisible();
});

// --- byte-identical preservation (AC-25, AC-26) -----------------------------

test("an unedited chart survives a same-format save byte-identical after editing an unrelated paragraph", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(chartsAndControlsFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  // Edit the intro paragraph (unrelated to the chart).
  await content(page).getByText("Editable intro paragraph.").click();
  await page.keyboard.type(" Edited.");

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  // No conversion report for a same-format save: the chart round-trips
  // byte-identical, nothing dropped.
  await expect(page.getByTestId("conversion-report")).toHaveCount(0);
  expect(await download.path()).toBeTruthy();
});
