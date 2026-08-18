import { test, expect, type Download, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";
import { MOD } from "./keys";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-023 — table formatting and cell operations: live column/row
// resizing with minimums, cell merge/split, the table properties dialog
// (borders, shading, padding, vertical alignment, table alignment/width),
// border presets, and .docx/.odt round trips of merged cells, borders,
// shading and widths, including the table-merged.docx fixture (vertically
// merged cells with mixed borders).

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const mergedFixture = join(fixturesDir, "table-merged.docx");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const table = (page: Page) => content(page).locator("table.doc-table");
const cellAt = (page: Page, row: number, col: number) => page.getByTestId(`table-cell-${row}-${col}`);
const cellParagraph = (page: Page, row: number, col: number) => cellAt(page, row, col).locator(".doc-paragraph").first();

async function insertViaPicker(page: Page, rows: number, cols: number): Promise<void> {
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`table-grid-cell-${rows}-${cols}`).click();
  await expect(table(page)).toHaveCount(1);
}

async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("download-menu").click();
  const file: Download = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

/** Minimal central-directory walk over a downloaded archive — same helper
 * as e2e/SPEC-022.spec.ts. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd).toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let i = 0; i < totalEntries; i++) {
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

async function selectCellRange(page: Page, from: [number, number], to: [number, number]): Promise<void> {
  const fromBox = (await cellAt(page, from[0], from[1]).boundingBox())!;
  const toBox = (await cellAt(page, to[0], to[1]).boundingBox())!;
  await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(toBox.x + toBox.width / 2, toBox.y + toBox.height / 2, { steps: 8 });
  await page.mouse.up();
}

// --- column/row resize (AC-1..5) --------------------------------------------

test("dragging a column boundary resizes columns live, enforces a minimum and commits once on release (AC-1..3)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 2);
  await cellParagraph(page, 0, 0).hover();
  const handle = page.getByTestId("table-resize-col-0");
  await expect(handle).toBeVisible();
  const before = await table(page).locator("col").first().evaluate((el) => (el as HTMLElement).style.width);
  const box = (await handle.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + 60, box.y + box.height / 2, { steps: 6 });
  // Live during the drag: the DOM column width already changed.
  const during = await table(page).locator("col").first().evaluate((el) => (el as HTMLElement).style.width);
  expect(during).not.toBe(before);
  // Dragging far past the other column's start still leaves it at (or above)
  // the minimum width rather than collapsing to zero (AC-2).
  await page.mouse.move(box.x + 2000, box.y + box.height / 2, { steps: 6 });
  await page.mouse.up();
  const secondColWidth = await table(page).locator("col").nth(1).evaluate((el) => parseFloat((el as HTMLElement).style.width));
  expect(secondColWidth).toBeGreaterThan(0);
  expect(secondColWidth).toBeLessThan(1);
  // One undo step reverts the whole drag (AC-3, AC-34).
  await page.keyboard.press(`${MOD}+z`);
  const after = await table(page).locator("col").first().evaluate((el) => (el as HTMLElement).style.width);
  expect(after).toBe(before);
});

test("dragging a row boundary sets a minimum row height; the row still grows for its content (AC-4, AC-5)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 1);
  await cellParagraph(page, 0, 0).hover();
  const handle = page.getByTestId("table-resize-row-0");
  await expect(handle).toBeVisible();
  const box = (await handle.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2, box.y + 150, { steps: 6 });
  await page.mouse.up();
  const rowHeight = (await table(page).locator("tr").first().boundingBox())!.height;
  expect(rowHeight).toBeGreaterThan(100);
  // The row still grows taller than its minimum for a lot of content.
  await cellParagraph(page, 0, 0).click();
  for (let i = 0; i < 15; i++) {
    await page.keyboard.type(`line ${i}`);
    await page.keyboard.press("Enter");
  }
  const grownHeight = (await table(page).locator("tr").first().boundingBox())!.height;
  expect(grownHeight).toBeGreaterThan(rowHeight);
});

// --- merge and split (AC-6..9) ----------------------------------------------

test("selecting a range and choosing Merge cells produces one cell with content joined as paragraphs (AC-6, AC-7)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await cellParagraph(page, 0, 0).click();
  await page.keyboard.type("A1");
  await cellParagraph(page, 0, 1).click();
  await page.keyboard.type("B1");
  await selectCellRange(page, [0, 0], [0, 1]);
  await cellAt(page, 0, 0).click({ button: "right" });
  await expect(page.getByTestId("context-menu-merge-cells")).toBeVisible();
  await page.getByTestId("context-menu-merge-cells").click();
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(1);
  await expect(cellAt(page, 0, 0)).toHaveAttribute("colspan", "2");
  await expect(cellAt(page, 0, 0).locator(".doc-paragraph")).toHaveCount(2);
  await expect(cellAt(page, 0, 0).locator(".doc-paragraph").nth(0)).toHaveText("A1");
  await expect(cellAt(page, 0, 0).locator(".doc-paragraph").nth(1)).toHaveText("B1");
  // Single undo step (AC-34).
  await page.keyboard.press(`${MOD}+z`);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(2);
});

test("Split cell… asks for rows/columns and splits the cell accordingly (AC-8, AC-9)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 1);
  await cellParagraph(page, 0, 0).click();
  await page.keyboard.type("original");
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-split-cell").click();
  await expect(page.getByTestId("split-cell-dialog")).toBeVisible();
  await page.getByTestId("split-cell-rows").fill("2");
  await page.getByTestId("split-cell-cols").fill("2");
  await page.getByTestId("split-cell-apply").click();
  await expect(table(page).locator("tr")).toHaveCount(2);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(2);
  await expect(cellParagraph(page, 0, 0)).toHaveText("original");
  await expect(cellParagraph(page, 0, 1)).toHaveText("");
  await expect(cellParagraph(page, 1, 0)).toHaveText("");
  // Single undo step (AC-34).
  await page.keyboard.press(`${MOD}+z`);
  await expect(table(page).locator("tr")).toHaveCount(1);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(1);
  await expect(cellParagraph(page, 0, 0)).toHaveText("original");
});

// --- table properties dialog (AC-10..23) ------------------------------------

test("the table properties dialog sets border style/width/colour per table, with the standard presets (AC-10..13, AC-18..21)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await expect(page.getByTestId("table-properties-dialog")).toBeVisible();
  await page.getByTestId("table-props-border-style").selectOption("dashed");
  await page.getByTestId("table-props-border-width").fill("2");
  await page.getByTestId("table-props-border-color").fill("#ff0000");
  await page.getByTestId("table-props-preset-all").click();
  await page.getByTestId("table-props-apply").click();
  const borderTop = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderTopStyle);
  expect(borderTop).toBe("dashed");

  // "Outer only" clears the inner shared edges.
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-preset-outer").click();
  await page.getByTestId("table-props-apply").click();
  const innerRight00 = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderRightStyle);
  expect(innerRight00).toBe("none");
  const outerTop00 = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderTopStyle);
  expect(outerTop00).toBe("dashed");

  // "Inner only" clears the outer edges.
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-preset-inner").click();
  await page.getByTestId("table-props-apply").click();
  const outerTopAfterInner = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderTopStyle);
  expect(outerTopAfterInner).toBe("none");
  const innerRightAfterInner = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderRightStyle);
  expect(innerRightAfterInner).toBe("dashed");

  // "None" clears every edge.
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-preset-none").click();
  await page.getByTestId("table-props-apply").click();
  const clearedRight = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderRightStyle);
  expect(clearedRight).toBe("none");
});

test("border settings apply to just the selected cells, leaving the rest of the table alone (AC-14)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  // A genuine multi-cell range — a drag that never leaves its origin cell
  // is a plain click, not a range selection (SPEC-022 AC-15).
  await selectCellRange(page, [0, 0], [0, 1]);
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-preset-all").click();
  await page.getByTestId("table-props-apply").click();
  const targetTop = await cellAt(page, 0, 0).evaluate((el) => getComputedStyle(el).borderTopStyle);
  expect(targetTop).toBe("solid");
  const selectedTop = await cellAt(page, 0, 1).evaluate((el) => getComputedStyle(el).borderTopStyle);
  expect(selectedTop).toBe("solid");
  const untouchedTop = await cellAt(page, 1, 1).evaluate((el) => getComputedStyle(el).borderTopColor);
  // The untouched cell keeps the canvas's default gridline colour rather
  // than the dialog's chosen colour.
  expect(untouchedTop).not.toBe("rgb(0, 0, 0)");
});

// Regression (2026-08-18, tracking issue #141): the table context menu grew
// past the viewport as sections accumulated (fields, lists, SPEC-055's
// caption entry) and its position:fixed placement had no clamping, so
// "Table properties…" — the last item — became unreachable on shorter
// viewports; the factory container's taller font metrics hit it at 720px
// before macOS did. A deliberately short viewport forces the overflow on
// every platform: the menu must shift to fit, keeping its last item fully
// inside the viewport and clickable.
test("the table context menu clamps into a short viewport, keeping Table properties reachable", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 500 });
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await cellAt(page, 0, 0).click({ button: "right" });
  const item = page.getByTestId("context-menu-table-properties");
  await expect(item).toBeVisible();
  const box = await item.boundingBox();
  expect(box).not.toBeNull();
  expect(box!.y).toBeGreaterThanOrEqual(0);
  expect(box!.y + box!.height).toBeLessThanOrEqual(500);
  await item.click();
  await expect(page.getByTestId("table-props-apply")).toBeVisible();
});

test("the dialog sets cell shading, padding and vertical alignment (AC-15..17)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 1);
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-shading-enabled").check();
  await page.getByTestId("table-props-shading-color").fill("#00ff00");
  await page.getByTestId("table-props-padding").fill("1");
  await page.getByTestId("table-props-valign").selectOption("bottom");
  await page.getByTestId("table-props-apply").click();
  await expect(cellAt(page, 0, 0)).toHaveCSS("background-color", "rgb(0, 255, 0)");
  const paddingTop = await cellAt(page, 0, 0).evaluate((el) => parseFloat(getComputedStyle(el).paddingTop));
  // 1cm at CSS's fixed 96px/inch ratio.
  expect(paddingTop).toBeCloseTo(37.795, 1);
  await expect(cellAt(page, 0, 0)).toHaveCSS("vertical-align", "bottom");
});

test("the dialog sets table alignment and width (AC-22, AC-23)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 2);
  // The text area is the page-margin box (since SPEC-028 the editable root
  // wraps whole pages).
  const contentBox = (await page.getByTestId("page-margin").boundingBox())!;
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-width-mode").selectOption("percent");
  await page.getByTestId("table-props-width-percent").fill("50");
  await page.getByTestId("table-props-alignment").selectOption("center");
  await page.getByTestId("table-props-apply").click();
  const tableBox = (await table(page).boundingBox())!;
  expect(tableBox.width).toBeLessThan(contentBox.width * 0.6);
  // Centred: roughly equal margins on both sides of the text area.
  const leftGap = tableBox.x - contentBox.x;
  const rightGap = contentBox.x + contentBox.width - (tableBox.x + tableBox.width);
  expect(Math.abs(leftGap - rightGap)).toBeLessThan(8);
  // The whole dialog session — border/shading/padding/valign/alignment/
  // width all touched together — is one undo step (AC-34).
  await page.keyboard.press(`${MOD}+z`);
  const revertedBox = (await table(page).boundingBox())!;
  expect(revertedBox.width).toBeGreaterThan(contentBox.width * 0.9);
});

// --- round trips (AC-24..31) -------------------------------------------------

test("merged cells and borders round-trip through .docx using w:gridSpan/w:vMerge/w:tcBorders (AC-24, AC-25)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await selectCellRange(page, [0, 0], [1, 0]);
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-merge-cells").click();
  // The merged cell's own bounding-box centre sits on the row-resize handle
  // (a thin strip along the internal row boundary it now spans); click near
  // its top-left corner instead.
  await cellAt(page, 0, 0).click({ button: "right", position: { x: 5, y: 5 } });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-preset-all").click();
  await page.getByTestId("table-props-apply").click();

  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("w:vMerge");
  expect(documentXml).toContain("w:tcBorders");

  await openFixture(page, path);
  await expect(cellAt(page, 0, 0)).toHaveAttribute("rowspan", "2");
});

test("shading and widths round-trip through .docx using w:shd and w:tblW (AC-26, AC-27)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 1);
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-shading-enabled").check();
  await page.getByTestId("table-props-shading-color").fill("#123456");
  await page.getByTestId("table-props-width-mode").selectOption("percent");
  await page.getByTestId("table-props-width-percent").fill("40");
  await page.getByTestId("table-props-apply").click();

  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml.toLowerCase()).toContain('w:fill="123456"');
  expect(documentXml).toContain('w:type="pct"');

  await openFixture(page, path);
  await expect(cellAt(page, 0, 0)).toHaveCSS("background-color", "rgb(18, 52, 86)");
});

test("merged cells, cell styles and widths round-trip through .odt (AC-28..31)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await selectCellRange(page, [0, 0], [1, 0]);
  await cellAt(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-merge-cells").click();
  await cellAt(page, 0, 0).click({ button: "right", position: { x: 5, y: 5 } });
  await page.getByTestId("context-menu-table-properties").click();
  await page.getByTestId("table-props-shading-enabled").check();
  await page.getByTestId("table-props-shading-color").fill("#654321");
  await page.getByTestId("table-props-apply").click();

  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")!.toString("utf8");
  expect(contentXml).toContain("table:number-rows-spanned");
  expect(contentXml).toContain("table:covered-table-cell");
  expect(contentXml.toLowerCase()).toContain("fo:background-color=\"#654321\"");

  await openFixture(page, path);
  await expect(cellAt(page, 0, 0)).toHaveAttribute("rowspan", "2");
  await expect(cellAt(page, 0, 0)).toHaveCSS("background-color", "rgb(101, 67, 33)");
});

// --- the .docx fixture (AC-32, AC-33) ----------------------------------------

test("a .docx fixture with vertically merged cells and mixed borders renders correctly (AC-32)", async ({ page }) => {
  await openFixture(page, mergedFixture);
  await expect(table(page)).toHaveCount(1);
  await expect(cellAt(page, 0, 0)).toHaveAttribute("rowspan", "2");
  await expect(table(page).locator("tr").nth(1).locator("td")).toHaveCount(2);
  await expect(cellAt(page, 0, 1)).toHaveCSS("border-bottom-style", "double");
  await expect(cellAt(page, 1, 1)).toHaveCSS("vertical-align", "middle");
});

test("the merged/bordered .docx fixture survives a save/re-open unchanged (AC-33)", async ({ page }) => {
  await openFixture(page, mergedFixture);
  const path = await download(page, "docx");
  await openFixture(page, path);
  await expect(table(page)).toHaveCount(1);
  await expect(cellAt(page, 0, 0)).toHaveAttribute("rowspan", "2");
  await expect(table(page).locator("tr").nth(1).locator("td")).toHaveCount(2);
  await expect(cellAt(page, 0, 1)).toHaveCSS("border-bottom-style", "double");
  await expect(cellAt(page, 2, 0).locator(".doc-paragraph").first()).toHaveText("Golf");
});
