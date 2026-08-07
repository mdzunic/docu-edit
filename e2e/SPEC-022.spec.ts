import { test, expect, type Download, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-022 — tables: the insert-table toolbar control (grid picker and
// dialog), typing/formatting/lists/styles inside cells, Tab and arrow
// navigation, rectangular cell selection, the row/column context menu and
// hover handles, delete-until-gone behaviour, and .docx/.odt round trips
// including the table.docx fixture's empty and multi-paragraph cells.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const tableFixture = join(fixturesDir, "table.docx");

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
const cellAt = (page: Page, row: number, col: number) =>
  page.getByTestId(`table-cell-${row}-${col}`);
const cellParagraph = (page: Page, row: number, col: number) =>
  cellAt(page, row, col).locator(".doc-paragraph").first();

/** Opens the toolbar grid picker and inserts rows×cols via its cells. */
async function insertViaPicker(page: Page, rows: number, cols: number): Promise<void> {
  await page.getByTestId("insert-table-button").click();
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

/** Minimal central-directory walk over a downloaded archive — the same
 * helper as e2e/SPEC-006.spec.ts. */
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

// --- inserting (AC-1..AC-4) -------------------------------------------------

test("the toolbar control opens a 10×10 grid picker that inserts the chosen size (AC-1, AC-3)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("insert-table-button").click();
  const grid = page.getByTestId("insert-table-grid");
  await expect(grid).toBeVisible();
  // The grid offers up to 10×10.
  await expect(page.getByTestId("table-grid-cell-10-10")).toBeVisible();
  await expect(page.getByTestId("table-grid-cell-11-1")).toHaveCount(0);
  // Hovering previews the size.
  await page.getByTestId("table-grid-cell-3-4").hover();
  await expect(page.getByTestId("insert-table-grid-label")).toHaveText("3 × 4");
  await page.getByTestId("table-grid-cell-3-4").click();
  await expect(table(page)).toHaveCount(1);
  await expect(table(page).locator("tr")).toHaveCount(3);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(4);
});

test("equal column widths span the text area (AC-3)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 4);
  const cols = table(page).locator("col");
  await expect(cols).toHaveCount(4);
  const widths = await cols.evaluateAll((els) => els.map((el) => (el as HTMLElement).style.width));
  expect(new Set(widths).size).toBe(1);
  // The table spans (approximately) the full text area (the page-margin
  // box — since SPEC-028 the editable root wraps whole pages).
  const tableWidth = (await table(page).boundingBox())!.width;
  const contentWidth = (await page.getByTestId("page-margin").boundingBox())!.width;
  expect(Math.abs(tableWidth - contentWidth)).toBeLessThan(8);
});

test("the Insert table… dialog inserts larger sizes (AC-2)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("insert-table-button").click();
  await page.getByTestId("insert-table-dialog-item").click();
  await expect(page.getByTestId("insert-table-dialog")).toBeVisible();
  await page.getByTestId("insert-table-rows").fill("12");
  await page.getByTestId("insert-table-cols").fill("3");
  await page.getByTestId("insert-table-confirm").click();
  await expect(table(page)).toHaveCount(1);
  await expect(table(page).locator("tr")).toHaveCount(12);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(3);
});

test("inserting puts the caret in the first cell, which accepts typing (AC-4, AC-5)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("hello cell");
  await expect(cellParagraph(page, 0, 0)).toHaveText("hello cell");
  await expect(cellParagraph(page, 0, 1)).toHaveText("");
});

// --- editing inside cells (AC-5..AC-9) --------------------------------------

test("cells accept character formatting (AC-6)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 2);
  await page.keyboard.type("bold me");
  await page.keyboard.press("Control+a");
  await page.getByTestId("format-bold").click();
  const span = cellParagraph(page, 0, 0).locator("span").first();
  await expect(span).toHaveCSS("font-weight", "700");
});

test("cells accept paragraph formatting (AC-7)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 2);
  await page.keyboard.type("centre me");
  await page.getByTestId("align-center").click();
  await expect(cellParagraph(page, 0, 0)).toHaveCSS("text-align", "center");
});

test("cells accept lists with their own numbering context (AC-8)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 1, 2);
  await page.keyboard.type("first");
  await page.getByTestId("list-numbered").click();
  await expect(cellParagraph(page, 0, 0)).toHaveAttribute("data-list-marker", "1.");
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("second");
  await expect(cellAt(page, 0, 0).locator(".doc-paragraph").nth(1)).toHaveAttribute(
    "data-list-marker",
    "2.",
  );
  // The neighbouring cell is its own numbering context: a fresh list
  // starts at 1 again.
  await cellParagraph(page, 0, 1).click();
  await page.keyboard.type("other");
  await page.getByTestId("list-numbered").click();
  await expect(cellParagraph(page, 0, 1)).toHaveAttribute("data-list-marker", "1.");
});

test("cells accept styles: a Heading1 cell paragraph renders as a heading (AC-9)", async ({ page }) => {
  await openFixture(page, tableFixture);
  // The fixture's first cell paragraph carries pStyle Heading1 — it
  // renders as an h1 inside the cell and keeps its cell address.
  const heading = cellAt(page, 0, 0).locator("h1.doc-paragraph");
  await expect(heading).toHaveCount(1);
  await expect(heading).toHaveText("Alpha");
});

// --- Tab and arrow navigation (AC-10..AC-14) --------------------------------

test("Tab moves to the next cell, Shift+Tab to the previous (AC-10, AC-11)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("a");
  await page.keyboard.press("Tab");
  await page.keyboard.type("b");
  await expect(cellParagraph(page, 0, 1)).toHaveText("b");
  // Tab at the end of a row wraps to the next row's first cell.
  await page.keyboard.press("Tab");
  await page.keyboard.type("c");
  await expect(cellParagraph(page, 1, 0)).toHaveText("c");
  // Shift+Tab moves back to the previous cell (its end).
  await page.keyboard.press("Shift+Tab");
  await page.keyboard.type("!");
  await expect(cellParagraph(page, 0, 1)).toHaveText("b!");
});

test("Tab in the last cell appends a new row (AC-12)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await cellParagraph(page, 1, 1).click();
  await page.keyboard.press("Tab");
  await expect(table(page).locator("tr")).toHaveCount(3);
  await page.keyboard.type("new row");
  await expect(cellParagraph(page, 2, 0)).toHaveText("new row");
});

test("arrow keys move between cells at cell boundaries (AC-13)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("ab");
  // ArrowRight at the cell's end enters the next cell.
  await page.keyboard.press("ArrowRight");
  await page.keyboard.type("x");
  await expect(cellParagraph(page, 0, 1)).toHaveText("x");
  // ArrowDown moves to the cell below.
  await page.keyboard.press("ArrowDown");
  await page.keyboard.type("y");
  await expect(cellParagraph(page, 1, 1)).toHaveText("y");
  // ArrowLeft at the cell's start enters the previous cell.
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.type("z");
  await expect(cellParagraph(page, 1, 0)).toHaveText("z");
});

test("the caret never gets trapped: arrows exit the table at its edges (AC-14)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("below");
  await page.keyboard.press("Home");
  await insertViaPicker(page, 1, 2);
  // Down from the last row leaves the table into the following paragraph.
  await cellParagraph(page, 0, 1).click();
  await page.keyboard.press("ArrowDown");
  await page.keyboard.type("!");
  await expect(content(page).locator("p.doc-paragraph[data-block]").last()).toContainText("!");
});

// --- cell-range selection (AC-15, AC-16) ------------------------------------

test("dragging across cells selects a highlighted rectangular range (AC-15, AC-16)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 3, 3);
  const from = (await cellAt(page, 0, 0).boundingBox())!;
  const to = (await cellAt(page, 1, 1).boundingBox())!;
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 8 });
  await page.mouse.up();
  await expect(table(page).locator("td.cell-selected")).toHaveCount(4);
  await expect(cellAt(page, 0, 0)).toHaveClass(/cell-selected/);
  await expect(cellAt(page, 1, 1)).toHaveClass(/cell-selected/);
  await expect(cellAt(page, 2, 2)).not.toHaveClass(/cell-selected/);
  // Escape clears the highlight.
  await page.keyboard.press("Escape");
  await expect(table(page).locator("td.cell-selected")).toHaveCount(0);
});

// --- context menu (AC-17..AC-19) --------------------------------------------

test("the cell context menu inserts rows above/below and columns left/right (AC-17, AC-18)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("origin");

  await cellParagraph(page, 0, 0).click({ button: "right" });
  await expect(page.getByTestId("context-menu-insert-row-below")).toBeVisible();
  await page.getByTestId("context-menu-insert-row-below").click();
  await expect(table(page).locator("tr")).toHaveCount(3);
  await expect(cellParagraph(page, 1, 0)).toHaveText("");

  await cellParagraph(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-insert-row-above").click();
  await expect(table(page).locator("tr")).toHaveCount(4);
  // The original content moved down one row.
  await expect(cellParagraph(page, 1, 0)).toHaveText("origin");

  await cellParagraph(page, 1, 0).click({ button: "right" });
  await page.getByTestId("context-menu-insert-col-left").click();
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(3);
  await expect(cellParagraph(page, 1, 1)).toHaveText("origin");

  await cellParagraph(page, 1, 1).click({ button: "right" });
  await page.getByTestId("context-menu-insert-col-right").click();
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(4);
});

test("the cell context menu deletes rows, columns and the whole table (AC-19)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 3, 3);
  await page.keyboard.type("keep");

  await cellParagraph(page, 1, 0).click({ button: "right" });
  await page.getByTestId("context-menu-delete-row").click();
  await expect(table(page).locator("tr")).toHaveCount(2);
  await expect(cellParagraph(page, 0, 0)).toHaveText("keep");

  await cellParagraph(page, 0, 1).click({ button: "right" });
  await page.getByTestId("context-menu-delete-col").click();
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(2);

  await cellParagraph(page, 0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-delete-table").click();
  await expect(table(page)).toHaveCount(0);
});

test("a context menu outside a table offers no table actions", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("plain");
  await content(page).locator(".doc-paragraph").first().click({ button: "right" });
  await expect(page.getByTestId("context-menu-copy")).toBeVisible();
  await expect(page.getByTestId("context-menu-delete-table")).toHaveCount(0);
});

// --- hover handles (AC-20, AC-21) -------------------------------------------

test("hover handles insert rows and columns without opening a menu (AC-20)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("first");
  await cellParagraph(page, 0, 0).hover();
  await expect(page.getByTestId("table-handles")).toBeVisible();
  // Insert a column at the leading boundary.
  await page.getByTestId("table-handle-insert-col-0").click();
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(3);
  await expect(cellParagraph(page, 0, 1)).toHaveText("first");
  // Insert a row at the trailing boundary.
  await cellParagraph(page, 0, 1).hover();
  await page.getByTestId("table-handle-insert-row-2").click();
  await expect(table(page).locator("tr")).toHaveCount(3);
});

test("hover handles delete rows and columns; deleting every row deletes the table and the caret lands where it was (AC-21..AC-23)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("after");
  await page.keyboard.press("Home");
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("doomed");

  await cellParagraph(page, 0, 0).hover();
  await page.getByTestId("table-handle-delete-col-1").click();
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(1);

  await cellParagraph(page, 0, 0).hover();
  await page.getByTestId("table-handle-delete-row-1").click();
  await expect(table(page).locator("tr")).toHaveCount(1);

  // Deleting the last row deletes the table (AC-22) and the caret lands
  // where the table was (AC-23): typing goes into the paragraph that took
  // its place.
  await cellParagraph(page, 0, 0).hover();
  await page.getByTestId("table-handle-delete-row-0").click();
  await expect(table(page)).toHaveCount(0);
  await page.keyboard.type("landed ");
  await expect(content(page).locator(".doc-paragraph").first()).toHaveText("landed after");
});

// --- round trips (AC-24, AC-25) ---------------------------------------------

test("tables round-trip through .docx with grid and cell widths (AC-24)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 3);
  await page.keyboard.type("R1C1");
  await page.keyboard.press("Tab");
  await page.keyboard.type("R1C2");

  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("<w:tbl>");
  expect((documentXml.match(/<w:gridCol /g) ?? []).length).toBe(3);
  expect((documentXml.match(/<w:tr>/g) ?? []).length).toBe(2);
  expect(documentXml).toContain('<w:tcW w:w="');
  expect(documentXml).toContain("R1C1");

  // Re-open the saved file: same dimensions and content (AC-24).
  await openFixture(page, path);
  await expect(table(page).locator("tr")).toHaveCount(2);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(3);
  await expect(cellParagraph(page, 0, 0)).toHaveText("R1C1");
  await expect(cellParagraph(page, 0, 1)).toHaveText("R1C2");
});

test("tables round-trip through .odt with column definitions (AC-25)", async ({ page }) => {
  await createDocument(page);
  await insertViaPicker(page, 2, 2);
  await page.keyboard.type("odt cell");

  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")!.toString("utf8");
  expect(contentXml).toContain("<table:table");
  expect((contentXml.match(/<table:table-column /g) ?? []).length).toBe(2);
  expect((contentXml.match(/<table:table-row/g) ?? []).length).toBe(2);
  expect(contentXml).toContain('style:family="table-column"');
  expect(contentXml).toContain("style:column-width=");
  expect(contentXml).toContain("odt cell");

  await openFixture(page, path);
  await expect(table(page).locator("tr")).toHaveCount(2);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(2);
  await expect(cellParagraph(page, 0, 0)).toHaveText("odt cell");
});

// --- the .docx fixture (AC-26..AC-28) ---------------------------------------

test("a .docx fixture with a table renders with the right dimensions and content (AC-26)", async ({ page }) => {
  await openFixture(page, tableFixture);
  await expect(table(page)).toHaveCount(1);
  await expect(table(page).locator("tr")).toHaveCount(3);
  await expect(table(page).locator("tr").first().locator("td")).toHaveCount(2);
  await expect(cellParagraph(page, 0, 0)).toHaveText("Alpha");
  await expect(cellParagraph(page, 0, 1)).toHaveText("Bravo");
  await expect(cellParagraph(page, 1, 1)).toHaveText("Charlie");
  await expect(cellParagraph(page, 2, 1)).toHaveText("Delta");
  // The surrounding paragraphs render in order around the table.
  await expect(content(page).locator("p.doc-paragraph[data-block]").first()).toHaveText(
    "Before the table",
  );
  await expect(content(page).locator("p.doc-paragraph[data-block]").last()).toHaveText(
    "After the table",
  );
});

test("an empty fixture cell renders and accepts typing (AC-27)", async ({ page }) => {
  await openFixture(page, tableFixture);
  await expect(cellAt(page, 1, 0).locator(".doc-paragraph")).toHaveCount(1);
  await expect(cellParagraph(page, 1, 0)).toHaveText("");
  await cellParagraph(page, 1, 0).click();
  await page.keyboard.type("filled");
  await expect(cellParagraph(page, 1, 0)).toHaveText("filled");
});

test("a multi-paragraph fixture cell renders both paragraphs (AC-28)", async ({ page }) => {
  await openFixture(page, tableFixture);
  const paragraphs = cellAt(page, 2, 0).locator(".doc-paragraph");
  await expect(paragraphs).toHaveCount(2);
  await expect(paragraphs.nth(0)).toHaveText("First line");
  await expect(paragraphs.nth(1)).toHaveText("Second line");
});
