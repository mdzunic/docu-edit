import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

// SPEC-124 — rows, columns and merged cells (AC-35), with the real pointer
// and keyboard: open spec124-structure.xlsx → right-click row header 3 (row
// 3 snaps to rows 2–3, the M2:N3 merge) → "Insert 2 rows above" → on Sheet2 the cell that read
// =Sheet1!A5 reads =Sheet1!A7 → delete column B from its header menu → the
// formula that referenced B5 (H5 =B5*2, now G7) shows #REF! → drag column
// C's boundary wider → double-click column D's boundary (it grows to its
// longest text) → hide row 10, unhide it → select A1:C1 (B1 holds a value)
// → Merge & Centre → the warning → Merge → View ▸ Freeze ▸ Top row → undo,
// redo → download → reopen: merges, freeze and formulas as left, and the
// package's conditional-format sqref, defined name, table ref and chart c:f
// shifted (2 rows down, then one column left). Then on spec124-structure.ods
// insert a column → after download and reopen its named range and
// conditional-format target have moved.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

function colHeader(page: Page, letters: string): Locator {
  return page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) });
}

function rowHeader(page: Page, row: number): Locator {
  return page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${row}$`) });
}

function sheetTab(page: Page, name: string): Locator {
  return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string, modifiers: ("Shift" | "ControlOrMeta")[] = []): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await colHeader(page, letters!).boundingBox();
  const row = await rowHeader(page, Number(digits)).boundingBox();
  for (const key of modifiers) {
    await page.keyboard.down(key);
  }
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  for (const key of modifiers) {
    await page.keyboard.up(key);
  }
}

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
}

async function download(page: Page, fileName: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = await downloadPromise;
  const path = join(tmpdir(), `spec124-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

async function headerMenu(page: Page, header: Locator, entry: string): Promise<Locator> {
  await header.click({ button: "right" });
  const menu = page.getByTestId("sheet-header-menu");
  await expect(menu).toBeVisible();
  return menu.getByTestId(entry);
}

const width = async (locator: Locator): Promise<number> => (await locator.boundingBox())!.width;

async function freezeLines(page: Page): Promise<string[]> {
  return page.getByTestId("sheet-freeze-line").evaluateAll((lines) => lines.map((line) => (line.classList.contains("is-vertical") ? "vertical" : "horizontal")));
}

const text = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes);

test("insert, delete, resize, autofit, hide, merge and freeze on spec124-structure.xlsx, then download and reopen (AC-35)", async ({ page }) => {
  const source = await unzip(new Uint8Array(readFileSync(join(fixtures, "spec124-structure.xlsx"))));
  await open(page, join(fixtures, "spec124-structure.xlsx"));
  // The source freezes the first column.
  expect(await freezeLines(page)).toEqual(["vertical"]);

  // Right-click row header 3: it selects row 3 first, which snaps to rows
  // 2–3 (the M2:N3 merge) — "Insert 2 rows above" puts two rows at row 2.
  const insert = await headerMenu(page, rowHeader(page, 3), "sheet-insert-before");
  await expect(insert).toHaveText("Insert 2 rows above");
  await insert.click();
  await expect(page.getByTestId("sheet-header-menu")).toBeHidden();
  await expect(cell(page, "A7")).toHaveText("Item 4");

  // Sheet2's =Sheet1!A5 follows the rows.
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A1");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("=Sheet1!A7");
  await sheetTab(page, "Sheet1").click();

  // Delete column B: H5's =B5*2 (now G7) shows #REF!.
  const deleteColumn = await headerMenu(page, colHeader(page, "B"), "sheet-delete-band");
  await expect(deleteColumn).toHaveText("Delete column");
  await deleteColumn.click();
  await expect(cell(page, "B1")).toHaveText("Price");
  await expect(cell(page, "G7")).toHaveText("#REF!");

  // Drag column C's right boundary 40 px wider: the guide and tooltip show while dragging.
  const before = (await colHeader(page, "C").boundingBox())!;
  await page.mouse.move(before.x + before.width - 1, before.y + before.height / 2);
  await expect(page.getByTestId("spreadsheet-grid")).toHaveCSS("cursor", "col-resize");
  await page.mouse.down();
  await page.mouse.move(before.x + before.width + 40, before.y + before.height / 2, { steps: 5 });
  await expect(page.getByTestId("sheet-resize-guide")).toBeVisible();
  await expect(page.getByTestId("sheet-resize-tooltip")).toContainText(/^Width: \d+\.\d\d \(\d+ pixels\)$/);
  await page.mouse.up();
  await expect(page.getByTestId("sheet-resize-guide")).toHaveCount(0);
  await expect.poll(() => width(colHeader(page, "C"))).toBeGreaterThan(before.width + 30);

  // Double-click column D's boundary: it grows to its longest text.
  const d = (await colHeader(page, "D").boundingBox())!;
  await page.mouse.dblclick(d.x + d.width - 1, d.y + d.height / 2);
  await expect.poll(() => width(colHeader(page, "D"))).toBeGreaterThan(d.width * 3);

  // Hide row 10, then unhide it from a selection spanning it.
  await (await headerMenu(page, rowHeader(page, 10), "sheet-hide-band")).click();
  await expect(rowHeader(page, 10)).toHaveCount(0);
  await expect(page.locator('[data-testid="sheet-hidden-marker"][title="Hidden row 10"]')).toHaveCount(1);
  await rowHeader(page, 9).click();
  await rowHeader(page, 11).click({ modifiers: ["Shift"] });
  const unhide = await headerMenu(page, rowHeader(page, 9), "sheet-unhide-band");
  await expect(unhide).toBeEnabled();
  await unhide.click();
  await expect(rowHeader(page, 10)).toHaveCount(1);

  // Merge & Centre A1:C1: B1 holds a value, so the warning asks first.
  await clickCell(page, "A1");
  await clickCell(page, "C1", ["Shift"]);
  await page.getByTestId("sheet-merge").click();
  const warning = page.getByTestId("sheet-merge-warning");
  await expect(warning).toContainText("Merging keeps only the upper-left value");
  await page.getByTestId("sheet-merge-confirm").click();
  await expect(warning).toHaveCount(0);
  await expect(page.getByTestId("sheet-merge")).toHaveAttribute("aria-pressed", "true");
  await expect.poll(async () => width(cell(page, "A1"))).toBeGreaterThan((await width(colHeader(page, "A"))) * 2);

  // View ▸ Freeze ▸ Top row; undo; redo.
  await page.getByRole("menubar").getByText("View", { exact: true }).click();
  await page.getByTestId("sheet-menu-freeze").click();
  await expect(page.getByTestId("sheet-menu-unfreeze")).toBeEnabled();
  await page.getByTestId("sheet-menu-freeze-top-row").click();
  await expect.poll(() => freezeLines(page)).toEqual(["horizontal"]);
  await page.getByTestId("spreadsheet-grid").focus();
  await page.keyboard.press("ControlOrMeta+z");
  await expect.poll(() => freezeLines(page)).toEqual(["vertical"]);
  await page.keyboard.press("ControlOrMeta+Shift+z");
  await expect.poll(() => freezeLines(page)).toEqual(["horizontal"]);

  // Download: the retained references moved 2 rows down (from row 2), then
  // one column left.
  const path = await download(page, "structure.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(path)));
  const sheet1 = text(saved.get("xl/worksheets/sheet1.xml"));
  expect(sheet1).toContain(`<conditionalFormatting sqref="B4:B33">`);
  expect(sheet1).toContain(`<mergeCell ref="A1:C1"/>`);
  expect(sheet1).toMatch(/<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/);
  expect(text(saved.get("xl/workbook.xml"))).toContain(`<definedName name="Prices">Sheet1!$B$4:$B$33</definedName>`);
  expect(text(saved.get("xl/tables/table1.xml"))).toContain(`ref="I1:J8"`);
  expect(text(saved.get("xl/charts/chart1.xml"))).toContain(`<c:f>Sheet1!$B$4:$B$33</c:f>`);
  expect(saved.has("xl/calcChain.xml")).toBe(false);
  // Parts the edits never touched keep their bytes.
  for (const name of ["xl/theme/theme1.xml", "docProps/core.xml", "xl/media/image1.png"]) {
    expect(equalBytes(saved.get(name)!, source.get(name)!), name).toBe(true);
  }

  // Reopen: merges, freeze and formulas as left.
  await open(page, path);
  await expect.poll(() => freezeLines(page)).toEqual(["horizontal"]);
  await clickCell(page, "A1");
  await expect(page.getByTestId("sheet-merge")).toHaveAttribute("aria-pressed", "true");
  await expect(cell(page, "G7")).toHaveText("#REF!");
  await clickCell(page, "G7");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("=#REF!*2");
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A1");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("=Sheet1!A7");
});

test("insert a column on spec124-structure.ods: its named range and conditional-format target move (AC-35)", async ({ page }) => {
  await open(page, join(fixtures, "spec124-structure.ods"));
  await expect(cell(page, "B1")).toHaveText("Amount");
  // A whole column B would snap to A:C (the A13:C13 merge): insert ONE
  // column through the keys instead — on a cell they open the dialog.
  await clickCell(page, "B5");
  await page.keyboard.press("ControlOrMeta+Alt+Equal");
  const dialog = page.getByTestId("sheet-insert-dialog");
  await expect(dialog).toBeVisible();
  await page.getByTestId("sheet-insert-delete-entire-column").click();
  await expect(dialog).toHaveCount(0);
  await expect(cell(page, "C1")).toHaveText("Amount");

  const path = await download(page, "structure.ods");
  const content = text((await unzip(new Uint8Array(readFileSync(path)))).get("content.xml"));
  expect(content).toContain(`table:cell-range-address="$Sheet1.$C$2:.$C$11"`);
  expect(content).toContain(`calcext:target-range-address="Sheet1.C2:Sheet1.C11"`);

  await open(page, path);
  await expect(cell(page, "C1")).toHaveText("Amount");
  await expect(cell(page, "B1")).toHaveText("");
});
