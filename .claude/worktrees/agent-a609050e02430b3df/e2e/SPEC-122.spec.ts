import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

// SPEC-122 — cell entry and the sheet writer (AC-60), with Playwright's real
// keyboard: New spreadsheet → "Revenue" in A1, Enter → A2 is active →
// "1,200", Tab → A2 reads "1,200" (#,##0 applied) → "=A2*2" in B2, Enter →
// "2400" → in B3 "=SUM(", ArrowUp twice, Shift+ArrowDown, ")" → the editor
// reads "=SUM(B1:B2)", Enter → "2400" → C1 "10%" → C2 "2026-10-08" is the
// en-US short date "10/8/2026" (its formula bar; the default 64 px column is
// narrower than SPEC-118's width estimate of that text, so the cell shows
// SPEC-118's "#" fill — widening a column on entry is SPEC-124's) → C3 two
// lines with Alt+Enter wraps (its row grows) → D1:D3
// "x" with Ctrl+Enter → Delete → undo → "=1+" in E1 shows the formula error
// → Escape → download and re-open: everything as left, B2's formula bar
// "=A2*2" → a copy as .ods re-opens with the same displayed values. Then
// spec122-shapes.xlsx: edit one cell on "Edit" → download → every other
// sheet part equalBytes with the fixture and no xl/calcChain.xml → on
// "Locked", typing into a locked cell shows "This cell is on a protected
// sheet".

const shapesFixture = join(dirname(fileURLToPath(import.meta.url)), "fixtures", "spec122-shapes.xlsx");

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** A position shows nothing: no cell element there with any text (the
 * grid mounts empty positions too while it renders densely). */
async function expectEmpty(page: Page, ref: string): Promise<void> {
  await expect(cell(page, ref).filter({ hasText: /\S/ })).toHaveCount(0);
}

function editor(page: Page): Locator {
  return page.getByTestId("cell-editor");
}

function mode(page: Page): Locator {
  return page.getByTestId("sheet-status-mode");
}

function tab(page: Page, name: string): Locator {
  return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
}

function rowHeader(page: Page, row: number): Locator {
  return page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${row}$`) });
}

/** Clicks the centre of a cell — empty cells have no element of their own,
 * so the point comes from its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await rowHeader(page, Number(digits)).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

/** Types an entry into the active cell and commits it with `key`. */
async function enter(page: Page, text: string, key = "Enter"): Promise<void> {
  await page.keyboard.type(text);
  await page.keyboard.press(key);
}

/** Downloads through `trigger`; resolves to the saved file's path. */
async function download(page: Page, trigger: () => Promise<void>, fileName: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const saved = await downloadPromise;
  const path = join(tmpdir(), `spec122-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

async function reopen(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
}

/** C2 holds the date 2026-10-08 as the en-US short date: the formula bar
 * shows "10/8/2026"; the cell, in a default-width column, SPEC-118's "#"
 * overflow fill (its width estimate of "10/8/2026" is 63 px of 59). */
async function expectDateInC2(page: Page): Promise<void> {
  await expect(cell(page, "C2")).toHaveText(/^#+$/);
  await clickCell(page, "C2");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("10/8/2026");
}

/** A1–D3 as the flow leaves them. */
async function expectEnteredValues(page: Page): Promise<void> {
  await expect(cell(page, "A1")).toHaveText("Revenue");
  await expect(cell(page, "A2")).toHaveText("1,200");
  await expect(cell(page, "B2")).toHaveText("2400");
  await expect(cell(page, "B3")).toHaveText("2400");
  await expect(cell(page, "C1")).toHaveText("10%");
  await expectDateInC2(page);
  await expect(cell(page, "C3")).toContainText("line one");
  await expect(cell(page, "C3")).toContainText("line two");
  for (const ref of ["D1", "D2", "D3"]) {
    await expect(cell(page, ref)).toHaveText("x");
  }
}

test("type values and formulas, point at a range, fill, clear, undo, and round-trip both formats (AC-60)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("new-spreadsheet").first().click();
  await expect(page).toHaveURL(/\/sheets\/.+/);

  await clickCell(page, "A1");
  await expect(editor(page)).toBeFocused();
  await expect(mode(page)).toHaveText("Ready");
  await page.keyboard.type("Revenue");
  await expect(mode(page)).toHaveText("Enter");
  await expect(editor(page)).toHaveText("Revenue");
  await page.keyboard.press("Enter");
  await expect(cell(page, "A1")).toHaveText("Revenue");
  await expect(page.getByTestId("name-box")).toHaveValue("A2");
  await expect(mode(page)).toHaveText("Ready");

  await enter(page, "1,200", "Tab");
  await expect(cell(page, "A2")).toHaveText("1,200");
  await expect(page.getByTestId("name-box")).toHaveValue("B2");

  await enter(page, "=A2*2");
  await expect(cell(page, "B2")).toHaveText("2400");
  // Enter after the Tab returns to the run's starting column, one row down.
  await expect(page.getByTestId("name-box")).toHaveValue("A3");
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("name-box")).toHaveValue("B3");

  // Point mode: the arrows insert, then replace, the reference at the caret.
  await page.keyboard.type("=SUM(");
  await page.keyboard.press("ArrowUp");
  await expect(mode(page)).toHaveText("Point");
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.type(")");
  await expect(editor(page)).toHaveText("=SUM(B1:B2)");
  await page.keyboard.press("Enter");
  await expect(cell(page, "B3")).toHaveText("2400");

  await clickCell(page, "C1");
  await enter(page, "10%");
  await expect(cell(page, "C1")).toHaveText("10%");
  await expect(page.getByTestId("name-box")).toHaveValue("C2");
  await enter(page, "2026-10-08");
  await expectDateInC2(page);
  await page.keyboard.press("ArrowDown");
  await expect(page.getByTestId("name-box")).toHaveValue("C3");

  // Alt+Enter breaks the line; the committed cell wraps and its row grows.
  const rowBefore = (await rowHeader(page, 3).boundingBox())!.height;
  await page.keyboard.type("line one");
  await page.keyboard.press("Alt+Enter");
  await page.keyboard.type("line two");
  await page.keyboard.press("Enter");
  await expect(cell(page, "C3")).toContainText("line two");
  await expect.poll(async () => (await rowHeader(page, 3).boundingBox())!.height).toBeGreaterThan(rowBefore);

  // Ctrl+Enter writes the entry into every cell of the selection.
  await clickCell(page, "D1");
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.type("x");
  await page.keyboard.press("Control+Enter");
  for (const ref of ["D1", "D2", "D3"]) {
    await expect(cell(page, ref)).toHaveText("x");
  }
  await page.keyboard.press("Delete");
  for (const ref of ["D1", "D2", "D3"]) {
    await expectEmpty(page, ref);
  }
  await page.keyboard.press("Control+z");
  for (const ref of ["D1", "D2", "D3"]) {
    await expect(cell(page, ref)).toHaveText("x");
  }

  // A formula that does not parse keeps the editor open with the message.
  await clickCell(page, "E1");
  await enter(page, "=1+");
  await expect(page.getByTestId("cell-editor-error")).toHaveText("There's a problem with this formula");
  await expect(mode(page)).not.toHaveText("Ready");
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("cell-editor-error")).toHaveCount(0);
  await expect(mode(page)).toHaveText("Ready");
  await expectEmpty(page, "E1");
  const blankUrl = page.url();

  const xlsx = await download(page, () => page.getByTestId("spreadsheet-download").click(), "entry.xlsx");
  await reopen(page, xlsx);
  await expectEnteredValues(page);
  await cell(page, "B2").click();
  await expect(page.getByTestId("formula-bar-content")).toHaveText("=A2*2");

  await page.goto(blankUrl);
  await expect(cell(page, "A1")).toHaveText("Revenue");
  const ods = await download(
    page,
    async () => {
      await page.getByTestId("spreadsheet-download-menu").click();
      await page.getByTestId("spreadsheet-download-copy").click();
    },
    "entry.ods",
  );
  await reopen(page, ods);
  await expect(page.getByTestId("spreadsheet-format")).toHaveText("OpenDocument Spreadsheet (.ods)");
  await expectEnteredValues(page);
});

test("an edit to spec122-shapes.xlsx rebuilds only its own sheet; a protected sheet refuses entry (AC-60)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(shapesFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(tab(page, "Edit")).toHaveAttribute("aria-selected", "true");

  await clickCell(page, "B2");
  await enter(page, "99");
  await expect(cell(page, "B2")).toHaveText("99");

  const path = await download(page, () => page.getByTestId("spreadsheet-download").click(), "shapes.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(path)));
  const source = await unzip(new Uint8Array(readFileSync(shapesFixture)));
  expect(saved.has("xl/calcChain.xml")).toBe(false);
  const editPart = "xl/worksheets/sheet1.xml";
  expect(equalBytes(saved.get(editPart)!, source.get(editPart)!)).toBe(false);
  for (const [name, bytes] of source) {
    if (name.startsWith("xl/worksheets/") && name !== editPart) {
      expect(saved.has(name), name).toBe(true);
      expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
    }
  }

  await tab(page, "Locked").click();
  await clickCell(page, "A1");
  await page.keyboard.type("a");
  await expect(page.getByTestId("toast").filter({ hasText: "This cell is on a protected sheet" })).toBeVisible();
  await expect(mode(page)).toHaveText("Ready");
});
