import { test, expect, type Locator, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-134 — find and replace, and Go To, over e2e/fixtures/spec134-find.xlsx
// (built by src/lib/spreadsheet/fixtures/build.ts: "Sales" active with the
// regions, currency amounts and IF flags; "Locked" protected; "Archive"
// hidden; the name Totals → Sales!$B$10), with the real keyboard and pointer:
// Cmd/Ctrl+F → "north" → 1 of 3 at A2 → Enter → 2 of 3 at A4 → Match entire
// cell → 2 of 2 (A6) → off → 3 of 3 → Within Workbook → 3 of 4 → Look in
// Formulas → 6 of 8 → Find all lists eight cells → Cmd/Ctrl+H → "N." →
// Replace all reports 7 in 7 and the protected skip → Escape → C2's formula
// reads "N." → one undo restores value and formula → Ctrl+G Sales!B3:C4 →
// F5 Totals → F5 Special… Formulas → the Name Box jumps to D20 → an invalid
// new name toasts SPEC-133's message.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
}

async function open(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(join(fixtures, "spec134-find.xlsx"));
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
  await expect(cell(page, "A2")).toHaveText("North");
}

/** The count once the debounced, chunked walk has settled ("Searching…" gone). */
async function expectCount(page: Page, text: string): Promise<void> {
  await expect(page.getByTestId("find-count")).toHaveText(text);
  await expect(page.getByTestId("find-status").filter({ hasText: "Searching…" })).toHaveCount(0);
}

test("find, replace, Go To, Go To Special and the Name Box over the fixture (SPEC-134)", async ({ page }) => {
  await open(page);
  const nameBox = page.getByTestId("name-box");
  await clickCell(page, "A1");
  await expect(nameBox).toHaveValue("A1");

  // Cmd/Ctrl+F opens the sheet's bar, not the browser's.
  await page.keyboard.press("ControlOrMeta+f");
  const bar = page.getByTestId("find-bar");
  await expect(bar).toBeVisible();
  await expect(page.getByTestId("find-query")).toBeFocused();
  await expect(page.getByTestId("find-query")).toHaveAttribute("placeholder", "Find in sheet");
  await expect(page.getByTestId("find-scope")).toHaveText("Cells only");
  await expect(page.getByTestId("find-whole-words")).toHaveCount(0);

  // As you type: the first match at or after A1 is current and selected.
  await page.keyboard.type("north");
  await expectCount(page, "1 of 3");
  await expect(nameBox).toHaveValue("A2");
  await expect(cell(page, "A2")).toHaveAttribute("data-find-current", "true");
  await expect(cell(page, "A4")).toHaveAttribute("data-find-match", "true");
  await expect(cell(page, "A3")).not.toHaveAttribute("data-find-match", "true");

  // Enter: the next match.
  await page.keyboard.press("Enter");
  await expectCount(page, "2 of 3");
  await expect(nameBox).toHaveValue("A4");

  // Match entire cell: "Northwest" no longer matches; the first at/after A4 is A6.
  await page.getByTestId("sheet-find-entire-cell").click();
  await expectCount(page, "2 of 2");
  await expect(nameBox).toHaveValue("A6");
  await page.getByTestId("sheet-find-entire-cell").click();
  await expectCount(page, "3 of 3");

  // Within Workbook adds Locked!A1 (the hidden Archive is never searched).
  await page.getByTestId("sheet-find-within").selectOption("workbook");
  await expect(page.getByTestId("find-query")).toHaveAttribute("placeholder", "Find in workbook");
  await expectCount(page, "3 of 4");

  // Look in Formulas adds the four IF formulas naming "North".
  await page.getByTestId("sheet-find-look-in").selectOption("formulas");
  await expectCount(page, "6 of 8");
  await expect(nameBox).toHaveValue("A6");

  // Find all: one row per matching cell, in search order.
  await page.getByTestId("sheet-find-all").click();
  const rows = page.getByTestId("sheet-find-result");
  await expect(rows).toHaveCount(8);
  await expect(rows.first()).toHaveAttribute("data-ref", "A2");
  await expect(rows.last()).toHaveAttribute("data-sheet", "Locked");
  await expect(rows.nth(1)).toContainText('=IF(A2="North",1,0)');

  // Cmd/Ctrl+H reveals the replace row; Replace all is one step and reports the skip.
  await page.keyboard.press("ControlOrMeta+h");
  await expect(page.getByTestId("find-replace-input")).toBeFocused();
  await page.keyboard.type("N.");
  await page.getByTestId("find-replace-all").click();
  await expect(page.getByTestId("find-status")).toHaveText("Replaced 7 occurrences in 7 cells · skipped 1 protected");

  // Escape closes the bar and clears the marks.
  await page.keyboard.press("Escape");
  await expect(bar).toHaveCount(0);
  await expect(page.locator('[data-find-match="true"]')).toHaveCount(0);
  await expect(cell(page, "A2")).toHaveText("N.");
  await clickCell(page, "C2");
  await expect(nameBox).toHaveValue("C2");
  await expect(page.getByTestId("formula-bar-content")).toHaveText('=IF(A2="N.",1,0)');

  // One undo restores every replaced value and formula.
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "A2")).toHaveText("North");
  await expect(cell(page, "A4")).toHaveText("Northwest");
  await clickCell(page, "C2");
  await expect(page.getByTestId("formula-bar-content")).toHaveText('=IF(A2="North",1,0)');

  // Ctrl+G (Control on every platform): Go To a sheet-qualified range.
  await page.keyboard.press("Control+g");
  const goTo = page.getByTestId("sheet-goto-dialog");
  await expect(goTo).toBeVisible();
  await page.getByTestId("sheet-goto-reference").fill("Sales!B3:C4");
  await page.keyboard.press("Enter");
  await expect(goTo).toHaveCount(0);
  await expect(nameBox).toHaveValue("B3");
  await expect(cell(page, "C4")).toHaveAttribute("aria-selected", "true");
  await expect(cell(page, "B3")).toHaveAttribute("aria-selected", "true");

  // F5: the list holds the name Totals, then the origin just left.
  await page.keyboard.press("F5");
  await expect(goTo).toBeVisible();
  await expect(page.getByTestId("sheet-goto-item").filter({ hasText: "Sales!C2" })).toHaveCount(1);
  await page.getByTestId("sheet-goto-item").filter({ hasText: /^\s*Totals\s*$/ }).click();
  await expect(page.getByTestId("sheet-goto-reference")).toHaveValue("Totals");
  await page.getByTestId("sheet-goto-ok").click();
  await expect(nameBox).toHaveValue("B10");

  // An unknown reference errors in the dialog and leaves the selection alone.
  await page.keyboard.press("F5");
  await page.getByTestId("sheet-goto-reference").fill("Nowhere!A1");
  await page.getByTestId("sheet-goto-ok").click();
  await expect(page.getByTestId("sheet-goto-error")).toHaveText("That isn't a cell reference or a defined name");
  await expect(nameBox).toHaveValue("B10");

  // Special… ▸ Formulas: every formula cell of the used range.
  await page.getByTestId("sheet-goto-special").click();
  const special = page.getByTestId("sheet-goto-special-dialog");
  await expect(special).toBeVisible();
  await page.getByTestId("sheet-goto-special-formulas").check();
  await expect(page.getByTestId("sheet-goto-type-numbers")).toBeEnabled();
  await page.getByTestId("sheet-goto-special-ok").click();
  await expect(special).toHaveCount(0);
  for (const ref of ["C2", "C3", "C4", "C5", "C6", "B10"]) {
    await expect(cell(page, ref)).toHaveAttribute("aria-selected", "true");
  }
  await expect(cell(page, "A2")).toHaveAttribute("aria-selected", "false");

  // The Name Box navigates through the same resolver.
  await nameBox.click();
  await nameBox.fill("D20");
  await nameBox.press("Enter");
  await expect(nameBox).toHaveValue("D20");

  // An invalid new name toasts SPEC-133's message.
  await nameBox.click();
  await nameBox.fill("Nowhere!");
  await nameBox.press("Enter");
  await expect(page.getByTestId("toast").filter({ hasText: "Names must start with a letter, underscore or backslash and can't contain spaces or symbols" })).toBeVisible();
});

test("find keys: Shift+Enter steps back with the wrap hint, F5 never reloads, the editor consumes the keys (SPEC-134)", async ({ page }) => {
  await open(page);
  const nameBox = page.getByTestId("name-box");
  await clickCell(page, "A1");

  await page.keyboard.press("ControlOrMeta+f");
  await page.keyboard.type("north");
  await expectCount(page, "1 of 3");
  // Shift+Enter from the first match wraps to the last.
  await page.keyboard.press("Shift+Enter");
  await expectCount(page, "3 of 3");
  await expect(nameBox).toHaveValue("A6");
  await expect(page.getByTestId("find-status")).toHaveText("Wrapped to top");
  // An invalid expression shows the inline error and searches nothing.
  await page.getByTestId("find-regex").click();
  await page.getByTestId("find-query").fill("(north");
  await expect(page.getByTestId("find-status")).toHaveText("Invalid regular expression");
  await expect(page.getByTestId("find-count")).toHaveText("0 of 0");
  await page.getByTestId("find-close").click();
  await expect(page.getByTestId("find-bar")).toHaveCount(0);

  // F5 opens Go To — the page is not reloaded — and Escape closes it.
  await page.evaluate(() => Reflect.set(window, "__spec134Marker", true));
  await page.keyboard.press("F5");
  await expect(page.getByTestId("sheet-goto-dialog")).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("sheet-goto-dialog")).toHaveCount(0);
  expect(await page.evaluate(() => Reflect.get(window, "__spec134Marker"))).toBe(true);

  // While the cell editor is open, Ctrl+G, F5 and Cmd/Ctrl+F do nothing.
  await clickCell(page, "D1");
  await page.keyboard.type("typing");
  await page.keyboard.press("Control+g");
  await page.keyboard.press("F5");
  await page.keyboard.press("ControlOrMeta+f");
  await expect(page.getByTestId("sheet-goto-dialog")).toHaveCount(0);
  await expect(page.getByTestId("find-bar")).toHaveCount(0);
  expect(await page.evaluate(() => Reflect.get(window, "__spec134Marker"))).toBe(true);
  await page.keyboard.press("Escape");
});
