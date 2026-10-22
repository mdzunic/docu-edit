import { test, expect, type Locator, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-127 — the function picker (AC-52), with Playwright's real keyboard:
// open mini-book.xlsx → the empty J1 → type "=XL" → function-autocomplete
// lists XLOOKUP → Tab → the editor reads "=XLOOKUP(" and function-arg-current
// "lookup_value" → type "x", → "lookup_array" → type ROUND( → the hint shows
// ROUND with "number" current → Escape cancels the edit → formula-bar-fx →
// category Financial → search "payment" → PMT → Insert → the editor reads
// "=PMT(" → replace it with =ROUND(PMT(0.05/12,60,10000),2) → Enter → J1
// reads "-188.71" → in J2 =SEQUENCE(3) → Enter → J2 reads "#NAME?",
// data-stale="true" → Shift+F3 opens the dialog, with no SEQUENCE item.

const miniBook = join(dirname(fileURLToPath(import.meta.url)), "fixtures", "mini-book.xlsx");

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

function editor(page: Page): Locator {
  return page.getByTestId("cell-editor");
}

/** Clicks the centre of a cell from its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

test("autocomplete, argument hint and Insert Function; a dynamic array is #NAME?, stale, and never offered (AC-52)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(miniBook);
  await expect(page).toHaveURL(/\/sheets\/.+/);

  await clickCell(page, "J1");
  await expect(editor(page)).toBeFocused();
  await page.keyboard.type("=XL");
  const list = page.getByTestId("function-autocomplete");
  await expect(list).toBeVisible();
  await expect(list).toHaveAttribute("role", "listbox");
  await expect(list.getByTestId("function-autocomplete-item")).toHaveText(["XLOOKUP"]);
  await expect(page.getByTestId("function-autocomplete-summary")).not.toBeEmpty();

  await page.keyboard.press("Tab");
  await expect(editor(page)).toHaveText("=XLOOKUP(");
  await expect(list).toBeHidden();
  await expect(page.getByTestId("function-arg-hint")).toHaveAttribute("role", "tooltip");
  await expect(page.getByTestId("function-arg-current")).toHaveText("lookup_value");

  await page.keyboard.type('"x",');
  await expect(page.getByTestId("function-arg-current")).toHaveText("lookup_array");

  await page.keyboard.type("ROUND(");
  await expect(page.getByTestId("function-arg-hint")).toContainText("ROUND(number, num_digits)");
  await expect(page.getByTestId("function-arg-current")).toHaveText("number");

  await page.keyboard.press("Escape");
  await expect(page.getByTestId("sheet-status-mode")).toHaveText("Ready");
  await expect(page.getByTestId("function-arg-hint")).toBeHidden();

  const fx = page.getByTestId("formula-bar-fx");
  await expect(fx).toHaveAttribute("aria-label", "Insert function");
  await fx.click();
  const dialog = page.getByTestId("insert-function-dialog");
  await expect(dialog).toBeVisible();
  await page.getByTestId("insert-function-category").selectOption("Financial");
  await page.getByTestId("insert-function-search").fill("payment");
  await page.locator('[data-testid="insert-function-item"][data-name="PMT"]').click();
  await expect(page.getByTestId("insert-function-detail")).toContainText("PMT(rate, nper, pv, [fv], [type])");
  await page.getByTestId("insert-function-insert").click();
  await expect(dialog).toBeHidden();
  await expect(editor(page)).toHaveText("=PMT(");
  await expect(editor(page)).toBeFocused();

  await page.keyboard.press("ControlOrMeta+A");
  await page.keyboard.type("=ROUND(PMT(0.05/12,60,10000),2)");
  await page.keyboard.press("Enter");
  await expect(cell(page, "J1")).toHaveText("-188.71");

  await expect(page.getByTestId("name-box")).toHaveValue("J2");
  await page.keyboard.type("=SEQUENCE(3)");
  await page.keyboard.press("Enter");
  await expect(cell(page, "J2")).toHaveText("#NAME?");
  await expect(cell(page, "J2")).toHaveAttribute("data-stale", "true");

  await page.keyboard.press("Shift+F3");
  await expect(dialog).toBeVisible();
  await expect(page.locator('[data-testid="insert-function-item"]').first()).toBeVisible();
  await expect(page.locator('[data-testid="insert-function-item"][data-name="SEQUENCE"]')).toHaveCount(0);
});
