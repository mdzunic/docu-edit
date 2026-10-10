import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-121 — create a workbook and manage its sheets (AC-24): New spreadsheet
// → "+" twice → Sheet1, Sheet2, Sheet3 → rename Sheet2 to "Budget" →
// duplicate it ("Budget (2)") → move the copy first through the tab's context
// menu by keyboard → hide Sheet3 and unhide it through the dialog → colour a
// tab → delete Sheet1 → undo → download and re-open: the tabs read as left.
// Then spec121-links.xlsx: rename "Data" to "Facts" → Summary!C2's formula
// reads "=Facts!B2*2" → delete "Lookup" after confirming → a cell that
// referenced it reads "#REF!" → download and re-open: both hold. Last, a copy
// of the blank workbook as .ods re-opens with the same tab names.

const fixture = join(dirname(fileURLToPath(import.meta.url)), "fixtures", "spec121-links.xlsx");

function tabs(page: Page) {
  return page.getByTestId("sheet-tab");
}

function tab(page: Page, name: string) {
  return tabs(page).filter({ hasText: new RegExp(`^\\s*${name.replace(/[()]/g, "\\$&")}\\s*$`) });
}

function cell(page: Page, ref: string) {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

async function tabMenu(page: Page, name: string, action: string): Promise<void> {
  await tab(page, name).click({ button: "right" });
  await expect(page.getByTestId("sheet-tab-menu")).toBeVisible();
  await page.getByTestId(`sheet-tab-menu-${action}`).click();
}

async function rename(page: Page, from: string, to: string): Promise<void> {
  await tab(page, from).dblclick();
  const input = page.getByTestId("sheet-tab-rename-input");
  await expect(input).toBeVisible();
  await input.fill(to);
  await input.press("Enter");
  await expect(input).toHaveCount(0);
}

/** Downloads through `trigger`, then opens the saved file from the library. */
async function downloadAndReopen(page: Page, trigger: () => Promise<void>, fileName: string): Promise<void> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const download = await downloadPromise;
  const path = join(tmpdir(), `spec121-${Date.now()}-${fileName}`);
  await download.saveAs(path);
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
}

test("create a workbook, manage its sheets, undo, and round-trip the download (AC-24)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("new-spreadsheet").first().click();
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const blankUrl = page.url();
  await expect(tabs(page)).toHaveText(["Sheet1"]);

  await page.getByTestId("sheet-add").click();
  await page.getByTestId("sheet-add").click();
  await expect(tabs(page)).toHaveText(["Sheet1", "Sheet2", "Sheet3"]);

  await rename(page, "Sheet2", "Budget");
  await expect(tabs(page)).toHaveText(["Sheet1", "Budget", "Sheet3"]);

  await tabMenu(page, "Budget", "duplicate");
  await expect(tabs(page)).toHaveText(["Sheet1", "Budget", "Budget (2)", "Sheet3"]);
  await expect(tab(page, "Budget (2)")).toHaveAttribute("aria-selected", "true");

  // Move it first by keyboard: Shift+F10 on the focused tab opens its menu on
  // "Insert sheet"; four steps down is "Move left".
  for (let step = 0; step < 2; step += 1) {
    await tab(page, "Budget (2)").focus();
    await page.keyboard.press("Shift+F10");
    await expect(page.getByTestId("sheet-tab-menu")).toBeVisible();
    for (let down = 0; down < 4; down += 1) {
      await page.keyboard.press("ArrowDown");
    }
    await expect(page.getByTestId("sheet-tab-menu-move-left")).toBeFocused();
    await page.keyboard.press("Enter");
  }
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget", "Sheet3"]);
  await expect(page.getByTestId("sheet-tabs-announcement")).toHaveText('Sheet "Budget (2)" moved to position 1');

  await tabMenu(page, "Sheet3", "hide");
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget"]);
  await tabMenu(page, "Budget", "unhide");
  const dialog = page.getByTestId("sheet-unhide-dialog");
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId("sheet-unhide-option")).toHaveText(["Sheet3"]);
  await dialog.getByTestId("sheet-unhide-option").click();
  await dialog.getByTestId("sheet-unhide-confirm").click();
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget", "Sheet3"]);
  await expect(tab(page, "Sheet3")).toHaveAttribute("aria-selected", "true");

  await tabMenu(page, "Budget", "tab-color");
  await expect(page.getByTestId("sheet-tab-color-picker")).toBeVisible();
  await page.getByTestId("text-color-swatch").first().click();
  await expect(tab(page, "Budget")).toHaveAttribute("data-tab-color", /^#/);

  // Sheet1 is empty: it goes without a confirmation; undo brings it back.
  await tabMenu(page, "Sheet1", "delete");
  await expect(tabs(page)).toHaveText(["Budget (2)", "Budget", "Sheet3"]);
  await page.getByTestId("spreadsheet-undo").click();
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget", "Sheet3"]);

  await downloadAndReopen(page, () => page.getByTestId("spreadsheet-download").click(), "blank.xlsx");
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget", "Sheet3"]);
  await expect(tab(page, "Budget")).toHaveAttribute("data-tab-color", /^#/);

  // spec121-links.xlsx: references follow a rename and break on a delete.
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(fixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await rename(page, "Data", "Facts");
  await tab(page, "Summary").click();
  await cell(page, "C2").click();
  await expect(page.getByTestId("formula-bar-content")).toHaveText("=Facts!B2*2");

  await tabMenu(page, "Lookup", "delete");
  const confirm = page.getByTestId("sheet-tab-delete-dialog");
  await expect(confirm).toContainText('Delete "Lookup"? Its cells and anything on it will be removed.');
  await page.getByTestId("sheet-tab-delete-confirm").click();
  await expect(tab(page, "Lookup")).toHaveCount(0);
  await tab(page, "Summary").click();
  await expect(cell(page, "E2")).toHaveText("#REF!");

  await downloadAndReopen(page, () => page.getByTestId("spreadsheet-download").click(), "links.xlsx");
  await expect(tabs(page)).toHaveText(["Summary", "Facts"]);
  await cell(page, "C2").click();
  await expect(page.getByTestId("formula-bar-content")).toHaveText("=Facts!B2*2");
  await expect(cell(page, "E2")).toHaveText("#REF!");

  // A copy of the blank workbook as .ods re-opens with the same tab names.
  await page.goto(blankUrl);
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget", "Sheet3"]);
  await downloadAndReopen(
    page,
    async () => {
      await page.getByTestId("spreadsheet-download-menu").click();
      await expect(page.getByTestId("spreadsheet-download-copy")).toHaveText("Download a copy as OpenDocument Spreadsheet (.ods)…");
      await page.getByTestId("spreadsheet-download-copy").click();
    },
    "blank.ods",
  );
  await expect(page.getByTestId("spreadsheet-format")).toHaveText("OpenDocument Spreadsheet (.ods)");
  await expect(tabs(page)).toHaveText(["Budget (2)", "Sheet1", "Budget", "Sheet3"]);
});
