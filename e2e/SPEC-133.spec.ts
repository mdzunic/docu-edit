import { test, expect, type Locator, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-133 — named ranges and tables (AC-61–AC-63), with the real pointer and
// keyboard: open spec133-names-tables.xlsx → SUM(Sales[Amount]) shows its
// value → change an amount → it recalculates → type Margin in the Name Box
// over B2:B5 → =SUM(Margin) evaluates → the Name Manager lists Margin and
// TaxRate but not Print_Area → (totals off) type a value under the table → it
// expands with banding → type =[@Price]*2 in the new column → the column
// fills → toggle the totals row → pick Average → rename the table to Orders →
// the formulas read Orders[...] → Insert ▸ Table on a fresh block → a Table2
// appears styled → download, reopen → the names, both tables and the
// formulas persist; download a copy as .ods, reopen → the values are
// identical, there are no tables, the names are there.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string, modifiers: ("Shift" | "ControlOrMeta")[] = []): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  for (const key of modifiers) {
    await page.keyboard.down(key);
  }
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  for (const key of modifiers) {
    await page.keyboard.up(key);
  }
}

/** Types `text` into a cell and commits it with Enter. */
async function enter(page: Page, ref: string, text: string): Promise<void> {
  await clickCell(page, ref);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
  await page.keyboard.type(text);
  await page.keyboard.press("Enter");
}

async function formulaOf(page: Page, ref: string): Promise<string> {
  await clickCell(page, ref);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
  return (await page.getByTestId("formula-bar-content").textContent()) ?? "";
}

/** A cell's painted background colour. */
const background = (target: Locator): Promise<string> => target.evaluate((el) => getComputedStyle(el).backgroundColor);

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
}

async function downloadAndReopen(page: Page, trigger: () => Promise<void>, fileName: string): Promise<void> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const saved = await downloadPromise;
  const path = join(tmpdir(), `spec133-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  await open(page, path);
}

async function nameManagerNames(page: Page): Promise<string[]> {
  await page.getByTestId("sheet-menu-formulas").click();
  await page.getByTestId("sheet-menu-name-manager").click();
  const manager = page.getByTestId("name-manager");
  await expect(manager).toBeVisible();
  const names = await manager.getByTestId("name-manager-row").locator("[data-testid='name-manager-row-name']").allTextContents();
  await manager.getByTestId("name-manager-close").click();
  await expect(manager).toBeHidden();
  return names;
}

test("names, tables and both formats (AC-61–AC-63)", async ({ page }) => {
  await open(page, join(fixtures, "spec133-names-tables.xlsx"));

  // AC-61: SUM(Sales[Amount]) shows its value; changing an amount recalculates it.
  await expect(cell(page, "G2")).toHaveText("140");
  await enter(page, "D2", "100");
  await expect(cell(page, "G2")).toHaveText("220");

  // Margin over B2:B5 through the Name Box, then =SUM(Margin).
  await clickCell(page, "B2");
  await clickCell(page, "B5", ["Shift"]);
  const nameBox = page.getByTestId("name-box");
  await nameBox.click();
  await nameBox.fill("Margin");
  await nameBox.press("Enter");
  await enter(page, "G9", "=SUM(Margin)");
  await expect(cell(page, "G9")).toHaveText("18");

  // The Name Manager lists Margin and TaxRate, never the reserved Print_Area.
  const listed = await nameManagerNames(page);
  expect(listed).toContain("Margin");
  expect(listed).toContain("TaxRate");
  expect(listed.some((name) => /print_area/i.test(name))).toBe(false);

  // AC-62: with the totals row off, a value typed under the table expands it.
  await clickCell(page, "B3");
  const design = page.getByTestId("sheet-table-design");
  await expect(design).toBeVisible();
  await expect(page.getByTestId("sheet-table-name")).toHaveValue("Sales");
  await page.getByTestId("sheet-table-option-totals-row").click();
  await expect(cell(page, "A6")).toHaveText("");
  await enter(page, "A6", "Kiwis");
  await clickCell(page, "A6");
  await expect(design).toBeVisible();
  // Banding: the new fifth data row is striped like the third, unlike the fourth.
  expect(await background(cell(page, "A6"))).toBe(await background(cell(page, "A4")));
  expect(await background(cell(page, "A6"))).not.toBe(await background(cell(page, "A5")));

  // =[@Price]*2 right of the last column grows a column and fills it.
  await enter(page, "F2", "=[@Price]*2");
  await expect(cell(page, "F2")).toHaveText("4");
  await expect(cell(page, "F3")).toHaveText("6");
  await expect(cell(page, "F5")).toHaveText("18");
  expect(await formulaOf(page, "F4")).toBe("=Sales[@Price]*2");

  // Totals back on, then Average for Amount.
  await clickCell(page, "B3");
  await page.getByTestId("sheet-table-option-totals-row").click();
  await expect(cell(page, "A7")).toHaveText("Total");
  await clickCell(page, "D7");
  await page.getByTestId("sheet-table-totals-menu-button").click();
  await page.getByTestId("sheet-table-totals-average").click();
  expect(await formulaOf(page, "D7")).toBe("=SUBTOTAL(101,Sales[Amount])");

  // Rename the table: every formula follows.
  await clickCell(page, "B3");
  const tableName = page.getByTestId("sheet-table-name");
  await tableName.fill("Orders");
  await tableName.press("Enter");
  expect(await formulaOf(page, "G2")).toBe("=SUM(Orders[Amount])");
  expect(await formulaOf(page, "F3")).toBe("=Orders[@Price]*2");
  expect(await formulaOf(page, "D7")).toBe("=SUBTOTAL(101,Orders[Amount])");

  // AC-63: Insert ▸ Table on a fresh block.
  await enter(page, "I12", "Name");
  await enter(page, "J12", "Score");
  await enter(page, "I13", "Ann");
  await enter(page, "J13", "7");
  await enter(page, "I14", "Bob");
  await enter(page, "J14", "9");
  await clickCell(page, "I12");
  await clickCell(page, "J14", ["Shift"]);
  await page.getByTestId("sheet-menu-insert").click();
  await page.getByTestId("sheet-menu-insert-table").click();
  const create = page.getByTestId("sheet-create-table");
  await expect(create).toBeVisible();
  await expect(page.getByTestId("sheet-create-table-range")).toHaveValue(/I\$?12:\$?J\$?14/);
  await expect(page.getByTestId("sheet-create-table-headers")).toBeChecked();
  await page.getByTestId("sheet-create-table-ok").click();
  await expect(create).toBeHidden();
  await clickCell(page, "J13");
  await expect(page.getByTestId("sheet-table-name")).toHaveValue("Table2");
  // Styled: TableStyleMedium2's header is filled.
  expect(await background(cell(page, "I12"))).not.toBe(await background(cell(page, "L12")));

  // Download, reopen: names, both tables and formulas persist.
  await downloadAndReopen(page, () => page.getByTestId("spreadsheet-download").click(), "names-tables.xlsx");
  await expect(cell(page, "G9")).toHaveText("18");
  expect(await formulaOf(page, "G2")).toBe("=SUM(Orders[Amount])");
  expect(await formulaOf(page, "F3")).toBe("=Orders[@Price]*2");
  await expect(page.getByTestId("sheet-table-name")).toHaveValue("Orders");
  await clickCell(page, "J13");
  await expect(page.getByTestId("sheet-table-name")).toHaveValue("Table2");
  const names = await nameManagerNames(page);
  expect(names).toContain("Margin");
  expect(names).toContain("TaxRate");
  const values = await Promise.all(["G2", "G3", "G9", "D7", "F3", "J14"].map((ref) => cell(page, ref).textContent()));

  // A copy as .ods: the values are identical, there are no tables, the names are there.
  await downloadAndReopen(
    page,
    async () => {
      await page.getByTestId("spreadsheet-download-menu").click();
      await page.getByTestId("spreadsheet-download-copy").click();
    },
    "names-tables.ods",
  );
  await expect(page.getByTestId("spreadsheet-format")).toHaveText("OpenDocument Spreadsheet (.ods)");
  expect(await Promise.all(["G2", "G3", "G9", "D7", "F3", "J14"].map((ref) => cell(page, ref).textContent()))).toEqual(values);
  await clickCell(page, "B3");
  await expect(page.getByTestId("sheet-table-design")).toBeHidden();
  expect(await formulaOf(page, "G2")).toBe("=SUM($D$2:$D$6)");
  const odsNames = await nameManagerNames(page);
  expect(odsNames).toContain("Margin");
  expect(odsNames).toContain("TaxRate");
});
