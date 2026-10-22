import { test, expect, type Locator, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { SPEC128_FILTER_XLSX } from "../src/lib/spreadsheet/fixtures/build";

// SPEC-128 — sort and filter (AC-42, AC-43), with the real pointer and
// keyboard: open spec128-filter.xlsx → the header row shows filter buttons
// and the status bar reads "8 of 12 records found" → open Region's popover →
// uncheck "West" → OK → "6 of 12 records found" and the West rows are gone →
// Data ▸ Clear filters → the user-hidden row is still hidden → click an
// Amount cell → Data ▸ Sort Z to A → the visible Amount column reads
// descending → Ctrl/Cmd+Z restores the order → Ctrl/Cmd+Shift+L → no filter
// button remains → download → reopen → no autofilter. Then
// spec128-filter.ods: its table:visibility="filter" rows open hidden and
// Data ▸ Reapply filter leaves the same rows hidden.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const SALES = SPEC128_FILTER_XLSX.sales;

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

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await colHeader(page, letters!).boundingBox();
  const row = await rowHeader(page, Number(digits)).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
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
  const path = join(tmpdir(), `spec128-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

async function dataMenu(page: Page, item: string): Promise<void> {
  await page.getByTestId("menu-title-data").click();
  await page.getByTestId(item).click();
}

/** The Amount column's shown cells, top to bottom, over the data rows. */
async function visibleAmounts(page: Page): Promise<string[]> {
  const out: string[] = [];
  for (let row = 2; row <= 13; row += 1) {
    if (await rowHeader(page, row).isVisible()) {
      out.push((await cell(page, `C${row}`).innerText()).trim());
    }
  }
  return out;
}

const texts = (values: readonly number[]): string[] => values.map(String);

/** Rows hidden either way: every data row whose header does not show. */
async function hiddenDataRows(page: Page): Promise<number[]> {
  const out: number[] = [];
  for (let row = 2; row <= 13; row += 1) {
    if (!(await rowHeader(page, row).isVisible())) {
      out.push(row);
    }
  }
  return out;
}

test("SPEC-128: filter, clear, sort, undo and toggle an Excel autofilter; reapply a LibreOffice one", async ({ page }) => {
  await open(page, join(fixtures, "spec128-filter.xlsx"));

  // The header row carries a filter button per column; the status bar counts records.
  const buttons = page.getByTestId("filter-button");
  await expect(buttons).toHaveCount(4);
  const region = page.getByRole("button", { name: 'Filter "Region"' });
  await expect(region).toHaveAttribute("aria-haspopup", "dialog");
  await expect(region).toHaveAttribute("aria-expanded", "false");
  await expect(region).toHaveAttribute("data-state", "filtered");
  await expect(page.getByTestId("sheet-status-filter")).toHaveText(SALES.status);
  await expect.poll(() => hiddenDataRows(page)).toEqual([...SALES.filterHiddenRows, SALES.userHiddenRow].toSorted((a, b) => a - b));
  await expect(rowHeader(page, 2)).toHaveAttribute("data-filtered", "true");

  // Region's popover: uncheck West → OK.
  await region.click();
  const popover = page.getByTestId("filter-popover");
  await expect(popover).toBeVisible();
  await expect(region).toHaveAttribute("aria-expanded", "true");
  await popover.getByTestId("filter-value").filter({ hasText: /^West$/ }).click();
  await popover.getByTestId("filter-ok").click();
  await expect(popover).toBeHidden();
  await expect(page.getByTestId("sheet-status-filter")).toHaveText(SALES.statusWithoutWest);
  await expect(rowHeader(page, 3)).toBeHidden();
  await expect(rowHeader(page, 8)).toBeHidden();

  // Data ▸ Clear filters: every filter-hidden row shows; the user-hidden one stays hidden.
  await dataMenu(page, "sheet-menu-clear-filters");
  await expect(page.getByTestId("sheet-status-filter")).toHaveCount(0);
  await expect.poll(() => hiddenDataRows(page)).toEqual([SALES.userHiddenRow]);
  await expect(buttons).toHaveCount(4);
  await expect.poll(() => visibleAmounts(page)).toEqual(texts(SALES.visibleAmountsCleared));

  // Click an Amount cell → Data ▸ Sort Z to A: the visible Amount column reads descending.
  await clickCell(page, "C6");
  await dataMenu(page, "sheet-menu-sort-desc");
  await expect.poll(() => visibleAmounts(page)).toEqual(texts(SALES.visibleAmountsDescending));
  await expect.poll(() => hiddenDataRows(page)).toEqual([SALES.userHiddenRow]);
  await expect(page.getByRole("button", { name: 'Filter "Amount"' })).toHaveAttribute("data-state", "sorted-desc");

  // Ctrl/Cmd+Z restores the order (one undo step labelled "Sort").
  await page.getByTestId("spreadsheet-grid").focus();
  await page.keyboard.press("ControlOrMeta+z");
  await expect.poll(() => visibleAmounts(page)).toEqual(texts(SALES.visibleAmountsCleared));

  // Ctrl/Cmd+Shift+L turns the filter off: no button remains.
  await page.getByTestId("spreadsheet-grid").focus();
  await page.keyboard.press("ControlOrMeta+Shift+L");
  await expect(buttons).toHaveCount(0);
  await expect.poll(() => hiddenDataRows(page)).toEqual([SALES.userHiddenRow]);

  // Download → reopen: no autofilter.
  const saved = await download(page, "spec128-filter.xlsx");
  await open(page, saved);
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-label", "Sales");
  await expect(cell(page, "A1")).toHaveText("Region");
  await expect(page.getByTestId("filter-button")).toHaveCount(0);
  await expect(page.getByTestId("sheet-status-filter")).toHaveCount(0);
  await page.getByTestId("menu-title-data").click();
  await expect(page.getByTestId("sheet-menu-filter")).toHaveAttribute("aria-checked", "false");
  await page.keyboard.press("Escape");

  // spec128-filter.ods: the table:visibility="filter" rows (and the collapsed one) open hidden…
  await open(page, join(fixtures, "spec128-filter.ods"));
  await expect(page.getByTestId("filter-button")).toHaveCount(4);
  const expectedHidden = [...SALES.filterHiddenRows, SALES.userHiddenRow].toSorted((a, b) => a - b);
  await expect.poll(() => hiddenDataRows(page)).toEqual(expectedHidden);
  await expect(page.getByTestId("sheet-status-filter")).toHaveText(SALES.status);

  // …and Data ▸ Reapply filter leaves the same rows hidden.
  await dataMenu(page, "sheet-menu-reapply-filter");
  await expect(page.getByTestId("sheet-status-filter")).toHaveText(SALES.status);
  await expect.poll(() => hiddenDataRows(page)).toEqual(expectedHidden);
});
