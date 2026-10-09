import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

// SPEC-117 — the sheet grid's styles, sizes, merges, frozen panes and
// selection, end to end (AC-58..62): open spec117-styles.xlsx (Excel-shaped)
// → B2 renders bold red on yellow → C5 is one element as wide as C5:E5 → the
// header boundary between F and H carries a "Hidden column G" marker → the
// frozen rows stay put while 200 rows scroll → the file's 85% zoom → Ctrl/Cmd
// +Down from B3 lands on B12 → B3 + Shift+click B5 aggregates 10, 20, 30 →
// Ctrl/Cmd+A twice selects the whole sheet → Gridlines off survives a reload
// with the selection and the zoom → dark mode leaves the paper's colours
// alone → Download is byte-identical → the same open-and-check for
// spec117-styles.ods (LibreOffice-shaped).

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const xlsxFixture = join(fixturesDir, "spec117-styles.xlsx");
const odsFixture = join(fixturesDir, "spec117-styles.ods");

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ODS_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

async function dropFile(page: Page, name: string, mime: string, path: string) {
  const transfer = await page.evaluateHandle(
    (item) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(item.base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], item.name, { type: item.mime }));
      return dt;
    },
    { name, mime, base64: readFileSync(path).toString("base64") },
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
}

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

function columnHeader(page: Page, letters: string): Locator {
  return page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) });
}

function rowHeader(page: Page, row: number): Locator {
  return page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${row}$`) });
}

async function box(locator: Locator): Promise<{ x: number; y: number; width: number; height: number }> {
  const rect = await locator.boundingBox();
  expect(rect, "element has a box").not.toBeNull();
  return rect!;
}

/** B2: bold red text on a yellow fill — the computed styles of its cell box
 * and its text. */
async function expectBoldRedOnYellow(page: Page): Promise<void> {
  const b2 = cell(page, "B2");
  await expect(b2).toHaveCount(1);
  await expect(b2).toHaveCSS("background-color", "rgb(255, 255, 0)");
  const text = b2.locator(".sheet-cell-text");
  await expect(text).toHaveCSS("color", "rgb(255, 0, 0)");
  await expect(text).toHaveCSS("font-weight", "700");
}

/** C5 is one element exactly as wide as columns C, D and E together, and
 * D5/E5 render nothing of their own. */
async function expectMergedTitle(page: Page): Promise<void> {
  const c5 = cell(page, "C5");
  await expect(c5).toHaveCount(1);
  await expect(c5).toHaveAttribute("aria-colspan", "3");
  await expect(cell(page, "D5")).toHaveCount(0);
  await expect(cell(page, "E5")).toHaveCount(0);
  const widths = await Promise.all(["C", "D", "E"].map(async (letters) => (await box(columnHeader(page, letters))).width));
  const merged = await box(c5);
  expect(Math.abs(merged.width - widths.reduce((a, b) => a + b, 0))).toBeLessThan(1.5);
}

/** The "Hidden column G" marker sits on the header boundary between F and H. */
async function expectHiddenColumnG(page: Page): Promise<void> {
  await expect(columnHeader(page, "G")).toHaveCount(0);
  const marker = page.locator('[data-testid="sheet-hidden-marker"][title="Hidden column G"]');
  await expect(marker).toHaveCount(1);
  const f = await box(columnHeader(page, "F"));
  const h = await box(columnHeader(page, "H"));
  const m = await box(marker);
  const centre = m.x + m.width / 2;
  expect(centre).toBeGreaterThanOrEqual(f.x + f.width - 4);
  expect(centre).toBeLessThanOrEqual(h.x + 4);
}

/** Scrolls the grid down `rows` default-height rows (at the current zoom). */
async function scrollDownRows(page: Page, rows: number, rowPx: number): Promise<void> {
  await page.getByTestId("spreadsheet-grid").evaluate((element, by) => {
    element.scrollTop += by;
    element.dispatchEvent(new Event("scroll"));
  }, rows * rowPx);
}

/** Rows 1–2 are frozen: still on screen, under the column letters, after the
 * scroll; row 3 is not, and a row 200 rows on is. */
async function expectFrozenRowsStay(page: Page, rowPx: number): Promise<void> {
  const grid = await box(page.getByTestId("spreadsheet-grid"));
  await scrollDownRows(page, 200, rowPx);
  await expect(rowHeader(page, 210)).toBeVisible();
  for (const row of [1, 2]) {
    const header = await box(rowHeader(page, row));
    expect(header.y).toBeGreaterThanOrEqual(grid.y);
    expect(header.y).toBeLessThan(grid.y + 80);
  }
  await expect(cell(page, "A1")).toBeVisible();
  await expect(page.getByTestId("sheet-freeze-line")).toHaveCount(2);
  await page.getByTestId("spreadsheet-grid").evaluate((element) => {
    element.scrollTop = 0;
    element.dispatchEvent(new Event("scroll"));
  });
}

/** Whether the workbook's stored shell has reached IndexedDB with sheet 0's
 * gridlines off — the view-state write runs at most once a second. */
async function storedGridLinesOff(page: Page, id: string): Promise<boolean> {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const record = await new Promise<{ model?: { sheets?: Array<{ showGridLines?: boolean }> } } | undefined>((resolve, reject) => {
      const request = db.transaction("documents").objectStore("documents").get(docId);
      request.onsuccess = () => resolve(request.result as { model?: { sheets?: Array<{ showGridLines?: boolean }> } } | undefined);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return record?.model?.sheets?.[0]?.showGridLines === false;
  }, id);
}

/** Every selected-state check of "the whole sheet is selected". */
async function expectWholeSheetSelected(page: Page): Promise<void> {
  for (const letters of ["A", "B", "F", "H"]) {
    await expect(columnHeader(page, letters)).toHaveClass(/is-selected/);
  }
  for (const row of [1, 3, 20]) {
    await expect(rowHeader(page, row)).toHaveClass(/is-selected/);
  }
  await expect(cell(page, "A1")).toHaveAttribute("aria-selected", "true");
  await expect(cell(page, "H20")).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("sheet-selection")).toHaveCount(1);
}

async function openView(page: Page): Promise<void> {
  await page.getByRole("menubar").getByText("View", { exact: true }).click();
}

test("spec117-styles.xlsx renders, navigates, remembers its view and downloads byte-identical (AC-58..61)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec117-styles.xlsx", XLSX_TYPE, xlsxFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const id = new URL(page.url()).pathname.split("/").pop()!;
  const grid = page.getByTestId("spreadsheet-grid");
  await expect(grid).toHaveAttribute("aria-label", "Styled");

  // Styles, the merge and the hidden column.
  await expectBoldRedOnYellow(page);
  await expectMergedTitle(page);
  await expectHiddenColumnG(page);

  // The file's zoom, and the frozen rows 1–2 (rows are 20 px at 100%).
  await expect(page.getByTestId("zoom-percent")).toHaveValue("85%");
  await expectFrozenRowsStay(page, 20 * 0.85);

  // The file's selection opens restored: B3 active.
  await expect(page.getByTestId("name-box")).toHaveValue("B3");
  await expect(grid).toHaveAttribute("aria-activedescendant", "sheet-cell-B3");

  // Ctrl/Cmd+Down from B3 runs to the end of the data region: B12.
  await cell(page, "B3").click();
  await page.keyboard.press("ControlOrMeta+ArrowDown");
  await expect(page.getByTestId("name-box")).toHaveValue("B12");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("100");

  // B3, Shift+click B5: 10, 20, 30.
  await cell(page, "B3").click();
  await cell(page, "B5").click({ modifiers: ["Shift"] });
  await expect(page.getByTestId("sheet-status-average")).toHaveText("Average: 20");
  await expect(page.getByTestId("sheet-status-count")).toHaveText("Count: 3");
  await expect(page.getByTestId("sheet-status-sum")).toHaveText("Sum: 60");
  await expect(cell(page, "B4")).toHaveAttribute("aria-selected", "true");

  // Ctrl/Cmd+A: the current region, then the whole sheet.
  await page.keyboard.press("ControlOrMeta+a");
  await expect(columnHeader(page, "H")).not.toHaveClass(/is-selected/);
  await page.keyboard.press("ControlOrMeta+a");
  await expectWholeSheetSelected(page);

  // View → Gridlines off.
  await expect(page.locator(".sheet-gridline").first()).toBeAttached();
  await openView(page);
  const gridlines = page.getByTestId("sheet-menu-gridlines");
  await expect(gridlines).toHaveAttribute("role", "menuitemcheckbox");
  await expect(gridlines).toHaveAttribute("aria-checked", "true");
  await gridlines.click();
  await expect(page.locator(".sheet-gridline")).toHaveCount(0);

  // Reload: the selection, the zoom and the gridlines are as left.
  await expect.poll(() => storedGridLinesOff(page, id)).toBe(true);
  await page.reload();
  await expect(grid).toHaveAttribute("aria-label", "Styled");
  await expect(page.getByTestId("zoom-percent")).toHaveValue("85%");
  await expect(page.locator(".sheet-gridline")).toHaveCount(0);
  await expect(page.getByTestId("name-box")).toHaveValue("B3");
  await expectWholeSheetSelected(page);
  await openView(page);
  await expect(page.getByTestId("sheet-menu-gridlines")).toHaveAttribute("aria-checked", "false");
  await page.keyboard.press("Escape");

  // Dark mode: the chrome follows the app theme, the paper keeps the file's
  // colours.
  await page.evaluate(() => {
    const key = "docu-edit:preferences";
    const prefs = JSON.parse(localStorage.getItem(key) ?? "{}") as Record<string, unknown>;
    localStorage.setItem(key, JSON.stringify({ ...prefs, theme: "dark" }));
  });
  await page.reload();
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  await expect(grid).toHaveAttribute("aria-label", "Styled");
  await expect(cell(page, "B2")).toHaveCSS("background-color", "rgb(255, 255, 0)");
  await expect(page.locator(".sheet-paper")).toHaveCSS("background-color", "rgb(255, 255, 255)");

  // Download: every part byte-identical to the fixture's, after all of it.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("spec117-styles.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(await download.path())));
  const source = await unzip(new Uint8Array(readFileSync(xlsxFixture)));
  expect([...saved.keys()]).toEqual([...source.keys()]);
  for (const [name, bytes] of source) {
    expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
  }
});

test("spec117-styles.ods opens with its frozen pane, merge, hidden column and file zoom (AC-62)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec117-styles.ods", ODS_TYPE, odsFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-label", "Styled");

  await expectBoldRedOnYellow(page);
  await expectMergedTitle(page);
  await expectHiddenColumnG(page);
  await expect(page.getByTestId("zoom-percent")).toHaveValue("85%");
  // LibreOffice's default rows are 0.178 in (17 px at 100%).
  await expectFrozenRowsStay(page, 17 * 0.85);
  await expect(page.getByTestId("name-box")).toHaveValue("B3");

  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const download = await downloadPromise;
  const saved = await unzip(new Uint8Array(readFileSync(await download.path())));
  const source = await unzip(new Uint8Array(readFileSync(odsFixture)));
  expect([...saved.keys()]).toEqual([...source.keys()]);
  for (const [name, bytes] of source) {
    expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
  }
});
