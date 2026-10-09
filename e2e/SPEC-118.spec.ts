import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

// SPEC-118 — number formats and dates, end to end (AC-38..41): open
// spec118-formats.xlsx (Excel-shaped) → A2 (built-in 14) "10/8/2026", A3 (22)
// "10/8/2026 13:45", A4 ([$-F800]) "Thursday, October 8, 2026", B2
// "(1,234.50)" in red, C2 "12.50%", D2 "36:00:00" → the narrow F shows
// "1.23E+11" and the too-narrow date in G only # → A4 in the formula bar is
// "10/8/2026 1:45:00 PM" → choosing sheet-locale-de-DE re-renders at once
// (A2 "08.10.2026", A4 "Donnerstag, 8. Oktober 2026", B5 "1.234,50", the bar
// "08.10.2026 13:45:00") without touching the stored record → a reload keeps
// de-DE → the download is byte-identical → spec118-formats.ods renders
// LibreOffice's styles ("-1,00" in red, "3 14/99", "Q4 2026", the unstyled
// 1899-12-25 date as "1899-12-25").

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const xlsxFixture = join(fixturesDir, "spec118-formats.xlsx");
const odsFixture = join(fixturesDir, "spec118-formats.ods");

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

/** The stored record's lastModified (undefined until the idle write lands). */
async function storedLastModified(page: Page, id: string): Promise<number | undefined> {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const record = await new Promise<{ lastModified?: number } | undefined>((resolve, reject) => {
      const request = db.transaction("documents").objectStore("documents").get(docId);
      request.onsuccess = () => resolve(request.result as { lastModified?: number } | undefined);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return record?.lastModified;
  }, id);
}

test("spec118-formats.xlsx renders its number formats, follows the locale preference and downloads byte-identical; spec118-formats.ods renders LibreOffice's styles (AC-38..41)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec118-formats.xlsx", XLSX_TYPE, xlsxFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const id = new URL(page.url()).pathname.split("/").pop()!;
  const grid = page.getByTestId("spreadsheet-grid");
  await expect(grid).toHaveAttribute("aria-label", "Formats");

  // AC-38: the built-ins, the system long date, the red negative, the
  // percentage and the elapsed hours, as en-US Excel shows them.
  await expect(page.getByTestId("sheet-locale")).toHaveText("en-US");
  await expect(cell(page, "A2")).toHaveText("10/8/2026");
  await expect(cell(page, "A3")).toHaveText("10/8/2026 13:45");
  await expect(cell(page, "A4")).toHaveText("Thursday, October 8, 2026");
  await expect(cell(page, "B2")).toHaveText("(1,234.50)");
  await expect(cell(page, "B2").locator(".sheet-cell-text")).toHaveCSS("color", "rgb(255, 0, 0)");
  await expect(cell(page, "C2")).toHaveText("12.50%");
  await expect(cell(page, "D2")).toHaveText("36:00:00");

  // AC-39: General shrinks to the narrow F; a date never shrinks — G is #.
  await expect(cell(page, "F2")).toHaveText("1.23E+11");
  await expect(cell(page, "G2")).toHaveText(/^#+$/);
  await cell(page, "A4").click();
  await expect(page.getByTestId("name-box")).toHaveValue("A4");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("10/8/2026 1:45:00 PM");

  // AC-40: the locale is a preference — it re-renders at once and touches
  // nothing stored.
  await expect.poll(() => storedLastModified(page, id)).not.toBeUndefined();
  const lastModified = await storedLastModified(page, id);
  await page.getByTestId("sheet-locale").click();
  const german = page.getByTestId("sheet-locale-de-DE");
  await expect(german).toHaveAttribute("role", "menuitemradio");
  await expect(german).toHaveAttribute("aria-checked", "false");
  await german.click();
  await expect(page.getByTestId("sheet-locale")).toHaveText("de-DE");
  await expect(cell(page, "A2")).toHaveText("08.10.2026");
  await expect(cell(page, "A4")).toHaveText("Donnerstag, 8. Oktober 2026");
  await expect(cell(page, "B5")).toHaveText("1.234,50");
  await expect(page.getByTestId("formula-bar-content")).toHaveText("08.10.2026 13:45:00");
  expect(await storedLastModified(page, id)).toBe(lastModified);

  // A reload keeps de-DE.
  await page.reload();
  await expect(grid).toHaveAttribute("aria-label", "Formats");
  await expect(page.getByTestId("sheet-locale")).toHaveText("de-DE");
  await expect(cell(page, "A2")).toHaveText("08.10.2026");
  await expect(cell(page, "A4")).toHaveText("Donnerstag, 8. Oktober 2026");

  // Download: every part byte-identical to the fixture's.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("spec118-formats.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(await download.path())));
  const source = await unzip(new Uint8Array(readFileSync(xlsxFixture)));
  expect([...saved.keys()]).toEqual([...source.keys()]);
  for (const [name, bytes] of source) {
    expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
  }

  // AC-41: LibreOffice's styles, still in de-DE.
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec118-formats.ods", ODS_TYPE, odsFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(grid).toHaveAttribute("aria-label", "Formats");
  await expect(cell(page, "B2")).toHaveText("-1,00");
  await expect(cell(page, "B2").locator(".sheet-cell-text")).toHaveCSS("color", "rgb(255, 0, 0)");
  await expect(cell(page, "B6")).toHaveText("3 14/99");
  await expect(cell(page, "B11")).toHaveText("Q4 2026");
  await expect(cell(page, "B14")).toHaveText("1899-12-25");
});
