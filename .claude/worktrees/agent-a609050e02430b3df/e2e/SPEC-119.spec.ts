import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

// SPEC-119 — the formula language, end to end (AC-70): drop
// spec119-formulas.xlsx (Excel-shaped) → B5's formula bar reads
// "=SUM(B2:B4)" → C3, a shared-formula follower, "=B3*$F$1" → E3, a CSE
// array member, "{=A2:A4*2}" → G2 `=IFS(A2>1,"big",TRUE,"small")` (no
// _xlfn.) → Ctrl+` shows formulas: B5's cell reads "=SUM(B2:B4)" and View ▸
// Show formulas is checked → Ctrl+` again: B5 reads "60" → Close → drop
// spec119-formulas.ods (LibreOffice-shaped) → B5 "=SUM(B2:B4)", B7
// "='My Sheet'!A1+Sheet2!B1" → the download is part-for-part the fixture.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const xlsxFixture = join(fixturesDir, "spec119-formulas.xlsx");
const odsFixture = join(fixturesDir, "spec119-formulas.ods");

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

/** Selects a cell and returns the formula bar's content. */
async function select(page: Page, ref: string): Promise<Locator> {
  await cell(page, ref).click();
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
  return page.getByTestId("formula-bar-content");
}

test("spec119-formulas.xlsx shows its formulas in the bar and, with Ctrl+`, in the grid; spec119-formulas.ods shows OpenFormula as canonical text and downloads byte-identical (AC-70)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec119-formulas.xlsx", XLSX_TYPE, xlsxFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const grid = page.getByTestId("spreadsheet-grid");
  await expect(grid).toHaveAttribute("aria-label", "Calc");

  // The formula bar: a formula, a shared follower, an array member and an
  // Excel future function, each in canonical (UI) form.
  await expect(await select(page, "B5")).toHaveText("=SUM(B2:B4)");
  await expect(cell(page, "B5")).toHaveText("60");
  await expect(await select(page, "C3")).toHaveText("=B3*$F$1");
  await expect(await select(page, "E3")).toHaveText("{=A2:A4*2}");
  const ifs = await select(page, "G2");
  await expect(ifs).toHaveText('=IFS(A2>1,"big",TRUE,"small")');
  await expect(ifs).not.toContainText("_xlfn.");
  // SPEC-122 made the bar editable: a textbox holding the cell editor's draft.
  await expect(ifs).toHaveAttribute("contenteditable", "true");

  // Ctrl+`: formulas in the grid, and View ▸ Show formulas checked.
  await page.keyboard.press("Control+Backquote");
  await expect(grid).toHaveAttribute("data-show-formulas", "true");
  await expect(cell(page, "B5")).toHaveText("=SUM(B2:B4)");
  await page.getByRole("menubar").getByText("View", { exact: true }).click();
  const showFormulas = page.getByTestId("sheet-menu-show-formulas");
  await expect(showFormulas).toHaveAttribute("role", "menuitemcheckbox");
  await expect(showFormulas).toHaveAttribute("aria-checked", "true");
  await page.keyboard.press("Escape");

  // Ctrl+` again: the values are back.
  await page.keyboard.press("Control+Backquote");
  await expect(grid).not.toHaveAttribute("data-show-formulas", "true");
  await expect(cell(page, "B5")).toHaveText("60");

  // Close, then the LibreOffice-shaped workbook.
  await page.getByTestId("close-document").click();
  await expect(page).toHaveURL(/\/$/);
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec119-formulas.ods", ODS_TYPE, odsFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(grid).toHaveAttribute("aria-label", "Calc");
  await expect(await select(page, "B5")).toHaveText("=SUM(B2:B4)");
  await expect(await select(page, "B7")).toHaveText("='My Sheet'!A1+Sheet2!B1");

  // Download: every part byte-identical to the fixture's.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("spec119-formulas.ods");
  const saved = await unzip(new Uint8Array(readFileSync(await download.path())));
  const source = await unzip(new Uint8Array(readFileSync(odsFixture)));
  expect([...saved.keys()]).toEqual([...source.keys()]);
  for (const [name, bytes] of source) {
    expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
  }
});
