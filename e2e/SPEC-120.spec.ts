import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

// SPEC-120 — the calculation engine, end to end (AC-43, AC-44), with the
// clock pinned before navigation to 2026-10-08 09:30: drop spec120-calc.xlsx
// → B2 (=SUM(A1:A3), cached 999) reads "999" — the cache is trusted → D2
// (=TODAY()) reads "10/8/2026" once the idle open pass recomputes volatile
// cells → the status bar reports "Circular reference: Loop!A1" → F9 → B2
// reads "6" → E2 (=_xll.CUSTOM(B2), cached 42) still reads "42", stale → the
// Data menu ends with "Recalculate" (F9) → Close → drop
// spec120-fullcalc.xlsx (fullCalcOnLoad, every formula cached 0) → its totals
// read their computed values → the download is part-for-part the fixture →
// Close → drop spec120-manual.xlsx → the manual-mode toast.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const calcFixture = join(fixturesDir, "spec120-calc.xlsx");
const fullCalcFixture = join(fixturesDir, "spec120-fullcalc.xlsx");
const manualFixture = join(fixturesDir, "spec120-manual.xlsx");

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function dropFile(page: Page, name: string, path: string) {
  const transfer = await page.evaluateHandle(
    (item) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(item.base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], item.name, { type: item.mime }));
      return dt;
    },
    { name, mime: XLSX_TYPE, base64: readFileSync(path).toString("base64") },
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
}

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

async function closeWorkbook(page: Page) {
  await page.getByTestId("close-document").click();
  await expect(page).toHaveURL(/\/$/);
  await expect(page.getByTestId("app-bar")).toBeVisible();
}

test("spec120-calc.xlsx trusts its cache, recomputes TODAY at open, reports its cycle and recalculates on F9; spec120-fullcalc.xlsx computes on open and downloads byte-identical; spec120-manual.xlsx raises the manual-mode toast (AC-43, AC-44)", async ({ page }) => {
  await page.clock.setFixedTime(new Date(2026, 9, 8, 9, 30));
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();

  // Open: cached values are trusted; once idle, the volatile cells recompute.
  await dropFile(page, "spec120-calc.xlsx", calcFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-label", "Calc");
  await expect(cell(page, "B2")).toHaveText("999");
  await expect(cell(page, "D2")).toHaveText("10/8/2026");
  await expect(cell(page, "B2")).toHaveText("999");
  await expect(page.getByTestId("sheet-status-circular")).toHaveText("Circular reference: Loop!A1");
  await expect(cell(page, "E2")).not.toHaveAttribute("data-stale", "true");

  // F9: a full recalculation. The add-in call keeps its cached value, stale.
  await page.keyboard.press("F9");
  await expect(cell(page, "B2")).toHaveText("6");
  await expect(cell(page, "E2")).toHaveText("42");
  await expect(cell(page, "E2")).toHaveAttribute("data-stale", "true");
  await expect(cell(page, "E2")).toHaveAttribute("title", "This value couldn't be recalculated here — _XLL.CUSTOM isn't supported yet");
  await expect(page.getByTestId("sheet-status-circular")).toHaveText("Circular reference: Loop!A1");

  // The Data menu, after View, ends with Recalculate (F9).
  const titles = page.getByRole("menubar").locator("[data-menu-title]");
  await expect(titles).toHaveText(["File", "View", "Data"]);
  await page.getByTestId("menu-title-data").click();
  const recalculate = page.getByTestId("sheet-menu-recalculate");
  await expect(recalculate).toBeVisible();
  await expect(recalculate).toContainText("Recalculate");
  await expect(recalculate).toContainText("F9");
  await page.keyboard.press("Escape");
  await closeWorkbook(page);

  // A library-written workbook: fullCalcOnLoad, every formula cached 0.
  await dropFile(page, "spec120-fullcalc.xlsx", fullCalcFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-label", "Totals");
  await expect(cell(page, "B4")).toHaveText("60");
  await expect(cell(page, "C4")).toHaveText("120");
  await expect(cell(page, "D4")).toHaveText("20");
  await expect(cell(page, "C2")).toHaveText("40");

  // The open-time pass changed nothing the file holds.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("spec120-fullcalc.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(await download.path())));
  const source = await unzip(new Uint8Array(readFileSync(fullCalcFixture)));
  expect([...saved.keys()]).toEqual([...source.keys()]);
  for (const [name, bytes] of source) {
    expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
  }
  await closeWorkbook(page);

  // A workbook saved in manual mode recalculates here all the same.
  await dropFile(page, "spec120-manual.xlsx", manualFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("toast").filter({ hasText: "calculate manually" })).toHaveText(
    "This workbook was set to calculate manually. It recalculates automatically here.",
  );
});
