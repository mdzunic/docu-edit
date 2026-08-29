import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-091 — tables on slides. Open a deck whose first slide carries a modelled
// 3×3 table (header + banded rows) and one deliberately out-of-subset table (a
// preserved foreign frame). The table renders as a real <table>; select it, open
// the Table panel, insert a row and merge the two header cells, then download →
// reopen and confirm the grid, the span and the preserved foreign table all
// come back.
const here = dirname(fileURLToPath(import.meta.url));
const PPTX = join(here, "fixtures", "spec091-tables.pptx");

async function openFixture(page: Page, file: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("pptx: table renders, edits through the panel, and round-trips", async ({ page }) => {
  await openFixture(page, PPTX);

  const stage = page.getByTestId("stage-workspace");
  const table = stage.getByTestId("slide-table");
  // AC-19a: the table renders as a real <table> with its cells.
  await expect(table).toBeVisible();
  await expect(stage.getByText("Region")).toBeVisible();
  await expect(table.locator("tbody > tr")).toHaveCount(3);
  // The deliberately out-of-subset table stays a foreign box (AC-17c).
  await expect(stage.locator(".foreign-label")).toHaveText("Table");

  // Select the table and open the Table panel.
  await stage.getByTestId("slide-table-cell-0-0").click();
  await page.getByTestId("slides-table-toggle").click();
  await expect(page.getByTestId("slides-table-panel")).toBeVisible();

  // AC-19c: insert a row below the anchor cell → four rows.
  await page.getByTestId("table-insert-row-below").click();
  await expect(table.locator("tbody > tr")).toHaveCount(4);

  // AC-19d: merge the two header cells → the first header cell spans 2 columns.
  await page.getByTestId("table-merge-cells").click();
  await expect(stage.getByTestId("slide-table-cell-0-0")).toHaveAttribute("colspan", "2");

  // AC-19e: banded rows toggle is a real control on the panel.
  await expect(page.getByTestId("table-band-rows")).toBeVisible();

  // AC-19f: download → reopen → the same grid, span, text and the preserved
  // out-of-subset table all come back.
  const file = await downloadTo(page, "spec091-tables-roundtrip.pptx");
  await openFixture(page, file);
  const stage2 = page.getByTestId("stage-workspace");
  await expect(stage2.getByTestId("slide-table")).toBeVisible();
  await expect(stage2.getByText("Region")).toBeVisible();
  await expect(stage2.getByTestId("slide-table").locator("tbody > tr")).toHaveCount(4);
  await expect(stage2.getByTestId("slide-table-cell-0-0")).toHaveAttribute("colspan", "2");
  await expect(stage2.locator(".foreign-label")).toHaveText("Table");
});
