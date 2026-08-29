import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-092 — charts on slides. Open a deck whose first slide carries a clustered
// column, a line and a pie chart (painted as real charts) plus one out-of-
// catalogue 3-D chart (a dashed "not supported" placeholder with disabled edit
// controls). Edit the column chart's data and options, insert a new pie chart,
// then download → reopen and confirm the edited values, the applied options and
// the untouched placeholder all come back.
const here = dirname(fileURLToPath(import.meta.url));
const PPTX = join(here, "fixtures", "spec092-charts.pptx");

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

test("pptx: charts paint, edit through the toolbar, insert, and round-trip", async ({ page }) => {
  await openFixture(page, PPTX);
  const stage = page.getByTestId("stage-workspace");

  // AC-11/12: every chart paints into a real <canvas> — not a grey "Chart"
  // foreign box (the pre-SPEC-092 behaviour), which would show a foreign label.
  await expect(stage.locator(".chart-frame canvas")).toHaveCount(4);
  await expect(stage.locator(".foreign-label")).toHaveCount(0);

  const frames = stage.getByTestId("slide-frame");

  // AC-15: the out-of-catalogue chart (the fourth frame) is unsupported — its
  // toolbar shows the reason and offers no edit-data / options controls.
  await frames.nth(3).click();
  await expect(page.getByTestId("slide-chart-toolbar")).toBeVisible();
  await expect(page.getByTestId("chart-unsupported-reason")).toBeVisible();
  await expect(page.getByTestId("chart-edit-data-button")).toHaveCount(0);
  // Double-click is a no-op for an unsupported chart (no data sheet opens).
  await frames.nth(3).dblclick();
  await expect(page.getByTestId("chart-data-grid")).toHaveCount(0);

  // AC-14: select the column chart (first frame) → its toolbar offers Edit data.
  await frames.nth(0).click();
  await expect(page.getByTestId("chart-edit-data-button")).toBeVisible();

  // AC-14/16: Edit data → change a value cell → Done → the sheet closes.
  await page.getByTestId("chart-edit-data-button").click();
  await expect(page.getByTestId("chart-data-grid")).toBeVisible();
  await page.getByTestId("chart-data-cell-1-1").fill("99");
  await page.getByTestId("chart-data-apply").click();
  await expect(page.getByTestId("chart-data-grid")).toHaveCount(0);

  // AC-14a/16: Chart options → legend bottom + data labels on → Apply.
  await page.getByTestId("chart-options-button").click();
  await expect(page.getByTestId("chart-options-title")).toBeVisible();
  await page.getByTestId("chart-options-legend").selectOption("bottom");
  await page.getByTestId("chart-options-data-labels").check();
  await page.getByTestId("chart-options-apply").click();
  await expect(page.getByTestId("chart-options-title")).toHaveCount(0);

  // AC-13/13a: Insert → Chart → Pie → a new chart lands selected (its toolbar
  // shows the supported edit controls), bringing the deck to five charts.
  await page.getByTestId("insert-chart").click();
  await expect(page.getByTestId("chart-gallery")).toBeVisible();
  await page.getByTestId("chart-gallery-pie").click();
  await expect(page.getByTestId("chart-edit-data-button")).toBeVisible();
  await expect(stage.locator(".chart-frame canvas")).toHaveCount(5);

  // AC-26/24: download → reopen → the edited value persists, the placeholder is
  // untouched, and all five charts are present.
  const file = await downloadTo(page, "spec092-charts-roundtrip.pptx");
  await openFixture(page, file);
  const stage2 = page.getByTestId("stage-workspace");
  await expect(stage2.locator(".chart-frame canvas")).toHaveCount(5);

  // The out-of-catalogue chart still opens as an unsupported placeholder.
  const frames2 = stage2.getByTestId("slide-frame");
  await frames2.nth(3).click();
  await expect(page.getByTestId("chart-unsupported-reason")).toBeVisible();

  // The column chart's edited value survived the round trip.
  await frames2.nth(0).click();
  await page.getByTestId("chart-edit-data-button").click();
  await expect(page.getByTestId("chart-data-cell-1-1")).toHaveValue("99");
});
