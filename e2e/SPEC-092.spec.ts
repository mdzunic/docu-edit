import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { openInsertTab } from "./slides";

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

// Read the first data cell (row 1, col 1) of every supported chart on the slide,
// draining as it goes: select the topmost frame, read it (unsupported charts have
// no Edit-data control and are skipped), then delete it to expose the next. This
// stays robust when charts overlap — a duplicate lands stacked on its source, so
// index-based clicks would be intercepted. The deck is discarded after, so the
// destructive read is harmless.
async function drainChartFirstValues(page: Page): Promise<string[]> {
  const stage = page.getByTestId("stage-workspace");
  const values: string[] = [];
  for (let guard = 0; guard < 20; guard++) {
    const frames = stage.getByTestId("slide-frame");
    const n = await frames.count();
    if (n === 0) break;
    await frames.nth(n - 1).click({ force: true });
    const editBtn = page.getByTestId("chart-edit-data-button");
    if ((await editBtn.count()) > 0) {
      await editBtn.click();
      values.push(await page.getByTestId("chart-data-cell-1-1").inputValue());
      await page.keyboard.press("Escape");
      await expect(page.getByTestId("chart-data-grid")).toHaveCount(0);
    }
    await page.getByTestId("chart-delete-button").click();
    await expect(stage.getByTestId("slide-frame")).toHaveCount(n - 1);
  }
  return values;
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

  // AC-14: double-clicking a supported chart opens its data sheet directly.
  await frames.nth(0).dblclick();
  await expect(page.getByTestId("chart-data-grid")).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("chart-data-grid")).toHaveCount(0);

  // AC-14: the column chart stays selected → its toolbar offers Edit data.
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
  await openInsertTab(page);
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

// AC-18a/AC-25/AC-18/AC-30 — the regression guard. Two frames may share one
// sourcePart (Duplicate, per D8); editing one must mint a fresh part and leave
// the other's data intact. A prior round shipped a writer that silently
// overwrote the shared part, so the duplicated sibling came back corrupted.
test("pptx: duplicating an imported chart then editing the copy leaves the original intact", async ({ page }) => {
  await openFixture(page, PPTX);
  const stage = page.getByTestId("stage-workspace");
  const frames = stage.getByTestId("slide-frame");

  // Select the imported line chart (frame 1) and duplicate it. The copy shares
  // the original's sourcePart (ppt/charts/chart2.xml) until the deck is written.
  await frames.nth(1).click();
  await page.getByTestId("chart-duplicate-button").click();
  // Five charts now paint (the four fixture charts plus the duplicate).
  await expect(stage.locator(".chart-frame canvas")).toHaveCount(5);

  // The duplicate lands selected, stacked on its source. Edit its first value to
  // a sentinel (99). The original line chart's Q1 is 3 and must not move.
  await page.getByTestId("chart-edit-data-button").click();
  await expect(page.getByTestId("chart-data-grid")).toBeVisible();
  await expect(page.getByTestId("chart-data-cell-1-1")).toHaveValue("3");
  await page.getByTestId("chart-data-cell-1-1").fill("99");
  await page.getByTestId("chart-data-apply").click();
  await expect(page.getByTestId("chart-data-grid")).toHaveCount(0);

  // Download → reopen and read every supported chart's first value. The edit
  // must land in exactly one chart (a freshly-minted part) while the untouched
  // sibling keeps its original 3 — if the shared part had been overwritten, the
  // sentinel would appear twice and the 3 would be gone.
  const file = await downloadTo(page, "spec092-charts-duplicate.pptx");
  await openFixture(page, file);
  const values = await drainChartFirstValues(page);
  expect(values.filter((v) => v === "99")).toHaveLength(1);
  expect(values).toContain("3");
});
