import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";

// SPEC-132: sheet charts in spec132-charts.xlsx — live bars from the cells,
// placeholders for what is not modelled, insert from selection, options, the
// chart sheet, and a download that keeps the edits.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const decoder = new TextDecoder();

test.use({ viewport: { width: 1440, height: 900 } });

function drawing(page: Page, name: string): Locator {
  return page.locator(`[data-testid="sheet-drawing"][data-primary="true"][data-drawing-name="${name}"]`);
}

function tab(page: Page, name: string): Locator {
  return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
}

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
}

async function goTo(page: Page, reference: string): Promise<void> {
  const nameBox = page.getByTestId("name-box");
  await nameBox.fill(reference);
  await nameBox.press("Enter");
}

const COLUMN_VALUES = [
  [10, 20, 30, 40],
  [15, 25, 35, 45],
  [12, 30, 18, 24],
];

test("SPEC-132: the column chart follows its cells and a row insert; the 3-D and waterfall charts are placeholders; a line chart inserts from a selection; the chart sheet paints; the download keeps it all", async ({ page }) => {
  await open(page, join(fixtures, "spec132-charts.xlsx"));

  // The clustered column paints its bars from the caches.
  const column = drawing(page, "Chart 1").getByTestId("sheet-chart-canvas");
  await expect(column).toHaveAttribute("data-chart-type", "columnClustered");
  await expect(column).toHaveAttribute("data-chart-values", JSON.stringify(COLUMN_VALUES));

  // The 3-D chart is paintChart's placeholder; selected, its edit controls give way to the reason.
  const threeD = drawing(page, "Chart 5");
  await expect(threeD.getByTestId("sheet-chart-canvas")).toHaveAttribute("data-chart-unsupported", "true");
  await threeD.click();
  const toolbar = page.getByTestId("sheet-chart-toolbar");
  await expect(toolbar.getByTestId("chart-unsupported-reason")).toBeVisible();
  await expect(toolbar.getByTestId("chart-edit-data-button")).toHaveCount(0);
  await expect(toolbar.getByTestId("chart-options-button")).toHaveCount(0);
  // The waterfall (chartex) keeps SPEC-131's "Chart" placeholder.
  await expect(drawing(page, "Chart 7")).toContainText("Chart");
  await page.keyboard.press("Escape");

  // A larger number in Data!D3 grows its bar.
  await goTo(page, "D3");
  await page.keyboard.type("90");
  await page.keyboard.press("Enter");
  const edited = [COLUMN_VALUES[0], COLUMN_VALUES[1], [12, 90, 18, 24]];
  await expect(column).toHaveAttribute("data-chart-values", JSON.stringify(edited));

  // A row inserted above the data: the chart still plots the same values.
  await page.getByTestId("sheet-row-header").filter({ hasText: /^1$/ }).click({ button: "right" });
  await page.getByTestId("sheet-header-menu").getByTestId("sheet-insert-before").click();
  await expect(column).toHaveAttribute("data-chart-values", JSON.stringify(edited));

  // Summary: A1:C4 → Insert ▸ Chart ▸ Line.
  await tab(page, "Summary").click();
  await goTo(page, "A1:C4");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("sheet-menu-insert-chart").click();
  await expect(page.getByTestId("sheet-chart-gallery-popover")).toBeVisible();
  await page.getByTestId("chart-gallery-line").click();
  const created = page.locator('[data-testid="sheet-drawing"][data-primary="true"][aria-selected="true"]');
  const createdCanvas = created.getByTestId("sheet-chart-canvas");
  await expect(createdCanvas).toHaveAttribute("data-chart-type", "line");
  await expect(createdCanvas).toHaveAttribute("data-chart-values", JSON.stringify([[4, 6, 8], [5, 5, 9]]));

  // Chart options → legend "bottom" → Apply.
  await page.getByTestId("sheet-chart-toolbar").getByTestId("chart-options-button").click();
  await page.getByTestId("chart-options-legend").selectOption("bottom");
  await page.getByTestId("chart-options-apply").click();
  await expect(createdCanvas).toHaveAttribute("data-chart-legend", "bottom");

  // The chart sheet paints its chart.
  await tab(page, "Sales Chart").click();
  const chartSheet = page.getByTestId("sheet-chart-sheet");
  await expect(chartSheet).toBeVisible();
  await expect(chartSheet.getByTestId("sheet-chart-canvas")).toHaveAttribute("data-chart-type", "columnClustered");

  // Download: the edited value's cache, the moved references, the new chart and its options persist;
  // the 3-D chart's part differs from the source only in its adjusted c:f texts.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = join(tmpdir(), `spec132-${Date.now()}-charts.xlsx`);
  await (await downloadPromise).saveAs(saved);
  const parts = await unzip(readFileSync(saved));
  const source = await unzip(readFileSync(join(fixtures, "spec132-charts.xlsx")));
  const text = (files: Map<string, Uint8Array>, name: string): string => decoder.decode(files.get(name)!);
  const chart1 = text(parts, "xl/charts/chart1.xml");
  expect(chart1).toContain("<c:f>Data!$D$3:$D$6</c:f>");
  expect(chart1).toContain("<c:v>90</c:v>");
  const withoutFormulas = (xml: string): string => xml.replace(/<c:f>[^<]*<\/c:f>/g, "<c:f/>");
  expect(text(parts, "xl/charts/chart5.xml")).not.toBe(text(source, "xl/charts/chart5.xml"));
  expect(withoutFormulas(text(parts, "xl/charts/chart5.xml"))).toBe(withoutFormulas(text(source, "xl/charts/chart5.xml")));
  const added = [...parts.keys()].filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name) && !source.has(name));
  expect(added).toHaveLength(1);
  expect(text(parts, added[0]!)).toContain("<c:lineChart>");
  expect(text(parts, added[0]!)).toContain('<c:legendPos val="b"/>');

  // Reopened, the chart plots the edited cells.
  await open(page, saved);
  await expect(drawing(page, "Chart 1").getByTestId("sheet-chart-canvas")).toHaveAttribute("data-chart-values", JSON.stringify(edited));
});
