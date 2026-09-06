import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";
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

  // AC-19b: double-click a cell, type, and see the text — the same in-place
  // text machinery SPEC-081 gives text frames, now editing a cell's paragraphs.
  await stage.getByTestId("slide-table-cell-1-1").dblclick();
  await expect(page.getByTestId("slide-table-cell-edit")).toBeVisible();
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.type("Growth");
  await page.keyboard.press("Escape");
  await expect(stage.getByTestId("slide-table-cell-1-1")).toContainText("Growth");

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

test("pptx: dragging a table grid boundary resizes a column and a row (AC-8a,8b)", async ({ page }) => {
  await openFixture(page, PPTX);
  const stage = page.getByTestId("stage-workspace");

  // Selecting the table (single click on a cell) reveals the internal grid
  // boundary handles — one per gap in a 3×3 table.
  await stage.getByTestId("slide-table-cell-0-0").click();
  const colHandles = stage.getByTestId("table-col-resize");
  const rowHandles = stage.getByTestId("table-row-resize");
  await expect(colHandles).toHaveCount(2);
  await expect(rowHandles).toHaveCount(2);

  // AC-8a: drag the first column boundary right → column 0 gets wider.
  const cell = stage.getByTestId("slide-table-cell-0-0");
  const widthBefore = (await cell.boundingBox())!.width;
  const colBox = (await colHandles.first().boundingBox())!;
  await page.mouse.move(colBox.x + colBox.width / 2, colBox.y + colBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(colBox.x + 60, colBox.y + colBox.height / 2, { steps: 5 });
  await page.mouse.up();
  await expect.poll(async () => (await cell.boundingBox())!.width).toBeGreaterThan(widthBefore + 5);

  // AC-8b: drag the first row boundary down → row 0 gets taller.
  const heightBefore = (await cell.boundingBox())!.height;
  const rowBox = (await rowHandles.first().boundingBox())!;
  await page.mouse.move(rowBox.x + rowBox.width / 2, rowBox.y + rowBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(rowBox.x + rowBox.width / 2, rowBox.y + 50, { steps: 5 });
  await page.mouse.up();
  await expect.poll(async () => (await cell.boundingBox())!.height).toBeGreaterThan(heightBefore + 5);
});

// The cell currently hosting the in-place editor, as "slide-table-cell-{r}-{c}"
// (the editor host is nested inside the addressed <td>). Used to observe Tab /
// Shift+Tab navigation on the real contenteditable — the jsdom unit test invokes
// onKeydown directly and cannot prove the DOM key event drives the gesture.
function editingCell(page: Page): Promise<string> {
  return page.evaluate(() => {
    const edit = document.querySelector('[data-testid="slide-table-cell-edit"]');
    const cell = edit?.parentElement?.closest('[data-testid^="slide-table-cell-"]');
    return cell?.getAttribute("data-testid") ?? "none";
  });
}

test("pptx: Tab/Shift+Tab move between cells, arrows do not (AC-6b, AC-6c)", async ({ page }) => {
  await openFixture(page, PPTX);
  const stage = page.getByTestId("stage-workspace");
  const table = stage.getByTestId("slide-table");

  // Start a cell session (AC-6a) at the top-left cell.
  await stage.getByTestId("slide-table-cell-0-0").dblclick();
  await expect(page.getByTestId("slide-table-cell-edit")).toBeVisible();
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-0-0");

  // AC-6b: Tab walks the row and wraps to the next row's first cell.
  await page.keyboard.press("Tab");
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-0-1");
  await page.keyboard.press("Tab");
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-0-2");
  await page.keyboard.press("Tab");
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-1-0");
  // Shift+Tab steps back across the same wrap.
  await page.keyboard.press("Shift+Tab");
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-0-2");

  // AC-6c: arrow keys are NOT trapped for grid navigation — the session stays
  // in the same cell (the caret moves within / leaves at the table edge).
  await page.keyboard.press("ArrowRight");
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-0-2");

  // AC-6b: Tab in the last cell appends a row and lands in its first cell.
  await page.keyboard.press("Escape");
  await stage.getByTestId("slide-table-cell-2-2").dblclick();
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-2-2");
  await expect(table.locator("tbody > tr")).toHaveCount(3);
  await page.keyboard.press("Tab");
  await expect(table.locator("tbody > tr")).toHaveCount(4);
  await expect.poll(() => editingCell(page)).toBe("slide-table-cell-3-0");
});

// The Table panel's own promise. Opening its ribbon toggle used to light the
// toggle and render nothing at all: SlidePanelStack gated the panel on a table
// being selected as well as the toggle being on, so "Table" was a control that
// visibly did nothing, and a panel already open vanished the moment you clicked
// off the table. The panel now behaves like the other four — it appears when
// its toggle is on — and states what to select, with its commands disabled
// rather than live-looking and silently inert.
test("the Table panel opens with no table selected, explains itself, and goes live on selection", async ({ page }) => {
  await openFixture(page, PPTX);

  const rowBelow = page.getByTestId("table-insert-row-below");

  // Nothing selected: the toggle turns the panel on, and the panel says why it
  // cannot act yet.
  await page.getByTestId("slides-table-toggle").click();
  await expect(page.getByTestId("slides-table-panel")).toBeVisible();
  await expect(page.getByTestId("slides-table-empty")).toBeVisible();
  await expect(rowBelow).toBeDisabled();

  // Selecting a table arms it, with no second trip to the toggle.
  const stage = page.getByTestId("stage-workspace");
  await stage.getByTestId("slide-table-cell-0-0").click();
  await expect(page.getByTestId("slides-table-empty")).toHaveCount(0);
  await expect(rowBelow).toBeEnabled();
  await expect(stage.getByTestId("slide-table").locator("tbody > tr")).toHaveCount(3);
  await rowBelow.click();
  await expect(stage.getByTestId("slide-table").locator("tbody > tr")).toHaveCount(4);

  // Clicking off the table disarms the commands but keeps the panel — the
  // toggle stays pressed, so what it controls stays on screen.
  await stage.getByTestId("slide-stage").click({ position: { x: 6, y: 6 } });
  await expect(page.getByTestId("slides-table-panel")).toBeVisible();
  await expect(page.getByTestId("slides-table-empty")).toBeVisible();
  await expect(rowBelow).toBeDisabled();
  await expect(page.getByTestId("slides-table-toggle")).toHaveAttribute("aria-pressed", "true");
});

// The ribbon's two kinds of control (SPEC-061 Toolbar Redesign). An open panel
// used to render as a second *selected tab* beside Home/Insert — two
// aria-selected tabs in one tablist, and two accent-underlined tabs on screen.
test("an open panel is a pressed toggle, never a second selected tab", async ({ page }) => {
  await openFixture(page, PPTX);

  const tabs = page.getByTestId("slides-ribbon-tabs").getByRole("tab");
  await expect(tabs).toHaveCount(2);

  await page.getByTestId("slides-tab-insert").click();
  await page.getByTestId("slides-table-toggle").click();
  await page.getByTestId("slides-media-toggle").click();

  // Exactly one selected tab, whatever is open beside it.
  await expect(page.getByTestId("slides-ribbon-tabs").locator('[aria-selected="true"]')).toHaveCount(1);
  await expect(page.getByTestId("slides-tab-insert")).toHaveAttribute("aria-selected", "true");
  // The panel toggles are pressed buttons, and more than one may be pressed.
  await expect(page.getByTestId("slides-table-toggle")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("slides-media-toggle")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("slides-ribbon-panels").getByRole("tab")).toHaveCount(0);
});
