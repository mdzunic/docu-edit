import { test, expect, type Page } from "@playwright/test";

// SPEC-055 — captions and the table of figures. This focused e2e covers the
// three flows that decide the feature's usability: captioning a table selected
// the ordinary way (a click into a cell, AC-1/AC-2), a freshly generated table
// of figures NOT showing a spurious update control (AC-32), and no uncaught
// exception during the idle pagination pass that follows (a ToF is a block-
// level atom, not a paragraph). The exhaustive round-trip/numbering coverage
// lives in the src/lib unit specs and fixtures.captions.spec.ts.

const content = (page: Page) => page.getByTestId("document-content");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

test("caption a table selected by clicking a cell, then a fresh table of figures", async ({ page }) => {
  const errors: string[] = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") errors.push(msg.text());
  });
  page.on("pageerror", (err) => errors.push(String(err)));

  await createDocument(page);

  // Insert a 3×3 table via the grid picker.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-3-3").click();
  await expect(content(page).locator("table")).toHaveCount(1);

  // Click into a table cell — the ordinary way to "select" a table.
  await content(page).locator("table td").first().click();

  // Insert -> Caption opens the dialog (the AC-1/AC-2 fix).
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("insert-caption-button").click();
  await expect(page.getByTestId("caption-dialog")).toBeVisible();

  // Default label is Figure; switch to Table to caption the table.
  await page.getByTestId("caption-label").selectOption("Table");
  await page.getByTestId("caption-insert").click();
  await expect(page.getByTestId("caption-dialog")).toHaveCount(0);

  // A Caption paragraph now shows "Table 1".
  await expect(content(page).getByText(/Table\s*1/).first()).toBeVisible();

  // Type some caption text so it reads like a real caption.
  await page.keyboard.type("Quarterly results");

  // Insert -> Table of figures for the Figure label is empty; use Table.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("insert-table-of-figures-button").click();
  await expect(page.getByTestId("table-of-figures-dialog")).toBeVisible();
  await page.getByTestId("tof-label").selectOption("Table");
  await page.getByTestId("tof-insert").click();
  await expect(page.getByTestId("table-of-figures-dialog")).toHaveCount(0);

  // The ToF block generated with one entry.
  await expect(page.getByTestId("tof-block")).toBeVisible();
  await expect(page.getByTestId("tof-entry")).toHaveCount(1);

  // AC-32: a freshly generated ToF shows NO update bar.
  await expect(page.getByTestId("tof-update-bar")).toHaveCount(0);

  // Give the idle pagination pass a moment to run (the prior crash site).
  await page.waitForTimeout(800);
  await expect(page.getByTestId("tof-block")).toBeVisible();
  await expect(page.getByTestId("tof-update-bar")).toHaveCount(0);

  expect(errors, `console errors:\n${errors.join("\n")}`).toEqual([]);
});
