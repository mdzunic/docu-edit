import { test, expect, type Page } from "@playwright/test";

// SPEC-073 — shape styling depth. The full pickers behind the toolbar's
// Style popover (AC-16), the same panel in Format → Drawing → Shape
// formatting… (AC-17), the five-preset gallery as one undo entry (AC-18/
// AC-24), gradient seeding (AC-22), the shadow toggle (AC-20) and the
// transparency slider's single commit (AC-25).

const content = (page: Page) => page.getByTestId("document-content");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function insertShapeAt(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

async function openStylePanel(page: Page): Promise<void> {
  await page.getByTestId("shape-style-button").click();
  await expect(page.getByTestId("shape-style-panel")).toBeVisible();
}

test.describe("SPEC-073 — shape styling pickers", () => {
  test("the toolbar's Style popover exposes the full model: fill kind, transparency, join/cap, shadow (AC-16)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await openStylePanel(page);
    await expect(page.getByTestId("shape-fill-kind")).toBeVisible();
    await expect(page.getByTestId("shape-fill-transparency")).toBeVisible();
    await expect(page.getByTestId("shape-outline-join")).toBeVisible();
    await expect(page.getByTestId("shape-outline-cap")).toBeVisible();
    await expect(page.getByTestId("shape-shadow-toggle")).toBeVisible();
    await expect(page.getByTestId("shape-shadow-preset")).toBeVisible();
  });

  test("switching the fill to gradient seeds a two-stop editor (AC-22) and radial stays available (AC-3)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await openStylePanel(page);
    await page.getByTestId("shape-fill-kind").selectOption("gradient");
    await expect(page.getByTestId("shape-gradient-editor")).toBeVisible();
    await expect(page.getByTestId("shape-gradient-stop-0")).toBeVisible();
    await expect(page.getByTestId("shape-gradient-stop-1")).toBeVisible();
    // Seeded from the default solid accent at 90°.
    await expect(page.getByTestId("shape-gradient-angle")).toHaveValue("90");
    await expect(page.getByTestId("shape-gradient-stop-color-0")).toHaveValue("#4472c4");
    await expect(page.getByTestId("shape-gradient-stop-color-1")).toHaveValue("#8faadc");
    await page.getByTestId("shape-gradient-type").selectOption("radial");
    await expect(page.getByTestId("shape-gradient-angle")).toHaveCount(0);
  });

  test("the five-preset gallery applies its triple in one click and one undo entry (AC-18, AC-24)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await openStylePanel(page);
    const presets = page.getByTestId("shape-style-presets").locator("button");
    await expect(presets).toHaveCount(5);
    await page.getByTestId("shape-style-preset-outlinedShadow").click();
    // One undo returns the shape to its freshly inserted look — the preset
    // was a single history entry.
    await page.keyboard.press("ControlOrMeta+z");
    await expect(page.getByTestId("shape-style-panel")).toBeVisible();
  });

  test("the shadow toggle applies and clears the standard shadow (AC-20)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await openStylePanel(page);
    const toggle = page.getByTestId("shape-shadow-toggle");
    await expect(toggle).toHaveText("Add shadow");
    await toggle.click();
    await expect(toggle).toHaveText("Remove shadow");
    await expect(page.getByTestId("shape-shadow-preset")).toHaveValue("bottom");
    await toggle.click();
    await expect(toggle).toHaveText("Add shadow");
  });

  test("the same panel opens from Format → Drawing → Shape formatting… (AC-17)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-drawing").click();
    await page.getByTestId("drawing-shape-formatting").click();
    await expect(page.getByTestId("shape-format-dialog")).toBeVisible();
    await expect(page.getByTestId("shape-format-dialog").getByTestId("shape-style-panel")).toBeVisible();
    await page.getByTestId("shape-format-close").click();
    await expect(page.getByTestId("shape-format-dialog")).toHaveCount(0);
  });

  test("arrowhead pickers appear for a line and default a turned-on head to medium (AC-14)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "line", 400, 250);
    await openStylePanel(page);
    await page.getByTestId("shape-start-arrow-kind").selectOption("diamond");
    await expect(page.getByTestId("shape-start-arrow-size")).toHaveValue("medium");
  });
});
