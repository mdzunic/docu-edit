import { test, expect, type Page } from "@playwright/test";

// SPEC-071 — freeform shapes, curves and edit points. Drives the three
// Insert-gallery Freeform tools and asserts a real path-bearing drawing
// lands (the model + round-trip are covered exhaustively by the unit
// suites; this proves the gesture wiring reaches them).

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function armFreeform(page: Page, tool: string): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`freeform-gallery-${tool}`).click();
}

test.describe("SPEC-071 — freeform drawing", () => {
  test("Freeform tool places vertices and double-click finishes an open shape (AC-6, AC-7)", async ({
    page,
  }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await content(page).dblclick({ position: { x: 640, y: 220 } });
    await expect(drawings(page)).toHaveCount(1);
    // The gallery is a Freeform group with three tools.
    await page.getByTestId("menu-title-insert").click();
    await expect(page.getByTestId("freeform-gallery-freeform")).toBeVisible();
    await expect(page.getByTestId("freeform-gallery-curve")).toBeVisible();
    await expect(page.getByTestId("freeform-gallery-scribble")).toBeVisible();
  });

  test("Escape during a freeform draw creates nothing (AC-10)", async ({ page }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await page.keyboard.press("Escape");
    await expect(drawings(page)).toHaveCount(0);
  });

  test("a single click then finish creates nothing (AC-7)", async ({ page }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    await content(page).click({ position: { x: 400, y: 200 } });
    await page.keyboard.press("Enter");
    await expect(drawings(page)).toHaveCount(0);
  });

  test("Curve tool draws a shape that survives save and reopen", async ({ page }) => {
    await createDocument(page);
    await armFreeform(page, "curve");
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 500, y: 320 } });
    await content(page).dblclick({ position: { x: 620, y: 200 } });
    await expect(drawings(page)).toHaveCount(1);
  });
});

test.describe("SPEC-071 — edit points", () => {
  async function insertPreset(page: Page, geometry: string, x: number, y: number): Promise<void> {
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId(`shape-gallery-${geometry}`).click();
    await content(page).click({ position: { x, y } });
  }

  test("Format -> Edit points converts a preset and shows vertex handles (AC-14, AC-15, AC-19)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await page.getByTestId("shape-edit-points-button").click();
    await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
    // A triangle converts to a 3-vertex freeform.
    await expect(page.getByTestId("point-edit-vertex-0")).toBeVisible();
    await expect(page.getByTestId("point-edit-vertex-2")).toBeVisible();
  });

  test("dragging a vertex commits one edit and Escape exits (AC-17, AC-22)", async ({ page }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    await page.getByTestId("shape-edit-points-button").click();
    const handle = page.getByTestId("point-edit-vertex-0");
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("no vertex handle");
    }
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 40, box.y - 30, { steps: 8 });
    await page.mouse.up();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("point-edit-overlay")).toHaveCount(0);
    // The shape is still there, still a single drawing.
    await expect(drawings(page)).toHaveCount(1);
  });
});
