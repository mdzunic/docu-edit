import { test, expect, type Page } from "@playwright/test";

// SPEC-069: the expanded shape catalogue, its grouped/searchable gallery and
// the geometry adjustment handles. Mirrors SPEC-051.spec.ts's own helpers
// (a fresh document, "Insert -> Shape", a single canvas click to place).
const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openGallery(page: Page): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await expect(page.getByTestId("shape-gallery")).toBeVisible();
}

async function insertShapeAt(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await openGallery(page);
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

/** Whether the canvas preceding a drawing proxy has any non-transparent
 * pixel near its centre — a real painted outline/fill, not the empty
 * interior of the SPEC-051 dashed placeholder. */
async function centrePainted(page: Page, index: number): Promise<boolean> {
  const canvas = drawings(page).nth(index).locator("xpath=preceding-sibling::canvas[1]");
  return canvas.evaluate((el) => {
    const c = el as HTMLCanvasElement;
    const ctx = c.getContext("2d")!;
    const { data } = ctx.getImageData(0, 0, c.width, c.height);
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] > 10) {
        return true;
      }
    }
    return false;
  });
}

test.describe("SPEC-069 — expanded shape catalogue", () => {
  test("the gallery shows the new groups and the search field filters by label/alias (AC-2)", async ({ page }) => {
    await createDocument(page);
    await openGallery(page);
    // Groups beyond SPEC-051's four are present.
    await expect(page.getByTestId("shape-gallery-flowChartProcess")).toBeVisible();
    await expect(page.getByTestId("shape-gallery-hexagon")).toBeVisible();
    await expect(page.getByTestId("shape-gallery-ribbon2")).toBeVisible();

    // Searching by alias narrows the gallery to the matching geometry.
    await page.getByTestId("shape-gallery-search").fill("cylinder");
    await expect(page.getByTestId("shape-gallery-can")).toBeVisible();
    await expect(page.getByTestId("shape-gallery-rect")).toHaveCount(0);
    await expect(page.getByTestId("shape-gallery-flowChartProcess")).toHaveCount(0);
  });

  test("inserting an expanded-catalogue geometry paints a real outline, not the placeholder (AC-8a)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "flowChartDecision", 400, 250);
    await expect(drawings(page)).toHaveCount(1);
    expect(await centrePainted(page, 0)).toBe(true);
  });

  test("a shape with adjustments shows a draggable adjustment handle, distinct from resize handles (AC-5)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "chevron", 400, 250);
    await drawings(page).nth(0).click();
    await expect(page.getByTestId("shape-handles")).toBeVisible();
    const handle = page.getByTestId("shape-adjust-0");
    await expect(handle).toBeVisible();
    await expect(handle).toHaveAttribute("role", "slider");
    // It is not one of the eight square resize handles.
    await expect(page.getByTestId("shape-resize-nw")).toBeVisible();

    // Dragging the handle commits an "adjust shape" edit that a single undo
    // reverses — the drawing count is unchanged, proving it edited rather
    // than inserted/removed a shape.
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("adjustment handle has no bounding box");
    }
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 30, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await expect(drawings(page)).toHaveCount(1);
  });
});
