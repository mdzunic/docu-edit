import { test, expect, type Page, type Locator } from "@playwright/test";

// SPEC-070 — connectors: the Insert gallery's Connectors group arms a
// two-click draw mode; drawing between two shapes attaches each endpoint to
// the shape under it and paints a real line (not the round-1 degenerate
// zero-length line pinned to the page top); moving a shape re-routes every
// connector attached to it live, with no separate commit.

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');
const connectors = (page: Page) => page.locator('path[data-testid^="connector-"]');

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function boxOf(locator: Locator): Promise<{ x: number; y: number; width: number; height: number }> {
  const box = await locator.boundingBox();
  if (!box) {
    throw new Error("locator has no bounding box");
  }
  return box;
}

/** "Insert → Shape", then a single click placing `geometry` at its default
 * size at (x, y) in the content box. */
async function insertShape(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

/** Arms a connector kind from the Insert gallery's Connectors group. */
async function armConnector(page: Page, kind: string): Promise<void> {
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`connector-gallery-${kind}`).click();
}

/** The two endpoints of a connector's `d` (`M x y L … x y` or the curved
 * `M x y C …`). Returns the first and last coordinate pair. */
function endpointsOfPath(d: string): { a: [number, number]; b: [number, number] } {
  const nums = d.match(/-?\d+(?:\.\d+)?/g)?.map(Number) ?? [];
  return { a: [nums[0], nums[1]], b: [nums[nums.length - 2], nums[nums.length - 1]] };
}

async function clickCenter(page: Page, locator: Locator): Promise<void> {
  const box = await boxOf(locator);
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
}

test.describe("SPEC-070 — connectors", () => {
  test("the Insert gallery offers the three connector kinds (AC-2)", async ({ page }) => {
    await createDocument(page);
    await page.getByTestId("menu-title-insert").click();
    await expect(page.getByTestId("connector-gallery-straight")).toBeVisible();
    await expect(page.getByTestId("connector-gallery-elbow")).toBeVisible();
    await expect(page.getByTestId("connector-gallery-curved")).toBeVisible();
  });

  test("drawing a connector between two shapes paints a real line and re-routes on a move (AC-3..7)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");

    // Two shapes, well apart horizontally.
    await insertShape(page, "rect", 320, 260);
    await insertShape(page, "rect", 520, 260);
    await expect(drawings(page)).toHaveCount(2);
    const shapeA = drawings(page).nth(0);
    const shapeB = drawings(page).nth(1);

    // Two-click draw: source shape, then target shape (AC-3, AC-5).
    await armConnector(page, "straight");
    await clickCenter(page, shapeA);
    await clickCenter(page, shapeB);

    // A connector exists and its line is NOT degenerate (AC-5): the two
    // endpoints are far apart, running between the shapes.
    await expect(connectors(page)).toHaveCount(1);
    const path = connectors(page).first();
    const d1 = (await path.getAttribute("d")) ?? "";
    const { a, b } = endpointsOfPath(d1);
    expect(Math.hypot(b[0] - a[0], b[1] - a[1])).toBeGreaterThan(80);

    // Moving shape A re-routes the connector live (AC-7): its `d` changes.
    const boxA = await boxOf(shapeA);
    await page.mouse.move(boxA.x + boxA.width / 2, boxA.y + boxA.height / 2);
    await page.mouse.down();
    await page.mouse.move(boxA.x + boxA.width / 2, boxA.y + boxA.height / 2 + 120, { steps: 8 });
    await page.mouse.up();

    const d2 = (await connectors(page).first().getAttribute("d")) ?? "";
    expect(d2).not.toBe(d1);
  });

  test("hovering a shape while drawing reveals its connection sites (AC-4)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");
    await insertShape(page, "rect", 360, 260);
    const shapeA = drawings(page).nth(0);

    await armConnector(page, "elbow");
    // Hover the shape (move the pointer over it) — the nine sites appear.
    const box = await boxOf(shapeA);
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await expect(page.getByTestId("connector-site")).toHaveCount(9);
  });
});
