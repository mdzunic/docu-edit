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
    // ...and it stays a real line after the drop, not a zero-length collapse:
    // re-anchoring on release re-indexes the connector's attachment so it still
    // resolves to the moved shape rather than a stale target (AC-7 regression).
    const moved = endpointsOfPath(d2);
    expect(Math.hypot(moved.b[0] - moved.a[0], moved.b[1] - moved.a[1])).toBeGreaterThan(80);
  });

  test("a selected connector shows handles and a formatting toolbar; kind and delete work (AC-15..21)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");
    await insertShape(page, "rect", 320, 260);
    await insertShape(page, "rect", 520, 260);
    await armConnector(page, "straight");
    await clickCenter(page, drawings(page).nth(0));
    await clickCenter(page, drawings(page).nth(1));
    await expect(connectors(page)).toHaveCount(1);

    // Select the connector by clicking its line — handles and toolbar appear.
    const box = await boxOf(connectors(page).first());
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
    await expect(page.getByTestId("connector-toolbar")).toBeVisible();
    await expect(page.getByTestId("connector-handle-start")).toBeVisible();
    await expect(page.getByTestId("connector-handle-end")).toBeVisible();

    // Change the kind to elbow → an orthogonal path (mid handle appears).
    await page.getByTestId("connector-kind").selectOption("elbow");
    await expect(page.getByTestId("connector-handle-mid")).toBeVisible();

    // Delete removes it.
    await page.getByTestId("connector-delete").click();
    await expect(connectors(page)).toHaveCount(0);
  });

  test("rotating an attached shape re-routes its connector live (AC-9)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");
    await insertShape(page, "rect", 320, 260);
    await insertShape(page, "rect", 520, 260);
    const shapeA = drawings(page).nth(0);
    await armConnector(page, "straight");
    await clickCenter(page, shapeA);
    await clickCenter(page, drawings(page).nth(1));
    await expect(connectors(page)).toHaveCount(1);
    const d1 = (await connectors(page).first().getAttribute("d")) ?? "";

    // Select shape A (near its top-left interior, clear of the line) and turn
    // it 60° — its connection sites turn with it, so the connector re-routes.
    const boxA = await boxOf(shapeA);
    await page.mouse.click(boxA.x + boxA.width * 0.3, boxA.y + boxA.height * 0.3);
    const rotation = page.getByTestId("shape-rotation");
    await rotation.fill("60");
    await rotation.dispatchEvent("change");

    const d2 = (await connectors(page).first().getAttribute("d")) ?? "";
    expect(d2).not.toBe(d1);
  });

  test("deleting a shape keeps its connector, freezing the freed end at its last position (AC-11)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");
    await insertShape(page, "rect", 320, 260);
    await insertShape(page, "rect", 520, 260);
    const shapeA = drawings(page).nth(0);
    await armConnector(page, "straight");
    await clickCenter(page, shapeA);
    await clickCenter(page, drawings(page).nth(1));
    await expect(connectors(page)).toHaveCount(1);
    // The freed end is the start (attached to shape A): record its y before.
    const before = endpointsOfPath((await connectors(page).first().getAttribute("d")) ?? "");

    // Select shape A and delete it via its toolbar.
    const boxA = await boxOf(shapeA);
    await page.mouse.click(boxA.x + boxA.width * 0.3, boxA.y + boxA.height * 0.3);
    await page.getByTestId("shape-delete-button").click();
    await expect(drawings(page)).toHaveCount(1);

    // The connector is NOT deleted (AC-11), and its freed start endpoint stays
    // put — it must not spring up to the connector's paragraph anchor at the
    // page top (the pre-fix regression jumped it ~180px).
    await expect(connectors(page)).toHaveCount(1);
    const after = endpointsOfPath((await connectors(page).first().getAttribute("d")) ?? "");
    expect(Math.abs(after.a[1] - before.a[1])).toBeLessThan(40);
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

  test("resizing an attached shape re-routes its connector and it stays a real line (AC-8)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");
    await insertShape(page, "rect", 320, 260);
    await insertShape(page, "rect", 560, 260);
    const shapeA = drawings(page).nth(0);
    await armConnector(page, "straight");
    await clickCenter(page, shapeA);
    await clickCenter(page, drawings(page).nth(1));
    await expect(connectors(page)).toHaveCount(1);
    const d1 = (await connectors(page).first().getAttribute("d")) ?? "";

    // Select shape A, then drag its SE resize handle outward.
    const boxA = await boxOf(shapeA);
    await page.mouse.click(boxA.x + boxA.width * 0.3, boxA.y + boxA.height * 0.3);
    const se = boxA.x + boxA.width;
    const sey = boxA.y + boxA.height;
    await page.mouse.move(se, sey);
    await page.mouse.down();
    await page.mouse.move(se + 60, sey + 60, { steps: 6 });
    await page.mouse.up();

    // The path changed (re-routed) and remains a real, non-degenerate line —
    // resize must not collapse the attachment.
    const d2 = (await connectors(page).first().getAttribute("d")) ?? "";
    expect(d2).not.toBe(d1);
    const { a, b } = endpointsOfPath(d2);
    expect(Math.hypot(b[0] - a[0], b[1] - a[1])).toBeGreaterThan(20);
  });

  test("releasing on empty canvas leaves that endpoint free; moving the shape re-routes only the attached end (AC-6)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Diagram");
    await insertShape(page, "rect", 320, 260);
    const shapeA = drawings(page).nth(0);

    // Draw from shape A's centre to an empty point on the canvas.
    await armConnector(page, "straight");
    await clickCenter(page, shapeA);
    const canvas = await boxOf(content(page));
    const freeX = canvas.x + 700;
    const freeY = canvas.y + 300;
    await page.mouse.click(freeX, freeY);
    await expect(connectors(page)).toHaveCount(1);

    // start = attached to shape A, end = free at the clicked point.
    const before = endpointsOfPath((await connectors(page).first().getAttribute("d")) ?? "");

    // Move shape A down: the attached start follows it, the free end stays put.
    const boxA = await boxOf(shapeA);
    await page.mouse.move(boxA.x + boxA.width / 2, boxA.y + boxA.height / 2);
    await page.mouse.down();
    await page.mouse.move(boxA.x + boxA.width / 2, boxA.y + boxA.height / 2 + 100, { steps: 6 });
    await page.mouse.up();

    const after = endpointsOfPath((await connectors(page).first().getAttribute("d")) ?? "");
    // Attached start moved appreciably…
    expect(Math.abs(after.a[1] - before.a[1])).toBeGreaterThan(40);
    // …while the free end stayed at its point.
    expect(Math.abs(after.b[0] - before.b[0])).toBeLessThan(5);
    expect(Math.abs(after.b[1] - before.b[1])).toBeLessThan(5);
  });
});
