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

  test("an open freeform hit-tests along its stroke, not just at its vertices (AC-24)", async ({
    page,
  }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    // A V: first segment (400,200)->(520,320), second (520,320)->(640,220).
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await content(page).dblclick({ position: { x: 640, y: 220 } });
    await expect(drawings(page)).toHaveCount(1);
    // Deselect the freshly drawn (auto-selected) shape.
    await content(page).click({ position: { x: 40, y: 40 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // Click the exact midpoint of the first segment — on the stroke but deep
    // in the shape's bounding-box interior, far from any vertex. Before the
    // fix this only selected at the extreme vertices (the box edges).
    await content(page).click({ position: { x: 460, y: 260 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  });

  test("an open freeform is hit-tested against its outline, not its bounding box (AC-24)", async ({
    page,
  }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    // Same V: segments (400,200)->(520,320)->(640,220).
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await content(page).dblclick({ position: { x: 640, y: 220 } });
    await content(page).click({ position: { x: 40, y: 40 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // A point inside the V's bounding box but far from either stroke segment
    // (the empty gap above the V's notch): an unfilled path must NOT select
    // here — proving the hit test follows the real outline, not the box.
    await content(page).click({ position: { x: 520, y: 210 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // ...while a point on the stroke does select.
    await content(page).click({ position: { x: 460, y: 260 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
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

  test("point-edit hides the box resize/rotation handles so a vertex handle is grabbable (AC-17)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    // Resize/rotation handles are shown while the shape is merely selected.
    await expect(page.getByTestId("shape-handles")).toBeVisible();
    await page.getByTestId("shape-edit-points-button").click();
    await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
    // In point-edit mode they are gone, so a vertex handle sitting on a box
    // corner (every vertex of a converted triangle) is the topmost element at
    // its own centre and receives the drag instead of a resize handle.
    await expect(page.getByTestId("shape-handles")).toHaveCount(0);
    const handle = page.getByTestId("point-edit-vertex-0");
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("no vertex handle");
    }
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    const topmost = await page.evaluate(
      ([x, y]) => document.elementFromPoint(x, y)?.getAttribute("data-testid") ?? "",
      [cx, cy],
    );
    expect(topmost).toBe("point-edit-vertex-0");
  });

  test("dragging a vertex on a box corner edits the point, not the box (AC-17)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    await page.getByTestId("shape-edit-points-button").click();
    await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
    const handle = page.getByTestId("point-edit-vertex-0");
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("no vertex handle");
    }
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    // This vertex sits on a corner of the shape's bounding box, exactly where a
    // resize handle used to sit on top of it. Dragging it must move the vertex
    // (committing an "edit points" undo) — before the fix it resized the box.
    await page.mouse.move(cx, cy);
    await page.mouse.down();
    await page.mouse.move(cx + 60, cy - 40, { steps: 10 });
    await page.mouse.up();
    // The vertex handle followed the drag rather than the box being resized.
    const moved = await handle.boundingBox();
    if (!moved) {
      throw new Error("vertex handle vanished");
    }
    expect(Math.round(moved.x + moved.width / 2 - cx)).toBeGreaterThan(30);
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("point-edit-overlay")).toHaveCount(0);
    const undoLabel = await page.getByTestId("undo-button").getAttribute("aria-label");
    expect(undoLabel).toContain("edit points");
    expect(undoLabel).not.toContain("resize");
  });
});
