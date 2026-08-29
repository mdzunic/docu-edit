import { test, expect, type Page, type Locator } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-051 — drawing shapes and text boxes: AC-63 (opening
// fixtures/docx/word-shapes.docx) and AC-64 (opening fixtures/odt/lo-shapes.odt).
// Both fixtures are hand-built by scripts/build-spec051-fixtures.mjs — NOT
// round-tripped through this project's own writer — and contain, in this
// fixed document order, every "Shape catalogue" geometry, a grouped pair, a
// text box and one geometry outside the catalogue:
//   0  line              7  star5              14 group member A (rect)
//   1  arrow             8  rightArrow         15 group member B (roundRect)
//   2  rect              9  leftArrow          16 text box (rect, no fill/outline)
//   3  roundRect        10  upArrow            17 unsupported geometry ("pentagon")
//   4  ellipse          11  downArrow
//   5  triangle         12  wedgeRectCallout
//   6  diamond          13  wedgeRoundRectCallout

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxShapesFixture = join(fixturesDir, "word-shapes.docx");
const odtShapesFixture = join(fixturesDir, "lo-shapes.odt");

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');
const drawingAt = (page: Page, index: number) => content(page).locator(`[data-testid="drawing-${index}"]`);

// The catalogue geometries in the fixed fixture order above, with the alt
// text scripts/build-spec051-fixtures.mjs gives each one ("<label> shape").
const CATALOGUE_LABELS = [
  "Line",
  "Arrow",
  "Rectangle",
  "Rounded rectangle",
  "Ellipse",
  "Triangle",
  "Diamond",
  "Star (5-point)",
  "Right arrow",
  "Left arrow",
  "Up arrow",
  "Down arrow",
  "Rectangular callout",
  "Rounded callout",
];

/** Samples one pixel from the canvas immediately preceding a drawing proxy
 * in the DOM (DocShapeBlock.vue paints geometry/fill/outline on a `<canvas
 * class="doc-shape-canvas">` sibling before the ARIA proxy) — used to check
 * that a catalogue shape's solid fill actually painted, and that the
 * unsupported-geometry placeholder painted no fill at all (AC-60), since
 * neither is otherwise visible in the DOM. Coordinates are fractions (0..1)
 * of the canvas's own backing-store size, so devicePixelRatio scaling
 * doesn't matter. */
async function canvasPixelAt(
  proxy: Locator,
  xFrac: number,
  yFrac: number,
): Promise<{ r: number; g: number; b: number; a: number }> {
  const canvas = proxy.locator("xpath=preceding-sibling::canvas[1]");
  const [r, g, b, a] = await canvas.evaluate(
    (el, [xf, yf]) => {
      const canvasEl = el as HTMLCanvasElement;
      const ctx = canvasEl.getContext("2d")!;
      const x = Math.min(canvasEl.width - 1, Math.max(0, Math.floor(canvasEl.width * xf)));
      const y = Math.min(canvasEl.height - 1, Math.max(0, Math.floor(canvasEl.height * yf)));
      const data = ctx.getImageData(x, y, 1, 1).data;
      return [data[0], data[1], data[2], data[3]];
    },
    [xFrac, yFrac] as const,
  );
  return { r, g, b, a };
}

/** Whether ANY pixel on the canvas immediately preceding `proxy` is opaque
 * (alpha > 0) — used where a single sampled point can't reliably land on a
 * thin painted line (AC-15's outline-removal check). */
async function canvasHasOpaquePixel(proxy: Locator): Promise<boolean> {
  const canvas = proxy.locator("xpath=preceding-sibling::canvas[1]");
  return canvas.evaluate((el) => {
    const canvasEl = el as HTMLCanvasElement;
    const ctx = canvasEl.getContext("2d")!;
    const data = ctx.getImageData(0, 0, canvasEl.width, canvasEl.height).data;
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] > 0) {
        return true;
      }
    }
    return false;
  });
}

/** Registers the shared AC-63/AC-64 assertions against one fixture — the
 * two fixtures are structurally identical (same geometries, same document
 * order, same alt text), so this runs once per format instead of
 * duplicating every assertion. */
function registerFixtureTests(formatLabel: string, fixturePath: string, acTag: string): void {
  test.describe(`SPEC-051 — opening ${formatLabel}'s shape fixture (${acTag})`, () => {
    test(`renders all 18 shapes, one proxy per drawing, in document order (${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      await expect(drawings(page)).toHaveCount(18);
    });

    test(`renders every catalogue geometry with a focusable, correctly labelled ARIA proxy (${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      for (let i = 0; i < CATALOGUE_LABELS.length; i++) {
        const proxy = drawingAt(page, i);
        await expect(proxy, CATALOGUE_LABELS[i]).toBeVisible();
        await expect(proxy).toHaveAttribute("role", "img");
        await expect(proxy).toHaveAttribute("aria-label", `${CATALOGUE_LABELS[i]} shape`);
        await expect(proxy).toHaveAttribute("tabindex", "0");
      }
    });

    test(`a filled catalogue shape (rectangle) actually paints its fill colour (${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      const rect = drawingAt(page, 2); // index 2 = Rectangle, fill #3366cc
      const pixel = await canvasPixelAt(rect, 0.5, 0.5);
      expect(pixel.a).toBeGreaterThan(200);
      expect(pixel.r).toBeCloseTo(0x33, -1);
      expect(pixel.g).toBeCloseTo(0x66, -1);
      expect(pixel.b).toBeCloseTo(0xcc, -1);
    });

    test(`the grouped pair share group behaviour — selecting one member exposes Ungroup (${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      const memberA = drawingAt(page, 14);
      const memberB = drawingAt(page, 15);
      await expect(memberA).toHaveAttribute("aria-label", "Group shape A");
      await expect(memberB).toHaveAttribute("aria-label", "Group shape B");

      await memberA.click();
      await expect(page.getByTestId("shape-toolbar")).toBeVisible();
      await expect(page.getByTestId("shape-ungroup-button")).toBeVisible();
    });

    test(`the text box renders its own text and role=group (${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      const textBox = drawingAt(page, 16);
      await expect(textBox).toHaveAttribute("role", "group");
      await expect(textBox).toHaveAttribute("aria-label", "Text box");
      const text = textBox.locator('[data-testid="shape-text"]');
      await expect(text).toContainText("First paragraph in the text box.");
      await expect(text).toContainText("Second paragraph in the text box.");
    });

    test(`the unsupported geometry renders as a selectable placeholder with no fill painted (AC-60, ${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      const placeholder = drawingAt(page, 17);
      await expect(placeholder).toHaveAttribute("role", "img");
      await expect(placeholder).toHaveAttribute("aria-label", "Unsupported shape");
      // No mapped geometry is ever painted for it — the placeholder is a
      // dashed outline only (shapeCanvas.ts), so its interior stays empty
      // (alpha 0) instead of showing the #cc6600 fill the source XML set.
      const pixel = await canvasPixelAt(placeholder, 0.5, 0.5);
      expect(pixel.a).toBe(0);
      // Still selectable per AC-60 ("selectable and movable but not
      // re-geometried") — clicking it opens the same shape toolbar.
      await placeholder.click();
      await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    });
  });
}

registerFixtureTests("word-shapes.docx", docxShapesFixture, "AC-63");
registerFixtureTests("lo-shapes.odt", odtShapesFixture, "AC-64");

// --- interactive editing (fix-forward retry: AC-11, AC-17, AC-18, AC-27,
// AC-28, AC-34 — previously unreachable from the UI or wrong) -------------

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** "Insert -> Shape" (AC-3) then a single click on the canvas at the given
 * position places `geometry` at the default 3cm x 2cm size (AC-5). The
 * caret first goes to a fixed, shape-free corner so the insertion always
 * anchors near the same paragraph regardless of what an earlier insertion
 * placed. */
async function insertShapeAt(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

test.describe("SPEC-051 — interactive shape editing", () => {
  test("shift-clicking two shapes offers Group; grouping exposes Ungroup on either member (AC-34)", async ({
    page,
  }) => {
    await createDocument(page);
    // Well clear of the left margin/ruler, so neither shape's proxy ever
    // lands partly under the page furniture (a click there can hit the
    // wrong element).
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await expect(page.getByTestId("shape-group-button")).toHaveCount(0);
    await expect(drawings(page)).toHaveCount(2);
    // A fresh shape lands relative to the caret's own block, an atom like an
    // image (see shapes.ts's header comment) — inserting the second next to
    // the first can land it either side, so identify the two by their
    // left-to-right screen order rather than assume a fixed drawing index.
    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    const rightBox = firstBox.x <= secondBox.x ? secondBox : firstBox;
    // A plain click establishes a known single-selection baseline, then
    // Shift-click the other shape to build a two-member selection. Both
    // clicks land in the shape's LOWER half: the floating shape toolbar
    // renders across the page top overlapping the shapes' upper halves, and
    // under CPU load its mount/positioning races the click's hit-target
    // check — a centre click can land on the toolbar instead of the shape
    // (the factory-gate flake). 72% height is still well inside the painted
    // ellipse. Waiting for the toolbar between the clicks pins its final
    // position before the second click's hit-testing starts.
    await leftShape.click({ position: { x: firstBox.width / 2, y: firstBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({
      modifiers: ["Shift"],
      position: { x: rightBox.width / 2, y: rightBox.height * 0.72 },
    });
    await expect(page.getByTestId("shape-group-button")).toBeVisible();
    await page.getByTestId("shape-group-button").click();
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await expect(page.getByTestId("shape-ungroup-button")).toBeVisible();
    // The other member reports the same group once selected on its own.
    await leftShape.click();
    await expect(page.getByTestId("shape-ungroup-button")).toBeVisible();
  });

/** Groups the two insert-order-0/1 shapes and returns their locators,
 * left-to-right, with the group selected (leftShape single-selected, per
 * handleGroupSelected's own "leaves the group's first member selected"
 * contract — so leftShape is also what the resize/rotate handles measure
 * against). */
async function groupFirstTwoShapes(page: Page): Promise<{ leftShape: Locator; rightShape: Locator }> {
  const [first, second] = [drawingAt(page, 0), drawingAt(page, 1)];
  const firstBox = await first.boundingBox();
  const secondBox = await second.boundingBox();
  if (!firstBox || !secondBox) {
    throw new Error("a drawing has no bounding box");
  }
  const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
  const [leftBox, rightBox] = firstBox.x <= secondBox.x ? [firstBox, secondBox] : [secondBox, firstBox];
  // Lower-half clicks with the toolbar settled in between — see the AC-34
  // test's own comment: the floating toolbar overlaps the shapes' upper
  // halves, and under load a centre click can hit it instead of the shape.
  await leftShape.click({ position: { x: leftBox.width / 2, y: leftBox.height * 0.72 } });
  await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  await rightShape.click({
    modifiers: ["Shift"],
    position: { x: rightBox.width / 2, y: rightBox.height * 0.72 },
  });
  await expect(page.getByTestId("shape-group-button")).toBeVisible();
  await page.getByTestId("shape-group-button").click();
  await expect(page.getByTestId("shape-ungroup-button")).toBeVisible();
  return { leftShape, rightShape };
}

test.describe("SPEC-051 — group move/resize/rotate as one unit (AC-36..38)", () => {
  test("dragging a grouped member moves every member by the same delta, sizes untouched (AC-36)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    const { leftShape, rightShape } = await groupFirstTwoShapes(page);

    const beforeA = await leftShape.boundingBox();
    const beforeB = await rightShape.boundingBox();
    if (!beforeA || !beforeB) {
      throw new Error("a group member has no bounding box");
    }
    await leftShape.hover();
    await page.mouse.down();
    await page.mouse.move(beforeA.x + beforeA.width / 2 + 80, beforeA.y + beforeA.height / 2 + 60, { steps: 10 });
    await page.mouse.up();

    const afterA = await leftShape.boundingBox();
    const afterB = await rightShape.boundingBox();
    if (!afterA || !afterB) {
      throw new Error("a group member has no bounding box after the drag");
    }
    // Both members move by the same on-screen delta — relative positions
    // and sizes are preserved.
    expect(afterA.x - beforeA.x).toBeCloseTo(afterB.x - beforeB.x, 0);
    expect(afterA.y - beforeA.y).toBeCloseTo(afterB.y - beforeB.y, 0);
    expect(afterA.x).toBeGreaterThan(beforeA.x + 40);
    expect(afterA.width).toBeCloseTo(beforeA.width, 0);
    expect(afterB.width).toBeCloseTo(beforeB.width, 0);
  });

  test("dragging a resize handle scales every member about the group's bounding-box centre (AC-37)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    const { leftShape, rightShape } = await groupFirstTwoShapes(page);

    const beforeA = await leftShape.boundingBox();
    const beforeB = await rightShape.boundingBox();
    const handle = page.getByTestId("shape-resize-se");
    const handleBox = await handle.boundingBox();
    if (!beforeA || !beforeB || !handleBox) {
      throw new Error("missing bounding box");
    }
    await handle.hover();
    await page.mouse.down();
    await page.mouse.move(
      handleBox.x + handleBox.width / 2 + 60,
      handleBox.y + handleBox.height / 2 + 40,
      { steps: 8 },
    );
    await page.mouse.up();

    const afterA = await leftShape.boundingBox();
    const afterB = await rightShape.boundingBox();
    if (!afterA || !afterB) {
      throw new Error("a group member has no bounding box after the resize");
    }
    // Both members grow — the resize applies to the whole group, not just
    // the member whose own handle was dragged.
    expect(afterA.width).toBeGreaterThan(beforeA.width + 5);
    expect(afterB.width).toBeGreaterThan(beforeB.width + 5);
  });

  test("dragging the rotation handle rotates every member (AC-38)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await groupFirstTwoShapes(page);

    const handle = page.getByTestId("shape-rotation-handle");
    const handleBox = await handle.boundingBox();
    if (!handleBox) {
      throw new Error("rotation handle has no bounding box");
    }
    await handle.hover();
    await page.mouse.down();
    // Sweep the handle a quarter-turn around the shape's centre.
    await page.mouse.move(handleBox.x + 80, handleBox.y + 80, { steps: 8 });
    await page.mouse.up();

    // Both members carry a live rotation — read it back off the model via
    // the shape toolbar's rotation input, which reflects the currently
    // selected (first) member; the other member's own rotation is asserted
    // at the unit level (rotateGroup's own tests) since only one member is
    // ever individually selected at a time in this UI.
    const rotationInput = page.getByTestId("shape-rotation");
    await expect(rotationInput).toBeVisible();
    const rotation = Number(await rotationInput.inputValue());
    expect(rotation).not.toBe(0);
  });
});

  test("an unfilled, textless shape's interior is click-through, never selectable by bounding box (AC-11)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    await page.getByTestId("shape-fill-none").click();
    await page.getByTestId("shape-outline-none").click();
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await content(page).click({ position: { x: 550, y: 550 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // A click at the (now paint-free) shape's own bounding-box centre must
    // not select it — with nothing filled, no outline and no text, its
    // painted region is empty.
    await drawingAt(page, 0).click({ force: true });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
  });

  test("double-clicking a shape enters text edit mode and typed text persists in the model (AC-17, AC-18)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    await drawingAt(page, 0).dblclick();
    const text = drawingAt(page, 0).locator('[data-testid="shape-text"]');
    await expect(text).toHaveAttribute("contenteditable", "true");
    await page.keyboard.type("Hello shape");
    await expect(text).toContainText("Hello shape");
    // Click elsewhere to leave edit mode and re-render from the model —
    // the text surviving that round trip proves it landed in Shape.text,
    // not just the transient contenteditable DOM.
    await content(page).click({ position: { x: 550, y: 550 } });
    await expect(text).toContainText("Hello shape");
  });

  test("selecting a shape shows an anchor marker, and dragging it moves the shape (AC-27, AC-28)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    await expect(page.getByTestId("shape-anchor-marker")).toBeVisible();
    const proxy = drawingAt(page, 0);
    const before = await proxy.boundingBox();
    if (!before) {
      throw new Error("shape proxy has no bounding box");
    }
    await proxy.hover();
    await page.mouse.down();
    await page.mouse.move(before.x + before.width / 2 + 90, before.y + before.height / 2 + 90, { steps: 6 });
    await page.mouse.up();
    const after = await proxy.boundingBox();
    if (!after) {
      throw new Error("shape proxy lost its bounding box after the drag");
    }
    expect(after.x).toBeGreaterThan(before.x + 40);
    expect(after.y).toBeGreaterThan(before.y + 40);
  });

  test("drawing a shape previews it live, stretching with the pointer, and lands where it was dragged (AC-3, AC-4)", async ({
    page,
  }) => {
    await createDocument(page);
    const content = page.getByTestId("document-content");
    await content.click({ position: { x: 20, y: 20 } });
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("shape-gallery-ellipse").click();

    const preview = page.getByTestId("shape-draw-preview");
    await expect(preview).toHaveCount(0);

    const box = (await content.boundingBox())!;
    const startX = box.x + 150;
    const startY = box.y + 150;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Pressing already previews what releasing here would insert.
    await expect(preview).toBeVisible();

    await page.mouse.move(startX + 120, startY + 90, { steps: 6 });
    const small = (await preview.boundingBox())!;
    await page.mouse.move(startX + 240, startY + 180, { steps: 6 });
    const large = (await preview.boundingBox())!;
    // It stretches: same origin, bigger box.
    expect(large.width).toBeGreaterThan(small.width + 50);
    expect(large.height).toBeGreaterThan(small.height + 50);
    expect(large.x).toBeCloseTo(small.x, 0);
    expect(large.y).toBeCloseTo(small.y, 0);

    // The preview is the drag rectangle itself, not an approximation of it
    // — to within the whole pixel the browser quantises pointer
    // coordinates to.
    expect(Math.abs(large.x - startX)).toBeLessThanOrEqual(1);
    expect(Math.abs(large.y - startY)).toBeLessThanOrEqual(1);
    expect(Math.abs(large.width - 240)).toBeLessThanOrEqual(1);
    expect(Math.abs(large.height - 180)).toBeLessThanOrEqual(1);

    await page.mouse.up();
    await expect(preview).toHaveCount(0);

    // What was dragged out is the size AND position that landed —
    // drawGestureBox is the single source both the preview and the
    // insertion read, and SPEC-078's explicit anchors made the render
    // origin the anchor paragraph's own top (the same origin draw mode
    // measures against), so the vertical position is now exact too. This
    // used to be x-only, with a pinned ~14px vertical sag documented here;
    // SPEC-078 resolved it.
    const drawn = (await drawingAt(page, 0).boundingBox())!;
    expect(drawn.width).toBeCloseTo(large.width, 0);
    expect(drawn.height).toBeCloseTo(large.height, 0);
    expect(drawn.x).toBeCloseTo(large.x, 0);
    expect(drawn.y).toBeCloseTo(large.y, 0);
  });

  test("a resize drag stretches the shape and its handles live, not just on release (AC-4)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    const proxy = drawingAt(page, 0);
    const before = (await proxy.boundingBox())!;
    const handle = page.getByTestId("shape-resize-se");
    const handleBefore = (await handle.boundingBox())!;

    await handle.hover();
    await page.mouse.down();
    await page.mouse.move(handleBefore.x + 120, handleBefore.y + 90, { steps: 6 });

    // Mid-drag, before any mouseup: the painted shape is already the size
    // being dragged, and the handle the pointer is holding has come with it.
    const during = (await proxy.boundingBox())!;
    expect(during.width).toBeGreaterThan(before.width + 80);
    expect(during.height).toBeGreaterThan(before.height + 60);
    const handleDuring = (await handle.boundingBox())!;
    expect(handleDuring.x).toBeGreaterThan(handleBefore.x + 80);
    expect(handleDuring.y).toBeGreaterThan(handleBefore.y + 60);
    // The canvas is repainted at the dragged size, not a stale box stretched
    // by CSS: its own centre is still filled.
    expect((await canvasPixelAt(proxy, 0.5, 0.5)).a).toBeGreaterThan(200);

    await page.mouse.up();
    const after = (await proxy.boundingBox())!;
    expect(after.width).toBeCloseTo(during.width, 0);
    expect(after.height).toBeCloseTo(during.height, 0);
  });

  test("dragging a shape takes its handles along and hides the toolbar until the drop (AC-28)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    const toolbar = page.getByTestId("shape-toolbar");
    await expect(toolbar).toBeVisible();
    const proxy = drawingAt(page, 0);
    const before = (await proxy.boundingBox())!;
    const handleBefore = (await page.getByTestId("shape-resize-nw").boundingBox())!;

    await proxy.hover();
    await page.mouse.down();
    await page.mouse.move(before.x + before.width / 2 + 110, before.y + before.height / 2 + 70, { steps: 6 });

    // The toolbar is out of the way for the duration of the drag...
    await expect(toolbar).toHaveCount(0);
    // ...and the handles have travelled with the shape rather than staying
    // behind, outlining where it no longer is.
    const handleDuring = (await page.getByTestId("shape-resize-nw").boundingBox())!;
    expect(handleDuring.x - handleBefore.x).toBeCloseTo(110, 0);
    expect(handleDuring.y - handleBefore.y).toBeCloseTo(70, 0);

    await page.mouse.up();
    await expect(toolbar).toBeVisible();
  });

  test("rotating keeps the shape's whole form — no corner cropped off (AC-14)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    const proxy = drawingAt(page, 0);
    const box = (await proxy.boundingBox())!;

    await page.getByTestId("shape-rotation").fill("35");
    await page.getByTestId("shape-rotation").press("Enter");

    // The canvas grows to the rotated bounding box and shifts back by half
    // the growth, so the shape's own box stays where it was. Painted into
    // the unrotated box instead, a rotated rectangle loses all four corners
    // and comes out an octagon.
    const surface = await proxy.evaluate((el) => {
      const canvas = el.previousElementSibling as HTMLCanvasElement;
      const rect = canvas.getBoundingClientRect();
      return { width: rect.width, height: rect.height, left: canvas.style.left, top: canvas.style.top };
    });
    const rad = (35 * Math.PI) / 180;
    const expectedW = box.width * Math.cos(rad) + box.height * Math.sin(rad);
    const expectedH = box.width * Math.sin(rad) + box.height * Math.cos(rad);
    expect(Math.abs(surface.width - expectedW)).toBeLessThanOrEqual(2);
    expect(Math.abs(surface.height - expectedH)).toBeLessThanOrEqual(2);
    expect(Number.parseFloat(surface.left)).toBeLessThan(0);
    expect(Number.parseFloat(surface.top)).toBeLessThan(0);

    // Every rotated corner is actually painted: sample just inside each one,
    // in the padded surface's own coordinates.
    const cx = surface.width / 2;
    const cy = surface.height / 2;
    for (const [sx, sy] of [
      [-1, -1],
      [1, -1],
      [1, 1],
      [-1, 1],
    ]) {
      const dx = (sx * box.width) / 2;
      const dy = (sy * box.height) / 2;
      // 85% of the way out to the corner — inside the fill, past where the
      // unrotated box would have clipped it.
      const x = cx + 0.85 * (dx * Math.cos(rad) - dy * Math.sin(rad));
      const y = cy + 0.85 * (dx * Math.sin(rad) + dy * Math.cos(rad));
      const pixel = await canvasPixelAt(proxy, x / surface.width, y / surface.height);
      expect(pixel.a, `corner ${sx},${sy}`).toBeGreaterThan(200);
    }
  });

  test("the resize and rotation handles ride round with a rotated shape (AC-10, AC-13, AC-16)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    const proxy = drawingAt(page, 0);
    const box = (await proxy.boundingBox())!;
    const deg = 35;
    await page.getByTestId("shape-rotation").fill(String(deg));
    await page.getByTestId("shape-rotation").press("Enter");

    // The proxy never rotates — it stays the shape's upright box, which is
    // what every measurement and hit test is built on — so handles left at
    // that box's own corners end up nowhere near the shape a user sees.
    const after = (await proxy.boundingBox())!;
    expect(after.width).toBeCloseTo(box.width, 0);
    expect(after.height).toBeCloseTo(box.height, 0);
    const cx = after.x + after.width / 2;
    const cy = after.y + after.height / 2;
    const rad = (deg * Math.PI) / 180;
    const rotated = (dx: number, dy: number) => ({
      x: cx + dx * Math.cos(rad) - dy * Math.sin(rad),
      y: cy + dx * Math.sin(rad) + dy * Math.cos(rad),
    });
    const centreOf = async (testid: string) => {
      const handle = (await page.getByTestId(testid).boundingBox())!;
      return { x: handle.x + handle.width / 2, y: handle.y + handle.height / 2 };
    };

    for (const [testid, sx, sy] of [
      ["shape-resize-nw", -1, -1],
      ["shape-resize-ne", 1, -1],
      ["shape-resize-se", 1, 1],
      ["shape-resize-sw", -1, 1],
    ] as const) {
      const expected = rotated((sx * after.width) / 2, (sy * after.height) / 2);
      const actual = await centreOf(testid);
      expect(Math.hypot(actual.x - expected.x, actual.y - expected.y), testid).toBeLessThanOrEqual(2);
    }

    // The rotation handle hangs off the middle of the shape's OWN top edge,
    // so it swings round with it too — at 35° that is up and to the right.
    const rotationHandle = await centreOf("shape-rotation-handle");
    const topEdge = rotated(0, -after.height / 2);
    expect(rotationHandle.x).toBeGreaterThan(topEdge.x);
    expect(rotationHandle.y).toBeLessThan(topEdge.y);
    // ...and it stays on the line the shape's own "up" runs along: the
    // handle, the top edge's midpoint and the centre are collinear.
    const cross =
      (topEdge.x - cx) * (rotationHandle.y - cy) - (topEdge.y - cy) * (rotationHandle.x - cx);
    expect(Math.abs(cross)).toBeLessThanOrEqual(80);

    // Each corner handle is on the painted shape's own corner — the fill is
    // right there under it — while the upright box's corners, where the
    // handles used to sit, are over the empty padding the rotated shape
    // leans out of.
    const surface = await proxy.evaluate((el) => {
      const canvas = el.previousElementSibling as HTMLCanvasElement;
      const rect = canvas.getBoundingClientRect();
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    });
    const pixelAtPoint = (point: { x: number; y: number }) =>
      canvasPixelAt(proxy, (point.x - surface.x) / surface.width, (point.y - surface.y) / surface.height);
    for (const [testid, sx, sy] of [
      ["shape-resize-nw", -1, -1],
      ["shape-resize-ne", 1, -1],
      ["shape-resize-se", 1, 1],
      ["shape-resize-sw", -1, 1],
    ] as const) {
      const handle = await centreOf(testid);
      const length = Math.hypot(cx - handle.x, cy - handle.y);
      const inside = {
        x: handle.x + ((cx - handle.x) / length) * 6,
        y: handle.y + ((cy - handle.y) / length) * 6,
      };
      expect((await pixelAtPoint(inside)).a, `${testid} sits on painted shape`).toBeGreaterThan(200);
      const uprightCorner = {
        x: cx + (sx * (after.width / 2 - 4)),
        y: cy + (sy * (after.height / 2 - 4)),
      };
      expect((await pixelAtPoint(uprightCorner)).a, `${testid}'s old upright corner`).toBeLessThan(50);
    }

    // The cursor turns with the handle, so it still points along the edge
    // the drag will move: a quarter turn takes "n" two compass points on.
    await page.getByTestId("shape-rotation").fill("90");
    await page.getByTestId("shape-rotation").press("Enter");
    await expect
      .poll(() => page.getByTestId("shape-resize-n").evaluate((el) => getComputedStyle(el).cursor))
      .toBe("e-resize");
  });

  test("dragging a rotated shape's handle sizes it along its own edges, not the screen's (AC-4, AC-13)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    const proxy = drawingAt(page, 0);
    const before = (await proxy.boundingBox())!;
    await page.getByTestId("shape-rotation").fill("90");
    await page.getByTestId("shape-rotation").press("Enter");

    // At a quarter turn the shape's own width runs DOWN the screen, so a
    // drag straight down the "e" handle is a drag along its width — it must
    // widen the shape and leave the height alone. Taken raw, that same drag
    // reads as pure vertical travel and grows the height instead.
    const handle = (await page.getByTestId("shape-resize-e").boundingBox())!;
    await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
    await page.mouse.down();
    await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2 + 40, { steps: 8 });
    await page.mouse.up();

    const after = (await proxy.boundingBox())!;
    expect(after.width - before.width).toBeGreaterThan(30);
    expect(Math.abs(after.height - before.height)).toBeLessThanOrEqual(2);
  });

  test("the floating toolbar sits fully clear above the shape, rotation handle included (AC-12, AC-16)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    const toolbar = page.getByTestId("shape-toolbar");
    await expect(toolbar).toBeVisible();
    const toolbarBox = await toolbar.boundingBox();
    const shapeBox = await drawingAt(page, 0).boundingBox();
    const handleBox = await page.getByTestId("shape-rotation-handle").boundingBox();
    if (!toolbarBox || !shapeBox || !handleBox) {
      throw new Error("the toolbar, shape or rotation handle has no bounding box");
    }
    // The toolbar wraps to more than one row, so a fixed offset from the
    // shape's top edge used to leave its lower rows lying across the shape
    // and over the rotation handle — the handle stayed visible but never
    // received the drag that starts a rotation.
    expect(toolbarBox.height).toBeGreaterThan(40);
    expect(toolbarBox.y + toolbarBox.height).toBeLessThanOrEqual(handleBox.y);
    expect(toolbarBox.y + toolbarBox.height).toBeLessThanOrEqual(shapeBox.y);
    // Nothing of the toolbar is pushed off the top of the viewport either.
    expect(toolbarBox.y).toBeGreaterThanOrEqual(0);
  });

  test("Insert -> Text box lands already in text-editing mode with the caret inside (AC-9, AC-51)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-textbox-button").click();
    const proxy = drawingAt(page, 0);
    const text = proxy.locator('[data-testid="shape-text"]');
    await expect(text).toHaveAttribute("contenteditable", "true");
    await page.keyboard.type("New text box");
    await expect(text).toContainText("New text box");
    // block.text is now populated, so the proxy's role switches from img
    // to group (AC-51).
    await expect(proxy).toHaveAttribute("role", "group");
  });

  test("In line with text keeps the painted canvas, the handles and undo in step (regression)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rightArrow", 400, 250);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    const proxy = drawingAt(page, 0);
    const canvas = proxy.locator("xpath=preceding-sibling::canvas[1]");

    // Switching a floating shape into the text flow used to break twice
    // over: .doc-shape lost its positioned ancestor in inline mode, so the
    // absolutely positioned canvas resolved against the page and painted the
    // geometry at the page's top-left corner — while the proxy (and the
    // selection handles measured from it) sat correctly in the line; and the
    // wrap change never re-measured the handles at all, leaving them at the
    // shape's old floating position.
    await page.getByTestId("shape-wrap").selectOption("inline");
    await expect(page.getByTestId("shape-wrap")).toHaveValue("inline");

    const proxyBox = (await proxy.boundingBox())!;
    const canvasBox = (await canvas.boundingBox())!;
    expect(Math.abs(canvasBox.x - proxyBox.x)).toBeLessThanOrEqual(1.5);
    expect(Math.abs(canvasBox.y - proxyBox.y)).toBeLessThanOrEqual(1.5);
    const se = (await page.getByTestId("shape-resize-se").boundingBox())!;
    expect(Math.abs(se.x + se.width / 2 - (proxyBox.x + proxyBox.width))).toBeLessThanOrEqual(6);
    expect(Math.abs(se.y + se.height / 2 - (proxyBox.y + proxyBox.height))).toBeLessThanOrEqual(6);

    // Undo moves the shape back to its floating spot — the handles and the
    // canvas must follow there too (the history-restore path used to skip
    // the re-measure, stranding the handles at the inline position).
    await page.getByTestId("undo-button").click();
    await expect(page.getByTestId("shape-wrap")).toHaveValue("square");
    const proxyAfter = (await proxy.boundingBox())!;
    const canvasAfter = (await canvas.boundingBox())!;
    expect(Math.abs(canvasAfter.x - proxyAfter.x)).toBeLessThanOrEqual(1.5);
    expect(Math.abs(canvasAfter.y - proxyAfter.y)).toBeLessThanOrEqual(1.5);
    const seAfter = (await page.getByTestId("shape-resize-se").boundingBox())!;
    expect(Math.abs(seAfter.x + seAfter.width / 2 - (proxyAfter.x + proxyAfter.width))).toBeLessThanOrEqual(6);
    expect(Math.abs(seAfter.y + seAfter.height / 2 - (proxyAfter.y + proxyAfter.height))).toBeLessThanOrEqual(6);
  });

  test("square and topAndBottom wrap actually keep text clear of the shape; infront overlaps (regression)", async ({
    page,
  }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type(
      "The quick brown fox jumps over the lazy dog while the calm river flows past the old stone bridge. " +
        "Every morning the village bakery fills the narrow street with the warm smell of fresh bread and coffee. " +
        "Children hurry along the cobblestones on their way to school, greeting the shopkeepers who sweep their doorsteps. " +
        "Travellers who pass through often stop to rest under the plane trees and watch the light change on the hills.",
    );
    // Draw the shape ONTO the paragraph. The caret sits at the text's end,
    // in the same paragraph, so the drawn offsets and the rendered anchor
    // must agree — the shape lands at the pointer (this used to be
    // displaced by the anchor paragraph's height: offsets were measured
    // from the paragraph's top but rendered from the line after it).
    const paraBox = (await content(page).locator(".doc-paragraph").first().boundingBox())!;
    const drawX = paraBox.x + 150;
    const drawY = paraBox.y + 30;
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("shape-gallery-rect").click();
    await page.mouse.click(drawX, drawY);
    const proxy = drawingAt(page, 0);
    const box = (await proxy.boundingBox())!;
    // A single click places the default-size box with its top-left corner at
    // the pointer (drawGestureBox's click behaviour).
    expect(Math.abs(box.x - drawX)).toBeLessThanOrEqual(8);
    expect(Math.abs(box.y - drawY)).toBeLessThanOrEqual(8);

    /** The paragraph's line boxes that intersect the shape's own box. */
    async function overlappingLines(): Promise<number> {
      const shapeBox = (await drawingAt(page, 0).boundingBox())!;
      return page.evaluate((sb) => {
        const para = document.querySelector('[data-testid="document-content"] .doc-paragraph');
        if (!para) {
          return -1;
        }
        const range = document.createRange();
        range.selectNodeContents(para);
        let count = 0;
        for (const rect of range.getClientRects()) {
          if (rect.width < 2 || rect.height < 2) {
            continue;
          }
          const vertical = rect.bottom > sb.y + 2 && rect.top < sb.y + sb.height - 2;
          const horizontal = rect.right > sb.x + 2 && rect.left < sb.x + sb.width - 2;
          if (vertical && horizontal) {
            count += 1;
          }
        }
        return count;
      }, shapeBox);
    }

    // Square (the drawn default): the wrap spacer exists and no line box
    // runs under the shape.
    await expect(page.getByTestId("shape-wrap-spacer")).toHaveCount(1);
    expect(await overlappingLines()).toBe(0);

    // Top and bottom: the full band is clear too.
    await page.getByTestId("shape-wrap").selectOption("topAndBottom");
    await expect(page.getByTestId("shape-wrap")).toHaveValue("topAndBottom");
    expect(await overlappingLines()).toBe(0);

    // In front of text: no spacer, text overlaps by design — proving the
    // wrap modes actually render differently from one another.
    await page.getByTestId("shape-wrap").selectOption("infront");
    await expect(page.getByTestId("shape-wrap-spacer")).toHaveCount(0);
    expect(await overlappingLines()).toBeGreaterThan(0);
  });

  test("removing a Line's outline stops it painting (AC-15)", async ({ page }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("shape-gallery-line").click();
    await content(page).click({ position: { x: 300, y: 300 } });
    const proxy = drawingAt(page, 0);
    await expect(canvasHasOpaquePixel(proxy)).resolves.toBe(true);
    await page.getByTestId("shape-outline-none").click();
    await expect(canvasHasOpaquePixel(proxy)).resolves.toBe(false);
  });

  test("the overflow indicator shows only while a shape with clipped text is selected (AC-22, AC-23)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 250, 250);
    const proxy = drawingAt(page, 0);
    await proxy.dblclick();
    const text = proxy.locator('[data-testid="shape-text"]');
    await expect(text).toHaveAttribute("contenteditable", "true");
    // 15 short lines overflow a 3x2cm box (matches AC-21's own coverage).
    for (let i = 0; i < 15; i++) {
      await page.keyboard.type(`line ${i}`);
      await page.keyboard.press("Enter");
    }
    const indicator = page.getByTestId("shape-overflow-indicator");
    await expect(indicator).toBeVisible();
    // Deselecting hides it even though the text is still clipped (AC-23).
    await content(page).click({ position: { x: 550, y: 550 } });
    await expect(indicator).toHaveCount(0);
  });

  test("z-order commands change which shape paints on top (AC-29..33)", async ({ page }) => {
    await createDocument(page);
    // Two overlapping filled shapes so one visibly occludes the other.
    await insertShapeAt(page, "rect", 400, 300);
    await insertShapeAt(page, "rect", 430, 300);
    // Insertion leaves the new shape selected, and the floating shape
    // toolbar (a page-level overlay, always painted above a page's own
    // content) sits right above a small shape's top edge — close enough to
    // the two rects' own overlap region to cover it. Deselect first (via
    // Escape, not a document click, since a click at a fixed viewport
    // position can itself scroll the page and invalidate the bounding
    // boxes read below) so the baseline z-order check hits the shapes
    // themselves, not the toolbar.
    await page.keyboard.press("Escape");
    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    // The actual pixel overlap between the two shapes' boxes.
    const overlapPoint = {
      x: Math.max(firstBox.x, secondBox.x) + 5,
      y: Math.max(firstBox.y, secondBox.y) + 5,
    };
    const topmostTestId = () =>
      page.evaluate(
        (p) => document.elementFromPoint(p.x, p.y)?.closest("[data-testid^='drawing-']")?.getAttribute("data-testid") ?? null,
        overlapPoint,
      );
    // The later-inserted shape (second) paints on top by default (higher
    // zOrder from insertion order).
    await expect.poll(topmostTestId).toBe(await second.getAttribute("data-testid"));
    // Select the bottom shape and send it to the front.
    await first.click({ position: { x: 5, y: 5 }, force: true });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await page.getByTestId("shape-zorder").selectOption("front");
    // Deselect again (Escape, for the same scroll-safety reason as above) so
    // the toolbar doesn't shadow the repainted stacking order.
    await page.keyboard.press("Escape");
    await expect.poll(topmostTestId).toBe(await first.getAttribute("data-testid"));
  });
});

test.describe("SPEC-068 — drawing handle degrades to a numeric position dialog (AC-42)", () => {
  test("the shape toolbar's Position & size dialog resizes and repositions a shape numerically", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 300);
    // Insertion leaves the new shape selected with its toolbar up.
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    const drawing = drawingAt(page, 0);
    const before = await drawing.boundingBox();
    if (!before) {
      throw new Error("the drawing has no bounding box");
    }

    // Open the numeric fallback — the coarse-pointer path for a handle drag.
    await page.getByTestId("shape-position-button").click();
    await expect(page.getByTestId("shape-position-dialog")).toBeVisible();

    // Enlarge the shape well beyond the 3cm x 2cm default. Changes apply LIVE
    // as each field is committed (SPEC-072 AC-43) — there is no OK button.
    await page.getByTestId("shape-position-width").fill("12");
    await page.getByTestId("shape-position-width").press("Tab");
    await page.getByTestId("shape-position-height").fill("8");
    await page.getByTestId("shape-position-height").press("Tab");

    // The live change shows up as a visibly larger rendered drawing.
    await expect
      .poll(async () => (await drawing.boundingBox())?.width ?? 0)
      .toBeGreaterThan(before.width + 20);

    // Close, then re-open; the dialog reflects the persisted numeric size.
    await page.getByTestId("shape-position-done").click();
    await expect(page.getByTestId("shape-position-dialog")).toHaveCount(0);
    await page.getByTestId("shape-position-button").click();
    await expect(page.getByTestId("shape-position-dialog")).toBeVisible();
    await expect(page.getByTestId("shape-position-width")).toHaveValue("12");
    await expect(page.getByTestId("shape-position-height")).toHaveValue("8");
  });
});
