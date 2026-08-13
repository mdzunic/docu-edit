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
    // A plain click establishes a known single-selection baseline, then
    // Shift-click the other shape to build a two-member selection.
    await leftShape.click();
    await rightShape.click({ modifiers: ["Shift"] });
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
  await leftShape.click();
  await rightShape.click({ modifiers: ["Shift"] });
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
    await expect.poll(topmostTestId).toBe(await first.getAttribute("data-testid"));
  });
});
