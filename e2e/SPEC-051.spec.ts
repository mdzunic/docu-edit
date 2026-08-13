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
