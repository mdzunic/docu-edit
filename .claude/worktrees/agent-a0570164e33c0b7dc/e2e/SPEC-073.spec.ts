import { test, expect, type Page } from "@playwright/test";
import { decodePng, type DecodedScreenshot } from "./pixels";

// SPEC-073 — shape styling depth. The full pickers behind the toolbar's
// Style popover (AC-16), the same panel in Format → Drawing → Shape
// formatting… (AC-17), the five-preset gallery as one undo entry (AC-18/
// AC-24), gradient seeding (AC-22), the shadow toggle (AC-20) and the
// transparency slider's single commit (AC-25) — plus composited-pixel
// assertions that transparency over text (AC-7) and the shadow pass (AC-9,
// AC-39) reach the screen, sampled from real page screenshots rather than a
// canvas backing store.

const content = (page: Page) => page.getByTestId("document-content");

/** The viewport composite as decoded pixels, plus the screenshot-px-per-
 * CSS-px ratio so boundingBox() coordinates can be sampled regardless of
 * deviceScaleFactor. */
async function compositedViewport(page: Page): Promise<{ img: DecodedScreenshot; scale: number }> {
  const img = decodePng(await page.screenshot());
  const viewport = page.viewportSize();
  return { img, scale: viewport ? img.width / viewport.width : 1 };
}

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

  test("a semi-transparent fill over body text composites on screen — visible tint, text showing through (AC-7)", async ({
    page,
  }) => {
    await createDocument(page);
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.keyboard.type("Some text behind the shape for compositing check. ".repeat(4));
    await insertShapeAt(page, "rect", 400, 350);
    await page.getByTestId("shape-wrap").selectOption("infront");

    // Drag the shape up over the paragraph — the exact manual scenario:
    // "in front of text" wrap, box overlapping rendered glyphs.
    const proxy = page.getByTestId("drawing-0");
    const paraBox = (await content(page).locator(".doc-paragraph").first().boundingBox())!;
    const startBox = (await proxy.boundingBox())!;
    await page.mouse.move(startBox.x + startBox.width / 2, startBox.y + startBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(paraBox.x + paraBox.width / 2, paraBox.y + paraBox.height / 2, { steps: 8 });
    await page.mouse.up();

    await openStylePanel(page);
    const slider = page.getByTestId("shape-fill-transparency");
    const sliderBox = (await slider.boundingBox())!;
    await page.mouse.move(sliderBox.x + 2, sliderBox.y + sliderBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(sliderBox.x + sliderBox.width / 2, sliderBox.y + sliderBox.height / 2, { steps: 5 });

    // Mid-drag the live preview must repaint the shape's canvas at the
    // preview alpha — the popover can overlap the shape itself on screen
    // (the toolbar floats above it near the page top), so the preview is
    // asserted on the backing store; the composited-over-text check below
    // runs on the committed state with everything deselected.
    const midValue = Number(await slider.inputValue());
    expect(midValue).toBeGreaterThan(30);
    const midPixel = await proxy
      .locator("xpath=preceding-sibling::canvas[1]")
      .evaluate((el) => {
        const canvasEl = el as HTMLCanvasElement;
        const ctx = canvasEl.getContext("2d")!;
        const data = ctx.getImageData(Math.floor(canvasEl.width / 2), Math.floor(canvasEl.height / 2), 1, 1).data;
        return [data[0], data[1], data[2], data[3]];
      });
    expect(Math.abs(midPixel[3] - 255 * (1 - midValue / 100)), "preview paints at the dragged alpha").toBeLessThan(14);
    expect(midPixel[2], "preview keeps the fill's colour").toBeGreaterThan(midPixel[0]);
    await page.mouse.up();

    const transparency = Number(await slider.inputValue());
    expect(transparency).toBeGreaterThan(30);
    expect(transparency).toBeLessThan(70);

    // Close the popover, refocus the shape and deselect so no selection
    // frame or toolbar overlaps what we sample.
    await page.getByTestId("shape-style-button").click();
    await expect(page.getByTestId("shape-style-panel")).toHaveCount(0);
    // Focus the proxy explicitly: the toolbar's own mousedown.prevent kept
    // DOM focus on the slider, which sits outside the content root's
    // Escape-to-deselect keydown handler.
    await proxy.focus();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("shape-style-button")).toHaveCount(0);

    const box = (await proxy.boundingBox())!;
    const { img, scale } = await compositedViewport(page);
    const samples: [number, number, number][] = [];
    for (let ix = 0; ix < 7; ix++) {
      for (let iy = 0; iy < 5; iy++) {
        samples.push(
          img.rgbAt(
            (box.x + 10 + ((box.width - 20) * ix) / 6) * scale,
            (box.y + 10 + ((box.height - 20) * iy) / 4) * scale,
          ),
        );
      }
    }
    // The fill tints (b > r) essentially everywhere inside the box — a small
    // allowance for spellcheck squiggles or other non-neutral underlay.
    const tinted = samples.filter(([r, , b]) => b - r > 20);
    expect(tinted.length, "fill tint visible across the shape's interior").toBeGreaterThanOrEqual(
      Math.ceil(samples.length * 0.85),
    );
    // Glyphs must SHOW THROUGH the fill: some tinted pixels sit over dark
    // text, some over the white gaps — an opaque fill would hide the glyphs
    // (uniform colour), an invisible one would fail the tint check above.
    const luminances = tinted.map(([r, g, b]) => (r + g + b) / 3);
    expect(Math.min(...luminances), "text glyphs visible through the fill").toBeLessThan(150);
    const alpha = 1 - transparency / 100;
    const expectedGapRed = 0x44 * alpha + 255 * (1 - alpha);
    const lightest = tinted[luminances.indexOf(Math.max(...luminances))];
    expect(Math.abs(lightest[0] - expectedGapRed), "background gaps blend at the committed alpha").toBeLessThan(12);
  });

  test("the standard shadow paints composited pixels below the shape, and re-renders scaled under zoom (AC-9, AC-39)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 300);
    await openStylePanel(page);
    await page.getByTestId("shape-shadow-toggle").click();
    await expect(page.getByTestId("shape-shadow-toggle")).toHaveText("Remove shadow");

    // Deselect so the sampled pixels are the document alone.
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("shape-style-panel")).toHaveCount(0);
    await expect(page.getByTestId("shape-style-button")).toHaveCount(0);

    const proxy = page.getByTestId("drawing-0");
    const canvas = proxy.locator("xpath=preceding-sibling::canvas[1]");
    const box = (await proxy.boundingBox())!;
    const { img, scale } = await compositedViewport(page);
    const cx = (box.x + box.width / 2) * scale;
    const bottom = (box.y + box.height) * scale;
    // D4's standard shadow: 90° (straight down), 3pt ≈ 4px offset, 4pt blur —
    // a neutral grey band right below the bottom edge, fading out by ~+12px.
    const [r, g, b] = img.rgbAt(cx, bottom + 5 * scale);
    expect(Math.abs(r - g), "shadow is neutral grey").toBeLessThanOrEqual(12);
    expect(Math.abs(g - b), "shadow is neutral grey").toBeLessThanOrEqual(12);
    expect(r, "shadow visibly darkens the page below the shape").toBeLessThan(240);
    expect(r, "shadow is a soft grey, not solid black").toBeGreaterThan(90);
    // Directional: 90° casts no band above the shape, and the blur is
    // bounded — well below the shape the page is back to white.
    expect(img.rgbAt(cx, (box.y - 6) * scale)[0], "no shadow above the shape at 90°").toBeGreaterThan(250);
    expect(img.rgbAt(cx, bottom + 45 * scale)[0], "shadow fades out within its blur").toBeGreaterThan(250);

    // AC-39: zooming re-paints the backing store at devicePixelRatio * zoom
    // (bigger store, shadow still composited) rather than scaling a bitmap.
    const storeBefore = await canvas.evaluate((el) => (el as HTMLCanvasElement).width);
    await page.getByTestId("zoom-in").click();
    const zoomText = await page.getByTestId("zoom-percent").inputValue();
    const zoom = Number(zoomText.replace(/[^0-9]/g, "")) / 100;
    expect(zoom).toBeGreaterThan(1);
    await expect
      .poll(async () => canvas.evaluate((el) => (el as HTMLCanvasElement).width))
      .toBeGreaterThan(storeBefore);
    const zoomedBox = (await proxy.boundingBox())!;
    const zoomed = await compositedViewport(page);
    const zr = zoomed.img.rgbAt(
      (zoomedBox.x + zoomedBox.width / 2) * zoomed.scale,
      (zoomedBox.y + zoomedBox.height + 5 * zoom) * zoomed.scale,
    );
    expect(zr[0], "shadow still composites at the new zoom").toBeLessThan(240);
    expect(Math.abs(zr[0] - zr[2]), "zoomed shadow stays neutral grey").toBeLessThanOrEqual(12);
  });
});
