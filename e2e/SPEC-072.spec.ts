import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { unzip } from "../src/lib/core/zip";
import { cmToEmu } from "../src/lib/core/units";

// SPEC-072 — align and distribute produce EXACT model geometry, not
// approximate. Three shapes are placed at arbitrary (not pre-aligned, not
// pre-spaced) positions, aligned left, then distributed vertically; their model
// offsets are read back from the exported .docx (`<wp:posOffset>`, in EMU) and
// asserted for exact equality / an exact arithmetic progression — no tolerance.
// The chosen vertical positions (2cm, 8cm) make the ideal even spacing land on
// whole EMU, so exactness is real rather than lucky rounding.

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');

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

/** Clicks a drawing's LOWER half (the floating toolbar overlaps the upper half
 * — see SPEC-051), optionally Shift-extending the selection. */
async function clickDrawing(page: Page, index: number, shift = false): Promise<void> {
  const drawing = drawings(page).nth(index);
  const box = (await drawing.boundingBox())!;
  await drawing.click({
    modifiers: shift ? ["Shift"] : [],
    position: { x: box.width / 2, y: box.height * 0.72 },
  });
}

/** Selects one drawing (by document-order index) and sets its exact horizontal
 * and vertical position through the Position & size dialog's live fields. */
async function setPosition(page: Page, index: number, xCm: number, yCm: number): Promise<void> {
  await clickDrawing(page, index);
  await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  await page.getByTestId("shape-position-button").click();
  await expect(page.getByTestId("shape-position-dialog")).toBeVisible();
  // Pin a common height so the distributed top offsets form an arithmetic
  // progression (equal-height shapes => equal gaps <=> equal top spacing).
  await page.getByTestId("shape-position-height").fill("2");
  await page.getByTestId("shape-position-height").press("Tab");
  await page.getByTestId("shape-position-x").fill(String(xCm));
  await page.getByTestId("shape-position-x").press("Tab");
  await page.getByTestId("shape-position-y").fill(String(yCm));
  await page.getByTestId("shape-position-y").press("Tab");
  await page.getByTestId("shape-position-done").click();
}

/** Exports to .docx and returns each drawing's model offset in EMU, in document
 * order. The docx writer serialises a floating drawing's offset as
 * `<wp:positionH…><wp:posOffset>` / `<wp:positionV…>` (src/lib/docx/shape.ts). */
async function readOffsetsEmu(page: Page): Promise<Array<{ x: number; y: number }>> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const parts = await unzip(readFileSync((await download.path())!));
  const xml = new TextDecoder().decode(parts.get("word/document.xml")!);
  const pattern =
    /<wp:positionH[^>]*><wp:posOffset>(-?\d+)<\/wp:posOffset><\/wp:positionH><wp:positionV[^>]*><wp:posOffset>(-?\d+)<\/wp:posOffset>/g;
  return [...xml.matchAll(pattern)].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }));
}

test.describe("SPEC-072 — exact align and distribute", () => {
  test("align-left then distribute-vertical yields exact equal / evenly-spaced offsets", async ({ page }) => {
    await createDocument(page);
    // Three shapes at distinct, non-overlapping screen positions.
    await insertShapeAt(page, "rect", 320, 220);
    await insertShapeAt(page, "ellipse", 540, 220);
    await insertShapeAt(page, "hexagon", 400, 420);
    await expect(drawings(page)).toHaveCount(3);

    // Arbitrary positions: horizontals differ (not pre-aligned), verticals are
    // unevenly spaced (2, 3, 8 — gaps 1 and 5). The outer verticals (2cm, 8cm)
    // and the ideal middle (5cm) are all whole EMU.
    await setPosition(page, 0, 1, 2);
    await setPosition(page, 1, 4, 3);
    await setPosition(page, 2, 2, 8);

    // Select all three (single, then Shift-extend) and act through the
    // contextual toolbar's Align / Distribute controls (AC-15).
    await clickDrawing(page, 0);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await clickDrawing(page, 1, true);
    await clickDrawing(page, 2, true);

    await page.getByTestId("shape-align").selectOption("left");
    await page.getByTestId("shape-distribute").selectOption("vertical");

    const offsets = await readOffsetsEmu(page);
    expect(offsets).toHaveLength(3);

    // Align-left: every horizontal offset is exactly equal (they all take the
    // leftmost's offset, 1cm) — asserted with toBe, no tolerance.
    const xs = offsets.map((o) => o.x);
    expect(xs[0]).toBe(cmToEmu(1));
    expect(xs[1]).toBe(cmToEmu(1));
    expect(xs[2]).toBe(cmToEmu(1));

    // Distribute-vertical: the successive top offsets form an exact arithmetic
    // progression; the two integer-EMU differences are exactly equal.
    const ys = offsets.map((o) => o.y).sort((a, b) => a - b);
    expect(ys[1] - ys[0]).toBe(ys[2] - ys[1]);
    // The two outermost shapes did not move (2cm and 8cm).
    expect(ys[0]).toBe(cmToEmu(2));
    expect(ys[2]).toBe(cmToEmu(8));
    // The middle landed on the exact midpoint, 5cm.
    expect(ys[1]).toBe(cmToEmu(5));
  });
});
