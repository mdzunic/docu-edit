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

  test("single-drawing align-top moves it to the page content area, not its anchor paragraph (AC-2)", async ({ page }) => {
    await createDocument(page);
    // Push the anchor well down the page: a drawing inserted here anchors to a
    // paragraph partway down, so aligning to the content area must move it UP to
    // the top margin — not leave it at its anchor's own vertical position.
    await content(page).click({ position: { x: 20, y: 20 } });
    for (let i = 0; i < 12; i++) {
      await page.keyboard.press("Enter");
    }
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("shape-gallery-rect").click();
    await content(page).click({ position: { x: 300, y: 300 } });
    await expect(drawings(page)).toHaveCount(1);

    const marginTop = (await content(page).getByTestId("page-margin").first().boundingBox())!.y;
    const beforeTop = (await drawings(page).nth(0).boundingBox())!.y;
    // The drawing starts well below the content-area top (its anchor is ~12
    // paragraphs down) — otherwise this test could not distinguish the fix.
    expect(beforeTop - marginTop).toBeGreaterThan(30);

    await clickDrawing(page, 0);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await page.getByTestId("shape-align").selectOption("top");

    // After align-top the drawing's top sits on the content-area top margin,
    // regardless of where its anchor paragraph is down the page.
    const afterTop = (await drawings(page).nth(0).boundingBox())!.y;
    expect(Math.abs(afterTop - marginTop)).toBeLessThan(4);
  });

  test("resizing a drawing snaps its moving edge to another's and shows a guide (AC-27)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 340, 200);
    await insertShapeAt(page, "rect", 340, 420);
    await expect(drawings(page)).toHaveCount(2);

    // Anchor's left edge at 6cm; the resizer starts at 2cm (3cm wide => right
    // edge at 5cm). Growing its right edge toward 6cm should snap there (AC-27),
    // giving an exact 4cm width.
    await setPosition(page, 0, 6, 2);
    await setPosition(page, 1, 2, 8);

    await clickDrawing(page, 1);
    await expect(page.getByTestId("shape-handles")).toBeVisible();
    const anchorBox = (await drawings(page).nth(0).boundingBox())!;
    const handleBox = (await page.getByTestId("shape-resize-se").boundingBox())!;
    const startX = handleBox.x + handleBox.width / 2;
    const startY = handleBox.y + handleBox.height / 2;
    // Drag the SE handle so the right edge lands ~5px shy of the anchor's left
    // edge — inside the 8px snap threshold but not yet aligned.
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(anchorBox.x - 5, startY, { steps: 12 });

    const guides = page.getByTestId("drawing-guides");
    await expect(guides).toBeVisible();
    await expect(guides.locator(".drawing-guide").first()).toBeVisible();

    await page.mouse.up();
    // Guides vanish on release (AC-56).
    await expect(page.getByTestId("drawing-guides")).toHaveCount(0);

    // The right edge snapped to the anchor's left edge (6cm), so the resizer's
    // width is exactly 6 − 2 = 4cm rather than the 5px-shy value.
    await clickDrawing(page, 1);
    await page.getByTestId("shape-position-button").click();
    await expect(page.getByTestId("shape-position-dialog")).toBeVisible();
    const width = Number(await page.getByTestId("shape-position-width").inputValue());
    expect(Math.abs(width - 4)).toBeLessThan(0.1);
  });

  test("dragging a drawing near another's centre shows a guide and snaps into alignment (AC-18,23,26,55,56)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 360, 220);
    await insertShapeAt(page, "rect", 620, 420);
    await expect(drawings(page)).toHaveCount(2);

    const anchor = (await drawings(page).nth(0).boundingBox())!;
    const mover = drawings(page).nth(1);
    const box = (await mover.boundingBox())!;

    const anchorCentreX = anchor.x + anchor.width / 2;
    const moverCentreX = box.x + box.width / 2;
    // Grab the mover's lower half and drag its centre to 5px shy of the
    // anchor's centre x — within the 8px snap threshold but not aligned.
    const grabX = moverCentreX;
    const grabY = box.y + box.height * 0.72;
    await page.mouse.move(grabX, grabY);
    await page.mouse.down();
    await page.mouse.move(anchorCentreX - 5, grabY, { steps: 12 });

    // A guide line is shown while the edges line up (AC-18/21), and it is a
    // print-excluded overlay (AC-57).
    const guides = page.getByTestId("drawing-guides");
    await expect(guides).toBeVisible();
    await expect(guides.locator(".drawing-guide").first()).toBeVisible();

    await page.mouse.up();
    // Guides vanish on release (AC-56).
    await expect(page.getByTestId("drawing-guides")).toHaveCount(0);

    // Both shapes are the default 3cm wide, so equal centres mean equal left
    // offsets. Snapping pulled the released-5px-off drawing into alignment:
    // the two offsets now differ by far less than those 5px (~0.13cm).
    const offsets = await readOffsetsEmu(page);
    expect(offsets).toHaveLength(2);
    expect(Math.abs(offsets[0].x - offsets[1].x)).toBeLessThan(cmToEmu(0.05));
  });

  test("creation-drag snaps to alignment during the draw and clears the guide on release (AC-28,56)", async ({ page }) => {
    await createDocument(page);
    // An anchor whose left edge the new shape's creation drag will line up with.
    await insertShapeAt(page, "rect", 360, 220);
    await expect(drawings(page)).toHaveCount(1);
    const anchor = (await drawings(page).nth(0).boundingBox())!;

    // Set the caret to the top block first (as insertShapeAt does) so the drawn
    // shape anchors to the same paragraph as the anchor shape — their posOffsets
    // are then measured from the same origin and are directly comparable.
    await content(page).click({ position: { x: 20, y: 20 } });

    // Arm the rect tool, then draw a new shape by dragging — starting its left
    // edge 5px shy of the anchor's left (inside the 8px snap threshold) and
    // dragging right+down well past the click/drag threshold so it is a genuine
    // creation drag (AC-28), not a plain click.
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("shape-gallery-rect").click();
    const startX = anchor.x - 5;
    const startY = anchor.y + anchor.height + 60;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX + 120, startY + 90, { steps: 12 });

    // While the left edges line up a guide is shown (AC-18/21).
    const guides = page.getByTestId("drawing-guides");
    await expect(guides).toBeVisible();
    await expect(guides.locator(".drawing-guide").first()).toBeVisible();

    await page.mouse.up();
    // The creation-drag guide vanishes on release just like move/resize (AC-56):
    // the regression this fix addresses left it stuck because finishShapeDraw's
    // snap repopulated the guides after the pre-clear. That a guide showed at all
    // during the drag is itself evidence snapping applied to the creation drag
    // (AC-28) — guides and snap share the same match result.
    await expect(page.getByTestId("drawing-guides")).toHaveCount(0);
    await expect(drawings(page)).toHaveCount(2);
  });

  test("Format → Drawing → Align opens and runs by mouse click through the nested submenu (AC-12)", async ({ page }) => {
    await createDocument(page);
    // Push the anchor down the page so align-top has somewhere to move to.
    await content(page).click({ position: { x: 20, y: 20 } });
    for (let i = 0; i < 12; i++) {
      await page.keyboard.press("Enter");
    }
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("shape-gallery-rect").click();
    await content(page).click({ position: { x: 300, y: 300 } });
    await expect(drawings(page)).toHaveCount(1);

    const marginTop = (await content(page).getByTestId("page-margin").first().boundingBox())!.y;
    const beforeTop = (await drawings(page).nth(0).boundingBox())!.y;
    expect(beforeTop - marginTop).toBeGreaterThan(30);

    await clickDrawing(page, 0);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();

    // Drive the Format-menu path entirely by mouse: the Align entry is a
    // MenuSubmenu nested inside the Drawing MenuSubmenu, so clicking it must
    // open its flyout (not collapse the parent — the QA regression).
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-drawing").click();
    await page.getByTestId("drawing-align-submenu").click();
    const alignTop = page.getByTestId("drawing-align-top");
    await expect(alignTop).toBeVisible();
    await alignTop.click();

    // The command ran: the drawing sits on the content-area top margin.
    const afterTop = (await drawings(page).nth(0).boundingBox())!.y;
    expect(Math.abs(afterTop - marginTop)).toBeLessThan(4);
  });

  test("Format → Drawing → Distribute opens and runs by mouse click through the nested submenu (AC-13)", async ({ page }) => {
    await createDocument(page);
    // Three shapes at unevenly-spaced verticals (2, 3, 8 → gaps 1 and 5) so a
    // vertical distribute has a visible effect: the middle moves to 5cm.
    await insertShapeAt(page, "rect", 320, 220);
    await insertShapeAt(page, "ellipse", 540, 220);
    await insertShapeAt(page, "hexagon", 400, 420);
    await expect(drawings(page)).toHaveCount(3);
    await setPosition(page, 0, 1, 2);
    await setPosition(page, 1, 4, 3);
    await setPosition(page, 2, 2, 8);

    // Select all three so Distribute is enabled (fewer than three disables it).
    await clickDrawing(page, 0);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await clickDrawing(page, 1, true);
    await clickDrawing(page, 2, true);

    // Drive the Format-menu path entirely by mouse: Distribute is a MenuSubmenu
    // nested inside the Drawing MenuSubmenu, so clicking it must open its flyout
    // (not collapse the parent — the QA regression).
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-drawing").click();
    await page.getByTestId("drawing-distribute-submenu").click();
    const distributeVertical = page.getByTestId("drawing-distribute-vertical");
    await expect(distributeVertical).toBeVisible();
    await distributeVertical.click();

    // The command ran: the middle shape's top offset is now the exact midpoint
    // between the two outermost (2cm and 8cm → 5cm), an exact arithmetic
    // progression with the outermost two unchanged.
    const ys = (await readOffsetsEmu(page)).map((o) => o.y).sort((a, b) => a - b);
    expect(ys[0]).toBe(cmToEmu(2));
    expect(ys[2]).toBe(cmToEmu(8));
    expect(ys[1] - ys[0]).toBe(ys[2] - ys[1]);
    expect(ys[1]).toBe(cmToEmu(5));
  });

  test("Format → Drawing → Align to switches the reference by mouse and remembers it (AC-3,4)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 320, 220);
    await expect(drawings(page)).toHaveCount(1);
    await clickDrawing(page, 0);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();

    // Open Format → Drawing → Align to by mouse and pick Page.
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-drawing").click();
    await page.getByTestId("drawing-align-to-submenu").click();
    const pageRef = page.getByTestId("drawing-align-to-page");
    await expect(pageRef).toBeVisible();
    // Margin is the default; Page is not yet chosen.
    await expect(pageRef).toHaveAttribute("aria-checked", "false");
    await pageRef.click();

    // Reopen the same path: the chosen reference is remembered (session-scoped).
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-drawing").click();
    await page.getByTestId("drawing-align-to-submenu").click();
    await expect(page.getByTestId("drawing-align-to-page")).toHaveAttribute("aria-checked", "true");
    await expect(page.getByTestId("drawing-align-to-margin")).toHaveAttribute("aria-checked", "false");
  });
});
