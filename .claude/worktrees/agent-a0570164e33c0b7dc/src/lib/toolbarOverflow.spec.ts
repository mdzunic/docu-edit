import { describe, expect, it } from "vitest";
import { COLLAPSE_ORDER, collapsedBands, fitsWholly, planCollapse } from "./toolbarOverflow";

// Representative band widths (px) for a full eight-band strip. Bands 1–4 are
// permanent; 5 (Colour), 6 (Paragraph), 7 (Lists), 8 (extras) collapse in the
// order 8→7→6→5.
const WIDTHS: Record<number, number> = {
  1: 72, // History
  2: 120, // Style
  3: 150, // Type
  4: 190, // Character
  5: 72, // Colour
  6: 150, // Paragraph
  7: 130, // Lists
  8: 72, // Paragraph extras
};
const ORDER = [1, 2, 3, 4, 5, 6, 7, 8];
const GAP = 12;
const OVERFLOW = 44;

function plan(available: number, previousCollapsed = 0) {
  return planCollapse({
    bandWidths: WIDTHS,
    bandOrder: ORDER,
    gap: GAP,
    available,
    overflowWidth: OVERFLOW,
    previousCollapsed,
    hysteresis: 24,
  });
}

describe("planCollapse", () => {
  it("collapses nothing when everything fits", () => {
    // Sum of all bands + 7 gaps = 956 + 84 = 1040.
    expect(plan(1100)).toBe(0);
  });

  it("collapses the highest bands first, one at a time", () => {
    // Just under the full width: band 8 collapses first.
    expect(plan(1000)).toBe(1);
    expect(collapsedBands(plan(1000))).toEqual(new Set([8]));
  });

  it("keeps the paragraph band (6) at a realistic 1024px (AC-36)", () => {
    // At ~1000px available (a 1024px viewport minus strip padding), bands 8
    // and 7 collapse but 6 (Paragraph) and 4 (Character) stay in the strip.
    const collapsed = collapsedBands(plan(1000));
    expect(collapsed.has(6)).toBe(false);
    expect(collapsed.has(4)).toBe(false);
  });

  it("never collapses a band below 5", () => {
    // Absurdly narrow: all four collapsible bands go, but no more.
    expect(plan(100)).toBe(COLLAPSE_ORDER.length);
    const collapsed = collapsedBands(plan(100));
    expect(collapsed.has(4)).toBe(false);
    expect(collapsed.has(1)).toBe(false);
  });

  it("collapses progressively as width shrinks", () => {
    const wide = plan(1040);
    const mid = plan(900);
    const narrow = plan(760);
    expect(mid).toBeGreaterThanOrEqual(wide);
    expect(narrow).toBeGreaterThanOrEqual(mid);
  });

  it("applies hysteresis so a band does not flap at the boundary", () => {
    // Find a width where 1 band is collapsed with no slack.
    const w = requiredForOneCollapsed();
    // Already collapsed: at exactly that width we should not re-expand even
    // though 0 collapsed would also nominally fit within a px or two.
    const held = plan(w + 1, 1);
    expect(held).toBe(1);
    // Given generous extra slack, it re-expands.
    expect(plan(w + 500, 1)).toBe(0);
  });
});

// Width at which exactly one band (8) is collapsed and it only just fits.
function requiredForOneCollapsed() {
  const visible = [1, 2, 3, 4, 5, 6, 7];
  const w =
    visible.reduce((s, b) => s + WIDTHS[b], 0) + GAP * (visible.length - 1) + OVERFLOW;
  return w;
}

// The Toolbar Redesign degrades in two steps: the strip first drops its
// control captions, and only when the icon-only geometry still does not fit
// does it collapse a band. `fitsWholly` is the first question — asked against
// the CAPTIONED band widths, which are wider than the icon-only ones.
describe("fitsWholly (the captions-before-bands step)", () => {
  // The same strip with captions under every icon: every band grows.
  const CAPTIONED: Record<number, number> = {
    1: 92, 2: 128, 3: 172, 4: 240, 5: 104, 6: 200, 7: 176, 8: 92,
  };
  const fits = (available: number, bandWidths = CAPTIONED) =>
    fitsWholly({ bandWidths, bandOrder: ORDER, gap: GAP, available, overflowWidth: OVERFLOW });

  it("is true when every captioned band fits with nothing collapsed", () => {
    // 1204 of bands + 7 gaps (84) = 1288.
    expect(fits(1288)).toBe(true);
    expect(fits(1400)).toBe(true);
  });

  it("is false one pixel below the captioned strip's own width", () => {
    expect(fits(1287)).toBe(false);
  });

  it("is answered independently of the icon-only geometry, which still fits", () => {
    // At 1100 the captions do not fit, so the strip drops them — and the
    // icon-only widths then need no band collapsed at all. Losing the words
    // under the icons costs no commands; collapsing a band would.
    expect(fits(1100)).toBe(false);
    expect(plan(1100)).toBe(0);
  });

  it("adds no overflow button when nothing is collapsed", () => {
    // A captioned strip that exactly fills the width fits: `fitsWholly` asks
    // about the zero-collapse case, where no overflow button is rendered.
    const single = { 1: 100 };
    expect(fitsWholly({ bandWidths: single, bandOrder: [1], gap: GAP, available: 100, overflowWidth: OVERFLOW })).toBe(true);
    expect(fitsWholly({ bandWidths: single, bandOrder: [1], gap: GAP, available: 99, overflowWidth: OVERFLOW })).toBe(false);
  });
});
