import { describe, expect, it } from "vitest";
import { applyGuides, clientToCm, DRAG_ENGAGE_CM, resizeRect, resizeRotatedRect } from "./stageMath";

const surface = { left: 100, top: 50, width: 960, height: 540, right: 1060, bottom: 590 } as DOMRect;

describe("clientToCm", () => {
  it("converts a viewport point to slide cm via the surface rect", () => {
    // 960px wide surface for a 24cm slide → 40px/cm; (100+400)-100 = 400px = 10cm.
    const p = clientToCm(500, 50, surface, 24, 13.5);
    expect(p.xCm).toBeCloseTo(10, 3);
    expect(p.yCm).toBeCloseTo(0, 3);
  });
});

describe("resizeRect", () => {
  const orig = { xCm: 4, yCm: 4, widthCm: 6, heightCm: 3 };

  it("east handle moves the right edge, keeps the left anchored", () => {
    const r = resizeRect("e", orig, { xCm: 12, yCm: 0 }, false);
    expect(r.xCm).toBe(4);
    expect(r.widthCm).toBe(8);
  });

  it("north-west handle moves the top-left, keeps the bottom-right anchored", () => {
    const r = resizeRect("nw", orig, { xCm: 2, yCm: 2 }, false);
    expect(r.xCm).toBe(2);
    expect(r.yCm).toBe(2);
    expect(r.widthCm).toBe(8);
    expect(r.heightCm).toBe(5);
  });

  it("enforces the 0.5cm minimum", () => {
    const r = resizeRect("e", orig, { xCm: 4, yCm: 0 }, false);
    expect(r.widthCm).toBe(0.5);
  });

  it("aspect keeps the original ratio on a corner", () => {
    const r = resizeRect("se", orig, { xCm: 16, yCm: 5 }, true);
    expect(r.widthCm / r.heightCm).toBeCloseTo(orig.widthCm / orig.heightCm, 3);
  });
});

describe("applyGuides", () => {
  it("snaps the centre to the slide centreline within threshold", () => {
    const rect = { xCm: 9.9, yCm: 2, widthCm: 4, heightCm: 2 };
    const { rect: snapped, guides } = applyGuides(rect, 24, 13.5, 0.15);
    // centre would be 11.9; slide centre 12 → within 0.15 → snap centre to 12.
    expect(snapped.xCm + snapped.widthCm / 2).toBeCloseTo(12, 3);
    expect(guides.vertical).toContain(12);
  });

  it("snaps the left edge to zero", () => {
    const rect = { xCm: 0.1, yCm: 5, widthCm: 4, heightCm: 2 };
    const { rect: snapped, guides } = applyGuides(rect, 24, 13.5, 0.15);
    expect(snapped.xCm).toBe(0);
    expect(guides.vertical).toContain(0);
  });

  it("leaves a far-from-guide rect untouched", () => {
    const rect = { xCm: 3, yCm: 3, widthCm: 4, heightCm: 2 };
    const { rect: snapped, guides } = applyGuides(rect, 24, 13.5, 0.15);
    expect(snapped.xCm).toBe(3);
    expect(guides.vertical).toHaveLength(0);
    expect(guides.horizontal).toHaveLength(0);
  });

  // SPEC-096: the optional fifth argument adds other frames' edges/centres.
  it("snaps to another frame's edge through `others`, kind 'frame'", () => {
    const rect = { xCm: 3, yCm: 3, widthCm: 4, heightCm: 2 };
    const others = [{ xCm: 3.1, yCm: 8, widthCm: 5, heightCm: 2 }];
    const { rect: snapped, guides } = applyGuides(rect, 24, 13.5, 0.15, others);
    expect(snapped.xCm).toBeCloseTo(3.1, 6);
    expect(guides.vertical).toContain(3.1);
    expect(guides.verticalKinds).toContain("frame");
  });

  it("marks a slide-line snap with kind 'slide'", () => {
    const rect = { xCm: 0.1, yCm: 5, widthCm: 4, heightCm: 2 };
    const { guides } = applyGuides(rect, 24, 13.5, 0.15, []);
    expect(guides.vertical).toContain(0);
    expect(guides.verticalKinds).toContain("slide");
  });

  it("prefers the slide centre over a nearer frame line", () => {
    const rect = { xCm: 9.9, yCm: 5, widthCm: 4, heightCm: 2 }; // centre 11.9
    const others = [{ xCm: 0, yCm: 0, widthCm: 9.95, heightCm: 1 }]; // edge 0.05 away
    const { rect: snapped, guides } = applyGuides(rect, 24, 13.5, 0.15, others);
    expect(snapped.xCm + snapped.widthCm / 2).toBeCloseTo(12, 6);
    expect(guides.vertical).toContain(12);
  });
});

describe("DRAG_ENGAGE_CM", () => {
  it("keeps the 3px-at-96dpi value the drag closure used", () => {
    expect(DRAG_ENGAGE_CM).toBeCloseTo((3 * 2.54) / 96, 10);
  });
});

describe("resizeRotatedRect", () => {
  const orig = { xCm: 4, yCm: 4, widthCm: 6, heightCm: 3 };

  it("is bit-identical to resizeRect at 0°", () => {
    const pointer = { xCm: 12, yCm: 5 };
    expect(resizeRotatedRect("se", orig, 0, pointer, false)).toEqual(
      resizeRect("se", orig, pointer, false),
    );
  });

  it("keeps the opposite corner fixed in slide space at 90°", () => {
    // At 90° the nw corner (4,4) paints at the rotated position; drag the se
    // handle and the nw corner's slide-space position must not move.
    const cx = orig.xCm + orig.widthCm / 2;
    const cy = orig.yCm + orig.heightCm / 2;
    // nw corner rotated 90° about the centre: (cx + (4−cy)·(−1)…) — compute
    // directly: R(90°)·(dx,dy) = (−dy, dx).
    const dx = orig.xCm - cx;
    const dy = orig.yCm - cy;
    const anchorBefore = { x: cx - dy, y: cy + dx };
    const out = resizeRotatedRect("se", orig, 90, { xCm: 8, yCm: 9 }, false);
    const ncx = out.xCm + out.widthCm / 2;
    const ncy = out.yCm + out.heightCm / 2;
    const ndx = out.xCm - ncx;
    const ndy = out.yCm - ncy;
    const anchorAfter = { x: ncx - ndy, y: ncy + ndx };
    expect(anchorAfter.x).toBeCloseTo(anchorBefore.x, 6);
    expect(anchorAfter.y).toBeCloseTo(anchorBefore.y, 6);
  });

  it("sizes along the rotated axes: at 90° a vertical drag grows the width", () => {
    // The e handle of a 90°-rotated frame points screen-south; dragging the
    // pointer down (greater y) extends the frame's own width.
    const out = resizeRotatedRect("e", orig, 90, { xCm: 7, yCm: 10 }, false);
    expect(out.widthCm).toBeGreaterThan(orig.widthCm);
    expect(out.heightCm).toBeCloseTo(orig.heightCm, 6);
  });
});
