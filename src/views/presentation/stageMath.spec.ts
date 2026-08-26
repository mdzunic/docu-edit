import { describe, expect, it } from "vitest";
import { applyGuides, clientToCm, resizeRect } from "./stageMath";

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
});
