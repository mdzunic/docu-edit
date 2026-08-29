import { describe, it, expect } from "vitest";
import {
  alignRects,
  distributeRects,
  type AlignRect,
} from "./drawingAlign";

// The gap-equalisation invariant asserted to a tight epsilon (SPEC-072). Cm
// floats can't be compared exactly after arithmetic, but the gaps must agree
// far below any visible or storable resolution.
const EPS = 1e-9;

describe("alignRects", () => {
  // Reference chosen so every computed centre/right result lands on an exact
  // binary fraction — each output is a single assigned coordinate, so === holds.
  const reference: AlignRect = { xCm: 2, yCm: 4, wCm: 10, hCm: 8 };
  const rects: AlignRect[] = [
    { xCm: 3, yCm: 5, wCm: 2, hCm: 1 },
    { xCm: 7, yCm: 9, wCm: 4, hCm: 3 },
  ];

  it("aligns left to the reference's left edge without moving y", () => {
    const out = alignRects(rects, "left", reference);
    expect(out[0].xCm).toBe(2);
    expect(out[1].xCm).toBe(2);
    expect(out[0].yCm).toBe(5);
    expect(out[1].yCm).toBe(9);
  });

  it("aligns horizontal centre to the reference's centre", () => {
    const out = alignRects(rects, "centerH", reference);
    // reference centre x = 2 + 10/2 = 7
    expect(out[0].xCm).toBe(6); // 7 - 2/2
    expect(out[1].xCm).toBe(5); // 7 - 4/2
    expect(out[0].yCm).toBe(5);
  });

  it("aligns right to the reference's right edge", () => {
    const out = alignRects(rects, "right", reference);
    // reference right = 2 + 10 = 12
    expect(out[0].xCm).toBe(10); // 12 - 2
    expect(out[1].xCm).toBe(8); // 12 - 4
  });

  it("aligns top to the reference's top edge without moving x", () => {
    const out = alignRects(rects, "top", reference);
    expect(out[0].yCm).toBe(4);
    expect(out[1].yCm).toBe(4);
    expect(out[0].xCm).toBe(3);
    expect(out[1].xCm).toBe(7);
  });

  it("aligns vertical middle to the reference's middle", () => {
    const out = alignRects(rects, "middleV", reference);
    // reference middle y = 4 + 8/2 = 8
    expect(out[0].yCm).toBe(7.5); // 8 - 1/2
    expect(out[1].yCm).toBe(6.5); // 8 - 3/2
  });

  it("aligns bottom to the reference's bottom edge", () => {
    const out = alignRects(rects, "bottom", reference);
    // reference bottom = 4 + 8 = 12
    expect(out[0].yCm).toBe(11); // 12 - 1
    expect(out[1].yCm).toBe(9); // 12 - 3
  });

  it("does not move the outermost drawing when the reference is the union", () => {
    // Union of the two rects: x 3..11, so left edge 3.
    const union: AlignRect = { xCm: 3, yCm: 5, wCm: 8, hCm: 7 };
    const out = alignRects(rects, "left", union);
    expect(out[0].xCm).toBe(3); // leftmost already at the edge — unchanged
  });

  it("is pure: does not mutate its inputs", () => {
    const input: AlignRect[] = [{ xCm: 3, yCm: 5, wCm: 2, hCm: 1 }];
    const snapshot = JSON.parse(JSON.stringify(input));
    alignRects(input, "right", reference);
    expect(input).toEqual(snapshot);
  });
});

describe("distributeRects", () => {
  it("returns fewer-than-three inputs unchanged", () => {
    const two: AlignRect[] = [
      { xCm: 1, yCm: 0, wCm: 2, hCm: 2 },
      { xCm: 9, yCm: 0, wCm: 2, hCm: 2 },
    ];
    expect(distributeRects(two, "h")).toEqual([
      { xCm: 1, yCm: 0 },
      { xCm: 9, yCm: 0 },
    ]);
    expect(distributeRects(two, "v")).toEqual([
      { xCm: 1, yCm: 0 },
      { xCm: 9, yCm: 0 },
    ]);
  });

  // Differently-sized rectangles so gap-equalisation and centre-equalisation
  // give different answers — the distinction the spec calls out.
  const varied: AlignRect[] = [
    { xCm: 0, yCm: 0, wCm: 1, hCm: 3 },
    { xCm: 4, yCm: 5, wCm: 3, hCm: 1 },
    { xCm: 10, yCm: 12, wCm: 2, hCm: 4 },
  ];

  it("equalises horizontal gaps, not centre distances", () => {
    const out = distributeRects(varied, "h");
    // Outermost two fixed.
    expect(out[0].xCm).toBe(0);
    expect(out[2].xCm).toBe(10);
    // Adjacent gaps (right edge -> next left edge) equal.
    const gap1 = out[1].xCm - (out[0].xCm + varied[0].wCm);
    const gap2 = out[2].xCm - (out[1].xCm + varied[1].wCm);
    expect(Math.abs(gap1 - gap2)).toBeLessThan(EPS);
    // Centre distances are NOT equal (proving it's gap- not centre-based).
    const c0 = out[0].xCm + varied[0].wCm / 2;
    const c1 = out[1].xCm + varied[1].wCm / 2;
    const c2 = out[2].xCm + varied[2].wCm / 2;
    expect(Math.abs((c1 - c0) - (c2 - c1))).toBeGreaterThan(EPS);
    // y untouched on the h axis.
    expect(out[1].yCm).toBe(5);
  });

  it("equalises vertical gaps, not centre distances", () => {
    const out = distributeRects(varied, "v");
    expect(out[0].yCm).toBe(0);
    expect(out[2].yCm).toBe(12);
    const gap1 = out[1].yCm - (out[0].yCm + varied[0].hCm);
    const gap2 = out[2].yCm - (out[1].yCm + varied[1].hCm);
    expect(Math.abs(gap1 - gap2)).toBeLessThan(EPS);
    const c0 = out[0].yCm + varied[0].hCm / 2;
    const c1 = out[1].yCm + varied[1].hCm / 2;
    const c2 = out[2].yCm + varied[2].hCm / 2;
    expect(Math.abs((c1 - c0) - (c2 - c1))).toBeGreaterThan(EPS);
    expect(out[1].xCm).toBe(4);
  });

  it("returns positions in input order regardless of spatial order", () => {
    // Input is not sorted by position; output keeps input order.
    const scrambled: AlignRect[] = [
      { xCm: 10, yCm: 0, wCm: 2, hCm: 2 },
      { xCm: 0, yCm: 0, wCm: 2, hCm: 2 },
      { xCm: 5, yCm: 0, wCm: 2, hCm: 2 },
    ];
    const out = distributeRects(scrambled, "h");
    // The outermost (leftmost x=0 and rightmost x=10) stay put.
    expect(out[0].xCm).toBe(10);
    expect(out[1].xCm).toBe(0);
    // Middle one repositioned; equal-size rects => equal gaps.
    const gapA = out[2].xCm - (out[1].xCm + 2);
    const gapB = out[0].xCm - (out[2].xCm + 2);
    expect(Math.abs(gapA - gapB)).toBeLessThan(EPS);
  });

  it("is pure: does not mutate its inputs", () => {
    const snapshot = JSON.parse(JSON.stringify(varied));
    distributeRects(varied, "h");
    expect(varied).toEqual(snapshot);
  });
});
