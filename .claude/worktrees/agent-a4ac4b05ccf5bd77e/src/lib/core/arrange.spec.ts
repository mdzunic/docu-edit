// SPEC-096: the pure cm-rect algebra behind Align / Distribute / smart
// guides — six edges against all three references, distribute with 2/3/many
// and ragged rects, snap priority and threshold boundaries.
import { describe, expect, it } from "vitest";
import {
  alignRects,
  boundingRect,
  distributeRects,
  guideCandidates,
  snapRect,
  type Rect,
} from "./arrange";

const container: Rect = { xCm: 0, yCm: 0, widthCm: 24, heightCm: 13.5 };

const a: Rect = { xCm: 2, yCm: 1, widthCm: 4, heightCm: 2 };
const b: Rect = { xCm: 8, yCm: 4, widthCm: 2, heightCm: 3 };
const c: Rect = { xCm: 14, yCm: 9, widthCm: 6, heightCm: 1 };

describe("boundingRect", () => {
  it("is null for an empty list", () => {
    expect(boundingRect([])).toBeNull();
  });

  it("is the union box", () => {
    expect(boundingRect([a, b, c])).toEqual({ xCm: 2, yCm: 1, widthCm: 18, heightCm: 9 });
  });
});

describe("alignRects", () => {
  it("aligns left to the selection's bounding box", () => {
    const out = alignRects([a, b, c], "left", "selection", container);
    expect(out.map((r) => r.xCm)).toEqual([2, 2, 2]);
    // Only the aligned axis moves.
    expect(out.map((r) => r.yCm)).toEqual([1, 4, 9]);
  });

  it("aligns right to the selection's bounding box", () => {
    const out = alignRects([a, b, c], "right", "selection", container);
    expect(out.map((r) => r.xCm + r.widthCm)).toEqual([20, 20, 20]);
  });

  it("centres horizontally on the selection box", () => {
    const out = alignRects([a, b, c], "centerX", "selection", container);
    for (const r of out) {
      expect(r.xCm + r.widthCm / 2).toBeCloseTo(11, 6);
    }
  });

  it("aligns top/middle/bottom on the selection box", () => {
    expect(alignRects([a, b, c], "top", "selection", container).map((r) => r.yCm)).toEqual([1, 1, 1]);
    for (const r of alignRects([a, b, c], "middleY", "selection", container)) {
      expect(r.yCm + r.heightCm / 2).toBeCloseTo(5.5, 6);
    }
    expect(
      alignRects([a, b, c], "bottom", "selection", container).map((r) => r.yCm + r.heightCm),
    ).toEqual([10, 10, 10]);
  });

  it("aligns against the first rect", () => {
    const out = alignRects([b, a, c], "left", "first", container);
    expect(out.map((r) => r.xCm)).toEqual([8, 8, 8]);
  });

  it("aligns against the container", () => {
    const left = alignRects([a, b], "left", "container", container);
    expect(left.map((r) => r.xCm)).toEqual([0, 0]);
    const right = alignRects([a, b], "right", "container", container);
    expect(right.map((r) => r.xCm + r.widthCm)).toEqual([24, 24]);
    const mid = alignRects([a], "middleY", "container", container);
    expect(mid[0].yCm + mid[0].heightCm / 2).toBeCloseTo(6.75, 6);
  });

  it("returns rects in the input order", () => {
    const out = alignRects([c, a, b], "top", "selection", container);
    expect(out.map((r) => r.widthCm)).toEqual([6, 4, 2]);
  });
});

describe("distributeRects", () => {
  it("returns fewer than three rects unchanged", () => {
    expect(distributeRects([a, b], "x")).toEqual([a, b]);
  });

  it("holds the outermost two and equalises gaps for three", () => {
    const out = distributeRects([a, b, c], "x");
    expect(out[0]).toEqual(a); // leftmost held
    expect(out[2]).toEqual(c); // rightmost held
    const gap1 = out[1].xCm - (out[0].xCm + out[0].widthCm);
    const gap2 = out[2].xCm - (out[1].xCm + out[1].widthCm);
    expect(gap1).toBeCloseTo(gap2, 6);
    expect(gap1).toBeCloseTo(3, 6); // (20−2) − (4+2+6) = 6 over 2 gaps
  });

  it("equalises ragged sizes by gap, not centre", () => {
    const wide: Rect = { xCm: 0, yCm: 0, widthCm: 8, heightCm: 1 };
    const thin: Rect = { xCm: 9, yCm: 0, widthCm: 1, heightCm: 1 };
    const end: Rect = { xCm: 19, yCm: 0, widthCm: 1, heightCm: 1 };
    const out = distributeRects([wide, thin, end], "x");
    const gap1 = out[1].xCm - 8;
    const gap2 = out[2].xCm - (out[1].xCm + 1);
    expect(gap1).toBeCloseTo(gap2, 6);
  });

  it("distributes many rects vertically, preserving input order", () => {
    const rects: Rect[] = [
      { xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 },
      { xCm: 0, yCm: 9, widthCm: 1, heightCm: 1 },
      { xCm: 0, yCm: 2, widthCm: 1, heightCm: 2 },
      { xCm: 0, yCm: 6, widthCm: 1, heightCm: 1 },
    ];
    const out = distributeRects(rects, "y");
    expect(out[0].yCm).toBe(0);
    expect(out[1].yCm).toBe(9);
    // Sorted by centre: [0], [2], [3], [1] — gaps all equal.
    const gapA = out[2].yCm - 1;
    const gapB = out[3].yCm - (out[2].yCm + 2);
    const gapC = out[1].yCm - (out[3].yCm + 1);
    expect(gapA).toBeCloseTo(gapB, 6);
    expect(gapB).toBeCloseTo(gapC, 6);
  });
});

describe("guideCandidates / snapRect", () => {
  it("emits the container tier first", () => {
    const candidates = guideCandidates([a], container);
    expect(candidates[0]).toEqual({ axis: "x", line: 12, kind: "centre" });
    expect(candidates[1]).toEqual({ axis: "y", line: 6.75, kind: "centre" });
    // Frame lines follow: a's left/right/centreX then top/bottom/middleY.
    expect(candidates.slice(6).map((cnd) => cnd.line)).toEqual([2, 6, 4, 1, 3, 2]);
  });

  it("snaps to the container centre before a nearer frame edge", () => {
    // Rect centre at 11.9 (0.1 from slide centre); frame edge at 9.95 is only
    // 0.05 from the rect's left edge — the slide centre still wins.
    const other: Rect = { xCm: 0, yCm: 0, widthCm: 9.95, heightCm: 1 };
    const rect: Rect = { xCm: 9.9, yCm: 5, widthCm: 4, heightCm: 2 };
    const { rect: snapped, fired } = snapRect(rect, guideCandidates([other], container), 0.15);
    expect(snapped.xCm + snapped.widthCm / 2).toBeCloseTo(12, 6);
    expect(fired.find((f) => f.axis === "x")?.line).toBe(12);
  });

  it("snaps to the nearest frame line when the container is out of range", () => {
    const near: Rect = { xCm: 5.08, yCm: 0, widthCm: 1, heightCm: 1 };
    const far: Rect = { xCm: 5.2, yCm: 0, widthCm: 1, heightCm: 1 };
    const rect: Rect = { xCm: 5, yCm: 5, widthCm: 2, heightCm: 2 };
    const { rect: snapped, fired } = snapRect(rect, guideCandidates([far, near], container), 0.15);
    expect(snapped.xCm).toBeCloseTo(5.08, 6);
    expect(fired.find((f) => f.axis === "x")?.line).toBeCloseTo(5.08, 6);
  });

  it("snaps an edge to another frame's edge and a centre to its centre", () => {
    const other: Rect = { xCm: 10, yCm: 2, widthCm: 4, heightCm: 4 };
    const rect: Rect = { xCm: 10.1, yCm: 8, widthCm: 2, heightCm: 2 };
    const { rect: snapped } = snapRect(rect, guideCandidates([other], container), 0.15);
    expect(snapped.xCm).toBeCloseTo(10, 6);
  });

  it("fires at most one candidate per axis", () => {
    const rect: Rect = { xCm: 0.05, yCm: 0.05, widthCm: 4, heightCm: 2 };
    const { fired } = snapRect(rect, guideCandidates([], container), 0.15);
    expect(fired.filter((f) => f.axis === "x")).toHaveLength(1);
    expect(fired.filter((f) => f.axis === "y")).toHaveLength(1);
  });

  it("respects the threshold boundary exactly", () => {
    const at: Rect = { xCm: 0.15, yCm: 5, widthCm: 2, heightCm: 2 };
    expect(snapRect(at, guideCandidates([], container), 0.15).fired.length).toBeGreaterThan(0);
    const past: Rect = { xCm: 0.1501, yCm: 5, widthCm: 2, heightCm: 2 };
    expect(
      snapRect(past, guideCandidates([], container), 0.15).fired.filter((f) => f.axis === "x"),
    ).toHaveLength(0);
  });

  it("does not snap a centre candidate to an edge", () => {
    // A frame centre at 5; the dragged rect's left edge at 5.05 — centres
    // attract centres only, so nothing fires.
    const other: Rect = { xCm: 4, yCm: 0, widthCm: 2, heightCm: 1 };
    const rect: Rect = { xCm: 5.05, yCm: 6, widthCm: 3, heightCm: 1 };
    const { fired } = snapRect(rect, guideCandidates([other], container), 0.02);
    expect(fired).toHaveLength(0);
  });
});
