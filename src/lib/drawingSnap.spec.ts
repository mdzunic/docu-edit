import { describe, it, expect } from "vitest";
import {
  computeDrawingSnap,
  gridCandidates,
  type SnapCandidate,
} from "./drawingSnap";
import { type AlignRect } from "./drawingAlign";

const rect: AlignRect = { xCm: 5, yCm: 5, wCm: 2, hCm: 2 };

describe("computeDrawingSnap", () => {
  it("snaps the left edge onto a nearby vertical candidate", () => {
    // A vertical line at 4.9cm, 0.1cm from the left edge (5cm), threshold 0.2.
    const candidates: SnapCandidate[] = [{ axis: "x", coordCm: 4.9, startCm: 0, endCm: 10 }];
    const result = computeDrawingSnap(rect, candidates, 0.2);
    expect(result.xCm).toBeCloseTo(4.9, 9);
    expect(result.yCm).toBe(5); // unchanged
    expect(result.guides).toHaveLength(1);
    expect(result.guides[0]).toMatchObject({ axis: "x", coordCm: 4.9 });
  });

  it("does not snap when the nearest candidate is beyond the threshold", () => {
    const candidates: SnapCandidate[] = [{ axis: "x", coordCm: 4.5, startCm: 0, endCm: 10 }];
    const result = computeDrawingSnap(rect, candidates, 0.2);
    expect(result.xCm).toBe(5);
    expect(result.guides).toHaveLength(0);
  });

  it("snaps the closest of several features (centre beats edge)", () => {
    // Candidate at 6.05cm: centre is at 6cm (0.05 away), right at 7cm (0.95).
    const candidates: SnapCandidate[] = [{ axis: "x", coordCm: 6.05, startCm: 0, endCm: 10 }];
    const result = computeDrawingSnap(rect, candidates, 0.2);
    // Centre snaps: new centre = 6.05 => new x = 6.05 - 1 = 5.05.
    expect(result.xCm).toBeCloseTo(5.05, 9);
  });

  it("snaps both axes at once and draws a guide for each (AC-22)", () => {
    const candidates: SnapCandidate[] = [
      { axis: "x", coordCm: 4.95, startCm: 0, endCm: 10 },
      { axis: "y", coordCm: 5.05, startCm: 0, endCm: 10 },
    ];
    const result = computeDrawingSnap(rect, candidates, 0.2);
    expect(result.xCm).toBeCloseTo(4.95, 9);
    expect(result.yCm).toBeCloseTo(5.05, 9);
    expect(result.guides).toHaveLength(2);
  });

  it("draws a guide for every target sharing the matched coordinate (AC-22)", () => {
    // Two drawings whose right edges both sit at 4.95cm.
    const candidates: SnapCandidate[] = [
      { axis: "x", coordCm: 4.95, startCm: 0, endCm: 3 },
      { axis: "x", coordCm: 4.95, startCm: 8, endCm: 9 },
    ];
    const result = computeDrawingSnap(rect, candidates, 0.2);
    expect(result.guides).toHaveLength(2);
    // The guide spans from the dragged drawing to the matched target.
    expect(result.guides[0].coordCm).toBe(4.95);
  });

  it("is pure: does not mutate its inputs", () => {
    const candidates: SnapCandidate[] = [{ axis: "x", coordCm: 4.9, startCm: 0, endCm: 10 }];
    const snapshot = JSON.parse(JSON.stringify({ rect, candidates }));
    computeDrawingSnap(rect, candidates, 0.2);
    expect({ rect, candidates }).toEqual(snapshot);
  });
});

describe("gridCandidates", () => {
  it("covers [0, extent] at the given spacing", () => {
    const lines = gridCandidates("x", 1, 3, 0, 10);
    expect(lines.map((c) => c.coordCm)).toEqual([0, 1, 2, 3]);
    expect(lines.every((c) => c.axis === "x" && c.startCm === 0 && c.endCm === 10)).toBe(true);
  });

  it("returns nothing for a non-positive spacing (no divide-by-zero)", () => {
    expect(gridCandidates("y", 0, 10, 0, 5)).toEqual([]);
  });

  it("lets a dragged rect snap onto the nearest grid line", () => {
    const lines = gridCandidates("x", 1, 20, 0, 20);
    // Left edge at 5.05 snaps to the 5cm grid line within a 0.2 threshold.
    const result = computeDrawingSnap({ xCm: 5.05, yCm: 0, wCm: 2, hCm: 2 }, lines, 0.2);
    expect(result.xCm).toBeCloseTo(5, 9);
  });
});
