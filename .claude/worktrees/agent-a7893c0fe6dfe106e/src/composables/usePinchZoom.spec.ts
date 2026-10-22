import { describe, it, expect } from "vitest";
import { pinchCentroid, pinchDistance, pinchZoomPercent } from "./usePinchZoom";

describe("pinch math (SPEC-068 AC-32/33)", () => {
  it("measures the distance between two points", () => {
    expect(pinchDistance({ clientX: 0, clientY: 0 }, { clientX: 3, clientY: 4 })).toBe(5);
  });

  it("finds the centroid of two points", () => {
    expect(pinchCentroid({ clientX: 0, clientY: 0 }, { clientX: 10, clientY: 20 })).toEqual({ x: 5, y: 10 });
  });

  it("scales the zoom percent by the ratio of finger distances", () => {
    // Fingers spread to double the distance → zoom doubles.
    expect(pinchZoomPercent(100, 100, 200)).toBe(200);
    // Fingers pinch to half the distance → zoom halves.
    expect(pinchZoomPercent(150, 200, 100)).toBe(75);
  });

  it("is a no-op when the initial distance is degenerate", () => {
    expect(pinchZoomPercent(120, 0, 300)).toBe(120);
  });
});
