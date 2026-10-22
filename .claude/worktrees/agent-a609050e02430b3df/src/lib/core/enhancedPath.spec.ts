import { describe, it, expect } from "vitest";
import { parseEnhancedPath, enhancedPathString, enhancedPathViewBox } from "./enhancedPath";
import type { ShapePath } from "./shapePath";

describe("enhancedPathString / parseEnhancedPath round-trip", () => {
  const path: ShapePath = {
    subpaths: [
      {
        steps: [
          { op: "move", x: 0, y: 1 },
          { op: "line", x: 1, y: 1 },
          { op: "cubic", x1: 0.75, y1: 0.5, x2: 0.25, y2: 0.5, x: 0.5, y: 0 },
          { op: "close" },
        ],
        closed: true,
      },
    ],
  };

  it("emits M/L/C/Z terminated by N on the 21600 grid", () => {
    expect(enhancedPathString(path)).toBe("M 0 21600 L 21600 21600 C 16200 10800 5400 10800 10800 0 Z N");
    expect(enhancedPathViewBox()).toBe("0 0 21600 21600");
  });

  it("parses its own output back to an equal path", () => {
    const parsed = parseEnhancedPath(enhancedPathString(path), 21600, 21600);
    expect(parsed?.converted).toBe(false);
    expect(parsed?.path).toEqual(path);
  });
});

describe("parseEnhancedPath — primitive reduction", () => {
  it("converts a quadratic Q to a cubic exactly", () => {
    const parsed = parseEnhancedPath("M 0 0 Q 10800 21600 21600 0 N", 21600, 21600);
    expect(parsed?.converted).toBe(true);
    const steps = parsed?.path.subpaths[0].steps ?? [];
    expect(steps[1].op).toBe("cubic");
    if (steps[1].op === "cubic") {
      expect(steps[1].x1).toBeCloseTo(1 / 3);
      expect(steps[1].y1).toBeCloseTo(2 / 3);
    }
  });

  it("converts an arc A to cubics", () => {
    // Bounding box (0,0)-(21600,21600): unit circle centred at 0.5,0.5,
    // radius 0.5; start (1,0.5) end (0.5,0).
    const parsed = parseEnhancedPath("M 21600 10800 A 0 0 21600 21600 21600 10800 10800 0 N", 21600, 21600);
    expect(parsed?.converted).toBe(true);
    expect(parsed?.path.subpaths[0].steps.some((s) => s.op === "cubic")).toBe(true);
  });
});

describe("parseEnhancedPath — preservation cases", () => {
  it("returns null for a formula-driven path", () => {
    expect(parseEnhancedPath("M 0 0 L ?f0 21600 N", 21600, 21600)).toBeNull();
  });

  it("returns null for an unmodelled command", () => {
    expect(parseEnhancedPath("M 0 0 W 0 0 21600 21600 0 0 21600 0 N", 21600, 21600)).toBeNull();
  });

  it("returns null for an empty path", () => {
    expect(parseEnhancedPath("N", 21600, 21600)).toBeNull();
  });

  it("handles multiple implicit line points", () => {
    const parsed = parseEnhancedPath("M 0 0 L 21600 0 21600 21600 0 21600 Z N", 21600, 21600);
    expect(parsed?.path.subpaths[0].steps.filter((s) => s.op === "line")).toHaveLength(3);
    expect(parsed?.path.subpaths[0].closed).toBe(true);
  });
});
