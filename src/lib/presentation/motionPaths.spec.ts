// SPEC-105 motion-path geometry: preset resolution (slide fractions from the
// frame's centre), the path grammar the readers accept, recognising a read
// path as its preset again, and the endpoint drag's snapping.
import { describe, expect, it } from "vitest";
import {
  formatPathD,
  inferMotionPath,
  motionPathEnd,
  motionPathFromHandle,
  motionPathHandle,
  parseMotionPathD,
  pathDInPixels,
  resolveMotionPath,
  samplePath,
} from "./motionPaths";
import type { MotionPath } from "./model";

const d = (path: MotionPath) => formatPathD(resolveMotionPath(path) ?? []);

describe("resolveMotionPath", () => {
  it("line: a straight direction × distance move (absent = right, 25%)", () => {
    expect(d({ preset: "line" })).toBe("M 0 0 L 0.25 0");
    expect(d({ preset: "line", direction: "up", distanceFraction: 0.5 })).toBe("M 0 0 L 0 -0.5");
  });

  it("arc: the same end, bowing perpendicular by half the distance", () => {
    expect(d({ preset: "arc", distanceFraction: 0.3 })).toBe("M 0 0 C 0 -0.2 0.3 -0.2 0.3 0");
    // The bow's midpoint sits half the distance off the chord.
    const mid = samplePath(resolveMotionPath({ preset: "arc", distanceFraction: 0.3 })!, 2)[1];
    expect(mid.x).toBeCloseTo(0.15);
    expect(mid.y).toBeCloseTo(-0.15);
  });

  it("turn: an L — along the direction, then a quarter clockwise", () => {
    expect(d({ preset: "turn" })).toBe("M 0 0 L 0.25 0 L 0.25 0.25");
    expect(d({ preset: "turn", direction: "left", distanceFraction: 0.1 })).toBe("M 0 0 L -0.1 0 L -0.1 -0.1");
  });

  it("loop: a closed circle of radius distance / 2", () => {
    const loop = resolveMotionPath({ preset: "loop", distanceFraction: 0.4 })!;
    expect(loop.at(-1)).toEqual({ op: "Z", args: [] });
    const points = samplePath(loop, 16);
    expect(Math.max(...points.map((p) => p.x))).toBeCloseTo(0.4);
    expect(Math.min(...points.map((p) => p.y))).toBeCloseTo(-0.2, 2);
    expect(motionPathEnd({ preset: "loop" })).toEqual({ x: 0, y: 0 });
  });

  it("custom: its own path, when in subset", () => {
    expect(d({ preset: "custom", d: "M 0 0 L 0.1 0.2" })).toBe("M 0 0 L 0.1 0.2");
    expect(resolveMotionPath({ preset: "custom", d: "M 0 0 Q 1 1 2 2" })).toBeUndefined();
  });
});

describe("parseMotionPathD", () => {
  it("accepts M/L/C/Z with numbers, relative forms made absolute, a trailing E tolerated", () => {
    expect(formatPathD(parseMotionPathD("m 0 0 l 0.25 0 l 0 0.1")!)).toBe("M 0 0 L 0.25 0 L 0.25 0.1");
    expect(formatPathD(parseMotionPathD("M0,0L0.25,0E")!)).toBe("M 0 0 L 0.25 0");
    expect(formatPathD(parseMotionPathD("M 0 0 L 1 0 2 0")!)).toBe("M 0 0 L 1 0 L 2 0");
    expect(formatPathD(parseMotionPathD("M -4.16667E-6 0 L 0.25 0")!)).toBe("M 0 0 L 0.25 0");
  });

  it("refuses anything else", () => {
    for (const bad of ["", "L 0 0", "M 0 0 Q 1 1 2 2", "M 0 0 L 1", "M 0 0 A 1 1 0 0 1 2 2", "M x y", "M 0 0 E L 1 1"]) {
      expect(parseMotionPathD(bad), bad).toBeUndefined();
    }
  });
});

describe("inferMotionPath", () => {
  const presets: MotionPath[] = [
    { preset: "line" },
    { preset: "line", direction: "down", distanceFraction: 1 },
    { preset: "arc", direction: "left", distanceFraction: 0.05 },
    { preset: "turn", direction: "up", distanceFraction: 0.33 },
    { preset: "loop", distanceFraction: 0.6 },
  ];
  for (const path of presets) {
    it(`recognises ${JSON.stringify(path)} from its own path`, () => {
      expect(inferMotionPath(path.preset, parseMotionPathD(d(path))!)).toEqual(path);
    });
  }

  it("refuses a path its preset does not reproduce, or a distance out of bounds", () => {
    expect(inferMotionPath("line", parseMotionPathD("M 0 0 L 0.2 0.2")!)).toBeUndefined();
    expect(inferMotionPath("arc", parseMotionPathD("M 0 0 L 0.2 0")!)).toBeUndefined();
    expect(inferMotionPath("line", parseMotionPathD("M 0 0 L 0.01 0")!)).toBeUndefined();
    expect(inferMotionPath("line", parseMotionPathD("M 0 0 L 2 0")!)).toBeUndefined();
  });
});

describe("the endpoint drag", () => {
  it("sits on the path end (a loop: its far side)", () => {
    expect(motionPathHandle({ preset: "line", direction: "down" })).toEqual({ x: 0, y: 0.25 });
    expect(motionPathHandle({ preset: "turn" })).toEqual({ x: 0.25, y: 0.25 });
    expect(motionPathHandle({ preset: "loop", distanceFraction: 0.4 })).toEqual({ x: 0.4, y: 0 });
  });

  it("snaps a dragged handle to a direction and a whole-percent distance in 5–100%", () => {
    expect(motionPathFromHandle({ preset: "line" }, { x: 0.02, y: -0.403 })).toEqual({ preset: "line", direction: "up", distanceFraction: 0.4 });
    expect(motionPathFromHandle({ preset: "arc" }, { x: 0.251, y: 0.01 })).toEqual({ preset: "arc" });
    expect(motionPathFromHandle({ preset: "line" }, { x: -3, y: 0 })).toEqual({ preset: "line", direction: "left", distanceFraction: 1 });
    expect(motionPathFromHandle({ preset: "line" }, { x: 0.001, y: 0 })).toEqual({ preset: "line", distanceFraction: 0.05 });
    // A turn's end lies in one quadrant per direction.
    expect(motionPathFromHandle({ preset: "turn" }, { x: -0.3, y: 0.3 })).toEqual({ preset: "turn", direction: "down", distanceFraction: 0.3 });
    // A loop keeps its side and takes the reach.
    expect(motionPathFromHandle({ preset: "loop" }, { x: 0, y: -0.5 })).toEqual({ preset: "loop", distanceFraction: 0.5 });
  });

  it("draws in surface pixels from the frame centre", () => {
    expect(pathDInPixels(resolveMotionPath({ preset: "line" })!, { x: 100, y: 50 }, 1000, 500)).toBe("M 100.00 50.00 L 350.00 50.00");
  });
});
