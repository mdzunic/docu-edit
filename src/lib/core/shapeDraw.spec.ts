import { describe, it, expect } from "vitest";
import {
  polylinePath,
  catmullRomPath,
  presetToPath,
  drawnShapeBox,
  SCRIBBLE_TOLERANCE_CM,
  CLOSE_DISTANCE_CM,
} from "./shapeDraw";
import { pathToSvgD } from "./shapePath";

describe("polylinePath (Freeform)", () => {
  it("builds move + lines, corner vertices, open by default", () => {
    const p = polylinePath([{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }], false);
    expect(p.subpaths[0].steps.map((s) => s.op)).toEqual(["move", "line", "line"]);
    expect(p.subpaths[0].closed).toBe(false);
    expect(p.smooth).toBeUndefined();
  });

  it("closes when asked", () => {
    const p = polylinePath([{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }], true);
    expect(p.subpaths[0].closed).toBe(true);
    expect(p.subpaths[0].steps.at(-1)!.op).toBe("close");
  });
});

describe("catmullRomPath (Curve / scribble smoothing)", () => {
  it("passes smoothly through every placed point and marks all smooth", () => {
    const pts = [{ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 2, y: 0 }];
    const p = catmullRomPath(pts, false);
    const ends = p.subpaths[0].steps.filter((s) => s.op === "cubic").map((s) => (s.op === "cubic" ? { x: s.x, y: s.y } : null));
    expect(ends).toEqual([{ x: 1, y: 1 }, { x: 2, y: 0 }]);
    expect(p.smooth).toEqual([true, true, true]);
  });

  it("adds a closing cubic when closed", () => {
    const p = catmullRomPath([{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 1, y: 1 }, { x: 0, y: 1 }], true);
    expect(p.subpaths[0].closed).toBe(true);
    expect(p.subpaths[0].steps.filter((s) => s.op === "cubic")).toHaveLength(4);
  });
});

describe("presetToPath (preset -> freeform, AC-15)", () => {
  it("converts a triangle preset to a closed line path", () => {
    const p = presetToPath("triangle");
    expect(p).not.toBeNull();
    expect(p!.subpaths[0].closed).toBe(true);
    // Apex at top-centre survives.
    expect(pathToSvgD(p!, 100, 100)).toContain("M 50 0");
  });

  it("returns null for a line geometry", () => {
    expect(presetToPath("line")).toBeNull();
  });
});

describe("drawnShapeBox", () => {
  it("re-fits an overlay-space path to a 0..1 box and reports its extent", () => {
    // A path drawn from (2,3) spanning 4cm x 2cm.
    const raw = polylinePath([{ x: 2, y: 3 }, { x: 6, y: 3 }, { x: 6, y: 5 }], true);
    const { path, widthCm, heightCm, offset } = drawnShapeBox(raw);
    expect(widthCm).toBeCloseTo(4);
    expect(heightCm).toBeCloseTo(2);
    expect(offset).toEqual({ x: 2, y: 3 });
    // First vertex now at the box origin.
    expect(path.subpaths[0].steps[0]).toMatchObject({ op: "move", x: 0, y: 0 });
  });
});

describe("constants", () => {
  it("states the scribble tolerance and close distance", () => {
    expect(SCRIBBLE_TOLERANCE_CM).toBe(0.03);
    expect(CLOSE_DISTANCE_CM).toBe(0.2);
  });
});
