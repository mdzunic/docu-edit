import { describe, it, expect } from "vitest";
import {
  type ShapePath,
  pathBounds,
  normalisePath,
  pointInPath,
  distanceToPath,
  pathToSvgD,
  simplifyPolyline,
  snapPathToGrid,
  mapPath,
  quadraticToCubic,
  arcToCubics,
  PATH_GRID,
} from "./shapePath";

/** A unit square, closed, as four line steps. */
function square(): ShapePath {
  return {
    subpaths: [
      {
        steps: [
          { op: "move", x: 0, y: 0 },
          { op: "line", x: 1, y: 0 },
          { op: "line", x: 1, y: 1 },
          { op: "line", x: 0, y: 1 },
          { op: "close" },
        ],
        closed: true,
      },
    ],
  };
}

describe("pathBounds", () => {
  it("covers anchors and control points", () => {
    const path: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0.2, y: 0.2 },
            { op: "cubic", x1: -0.1, y1: 0.5, x2: 1.2, y2: 0.5, x: 0.8, y: 0.8 },
          ],
          closed: false,
        },
      ],
    };
    const b = pathBounds(path);
    expect(b.x).toBeCloseTo(-0.1);
    expect(b.y).toBeCloseTo(0.2);
    expect(b.w).toBeCloseTo(1.3);
    expect(b.h).toBeCloseTo(0.6);
  });

  it("returns a zero box for an empty path", () => {
    expect(pathBounds({ subpaths: [] })).toEqual({ x: 0, y: 0, w: 0, h: 0 });
  });
});

describe("normalisePath", () => {
  it("re-fits an oversized path to the unit square and reports offset/scale", () => {
    const path: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0, y: 0 },
            { op: "line", x: 1.5, y: 0 },
            { op: "line", x: 1.5, y: 1 },
            { op: "close" },
          ],
          closed: true,
        },
      ],
    };
    const { path: fitted, offset, scale } = normalisePath(path);
    const b = pathBounds(fitted);
    expect(b).toMatchObject({ x: 0, y: 0, w: 1, h: 1 });
    expect(offset).toEqual({ x: 0, y: 0 });
    expect(scale).toEqual({ x: 1.5, y: 1 });
  });

  it("shifts the offset for a vertex dragged past the left edge", () => {
    const path: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: -0.2, y: 0 },
            { op: "line", x: 1, y: 0 },
            { op: "line", x: 1, y: 1 },
            { op: "close" },
          ],
          closed: true,
        },
      ],
    };
    const { offset, scale } = normalisePath(path);
    expect(offset.x).toBeCloseTo(-0.2);
    expect(scale.x).toBeCloseTo(1.2);
  });

  it("keeps a zero-height axis at zero without dividing by zero", () => {
    const line: ShapePath = {
      subpaths: [{ steps: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 1, y: 0 }], closed: false }],
    };
    const { path: fitted, scale } = normalisePath(line);
    expect(scale.y).toBe(0);
    expect(fitted.subpaths[0].steps.every((s) => s.op === "close" || s.y === 0)).toBe(true);
  });
});

describe("pointInPath (non-zero winding)", () => {
  it("is inside a filled square", () => {
    expect(pointInPath({ x: 5, y: 5 }, square(), 10, 10)).toBe(true);
    expect(pointInPath({ x: 15, y: 5 }, square(), 10, 10)).toBe(false);
  });

  it("ignores open subpaths for fill", () => {
    const open: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0, y: 0 },
            { op: "line", x: 1, y: 0 },
            { op: "line", x: 1, y: 1 },
          ],
          closed: false,
        },
      ],
    };
    expect(pointInPath({ x: 5, y: 5 }, open, 10, 10)).toBe(false);
  });
});

describe("distanceToPath", () => {
  it("measures to the nearest edge of a closed path", () => {
    expect(distanceToPath({ x: 5, y: 0 }, square(), 10, 10)).toBeCloseTo(0);
    expect(distanceToPath({ x: 5, y: 2 }, square(), 10, 10)).toBeCloseTo(2);
  });

  it("includes the closing edge of a closed subpath", () => {
    // Point near the y=0..y=10 left edge (the close segment) of the square.
    expect(distanceToPath({ x: 1, y: 5 }, square(), 10, 10)).toBeCloseTo(1);
  });
});

describe("pathToSvgD", () => {
  it("scales fractions to real units and emits M/L/C/Z", () => {
    const path: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0, y: 0 },
            { op: "line", x: 1, y: 0 },
            { op: "cubic", x1: 1, y1: 0.5, x2: 0.5, y2: 1, x: 0, y: 1 },
            { op: "close" },
          ],
          closed: true,
        },
      ],
    };
    expect(pathToSvgD(path, 100, 200)).toBe("M 0 0 L 100 0 C 100 100 50 200 0 200 Z");
  });
});

describe("simplifyPolyline", () => {
  it("drops collinear points within tolerance", () => {
    const pts = [
      { x: 0, y: 0 },
      { x: 1, y: 0.001 },
      { x: 2, y: 0 },
      { x: 3, y: 0.001 },
      { x: 4, y: 0 },
    ];
    expect(simplifyPolyline(pts, 0.03)).toEqual([
      { x: 0, y: 0 },
      { x: 4, y: 0 },
    ]);
  });

  it("keeps a point that deviates beyond tolerance", () => {
    const pts = [
      { x: 0, y: 0 },
      { x: 1, y: 1 },
      { x: 2, y: 0 },
    ];
    expect(simplifyPolyline(pts, 0.03)).toHaveLength(3);
  });

  it("returns short polylines unchanged", () => {
    const pts = [{ x: 0, y: 0 }, { x: 1, y: 1 }];
    expect(simplifyPolyline(pts, 0.03)).toEqual(pts);
  });
});

describe("snapPathToGrid round-trips exactly", () => {
  it("keeps a snapped fraction identical after emit/read on the grid", () => {
    const snapped = snapPathToGrid(square());
    for (const s of snapped.subpaths[0].steps) {
      if (s.op === "close") {
        continue;
      }
      const emitted = Math.round(s.x * PATH_GRID);
      expect(emitted / PATH_GRID).toBe(s.x);
    }
  });

  it("snaps an arbitrary fraction to the nearest grid step", () => {
    const p: ShapePath = { subpaths: [{ steps: [{ op: "move", x: 1 / 3, y: 0 }], closed: false }] };
    const s = snapPathToGrid(p).subpaths[0].steps[0];
    if (s.op !== "move") {
      throw new Error("unreachable");
    }
    expect(s.x).toBe(Math.round((1 / 3) * PATH_GRID) / PATH_GRID);
  });
});

describe("mapPath", () => {
  it("transforms coordinates and preserves the smooth array", () => {
    const p: ShapePath = { ...square(), smooth: [false, true, false, false] };
    const out = mapPath(p, (pt) => ({ x: pt.x + 1, y: pt.y }));
    expect(out.smooth).toEqual([false, true, false, false]);
    const first = out.subpaths[0].steps[0];
    expect(first.op === "move" && first.x).toBe(1);
  });
});

/** Evaluate a cubic Bézier at `t` for one axis. */
function bez(p0: number, p1: number, p2: number, p3: number, t: number): number {
  const mt = 1 - t;
  return mt * mt * mt * p0 + 3 * mt * mt * t * p1 + 3 * mt * t * t * p2 + t * t * t * p3;
}

describe("quadraticToCubic", () => {
  it("is the exact degree elevation of a quadratic", () => {
    const c = quadraticToCubic({ x: 0, y: 0 }, { x: 6, y: 12 }, { x: 12, y: 0 });
    // Elevation: c1 = p0 + 2/3(ctrl-p0), c2 = p1 + 2/3(ctrl-p1).
    expect(c.x1).toBeCloseTo(4);
    expect(c.y1).toBeCloseTo(8);
    expect(c.x2).toBeCloseTo(8);
    expect(c.y2).toBeCloseTo(8);
    expect(c.x).toBe(12);
    expect(c.y).toBe(0);
    // Midpoint of the quadratic equals the midpoint of the elevated cubic.
    const quadMid = { x: 0.5 * 6 + 0.25 * 12, y: 0.5 * 12 };
    expect(bez(0, c.x1, c.x2, c.x, 0.5)).toBeCloseTo(quadMid.x);
    expect(bez(0, c.y1, c.y2, c.y, 0.5)).toBeCloseTo(quadMid.y);
  });
});

describe("arcToCubics", () => {
  it("approximates a quarter circle to within a pixel", () => {
    // Quarter of the unit circle, angle 0 -> pi/2, centred at origin.
    const segs = arcToCubics(0, 0, 1, 1, 0, Math.PI / 2);
    expect(segs).toHaveLength(1);
    const s = segs[0];
    // Endpoint at (0,1).
    expect(s.x).toBeCloseTo(0);
    expect(s.y).toBeCloseTo(1);
    // Sample the midpoint (t=0.5); it should sit on the unit circle. The
    // segment starts at (1,0).
    const x = bez(1, s.x1, s.x2, s.x, 0.5);
    const y = bez(0, s.y1, s.y2, s.y, 0.5);
    expect(Math.hypot(x, y)).toBeCloseTo(1, 3);
  });

  it("splits a full circle into four segments", () => {
    expect(arcToCubics(0, 0, 1, 1, 0, 2 * Math.PI)).toHaveLength(4);
  });
});
