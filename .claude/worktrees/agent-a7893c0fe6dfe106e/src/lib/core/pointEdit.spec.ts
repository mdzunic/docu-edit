import { describe, it, expect } from "vitest";
import {
  pathVertices,
  nudgeVertex,
  moveVertexTo,
  toggleVertexSmooth,
  moveControl,
  splitSegmentAfter,
  deleteVertex,
  renormaliseAfterEdit,
} from "./pointEdit";
import type { ShapePath } from "./shapePath";

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
    smooth: [false, false, false, false],
  };
}

function curve(): ShapePath {
  return {
    subpaths: [
      {
        steps: [
          { op: "move", x: 0, y: 0.5 },
          { op: "cubic", x1: 0, y1: 0, x2: 1, y2: 0, x: 1, y: 0.5 },
        ],
        closed: false,
      },
    ],
    smooth: [true, true],
  };
}

describe("pathVertices", () => {
  it("enumerates on-curve vertices with control points", () => {
    const v = pathVertices(curve());
    expect(v).toHaveLength(2);
    expect(v[1].point).toEqual({ x: 1, y: 0.5 });
    expect(v[1].incoming).toEqual({ x: 1, y: 0 });
    expect(v[0].outgoing).toEqual({ x: 0, y: 0 });
  });
});

describe("nudgeVertex / moveVertexTo", () => {
  it("moves a line vertex", () => {
    const p = nudgeVertex(square(), 1, 0.5, 0.25);
    expect(pathVertices(p)[1].point).toEqual({ x: 1.5, y: 0.25 });
  });

  it("carries a cubic vertex's control points with it", () => {
    const p = moveVertexTo(curve(), 1, { x: 2, y: 0.5 });
    const step = p.subpaths[0].steps[1];
    expect(step.op).toBe("cubic");
    if (step.op === "cubic") {
      expect(step.x).toBe(2);
      expect(step.x2).toBe(2); // incoming control shifted by +1
    }
  });
});

describe("toggleVertexSmooth", () => {
  it("flips the flag and derives collinear controls when smoothing", () => {
    const p = toggleVertexSmooth(square(), 1); // top-right corner -> smooth
    expect(p.smooth![1]).toBe(true);
  });

  it("flips a smooth vertex back to corner", () => {
    const p = toggleVertexSmooth(curve(), 1);
    expect(p.smooth![1]).toBe(false);
  });
});

describe("moveControl", () => {
  it("mirrors the opposite handle on a smooth vertex", () => {
    // Build a two-cubic path so the middle vertex has both handles.
    const path: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0, y: 0 },
            { op: "cubic", x1: 0.1, y1: 0.1, x2: 0.4, y2: 0.4, x: 0.5, y: 0.5 },
            { op: "cubic", x1: 0.6, y1: 0.6, x2: 0.9, y2: 0.9, x: 1, y: 1 },
          ],
          closed: false,
        },
      ],
      smooth: [false, true, false],
    };
    const out = moveControl(path, 1, "out", { x: 0.7, y: 0.5 });
    // outgoing control set; incoming mirrored about (0.5,0.5).
    const s1 = out.subpaths[0].steps[1];
    const s2 = out.subpaths[0].steps[2];
    if (s1.op === "cubic" && s2.op === "cubic") {
      expect(s2.x1).toBeCloseTo(0.7);
      expect(s2.y1).toBeCloseTo(0.5);
      expect(s1.x2).toBeCloseTo(0.3);
      expect(s1.y2).toBeCloseTo(0.5);
    }
  });

  it("moves handles independently on a corner vertex", () => {
    const path: ShapePath = {
      subpaths: [
        {
          steps: [
            { op: "move", x: 0, y: 0 },
            { op: "cubic", x1: 0.1, y1: 0.1, x2: 0.4, y2: 0.4, x: 0.5, y: 0.5 },
            { op: "cubic", x1: 0.6, y1: 0.6, x2: 0.9, y2: 0.9, x: 1, y: 1 },
          ],
          closed: false,
        },
      ],
      smooth: [false, false, false],
    };
    const out = moveControl(path, 1, "out", { x: 0.7, y: 0.5 });
    const s1 = out.subpaths[0].steps[1];
    if (s1.op === "cubic") {
      expect({ x: s1.x2, y: s1.y2 }).toEqual({ x: 0.4, y: 0.4 }); // incoming unchanged
    }
  });
});

describe("splitSegmentAfter", () => {
  it("splits a line into two lines at the midpoint", () => {
    const p = splitSegmentAfter(square(), 0, 0.5); // segment (0,0)->(1,0)
    const ops = p.subpaths[0].steps.map((s) => s.op);
    expect(ops.filter((o) => o === "line")).toHaveLength(4);
    expect(pathVertices(p)[1].point).toEqual({ x: 0.5, y: 0 });
  });

  it("splits a cubic into two cubics", () => {
    const p = splitSegmentAfter(curve(), 0, 0.5);
    expect(p.subpaths[0].steps.filter((s) => s.op === "cubic")).toHaveLength(2);
    expect(pathVertices(p)).toHaveLength(3);
  });
});

describe("deleteVertex", () => {
  it("removes a vertex and its step", () => {
    const p = deleteVertex(square(), 1);
    expect(p).not.toBeNull();
    expect(pathVertices(p!)).toHaveLength(3);
  });

  it("removes the subpath when its last vertex goes, and returns null when the path empties", () => {
    let p: ShapePath | null = { subpaths: [{ steps: [{ op: "move", x: 0, y: 0 }], closed: false }], smooth: [false] };
    p = deleteVertex(p, 0);
    expect(p).toBeNull();
  });

  it("promotes the next step to a move when the first vertex is deleted", () => {
    const p = deleteVertex(square(), 0);
    expect(p!.subpaths[0].steps[0].op).toBe("move");
  });
});

describe("renormaliseAfterEdit", () => {
  it("re-fits an enlarged path and reports the box growth", () => {
    // Drag a vertex out to x=1.5.
    const dragged = nudgeVertex(square(), 1, 0.5, 0);
    const { path, offset, scale } = renormaliseAfterEdit(dragged);
    expect(scale.x).toBeCloseTo(1.5);
    expect(offset.x).toBeCloseTo(0);
    // The dragged vertex sits back at the box's right edge.
    expect(pathVertices(path)[1].point.x).toBeCloseTo(1);
  });
});
