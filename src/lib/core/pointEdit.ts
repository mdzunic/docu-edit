// SPEC-071 point-edit mode as pure model operations on a `ShapePath`: the
// spec's whole premise is that "move a vertex", "toggle smooth/corner",
// "split a segment" and "delete a vertex" are model edits, not a new
// subsystem, so they live here — framework-free and unit-tested — and the
// Vue overlay only maps handle drags and keys onto them.
import type { Point } from "./shapeGeometry";
import { normalisePath, snapPathToGrid, type PathStep, type ShapePath } from "./shapePath";

/** One on-curve vertex, addressed by its subpath and the index of the step
 * whose end point it is. `incoming`/`outgoing` are its control points where
 * an adjacent segment is a cubic (absent for a line/at an open end). */
export interface Vertex {
  index: number;
  subpath: number;
  step: number;
  point: Point;
  smooth: boolean;
  incoming?: Point;
  outgoing?: Point;
}

function endOf(step: PathStep): Point | null {
  return step.op === "close" ? null : { x: step.x, y: step.y };
}

/** Enumerates a path's on-curve vertices in order across all subpaths,
 * pairing each with its neighbouring control points and its `smooth` flag. */
export function pathVertices(path: ShapePath): Vertex[] {
  const vertices: Vertex[] = [];
  let index = 0;
  path.subpaths.forEach((subpath, sp) => {
    const onCurve = subpath.steps
      .map((step, step_) => ({ step, step_ }))
      .filter((s) => s.step.op !== "close");
    onCurve.forEach(({ step, step_ }, localIndex) => {
      const point = endOf(step)!;
      const outgoingStep = subpath.steps[step_ + 1];
      const incoming = step.op === "cubic" ? { x: step.x2, y: step.y2 } : undefined;
      const outgoing =
        outgoingStep && outgoingStep.op === "cubic" ? { x: outgoingStep.x1, y: outgoingStep.y1 } : undefined;
      vertices.push({
        index,
        subpath: sp,
        step: step_,
        point,
        smooth: path.smooth?.[index] ?? false,
        incoming,
        outgoing,
        // localIndex kept implicitly by order; not stored.
      });
      index++;
      void localIndex;
    });
  });
  return vertices;
}

function clonePath(path: ShapePath): ShapePath {
  return {
    subpaths: path.subpaths.map((s) => ({ steps: s.steps.map((step) => ({ ...step })), closed: s.closed })),
    smooth: path.smooth ? [...path.smooth] : undefined,
  };
}

/** Moves vertex `index` by (dx, dy) in 0..1 space, carrying its attached
 * control points (the incoming cubic's second control and the outgoing
 * cubic's first control) so the curve around it translates rather than
 * kinking. */
export function nudgeVertex(path: ShapePath, index: number, dx: number, dy: number): ShapePath {
  const out = clonePath(path);
  const v = pathVertices(path)[index];
  if (!v) {
    return out;
  }
  const step = out.subpaths[v.subpath].steps[v.step];
  if (step.op === "line" || step.op === "move") {
    step.x += dx;
    step.y += dy;
  } else if (step.op === "cubic") {
    step.x += dx;
    step.y += dy;
    step.x2 += dx;
    step.y2 += dy;
  }
  const next = out.subpaths[v.subpath].steps[v.step + 1];
  if (next && next.op === "cubic") {
    next.x1 += dx;
    next.y1 += dy;
  }
  return out;
}

/** Sets vertex `index` to an absolute position. */
export function moveVertexTo(path: ShapePath, index: number, point: Point): ShapePath {
  const v = pathVertices(path)[index];
  if (!v) {
    return clonePath(path);
  }
  return nudgeVertex(path, index, point.x - v.point.x, point.y - v.point.y);
}

/** Toggles a vertex between corner and smooth. Making it smooth re-derives
 * its two control points collinear through it (a sixth of the chord between
 * its neighbours, matching the Catmull–Rom draw); making it a corner leaves
 * the control points where they are. */
export function toggleVertexSmooth(path: ShapePath, index: number): ShapePath {
  const out = clonePath(path);
  const vertices = pathVertices(path);
  const v = vertices[index];
  if (!v) {
    return out;
  }
  const smooth = out.smooth ?? vertices.map((vertex) => vertex.smooth);
  smooth[index] = !smooth[index];
  out.smooth = smooth;
  if (smooth[index]) {
    // Neighbour points within the same subpath (clamped at open ends).
    const sameSub = vertices.filter((w) => w.subpath === v.subpath);
    const local = sameSub.findIndex((w) => w.index === index);
    const prev = sameSub[local - 1] ?? v;
    const next = sameSub[local + 1] ?? v;
    const tx = (next.point.x - prev.point.x) / 6;
    const ty = (next.point.y - prev.point.y) / 6;
    const incomingStep = out.subpaths[v.subpath].steps[v.step];
    if (incomingStep.op === "cubic") {
      incomingStep.x2 = v.point.x - tx;
      incomingStep.y2 = v.point.y - ty;
    }
    const outStep = out.subpaths[v.subpath].steps[v.step + 1];
    if (outStep && outStep.op === "cubic") {
      outStep.x1 = v.point.x + tx;
      outStep.y1 = v.point.y + ty;
    }
  }
  return out;
}

/** Moves one control handle of a vertex. On a smooth vertex the opposite
 * handle mirrors the movement (kept collinear and equidistant); on a corner
 * vertex the two move independently (AC-18). `which` is the outgoing
 * (`"out"`) or incoming (`"in"`) handle. */
export function moveControl(path: ShapePath, index: number, which: "in" | "out", point: Point): ShapePath {
  const out = clonePath(path);
  const vertices = pathVertices(path);
  const v = vertices[index];
  if (!v) {
    return out;
  }
  const incomingStep = out.subpaths[v.subpath].steps[v.step];
  const outStep = out.subpaths[v.subpath].steps[v.step + 1];
  const setIn = (p: Point): void => {
    if (incomingStep.op === "cubic") {
      incomingStep.x2 = p.x;
      incomingStep.y2 = p.y;
    }
  };
  const setOut = (p: Point): void => {
    if (outStep && outStep.op === "cubic") {
      outStep.x1 = p.x;
      outStep.y1 = p.y;
    }
  };
  const mirror: Point = { x: 2 * v.point.x - point.x, y: 2 * v.point.y - point.y };
  if (which === "out") {
    setOut(point);
    if (v.smooth) {
      setIn(mirror);
    }
  } else {
    setIn(point);
    if (v.smooth) {
      setOut(mirror);
    }
  }
  return out;
}

/** Splits the segment leaving vertex `index` by inserting a new corner
 * vertex at parameter `t` (0..1). A line splits into two lines, a cubic into
 * two cubics via de Casteljau (AC-17). Returns the path unchanged when the
 * vertex has no following segment (an open end). */
export function splitSegmentAfter(path: ShapePath, index: number, t = 0.5): ShapePath {
  const out = clonePath(path);
  const vertices = pathVertices(path);
  const v = vertices[index];
  if (!v) {
    return out;
  }
  const steps = out.subpaths[v.subpath].steps;
  const start = v.point;
  const nextStep = steps[v.step + 1];
  // The following on-curve step is either the next step, or (past a close /
  // at the subpath end of a closed path) the subpath's first vertex.
  if (nextStep && nextStep.op === "line") {
    const mid = { x: start.x + (nextStep.x - start.x) * t, y: start.y + (nextStep.y - start.y) * t };
    steps.splice(v.step + 1, 0, { op: "line", x: mid.x, y: mid.y });
    insertSmooth(out, index + 1);
  } else if (nextStep && nextStep.op === "cubic") {
    const p0 = start;
    const p1 = { x: nextStep.x1, y: nextStep.y1 };
    const p2 = { x: nextStep.x2, y: nextStep.y2 };
    const p3 = { x: nextStep.x, y: nextStep.y };
    const a = lerp(p0, p1, t);
    const b = lerp(p1, p2, t);
    const c = lerp(p2, p3, t);
    const d = lerp(a, b, t);
    const e = lerp(b, c, t);
    const m = lerp(d, e, t);
    const first: PathStep = { op: "cubic", x1: a.x, y1: a.y, x2: d.x, y2: d.y, x: m.x, y: m.y };
    const second: PathStep = { op: "cubic", x1: e.x, y1: e.y, x2: c.x, y2: c.y, x: p3.x, y: p3.y };
    steps.splice(v.step + 1, 1, first, second);
    insertSmooth(out, index + 1);
  }
  return out;
}

function lerp(a: Point, b: Point, t: number): Point {
  return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
}

function insertSmooth(path: ShapePath, at: number): void {
  if (path.smooth) {
    path.smooth.splice(at, 0, false);
  }
}

/** Deletes vertex `index`. When it was the last vertex of its subpath the
 * subpath is removed; when that empties the path, null is returned so the
 * caller can delete the whole shape (AC-21). The step that ended at the
 * vertex is removed, keeping the subpath's remaining steps contiguous. */
export function deleteVertex(path: ShapePath, index: number): ShapePath | null {
  const out = clonePath(path);
  const vertices = pathVertices(path);
  const v = vertices[index];
  if (!v) {
    return out;
  }
  const subpath = out.subpaths[v.subpath];
  const onCurveCount = subpath.steps.filter((s) => s.op !== "close").length;
  if (onCurveCount <= 1) {
    out.subpaths.splice(v.subpath, 1);
  } else {
    const removed = subpath.steps[v.step];
    subpath.steps.splice(v.step, 1);
    // Removing the first vertex (a `move`) promotes the next step to a move
    // so the subpath still starts with one.
    if (removed.op === "move") {
      const first = subpath.steps.find((s) => s.op !== "close");
      if (first && first.op !== "move") {
        const i = subpath.steps.indexOf(first);
        subpath.steps[i] = { op: "move", x: first.x, y: first.y };
      }
    }
  }
  if (out.smooth) {
    out.smooth.splice(index, 1);
  }
  if (out.subpaths.length === 0) {
    return null;
  }
  return out;
}

/** Re-fits the path to 0..1 after an edit and returns the box adjustment the
 * caller folds into the shape (SPEC-071 AC-23): `xCm += offset.x * widthCm`,
 * `widthCm *= scale.x`, and likewise for y, so nothing moves on screen and a
 * vertex dragged outside the old box enlarges it. The path is re-snapped to
 * the shared grid. */
export function renormaliseAfterEdit(path: ShapePath): {
  path: ShapePath;
  offset: Point;
  scale: Point;
} {
  const { path: fitted, offset, scale } = normalisePath(path);
  return { path: snapPathToGrid(fitted), offset, scale };
}
