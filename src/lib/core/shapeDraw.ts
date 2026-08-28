// SPEC-071 drawing gestures, as pure geometry so the freeform/curve/scribble
// tools share one tested core and the Vue layer only wires pointer events to
// it. Every builder takes points in a single flat coordinate space (the draw
// overlay's cm) and returns a `ShapePath` in that same space; `drawnShapeBox`
// then re-fits it to a 0..1 box via `normalisePath`, giving the shape's own
// width/height/offset with the path snapped to the shared grid.
import type { Point } from "./shapeGeometry";
import type { ShapeGeometry } from "./shapes";
import { resolvedShapeAdjustments, shapeDefinition } from "./shapeCatalogue";
import { normalisePath, snapPathToGrid, type PathStep, type ShapePath } from "./shapePath";

/** Scribble simplification tolerance (SPEC-071 "Scribble simplification"):
 * a stated constant, not a magic number in the gesture code — the saved
 * file's size depends on it. */
export const SCRIBBLE_TOLERANCE_CM = 0.03;

/** How near the origin (in cm) a Freeform/Curve click must land to close the
 * subpath instead of extending it (SPEC-071 "Drawing gestures"). */
export const CLOSE_DISTANCE_CM = 0.2;

/** A straight-edged path through `points` (Freeform, AC-7): move to the
 * first, a line to each of the rest, closed with a `close` step when asked.
 * Every vertex is a corner (no `smooth` entry set). */
export function polylinePath(points: Point[], closed: boolean): ShapePath {
  const steps: PathStep[] = [{ op: "move", x: points[0].x, y: points[0].y }];
  for (let i = 1; i < points.length; i++) {
    steps.push({ op: "line", x: points[i].x, y: points[i].y });
  }
  if (closed) {
    steps.push({ op: "close" });
  }
  return { subpaths: [{ steps, closed }] };
}

/** A smooth cubic spline through `points` via Catmull–Rom-to-Bézier (Curve,
 * AC-8; the scribble smoother, AC-9), so the curve passes through every
 * placed point. Each control point is a sixth of the chord between a point's
 * neighbours. Every vertex is marked smooth. `closed` wraps the neighbour
 * lookup and adds a closing cubic back to the first point. */
export function catmullRomPath(points: Point[], closed: boolean): ShapePath {
  if (points.length < 2) {
    return polylinePath(points, closed);
  }
  const n = points.length;
  const at = (i: number): Point => {
    if (closed) {
      return points[((i % n) + n) % n];
    }
    return points[Math.max(0, Math.min(n - 1, i))];
  };
  const steps: PathStep[] = [{ op: "move", x: points[0].x, y: points[0].y }];
  const last = closed ? n : n - 1;
  for (let i = 0; i < last; i++) {
    const p0 = at(i - 1);
    const p1 = at(i);
    const p2 = at(i + 1);
    const p3 = at(i + 2);
    steps.push({
      op: "cubic",
      x1: p1.x + (p2.x - p0.x) / 6,
      y1: p1.y + (p2.y - p0.y) / 6,
      x2: p2.x - (p3.x - p1.x) / 6,
      y2: p2.y - (p3.y - p1.y) / 6,
      x: p2.x,
      y: p2.y,
    });
  }
  if (closed) {
    steps.push({ op: "close" });
  }
  // `smooth` is parallel to on-curve vertex order (one entry per non-close
  // step): the `move` plus one per cubic. A closed curve adds a closing cubic
  // back to the start, so it has one more on-curve vertex than input points —
  // the array must match or the trailing vertex reads as a corner in
  // point-edit.
  const vertexCount = 1 + last;
  return { subpaths: [{ steps, closed }], smooth: Array.from({ length: vertexCount }, () => true) };
}

/** Converts a preset geometry to an equivalent freeform `path` (SPEC-071
 * AC-15): SPEC-069's `ShapeDefinition.outline` already gives the outline as
 * a closed polygon in the unit square, so the path is that polygon as
 * move/line/close with corner vertices. Returns null for a geometry with no
 * closed outline (the two line geometries), which point-edit does not enter. */
export function presetToPath(geometry: ShapeGeometry, adjustments?: readonly number[]): ShapePath | null {
  const def = shapeDefinition(geometry);
  if (!def || def.isLine) {
    return null;
  }
  const outline = def.outline(resolvedShapeAdjustments(geometry, adjustments));
  if (outline.length < 2) {
    return null;
  }
  return snapPathToGrid(polylinePath(outline, true));
}

/** Re-fits a drawn path (in overlay cm) to a 0..1 box and reports the box
 * the caller gives the shape: `widthCm`/`heightCm` are the path's extent,
 * `offset` its top-left in the same cm space. The path is snapped to the
 * shared grid so a save→reopen is an exact identity. A zero-extent axis
 * (a perfectly straight horizontal/vertical draw) keeps a small minimum so
 * the shape has a clickable box. */
export function drawnShapeBox(rawPath: ShapePath): {
  path: ShapePath;
  widthCm: number;
  heightCm: number;
  offset: Point;
} {
  const { path, offset, scale } = normalisePath(rawPath);
  return {
    path: snapPathToGrid(path),
    widthCm: scale.x,
    heightCm: scale.y,
    offset,
  };
}
