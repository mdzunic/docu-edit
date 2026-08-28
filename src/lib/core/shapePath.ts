// SPEC-071: the path model for custom/freeform geometry. A `ShapePath` is a
// list of subpaths of move/line/cubic/close steps in the shape's own unit
// square (0..1 in each axis) — the same coordinate space a preset outline
// uses (SPEC-069's `ShapeOutlinePath`), so scaling, rotating and hit-testing
// a hand-drawn path is the same arithmetic as for a catalogue shape.
//
// Product-neutral core (boundary.spec.ts): no document model, no Vue, no
// docx/odf reader. The only import is the type-only `Point` from
// shapeGeometry.ts (erased at build, so no runtime cycle with the widened
// hit-testing there).
import type { Point } from "./shapeGeometry";

/** One step of a subpath. Coordinates are fractions of the shape's box
 *  (0..1 in each axis; values outside are legal mid-edit and renormalised
 *  on commit). A `cubic` carries its two control points then its end. */
export type PathStep =
  | { op: "move"; x: number; y: number }
  | { op: "line"; x: number; y: number }
  | { op: "cubic"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: "close" };

export interface ShapeSubpath {
  steps: PathStep[];
  /** True when the subpath ends with a close; fill applies only to closed
   *  subpaths. */
  closed: boolean;
}

export interface ShapePath {
  subpaths: ShapeSubpath[];
  /** Per-vertex smoothness, parallel to the on-curve vertex order across all
   *  subpaths, so entering point-edit mode on a re-opened document restores
   *  the right handles instead of inferring them from the control points. */
  smooth?: boolean[];
}

/** The whole-number grid both writers use (Word's own `a:custGeom` extent).
 *  Model coordinates are snapped to it on every commit so a save→reopen is
 *  an exact identity: a fraction is always `n / PATH_GRID`, the writer emits
 *  `n`, and the reader reads `n / PATH_GRID` back to the same double. */
export const PATH_GRID = 21600;

/** How finely a cubic is sampled when flattened for hit-testing and bounds
 *  — enough that a screen-sized curve reads as smooth without the winding
 *  test paying for hundreds of segments. */
const CUBIC_SAMPLES = 24;

/** The end (on-curve) point of a step, or null for a `close`. */
function stepEnd(step: PathStep): Point | null {
  if (step.op === "close") {
    return null;
  }
  return { x: step.x, y: step.y };
}

/** Every coordinate a step carries — anchors and control points — so bounds
 *  contain the curve's convex hull and never clip a control-heavy cubic. */
function stepPoints(step: PathStep): Point[] {
  switch (step.op) {
    case "move":
    case "line":
      return [{ x: step.x, y: step.y }];
    case "cubic":
      return [
        { x: step.x1, y: step.y1 },
        { x: step.x2, y: step.y2 },
        { x: step.x, y: step.y },
      ];
    case "close":
      return [];
  }
}

/** The axis-aligned extent of a path in its own fractional space, covering
 *  anchors and control points. An empty path is a zero box at the origin. */
export function pathBounds(path: ShapePath): { x: number; y: number; w: number; h: number } {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const subpath of path.subpaths) {
    for (const step of subpath.steps) {
      for (const p of stepPoints(step)) {
        minX = Math.min(minX, p.x);
        minY = Math.min(minY, p.y);
        maxX = Math.max(maxX, p.x);
        maxY = Math.max(maxY, p.y);
      }
    }
  }
  if (!Number.isFinite(minX)) {
    return { x: 0, y: 0, w: 0, h: 0 };
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

function mapStep(step: PathStep, fn: (p: Point) => Point): PathStep {
  switch (step.op) {
    case "move":
    case "line": {
      const p = fn({ x: step.x, y: step.y });
      return { op: step.op, x: p.x, y: p.y };
    }
    case "cubic": {
      const c1 = fn({ x: step.x1, y: step.y1 });
      const c2 = fn({ x: step.x2, y: step.y2 });
      const end = fn({ x: step.x, y: step.y });
      return { op: "cubic", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y };
    }
    case "close":
      return step;
  }
}

/** Applies a coordinate transform to every step of a path, preserving the
 *  `smooth` array (which is parallel to vertex order, not coordinates). */
export function mapPath(path: ShapePath, fn: (p: Point) => Point): ShapePath {
  const out: ShapePath = {
    subpaths: path.subpaths.map((s) => ({ steps: s.steps.map((step) => mapStep(step, fn)), closed: s.closed })),
  };
  if (path.smooth) {
    out.smooth = [...path.smooth];
  }
  return out;
}

/** Re-fits a path's extent to the unit square and returns the `offset` and
 *  `scale` (both in fractional units) the caller folds into the shape's own
 *  box so nothing moves on screen: `xCm += offset.x * widthCm`,
 *  `widthCm *= scale.x`, and likewise for y. A dragged vertex outside the
 *  old box thus enlarges the box rather than being clipped (AC-23). The
 *  direct analogue of SPEC-057's container renormalisation. A zero-width or
 *  zero-height extent (a horizontal/vertical line) keeps that axis at 0. */
export function normalisePath(path: ShapePath): {
  path: ShapePath;
  offset: { x: number; y: number };
  scale: { x: number; y: number };
} {
  const b = pathBounds(path);
  const sx = b.w === 0 ? 1 : b.w;
  const sy = b.h === 0 ? 1 : b.h;
  const refitted = mapPath(path, (p) => ({ x: (p.x - b.x) / sx, y: (p.y - b.y) / sy }));
  return { path: refitted, offset: { x: b.x, y: b.y }, scale: { x: b.w, y: b.h } };
}

/** Snaps every coordinate to the shared whole-number grid, so a committed
 *  model path round-trips through either format as an exact identity. */
export function snapPathToGrid(path: ShapePath): ShapePath {
  return mapPath(path, (p) => ({
    x: Math.round(p.x * PATH_GRID) / PATH_GRID,
    y: Math.round(p.y * PATH_GRID) / PATH_GRID,
  }));
}

/** Samples a cubic Bézier from `p0` through `p1`,`p2` to `p3` into
 *  `CUBIC_SAMPLES` points (excluding the start, which the caller already
 *  holds). */
function sampleCubic(p0: Point, p1: Point, p2: Point, p3: Point): Point[] {
  const out: Point[] = [];
  for (let i = 1; i <= CUBIC_SAMPLES; i++) {
    const t = i / CUBIC_SAMPLES;
    const mt = 1 - t;
    const a = mt * mt * mt;
    const b = 3 * mt * mt * t;
    const c = 3 * mt * t * t;
    const d = t * t * t;
    out.push({
      x: a * p0.x + b * p1.x + c * p2.x + d * p3.x,
      y: a * p0.y + b * p1.y + c * p2.y + d * p3.y,
    });
  }
  return out;
}

/** Flattens one subpath to a polyline in real (w x h) coordinates. The
 *  returned point list is the outline in order; closure is left implicit
 *  (the winding test closes it, the distance test does not). */
function flattenSubpath(subpath: ShapeSubpath, w: number, h: number): Point[] {
  const scale = (p: Point): Point => ({ x: p.x * w, y: p.y * h });
  const points: Point[] = [];
  let current: Point | null = null;
  for (const step of subpath.steps) {
    if (step.op === "move") {
      current = scale({ x: step.x, y: step.y });
      points.push(current);
    } else if (step.op === "line") {
      current = scale({ x: step.x, y: step.y });
      points.push(current);
    } else if (step.op === "cubic" && current) {
      const c1 = scale({ x: step.x1, y: step.y1 });
      const c2 = scale({ x: step.x2, y: step.y2 });
      const end = scale({ x: step.x, y: step.y });
      points.push(...sampleCubic(current, c1, c2, end));
      current = end;
    }
  }
  return points;
}

/** The signed contribution of edge `a`→`b` to the winding number around
 *  `pt` (positive to the left). */
function isLeft(a: Point, b: Point, pt: Point): number {
  return (b.x - a.x) * (pt.y - a.y) - (pt.x - a.x) * (b.y - a.y);
}

/** Whether `point` (in real w x h coordinates) lands inside the path's
 *  filled region under the **non-zero winding** rule (SPEC-071 states this
 *  as the fill rule). Only closed subpaths contribute. */
export function pointInPath(point: Point, path: ShapePath, w: number, h: number): boolean {
  let winding = 0;
  for (const subpath of path.subpaths) {
    if (!subpath.closed) {
      continue;
    }
    const poly = flattenSubpath(subpath, w, h);
    for (let i = 0; i < poly.length; i++) {
      const a = poly[i];
      const b = poly[(i + 1) % poly.length];
      if (a.y <= point.y) {
        if (b.y > point.y && isLeft(a, b, point) > 0) {
          winding++;
        }
      } else if (b.y <= point.y && isLeft(a, b, point) < 0) {
        winding--;
      }
    }
  }
  return winding !== 0;
}

function distanceToSegment(point: Point, a: Point, b: Point): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const lengthSq = dx * dx + dy * dy;
  if (lengthSq === 0) {
    return Math.hypot(point.x - a.x, point.y - a.y);
  }
  const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / lengthSq));
  return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
}

/** Shortest distance from `point` (real w x h coordinates) to any of the
 *  path's outlines — closed subpaths include their closing edge. */
export function distanceToPath(point: Point, path: ShapePath, w: number, h: number): number {
  let min = Infinity;
  for (const subpath of path.subpaths) {
    const poly = flattenSubpath(subpath, w, h);
    const last = subpath.closed ? poly.length : poly.length - 1;
    for (let i = 0; i < last; i++) {
      const a = poly[i];
      const b = poly[(i + 1) % poly.length];
      min = Math.min(min, distanceToSegment(point, a, b));
    }
    if (poly.length === 1) {
      min = Math.min(min, Math.hypot(point.x - poly[0].x, point.y - poly[0].y));
    }
  }
  return min;
}

function num(n: number): string {
  // Trim floating dust so the `d` string and the ODF/OOXML integers a save
  // produces stay stable and readable.
  return Number.isInteger(n) ? String(n) : n.toFixed(3).replace(/\.?0+$/, "");
}

/** An SVG path `d` string in real (w x h) coordinates — used by the page
 *  canvas and the gallery icon to paint the real outline. */
export function pathToSvgD(path: ShapePath, w: number, h: number): string {
  const parts: string[] = [];
  for (const subpath of path.subpaths) {
    for (const step of subpath.steps) {
      if (step.op === "move") {
        parts.push(`M ${num(step.x * w)} ${num(step.y * h)}`);
      } else if (step.op === "line") {
        parts.push(`L ${num(step.x * w)} ${num(step.y * h)}`);
      } else if (step.op === "cubic") {
        parts.push(
          `C ${num(step.x1 * w)} ${num(step.y1 * h)} ${num(step.x2 * w)} ${num(step.y2 * h)} ${num(step.x * w)} ${num(step.y * h)}`,
        );
      } else {
        parts.push("Z");
      }
    }
  }
  return parts.join(" ");
}

/** Ramer–Douglas–Peucker polyline simplification: drops points closer than
 *  `toleranceCm` to the chord they lie on. Used to turn a scribble's few
 *  hundred samples into a few dozen (SPEC-071 "Scribble simplification").
 *  The tolerance is in the same units as the points. */
export function simplifyPolyline(points: Point[], toleranceCm: number): Point[] {
  if (points.length <= 2) {
    return points.slice();
  }
  const keep = new Array<boolean>(points.length).fill(false);
  keep[0] = true;
  keep[points.length - 1] = true;
  const stack: [number, number][] = [[0, points.length - 1]];
  while (stack.length > 0) {
    const [start, end] = stack.pop() as [number, number];
    let maxDist = -1;
    let index = -1;
    for (let i = start + 1; i < end; i++) {
      const d = distanceToSegment(points[i], points[start], points[end]);
      if (d > maxDist) {
        maxDist = d;
        index = i;
      }
    }
    if (maxDist > toleranceCm && index !== -1) {
      keep[index] = true;
      stack.push([start, index], [index, end]);
    }
  }
  return points.filter((_, i) => keep[i]);
}

/** The flattened outline of the path's first subpath, in unit-square
 *  fractions — used to build a CSS `clip-path: polygon(...)` for a filled
 *  path shape's DOM proxy so a click in an empty corner passes through to
 *  the shape beneath, the same role `shapePolygon` plays for a preset. A
 *  polygon cannot express holes or disjoint subpaths, so only the first
 *  subpath is used; that is the common single-outline freeform, and the
 *  precise hit test (`pointInPath`) is the source of truth regardless. */
export function pathToClipPolygon(path: ShapePath): Point[] {
  const first = path.subpaths.find((s) => s.closed);
  if (!first) {
    return [];
  }
  return flattenSubpath(first, 1, 1);
}

// --- primitive reduction (SPEC-071 "Primitive reduction") -------------------
//
// The model carries only move/line/cubic/close. Both formats have primitives
// it does not — OOXML `a:quadBezTo`/`a:arcTo`, ODF `Q`/`T`/`A` — so a reader
// converts them to cubics at its boundary using these pure helpers.

/** The exact cubic equivalent of a quadratic Bézier `p0`→(control `c`)→`p1`.
 *  A quadratic raised to cubic degree has control points at
 *  `p + 2/3(c - p)` on each end; the conversion is lossless. */
export function quadraticToCubic(
  p0: Point,
  c: Point,
  p1: Point,
): { x1: number; y1: number; x2: number; y2: number; x: number; y: number } {
  return {
    x1: p0.x + (2 / 3) * (c.x - p0.x),
    y1: p0.y + (2 / 3) * (c.y - p0.y),
    x2: p1.x + (2 / 3) * (c.x - p1.x),
    y2: p1.y + (2 / 3) * (c.y - p1.y),
    x: p1.x,
    y: p1.y,
  };
}

/** Cubic segments approximating an elliptical arc centred at `(cx, cy)` with
 *  radii `rx`,`ry`, sweeping `deltaAngle` radians from `startAngle` (measured
 *  in the ellipse's own axes; positive is the +x→+y direction). The arc is
 *  split into ≤90° pieces, each approximated by the standard `k = 4/3
 *  tan(θ/4)` cubic — accurate to well within one screen pixel, which is the
 *  stated tolerance. Returns each piece's two control points and its end;
 *  the caller already holds the start. */
export function arcToCubics(
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  startAngle: number,
  deltaAngle: number,
): { x1: number; y1: number; x2: number; y2: number; x: number; y: number }[] {
  const segments = Math.max(1, Math.ceil(Math.abs(deltaAngle) / (Math.PI / 2)));
  const delta = deltaAngle / segments;
  const k = (4 / 3) * Math.tan(delta / 4);
  const out: { x1: number; y1: number; x2: number; y2: number; x: number; y: number }[] = [];
  let angle = startAngle;
  for (let i = 0; i < segments; i++) {
    const a1 = angle;
    const a2 = angle + delta;
    const p1 = { x: cx + rx * Math.cos(a1), y: cy + ry * Math.sin(a1) };
    const p2 = { x: cx + rx * Math.cos(a2), y: cy + ry * Math.sin(a2) };
    const t1 = { x: -rx * Math.sin(a1), y: ry * Math.cos(a1) };
    const t2 = { x: -rx * Math.sin(a2), y: ry * Math.cos(a2) };
    out.push({
      x1: p1.x + k * t1.x,
      y1: p1.y + k * t1.y,
      x2: p2.x - k * t2.x,
      y2: p2.y - k * t2.y,
      x: p2.x,
      y: p2.y,
    });
    angle = a2;
  }
  return out;
}
