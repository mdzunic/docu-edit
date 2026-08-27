// SPEC-070: the product-neutral routing half of connectors. A connector is a
// line that remembers what it is attached to; this module owns the pure
// geometry — the nine connection sites derived from a shape's bounding box,
// the straight/elbow/curved routing between two resolved endpoints, and the
// centre-site clipping against a shape's outline. Geometry in, points out: no
// document model, no Vue, no reader — unit-testable without a document and
// reusable by a future Slides host (the core layer's boundary rule).
import type { Point } from "./shapeGeometry";
import { rotatePoint } from "./shapeGeometry";

/** How a connector routes between its two endpoints. The document model's
 * `Connector.kind` is this type (re-exported from documentModel.ts). */
export type ConnectorKind = "straight" | "elbow" | "curved";

/** The nine connection sites per shape, in the fixed on-disk index order:
 * 0 top, 1 right, 2 bottom, 3 left, 4 top-left, 5 top-right, 6 bottom-right,
 * 7 bottom-left, 8 centre ("Connection sites"). */
export const CONNECTION_SITE_COUNT = 9;

/** An axis-aligned box (before rotation) in any single unit — cm or px. */
export interface ConnectorBox {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Whole degrees clockwise about the box centre; absent means 0. */
  rotationDeg?: number;
}

/** A resolved connection site: its absolute attachment point and the unit
 * outward normal of the edge it sits on (the direction a connector leaves
 * the shape). The centre site (8) has a zero normal until it is clipped to
 * the outline toward the other endpoint (see clipCentreSite). */
export interface ConnectionSite {
  point: Point;
  normal: Point;
}

/** Local (unit-square) position of each site, indexed 0-8. */
const SITE_LOCAL: readonly Point[] = [
  { x: 0.5, y: 0 }, // 0 top
  { x: 1, y: 0.5 }, // 1 right
  { x: 0.5, y: 1 }, // 2 bottom
  { x: 0, y: 0.5 }, // 3 left
  { x: 0, y: 0 }, // 4 top-left
  { x: 1, y: 0 }, // 5 top-right
  { x: 1, y: 1 }, // 6 bottom-right
  { x: 0, y: 1 }, // 7 bottom-left
  { x: 0.5, y: 0.5 }, // 8 centre
];

const SQRT1_2 = Math.SQRT1_2;
/** Unit outward normal of each site (y-down screen convention). */
const SITE_NORMAL: readonly Point[] = [
  { x: 0, y: -1 }, // 0 top
  { x: 1, y: 0 }, // 1 right
  { x: 0, y: 1 }, // 2 bottom
  { x: -1, y: 0 }, // 3 left
  { x: -SQRT1_2, y: -SQRT1_2 }, // 4 top-left
  { x: SQRT1_2, y: -SQRT1_2 }, // 5 top-right
  { x: SQRT1_2, y: SQRT1_2 }, // 6 bottom-right
  { x: -SQRT1_2, y: SQRT1_2 }, // 7 bottom-left
  { x: 0, y: 0 }, // 8 centre — no edge, resolved by clipping
];

/** Rotates a direction vector by `rotationDeg` clockwise (y-down), matching
 * shapeGeometry's rotatePoint about the origin. */
function rotateVector(v: Point, rotationDeg: number): Point {
  if (!rotationDeg) {
    return v;
  }
  const rad = (rotationDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return { x: v.x * cos - v.y * sin, y: v.x * sin + v.y * cos };
}

/** The absolute point of a single site on `box`, honouring its rotation. */
export function connectionSitePoint(box: ConnectorBox, site: number): Point {
  const local = SITE_LOCAL[clampSite(site)];
  const rotated = rotatePoint({ x: local.x * box.w, y: local.y * box.h }, box.w, box.h, box.rotationDeg ?? 0);
  return { x: box.x + rotated.x, y: box.y + rotated.y };
}

/** All nine connection sites of `box`, in index order, honouring rotation. */
export function connectionSites(box: ConnectorBox): ConnectionSite[] {
  const sites: ConnectionSite[] = [];
  for (let i = 0; i < CONNECTION_SITE_COUNT; i++) {
    sites.push({
      point: connectionSitePoint(box, i),
      normal: rotateVector(SITE_NORMAL[i], box.rotationDeg ?? 0),
    });
  }
  return sites;
}

/** Clamps an arbitrary index into the valid 0-8 range (an unmappable format
 * index falls back to the centre). */
export function clampSite(site: number): number {
  return Number.isInteger(site) && site >= 0 && site < CONNECTION_SITE_COUNT ? site : 8;
}

/** The nearest site index to an absolute point — the fallback the writers use
 * when a format's own glue-point numbering has no exact match, and the reader
 * uses for the same reason ("Connection sites"). */
export function nearestSite(box: ConnectorBox, point: Point): number {
  let best = 8;
  let bestDist = Infinity;
  for (let i = 0; i < CONNECTION_SITE_COUNT; i++) {
    const p = connectionSitePoint(box, i);
    const d = (p.x - point.x) ** 2 + (p.y - point.y) ** 2;
    if (d < bestDist) {
      bestDist = d;
      best = i;
    }
  }
  return best;
}

/** Intersection of segment [a,b] with a closed polygon, nearest to `a`, or
 * null when the segment misses every edge. Used to clip a centre-site anchor
 * to the shape's outline so a connector to a centre never draws inside it. */
export function segmentPolygonCrossing(a: Point, b: Point, polygon: Point[]): Point | null {
  let best: Point | null = null;
  let bestT = Infinity;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const hit = segmentSegmentIntersection(a, b, polygon[j], polygon[i]);
    if (hit && hit.t < bestT) {
      bestT = hit.t;
      best = hit.point;
    }
  }
  return best;
}

function segmentSegmentIntersection(
  p1: Point,
  p2: Point,
  p3: Point,
  p4: Point,
): { point: Point; t: number } | null {
  const d1x = p2.x - p1.x;
  const d1y = p2.y - p1.y;
  const d2x = p4.x - p3.x;
  const d2y = p4.y - p3.y;
  const denom = d1x * d2y - d1y * d2x;
  if (Math.abs(denom) < 1e-9) {
    return null;
  }
  const t = ((p3.x - p1.x) * d2y - (p3.y - p1.y) * d2x) / denom;
  const u = ((p3.x - p1.x) * d1y - (p3.y - p1.y) * d1x) / denom;
  if (t < 0 || t > 1 || u < 0 || u > 1) {
    return null;
  }
  return { point: { x: p1.x + t * d1x, y: p1.y + t * d1y }, t };
}

/** One resolved connector endpoint: its anchor point and, when attached, the
 * outward normal of its site (absent for a free end). */
export interface ConnectorEndpoint {
  point: Point;
  normal?: Point;
}

/** Resolves a centre-site (8) endpoint to where the line toward the other
 * endpoint crosses `outline` (absolute polygon), so a connector to a centre
 * stops at the shape's edge. Returns the crossing point and the local
 * outward normal (direction from centre to crossing). Falls back to the raw
 * centre when the segment misses the outline. */
export function clipCentreSite(centre: Point, toward: Point, outline: Point[]): ConnectorEndpoint {
  const crossing = segmentPolygonCrossing(centre, toward, outline);
  if (!crossing) {
    return { point: centre };
  }
  const nx = crossing.x - centre.x;
  const ny = crossing.y - centre.y;
  const len = Math.hypot(nx, ny) || 1;
  return { point: crossing, normal: { x: nx / len, y: ny / len } };
}

/** A routed connector path: the ordered points of a poly-line (straight/
 * elbow) plus, for a curved kind, the two cubic control points. Everything is
 * in the caller's own unit (px for rendering, cm for a fixture). */
export interface RoutedPath {
  kind: ConnectorKind;
  points: Point[];
  /** For "curved": [control1, control2] between points[0] and points[last]. */
  controls?: [Point, Point];
}

/** How far a stub leaves a site before the elbow turns, as a fraction of the
 * distance between the endpoints (clamped so tiny connectors still turn). */
const STUB_FRACTION = 0.5;

/** Routes a connector between two resolved endpoints. Pure: geometry in,
 * points out ("Routing").
 * - straight: a segment between the two points.
 * - elbow: three orthogonal segments — leave each endpoint perpendicular to
 *   its site for a stub, then one crossing segment at the adjustment
 *   fraction; degenerates to a single segment where the stubs already align.
 * - curved: a cubic whose control points extend along each outward normal. */
export function routeConnector(
  kind: ConnectorKind,
  from: ConnectorEndpoint,
  to: ConnectorEndpoint,
  adjustment = 0.5,
): RoutedPath {
  if (kind === "straight") {
    return { kind, points: [from.point, to.point] };
  }
  if (kind === "curved") {
    const dist = Math.hypot(to.point.x - from.point.x, to.point.y - from.point.y);
    const reach = Math.max(dist * 0.4, 1e-6);
    const c1 = controlAlong(from, to.point, reach);
    const c2 = controlAlong(to, from.point, reach);
    return { kind, points: [from.point, to.point], controls: [c1, c2] };
  }
  return { kind, points: elbowPoints(from, to, adjustment) };
}

function controlAlong(end: ConnectorEndpoint, other: Point, reach: number): Point {
  const n = end.normal && (end.normal.x || end.normal.y) ? end.normal : directionTo(end.point, other);
  return { x: end.point.x + n.x * reach, y: end.point.y + n.y * reach };
}

function directionTo(a: Point, b: Point): Point {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const len = Math.hypot(dx, dy) || 1;
  return { x: dx / len, y: dy / len };
}

/** Whether a normal is (near) horizontal — its dominant axis is x. */
function isHorizontal(n: Point): boolean {
  return Math.abs(n.x) >= Math.abs(n.y);
}

function elbowPoints(from: ConnectorEndpoint, to: ConnectorEndpoint, adjustment: number): Point[] {
  const a = from.point;
  const b = to.point;
  const na = from.normal && (from.normal.x || from.normal.y) ? from.normal : axisNormalToward(a, b);
  const nb = to.normal && (to.normal.x || to.normal.y) ? to.normal : axisNormalToward(b, a);
  const frac = Number.isFinite(adjustment) ? Math.min(1, Math.max(0, adjustment)) : 0.5;

  // Both endpoints leave along a horizontal axis: the crossing is a vertical
  // segment placed at the adjustment fraction between them.
  if (isHorizontal(na) && isHorizontal(nb)) {
    const midX = a.x + (b.x - a.x) * frac;
    if (Math.abs(a.y - b.y) < 1e-6) {
      return [a, b];
    }
    return [a, { x: midX, y: a.y }, { x: midX, y: b.y }, b];
  }
  // Both leave vertically: crossing is a horizontal segment.
  if (!isHorizontal(na) && !isHorizontal(nb)) {
    const midY = a.y + (b.y - a.y) * frac;
    if (Math.abs(a.x - b.x) < 1e-6) {
      return [a, b];
    }
    return [a, { x: a.x, y: midY }, { x: b.x, y: midY }, b];
  }
  // Mixed: one horizontal, one vertical — a single corner at their meeting.
  if (isHorizontal(na)) {
    return [a, { x: b.x, y: a.y }, b];
  }
  return [a, { x: a.x, y: b.y }, b];
}

/** The axis-aligned outward direction from `a` toward `b` (used to give a
 * free elbow/curved endpoint a sensible leave direction). */
function axisNormalToward(a: Point, b: Point): Point {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  if (Math.abs(dx) >= Math.abs(dy)) {
    return { x: Math.sign(dx) || 1, y: 0 };
  }
  return { x: 0, y: Math.sign(dy) || 1 };
}

/** Turns a routed path into an SVG `d` string in the caller's coordinate
 * unit — a poly-line for straight/elbow, a single cubic for curved. */
export function routedPathToSvg(path: RoutedPath): string {
  const pts = path.points;
  if (pts.length === 0) {
    return "";
  }
  if (path.kind === "curved" && path.controls) {
    const [c1, c2] = path.controls;
    const a = pts[0];
    const b = pts[pts.length - 1];
    return `M ${fmt(a.x)} ${fmt(a.y)} C ${fmt(c1.x)} ${fmt(c1.y)} ${fmt(c2.x)} ${fmt(c2.y)} ${fmt(b.x)} ${fmt(b.y)}`;
  }
  return `M ${pts.map((p) => `${fmt(p.x)} ${fmt(p.y)}`).join(" L ")}`;
}

function fmt(n: number): string {
  return (Math.round(n * 100) / 100).toString();
}

/** Whether an elbow path would run back through a box it is meant to avoid —
 * true if any interior vertex lands strictly inside `box`. The router already
 * leaves perpendicular to each site, so this is the guard AC-13 checks. */
export function elbowAvoidsBox(points: Point[], box: ConnectorBox): boolean {
  for (let i = 1; i < points.length - 1; i++) {
    if (pointStrictlyInBox(points[i], box)) {
      return false;
    }
  }
  return true;
}

function pointStrictlyInBox(p: Point, box: ConnectorBox): boolean {
  const local = rotatePoint({ x: p.x - box.x, y: p.y - box.y }, box.w, box.h, -(box.rotationDeg ?? 0));
  const eps = 1e-6;
  return local.x > eps && local.x < box.w - eps && local.y > eps && local.y < box.h - eps;
}
