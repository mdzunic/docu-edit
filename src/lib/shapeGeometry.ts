// Pure geometry for the Shape catalogue (SPEC-051): the same unit-square
// vertex lists paint the shape on the page's canvas layer and answer "does
// this point land on the shape" for hit testing ("Selection and hit
// testing"). Framework-free and canvas-API-free (jsdom's test environment
// has no 2D context) so it is unit-testable like every other lib module.
import { LINE_GEOMETRIES, type ShapeGeometry } from "./documentModel";

export interface Point {
  x: number;
  y: number;
}

/** A point in the shape's own (unrotated, 0..widthCm x 0..heightCm) local
 * space — the caller rotates a hit-test point into this space first. */
export type LocalPoint = Point;

const TAU = Math.PI * 2;

function star5Points(): Point[] {
  // Outer point on top, five points, standard prstGeom star5 proportions:
  // inner radius ~0.382x the outer (golden-ratio-ish, matches Word's default
  // adjustment) measured against a bounding circle inscribed in the box.
  const points: Point[] = [];
  const cx = 0.5;
  const cy = 0.5;
  const outerR = 0.5;
  const innerR = outerR * 0.382;
  for (let i = 0; i < 10; i++) {
    const r = i % 2 === 0 ? outerR : innerR;
    const angle = -Math.PI / 2 + (i * TAU) / 10;
    points.push({ x: cx + r * Math.cos(angle), y: cy + r * Math.sin(angle) });
  }
  return points;
}

/** The closed polygon for a catalogue geometry, in unit-square fractions
 * (0..1 of width, 0..1 of height) — multiply by widthCm/heightCm (or a pixel
 * equivalent) to get real coordinates. Returns null for `"ellipse"`, which
 * has no polygon representation (see ellipseHit below), and for the two
 * line geometries, which are a segment rather than a closed region (see
 * lineSegment below). */
export function shapePolygon(geometry: ShapeGeometry): Point[] | null {
  switch (geometry) {
    case "ellipse":
    case "line":
    case "arrow":
      return null;
    case "rect":
    case "roundRect":
      return [
        { x: 0, y: 0 },
        { x: 1, y: 0 },
        { x: 1, y: 1 },
        { x: 0, y: 1 },
      ];
    case "triangle":
      return [
        { x: 0.5, y: 0 },
        { x: 1, y: 1 },
        { x: 0, y: 1 },
      ];
    case "diamond":
      return [
        { x: 0.5, y: 0 },
        { x: 1, y: 0.5 },
        { x: 0.5, y: 1 },
        { x: 0, y: 0.5 },
      ];
    case "star5":
      return star5Points();
    case "rightArrow":
      return [
        { x: 0, y: 0.3 },
        { x: 0.6, y: 0.3 },
        { x: 0.6, y: 0 },
        { x: 1, y: 0.5 },
        { x: 0.6, y: 1 },
        { x: 0.6, y: 0.7 },
        { x: 0, y: 0.7 },
      ];
    case "leftArrow":
      return [
        { x: 1, y: 0.3 },
        { x: 0.4, y: 0.3 },
        { x: 0.4, y: 0 },
        { x: 0, y: 0.5 },
        { x: 0.4, y: 1 },
        { x: 0.4, y: 0.7 },
        { x: 1, y: 0.7 },
      ];
    case "upArrow":
      return [
        { x: 0.3, y: 1 },
        { x: 0.3, y: 0.4 },
        { x: 0, y: 0.4 },
        { x: 0.5, y: 0 },
        { x: 1, y: 0.4 },
        { x: 0.7, y: 0.4 },
        { x: 0.7, y: 1 },
      ];
    case "downArrow":
      return [
        { x: 0.3, y: 0 },
        { x: 0.7, y: 0 },
        { x: 0.7, y: 0.6 },
        { x: 1, y: 0.6 },
        { x: 0.5, y: 1 },
        { x: 0, y: 0.6 },
        { x: 0.3, y: 0.6 },
      ];
    case "wedgeRectCallout":
      return [
        { x: 0, y: 0 },
        { x: 1, y: 0 },
        { x: 1, y: 0.75 },
        { x: 0.35, y: 0.75 },
        { x: 0.15, y: 1 },
        { x: 0.05, y: 0.75 },
        { x: 0, y: 0.75 },
      ];
    case "wedgeRoundRectCallout":
      return [
        { x: 0, y: 0 },
        { x: 1, y: 0 },
        { x: 1, y: 0.75 },
        { x: 0.35, y: 0.75 },
        { x: 0.15, y: 1 },
        { x: 0.05, y: 0.75 },
        { x: 0, y: 0.75 },
      ];
    default:
      return null;
  }
}

/** The rounded-rectangle corner radius, in the same units as `width`/`height`
 * — used both to paint and (loosely) to hit-test `"roundRect"`/
 * `"wedgeRoundRectCallout"`. Word's own default adjustment is ~16.7% of the
 * shorter side. */
export function roundRectRadius(width: number, height: number): number {
  return Math.min(width, height) * 0.1667;
}

/** The two endpoints of a `"line"`/`"arrow"` geometry's segment, in real
 * units — corner to corner of the bounding box (no flip/direction field
 * exists on `Shape`, so the segment always runs top-left to bottom-right;
 * a horizontal segment is simply a box with `heightCm: 0`). */
export function lineSegment(width: number, height: number): [Point, Point] {
  return [
    { x: 0, y: 0 },
    { x: width, y: height },
  ];
}

/** A polygon scaled from unit-square fractions to real width/height units. */
export function scalePolygon(polygon: Point[], width: number, height: number): Point[] {
  return polygon.map((p) => ({ x: p.x * width, y: p.y * height }));
}

/** Rotates `point` by `-rotationDeg` around the box centre `(width/2,
 * height/2)` — the inverse of painting's rotation, so a hit test can work
 * entirely in the shape's own unrotated local space. */
export function unrotatePoint(point: Point, width: number, height: number, rotationDeg: number): Point {
  if (!rotationDeg) {
    return point;
  }
  const cx = width / 2;
  const cy = height / 2;
  const rad = (-rotationDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = point.x - cx;
  const dy = point.y - cy;
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}

/** Standard even-odd point-in-polygon test (ray casting). */
export function pointInPolygon(point: Point, polygon: Point[]): boolean {
  let inside = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const pi = polygon[i];
    const pj = polygon[j];
    const intersects =
      pi.y > point.y !== pj.y > point.y &&
      point.x < ((pj.x - pi.x) * (point.y - pi.y)) / (pj.y - pi.y) + pi.x;
    if (intersects) {
      inside = !inside;
    }
  }
  return inside;
}

/** Whether `point` lands inside the ellipse inscribed in `width` x `height`. */
export function pointInEllipse(point: Point, width: number, height: number): boolean {
  const rx = width / 2;
  const ry = height / 2;
  if (rx <= 0 || ry <= 0) {
    return false;
  }
  const dx = point.x - rx;
  const dy = point.y - ry;
  return (dx * dx) / (rx * rx) + (dy * dy) / (ry * ry) <= 1;
}

/** Shortest distance from `point` to the segment `[a, b]`. */
export function distanceToSegment(point: Point, a: Point, b: Point): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const lengthSq = dx * dx + dy * dy;
  if (lengthSq === 0) {
    return Math.hypot(point.x - a.x, point.y - a.y);
  }
  const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / lengthSq));
  const projX = a.x + t * dx;
  const projY = a.y + t * dy;
  return Math.hypot(point.x - projX, point.y - projY);
}

/** Distance from `point` to a polygon's outline (min over its edges). */
export function distanceToPolygonOutline(point: Point, polygon: Point[]): number {
  let min = Infinity;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    min = Math.min(min, distanceToSegment(point, polygon[j], polygon[i]));
  }
  return min;
}

export interface ShapeHitOptions {
  geometry: ShapeGeometry;
  widthCm: number;
  heightCm: number;
  rotationDeg?: number;
  /** Whether the shape has a fill — an unfilled shape is only hit through
   * its outline, widened by `outlineToleranceCm` ("Selection and hit
   * testing"). */
  filled: boolean;
  /** Whether the shape has a text frame with text — an unfilled text box is
   * also hit through its text block's bounds; approximated here as the
   * shape's own box, since the text frame never exceeds it (AC-19). */
  hasText: boolean;
  /** 3pt (~0.106cm) hit tolerance for an unfilled shape's outline, in the
   * same units as widthCm/heightCm. */
  outlineToleranceCm: number;
}

/** "Selection and hit testing": whether a point (in the same units as
 * widthCm/heightCm, relative to the shape's top-left corner before
 * rotation) lands on the shape's *painted* region — never its bounding box. */
export function pointHitsShape(point: Point, opts: ShapeHitOptions): boolean {
  const local = unrotatePoint(point, opts.widthCm, opts.heightCm, opts.rotationDeg ?? 0);

  if (LINE_GEOMETRIES.has(opts.geometry)) {
    const [a, b] = lineSegment(opts.widthCm, opts.heightCm);
    return distanceToSegment(local, a, b) <= opts.outlineToleranceCm;
  }

  if (opts.geometry === "ellipse") {
    if (opts.filled && pointInEllipse(local, opts.widthCm, opts.heightCm)) {
      return true;
    }
    if (!opts.filled) {
      // Distance-to-ellipse-outline approximated by the difference between
      // the point's ellipse "radius fraction" and 1 — accurate enough for a
      // hit-test tolerance band, not for painting.
      const rx = opts.widthCm / 2;
      const ry = opts.heightCm / 2;
      if (rx <= 0 || ry <= 0) {
        return false;
      }
      const dx = local.x - rx;
      const dy = local.y - ry;
      const frac = Math.sqrt((dx * dx) / (rx * rx) + (dy * dy) / (ry * ry));
      const bandCm = opts.outlineToleranceCm / Math.max(rx, ry);
      if (Math.abs(frac - 1) <= bandCm) {
        return true;
      }
    }
    return opts.hasText && pointInEllipse(local, opts.widthCm, opts.heightCm);
  }

  const unit = shapePolygon(opts.geometry);
  if (!unit) {
    return false;
  }
  const polygon = scalePolygon(unit, opts.widthCm, opts.heightCm);
  if (opts.filled && pointInPolygon(local, polygon)) {
    return true;
  }
  if (!opts.filled && distanceToPolygonOutline(local, polygon) <= opts.outlineToleranceCm) {
    return true;
  }
  if (opts.hasText && pointInPolygon(local, polygon)) {
    return true;
  }
  return false;
}

/** 3pt in cm — the fixed hit-tolerance "Selection and hit testing" defines
 * for an unfilled shape's outline. 1pt = 1/72in = 2.54/72cm. */
export const HIT_TOLERANCE_CM = (3 * 2.54) / 72;

/** A CSS `clip-path` value restricting a filled shape's DOM proxy to its
 * painted region ("Selection and hit testing" — a click must land on the
 * topmost *painted* region, never a bounding box). Returns null for an
 * unfilled shape (the proxy stays a full rectangle, since an unfilled
 * shape's own hit region is its outline/text — a coarser but harmless
 * simplification, see the implement summary) and for `"rect"`, whose clip
 * is the box itself. */
export function shapeClipPath(geometry: ShapeGeometry, filled: boolean): string | null {
  if (!filled || geometry === "rect" || LINE_GEOMETRIES.has(geometry)) {
    return null;
  }
  if (geometry === "ellipse") {
    return "ellipse(50% 50% at 50% 50%)";
  }
  const unit = shapePolygon(geometry);
  if (!unit) {
    return null;
  }
  const points = unit.map((p) => `${(p.x * 100).toFixed(2)}% ${(p.y * 100).toFixed(2)}%`).join(", ");
  return `polygon(${points})`;
}
