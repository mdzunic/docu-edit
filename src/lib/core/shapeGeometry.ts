// Pure geometry for the Shape catalogue (SPEC-051): the same unit-square
// vertex lists paint the shape on the page's canvas layer and answer "does
// this point land on the shape" for hit testing ("Selection and hit
// testing"). Framework-free and canvas-API-free (jsdom's test environment
// has no 2D context) so it is unit-testable like every other lib module.
import { LINE_GEOMETRIES, type ShapeGeometry } from "./shapes";
import { shapeAdjustmentDefaults, shapeDefinition } from "./shapeCatalogue";

export interface Point {
  x: number;
  y: number;
}

/** A point in the shape's own (unrotated, 0..widthCm x 0..heightCm) local
 * space — the caller rotates a hit-test point into this space first. */
export type LocalPoint = Point;

const TAU = Math.PI * 2;

/** How much of a callout's box its body occupies — the bottom quarter is
 * where the tail hangs. Kept here (mirrored from shapeCatalogue.ts's own
 * copy) because roundedCalloutOutline/Radius below are pinned by SPEC-051
 * tests and predate the SPEC-069 table. */
const CALLOUT_BODY_FRACTION = 0.75;

/** Where the rounded callout's tail leaves the body's bottom edge, as
 * fractions of the width — the rounded corner needs the tail set in so it
 * has room to curve (see roundedCalloutRadius). */
const ROUND_CALLOUT_TAIL = { left: 0.25, tip: 0.35, right: 0.55 };

/** The closed polygon for a catalogue geometry, in unit-square fractions
 * (0..1 of width, 0..1 of height) — multiply by widthCm/heightCm (or a pixel
 * equivalent) to get real coordinates. Derived from the single
 * `SHAPE_DEFINITIONS` table (SPEC-069 AC-15a): the geometry's own outline at
 * its default adjustments. Returns null for `"ellipse"`, which has no
 * polygon representation (see pointInEllipse below), and for the two line
 * geometries, which are a segment rather than a closed region (see
 * lineSegment below). */
export function shapePolygon(geometry: ShapeGeometry): Point[] | null {
  if (geometry === "ellipse") {
    return null;
  }
  const def = shapeDefinition(geometry);
  if (!def || def.isLine) {
    return null;
  }
  return def.outline(shapeAdjustmentDefaults(def));
}

/** The rounded-rectangle corner radius, in the same units as `width`/`height`
 * — used both to paint and (loosely) to hit-test `"roundRect"`/
 * `"wedgeRoundRectCallout"`. Word's own default adjustment is ~16.7% of the
 * shorter side. */
export function roundRectRadius(width: number, height: number): number {
  return Math.min(width, height) * 0.1667;
}

/** `"wedgeRoundRectCallout"`'s corner radius, in the same units as
 * `width`/`height`: roundRect's own 16.7% rule, but measured against the
 * BODY rather than the whole box (the tail's quarter is not part of the
 * rectangle being rounded), and never wider than the run of bottom edge
 * either side of the tail — a corner arc and the tail would otherwise
 * cross. */
export function roundedCalloutRadius(width: number, height: number): number {
  const body = height * CALLOUT_BODY_FRACTION;
  return Math.max(
    0,
    Math.min(
      roundRectRadius(width, body),
      ROUND_CALLOUT_TAIL.left * width,
      (1 - ROUND_CALLOUT_TAIL.right) * width,
      width / 2,
      body / 2,
    ),
  );
}

/** One leg of a rounded callout's outline: a straight run to `to`, or — when
 * `corner` is set — the 90° arc that turns at `corner` and rejoins the
 * outline at `to`. `to` always sits exactly `roundedCalloutRadius` along
 * the next edge, so a renderer needs no tangent maths of its own. */
export interface CalloutOutlineStep {
  to: Point;
  corner?: Point;
}

/** `"wedgeRoundRectCallout"`'s outline in real units, walked clockwise from
 * the top edge: a rounded-rectangle body whose bottom edge detours into the
 * (sharp) tail. Shared so the page canvas and the gallery icon draw one
 * outline in their own two idioms — `shapePolygon` stays the straight-edged
 * approximation hit testing works against, as it does for `"roundRect"`. */
export function roundedCalloutOutline(
  width: number,
  height: number,
): { start: Point; steps: CalloutOutlineStep[]; radius: number } {
  const radius = roundedCalloutRadius(width, height);
  const bottom = height * CALLOUT_BODY_FRACTION;
  const start = { x: radius, y: 0 };
  return {
    start,
    radius,
    steps: [
      { to: { x: width - radius, y: 0 } },
      { corner: { x: width, y: 0 }, to: { x: width, y: radius } },
      { to: { x: width, y: bottom - radius } },
      { corner: { x: width, y: bottom }, to: { x: width - radius, y: bottom } },
      { to: { x: ROUND_CALLOUT_TAIL.right * width, y: bottom } },
      { to: { x: ROUND_CALLOUT_TAIL.tip * width, y: height } },
      { to: { x: ROUND_CALLOUT_TAIL.left * width, y: bottom } },
      { to: { x: radius, y: bottom } },
      { corner: { x: 0, y: bottom }, to: { x: 0, y: bottom - radius } },
      { to: { x: 0, y: radius } },
      { corner: { x: 0, y: 0 }, to: start },
    ],
  };
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

/** The axis-aligned box a `width` x `height` shape needs once rotated about
 * its own centre, plus the padding that box adds either side.
 *
 * A rotated shape is painted into a canvas the size of its UNROTATED box
 * unless something asks this: the corners then swing outside that box and
 * are sliced off, turning a rotated rectangle into an octagon. The shape
 * keeps its own form and dimensions — only the surface it is painted on
 * has to grow, and shift back by the padding to leave the shape where it
 * was. Zero rotation pads nothing, so an unrotated shape is untouched. */
export function rotatedBounds(
  width: number,
  height: number,
  rotationDeg: number,
): { width: number; height: number; padX: number; padY: number } {
  if (!rotationDeg) {
    return { width, height, padX: 0, padY: 0 };
  }
  const rad = (rotationDeg * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad));
  const sin = Math.abs(Math.sin(rad));
  const rotatedWidth = width * cos + height * sin;
  const rotatedHeight = width * sin + height * cos;
  return {
    width: rotatedWidth,
    height: rotatedHeight,
    padX: (rotatedWidth - width) / 2,
    padY: (rotatedHeight - height) / 2,
  };
}

/** Rotates `point` by `rotationDeg` around the box centre `(width/2,
 * height/2)` — the same rotation painting applies, so a corner of the
 * unrotated box maps to where that corner is actually painted. What the
 * selection handles are placed with (AC-10, AC-16): a handle belongs on the
 * shape's own corner, not on the corner of the box the shape would occupy
 * if it were upright. */
export function rotatePoint(point: Point, width: number, height: number, rotationDeg: number): Point {
  if (!rotationDeg) {
    return point;
  }
  const cx = width / 2;
  const cy = height / 2;
  const rad = (rotationDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = point.x - cx;
  const dy = point.y - cy;
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}

/** Rotates `point` by `-rotationDeg` around the box centre `(width/2,
 * height/2)` — the inverse of painting's rotation, so a hit test can work
 * entirely in the shape's own unrotated local space. */
export function unrotatePoint(point: Point, width: number, height: number, rotationDeg: number): Point {
  return rotatePoint(point, width, height, -rotationDeg);
}

/** The eight resize-handle directions, clockwise from screen-north — the
 * order is the map from "how far round the compass" to a CSS cursor name. */
const RESIZE_DIRECTIONS = ["n", "ne", "e", "se", "s", "sw", "w", "nw"];

/** The `*-resize` cursor a handle shows once its shape is rotated: the "nw"
 * handle of a shape turned 90° sits on the shape's north-EAST on screen, and
 * a cursor pointing the other way tells the user the drag will do the
 * opposite of what it does. The eight directions are 45° apart, so the
 * handle's own screen angle rounds to the nearest of them. */
export function rotatedResizeCursor(handle: string, rotationDeg: number): string {
  const base = RESIZE_DIRECTIONS.indexOf(handle);
  if (base < 0) {
    return `${handle}-resize`;
  }
  const octant = Math.round(base + rotationDeg / 45);
  return `${RESIZE_DIRECTIONS[((octant % 8) + 8) % 8]}-resize`;
}

/** Rotates a screen-space drag delta into the shape's own unrotated axes,
 * so dragging a handle that now sits on a rotated edge sizes the shape
 * along that edge rather than along the screen's x/y (AC-4 under AC-13's
 * rotation). A pure delta has no origin to rotate about, hence the
 * zero-sized box. */
export function unrotateDelta(dx: number, dy: number, rotationDeg: number): { dx: number; dy: number } {
  const rotated = unrotatePoint({ x: dx, y: dy }, 0, 0, rotationDeg);
  return { dx: rotated.x, dy: rotated.y };
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
