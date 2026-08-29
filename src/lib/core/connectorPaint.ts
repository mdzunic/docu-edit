// SPEC-070: the *painted* description of a connector — the one place the
// line's stroke, dash and arrowheads are turned into drawable geometry, so
// the screen, the printed sheet, the PDF and the library miniature cannot
// disagree about what a connector looks like.
//
// Before this module the arrowheads were SVG `<marker>` elements owned by the
// interactive overlay: they existed only on screen, ignored the connector's
// `ArrowheadKind` entirely (every kind painted the same triangle), and could
// not be read back by the PDF extractor. Emitting real path geometry instead
// makes one description serve every surface.
//
// Product-neutral and framework-free (core layer): geometry in, paint data
// out.
import {
  arrowheadOutline,
  pointsToSvg,
  routedPathEnds,
  routedPathToSvg,
  type RoutedPath,
} from "./connectors";
import type { Point } from "./shapeGeometry";
import type { ArrowheadKind, ShapeOutline } from "./shapes";

/** The stroke a connector falls back to when it carries no outline — the
 * same dark blue `DEFAULT_CONNECTOR_OUTLINE` gives a freshly drawn one. */
export const DEFAULT_CONNECTOR_STROKE = "#2e5395";

/** Outline widths are stored in points; everything painted here is px. */
const PX_PER_PT = 96 / 72;

/** One arrowhead, ready to paint: its `d` plus whether it is filled solid
 * with the line colour (triangle/stealth/diamond/oval) or stroked like the
 * line (the "open" barb). */
export interface ConnectorHeadPaint {
  d: string;
  filled: boolean;
}

/** Everything needed to paint one connector, in the caller's own frame. */
export interface ConnectorPaint {
  /** Stable identity for the `v-for` (`"<section>:<block>"`). */
  key: string;
  /** The line's SVG `d`. */
  d: string;
  color: string;
  widthPx: number;
  /** SVG `stroke-dasharray`, or `""` for a solid line. */
  dash: string;
  /** The start head then the end head; either may be absent. */
  heads: ConnectorHeadPaint[];
}

/** A connector's painted stroke width in px — never thinner than a hairline,
 * so a 0pt outline still shows. */
export function connectorStrokeWidthPx(outline?: ShapeOutline): number {
  return Math.max(1, (outline?.widthPt || 1) * PX_PER_PT);
}

/** The dash pattern for an outline's dash style, in px, mirroring the shape
 * outline's own patterns. `""` is solid. */
export function connectorDashArray(dash?: ShapeOutline["dash"]): string {
  if (dash === "dash") {
    return "6 4";
  }
  if (dash === "dot") {
    return "1 3";
  }
  return "";
}

export interface ConnectorPaintOptions {
  outline?: ShapeOutline;
  startArrow?: ArrowheadKind;
  endArrow?: ArrowheadKind;
}

/** Turns a routed path plus its formatting into paint data. The arrowheads
 * are sized off the painted stroke width, so they scale with the line. */
export function buildConnectorPaint(
  key: string,
  path: RoutedPath,
  options: ConnectorPaintOptions,
): ConnectorPaint {
  const widthPx = connectorStrokeWidthPx(options.outline);
  const heads: ConnectorHeadPaint[] = [];
  const ends = routedPathEnds(path);
  if (ends) {
    const wanted: Array<[ArrowheadKind | undefined, { point: Point; along: Point }]> = [
      [options.startArrow, ends.start],
      [options.endArrow, ends.end],
    ];
    for (const [kind, end] of wanted) {
      const head = arrowheadOutline(kind, end.point, end.along, widthPx);
      if (head) {
        heads.push({ d: pointsToSvg(head.points, head.filled), filled: head.filled });
      }
    }
  }
  return {
    key,
    d: routedPathToSvg(path),
    color: options.outline?.color || DEFAULT_CONNECTOR_STROKE,
    widthPx,
    dash: connectorDashArray(options.outline?.dash),
    heads,
  };
}

/** The box a connector actually paints within, in the path's own unit:
 * the routed line grown by half its stroke width, plus whatever the
 * arrowheads reach. The library miniature sizes its box from this, and the
 * page painter uses it to decide which pages a connector touches — a line's
 * own points are not enough, since a head sticks out past the endpoint. */
export function connectorPaintBounds(
  path: RoutedPath,
  options: ConnectorPaintOptions,
): { left: number; top: number; right: number; bottom: number } {
  const widthPx = connectorStrokeWidthPx(options.outline);
  const points: Point[] = path.controls ? [...path.points, ...path.controls] : [...path.points];
  const ends = routedPathEnds(path);
  if (ends) {
    for (const [kind, end] of [
      [options.startArrow, ends.start],
      [options.endArrow, ends.end],
    ] as Array<[ArrowheadKind | undefined, { point: Point; along: Point }]>) {
      const head = arrowheadOutline(kind, end.point, end.along, widthPx);
      if (head) {
        points.push(...head.points);
      }
    }
  }
  if (points.length === 0) {
    return { left: 0, top: 0, right: 0, bottom: 0 };
  }
  const pad = widthPx / 2;
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const point of points) {
    left = Math.min(left, point.x);
    top = Math.min(top, point.y);
    right = Math.max(right, point.x);
    bottom = Math.max(bottom, point.y);
  }
  return { left: left - pad, top: top - pad, right: right + pad, bottom: bottom + pad };
}

/** The same routed path with every coordinate multiplied by `factor` — how a
 * path routed in cm (the model-level router) becomes one in px (a painter). */
export function scaleRoutedPath(path: RoutedPath, factor: number): RoutedPath {
  const scale = (p: Point): Point => ({ x: p.x * factor, y: p.y * factor });
  const scaled: RoutedPath = { kind: path.kind, points: path.points.map(scale) };
  if (path.controls) {
    scaled.controls = [scale(path.controls[0]), scale(path.controls[1])];
  }
  return scaled;
}

/** Typed `[]` seed for a connector-paint array prop in a `.vue` script (the
 * espree parser bars the annotation on a default — same idiom as
 * connectorOverlay's `no*` seeds). */
export function noConnectorPaint(): ConnectorPaint[] {
  return [];
}

/** Typed pass-through for a runtime-declared `.vue` prop. */
export function asConnectorPaint(value: unknown): ConnectorPaint[] {
  return value as ConnectorPaint[];
}
