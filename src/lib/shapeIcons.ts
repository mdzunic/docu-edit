// SVG icon paths for the "Insert -> Shape" gallery (SPEC-051 AC-1): every
// catalogue geometry drawn from the same unit-square vertex lists that paint
// it on the page (shapeGeometry.ts), so a gallery icon can never drift from
// what picking it actually draws. Pure path-string maths — no DOM and no
// canvas context — so it unit-tests like every other lib module.
import type { ShapeGeometry } from "./documentModel";
import {
  lineSegment,
  roundedCalloutOutline,
  roundRectRadius,
  scalePolygon,
  shapePolygon,
  type Point,
} from "./shapeGeometry";

/** Side of the square viewBox every icon path is authored in. */
export const SHAPE_ICON_SIZE = 24;

/** Padding kept inside the viewBox so a stroked outline never clips at the
 * edge — the geometry itself is scaled into the box this leaves. */
const ICON_INSET = 2;

const ICON_BOX = SHAPE_ICON_SIZE - ICON_INSET * 2;

/** Arrowhead length, as a fraction of the segment it caps. The canvas sizes
 * its own head against the outline width (shapeCanvas.ts's paintArrowhead),
 * which has no equivalent at icon scale — only the spread carries over. */
const ARROWHEAD_FRACTION = 0.3;
const ARROWHEAD_SPREAD = Math.PI / 7;

/** Two decimals is enough for a 24-unit box and keeps the emitted `d`
 * readable (and stable to assert on). */
function coord(value: number): string {
  return String(Math.round(value * 100) / 100);
}

/** A drawing-box point as viewBox coordinates. */
function at(p: Point): string {
  return `${coord(ICON_INSET + p.x)},${coord(ICON_INSET + p.y)}`;
}

function polygonPath(polygon: Point[]): string {
  const scaled = scalePolygon(polygon, ICON_BOX, ICON_BOX);
  return `${scaled.map((p, i) => `${i === 0 ? "M" : "L"} ${at(p)}`).join(" ")} Z`;
}

/** The ellipse inscribed in the (square) drawing box, as the two half-arcs
 * an SVG path needs — `<path>` keeps every icon a single element. */
function ellipsePath(): string {
  const radius = ICON_BOX / 2;
  const left = { x: 0, y: radius };
  const right = { x: ICON_BOX, y: radius };
  const arc = `A ${coord(radius)},${coord(radius)} 0 1 0`;
  return `M ${at(left)} ${arc} ${at(right)} ${arc} ${at(left)} Z`;
}

/** The same corner radius the canvas rounds a `"roundRect"` with. */
function roundRectPath(): string {
  const r = roundRectRadius(ICON_BOX, ICON_BOX);
  const arc = `A ${coord(r)},${coord(r)} 0 0 1`;
  return [
    `M ${at({ x: r, y: 0 })}`,
    `L ${at({ x: ICON_BOX - r, y: 0 })}`,
    `${arc} ${at({ x: ICON_BOX, y: r })}`,
    `L ${at({ x: ICON_BOX, y: ICON_BOX - r })}`,
    `${arc} ${at({ x: ICON_BOX - r, y: ICON_BOX })}`,
    `L ${at({ x: r, y: ICON_BOX })}`,
    `${arc} ${at({ x: 0, y: ICON_BOX - r })}`,
    `L ${at({ x: 0, y: r })}`,
    `${arc} ${at({ x: r, y: 0 })}`,
    "Z",
  ].join(" ");
}

/** The rounded callout, from the same outline the canvas paints: straight
 * runs plus four clockwise 90° corner arcs (sweep flag 1, as in
 * roundRectPath). */
function roundedCalloutPath(): string {
  const { start, steps, radius } = roundedCalloutOutline(ICON_BOX, ICON_BOX);
  const arc = `A ${coord(radius)},${coord(radius)} 0 0 1`;
  const legs = steps.map((step) => `${step.corner ? arc : "L"} ${at(step.to)}`);
  return `M ${at(start)} ${legs.join(" ")} Z`;
}

/** A line geometry: the corner-to-corner segment the canvas draws, plus (for
 * `"arrow"`) the filled head as a second, closed subpath. The segment
 * subpath stays open, so a `fill` on the element paints the head alone. */
function segmentPath(withHead: boolean): string {
  const [from, to] = lineSegment(ICON_BOX, ICON_BOX);
  const segment = `M ${at(from)} L ${at(to)}`;
  if (!withHead) {
    return segment;
  }
  const angle = Math.atan2(to.y - from.y, to.x - from.x);
  const size = Math.hypot(to.x - from.x, to.y - from.y) * ARROWHEAD_FRACTION;
  const wing = (sign: number): Point => ({
    x: to.x - size * Math.cos(angle + sign * ARROWHEAD_SPREAD),
    y: to.y - size * Math.sin(angle + sign * ARROWHEAD_SPREAD),
  });
  return `${segment} M ${at(to)} L ${at(wing(-1))} L ${at(wing(1))} Z`;
}

/** The `d` attribute for `geometry`'s gallery icon, in a
 * `SHAPE_ICON_SIZE` x `SHAPE_ICON_SIZE` viewBox. Line geometries come back
 * as an open segment (stroke them); everything else as a closed outline
 * (fill and stroke them) — LINE_GEOMETRIES tells the two apart. */
export function shapeIconPath(geometry: ShapeGeometry): string {
  switch (geometry) {
    case "line":
      return segmentPath(false);
    case "arrow":
      return segmentPath(true);
    case "ellipse":
      return ellipsePath();
    case "roundRect":
      return roundRectPath();
    case "wedgeRoundRectCallout":
      return roundedCalloutPath();
    default: {
      const polygon = shapePolygon(geometry);
      return polygon ? polygonPath(polygon) : "";
    }
  }
}
