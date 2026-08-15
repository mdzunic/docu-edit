// SVG icon marks for the "Insert -> Chart" gallery (SPEC-052 AC-2), the
// chart counterpart to shapeIcons.ts. A chart's appearance depends on its
// data, so unlike a shape there is no single outline to reuse — what the
// icons borrow from the painter instead is its *grammar*: the axis pair it
// draws (chartHasAxes), how much of a category slot a bar fills
// (BAR_CATEGORY_FILL), how big a doughnut's hole is (DOUGHNUT_HOLE_RATIO),
// and the fact that line/lineMarkers/area are one series drawn three ways.
// Pure path-string maths — no DOM, no canvas — so it unit-tests like every
// other lib module.
import type { ChartType } from "./charts";
import { BAR_CATEGORY_FILL, chartHasAxes, DOUGHNUT_HOLE_RATIO } from "./chartCanvas";

/** Side of the square viewBox every icon is authored in. */
export const CHART_ICON_SIZE = 24;

/** Padding kept inside the viewBox so a stroked mark never clips at the
 * edge — the plot itself is the box this leaves. */
const ICON_INSET = 2;

const PLOT = CHART_ICON_SIZE - ICON_INSET * 2;

/** How a mark is painted. `axis` is the plot frame, `fill` a filled and
 * outlined mark (bar, slice, area band), `line` a bare series line, `dot` a
 * solid marker or point. */
export type ChartIconRole = "axis" | "fill" | "line" | "dot";

export interface ChartIconMark {
  role: ChartIconRole;
  d: string;
}

/** The sample series every cartesian icon plots, as fractions of the plot
 * height — one shape drawn three ways for line/lineMarkers/area, exactly as
 * paintLineOrArea does from one point list. */
const SAMPLE_LINE = [0.35, 0.75, 0.45, 0.95];

/** The sample categories the clustered icons plot, and the two-segment
 * stacks the stacked one does. */
const SAMPLE_BARS = [0.5, 0.85, 0.35];
const SAMPLE_STACKS = [
  [0.3, 0.25],
  [0.45, 0.3],
  [0.25, 0.2],
];

const SAMPLE_SCATTER = [
  [0.15, 0.25],
  [0.35, 0.55],
  [0.55, 0.4],
  [0.75, 0.75],
  [0.9, 0.6],
];

/** Marker/point radius. The painter's own 2.5px is an absolute size on a
 * full-size chart, with no meaning at icon scale — only "a dot at each
 * plotted point" carries over. */
const DOT_RADIUS = 1.7;

function coord(value: number): string {
  return String(Math.round(value * 100) / 100);
}

/** A plot-space point (x right, y down from the plot's top-left) as viewBox
 * coordinates. */
function at(x: number, y: number): string {
  return `${coord(ICON_INSET + x)},${coord(ICON_INSET + y)}`;
}

/** Plot y for a value given as a fraction of the plot height. */
function valueY(fraction: number): number {
  return PLOT - fraction * PLOT;
}

function rect(x: number, y: number, width: number, height: number): string {
  return `M ${at(x, y)} L ${at(x + width, y)} L ${at(x + width, y + height)} L ${at(x, y + height)} Z`;
}

function circle(cx: number, cy: number, r: number): string {
  const arc = `A ${coord(r)},${coord(r)} 0 1 0`;
  return `M ${at(cx - r, cy)} ${arc} ${at(cx + r, cy)} ${arc} ${at(cx - r, cy)} Z`;
}

/** The L-shaped category/value axis pair, drawn exactly where paintAxes
 * draws it: down the plot's left edge, then along its bottom. */
function axisMark(): ChartIconMark {
  return { role: "axis", d: `M ${at(0, 0)} L ${at(0, PLOT)} L ${at(PLOT, PLOT)}` };
}

/** The x of each category's bar band, and its thickness — BAR_CATEGORY_FILL
 * of the slot, centred in it, as paintBarLike lays them out. */
function bands(count: number): { start: number; thickness: number }[] {
  const slot = PLOT / count;
  const thickness = slot * BAR_CATEGORY_FILL;
  return Array.from({ length: count }, (_, i) => ({
    start: i * slot + (slot - thickness) / 2,
    thickness,
  }));
}

/** The sample series' plotted points, evenly spread across the plot. */
function linePoints(): { x: number; y: number }[] {
  return SAMPLE_LINE.map((value, i) => ({
    x: (i / (SAMPLE_LINE.length - 1)) * PLOT,
    y: valueY(value),
  }));
}

function polyline(points: { x: number; y: number }[]): string {
  return points.map((p, i) => `${i === 0 ? "M" : "L"} ${at(p.x, p.y)}`).join(" ");
}

function columnMarks(): ChartIconMark[] {
  return bands(SAMPLE_BARS.length).map((band, i) => ({
    role: "fill" as const,
    d: rect(band.start, valueY(SAMPLE_BARS[i]), band.thickness, SAMPLE_BARS[i] * PLOT),
  }));
}

function stackedMarks(): ChartIconMark[] {
  const marks: ChartIconMark[] = [];
  bands(SAMPLE_STACKS.length).forEach((band, i) => {
    let top = PLOT;
    for (const segment of SAMPLE_STACKS[i]) {
      const height = segment * PLOT;
      top -= height;
      marks.push({ role: "fill", d: rect(band.start, top, band.thickness, height) });
    }
  });
  return marks;
}

function barMarks(): ChartIconMark[] {
  return bands(SAMPLE_BARS.length).map((band, i) => ({
    role: "fill" as const,
    d: rect(0, band.start, SAMPLE_BARS[i] * PLOT, band.thickness),
  }));
}

/** Pie/doughnut share a slice boundary: one wedge cut clockwise from twelve
 * o'clock, where paintPieLike starts its first slice. */
const SLICE_SWEEP_DEG = 110;

function radial(cx: number, cy: number, r: number, degFromTop: number): { x: number; y: number } {
  const rad = ((degFromTop - 90) * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

function pieMarks(): ChartIconMark[] {
  const centre = PLOT / 2;
  const r = PLOT / 2;
  const edge = radial(centre, centre, r, SLICE_SWEEP_DEG);
  return [
    { role: "fill", d: circle(centre, centre, r) },
    {
      role: "fill",
      d: `M ${at(centre, centre)} L ${at(centre, centre - r)} A ${coord(r)},${coord(r)} 0 0 1 ${at(edge.x, edge.y)} Z`,
    },
  ];
}

function doughnutMarks(): ChartIconMark[] {
  const centre = PLOT / 2;
  const r = PLOT / 2;
  const hole = r * DOUGHNUT_HOLE_RATIO;
  // Ring first (two subpaths, evenodd punches the hole), then the same
  // slice boundary the pie shows, clipped to the ring.
  const ring = `${circle(centre, centre, r)} ${circle(centre, centre, hole)}`;
  const boundary = (deg: number) => {
    const inner = radial(centre, centre, hole, deg);
    const outer = radial(centre, centre, r, deg);
    return `M ${at(inner.x, inner.y)} L ${at(outer.x, outer.y)}`;
  };
  return [
    { role: "fill", d: ring },
    { role: "line", d: boundary(0) },
    { role: "line", d: boundary(SLICE_SWEEP_DEG) },
  ];
}

function scatterMarks(): ChartIconMark[] {
  return SAMPLE_SCATTER.map(([x, y]) => ({
    role: "dot" as const,
    d: circle(x * PLOT, valueY(y), DOT_RADIUS),
  }));
}

function seriesMarks(chartType: ChartType): ChartIconMark[] {
  const points = linePoints();
  switch (chartType) {
    case "columnClustered":
      return columnMarks();
    case "columnStacked":
      return stackedMarks();
    case "barClustered":
      return barMarks();
    case "line":
      return [{ role: "line", d: polyline(points) }];
    case "lineMarkers":
      return [
        { role: "line", d: polyline(points) },
        ...points.map((p) => ({ role: "dot" as const, d: circle(p.x, p.y, DOT_RADIUS) })),
      ];
    case "area":
      // The line's own points, closed down to the baseline either end —
      // paintLineOrArea fills between exactly those and the zero line.
      return [
        {
          role: "fill",
          d: [
            `M ${at(points[0].x, PLOT)}`,
            ...points.map((p) => `L ${at(p.x, p.y)}`),
            `L ${at(points[points.length - 1].x, PLOT)}`,
            "Z",
          ].join(" "),
        },
      ];
    case "scatter":
      return scatterMarks();
    case "pie":
      return pieMarks();
    case "doughnut":
      return doughnutMarks();
    default:
      return [];
  }
}

/** The marks making up `chartType`'s gallery icon, back to front, in a
 * `CHART_ICON_SIZE` x `CHART_ICON_SIZE` viewBox. The axis comes first for
 * the types the painter gives one, so the series draws over it. */
export function chartIconMarks(chartType: ChartType): ChartIconMark[] {
  const series = seriesMarks(chartType);
  return chartHasAxes(chartType) ? [axisMark(), ...series] : series;
}
