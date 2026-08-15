// SPEC-052 AC-2: the "Insert -> Chart" gallery offers the closed chart
// catalogue. Now that a cell shows a sketch of the chart rather than its
// name, "the right chart" is a property of the emitted marks — these tests
// pin the parts the painter and the icon genuinely share, and pin that no
// two catalogue types draw the same picture (the trap the two callouts fell
// into in SPEC-051's own gallery).
import { describe, expect, it } from "vitest";
import { CHART_CATALOGUE } from "./charts";
import { BAR_CATEGORY_FILL, chartHasAxes, DOUGHNUT_HOLE_RATIO } from "./chartCanvas";
import { CHART_ICON_SIZE, chartIconMarks } from "./chartIcons";

const TYPES = CHART_CATALOGUE.map((entry) => entry.chartType);

/** Every coordinate pair in a mark's `d`, with any arc radii/flags stripped
 * first — those are written like a point but are not one. */
function markPoints(d: string): [number, number][] {
  const outline = d.replace(/A [\d.]+,[\d.]+ [01] [01] [01] /g, "");
  return [...outline.matchAll(/(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/g)].map(
    (match) => [Number(match[1]), Number(match[2])] as [number, number],
  );
}

function pointsOf(chartType: (typeof TYPES)[number]): [number, number][] {
  return chartIconMarks(chartType).flatMap((mark) => markPoints(mark.d));
}

describe("chartIconMarks", () => {
  it("draws every catalogue type", () => {
    for (const chartType of TYPES) {
      const marks = chartIconMarks(chartType);
      expect(marks.length, chartType).toBeGreaterThan(0);
      for (const mark of marks) {
        expect(mark.d, chartType).toMatch(/^M /);
      }
    }
  });

  it("keeps every icon inside the viewBox", () => {
    for (const chartType of TYPES) {
      for (const [x, y] of pointsOf(chartType)) {
        expect(x, `${chartType} x`).toBeGreaterThanOrEqual(0);
        expect(x, `${chartType} x`).toBeLessThanOrEqual(CHART_ICON_SIZE);
        expect(y, `${chartType} y`).toBeGreaterThanOrEqual(0);
        expect(y, `${chartType} y`).toBeLessThanOrEqual(CHART_ICON_SIZE);
      }
    }
  });

  it("gives an axis to exactly the types the painter gives one", () => {
    for (const chartType of TYPES) {
      const hasAxis = chartIconMarks(chartType).some((mark) => mark.role === "axis");
      expect(hasAxis, chartType).toBe(chartHasAxes(chartType));
    }
  });

  it("draws the axis as the painter's own L: down the left edge, then along the bottom", () => {
    const axis = chartIconMarks("columnClustered").find((mark) => mark.role === "axis")!;
    const [topLeft, bottomLeft, bottomRight] = markPoints(axis.d);
    expect(topLeft[0]).toBe(bottomLeft[0]);
    expect(bottomLeft[1]).toBe(bottomRight[1]);
    expect(bottomLeft[1]).toBeGreaterThan(topLeft[1]);
    expect(bottomRight[0]).toBeGreaterThan(bottomLeft[0]);
  });

  it("spaces bars at the painter's own category duty cycle", () => {
    for (const chartType of ["columnClustered", "barClustered"] as const) {
      const bars = chartIconMarks(chartType).filter((mark) => mark.role === "fill");
      expect(bars, chartType).toHaveLength(3);
      // Thickness across the category axis: width for columns, height for
      // the horizontal bars.
      const across = chartType === "columnClustered" ? 0 : 1;
      const spans = bars.map((bar) => {
        const values = markPoints(bar.d).map((point) => point[across]);
        return Math.max(...values) - Math.min(...values);
      });
      const starts = bars
        .map((bar) => Math.min(...markPoints(bar.d).map((point) => point[across])))
        .sort((a, b) => a - b);
      const slot = starts[1] - starts[0];
      for (const span of spans) {
        // Loose to a tenth: path coordinates are emitted rounded to two
        // decimals, so the slot measured back off them is not exact.
        expect(span, chartType).toBeCloseTo(slot * BAR_CATEGORY_FILL, 1);
      }
    }
  });

  it("stacks the stacked columns into segments that share their column", () => {
    const clustered = chartIconMarks("columnClustered").filter((mark) => mark.role === "fill");
    const stacked = chartIconMarks("columnStacked").filter((mark) => mark.role === "fill");
    // Two segments per category, against the clustered icon's one bar each.
    expect(stacked).toHaveLength(clustered.length * 2);
    const columnX = (mark: { d: string }) => Math.min(...markPoints(mark.d).map((point) => point[0]));
    expect(new Set(stacked.map(columnX)).size).toBe(clustered.length);
    // Segments meet: within a column, one segment's top is the next's bottom.
    const first = stacked.filter((mark) => columnX(mark) === columnX(stacked[0]));
    const ys = first.flatMap((mark) => markPoints(mark.d).map((point) => point[1]));
    expect(new Set(ys).size).toBe(3);
  });

  it("punches the doughnut's hole at the painter's own ratio", () => {
    const ring = chartIconMarks("doughnut").find((mark) => mark.role === "fill")!;
    const radii = [...ring.d.matchAll(/A ([\d.]+),/g)].map((match) => Number(match[1]));
    const outer = Math.max(...radii);
    const inner = Math.min(...radii);
    expect(inner).toBeCloseTo(outer * DOUGHNUT_HOLE_RATIO, 6);
    // Two subpaths, so evenodd has something to cancel against.
    expect(ring.d.match(/M /g)).toHaveLength(2);
  });

  it("draws pie and doughnut with no axis but a slice boundary each", () => {
    for (const chartType of ["pie", "doughnut"] as const) {
      const marks = chartIconMarks(chartType);
      expect(marks.some((mark) => mark.role === "axis"), chartType).toBe(false);
      expect(marks.length, chartType).toBeGreaterThan(1);
    }
  });

  it("draws line, lineMarkers and area from one and the same series", () => {
    const line = chartIconMarks("line").find((mark) => mark.role === "line")!;
    const withMarkers = chartIconMarks("lineMarkers").find((mark) => mark.role === "line")!;
    expect(withMarkers.d).toBe(line.d);

    // A marker on each plotted point, centred on it.
    const dots = chartIconMarks("lineMarkers").filter((mark) => mark.role === "dot");
    expect(dots).toHaveLength(markPoints(line.d).length);

    // The area band runs along those same points, closed to the baseline.
    const band = chartIconMarks("area").find((mark) => mark.role === "fill")!;
    for (const point of markPoints(line.d)) {
      expect(markPoints(band.d)).toContainEqual(point);
    }
    expect(band.d.endsWith("Z")).toBe(true);
  });

  it("gives scatter bare points and no connecting line", () => {
    const marks = chartIconMarks("scatter");
    expect(marks.filter((mark) => mark.role === "dot").length).toBeGreaterThan(2);
    expect(marks.some((mark) => mark.role === "line")).toBe(false);
    expect(marks.some((mark) => mark.role === "fill")).toBe(false);
  });

  it("draws no two catalogue types the same picture", () => {
    const drawn = new Map<string, string>();
    for (const chartType of TYPES) {
      const key = JSON.stringify(chartIconMarks(chartType));
      const twin = drawn.get(key);
      expect(twin, `${chartType} draws the same icon as ${twin}`).toBeUndefined();
      drawn.set(key, chartType);
    }
  });
});
