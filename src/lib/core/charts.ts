// The product-neutral half of SPEC-052's charts: the closed type catalogue,
// the data-sheet and options shapes, the default palette, the sample data a
// fresh chart starts with and the table-cells → chart-data mapping.
// Everything here is independent of the document model — a Slides product
// places the same charts on a slide instead of anchoring them to a
// paragraph. The document-specific half (anchoring, wrap, the editing
// commands) stays in lib/charts.ts and lib/documentModel.ts.

/** The closed set of chart types the editor renders, inserts and can
 * regenerate — see "Chart catalogue" in specs/052-charts.md. Everything else
 * arriving from a file is preserved per SPEC-008, not approximated into one
 * of these (mirrors ShapeGeometry's own closed-catalogue convention). */
export type ChartType =
  | "columnClustered"
  | "columnStacked"
  | "barClustered"
  | "line"
  | "lineMarkers"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter";

/** One data series in a chart's data sheet. `values` is index-aligned with
 * the chart's `categories`; `null` is a gap (an empty cell), distinct from
 * `0` — line charts break the line at a gap, bar/column/area charts draw
 * nothing there. */
export interface ChartSeries {
  name: string;
  values: (number | null)[];
  /** "#rrggbb". */
  color: string;
}

/** The chart options panel's settings (SPEC-052 AC-14..AC-20). */
export interface ChartOptions {
  title?: string;
  categoryAxisTitle?: string;
  valueAxisTitle?: string;
  legend: "none" | "top" | "right" | "bottom" | "left";
  dataLabels: boolean;
  gridlines: boolean;
  /** Absent = auto (computed from the data). */
  valueAxisMin?: number;
  valueAxisMax?: number;
}

/** Typed pass-through for a chart type coming from a `.vue` template's plain
 * string (espree bars the annotation in `.vue` scripts). */
export function asChartType(value: string): ChartType {
  return value as ChartType;
}

/** The closed catalogue chart types, in "Chart catalogue" order — the Insert
 * gallery and the docx/odf readers' type lookup both walk this list. */
export const CHART_CATALOGUE: { chartType: ChartType; label: string }[] = [
  { chartType: "columnClustered", label: "Clustered column" },
  { chartType: "columnStacked", label: "Stacked column" },
  { chartType: "barClustered", label: "Clustered bar" },
  { chartType: "line", label: "Line" },
  { chartType: "lineMarkers", label: "Line with markers" },
  { chartType: "area", label: "Area" },
  { chartType: "pie", label: "Pie" },
  { chartType: "doughnut", label: "Doughnut" },
  { chartType: "scatter", label: "Scatter (XY)" },
];

/** Chart types that render only their first series (pie/doughnut) — the data
 * sheet says so rather than silently ignoring the rest (SPEC-052 "Chart
 * catalogue"). */
export const SINGLE_SERIES_CHART_TYPES: ReadonlySet<ChartType> = new Set(["pie", "doughnut"]);

/** A default palette for newly created/inserted series — this spec's own,
 * deliberately not inherited from the document theme (SPEC-052 "Out of
 * scope"). */
export const DEFAULT_CHART_COLORS: readonly string[] = [
  "#4e79a7",
  "#f28e2b",
  "#e15759",
  "#76b7b2",
  "#59a14f",
  "#edc948",
  "#b07aa1",
  "#ff9da7",
  "#9c755f",
  "#bab0ac",
];

/** What painting a chart needs — the visual subset of the document model's
 * `Chart` block (which extends this with anchoring/wrap/preservation
 * bookkeeping). Any host (page canvas, slide canvas, gallery icon) can
 * paint one of these. */
export interface ChartVisual {
  chartType: ChartType;
  categories: string[];
  series: ChartSeries[];
  options: ChartOptions;
  /** Present when the source chart's type was outside the catalogue
   * (SPEC-008 preservation): the chart paints as a dashed placeholder at
   * its true bounds instead of `chartType`. */
  unsupportedChartXml?: string;
}

// --- insertion defaults and data mapping (SPEC-052) --------------------------

/** "Insert → Chart" size (SPEC-052 AC-3): 12cm x 7.5cm. */
export const NEW_CHART_WIDTH_CM = 12;
export const NEW_CHART_HEIGHT_CM = 7.5;

/** The smallest display size a resize may leave a chart. */
export const MIN_CHART_SIZE_CM = 1;

/** The sample data a freshly inserted chart carries (AC-3): three
 * categories, two series — enough to show every catalogue type's shape
 * (grouped bars/lines/slices) without the data sheet opening to an
 * intimidating wall of numbers. */
export function sampleChartCategories(): string[] {
  return ["Category 1", "Category 2", "Category 3"];
}
export function sampleChartSeries(): ChartSeries[] {
  return [
    { name: "Series 1", values: [4.3, 2.5, 3.5], color: DEFAULT_CHART_COLORS[0] },
    { name: "Series 2", values: [2.4, 4.4, 1.8], color: DEFAULT_CHART_COLORS[1] },
  ];
}

export function defaultChartOptions(): ChartOptions {
  return { legend: "right", dataLabels: false, gridlines: true };
}

/** "Chart from table" (AC-21, AC-22): builds a chart from a table's cells.
 * `unparsable` collects the text of every cell (beyond the header row/
 * category column) that could not be read as a number — reported by the
 * caller rather than silently dropped or coerced to 0. */
export interface ChartFromTableResult {
  categories: string[];
  series: ChartSeries[];
  unparsable: string[];
}

/** Reads a table's cells into chart data: the first row (from its second
 * cell on) gives series names, the first column (from its second row on)
 * gives categories, every other cell is a value. A one-time copy — this
 * module keeps no reference to the table afterwards (SPEC-052 "Out of
 * scope": no live link). */
export function chartDataFromTableCells(cellText: string[][]): ChartFromTableResult {
  const unparsable: string[] = [];
  if (cellText.length < 2 || (cellText[0]?.length ?? 0) < 2) {
    return { categories: [], series: [], unparsable };
  }
  const header = cellText[0];
  const seriesCount = header.length - 1;
  const series: ChartSeries[] = Array.from({ length: seriesCount }, (_, i) => ({
    name: header[i + 1]?.trim() || `Series ${i + 1}`,
    values: [],
    color: DEFAULT_CHART_COLORS[i % DEFAULT_CHART_COLORS.length],
  }));
  const categories: string[] = [];
  for (let r = 1; r < cellText.length; r++) {
    const row = cellText[r];
    categories.push(row[0]?.trim() ?? "");
    for (let s = 0; s < seriesCount; s++) {
      const raw = (row[s + 1] ?? "").trim();
      if (raw === "") {
        series[s].values.push(null);
        continue;
      }
      const n = Number(raw);
      if (Number.isFinite(n)) {
        series[s].values.push(n);
      } else {
        series[s].values.push(null);
        unparsable.push(raw);
      }
    }
  }
  return { categories, series, unparsable };
}
