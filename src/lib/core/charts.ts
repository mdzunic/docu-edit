// The product-neutral half of SPEC-052's charts: the closed type catalogue,
// the data-sheet and options shapes, the default palette, the sample data a
// fresh chart starts with and the table-cells → chart-data mapping.
// Everything here is independent of the document model — a Slides product
// places the same charts on a slide instead of anchoring them to a
// paragraph. The document-specific half (anchoring, wrap, the editing
// commands) stays in lib/charts.ts and lib/documentModel.ts.
import type { ModelColor, ThemeColorRef } from "./theme";

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
  /** A literal "#rrggbb" or a theme colour reference (SPEC-077). */
  color: ModelColor;
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

/** A newly created/inserted series' colour (SPEC-077 AC-46): the theme's six
 * accents, cycling — `{ theme: "accent1" }` … `{ theme: "accent6" }` — so a
 * theme change recolours a fresh chart. A series read from a file with a
 * literal colour keeps it. */
export function chartAccentRef(index: number): ThemeColorRef {
  const slot = `accent${(index % 6) + 1}`;
  return { theme: slot as ThemeColorRef["theme"] };
}

/** The read-side fallback palette for a source chart series that carries no
 * colour of its own — literals, so opening a document never invents theme
 * references (SPEC-077 AC-48). */
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
    { name: "Series 1", values: [4.3, 2.5, 3.5], color: chartAccentRef(0) },
    { name: "Series 2", values: [2.4, 4.4, 1.8], color: chartAccentRef(1) },
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

/** Reads one table cell's text as a number the way a person typing into a
 * table would expect, or null when it is not a number. Beyond what `Number()`
 * accepts: surrounding whitespace, a leading or trailing currency sign
 * ("$10", "10 €"), a trailing percent ("10%" → 10, the number shown), a
 * trailing unit word ("10 kg"), accounting negatives ("(1,000)"), and the
 * usual digit-grouping and decimal conventions — "1,000.5" and "1.000,5",
 * "1 000", "1'000", and a decimal comma ("10,5"). A lone point is always a
 * decimal point ("1.500" is one and a half, as in `Number()`); a lone comma
 * is a decimal comma except in the one shape that reads as digit grouping —
 * one to three digits, the comma, exactly three digits ("1,000", "12,500") —
 * so "0,125" and "1234,5" stay decimals. */
export function parseChartNumber(raw: string): number | null {
  let text = raw.trim();
  if (text === "") {
    return null;
  }
  // Scientific notation is already a number; nothing below applies to it.
  if (/^[-+−]?\d+(?:\.\d+)?e[-+]?\d+$/i.test(text)) {
    const value = Number(text.replace("−", "-"));
    return Number.isFinite(value) ? value : null;
  }
  // Accounting negative: "(1,000)".
  let negative = false;
  if (/^\(.*\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  // Leading sign, then a leading currency sign; a trailing percent, currency
  // sign or unit word (letters and degree/percent signs only) — each may be
  // separated from the figure by whitespace.
  const signMatch = /^([-+−])\s*/.exec(text);
  if (signMatch) {
    negative = negative || signMatch[1] !== "+";
    text = text.slice(signMatch[0].length);
  }
  text = text.replace(/^[$€£¥₹₽¢]\s*/, "").replace(/\s*(?:%|[$€£¥₹₽¢]|[A-Za-z°µ][A-Za-z°µ²³/.]*)$/, "");
  // Digit grouping by space, non-breaking/thin space or apostrophe.
  text = text.replace(/[\s\u00a0\u202f']/g, "");
  if (!/^[\d.,]+$/.test(text) || !/\d/.test(text)) {
    return null;
  }
  const commas = text.split(",").length - 1;
  const points = text.split(".").length - 1;
  /** Strips a digit-grouping separator, insisting the groups it separated
   * were real thousands groups ("1,000,000", never "1,2,3"). */
  const ungroup = (grouped: string, separator: string): string | null =>
    new RegExp(`^\\d{1,3}(?:\\${separator}\\d{3})*$`).test(grouped) ? grouped.split(separator).join("") : null;
  if (commas > 0 && points > 0) {
    // Both present: the last one is the decimal separator.
    const decimal = text.lastIndexOf(",") > text.lastIndexOf(".") ? "," : ".";
    const grouping = decimal === "," ? "." : ",";
    const [whole, fraction, ...rest] = text.split(decimal);
    const ungrouped = ungroup(whole, grouping);
    if (rest.length > 0 || ungrouped === null || fraction.includes(grouping)) {
      return null;
    }
    text = `${ungrouped}.${fraction}`;
  } else if (commas > 1 || points > 1) {
    // Repeated separators can only be digit grouping: "1,000,000" / "1.000.000".
    const ungrouped = ungroup(text, commas > 1 ? "," : ".");
    if (ungrouped === null) {
      return null;
    }
    text = ungrouped;
  } else if (commas === 1) {
    const [whole, fraction] = text.split(",");
    const grouping = whole.length >= 1 && whole.length <= 3 && whole !== "0" && fraction.length === 3;
    text = grouping ? whole + fraction : `${whole}.${fraction}`;
  }
  const value = Number(text);
  if (!Number.isFinite(value)) {
    return null;
  }
  return negative ? -value : value;
}

/** Reads a table's cells into chart data. The first column (from its second
 * row on, or from its first when there is no header row) gives categories —
 * even when those are numbers, like years. The first row gives series names
 * when it is a header row: any of its cells beyond the first that is not a
 * number marks it as one; a first row of nothing but numbers is data, and
 * the series are named "Series 1", "Series 2"… Every other cell is a value,
 * read by parseChartNumber. A one-time copy — this module keeps no reference
 * to the table afterwards (SPEC-052 "Out of scope": no live link). */
export function chartDataFromTableCells(cellText: string[][]): ChartFromTableResult {
  const unparsable: string[] = [];
  const columns = Math.max(0, ...cellText.map((row) => row.length));
  if (cellText.length < 1 || columns < 2) {
    return { categories: [], series: [], unparsable };
  }
  const first = cellText[0];
  const firstRowIsHeader = Array.from({ length: columns - 1 }, (_, i) => (first[i + 1] ?? "").trim()).some(
    (text) => text !== "" && parseChartNumber(text) === null,
  );
  const seriesCount = columns - 1;
  const series: ChartSeries[] = Array.from({ length: seriesCount }, (_, i) => ({
    name: (firstRowIsHeader ? first[i + 1]?.trim() : "") || `Series ${i + 1}`,
    values: [],
    color: chartAccentRef(i),
  }));
  const categories: string[] = [];
  for (let r = firstRowIsHeader ? 1 : 0; r < cellText.length; r++) {
    const row = cellText[r];
    categories.push(row[0]?.trim() ?? "");
    for (let s = 0; s < seriesCount; s++) {
      const raw = (row[s + 1] ?? "").trim();
      if (raw === "") {
        series[s].values.push(null);
        continue;
      }
      const n = parseChartNumber(raw);
      if (n !== null) {
        series[s].values.push(n);
      } else {
        series[s].values.push(null);
        unparsable.push(raw);
      }
    }
  }
  if (categories.length === 0) {
    // A header row with no data rows under it: nothing to chart.
    return { categories: [], series: [], unparsable };
  }
  return { categories, series, unparsable };
}
