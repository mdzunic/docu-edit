// Sheet charts (SPEC-132 "The chart model", "Resolving references",
// "Refresh", "Insert from selection", "Select Data", "Reference
// adjustment"): THIS FILE is the normative source for how a chart drawing's
// references are resolved against the grid, how its visual is refreshed from
// the cells, how charts follow structural edits, moves and sheet renames, and
// every chart command. Framework-free — core and spreadsheet siblings only;
// the view (useSheetCharts) wraps each command in one SPEC-121 undo step.
//
// SPEC-132 SKELETON: the signatures and contracts below are final; the bodies
// are filled in by the charts-model change.
import type { ChartOptions, ChartSeries, ChartType, ChartVisual } from "../core/charts";
import type { ModelColor } from "../core/theme";
import type { GridRange } from "./address";
import type { DrawingCommandResult } from "./drawings";
import type { CellKey } from "./formula/dependencies";
import type { SheetRemap } from "./formula/sheetRefs";
import type { GridMetrics } from "./geometry";
import type { CellValue, SheetChartContent, SheetChartRefs, SheetDrawing, SpreadsheetModel } from "./model";
import type { RangeMove } from "./move";
import type { SheetSelection } from "./selection";
import type { StructuralEdit } from "./structure";

// --- labels and messages ---------------------------------------------------------------------

/** The undo labels of the chart commands (SPEC-131's "Delete chart", "Move
 * chart" and "Resize chart" are drawings.ts's, unchanged). */
export const CHART_UNDO_LABELS = {
  insert: "Insert chart",
  data: "Edit chart data",
  options: "Chart options",
  seriesColor: "Chart series colour",
  source: "Chart data range",
  duplicate: "Duplicate chart",
} as const;

/** Insert ▸ Chart with nothing to chart (an empty selection whose current
 * region is empty). */
export const NO_CHART_DATA_MESSAGE = "Select the data you want to chart first";
/** Insert ▸ Chart over several areas sharing neither rows nor columns. */
export const CHART_AREAS_MESSAGE = "Select one block of data, or blocks that share the same rows";
/** The Select Data dialog over references that do not tile one rectangle. */
export const COMPLEX_CHART_RANGE_MESSAGE = "The data range is too complex to be displayed";
/** "Read-only charts": a `c:pivotSource` part. */
export const PIVOT_CHART_REASON = "Pivot charts can't be edited yet";
/** "Read-only charts": a `c:multiLvlStrRef` category reference. */
export const MULTI_LEVEL_CHART_REASON = "Charts with multi-level categories can't be edited yet";
/** A chart sheet's disabled Duplicate (its title). */
export const CHART_SHEET_DUPLICATE_TITLE = "Chart sheets can't be duplicated yet";

// --- the model ---------------------------------------------------------------------------------

/** A drawing's chart content, or undefined for any other kind. */
export function chartContentOf(drawing: SheetDrawing): SheetChartContent | undefined {
  return drawing.content.kind === "chart" ? drawing.content : undefined;
}

/** Whether a chart is modelled: a catalogue chart (no `unsupportedChartXml`).
 * Its `refs` are adjusted by the `charts` adjusters and (unless read-only)
 * refreshed; an unmodelled chart's part is left to SPEC-124's and SPEC-125's
 * `retained` rewrites. */
export function isModelledChart(content: SheetChartContent): boolean {
  return content.visual.unsupportedChartXml === undefined;
}

/** The visual of a chart outside the catalogue: painted as paintChart's
 * placeholder, never edited. */
export function unsupportedChartVisual(partText: string): ChartVisual {
  return { chartType: "columnClustered", categories: [], series: [], options: { legend: "none", dataLabels: false, gridlines: false }, unsupportedChartXml: partText };
}

// --- resolving references ------------------------------------------------------------------

/** SPEC-133's defined-name hook: the name's canonical reference text, or null. */
export type ChartNameResolver = (name: string, sheetId: string) => string | null;

export interface ResolvedChartArea {
  sheetId: string;
  range: GridRange;
}

/** A chart reference resolved against the grid ("Resolving references"):
 * a cell or range reference (no `sheetEnd`), a parenthesised union of them,
 * or (through `resolveName`) a defined name resolving to one. Unqualified
 * references are relative to `hostSheetId` (the chart's own sheet). Anything
 * else — a function, an external `[1]` reference, `#REF!`, a name without
 * the hook, a sheet that does not exist or is not a grid — is null. */
export function resolveChartRange(model: SpreadsheetModel, canonical: string, resolveName?: ChartNameResolver, hostSheetId?: string): ResolvedChartArea[] | null {
  void model;
  void canonical;
  void resolveName;
  void hostSheetId;
  return null;
}

/** One cell of a resolved reference, in reading order (a union's areas in
 * order; a row-shaped area left to right, a column-shaped one top to bottom,
 * a block row-major): its value (null when blank), its SPEC-118 formatValue
 * text, and whether its row or column is hidden (user- or filter-hidden). */
export interface ChartCell {
  sheetId: string;
  row: number;
  col: number;
  value: CellValue;
  text: string;
  hidden: boolean;
}

/** Every cell of a reference in reading order — blank cells included — or
 * null when it does not resolve. What the xlsx patch path writes caches from
 * ("Writing" 2). */
export function chartReferenceCells(model: SpreadsheetModel, canonical: string, hostSheetId: string, resolveName?: ChartNameResolver): ChartCell[] | null {
  void model;
  void canonical;
  void hostSheetId;
  void resolveName;
  return null;
}

// --- refresh ---------------------------------------------------------------------------------

/** The chart content refreshed from the cells ("Refresh"): categories, each
 * series' values and each referenced name rebuilt (hidden cells skipped
 * with plotVisibleOnly; blank and #N/A cells gaps; text and other errors
 * plot as 0; categories each cell's formatValue text, else "1", "2", …). A
 * series whose reference does not resolve keeps its cached values. The SAME
 * object when nothing changes, and for an unmodelled, read-only or
 * reference-less chart. Never sets a flag. */
export function refreshedChartContent(model: SpreadsheetModel, sheetId: string, content: SheetChartContent): SheetChartContent {
  void model;
  void sheetId;
  return content;
}

/** Refreshes one chart drawing in the model (copy-on-write over the sheet's
 * `drawings`; untouched drawings keep their identity): true when its visual
 * changed. A changed chart gains `cachesDirty` — unless `forDisplay` (open,
 * for a part with a reference but no cache), which changes only the visual.
 * Marks nothing else: the chart part is written by the chart write paths,
 * whatever the sheet's own dirty flags. */
export function refreshSheetChart(model: SpreadsheetModel, sheetId: string, drawingId: string, options?: { forDisplay?: boolean }): boolean {
  void model;
  void sheetId;
  void drawingId;
  void options;
  return false;
}

/** The modelled, refreshable charts (every sheet) whose resolved references
 * intersect a changed cell — an interval test over each chart's resolved
 * areas, cached per chart (by its `refs` object) and invalidated when the
 * refs change. */
export function chartsAffectedBy(model: SpreadsheetModel, changed: readonly CellKey[]): { sheetId: string; drawingId: string }[] {
  void model;
  void changed;
  return [];
}

/** The modelled, refreshable charts with `plotVisibleOnly` whose references
 * read cells on `sheetId` (refreshed when that sheet's row or column
 * visibility changes). */
export function chartsReadingSheet(model: SpreadsheetModel, sheetId: string): { sheetId: string; drawingId: string }[] {
  void model;
  void sheetId;
  return [];
}

// --- insert from selection --------------------------------------------------------------------

/** How a source range becomes series ("Insert from selection"). */
export interface ChartSourceFlags {
  seriesIn: "columns" | "rows";
  firstRowNames: boolean; // the first row (columns) / column (rows) holds series names
  firstColCategories: boolean; // the first column (columns) / row (rows) holds categories (scatter: names only)
}

export type ChartBuild =
  | { ok: true; visual: ChartVisual; refs: SheetChartRefs; source: { sheetId: string; ranges: GridRange[] } }
  | { ok: false; message: string };

/** A chart built from a selection ("Insert from selection"): the source (a
 * single cell expands to its currentRegion; several areas must share rows
 * or columns), header detection, orientation, scatter's x column/row,
 * accent colours and default options (title = the one series' name). */
export function buildChartFromRange(model: SpreadsheetModel, sheetId: string, selection: Pick<SheetSelection, "ranges" | "active">, chartType: ChartType): ChartBuild {
  void model;
  void sheetId;
  void selection;
  void chartType;
  return { ok: false, message: NO_CHART_DATA_MESSAGE };
}

/** The same builder with explicit flags (the Select Data dialog): the refs
 * of `range` (canonical, sheet-qualified; a union of areas allowed), or null
 * when it does not resolve to areas on one grid sheet. */
export function chartRefsFromRange(model: SpreadsheetModel, hostSheetId: string, range: string, flags: ChartSourceFlags, chartType: ChartType): SheetChartRefs | null {
  void model;
  void hostSheetId;
  void range;
  void flags;
  void chartType;
  return null;
}

/** The inverse of chartRefsFromRange: the one rectangle (canonical,
 * sheet-qualified, absolute) the refs tile and the flags that rebuild them
 * — null when they do not tile one rectangle on one sheet. */
export function chartRangeOf(refs: SheetChartRefs, chartType: ChartType): { range: string; flags: ChartSourceFlags } | null {
  void refs;
  void chartType;
  return null;
}

// --- commands ----------------------------------------------------------------------------------
// Each returns the new sheet (copy-on-write over `drawings`) for
// drawings.ts applyDrawingResult, or null when nothing changes; the view
// records one step labelled CHART_UNDO_LABELS[…] around it.

/** Insert ▸ Chart: the chart built from the selection, anchored twoCell
 * (editAs "twoCell") at NEW_CHART_WIDTH_CM × NEW_CHART_HEIGHT_CM with its
 * top-left two columns right of the source on the source's top row, on
 * top, named "Chart {k}", selected (`drawingId`). `{ message }` when there
 * is nothing to chart (NO_CHART_DATA_MESSAGE, CHART_AREAS_MESSAGE). */
export function insertSheetChart(
  model: SpreadsheetModel,
  sheetId: string,
  selection: Pick<SheetSelection, "ranges" | "active">,
  chartType: ChartType,
  metrics?: GridMetrics,
): DrawingCommandResult | { message: string } {
  void model;
  void sheetId;
  void selection;
  void chartType;
  void metrics;
  return { message: NO_CHART_DATA_MESSAGE };
}

/** The data sheet's Apply on a chart without refs (SPEC-052): categories
 * and series replaced, `edited`. */
export function setSheetChartData(model: SpreadsheetModel, sheetId: string, drawingId: string, categories: string[], series: ChartSeries[]): DrawingCommandResult | null {
  void model;
  void sheetId;
  void drawingId;
  void categories;
  void series;
  return null;
}

/** Chart options' Apply: options replaced, `edited`. */
export function setSheetChartOptions(model: SpreadsheetModel, sheetId: string, drawingId: string, options: ChartOptions): DrawingCommandResult | null {
  void model;
  void sheetId;
  void drawingId;
  void options;
  return null;
}

/** Chart options' series colour: one series' colour, `edited`. */
export function setSheetChartSeriesColor(model: SpreadsheetModel, sheetId: string, drawingId: string, seriesIndex: number, color: ModelColor): DrawingCommandResult | null {
  void model;
  void sheetId;
  void drawingId;
  void seriesIndex;
  void color;
  return null;
}

/** Select Data's Apply ("Select Data"): new refs. When the series count, the
 * chart type and the presence of every reference (name, categories, values)
 * per series are unchanged: `refsDirty` + refresh (the patch path);
 * otherwise the visual is rebuilt from the cells and the chart `edited`. */
export function setSheetChartSource(model: SpreadsheetModel, sheetId: string, drawingId: string, refs: SheetChartRefs): DrawingCommandResult | null {
  void model;
  void sheetId;
  void drawingId;
  void refs;
  return null;
}

/** Duplicate chart: a new chart from the source's visual + refs (no
 * sourcePart, no sourceXml — its own new part, regenerated), offset one row
 * and one column, on top, selected. */
export function duplicateSheetChart(model: SpreadsheetModel, sheetId: string, drawingId: string, metrics?: GridMetrics): DrawingCommandResult | null {
  void model;
  void sheetId;
  void drawingId;
  void metrics;
  return null;
}

// --- reference adjustment ----------------------------------------------------------------------

/** The `charts` REFERENCE_ADJUSTER ("Reference adjustment"): adjustFormula
 * over every modelled chart's refs on every sheet (host = the chart's
 * sheet), `refsDirty` where one changed, then a refresh (read-only charts
 * only adjust); ods: every chart frame's notify-on-update-of-ranges synced
 * (syncChartNotifyRanges). */
export function adjustChartsForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  void model;
  void edit;
}

/** The `charts` MOVE_ADJUSTER: repointFormula over every modelled chart's
 * refs, `refsDirty` + refresh as above. The drawing itself stays put. */
export function moveChartsForMove(model: SpreadsheetModel, move: RangeMove): void {
  void model;
  void move;
}

/** Sheet rename and delete (SPEC-121's commands call this with their
 * SheetRemap): remapSheetReferences over every chart's refs — a reference
 * into a deleted sheet becomes #REF! and its series keeps its caches —
 * `refsDirty` where one changed, then the notify sync (ods). */
export function remapChartReferences(model: SpreadsheetModel, remap: SheetRemap): void {
  void model;
  void remap;
}

/** ods: every chart frame whose `draw:notify-on-update-of-ranges` (in its
 * `sourceXml`) no longer matches its refs rewritten, its sheet marked
 * `dirty` (the frame lives in the table). A no-op for xlsx. */
export function syncChartNotifyRanges(model: SpreadsheetModel): void {
  void model;
}
