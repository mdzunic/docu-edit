// Sheet charts (SPEC-132 "The chart model", "Resolving references",
// "Refresh", "Insert from selection", "Select Data", "Reference
// adjustment"): THIS FILE is the normative source for how a chart drawing's
// references are resolved against the grid, how its visual is refreshed from
// the cells, how charts follow structural edits, moves and sheet renames, and
// every chart command. Framework-free — core and spreadsheet siblings only;
// the view (useSheetCharts) wraps each command in one SPEC-121 undo step.
//
// Resolving. A reference is canonical formula text (no "="), parsed with
// parseFormula and walked: a cell or range `ref` (no 3-D `sheetEnd`) on a
// GRID sheet — named case-insensitively, or the host (the chart's own sheet)
// when unqualified; a `paren` around it; a union (`,`, nested allowed, with
// or without the parentheses Excel writes); a `name` through the optional
// resolveName hook (SPEC-133's), whose canonical text is resolved again (at
// most MAX_NAME_DEPTH deep, so a name naming itself is null). Anything else —
// a function, an external `[1]` reference, `#REF!`, a name without the hook,
// a missing or non-grid sheet, text that does not parse — is null, and the
// series it feeds keeps its cached values. A whole-column (`B:B`) or
// whole-row reference is clipped to the sheet's used range for reading (a
// sheet storing nothing gives it no cells); chartsAffectedBy tests it
// unclipped, so typing below the used range still refreshes the chart.
// Reading order: a union's areas in order, each area row-major (a row-shaped
// area left to right, a column-shaped one top to bottom, a block row by
// row).
//
// Refresh ("Refresh"; AC-14):
// | Cell | values | categories | names |
// | number | the number | its formatValue text (scatter x: String(number)) | its text |
// | blank, #N/A | null (a gap) | "" / "#N/A" | "" / "#N/A" |
// | text, boolean, other error | 0 | its text | its text |
// With `plotVisibleOnly`, a cell in a user- or filter-hidden row or a hidden
// column is skipped by values and categories (never by a name). Categories
// come from the first series holding a categories reference, else "1", "2",
// … up to the longest series. A multi-cell name joins its non-empty texts
// with " ". A reference that does not resolve keeps what the visual holds.
// Colours, options and `unsupportedChartXml` are never touched; an equal
// visual is the SAME object (deep compare), and so is an unmodelled,
// read-only or reference-less chart. Display texts are SPEC-118's
// formatValue in the default locale (en-US), the workbook's date system,
// no width fitting.
//
// Insert from selection (buildChartFromRange):
// - Source: a multi-cell selection (a whole-row/column band clipped to the
//   used range); one cell (or one merged cell) expands to its currentRegion;
//   several areas must span the same rows (columns read left to right) or
//   the same columns (rows read top to bottom), else CHART_AREAS_MESSAGE; a
//   source holding no value at all is NO_CHART_DATA_MESSAGE.
// - A cell counts as a label when it holds text, is the empty top-left
//   corner, or holds a number whose format's first section is a date/time.
//   The first row is labels when every cell of it but the top-left counts
//   (a one-column block: the top-left itself) and some cell below it is a
//   non-label number; the first column is labels when every cell of it
//   below that row counts — unless no data column would remain.
// - Orientation over what remains: series down the columns when rows ≥
//   columns, else across the rows. The label row then holds series names
//   and the label column categories (rows: the other way round).
// - Scatter: no label column — the first data column (row) is every
//   series' x (`categories`, i.e. c:xVal), each following one a y series;
//   with a single data column there is no x (numbered). Labels are names.
// - Colours chartAccentRef(i); options defaultChartOptions(), titled with
//   the one series' name when there is exactly one and it names a cell.
//   Series without a name cell are "Series {i+1}". References are absolute
//   and sheet-qualified (`Data!$B$2:$B$5`; a series crossing several areas
//   is a parenthesised union); an ods chart also gets `plotArea` (the source).
//   The visual is computed from the cells exactly as a refresh would.
//
// Commands (each returns drawings.ts's DrawingCommandResult — the sheet with
// a new `drawings` array, untouched drawings by identity — or null when
// nothing changes; unmodelled and read-only charts refuse every data,
// options, colour, source and duplicate command):
// | Command | Label | Content | Layer marked changed |
// | insertSheetChart | Insert chart | new: visual + refs, no sourcePart | yes (xlsx drawingsDirty, ods dirty) |
// | duplicateSheetChart | Duplicate chart | new: a clone of visual + refs, anchor +1 row +1 column | yes |
// | setSheetChartData | Edit chart data | categories + series, `edited` (literal charts only) | no |
// | setSheetChartOptions | Chart options | options, `edited` | no |
// | setSheetChartSeriesColor | Chart series colour | one colour, `edited` | no |
// | setSheetChartSource | Chart data range | refs: patch (`refsDirty` + refresh) or regenerate (`edited`) | ods: when the frame's notify ranges change |
// A chart-content command marks nothing: the chart part is written by the
// chart write paths whatever the sheet's own dirty flags. Select Data keeps
// the chart's `plotVisibleOnly` (the dialog does not edit it).
//
// Reference adjustment. The `charts` REFERENCE_ADJUSTER (structure.ts, 12th)
// and MOVE_ADJUSTER (move.ts, 8th) map every modelled chart's references on
// every sheet (host = the chart's sheet) through adjustFormula /
// repointFormula; a changed reference set is a new `refs` object plus
// `refsDirty`; then every refreshable chart whose references changed or read
// the edited (moved-from or moved-to) sheet is refreshed, gaining
// `cachesDirty` when its visual changed (a read-only chart only adjusts). A
// reference turned `#REF!` no longer resolves, so its series keeps its
// caches. remapChartReferences applies SPEC-121's SheetRemap the same way
// without a refresh (a rename or delete changes no cell). All three end with
// syncChartNotifyRanges. Copy-on-write throughout: only what changes is
// replaced, and nothing an undo snapshot may hold is mutated.
import {
  chartAccentRef,
  defaultChartOptions,
  NEW_CHART_HEIGHT_CM,
  NEW_CHART_WIDTH_CM,
  type ChartOptions,
  type ChartSeries,
  type ChartType,
  type ChartVisual,
} from "../core/charts";
import type { ModelColor } from "../core/theme";
import { getCell, usedRange, type CellRef, type GridRange } from "./address";
import { anchorFromRectPx, EMU_PER_PX, nextDrawingId, nextDrawingNumber, withDrawings, type DrawingCommandResult } from "./drawings";
import { adjustFormula } from "./formula/adjust";
import type { FormulaNode } from "./formula/ast";
import { parseCellKey, type CellKey } from "./formula/dependencies";
import { parseFormula } from "./formula/parse";
import { formatReference } from "./formula/print";
import type { ReferenceNode } from "./formula/references";
import { remapSheetReferences, type SheetRemap } from "./formula/sheetRefs";
import { buildGridMetrics, type GridMetrics } from "./geometry";
import { buildMergeIndex } from "./merges";
import {
  isCellError,
  MAX_COLS,
  MAX_ROWS,
  rowIsHidden,
  type AnchorPoint,
  type CellValue,
  type DrawingAnchor,
  type Row,
  type Sheet,
  type SheetChartContent,
  type SheetChartRefs,
  type SheetChartSeriesRefs,
  type SheetDrawing,
  type SpreadsheetModel,
} from "./model";
import { repointFormula, type RangeMove } from "./move";
import { formatCodeOf } from "./numberFormat/builtin";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { parseFormatCode } from "./numberFormat/parse";
import { formatValue, plainText, workbookFormatContext, type FormatContext } from "./numberFormat/render";
import { notifyRangesOf, withNotifyRanges } from "./ods/chartXml";
import { currentRegion, type NavContext, type SheetSelection } from "./selection";
import { cellStyle, columnSpanAt, effectiveStyleId } from "./styles";
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

/** How deep a defined name may lead to another before resolution gives up. */
const MAX_NAME_DEPTH = 16;
/** CSS px per cm at 96 dpi. */
const PX_PER_CM = 96 / 2.54;

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

/** A modelled chart that is not read-only: one a refresh and the data
 * commands may change. */
function isEditableChart(content: SheetChartContent): boolean {
  return isModelledChart(content) && content.readOnlyReason === undefined;
}

/** A modelled, writable chart with references: one refresh rebuilds. */
function isRefreshable(content: SheetChartContent): content is SheetChartContent & { refs: SheetChartRefs } {
  return isEditableChart(content) && content.refs !== undefined;
}

// --- resolving references ------------------------------------------------------------------

/** SPEC-133's defined-name hook: the name's canonical reference text, or null. */
export type ChartNameResolver = (name: string, sheetId: string) => string | null;

export interface ResolvedChartArea {
  sheetId: string;
  range: GridRange;
}

const sheetById = (model: SpreadsheetModel, id: string): Sheet | undefined => model.sheets.find((sheet) => sheet.id === id);

function gridSheetByName(model: SpreadsheetModel, name: string): Sheet | undefined {
  const upper = name.toUpperCase();
  return model.sheets.find((sheet) => sheet.kind === "grid" && sheet.name.toUpperCase() === upper);
}

/** One reference's area: null when it cannot be one, undefined when it is
 * a whole row/column clipped to an empty sheet (resolved, no cells). */
function referenceArea(model: SpreadsheetModel, ref: ReferenceNode, hostSheetId: string | undefined, clip: boolean): ResolvedChartArea | null | undefined {
  if (ref.sheetEnd !== undefined) {
    return null;
  }
  const host = hostSheetId === undefined ? undefined : sheetById(model, hostSheetId);
  const sheet = ref.sheet !== undefined ? gridSheetByName(model, ref.sheet) : host?.kind === "grid" ? host : undefined;
  if (!sheet) {
    return null;
  }
  const start = ref.start;
  const end = ref.end ?? ref.start;
  const allRows = start.row === undefined || end.row === undefined;
  const allCols = start.col === undefined || end.col === undefined;
  const range: GridRange = {
    top: allRows ? 0 : Math.min(start.row!, end.row!),
    bottom: allRows ? MAX_ROWS - 1 : Math.max(start.row!, end.row!),
    left: allCols ? 0 : Math.min(start.col!, end.col!),
    right: allCols ? MAX_COLS - 1 : Math.max(start.col!, end.col!),
  };
  if (clip && (allRows || allCols)) {
    const used = usedRange(sheet);
    if (!used) {
      return undefined;
    }
    if (allRows) {
      range.top = used.top;
      range.bottom = used.bottom;
    }
    if (allCols) {
      range.left = used.left;
      range.right = used.right;
    }
  }
  return { sheetId: sheet.id, range };
}

/** resolveChartRange's walk (see the module header); `clip` clips whole
 * rows and columns to the used range. */
function resolveAreas(
  model: SpreadsheetModel,
  text: string,
  resolveName: ChartNameResolver | undefined,
  hostSheetId: string | undefined,
  clip: boolean,
  depth: number,
): ResolvedChartArea[] | null {
  if (depth > MAX_NAME_DEPTH) {
    return null;
  }
  let node: FormulaNode;
  try {
    node = parseFormula(text.startsWith("=") ? text.slice(1) : text);
  } catch {
    return null;
  }
  const areas: ResolvedChartArea[] = [];
  const walk = (each: FormulaNode): boolean => {
    switch (each.kind) {
      case "ref": {
        const area = referenceArea(model, each.ref, hostSheetId, clip);
        if (area === null) {
          return false;
        }
        if (area) {
          areas.push(area);
        }
        return true;
      }
      case "paren":
        return walk(each.inner);
      case "binary":
        return each.op === "," && walk(each.left) && walk(each.right);
      case "name": {
        if (!resolveName) {
          return false;
        }
        let scope = hostSheetId ?? "";
        if (each.sheet !== undefined) {
          const upper = each.sheet.toUpperCase();
          const scoped = model.sheets.find((sheet) => sheet.name.toUpperCase() === upper);
          if (!scoped) {
            return false;
          }
          scope = scoped.id;
        }
        const target = resolveName(each.name, scope);
        const inner = target === null ? null : resolveAreas(model, target, resolveName, hostSheetId, clip, depth + 1);
        if (!inner) {
          return false;
        }
        areas.push(...inner);
        return true;
      }
      default:
        return false;
    }
  };
  return walk(node) ? areas : null;
}

/** A chart reference resolved against the grid ("Resolving references"):
 * a cell or range reference (no `sheetEnd`), a parenthesised union of them,
 * or (through `resolveName`) a defined name resolving to one. Unqualified
 * references are relative to `hostSheetId` (the chart's own sheet). Anything
 * else — a function, an external `[1]` reference, `#REF!`, a name without
 * the hook, a sheet that does not exist or is not a grid — is null. A whole
 * row or column is clipped to the sheet's used range (none on an empty
 * sheet). */
export function resolveChartRange(model: SpreadsheetModel, canonical: string, resolveName?: ChartNameResolver, hostSheetId?: string): ResolvedChartArea[] | null {
  return resolveAreas(model, canonical, resolveName, hostSheetId, true, 0);
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

/** The context display texts format in (see the module header). */
function chartFormatContext(model: SpreadsheetModel): FormatContext {
  return workbookFormatContext(model, SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]);
}

/** The first index of `length` sorted positions at or after `target`. */
function lowerBound(length: number, position: (i: number) => number, target: number): number {
  let lo = 0;
  let hi = length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (position(mid) < target) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** Every cell of `areas` in reading order; texts only when `ctx` is given
 * (a values reference needs none). */
function readAreas(model: SpreadsheetModel, areas: readonly ResolvedChartArea[], ctx: FormatContext | undefined): ChartCell[] {
  const cells: ChartCell[] = [];
  for (const { sheetId, range } of areas) {
    const sheet = sheetById(model, sheetId);
    if (!sheet) {
      continue;
    }
    const colHidden: boolean[] = [];
    for (let col = range.left; col <= range.right; col++) {
      colHidden.push(columnSpanAt(sheet.columns, col)?.hidden === true);
    }
    const rows = sheet.rows;
    let rowAt = lowerBound(rows.length, (i) => rows[i]!.index, range.top);
    for (let r = range.top; r <= range.bottom; r++) {
      let row: Row | undefined;
      if (rows[rowAt]?.index === r) {
        row = rows[rowAt];
        rowAt += 1;
      }
      const rowHidden = rowIsHidden(row);
      const stored = row?.cells ?? [];
      let cellAt = lowerBound(stored.length, (i) => stored[i]!.col, range.left);
      for (let c = range.left; c <= range.right; c++) {
        let value: CellValue = null;
        if (stored[cellAt]?.col === c) {
          value = stored[cellAt]!.value;
          cellAt += 1;
        }
        const text = ctx === undefined || value === null ? "" : plainText(formatValue(value, formatCodeOf(cellStyle(model, effectiveStyleId(sheet, r, c)), ctx.locale), ctx));
        cells.push({ sheetId, row: r, col: c, value, text, hidden: rowHidden || colHidden[c - range.left]! });
      }
    }
  }
  return cells;
}

/** Every cell of a reference in reading order — blank cells included — or
 * null when it does not resolve. What the xlsx patch path writes caches from
 * ("Writing" 2). */
export function chartReferenceCells(model: SpreadsheetModel, canonical: string, hostSheetId: string, resolveName?: ChartNameResolver): ChartCell[] | null {
  const areas = resolveChartRange(model, canonical, resolveName, hostSheetId);
  return areas ? readAreas(model, areas, chartFormatContext(model)) : null;
}

// --- refresh ---------------------------------------------------------------------------------

/** A cell as a plotted value (see the module header's table). */
function plotValue(value: CellValue): number | null {
  if (value === null) {
    return null;
  }
  if (typeof value === "number") {
    return value;
  }
  if (isCellError(value)) {
    return value.error === "#N/A" ? null : 0;
  }
  return 0;
}

const sameStrings = (a: readonly string[], b: readonly string[]): boolean => a.length === b.length && a.every((each, i) => each === b[i]);
const sameValues = (a: readonly (number | null)[], b: readonly (number | null)[]): boolean => a.length === b.length && a.every((each, i) => Object.is(each, b[i]));

/** "1", "2", … `count`. */
function numbering(count: number): string[] {
  return Array.from({ length: count }, (_, i) => String(i + 1));
}

/** `visual` with its categories, values and referenced names read from the
 * cells under `refs` — the same object when nothing differs; an unchanged
 * series keeps its identity. */
function visualFromCells(model: SpreadsheetModel, hostSheetId: string, refs: SheetChartRefs, visual: ChartVisual, resolveName?: ChartNameResolver): ChartVisual {
  const ctx = chartFormatContext(model);
  const shown = (cells: ChartCell[]): ChartCell[] => (refs.plotVisibleOnly ? cells.filter((cell) => !cell.hidden) : cells);
  const read = (text: string, withText: boolean): ChartCell[] | null => {
    const areas = resolveChartRange(model, text, resolveName, hostSheetId);
    return areas ? readAreas(model, areas, withText ? ctx : undefined) : null;
  };
  let seriesChanged = false;
  const series = visual.series.map((each, i): ChartSeries => {
    const own = refs.series[i];
    if (!own) {
      return each;
    }
    let values = each.values;
    if (own.values !== undefined) {
      const cells = read(own.values, false);
      if (cells) {
        const next = shown(cells).map((cell) => plotValue(cell.value));
        values = sameValues(next, each.values) ? each.values : next;
      }
    }
    let name = each.name;
    if (own.name !== undefined) {
      const cells = read(own.name, true);
      if (cells) {
        name = cells
          .map((cell) => cell.text)
          .filter((text) => text !== "")
          .join(" ");
      }
    }
    if (values === each.values && name === each.name) {
      return each;
    }
    seriesChanged = true;
    return { ...each, name, values };
  });
  let categories = visual.categories;
  const categoryRef = refs.series.find((each) => each.categories !== undefined)?.categories;
  if (categoryRef !== undefined) {
    const cells = read(categoryRef, true);
    if (cells) {
      const scatter = visual.chartType === "scatter";
      categories = shown(cells).map((cell) => (scatter && typeof cell.value === "number" ? String(cell.value) : cell.text));
    }
  } else {
    categories = numbering(series.reduce((longest, each) => Math.max(longest, each.values.length), 0));
  }
  if (sameStrings(categories, visual.categories)) {
    categories = visual.categories;
  }
  if (!seriesChanged && categories === visual.categories) {
    return visual;
  }
  return { ...visual, categories, series };
}

/** The chart content refreshed from the cells ("Refresh"): categories, each
 * series' values and each referenced name rebuilt (hidden cells skipped
 * with plotVisibleOnly; blank and #N/A cells gaps; text and other errors
 * plot as 0; categories each cell's formatValue text, else "1", "2", …). A
 * series whose reference does not resolve keeps its cached values. The SAME
 * object when nothing changes, and for an unmodelled, read-only or
 * reference-less chart. Never sets a flag. */
export function refreshedChartContent(model: SpreadsheetModel, sheetId: string, content: SheetChartContent, resolveName?: ChartNameResolver): SheetChartContent {
  if (!isRefreshable(content)) {
    return content;
  }
  const visual = visualFromCells(model, sheetId, content.refs, content.visual, resolveName);
  return visual === content.visual ? content : { ...content, visual };
}

/** The sheet with drawing `index` replaced (a new `drawings` array, the
 * other drawings by identity); no dirty flag set. */
function withDrawingAt(sheet: Sheet, index: number, drawing: SheetDrawing): Sheet {
  const drawings = sheet.drawings!.slice();
  drawings[index] = drawing;
  return { ...sheet, drawings };
}

/** Refreshes one chart drawing in the model (copy-on-write over the sheet's
 * `drawings`; untouched drawings keep their identity): true when its visual
 * changed. A changed chart gains `cachesDirty` — unless `forDisplay` (open,
 * for a part with a reference but no cache), which changes only the visual.
 * Marks nothing else: the chart part is written by the chart write paths,
 * whatever the sheet's own dirty flags. */
export function refreshSheetChart(model: SpreadsheetModel, sheetId: string, drawingId: string, options?: { forDisplay?: boolean; resolveName?: ChartNameResolver }): boolean {
  const found = findChart(model, sheetId, drawingId);
  if (!found) {
    return false;
  }
  const next = refreshedChartContent(model, sheetId, found.content, options?.resolveName);
  if (next === found.content) {
    return false;
  }
  const content: SheetChartContent = options?.forDisplay ? next : { ...next, cachesDirty: true };
  const sheet = withDrawingAt(found.sheet, found.index, { ...found.drawing, content });
  model.sheets = model.sheets.map((each) => (each === found.sheet ? sheet : each));
  return true;
}

// --- which charts a change reaches ---------------------------------------------------------

interface AreaCacheEntry {
  key: string;
  areas: ResolvedChartArea[];
  /** SPEC-133: the names resolved through (copy-on-write), when a hook was given. */
  names?: unknown;
}

/** Each refs object's resolved areas (unclipped), keyed by the host and the
 * sheet list they were resolved against — a new refs object (refsDirty) or a
 * renamed, added or deleted sheet resolves again. */
const areaCache = new WeakMap<SheetChartRefs, AreaCacheEntry>();

/** What a resolution depends on: every sheet's id, name and kind. */
function sheetSignature(model: SpreadsheetModel): string {
  return model.sheets.map((sheet) => `${sheet.id}\u0000${sheet.name}\u0000${sheet.kind}`).join("\u0001");
}

/** Every area a chart's references read (names, categories, values), cached. */
function chartAreas(model: SpreadsheetModel, hostSheetId: string, refs: SheetChartRefs, signature: string, resolveName?: ChartNameResolver): ResolvedChartArea[] {
  const key = `${hostSheetId}\u0002${signature}`;
  const names = resolveName ? model.names : undefined;
  const hit = areaCache.get(refs);
  if (hit?.key === key && hit.names === names) {
    return hit.areas;
  }
  const areas: ResolvedChartArea[] = [];
  for (const series of refs.series) {
    for (const text of [series.name, series.categories, series.values]) {
      if (text !== undefined) {
        areas.push(...(resolveAreas(model, text, resolveName, hostSheetId, false, 0) ?? []));
      }
    }
  }
  areaCache.set(refs, { key, areas, names });
  return areas;
}

/** Every refreshable chart drawing, sheet by sheet in z-order. */
function* refreshableCharts(model: SpreadsheetModel): Generator<{ sheet: Sheet; drawing: SheetDrawing; refs: SheetChartRefs }> {
  for (const sheet of model.sheets) {
    for (const drawing of sheet.drawings ?? []) {
      const content = chartContentOf(drawing);
      if (content && isRefreshable(content)) {
        yield { sheet, drawing, refs: content.refs };
      }
    }
  }
}

/** Changed cells of one sheet, sorted by row (cols parallel). */
interface ChangedCells {
  rows: number[];
  cols: number[];
}

/** Whether any changed cell lies inside `range`. */
function touches(changed: ChangedCells, range: GridRange): boolean {
  const { rows, cols } = changed;
  for (let i = lowerBound(rows.length, (k) => rows[k]!, range.top); i < rows.length && rows[i]! <= range.bottom; i++) {
    if (cols[i]! >= range.left && cols[i]! <= range.right) {
      return true;
    }
  }
  return false;
}

/** The modelled, refreshable charts (every sheet) whose resolved references
 * intersect a changed cell — an interval test over each chart's resolved
 * areas, cached per chart (by its `refs` object) and invalidated when the
 * refs change. */
export function chartsAffectedBy(model: SpreadsheetModel, changed: readonly CellKey[], resolveName?: ChartNameResolver): { sheetId: string; drawingId: string }[] {
  if (changed.length === 0) {
    return [];
  }
  const points = new Map<string, { row: number; col: number }[]>();
  for (const key of changed) {
    const { sheetId, row, col } = parseCellKey(key);
    let list = points.get(sheetId);
    if (!list) {
      list = [];
      points.set(sheetId, list);
    }
    list.push({ row, col });
  }
  const bySheet = new Map<string, ChangedCells>();
  for (const [sheetId, list] of points) {
    list.sort((a, b) => a.row - b.row);
    bySheet.set(sheetId, { rows: list.map((each) => each.row), cols: list.map((each) => each.col) });
  }
  const signature = sheetSignature(model);
  const hits: { sheetId: string; drawingId: string }[] = [];
  for (const { sheet, drawing, refs } of refreshableCharts(model)) {
    const hit = chartAreas(model, sheet.id, refs, signature, resolveName).some((area) => {
      const cells = bySheet.get(area.sheetId);
      return cells !== undefined && touches(cells, area.range);
    });
    if (hit) {
      hits.push({ sheetId: sheet.id, drawingId: drawing.id });
    }
  }
  return hits;
}

/** The modelled, refreshable charts with `plotVisibleOnly` whose references
 * read cells on `sheetId` (refreshed when that sheet's row or column
 * visibility changes). */
export function chartsReadingSheet(model: SpreadsheetModel, sheetId: string): { sheetId: string; drawingId: string }[] {
  const signature = sheetSignature(model);
  const hits: { sheetId: string; drawingId: string }[] = [];
  for (const { sheet, drawing, refs } of refreshableCharts(model)) {
    if (refs.plotVisibleOnly && chartAreas(model, sheet.id, refs, signature).some((area) => area.sheetId === sheetId)) {
      hits.push({ sheetId: sheet.id, drawingId: drawing.id });
    }
  }
  return hits;
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

/** A source block: the rows and the columns it reads, ascending (several
 * areas sharing rows contribute their columns, and vice versa). */
interface Block {
  sheet: Sheet;
  rows: number[];
  cols: number[];
}

const span = (from: number, to: number): number[] => Array.from({ length: to - from + 1 }, (_, i) => from + i);
const sortedUnique = (values: number[]): number[] => [...new Set(values)].toSorted((a, b) => a - b);

/** The block several areas make, or null when they share neither rows nor
 * columns. */
function blockOf(sheet: Sheet, areas: readonly GridRange[]): Block | null {
  const first = areas[0]!;
  if (areas.every((area) => area.top === first.top && area.bottom === first.bottom)) {
    return { sheet, rows: span(first.top, first.bottom), cols: sortedUnique(areas.flatMap((area) => span(area.left, area.right))) };
  }
  if (areas.every((area) => area.left === first.left && area.right === first.right)) {
    return { sheet, rows: sortedUnique(areas.flatMap((area) => span(area.top, area.bottom))), cols: span(first.left, first.right) };
  }
  return null;
}

/** A selection range clipped to the used range along any axis it spans
 * whole (a selected column or row); null when nothing is left. */
function clippedSelectionRange(sheet: Sheet, range: GridRange): GridRange | null {
  const allRows = range.top === 0 && range.bottom === MAX_ROWS - 1;
  const allCols = range.left === 0 && range.right === MAX_COLS - 1;
  if (!allRows && !allCols) {
    return range;
  }
  const used = usedRange(sheet);
  if (!used) {
    return null;
  }
  const clipped = { ...range };
  if (allRows) {
    clipped.top = used.top;
    clipped.bottom = used.bottom;
  }
  if (allCols) {
    clipped.left = used.left;
    clipped.right = used.right;
  }
  return clipped;
}

function navContextOf(sheet: Sheet): NavContext {
  return {
    sheet,
    merges: buildMergeIndex(sheet.merges),
    isRowHidden: () => false,
    isColHidden: () => false,
    pageRows: 1,
    pageCols: 1,
  };
}

/** The selection's source areas ("Source"), or the message refusing them. */
function selectionSource(sheet: Sheet, selection: Pick<SheetSelection, "ranges" | "active">): GridRange[] | { message: string } {
  const at: CellRef = selection.active;
  const ranges = selection.ranges.length > 0 ? selection.ranges : [{ top: at.row, left: at.col, bottom: at.row, right: at.col }];
  if (ranges.length === 1) {
    const only = ranges[0]!;
    const oneCell = only.top === only.bottom && only.left === only.right;
    const oneMerge = sheet.merges.some((merge) => merge.top === only.top && merge.left === only.left && merge.bottom === only.bottom && merge.right === only.right);
    if (oneCell || oneMerge) {
      return [currentRegion({ row: only.top, col: only.left }, navContextOf(sheet))];
    }
  }
  const areas = ranges.map((range) => clippedSelectionRange(sheet, range)).filter((range): range is GridRange => range !== null);
  if (areas.length > 1 && !blockOf(sheet, areas)) {
    return { message: CHART_AREAS_MESSAGE };
  }
  return areas;
}

/** Whether any stored cell of `range` holds a value. */
function holdsValue(sheet: Sheet, range: GridRange): boolean {
  const rows = sheet.rows;
  for (let i = lowerBound(rows.length, (k) => rows[k]!.index, range.top); i < rows.length && rows[i]!.index <= range.bottom; i++) {
    if (rows[i]!.cells.some((cell) => cell.col >= range.left && cell.col <= range.right && cell.value !== null)) {
      return true;
    }
  }
  return false;
}

/** Whether a format code's first section renders a date or time. */
function isDateCode(code: string | undefined): boolean {
  if (code === undefined) {
    return false;
  }
  try {
    return parseFormatCode(code).sections[0]?.kind === "date";
  } catch {
    return false;
  }
}

/** The label rule ("Labels"): text, the empty top-left corner, or a number
 * formatted as a date or time. */
function isLabel(model: SpreadsheetModel, sheet: Sheet, row: number, col: number, corner: boolean): boolean {
  const value = getCell(sheet, row, col)?.value ?? null;
  if (typeof value === "string") {
    return true;
  }
  if (value === null) {
    return corner;
  }
  return typeof value === "number" && isDateCode(formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row, col)), SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]));
}

/** The flags a block's contents suggest ("Labels", "Orientation", "Scatter"). */
function detectedFlags(model: SpreadsheetModel, block: Block, chartType: ChartType): ChartSourceFlags {
  const { sheet, rows, cols } = block;
  const top = rows[0]!;
  const left = cols[0]!;
  const label = (row: number, col: number): boolean => isLabel(model, sheet, row, col, row === top && col === left);
  const headerCells = cols.length > 1 ? cols.slice(1) : cols;
  const labelRow =
    rows.length > 1 &&
    headerCells.every((col) => label(top, col)) &&
    rows.slice(1).some((row) => cols.some((col) => typeof getCell(sheet, row, col)?.value === "number" && !label(row, col)));
  const below = labelRow ? rows.slice(1) : rows;
  const labelCol = cols.length > 1 && below.every((row) => label(row, left));
  const dataRows = rows.length - (labelRow ? 1 : 0);
  const dataCols = cols.length - (labelCol ? 1 : 0);
  const seriesIn = dataRows >= dataCols ? "columns" : "rows";
  const names = seriesIn === "columns" ? labelRow : labelCol;
  const categories = seriesIn === "columns" ? labelCol : labelRow;
  // Scatter: the category line is the x data, whatever it holds.
  return { seriesIn, firstRowNames: names, firstColCategories: categories && chartType !== "scatter" };
}

/** An absolute, sheet-qualified reference to `range` on `sheetName`. */
function absoluteReference(sheetName: string, range: GridRange): string {
  const ref: ReferenceNode = { sheet: sheetName, start: { row: range.top, col: range.left, rowAbs: true, colAbs: true } };
  if (range.bottom !== range.top || range.right !== range.left) {
    ref.end = { row: range.bottom, col: range.right, rowAbs: true, colAbs: true };
  }
  return formatReference(ref);
}

/** Areas as one reference: the area's, or a parenthesised union. */
function areasReference(sheetName: string, areas: readonly GridRange[]): string {
  const parts = areas.map((area) => absoluteReference(sheetName, area));
  return parts.length === 1 ? parts[0]! : `(${parts.join(",")})`;
}

/** Ascending indices as runs of consecutive ones. */
function runsOf(indices: readonly number[]): [number, number][] {
  const runs: [number, number][] = [];
  for (const index of indices) {
    const last = runs.at(-1);
    if (last && last[1] + 1 === index) {
      last[1] = index;
    } else {
      runs.push([index, index]);
    }
  }
  return runs;
}

/** The reference to one line of a block: column `line` over `points` rows
 * (series in columns) or row `line` over `points` columns. */
function lineReference(sheetName: string, seriesIn: ChartSourceFlags["seriesIn"], line: number, points: readonly number[]): string {
  const areas = runsOf(points).map(([from, to]): GridRange => (seriesIn === "columns" ? { top: from, bottom: to, left: line, right: line } : { top: line, bottom: line, left: from, right: to }));
  return areasReference(sheetName, areas);
}

/** The refs a block makes under `flags` (null when no series or no point
 * would remain). */
function refsOfBlock(model: SpreadsheetModel, block: Block, flags: ChartSourceFlags, chartType: ChartType, sourceAreas: readonly GridRange[]): SheetChartRefs | null {
  const columns = flags.seriesIn === "columns";
  const lines = columns ? block.cols : block.rows;
  const points = columns ? block.rows : block.cols;
  const scatter = chartType === "scatter";
  const nameAt = flags.firstRowNames ? points[0] : undefined;
  const dataPoints = flags.firstRowNames ? points.slice(1) : points;
  let dataLines = lines;
  let categoryLine: number | undefined;
  if (!scatter && flags.firstColCategories) {
    categoryLine = lines[0];
    dataLines = lines.slice(1);
  } else if (scatter && lines.length > 1) {
    categoryLine = lines[0];
    dataLines = lines.slice(1);
  }
  if (dataPoints.length === 0 || dataLines.length === 0) {
    return null;
  }
  const sheetName = block.sheet.name;
  const categories = categoryLine === undefined ? undefined : lineReference(sheetName, flags.seriesIn, categoryLine, dataPoints);
  const series = dataLines.map((line): SheetChartSeriesRefs => {
    const refs: SheetChartSeriesRefs = {};
    if (nameAt !== undefined) {
      refs.name = absoluteReference(sheetName, columns ? { top: nameAt, bottom: nameAt, left: line, right: line } : { top: line, bottom: line, left: nameAt, right: nameAt });
    }
    if (categories !== undefined) {
      refs.categories = categories;
    }
    refs.values = lineReference(sheetName, flags.seriesIn, line, dataPoints);
    return refs;
  });
  const refs: SheetChartRefs = { series, plotVisibleOnly: true };
  if (model.format === "ods") {
    refs.plotArea = areasReference(sheetName, sourceAreas);
  }
  return refs;
}

/** A visual for `refs` computed from the cells as a refresh would: series
 * "Series {i+1}" until a name cell says otherwise, colours `colorOf(i)`. */
function visualOfRefs(model: SpreadsheetModel, hostSheetId: string, refs: SheetChartRefs, chartType: ChartType, options: ChartOptions, colorOf: (index: number) => ModelColor): ChartVisual {
  const base: ChartVisual = {
    chartType,
    categories: [],
    series: refs.series.map((_, i) => ({ name: `Series ${i + 1}`, values: [], color: colorOf(i) })),
    options,
  };
  return visualFromCells(model, hostSheetId, refs, base);
}

/** A chart built from a selection ("Insert from selection"): the source (a
 * single cell expands to its currentRegion; several areas must share rows
 * or columns), header detection, orientation, scatter's x column/row,
 * accent colours and default options (title = the one series' name). */
export function buildChartFromRange(model: SpreadsheetModel, sheetId: string, selection: Pick<SheetSelection, "ranges" | "active">, chartType: ChartType): ChartBuild {
  const sheet = sheetById(model, sheetId);
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NO_CHART_DATA_MESSAGE };
  }
  const source = selectionSource(sheet, selection);
  if (!Array.isArray(source)) {
    return { ok: false, message: source.message };
  }
  if (source.length === 0 || !source.some((range) => holdsValue(sheet, range))) {
    return { ok: false, message: NO_CHART_DATA_MESSAGE };
  }
  const block = blockOf(sheet, source)!;
  const refs = refsOfBlock(model, block, detectedFlags(model, block, chartType), chartType, source);
  if (!refs) {
    return { ok: false, message: NO_CHART_DATA_MESSAGE };
  }
  const visual = visualOfRefs(model, sheetId, refs, chartType, defaultChartOptions(), chartAccentRef);
  const only = visual.series.length === 1 ? visual.series[0]! : undefined;
  if (only && refs.series[0]!.name !== undefined && only.name !== "") {
    visual.options = { ...visual.options, title: only.name };
  }
  return { ok: true, visual, refs, source: { sheetId, ranges: source } };
}

/** The same builder with explicit flags (the Select Data dialog): the refs
 * of `range` (canonical, sheet-qualified; a union of areas allowed), or null
 * when it does not resolve to areas on one grid sheet. Scatter ignores
 * `firstColCategories` (its first data column/row is always the x). */
export function chartRefsFromRange(model: SpreadsheetModel, hostSheetId: string, range: string, flags: ChartSourceFlags, chartType: ChartType): SheetChartRefs | null {
  const areas = resolveChartRange(model, range, undefined, hostSheetId);
  if (!areas || areas.length === 0 || areas.some((area) => area.sheetId !== areas[0]!.sheetId)) {
    return null;
  }
  const sheet = sheetById(model, areas[0]!.sheetId)!;
  const ranges = areas.map((area) => area.range);
  const block = blockOf(sheet, ranges);
  return block ? refsOfBlock(model, block, flags, chartType, ranges) : null;
}

/** A reference written as one sheet-qualified area, or null. */
function qualifiedArea(text: string | undefined): { sheet: string; range: GridRange } | null {
  if (text === undefined) {
    return null;
  }
  let node: FormulaNode;
  try {
    node = parseFormula(text);
  } catch {
    return null;
  }
  while (node.kind === "paren") {
    node = node.inner;
  }
  if (node.kind !== "ref" || node.ref.sheet === undefined || node.ref.sheetEnd !== undefined) {
    return null;
  }
  const { start } = node.ref;
  const end = node.ref.end ?? start;
  if (start.row === undefined || start.col === undefined || end.row === undefined || end.col === undefined) {
    return null;
  }
  return {
    sheet: node.ref.sheet,
    range: { top: Math.min(start.row, end.row), left: Math.min(start.col, end.col), bottom: Math.max(start.row, end.row), right: Math.max(start.col, end.col) },
  };
}

const sameRange = (a: GridRange, b: GridRange): boolean => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;

/** chartRangeOf under one orientation, or null. */
function rangeOfOriented(refs: SheetChartRefs, chartType: ChartType, seriesIn: ChartSourceFlags["seriesIn"]): { range: string; flags: ChartSourceFlags } | null {
  const columns = seriesIn === "columns";
  const values = refs.series.map((each) => qualifiedArea(each.values));
  if (values.some((area) => area === null)) {
    return null;
  }
  const areas = values as { sheet: string; range: GridRange }[];
  const sheet = areas[0]!.sheet;
  const onSheet = (area: { sheet: string }): boolean => area.sheet.toUpperCase() === sheet.toUpperCase();
  // Each series one line: line index and its point span.
  const lineOf = (range: GridRange): number => (columns ? range.left : range.top);
  const first = areas[0]!.range;
  const pointFrom = columns ? first.top : first.left;
  const pointTo = columns ? first.bottom : first.right;
  for (const [i, area] of areas.entries()) {
    const r = area.range;
    const isLine = columns ? r.left === r.right && r.top === pointFrom && r.bottom === pointTo : r.top === r.bottom && r.left === pointFrom && r.right === pointTo;
    if (!onSheet(area) || !isLine || lineOf(r) !== lineOf(first) + i) {
      return null;
    }
  }
  const firstLine = lineOf(first);
  const lastLine = firstLine + areas.length - 1;
  const cellAt = (line: number, point: number): GridRange => (columns ? { top: point, bottom: point, left: line, right: line } : { top: line, bottom: line, left: point, right: point });
  const lineRange = (line: number): GridRange => (columns ? { top: pointFrom, bottom: pointTo, left: line, right: line } : { top: line, bottom: line, left: pointFrom, right: pointTo });
  // Names: all or none, each the cell before its line's first point.
  const named = refs.series.filter((each) => each.name !== undefined).length;
  if (named !== 0 && named !== refs.series.length) {
    return null;
  }
  const firstRowNames = named > 0;
  if (firstRowNames) {
    if (pointFrom === 0) {
      return null;
    }
    for (const [i, each] of refs.series.entries()) {
      const name = qualifiedArea(each.name);
      if (!name || !onSheet(name) || !sameRange(name.range, cellAt(firstLine + i, pointFrom - 1))) {
        return null;
      }
    }
  }
  // Categories (scatter: x): all or none, each the line before the first.
  const withCategories = refs.series.filter((each) => each.categories !== undefined).length;
  if (withCategories !== 0 && withCategories !== refs.series.length) {
    return null;
  }
  const scatter = chartType === "scatter";
  const hasCategoryLine = withCategories > 0;
  if (hasCategoryLine) {
    if (firstLine === 0) {
      return null;
    }
    for (const each of refs.series) {
      const categories = qualifiedArea(each.categories);
      if (!categories || !onSheet(categories) || !sameRange(categories.range, lineRange(firstLine - 1))) {
        return null;
      }
    }
  } else if (scatter && refs.series.length > 1) {
    return null; // rebuilt, the first y series would become the x
  }
  const startLine = hasCategoryLine ? firstLine - 1 : firstLine;
  const startPoint = firstRowNames ? pointFrom - 1 : pointFrom;
  const rect: GridRange = columns
    ? { top: startPoint, bottom: pointTo, left: startLine, right: lastLine }
    : { top: startLine, bottom: lastLine, left: startPoint, right: pointTo };
  return { range: absoluteReference(sheet, rect), flags: { seriesIn, firstRowNames, firstColCategories: hasCategoryLine && !scatter } };
}

/** The inverse of chartRefsFromRange: the one rectangle (canonical,
 * sheet-qualified, absolute) the refs tile and the flags that rebuild them
 * — null when they do not tile one rectangle on one sheet. */
export function chartRangeOf(refs: SheetChartRefs, chartType: ChartType): { range: string; flags: ChartSourceFlags } | null {
  if (refs.series.length === 0) {
    return null;
  }
  return rangeOfOriented(refs, chartType, "columns") ?? rangeOfOriented(refs, chartType, "rows");
}

// --- commands ----------------------------------------------------------------------------------
// Each returns the new sheet (copy-on-write over `drawings`) for
// drawings.ts applyDrawingResult, or null when nothing changes; the view
// records one step labelled CHART_UNDO_LABELS[…] around it.

interface FoundChart {
  sheet: Sheet;
  index: number;
  drawing: SheetDrawing;
  content: SheetChartContent;
}

function findChart(model: SpreadsheetModel, sheetId: string, drawingId: string): FoundChart | null {
  const sheet = sheetById(model, sheetId);
  const index = sheet?.drawings?.findIndex((drawing) => drawing.id === drawingId) ?? -1;
  if (!sheet || index < 0) {
    return null;
  }
  const drawing = sheet.drawings![index]!;
  const content = chartContentOf(drawing);
  return content ? { sheet, index, drawing, content } : null;
}

/** A chart-content command's result: the drawing's content replaced, the
 * drawing selected, no dirty flag (see the module header). */
function contentResult(found: FoundChart, content: SheetChartContent): DrawingCommandResult {
  return { sheet: withDrawingAt(found.sheet, found.index, { ...found.drawing, content }), drawingId: found.drawing.id };
}

/** A chart created on the sheet's name, "Chart {k}": k = nextDrawingNumber
 * (the cNvPr id the xlsx writer mints, − 1), raised past any drawing on the
 * sheet already so named — an imported chart's number need not follow its
 * cNvPr id (Excel numbers charts by a counter of its own). */
function newChartName(sheet: Sheet): string {
  const taken = new Set((sheet.drawings ?? []).map((drawing) => drawing.name.toLowerCase()));
  let k = nextDrawingNumber(sheet);
  while (taken.has(`chart ${k}`)) {
    k += 1;
  }
  return `Chart ${k}`;
}

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
  const built = buildChartFromRange(model, sheetId, selection, chartType);
  if (!built.ok) {
    return { message: built.message };
  }
  const sheet = sheetById(model, sheetId)!;
  const m = metrics ?? buildGridMetrics(sheet);
  const top = Math.min(...built.source.ranges.map((range) => range.top));
  const right = Math.max(...built.source.ranges.map((range) => range.right));
  const col = Math.min(right + 2, MAX_COLS - 1);
  const rect = { left: m.colLeft(col), top: m.rowTop(top), width: NEW_CHART_WIDTH_CM * PX_PER_CM, height: NEW_CHART_HEIGHT_CM * PX_PER_CM };
  const drawing: SheetDrawing = {
    id: nextDrawingId(sheet),
    anchor: anchorFromRectPx(rect, "moveAndSize", m),
    name: newChartName(sheet),
    content: { kind: "chart", visual: built.visual, refs: built.refs },
  };
  return { sheet: withDrawings(model, sheet, [...(sheet.drawings ?? []), drawing]), drawingId: drawing.id };
}

/** Plain-JSON deep equality (key order ignored; an undefined property is
 * an absent one). */
function sameJson(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) {
    return true;
  }
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null || Array.isArray(a) !== Array.isArray(b)) {
    return false;
  }
  if (Array.isArray(a)) {
    const other = b as unknown[];
    return a.length === other.length && a.every((each, i) => sameJson(each, other[i]));
  }
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  return [...new Set([...Object.keys(left), ...Object.keys(right)])].every((key) => sameJson(left[key], right[key]));
}

/** The data sheet's Apply on a chart without refs (SPEC-052): categories
 * and series replaced, `edited`. Null for a chart with refs (its data is
 * its cells), an unmodelled or read-only one, or unchanged data. */
export function setSheetChartData(model: SpreadsheetModel, sheetId: string, drawingId: string, categories: string[], series: ChartSeries[]): DrawingCommandResult | null {
  const found = findChart(model, sheetId, drawingId);
  if (!found || !isEditableChart(found.content) || found.content.refs !== undefined) {
    return null;
  }
  const { visual } = found.content;
  if (sameJson(categories, visual.categories) && sameJson(series, visual.series)) {
    return null;
  }
  return contentResult(found, { ...found.content, visual: { ...visual, categories: [...categories], series: series.map((each) => ({ ...each, values: [...each.values] })) }, edited: true });
}

/** Chart options' Apply: options replaced, `edited`. */
export function setSheetChartOptions(model: SpreadsheetModel, sheetId: string, drawingId: string, options: ChartOptions): DrawingCommandResult | null {
  const found = findChart(model, sheetId, drawingId);
  if (!found || !isEditableChart(found.content) || sameJson(options, found.content.visual.options)) {
    return null;
  }
  return contentResult(found, { ...found.content, visual: { ...found.content.visual, options: { ...options } }, edited: true });
}

/** Chart options' series colour: one series' colour, `edited`. */
export function setSheetChartSeriesColor(model: SpreadsheetModel, sheetId: string, drawingId: string, seriesIndex: number, color: ModelColor): DrawingCommandResult | null {
  const found = findChart(model, sheetId, drawingId);
  const target = found?.content.visual.series[seriesIndex];
  if (!found || !target || !isEditableChart(found.content) || sameJson(target.color, color)) {
    return null;
  }
  const series = found.content.visual.series.slice();
  series[seriesIndex] = { ...target, color: typeof color === "string" ? color : { ...color } };
  return contentResult(found, { ...found.content, visual: { ...found.content.visual, series }, edited: true });
}

/** Whether two series reference sets have the same shape ("Select Data"):
 * the same count and, per series, the same references present. */
const referencesPresent = (refs: SheetChartSeriesRefs): string => `${refs.name !== undefined}${refs.categories !== undefined}${refs.values !== undefined}`;

function sameReferenceShape(a: SheetChartRefs, b: SheetChartRefs): boolean {
  const present = referencesPresent;
  return a.series.length === b.series.length && a.series.every((each, i) => present(each) === present(b.series[i]!));
}

/** An ods chart frame's `draw:notify-on-update-of-ranges` synced with its
 * references: the drawing (same object when it already matches). */
function withSyncedNotify(model: Pick<SpreadsheetModel, "format">, drawing: SheetDrawing): SheetDrawing {
  const content = chartContentOf(drawing);
  if (model.format !== "ods" || drawing.sourceXml === undefined || !content || !isModelledChart(content) || !content.refs) {
    return drawing;
  }
  const sourceXml = withNotifyRanges(drawing.sourceXml, notifyRangesOf(content.refs));
  return sourceXml === drawing.sourceXml ? drawing : { ...drawing, sourceXml };
}

/** Select Data's Apply ("Select Data"): new refs. When the series count, the
 * chart type and the presence of every reference (name, categories, values)
 * per series are unchanged: `refsDirty` + refresh (the patch path);
 * otherwise the visual is rebuilt from the cells and the chart `edited`.
 * The chart keeps its own `plotVisibleOnly`; ods: the frame's notify ranges
 * follow (the sheet then `dirty`). */
export function setSheetChartSource(model: SpreadsheetModel, sheetId: string, drawingId: string, refs: SheetChartRefs): DrawingCommandResult | null {
  const found = findChart(model, sheetId, drawingId);
  if (!found || !isEditableChart(found.content)) {
    return null;
  }
  const current = found.content.refs;
  const nextRefs: SheetChartRefs = current ? { ...refs, plotVisibleOnly: current.plotVisibleOnly } : refs;
  if (current && sameJson(nextRefs, current)) {
    return null;
  }
  let content: SheetChartContent;
  if (current && sameReferenceShape(current, nextRefs)) {
    const patched: SheetChartContent = { ...found.content, refs: nextRefs, refsDirty: true };
    const refreshed = refreshedChartContent(model, sheetId, patched);
    content = refreshed === patched ? patched : { ...refreshed, cachesDirty: true };
  } else {
    const old = found.content.visual;
    const visual = visualOfRefs(model, sheetId, nextRefs, old.chartType, old.options, (i) => old.series[i]?.color ?? chartAccentRef(i));
    content = { ...found.content, visual, refs: nextRefs, edited: true };
  }
  const drawing = withSyncedNotify(model, { ...found.drawing, content });
  const sheet = withDrawingAt(found.sheet, found.index, drawing);
  if (drawing.sourceXml !== found.drawing.sourceXml) {
    sheet.dirty = true;
  }
  return { sheet, drawingId };
}

/** A point one row and one column further (offsets kept, clamped to the grid). */
function nextPoint(point: AnchorPoint): AnchorPoint {
  return { ...point, row: Math.min(point.row + 1, MAX_ROWS - 1), col: Math.min(point.col + 1, MAX_COLS - 1) };
}

/** An anchor offset by one row and one column: its points moved to the next
 * row and column; a position (posEmu) moved by the width of the column and
 * the height of the row it starts in. */
function offsetAnchor(anchor: DrawingAnchor, metrics: GridMetrics): DrawingAnchor {
  const next: DrawingAnchor = { ...anchor };
  if (anchor.from) {
    next.from = nextPoint(anchor.from);
  }
  if (anchor.to) {
    next.to = nextPoint(anchor.to);
  }
  if (anchor.posEmu) {
    const { x, y } = anchor.posEmu;
    const col = metrics.colAt(x / EMU_PER_PX);
    const row = metrics.rowAt(y / EMU_PER_PX);
    next.posEmu = { x: x + Math.round(metrics.colWidth(col) * EMU_PER_PX), y: y + Math.round(metrics.rowHeight(row) * EMU_PER_PX) };
  }
  return next;
}

function cloneVisual(visual: ChartVisual): ChartVisual {
  return {
    chartType: visual.chartType,
    categories: [...visual.categories],
    series: visual.series.map((each) => ({ name: each.name, values: [...each.values], color: typeof each.color === "string" ? each.color : { ...each.color } })),
    options: { ...visual.options },
  };
}

function cloneRefs(refs: SheetChartRefs): SheetChartRefs {
  return { ...refs, series: refs.series.map((each) => ({ ...each })) };
}

/** Duplicate chart: a new chart from the source's visual + refs (no
 * sourcePart, no sourceXml — its own new part, regenerated), offset one row
 * and one column, on top, selected. Null for an unmodelled or read-only
 * chart (a regenerated copy could not keep what makes it one) and on a
 * chart sheet (which holds exactly one drawing). */
export function duplicateSheetChart(model: SpreadsheetModel, sheetId: string, drawingId: string, metrics?: GridMetrics): DrawingCommandResult | null {
  const found = findChart(model, sheetId, drawingId);
  if (!found || found.sheet.kind !== "grid" || !isEditableChart(found.content)) {
    return null;
  }
  const { sheet, drawing, content } = found;
  const copy: SheetDrawing = {
    id: nextDrawingId(sheet),
    anchor: offsetAnchor(drawing.anchor, metrics ?? buildGridMetrics(sheet)),
    name: newChartName(sheet),
    content: content.refs ? { kind: "chart", visual: cloneVisual(content.visual), refs: cloneRefs(content.refs) } : { kind: "chart", visual: cloneVisual(content.visual) },
  };
  if (drawing.altText !== undefined) {
    copy.altText = drawing.altText;
  }
  return { sheet: withDrawings(model, sheet, [...sheet.drawings!, copy]), drawingId: copy.id };
}

// --- reference adjustment ----------------------------------------------------------------------

/** `refs` with every reference mapped: the same object when none changes. */
function mappedRefs(refs: SheetChartRefs, map: (text: string) => string): SheetChartRefs {
  let changed = false;
  const series = refs.series.map((each) => {
    const next: SheetChartSeriesRefs = { ...each };
    let seriesChanged = false;
    for (const key of ["name", "categories", "values"] as const) {
      const text = each[key];
      if (text !== undefined) {
        const mapped = map(text);
        if (mapped !== text) {
          next[key] = mapped;
          seriesChanged = true;
        }
      }
    }
    changed ||= seriesChanged;
    return seriesChanged ? next : each;
  });
  const plotArea = refs.plotArea === undefined ? undefined : map(refs.plotArea);
  if (!changed && plotArea === refs.plotArea) {
    return refs;
  }
  const next: SheetChartRefs = { ...refs, series };
  if (plotArea !== undefined) {
    next.plotArea = plotArea;
  }
  return next;
}

/** Every modelled chart's references mapped (host = its sheet); a chart
 * whose references changed gains `refsDirty`, and — with `refreshOn` —
 * a refreshable chart whose references changed or read one of those sheets
 * is refreshed (`cachesDirty` when its visual changed). Copy-on-write. */
function adjustCharts(model: SpreadsheetModel, map: (text: string, hostSheetId: string) => string, refreshOn?: readonly string[]): void {
  const signature = sheetSignature(model);
  let changed = false;
  const sheets = model.sheets.map((sheet) => {
    if (!sheet.drawings || sheet.drawings.length === 0) {
      return sheet;
    }
    let drawingsChanged = false;
    const drawings = sheet.drawings.map((drawing) => {
      const content = chartContentOf(drawing);
      if (!content || !isModelledChart(content) || !content.refs) {
        return drawing;
      }
      const refs = mappedRefs(content.refs, (text) => map(text, sheet.id));
      let next: SheetChartContent = refs === content.refs ? content : { ...content, refs, refsDirty: true };
      if (refreshOn && isRefreshable(next) && (refs !== content.refs || chartAreas(model, sheet.id, content.refs, signature).some((area) => refreshOn.includes(area.sheetId)))) {
        const refreshed = refreshedChartContent(model, sheet.id, next);
        if (refreshed !== next) {
          next = { ...refreshed, cachesDirty: true };
        }
      }
      if (next === content) {
        return drawing;
      }
      drawingsChanged = true;
      return { ...drawing, content: next };
    });
    if (!drawingsChanged) {
      return sheet;
    }
    changed = true;
    return { ...sheet, drawings };
  });
  if (changed) {
    model.sheets = sheets;
  }
}

/** Sheet id → name, for adjustFormula and repointFormula. */
function sheetNames(model: SpreadsheetModel): (sheetId: string) => string {
  const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
  return (sheetId) => names.get(sheetId) ?? "";
}

/** The `charts` REFERENCE_ADJUSTER ("Reference adjustment"): adjustFormula
 * over every modelled chart's refs on every sheet (host = the chart's
 * sheet), `refsDirty` where one changed, then a refresh (read-only charts
 * only adjust); ods: every chart frame's notify-on-update-of-ranges synced
 * (syncChartNotifyRanges). */
export function adjustChartsForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  const nameOf = sheetNames(model);
  adjustCharts(model, (text, hostSheetId) => adjustFormula(text, hostSheetId, edit, nameOf), [edit.sheetId]);
  syncChartNotifyRanges(model);
}

/** The `charts` MOVE_ADJUSTER: repointFormula over every modelled chart's
 * refs, `refsDirty` + refresh as above. The drawing itself stays put. */
export function moveChartsForMove(model: SpreadsheetModel, move: RangeMove): void {
  const nameOf = sheetNames(model);
  adjustCharts(model, (text, hostSheetId) => repointFormula(text, hostSheetId, move, nameOf), [move.fromSheetId, move.toSheetId]);
  syncChartNotifyRanges(model);
}

/** Sheet rename and delete (SPEC-121's commands call this with their
 * SheetRemap): remapSheetReferences over every chart's refs — a reference
 * into a deleted sheet becomes #REF! and its series keeps its caches —
 * `refsDirty` where one changed. No notify sync: a frame's notify list is
 * retained text in source-time names, which SPEC-121's write-time
 * sheet-name pass renames in a clean table (and the ods writer re-syncs a
 * rebuilt table's frames after it), so a rename leaves its sheet clean. */
export function remapChartReferences(model: SpreadsheetModel, remap: SheetRemap): void {
  adjustCharts(model, (text) => remapSheetReferences(text, remap));
}

/** ods: every chart frame whose `draw:notify-on-update-of-ranges` (in its
 * `sourceXml`) no longer matches its refs rewritten, its sheet marked
 * `dirty` (the frame lives in the table). A no-op for xlsx. */
export function syncChartNotifyRanges(model: SpreadsheetModel): void {
  if (model.format !== "ods") {
    return;
  }
  let changed = false;
  const sheets = model.sheets.map((sheet) => {
    if (!sheet.drawings || sheet.drawings.length === 0) {
      return sheet;
    }
    let drawingsChanged = false;
    const drawings = sheet.drawings.map((drawing) => {
      const next = withSyncedNotify(model, drawing);
      drawingsChanged ||= next !== drawing;
      return next;
    });
    if (!drawingsChanged) {
      return sheet;
    }
    changed = true;
    return { ...sheet, drawings, dirty: true as const };
  });
  if (changed) {
    model.sheets = sheets;
  }
}
