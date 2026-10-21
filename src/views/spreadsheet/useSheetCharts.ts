// The sheet charts' view wiring (SPEC-132 "View wiring", "Select Data",
// "Chart sheets", "Refresh"): which chart the toolbar acts on, Insert ▸
// Chart's gallery, the three dialogs, the rewrite notice and the refresh
// trigger; charts.ts owns the commands and the refresh itself.
// - The selected chart: useSheetDrawings' selection when it is a chart, or
//   a chart sheet's one chart (implicitly selected, never duplicated or
//   deleted as a drawing). One outside the catalogue or read-only (`locked`)
//   offers neither Edit data nor Chart options.
// - Every command is ONE SPEC-121 undo step on the chart's sheet: `record`
//   (which settles a running pass) before it, `commit` after, labelled
//   CHART_UNDO_LABELS' (Delete is SPEC-131's "Delete chart"); one that
//   changes nothing — an Insert ▸ Chart refused with a toast — records
//   nothing. Nothing runs while a cell is being edited (`busy`).
// - A dialog keeps the chart it opened on, so Select Data can point:
//   "Select cells" collapses it (`picking`) while the grid's selection, on
//   any sheet, fills its range; Done returns to the chart's sheet.
// - Refresh: each completed pass but an "open" one (open trusts caches)
//   refreshes the charts chartsAffectedBy names for its cells, plus those
//   chartsReadingSheet names for a sheet whose row or column visibility
//   changed (a per-sheet signature, recomputed only when its rows or columns
//   array changed). useRecalc reports a step's pass before the pipeline
//   writes through, so the refresh joins the triggering step's undo entry
//   (undo restores its snapshot) and never records one of its own.
import { computed, ref, shallowRef, triggerRef, watch, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { lostChartFeatures } from "../../lib/core/chartFeatures";
import { asChartType, SINGLE_SERIES_CHART_TYPES, type ChartOptions, type ChartSeries, type ChartType } from "../../lib/core/charts";
import { base64ToBytes, type OpaquePart } from "../../lib/core/packageParts";
import { DEFAULT_THEME, type ModelColor } from "../../lib/core/theme";
import { colToLetters, type GridRange } from "../../lib/spreadsheet/address";
import { CHART_UNDO_LABELS, chartContentOf, chartRangeOf, chartRefsFromRange, chartsAffectedBy, chartsReadingSheet, type ChartSourceFlags } from "../../lib/spreadsheet/charts";
import { duplicateSheetChart, insertSheetChart, refreshSheetChart, setSheetChartData, setSheetChartOptions, setSheetChartSeriesColor, setSheetChartSource } from "../../lib/spreadsheet/charts";
import { applyDrawingResult, type DrawingCommandResult } from "../../lib/spreadsheet/drawings";
import type { CellKey } from "../../lib/spreadsheet/formula/dependencies";
import { quoteSheetName } from "../../lib/spreadsheet/formula/print";
import type { RecalcReason } from "../../lib/spreadsheet/formula/recalc";
import type { Sheet, SheetChartContent, SheetDrawing, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { Recalc } from "./useRecalc";
import type { SheetDrawings } from "./useSheetDrawings";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

/** Select Data's Apply over a range that names no data. */
export const INVALID_CHART_RANGE_MESSAGE = "Enter a cell range on one sheet, such as Sheet1!$A$1:$D$5";
/** The flags a too-complex range opens with (Excel's defaults). */
export const DEFAULT_CHART_SOURCE_FLAGS: ChartSourceFlags = { seriesIn: "columns", firstRowNames: true, firstColCategories: true };

export interface SheetChartsDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeIndex: Ref<number>; activeSheet: Readonly<Ref<Sheet | null>>; sheetsChanged(): void };
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  recalc: Pick<Recalc, "onResult">;
  /** The grid selection a chart is built from (and Select Data points with). */
  selection: Pick<SheetSelectionController, "selection">;
  sheetMetrics: Pick<SheetMetrics, "metrics">;
  /** SPEC-131's layer: its selection, select, delete and where drawings go. */
  drawings: Pick<SheetDrawings, "selected" | "select" | "remove" | "canInsert">;
  /** Syncs the view state into the model before a step is recorded. */
  syncInto?: (model: SpreadsheetModel) => void;
  /** Whether a cell is being edited: charts are inert meanwhile. */
  busy?: () => boolean;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
}

/** The chart the toolbar acts on; `chartSheet`: a chart sheet's, implicitly
 * selected and never moved or deleted. */
export interface SelectedChart { sheetId: string; drawing: SheetDrawing; content: SheetChartContent; chartSheet: boolean } // prettier-ignore

interface DialogTarget { sheetId: string; drawingId: string } // prettier-ignore

/** The open dialog, on the chart it was opened for. */
export type SheetChartDialog =
  | (DialogTarget & { kind: "source"; chartType: ChartType; range: string | null; flags: ChartSourceFlags })
  | (DialogTarget & { kind: "dataSheet"; initial: { categories: string[]; series: ChartSeries[]; singleSeriesOnly: boolean } })
  | (DialogTarget & { kind: "options"; initial: ChartOptions & { series: { name: string; color: ModelColor }[]; hasAxes: boolean } });

/** A colour detached from any dialog draft (the model stays plain JSON). */
const plainColor = (color: ModelColor): ModelColor => (typeof color === "object" && color !== null ? { ...color } : color);

const absoluteRange = (range: GridRange): string => {
  const start = `$${colToLetters(range.left)}$${range.top + 1}`;
  return range.top === range.bottom && range.left === range.right ? start : `${start}:$${colToLetters(range.right)}$${range.bottom + 1}`;
};

/** A selection as a canonical, sheet-qualified, absolute reference — a
 * parenthesised union for several areas. */
export function selectionReference(sheetName: string, ranges: readonly GridRange[]): string {
  const areas = ranges.map((range) => `${quoteSheetName(sheetName)}!${absoluteRange(range)}`);
  return areas.length > 1 ? `(${areas.join(",")})` : (areas[0] ?? "");
}

/** The range field's text as chartRefsFromRange reads it: a leading "="
 * dropped, a comma-separated list made one union. */
export function sourceRangeText(text: string): string {
  const trimmed = text.trim().replace(/^=\s*/, "");
  return trimmed.includes(",") && !trimmed.startsWith("(") ? `(${trimmed})` : trimmed;
}

/** The rows hidden by the user or a filter and the hidden column spans. */
function visibilitySignature(sheet: Sheet): string {
  const rows = sheet.rows.filter((row) => row.hidden || row.filterHidden).map((row) => row.index);
  const columns = sheet.columns.filter((span) => span.hidden).map((span) => `${span.min}-${span.max}`);
  return `${rows.join(",")}|${columns.join(",")}`;
}

export function useSheetCharts(deps: SheetChartsDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const busy = (): boolean => deps.busy?.() === true;

  const selectedChart = computed((): SelectedChart | null => {
    const sheet = deps.book.activeSheet.value;
    const drawing = sheet?.kind === "chart" ? sheet.drawings?.find((entry) => entry.content.kind === "chart") : deps.drawings.selected.value?.drawing;
    const content = drawing ? chartContentOf(drawing) : undefined;
    return sheet && drawing && content ? { sheetId: sheet.id, drawing, content, chartSheet: sheet.kind === "chart" } : null;
  });
  /** Outside the catalogue: painted as paintChart's placeholder, never edited. */
  const isUnsupported = computed(() => selectedChart.value !== null && selectedChart.value.content.visual.unsupportedChartXml !== undefined);
  /** Why a read-only chart cannot be edited ("" for none). */
  const reason = computed(() => selectedChart.value?.content.readOnlyReason ?? "");
  /** Edit data and Chart options are off. */
  const locked = computed(() => isUnsupported.value || reason.value !== "");
  const theme = computed(() => model()?.theme ?? DEFAULT_THEME);

  /** One command as ONE undo step on `sheetId`; null: nothing ran. */
  function run(label: string, sheetId: string, command: (m: SpreadsheetModel) => DrawingCommandResult | null): DrawingCommandResult | null {
    const m = model();
    if (!m || busy()) {
      return null;
    }
    deps.syncInto?.(m);
    deps.undo.record(label, { sheetId });
    const result = command(m);
    if (result) {
      applyDrawingResult(m, result);
    }
    deps.undo.commit();
    return result;
  }

  /** Selects a chart again on its grid sheet when that sheet is on show. */
  function reselect(target: DialogTarget): void {
    const sheet = deps.book.activeSheet.value;
    if (sheet?.kind === "grid" && sheet.id === target.sheetId) {
      deps.drawings.select(target.drawingId);
    }
  }

  // --- Insert ▸ Chart ----------------------------------------------------------------------
  const galleryOpen = ref(false);
  const canInsert = computed(() => deps.drawings.canInsert.value);

  const openGallery = (): void => void (galleryOpen.value = canInsert.value && !busy());
  const closeGallery = (): void => void (galleryOpen.value = false);

  /** A gallery pick: the chart built from the selection, selected — or the
   * refusal's toast, nothing inserted. */
  function pickChartType(chartType: string): void {
    galleryOpen.value = false;
    const sheet = deps.book.activeSheet.value;
    if (!sheet || sheet.kind !== "grid" || !canInsert.value) {
      return;
    }
    const refusal = { message: "" };
    const result = run(CHART_UNDO_LABELS.insert, sheet.id, (m) => {
      const built = insertSheetChart(m, sheet.id, deps.selection.selection.value, asChartType(chartType), deps.sheetMetrics.metrics.value);
      if ("message" in built) {
        refusal.message = built.message;
        return null;
      }
      return built;
    });
    if (refusal.message) {
      toast(refusal.message);
    } else if (result?.drawingId) {
      deps.drawings.select(result.drawingId);
    }
  }

  // --- the dialogs ---------------------------------------------------------------------------
  const dialog = shallowRef<SheetChartDialog | null>(null);
  /** Select Data's inline error ("" for none). */
  const sourceError = ref("");
  /** Select Data is collapsed while the grid's selection fills its range. */
  const picking = ref(false);

  function closeDialogs(): void {
    dialog.value = null;
    picking.value = false;
    sourceError.value = "";
  }

  /** The toolbar's Edit data: Select Data on a chart with refs, else the data sheet. */
  function openData(): void {
    const target = selectedChart.value;
    if (!target || locked.value || busy()) {
      return;
    }
    const { visual, refs } = target.content;
    const at = { sheetId: target.sheetId, drawingId: target.drawing.id };
    sourceError.value = "";
    picking.value = false;
    if (refs) {
      const shown = chartRangeOf(refs, visual.chartType);
      dialog.value = { kind: "source", ...at, chartType: visual.chartType, range: shown?.range ?? null, flags: { ...(shown?.flags ?? DEFAULT_CHART_SOURCE_FLAGS) } };
      return;
    }
    const series = visual.series.map((entry) => ({ name: entry.name, color: entry.color, values: entry.values.slice() }));
    dialog.value = { kind: "dataSheet", ...at, initial: { categories: visual.categories.slice(), series, singleSeriesOnly: SINGLE_SERIES_CHART_TYPES.has(visual.chartType) } };
  }

  /** The toolbar's Chart options. */
  function openOptions(): void {
    const target = selectedChart.value;
    if (!target || locked.value || busy()) {
      return;
    }
    const { visual } = target.content;
    const series = visual.series.map((entry) => ({ name: entry.name, color: entry.color }));
    const hasAxes = visual.chartType !== "pie" && visual.chartType !== "doughnut";
    dialog.value = { kind: "options", sheetId: target.sheetId, drawingId: target.drawing.id, initial: { ...visual.options, series, hasAxes } };
  }

  /** A dialog's command as ONE step on the chart it opened on — closing it
   * (and reselecting the chart) unless it stays open (a live colour). */
  function commitDialog(kind: SheetChartDialog["kind"], label: string, command: (m: SpreadsheetModel, at: DialogTarget) => DrawingCommandResult | null, close = true): void {
    const open = dialog.value;
    if (open?.kind === kind) {
      dialog.value = close ? null : open;
      run(label, open.sheetId, (m) => command(m, open));
      if (close) {
        reselect(open);
      }
    }
  }

  /** The data sheet's Done: "Edit chart data". */
  const applyData = (categories: string[], series: ChartSeries[]): void => {
    const plain = series.map((entry) => ({ name: String(entry.name), values: entry.values.slice(), color: plainColor(entry.color) }));
    commitDialog("dataSheet", CHART_UNDO_LABELS.data, (m, at) => setSheetChartData(m, at.sheetId, at.drawingId, categories.slice(), plain));
  };
  /** Chart options' Apply: "Chart options". */
  const applyOptions = (options: ChartOptions): void =>
    commitDialog("options", CHART_UNDO_LABELS.options, (m, at) => setSheetChartOptions(m, at.sheetId, at.drawingId, { ...options }));
  /** Chart options' live series colour: "Chart series colour". */
  const applySeriesColor = (index: number, color: ModelColor): void =>
    commitDialog("options", CHART_UNDO_LABELS.seriesColor, (m, at) => setSheetChartSeriesColor(m, at.sheetId, at.drawingId, index, plainColor(color)), false);

  /** Select Data's Apply: "Chart data range" — false (and the inline error)
   * when the range names no data, nothing applied. */
  function applySource(range: string, flags: ChartSourceFlags): boolean {
    const open = dialog.value;
    const m = model();
    if (open?.kind !== "source" || !m) {
      return false;
    }
    const text = sourceRangeText(range);
    const refs = text ? chartRefsFromRange(m, open.sheetId, text, { ...flags }, open.chartType) : null;
    if (!refs) {
      sourceError.value = INVALID_CHART_RANGE_MESSAGE;
      return false;
    }
    closeDialogs();
    run(CHART_UNDO_LABELS.source, open.sheetId, (current) => setSheetChartSource(current, open.sheetId, open.drawingId, refs));
    reselect(open);
    return true;
  }

  /** The grid selection on show as Select Data's range ("" off a grid). */
  const selectionText = computed(() => {
    const sheet = deps.book.activeSheet.value;
    return sheet?.kind === "grid" ? selectionReference(sheet.name, deps.selection.selection.value.ranges) : "";
  });

  function startPick(): void {
    if (dialog.value?.kind === "source") {
      picking.value = true;
    }
  }

  /** Picking done (its range already taken by the dialog): back to the chart. */
  function endPick(): void {
    const open = dialog.value;
    picking.value = false;
    const index = open ? (model()?.sheets.findIndex((sheet) => sheet.id === open.sheetId) ?? -1) : -1;
    if (open && index >= 0) {
      deps.book.activeIndex.value = index;
      reselect(open);
    }
  }

  // --- duplicate and delete ------------------------------------------------------------------
  function duplicateChart(): void {
    const target = selectedChart.value;
    if (!target || target.chartSheet) {
      return;
    }
    const result = run(CHART_UNDO_LABELS.duplicate, target.sheetId, (m) => duplicateSheetChart(m, target.sheetId, target.drawing.id, deps.sheetMetrics.metrics.value));
    if (result?.drawingId) {
      deps.drawings.select(result.drawingId);
    }
  }

  /** SPEC-131's delete ("Delete chart"). */
  function deleteChart(): void {
    const target = selectedChart.value;
    if (target && !target.chartSheet) {
      deps.drawings.remove(target.drawing.id);
    }
  }

  // --- the rewrite notice ------------------------------------------------------------------
  const partTexts = new WeakMap<OpaquePart, string>();

  /** What regenerating an edited chart's part cost ("" when nothing). */
  const rewriteNotice = computed(() => {
    const m = model();
    const content = selectedChart.value?.content;
    const part = content?.edited && content.sourcePart ? m?.opaque.find((entry) => entry.name === content.sourcePart) : undefined;
    if (!m || !part) {
      return "";
    }
    const xml = partTexts.get(part) ?? new TextDecoder("utf-8").decode(base64ToBytes(part.data));
    partTexts.set(part, xml);
    const format = m.format === "ods" ? "odf" : "ooxml";
    const lost = lostChartFeatures(xml, format);
    if (format === "ooxml" && /<c:externalData[\s/>]/.test(xml)) {
      lost.push("its embedded data workbook");
    }
    return lost.length > 0 ? `Chart rebuilt from your edits — no longer includes ${lost.join("; ")}.` : "";
  });

  // --- refresh ---------------------------------------------------------------------------------
  const visibility = new Map<string, { rows: Sheet["rows"]; columns: Sheet["columns"]; signature: string }>();
  let visibilityOf: SpreadsheetModel | null = null;

  /** The grid sheets whose row or column visibility changed since the last look. */
  function visibilityChanges(m: SpreadsheetModel): string[] {
    if (visibilityOf !== m) {
      visibility.clear();
      visibilityOf = m;
    }
    const changed: string[] = [];
    for (const sheet of m.sheets) {
      const seen = visibility.get(sheet.id);
      if (sheet.kind !== "grid" || (seen && seen.rows === sheet.rows && seen.columns === sheet.columns)) {
        continue;
      }
      const signature = visibilitySignature(sheet);
      if (seen && seen.signature !== signature) {
        changed.push(sheet.id);
      }
      visibility.set(sheet.id, { rows: sheet.rows, columns: sheet.columns, signature });
    }
    return changed;
  }

  /** One completed pass's refresh (see the module header). */
  function refreshAfter(changed: readonly CellKey[], passReason?: RecalcReason): void {
    const m = model();
    const hiddenChanged = m ? visibilityChanges(m) : [];
    if (!m || passReason === "open") {
      return;
    }
    const targets = new Map<string, DialogTarget>();
    const add = (target: DialogTarget) => targets.set(`${target.sheetId}\n${target.drawingId}`, target);
    chartsAffectedBy(m, changed).forEach(add);
    hiddenChanged.forEach((sheetId) => chartsReadingSheet(m, sheetId).forEach(add));
    let refreshed = false;
    for (const target of targets.values()) {
      refreshed = refreshSheetChart(m, target.sheetId, target.drawingId) || refreshed;
    }
    if (refreshed) {
      triggerRef(deps.book.workbook);
      deps.book.sheetsChanged();
    }
  }

  deps.recalc.onResult(refreshAfter);
  // The signatures start from the workbook as opened.
  watch(() => deps.book.workbook.value, (m) => m && visibilityChanges(m), { immediate: true });

  // prettier-ignore
  return {
    selectedChart, isUnsupported, reason, locked, theme, canInsert, galleryOpen, dialog, sourceError, picking, selectionText, rewriteNotice,
    openGallery, closeGallery, pickChartType, openData, openOptions, closeDialogs, applyData, applyOptions, applySeriesColor, applySource,
    startPick, endPick, duplicateChart, deleteChart, refreshAfter,
  };
}

export type SheetCharts = ReturnType<typeof useSheetCharts>;

/** Narrows a component's runtime-typed `charts` prop (the .vue parser bars
 * casts). */
export function asSheetCharts(value: unknown): SheetCharts {
  return value as SheetCharts;
}
