// SheetGrid's wiring (SPEC-117): one sheet's viewport, the paper's render
// items, the overlays, and the keyboard and pointer handlers — assembled
// from the composables that own each (useSheetViewport, useSheetCells,
// useSheetOverlays, useSheetPointer) over the view-level state the grid is
// handed (useSheetMetrics, useSheetSelection, useSheetZoom,
// useSheetViewState). SheetGrid.vue binds what this returns; .vue scripts
// cannot carry TS syntax (the espree constraint), so the narrowing and the
// event logic live here. SPEC-119: in Show formulas mode the paper shows
// formula text (useSheetCells). SPEC-122: the in-cell editor (CellEditor)
// sits over the active cell and holds the grid's focus — keys reach this
// handler from it once the editor's key table has passed them on, and
// pointer presses go through useCellEditing first while a cell is being
// edited (pointing, or committing before the selection moves); a double-click
// opens Edit mode at the clicked character; the editing composable's moves
// are revealed here; the references of a formula being entered are outlined
// (useSheetOverlays' refHighlights).
import { computed, watch, type Ref } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { mergeAt } from "../../lib/spreadsheet/merges";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { cellEditorView } from "./cellEditorView";
import type { CellEditing } from "./useCellEditing";
import type { FormatPainter } from "./useFormatPainter";
import type { Formatting } from "./useFormatting";
import type { FormulaDisplay } from "./useFormulaDisplay";
import type { Recalc } from "./useRecalc";
import type { SheetMetrics } from "./useSheetMetrics";
import { useSheetCells } from "./useSheetCells";
import { useSheetOverlays } from "./useSheetOverlays";
import { useSheetPointer } from "./useSheetPointer";
import type { SheetSelectionController } from "./useSheetSelection";
import { COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX, useSheetViewport } from "./useSheetViewport";
import type { SheetViewState } from "./useSheetViewState";
import type { SheetZoom } from "./useSheetZoom";

/** What the view hands the grid: the view-level state of the sheet on show. */
export interface SheetGridContext {
  model: SpreadsheetModel;
  sheetMetrics: SheetMetrics;
  selection: SheetSelectionController;
  zoom: SheetZoom;
  viewState: SheetViewState;
  /** The number locale values render in (useSpreadsheetLocale, SPEC-118);
   * absent = the default locale. */
  locale?: SpreadsheetLocale;
  /** The formula bar's text and View ▸ Show formulas (SPEC-119). */
  formulas: FormulaDisplay;
  /** Recalculation's stale cells (SPEC-120); absent = none. */
  recalc?: Pick<Recalc, "staleTitleAt">;
  /** Cell entry (SPEC-122); absent = the grid is read-only. */
  cellEditing?: CellEditing;
  /** SPEC-123: formatting (the cell context menu and Format Cells) and the
   * format painter (the grid's data-format-painter); absent = none. */
  formatting?: Formatting;
  painter?: Pick<FormatPainter, "mode">;
}

/** Narrows SheetGrid's runtime-typed `context` prop (the .vue parser bars casts). */
export function asSheetGridContext(value: unknown): SheetGridContext {
  return value as SheetGridContext;
}

export function useSheetGrid(
  scroller: Readonly<Ref<HTMLElement | null>>,
  sheet: () => Sheet,
  context: () => SheetGridContext,
  onCycleSheet: (delta: number) => void,
  /** Show formulas mode (SPEC-119) — SheetGrid's prop. */
  showFormulas: () => boolean = () => false,
) {
  const metrics = () => context().sheetMetrics.metrics.value;
  const merges = () => context().sheetMetrics.merges.value;
  const selection = () => context().selection.selection.value;

  const viewport = useSheetViewport(scroller, {
    sheet,
    metrics,
    frozen: () => context().sheetMetrics.frozen.value,
    zoom: () => context().zoom.percent.value / 100,
    initialTopLeft: () => context().viewState.active.value.topLeftCell ?? sheet().topLeftCell,
    onTopLeftChange: (cell) => context().viewState.updateActive({ topLeftCell: cell }),
  });
  const cells = useSheetCells({ model: () => context().model, sheet, metrics, merges, selection, viewport, locale: () => context().locale, showFormulas });
  const editing = (): CellEditing | undefined => context().cellEditing;
  const overlays = useSheetOverlays({ sheet, metrics, merges, selection, viewport, highlights: () => editing()?.point.highlights.value ?? [] });
  /** Focus goes to the in-cell editor (the grid itself without one). */
  function focusEditor(): void {
    const host = scroller.value?.querySelector<HTMLElement>('[data-testid="cell-editor"]');
    (host ?? scroller.value)?.focus({ preventScroll: true });
  }
  const pointer = useSheetPointer({ scroller, metrics, viewport, selection: () => context().selection, editing, focus: focusEditor });

  /** The in-cell editor's box, text style and name (SPEC-122). */
  const editor = computed(() => {
    const ctx = context();
    const active = selection().active;
    const style = ctx.selection.styleAt(sheet(), active.row, active.col) ?? {};
    return cellEditorView({
      sheet: sheet(),
      active,
      style,
      rect: overlays.activeRect.value,
      room: viewport.viewCols.value,
      theme: ctx.model.theme ?? DEFAULT_THEME,
      format: ctx.selection.formatContext.value,
    });
  });
  // A move the editor made (a commit's Enter/Tab, Ready's Enter/Tab) brings
  // the new active cell into view.
  watch(
    () => editing()?.revealTick.value,
    () => viewport.ensureVisible(focusTarget(selection().active)),
  );

  /** A double-click on a cell opens Edit mode, caret at the clicked
   * character (SPEC-122). */
  function onDblClick(event: MouseEvent): void {
    const cellEditing = editing();
    if (!cellEditing || cellEditing.editing.value || event.button !== 0) {
      return;
    }
    viewport.ensureVisible(focusTarget(selection().active));
    if (cellEditing.openEdit({ point: { x: event.clientX, y: event.clientY } })) {
      focusEditor();
    }
  }

  const showGridLines = computed(() => context().viewState.active.value.showGridLines !== false);
  /** A stale cell's title on this sheet (SPEC-120), for the paper. */
  const staleTitle = (row: number, col: number): string | undefined => context().recalc?.staleTitleAt(sheet().id, row, col);
  /** The id of the active cell's element (a merge's anchor), for
   * aria-activedescendant. */
  const activeDescendant = computed(() => {
    const active = selection().active;
    return `sheet-cell-${cellAddress(active)}`;
  });

  function cellAddress(at: CellRef): string {
    const merge = mergeAt(merges(), at.row, at.col);
    const anchor = merge ? { row: merge.top, col: merge.left } : at;
    let letters = "";
    for (let n = anchor.col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
      letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
    }
    return `${letters}${anchor.row + 1}`;
  }

  /** The cell a keyboard move should bring into view: the active cell when
   * it moved, else the corner of the active range farthest from it (the
   * corner Shift+arrows move). */
  function focusTarget(before: CellRef): GridRange {
    const sel = selection();
    const active = sel.active;
    let target: CellRef = active;
    if (active.row === before.row && active.col === before.col) {
      const range = sel.ranges[sel.activeRangeIndex] ?? { top: active.row, left: active.col, bottom: active.row, right: active.col };
      target = {
        row: range.top === active.row ? range.bottom : range.top,
        col: range.left === active.col ? range.right : range.left,
      };
    }
    const merge = mergeAt(merges(), target.row, target.col);
    return merge ?? { top: target.row, left: target.col, bottom: target.row, right: target.col };
  }

  function onKeydown(event: KeyboardEvent): void {
    // A key on the grid element itself (not through the editor, whose own
    // handler ran first) still meets the editor's key table.
    if (event.target === scroller.value && editing()?.onKeydown(event)) {
      return;
    }
    if (event.ctrlKey && !event.altKey && !event.metaKey && !event.shiftKey && (event.key === "PageDown" || event.key === "PageUp")) {
      event.preventDefault();
      onCycleSheet(event.key === "PageDown" ? 1 : -1);
      return;
    }
    const before = selection().active;
    if (context().selection.handleKey(event, viewport.page.value)) {
      event.preventDefault();
      viewport.ensureVisible(focusTarget(before));
    }
  }

  function onWheel(event: WheelEvent): void {
    context().zoom.onWheel(event);
  }

  /** Style objects for the scaled content box and the scroll spacer. */
  const contentStyle = computed(() => ({
    transform: `scale(${viewport.z.value})`,
    width: `${viewport.viewW.value}px`,
    height: `${viewport.viewH.value}px`,
  }));
  const areaStyle = computed(() => ({
    width: `${viewport.canvasWidth.value + ROW_HEADER_WIDTH_PX * viewport.z.value}px`,
    height: `${viewport.canvasHeight.value + COLUMN_HEADER_HEIGHT_PX * viewport.z.value}px`,
  }));
  const canvasStyle = computed(() => ({
    left: `${ROW_HEADER_WIDTH_PX * viewport.z.value}px`,
    top: `${COLUMN_HEADER_HEIGHT_PX * viewport.z.value}px`,
    width: `${viewport.canvasWidth.value}px`,
    height: `${viewport.canvasHeight.value}px`,
  }));

  return {
    viewport,
    bands: cells.bands,
    overlays,
    pointer,
    editor,
    onDblClick,
    focusEditor,
    showGridLines,
    staleTitle,
    activeDescendant,
    onKeydown,
    onWheel,
    contentStyle,
    areaStyle,
    canvasStyle,
  };
}
