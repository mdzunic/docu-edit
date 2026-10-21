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
// (useSheetOverlays' refHighlights). SPEC-124: header-boundary resizing
// (useHeaderResize, through the pointer) and a double-click on a boundary
// autofitting; the grid binds the resize guide, tooltip and cursor.
// SPEC-125: the marching ants on the copied range, the fill handle on a
// single-range selection's bottom-right corner (a double-click on it fills
// down), the fill and drag-move previews (the move's with its destination's
// address), the fill options chip and the move cursor (`transfer`); the
// pointer hands the handle and the outline's presses to useFillHandle and
// useRangeMove. SPEC-130: the data-validation chrome over the active cell
// (`validationChrome` — the input message and the dropdown, placed from the
// active cell's rect, with the cell area's size to stay inside), the input
// message's id as the active cell's aria-describedby (`describedBy`), and
// Alt+Down opening the dropdown's list from the grid (useDataValidation's
// onGridKey, before SPEC-117's navigation). SPEC-131: the drawing layer
// (useDrawingGestures over the context's `drawings`) — its panes for SheetGrid,
// and its move, resize and draw gestures ahead of the pointer's own handlers.
import { computed, watch, type Ref } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { mergeAt } from "../../lib/spreadsheet/merges";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { cellEditorView } from "./cellEditorView";
import type { CellEditing } from "./useCellEditing";
import type { DataValidationView } from "./useDataValidation";
import { useDrawingGestures } from "./useDrawingGestures";
import type { SheetDrawings } from "./useSheetDrawings";
import type { ConditionalFormatsView } from "./useConditionalFormats";
import type { FunctionAssist } from "./useFunctionAssist";
import type { FillHandle } from "./useFillHandle";
import type { FormatPainter } from "./useFormatPainter";
import type { HeaderResize } from "./useHeaderResize";
import type { Merges } from "./useMerges";
import type { Structure } from "./useStructure";
import type { Formatting } from "./useFormatting";
import type { FormulaDisplay } from "./useFormulaDisplay";
import type { RangeMoveController } from "./useRangeMove";
import type { Recalc } from "./useRecalc";
import type { SheetClipboard } from "./useSheetClipboard";
import type { SheetMetrics } from "./useSheetMetrics";
import { useSheetCells } from "./useSheetCells";
import { useSheetOverlays } from "./useSheetOverlays";
import { useSheetPointer } from "./useSheetPointer";
import type { SheetSelectionController } from "./useSheetSelection";
import { COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX, useSheetViewport } from "./useSheetViewport";
import type { SheetViewState } from "./useSheetViewState";
import type { SheetZoom } from "./useSheetZoom";
import { VALIDATION_INPUT_MESSAGE_ID } from "./validationChrome";

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
  /** SPEC-129: each mounted cell's conditional-format result; absent = none. */
  conditionalFormats?: Pick<ConditionalFormatsView, "resultAt">;
  /** SPEC-127's function assists: the in-cell editor's ARIA. */
  assist?: Pick<FunctionAssist, "editorAria">;
  /** SPEC-123: formatting (the cell context menu and Format Cells) and the
   * format painter (the grid's data-format-painter); absent = none. */
  formatting?: Formatting;
  painter?: Pick<FormatPainter, "mode">;
  /** SPEC-124: rows and columns (the header menu and the dialogs), header
   * resizing (the cursor, guide, tooltip and drags) and merges; absent =
   * none. */
  structure?: Structure;
  headerResize?: HeaderResize;
  merges?: Merges;
  /** SPEC-125: the clipboard (the marching ants, its dialog), the fill
   * handle and its chip, drag-move; absent = none. */
  clipboard?: SheetClipboard;
  fillHandle?: FillHandle;
  rangeMove?: RangeMoveController;
  /** SPEC-130: data validation's input message and dropdown; absent = none. */
  validation?: DataValidationView;
  /** SPEC-131: the sheet's drawing layer; absent = none painted. */
  drawings?: SheetDrawings;
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
  const cells = useSheetCells({
    model: () => context().model,
    sheet,
    metrics,
    merges,
    selection,
    viewport,
    locale: () => context().locale,
    showFormulas,
    conditional: (row, col) => context().conditionalFormats?.resultAt(row, col),
  });
  const editing = (): CellEditing | undefined => context().cellEditing;
  const overlays = useSheetOverlays({ sheet, metrics, merges, selection, viewport, highlights: () => editing()?.point.highlights.value ?? [] });
  /** Focus goes to the in-cell editor (the grid itself without one). */
  function focusEditor(): void {
    const host = scroller.value?.querySelector<HTMLElement>('[data-testid="cell-editor"]');
    (host ?? scroller.value)?.focus({ preventScroll: true });
  }
  /** A single-range selection's outline (drag-move's band, the handle's
   * corner; SPEC-125). */
  const outline = computed(() => {
    const ranges = selection().ranges;
    return ranges.length === 1 ? overlays.rectOf(ranges[0]!, "outline") : null;
  });
  const pointer = useSheetPointer({
    scroller,
    metrics,
    viewport,
    selection: () => context().selection,
    editing,
    focus: focusEditor,
    resize: () => context().headerResize,
    fill: () => context().fillHandle,
    move: () => context().rangeMove,
    outline: () => outline.value,
    painting: () => Boolean(context().painter?.mode.value),
  });
  const drawingGestures = useDrawingGestures({ scroller, viewport, metrics, drawings: () => context().drawings, focusGrid: focusEditor });

  /** SPEC-125's overlays, in paper px (see the module header). */
  const transfer = computed(() => {
    const ctx = context();
    const id = sheet().id;
    const marquee = ctx.clipboard?.marquee.value;
    const ants = marquee && marquee.sheetId === id ? overlays.rectOf(marquee.range, "marquee") : null;
    const corner = ctx.fillHandle?.seed.value && !ctx.rangeMove?.drag.value ? outline.value : null;
    const fill = ctx.fillHandle?.preview.value;
    const move = ctx.rangeMove?.preview.value;
    const moveRect = move ? overlays.rectOf(move.range, "move") : null;
    const chip = ctx.fillHandle?.chip.value;
    const chipRect = chip && chip.sheetId === id && !ctx.fillHandle?.drag.value ? overlays.rectOf(chip.target, "chip") : null;
    return {
      ants: ants && marquee ? { ...ants, mode: marquee.mode } : null,
      handle: corner ? { left: corner.left + corner.width - 4, top: corner.top + corner.height - 4 } : null,
      fillPreview: fill ? overlays.rectOf(fill, "fill") : null,
      movePreview: moveRect && move ? { ...moveRect, address: move.address } : null,
      chip: chipRect && chip ? { x: chipRect.left + chipRect.width, y: chipRect.top + chipRect.height, applied: chip.applied } : null,
      cursor: ctx.rangeMove?.hover.value || ctx.rangeMove?.drag.value ? "move" : "",
    };
  });

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
    // SPEC-125: a double-click on the fill handle fills down.
    if (event.button === 0 && pointer.pressedOnHandle()) {
      context().fillHandle?.doubleClick();
      return;
    }
    // SPEC-124: a double-click on a header boundary autofits.
    const boundary = event.button === 0 ? pointer.boundaryAtClient(event.clientX, event.clientY) : null;
    if (boundary) {
      context().headerResize?.autofit(boundary.axis, boundary.index);
      return;
    }
    const cellEditing = editing();
    if (!cellEditing || cellEditing.editing.value || event.button !== 0) {
      return;
    }
    viewport.ensureVisible(focusTarget(selection().active));
    if (cellEditing.openEdit({ point: { x: event.clientX, y: event.clientY } })) {
      focusEditor();
    }
  }

  /** SPEC-130's chrome over the active cell (null: none shows): the cell's
   * rect and the cell area's size (paper px), the input message, the
   * dropdown and its open list. */
  const validationChrome = computed(() => {
    const validation = context().validation;
    const rect = overlays.activeRect.value;
    const message = validation?.inputMessage.value ?? null;
    const dropdown = validation?.dropdown.value ?? null;
    if (!validation || !rect || (!message && !dropdown)) {
      return null;
    }
    return {
      rect,
      bounds: { width: viewport.viewCols.value, height: viewport.viewRows.value },
      message,
      dropdown,
      list: validation.listView.value,
    };
  });
  /** The active cell's aria-describedby while the input message shows. */
  const describedBy = computed(() => (validationChrome.value?.message ? { ref: activeDescendant.value.slice("sheet-cell-".length), id: VALIDATION_INPUT_MESSAGE_ID } : null));

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
    // SPEC-130: Alt+Down opens the active cell's dropdown list.
    if (context().validation?.onGridKey(event)) {
      event.preventDefault();
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
    /** The grid element's pointer handlers: SPEC-131's drawing gestures first. */
    pointer: { ...pointer, ...drawingGestures.wrap(pointer) },
    drawingGestures,
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
    /** SPEC-124: header resizing (guide, tooltip, cursor), absent = none. */
    resize: computed(() => context().headerResize ?? null),
    transfer,
    /** SPEC-125: the fill handle the chip re-applies through; null = none. */
    fillChip: computed(() => context().fillHandle ?? null),
    validationChrome,
    describedBy,
    /** SPEC-130: data validation (null: not wired). */
    validation: computed(() => context().validation ?? null),
  };
}
