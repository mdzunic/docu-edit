// The grid's viewport (SPEC-116 "The read-only grid", SPEC-117 "Geometry",
// "Frozen panes and hidden rows and columns", "Zoom"): the scroll position →
// the windows of rows and columns SheetGrid renders (the viewport plus
// GRID_OVERSCAN_ROWS/COLS), the scroll extent its spacer is sized to, and
// where each row and column sits. The grid is DOM, not canvas — a visible
// cell needs real text layout, an accessibility node and a selector a test
// can find — and this composable is what bounds how many exist.
//
// - Geometry is GridMetrics (unscaled px). Zoom z is a view transform: the
//   grid scales one content box by z, the spacer is sized to total × z, and
//   every window here is computed from scroll / z.
// - The content box does not scroll: it is a sticky, viewport-sized box, and
//   everything inside is placed relative to it — the frozen rows and columns
//   (and their corner) at fixed positions, the rest offset by the scroll
//   position, whose window starts at the first unfrozen row and column. A
//   frozen count beyond the extent is clamped for display only.
// - The extent is max(usedRange.bottom + 100, 1000) rows by
//   max(usedRange.right + 10, 26) columns; scrolling within one screen of the
//   end grows it by the same margins, up to MAX_ROWS/MAX_COLS.
// - Above MAX_SCROLL_EXTENT_PX (under every target browser's maximum element
//   size) the spacer stops growing and scroll offsets map onto the sheet
//   proportionally, so a used range at row 1,000,000 is reachable everywhere.
// - SPEC-137: slots, the ring a far scroll step turns, and the viewport
//   alone while the grid settles from one (`overscanDrawn`) — sheetAxis.ts's
//   header.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch, type Ref } from "vue";
import { usedRange, type CellRef } from "../../lib/spreadsheet/address";
import { defaultColumnPx, defaultRowPx, type GridMetrics } from "../../lib/spreadsheet/geometry";
import { MAX_COLS, MAX_ROWS, type Sheet } from "../../lib/spreadsheet/model";
import { axisWindow, mapAxis, revealOffset, ScrollSettler, scrollFor } from "./sheetAxis";
import type { AxisWindow } from "./sheetRenderPlan";

export { MAX_SCROLL_EXTENT_PX, SETTLE_MS, type AxisItem } from "./sheetAxis";

export const GRID_OVERSCAN_ROWS = 8;
export const GRID_OVERSCAN_COLS = 4;
/** The row-number column (wide enough for "1048576") and the letter row,
 * unscaled — they scale with the zoom like everything in the content box. */
export const ROW_HEADER_WIDTH_PX = 56;
export const COLUMN_HEADER_HEIGHT_PX = 20;

const EXTENT_ROW_MARGIN = 100;
const MIN_EXTENT_ROWS = 1000;
const EXTENT_COL_MARGIN = 10;
const MIN_EXTENT_COLS = 26;

/** What an unmeasured grid (jsdom, a hidden tab) assumes it shows — a
 * desktop window, so the rendered window is never empty. */
const FALLBACK_WIDTH_PX = 1280;
const FALLBACK_HEIGHT_PX = 720;

/** Narrows the runtime-typed `sheet` prop (the .vue parser bars casts). */
export function asSheet(value: unknown): Sheet {
  return value as Sheet;
}

export interface SheetViewportDeps {
  sheet: () => Sheet;
  metrics: () => GridMetrics;
  frozen: () => { rows: number; cols: number };
  /** The zoom as a fraction (1 = 100%). */
  zoom: () => number;
  /** Where the sheet opens scrolled to (the file's or the stored top-left). */
  initialTopLeft: () => CellRef | undefined;
  /** Reported when the scroll position settles on a new top-left cell. */
  onTopLeftChange?: (cell: CellRef) => void;
}

export function useSheetViewport(scroller: Readonly<Ref<HTMLElement | null>>, deps: SheetViewportDeps) {
  const used = usedRange(deps.sheet());
  const start = deps.initialTopLeft();
  const extentRows = ref(
    Math.min(MAX_ROWS, Math.max((used?.bottom ?? 0) + EXTENT_ROW_MARGIN, MIN_EXTENT_ROWS, (start?.row ?? 0) + EXTENT_ROW_MARGIN)),
  );
  const extentCols = ref(
    Math.min(MAX_COLS, Math.max((used?.right ?? 0) + EXTENT_COL_MARGIN, MIN_EXTENT_COLS, (start?.col ?? 0) + EXTENT_COL_MARGIN)),
  );
  const scrollTop = ref(0);
  const scrollLeft = ref(0);
  /** The content offsets at the current zoom (see the zoom watch below). */
  let zoomAnchor = { x: 0, y: 0 };
  const width = ref(0);
  const height = ref(0);

  const z = computed(() => Math.max(0.01, deps.zoom()));
  /** The unscaled viewport. */
  const viewW = computed(() => (width.value || FALLBACK_WIDTH_PX) / z.value);
  const viewH = computed(() => (height.value || FALLBACK_HEIGHT_PX) / z.value);
  /** The unscaled px of sheet visible past the headers (frozen bands included). */
  const viewRows = computed(() => Math.max(0, viewH.value - COLUMN_HEADER_HEIGHT_PX));
  const viewCols = computed(() => Math.max(0, viewW.value - ROW_HEADER_WIDTH_PX));

  const frozenRows = computed(() => Math.min(deps.frozen().rows, extentRows.value));
  const frozenCols = computed(() => Math.min(deps.frozen().cols, extentCols.value));
  /** The frozen bands' sizes, clamped to the viewport for display. */
  const frozenH = computed(() => Math.min(deps.metrics().rowTop(frozenRows.value), viewRows.value));
  const frozenW = computed(() => Math.min(deps.metrics().colLeft(frozenCols.value), viewCols.value));

  const rowsAxis = computed(() => mapAxis(deps.metrics().totalHeight(extentRows.value), viewRows.value, scrollTop.value, z.value));
  const colsAxis = computed(() => mapAxis(deps.metrics().totalWidth(extentCols.value), viewCols.value, scrollLeft.value, z.value));
  /** The scroll offsets, unscaled content px. */
  const sy = computed(() => rowsAxis.value.offset);
  const sx = computed(() => colsAxis.value.offset);

  /** The spacer (the sheet's extent, headers excluded) and the scroll area. */
  const canvasWidth = computed(() => colsAxis.value.spacerPx);
  const canvasHeight = computed(() => rowsAxis.value.spacerPx);

  /** The rows window at `overscan` rows either side, its slots those of
   * `slotOverscan`'s (default: its own). */
  function rowWindow(overscan: number, slotOverscan = overscan): AxisWindow {
    const m = deps.metrics();
    const view = viewRows.value;
    return axisWindow({ count: frozenRows.value, frozenSize: frozenH.value, view, offset: sy.value, extent: extentRows.value, overscan, slotOverscan, ring: rowRing, defaultSize: defaultRowPx(deps.sheet()), startOf: m.rowTop, sizeOf: m.rowHeight, at: m.rowAt });
  }

  /** The columns window at `overscan` columns either side, its slots those
   * of `slotOverscan`'s (default: its own). */
  function colWindow(overscan: number, slotOverscan = overscan): AxisWindow {
    const m = deps.metrics();
    const view = viewCols.value;
    return axisWindow({ count: frozenCols.value, frozenSize: frozenW.value, view, offset: sx.value, extent: extentCols.value, overscan, slotOverscan, ring: colRing, defaultSize: defaultColumnPx(deps.sheet()), startOf: m.colLeft, sizeOf: m.colWidth, at: m.colAt });
  }

  /** SPEC-137: each axis's slot ring origin — plain state, changed only
   * with the scroll position, which recomputes every window — and how much
   * of the overscan the grid draws as it settles (sheetAxis.ts's header). */
  let rowRing = 0;
  let colRing = 0;
  const overscanDrawn = ref(1);
  const settler = new ScrollSettler((fraction) => (overscanDrawn.value = fraction));

  /** The first scrolled row and column the scroll position shows. */
  function scrolledEdge(): { row: number; col: number } {
    const m = deps.metrics();
    return { row: m.rowAt(frozenH.value + sy.value), col: m.colAt(frozenW.value + sx.value) };
  }

  /** The windows at the full overscan (the render plan may shrink it). */
  const rows = computed(() => rowWindow(GRID_OVERSCAN_ROWS));
  const cols = computed(() => colWindow(GRID_OVERSCAN_COLS));

  /** The cell at the scrolled area's top-left corner. */
  const topLeftCell = computed(
    (): CellRef => ({
      row: Math.max(frozenRows.value, deps.metrics().rowAt(frozenH.value + sy.value)),
      col: Math.max(frozenCols.value, deps.metrics().colAt(frozenW.value + sx.value)),
    }),
  );

  /** How many rows and columns PageDown / Alt+PageDown move: what the
   * scrolled area shows. */
  const page = computed(() => ({
    rows: Math.max(1, rows.value.lastVisible - rows.value.firstVisible),
    cols: Math.max(1, cols.value.lastVisible - cols.value.firstVisible),
  }));

  function measure(): void {
    const element = scroller.value;
    if (element) {
      width.value = element.clientWidth;
      height.value = element.clientHeight;
    }
  }

  /** Scrolls so the scrolled area starts at content offsets (sxNew, syNew). */
  function scrollToOffsets(sxNew: number, syNew: number): void {
    zoomAnchor = { x: Math.max(0, sxNew), y: Math.max(0, syNew) };
    const element = scroller.value;
    const top = scrollFor(rowsAxis.value, syNew, z.value);
    const left = scrollFor(colsAxis.value, sxNew, z.value);
    scrollTop.value = top;
    scrollLeft.value = left;
    if (element) {
      element.scrollTop = top;
      element.scrollLeft = left;
    }
  }

  /** Scrolls the cell (or range) into view, clear of the frozen panes. */
  function ensureVisible(target: { top: number; left: number; bottom: number; right: number }): void {
    const m = deps.metrics();
    if (target.bottom + EXTENT_ROW_MARGIN > extentRows.value) {
      extentRows.value = Math.min(MAX_ROWS, target.bottom + EXTENT_ROW_MARGIN);
    }
    if (target.right + EXTENT_COL_MARGIN > extentCols.value) {
      extentCols.value = Math.min(MAX_COLS, target.right + EXTENT_COL_MARGIN);
    }
    const mainH = Math.max(0, viewRows.value - frozenH.value);
    const mainW = Math.max(0, viewCols.value - frozenW.value);
    const nextSy =
      target.top >= frozenRows.value ? revealOffset(m.rowTop(target.top), m.rowTop(target.bottom + 1), frozenH.value, mainH, sy.value) : sy.value;
    const nextSx =
      target.left >= frozenCols.value ? revealOffset(m.colLeft(target.left), m.colLeft(target.right + 1), frozenW.value, mainW, sx.value) : sx.value;
    if (nextSy !== sy.value || nextSx !== sx.value) {
      scrollToOffsets(nextSx, nextSy);
    }
  }

  /** Scrolls so `cell` is the scrolled area's top-left. The extent first
   * grows to hold a whole viewport past it — a stored top-left (SPEC-138's
   * reopen) is otherwise clamped short of it, at the extent's end — and the
   * scroll then waits for the spacer to render at that size. */
  function scrollToTopLeft(cell: CellRef): void {
    const m = deps.metrics();
    const nextSy = cell.row >= frozenRows.value ? m.rowTop(cell.row) - m.rowTop(frozenRows.value) : 0;
    const nextSx = cell.col >= frozenCols.value ? m.colLeft(cell.col) - m.colLeft(frozenCols.value) : 0;
    const needRows = Math.min(MAX_ROWS, m.rowAt(m.rowTop(cell.row) + viewRows.value) + EXTENT_ROW_MARGIN);
    const needCols = Math.min(MAX_COLS, m.colAt(m.colLeft(cell.col) + viewCols.value) + EXTENT_COL_MARGIN);
    if (needRows <= extentRows.value && needCols <= extentCols.value) {
      scrollToOffsets(nextSx, nextSy);
      return;
    }
    extentRows.value = Math.max(extentRows.value, needRows);
    extentCols.value = Math.max(extentCols.value, needCols);
    void nextTick(() => scrollToOffsets(nextSx, nextSy));
  }

  /** Reads the scroll position, growing the extent when the viewport comes
   * within one screen of its end. */
  function onScroll(): void {
    const element = scroller.value;
    if (!element) {
      return;
    }
    // Far: past the overscan's reach on either axis (sheetAxis.ts's header).
    const far =
      Math.abs(element.scrollTop - scrollTop.value) > GRID_OVERSCAN_ROWS * defaultRowPx(deps.sheet()) * z.value ||
      Math.abs(element.scrollLeft - scrollLeft.value) > GRID_OVERSCAN_COLS * defaultColumnPx(deps.sheet()) * z.value;
    settler.step(far);
    const edge = far ? scrolledEdge() : undefined;
    scrollTop.value = element.scrollTop;
    scrollLeft.value = element.scrollLeft;
    if (edge) {
      const next = scrolledEdge();
      rowRing += next.row - edge.row;
      colRing += next.col - edge.col;
    }
    zoomAnchor = { x: sx.value, y: sy.value };
    const rowsView = viewRows.value * z.value;
    const colsView = viewCols.value * z.value;
    if (scrollTop.value + 2 * rowsView >= canvasHeight.value && extentRows.value < MAX_ROWS) {
      extentRows.value = Math.min(MAX_ROWS, extentRows.value + EXTENT_ROW_MARGIN);
    }
    if (scrollLeft.value + 2 * colsView >= canvasWidth.value && extentCols.value < MAX_COLS) {
      extentCols.value = Math.min(MAX_COLS, extentCols.value + EXTENT_COL_MARGIN);
    }
  }

  // Report a settled top-left cell (view state stores it as `topLeftCell`).
  watch(topLeftCell, (cell, previous) => {
    if (!previous || cell.row !== previous.row || cell.col !== previous.col) {
      deps.onTopLeftChange?.(cell);
    }
  });

  // A zoom change keeps the same content at the scrolled area's top-left:
  // the offsets are recorded wherever they change at a stable zoom (a scroll
  // event, a programmatic scroll) and re-applied under the new scale.
  watch(z, () => {
    const keep = { ...zoomAnchor };
    void nextTick(() => scrollToOffsets(keep.x, keep.y));
  });

  let observer: ResizeObserver | null = null;
  onMounted(() => {
    const initial = deps.initialTopLeft();
    const scrolled = initial !== undefined && (initial.row > 0 || initial.col > 0);
    const observed = typeof ResizeObserver === "function" && scroller.value !== null;
    // SPEC-137: with an observer, the first size comes from its first
    // callback, which the browser delivers straight after the first layout;
    // reading it here would force that layout — the whole view's, ~30 ms of
    // a large workbook's open — inside the task that mounted it (the first
    // render takes the fallback size either way). Scrolling to a stored
    // top-left lays the page out regardless, so that path measures first.
    if (!observed || scrolled) {
      measure();
    }
    if (scrolled) {
      scrollToTopLeft(initial);
    }
    if (observed) {
      observer = new ResizeObserver(measure);
      observer.observe(scroller.value!);
    }
  });
  onBeforeUnmount(() => {
    observer?.disconnect();
    settler.dispose();
  });

  return {
    extentRows,
    extentCols,
    canvasWidth,
    canvasHeight,
    scrollTop,
    scrollLeft,
    z,
    viewW,
    viewH,
    viewRows,
    viewCols,
    frozenRows,
    frozenCols,
    frozenH,
    frozenW,
    sx,
    sy,
    rows,
    cols,
    rowWindow,
    colWindow,
    overscanDrawn,
    page,
    topLeftCell,
    onScroll,
    ensureVisible,
    scrollToTopLeft,
    scrollToOffsets,
    measure,
  };
}

export type SheetViewport = ReturnType<typeof useSheetViewport>;
