// The read-only grid's viewport (SPEC-116 "The read-only grid"): the scroll
// position → the window of rows and columns SheetGrid renders (the viewport
// plus GRID_OVERSCAN_ROWS/COLS), the scroll extent its spacer is sized to,
// and where each row and column sits. The grid is DOM, not canvas — a
// visible cell needs real text layout, an accessibility node and a selector
// a test can find — and this composable is what bounds how many exist.
//
// - Geometry is uniform: DEFAULT_COL_WIDTH_PX × DEFAULT_ROW_HEIGHT_PX
//   (SPEC-117 replaces it with GridMetrics).
// - The extent is max(usedRange.bottom + 100, 1000) rows by
//   max(usedRange.right + 10, 26) columns; scrolling within one screen of the
//   end grows it by the same margins, up to MAX_ROWS/MAX_COLS.
// - Above MAX_SCROLL_EXTENT_PX (under every target browser's maximum element
//   size) the spacer stops growing and scroll offsets map onto rows
//   proportionally, so a used range at row 1,000,000 is reachable
//   everywhere.
import { computed, onBeforeUnmount, onMounted, ref, type Ref } from "vue";
import { colToLetters, findRow, formatA1, usedRange } from "../../lib/spreadsheet/address";
import { generalText } from "../../lib/spreadsheet/general";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type Sheet } from "../../lib/spreadsheet/model";

export const DEFAULT_COL_WIDTH_PX = 64;
export const DEFAULT_ROW_HEIGHT_PX = 20;
export const GRID_OVERSCAN_ROWS = 8;
export const GRID_OVERSCAN_COLS = 4;
export const MAX_SCROLL_EXTENT_PX = 15_000_000;
/** The row-number column (wide enough for "1048576") and the letter row. */
export const ROW_HEADER_WIDTH_PX = 56;
export const COLUMN_HEADER_HEIGHT_PX = DEFAULT_ROW_HEIGHT_PX;

const EXTENT_ROW_MARGIN = 100;
const MIN_EXTENT_ROWS = 1000;
const EXTENT_COL_MARGIN = 10;
const MIN_EXTENT_COLS = 26;

/** What an unmeasured grid (jsdom, a hidden tab) assumes it shows — a
 * desktop window, so the rendered window is never empty. */
const FALLBACK_WIDTH_PX = 1280;
const FALLBACK_HEIGHT_PX = 720;

/** One axis of the viewport: `count` items of `size` px in a spacer capped
 * at MAX_SCROLL_EXTENT_PX, scrolled to `scroll` with `view` px visible. */
interface Axis {
  /** The spacer's length in px. */
  extentPx: number;
  /** The unscaled offset of the viewport's leading edge into the items. */
  virtual: number;
  first: number;
  last: number;
}

function axis(count: number, size: number, scroll: number, view: number, overscan: number): Axis {
  const totalPx = count * size;
  const extentPx = Math.min(totalPx, MAX_SCROLL_EXTENT_PX);
  let virtual = scroll;
  if (totalPx > extentPx) {
    const range = extentPx - view;
    virtual = range > 0 ? (Math.min(scroll, range) / range) * (totalPx - view) : 0;
  }
  return {
    extentPx,
    virtual,
    first: Math.max(0, Math.floor(virtual / size) - overscan),
    last: Math.max(0, Math.min(count - 1, Math.floor((virtual + view) / size) + overscan)),
  };
}

/** How a value sits in its cell: numbers right, text left, booleans and
 * errors centred (Excel's General alignment). */
export function cellAlignment(value: CellValue): "number" | "text" | "centred" {
  if (typeof value === "number") {
    return "number";
  }
  return typeof value === "string" ? "text" : "centred";
}

/** Narrows the runtime-typed `sheet` prop (the .vue parser bars casts). */
export function asSheet(value: unknown): Sheet {
  return value as Sheet;
}

/** The index of the first cell at or right of `col` (cells ascend). */
function firstCellFrom(cells: readonly Cell[], col: number): number {
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (cells[mid]!.col < col) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

export function useSheetViewport(scroller: Readonly<Ref<HTMLElement | null>>, sheet: () => Sheet) {
  const used = usedRange(sheet());
  const extentRows = ref(Math.min(MAX_ROWS, Math.max((used?.bottom ?? 0) + EXTENT_ROW_MARGIN, MIN_EXTENT_ROWS)));
  const extentCols = ref(Math.min(MAX_COLS, Math.max((used?.right ?? 0) + EXTENT_COL_MARGIN, MIN_EXTENT_COLS)));
  const scrollTop = ref(0);
  const scrollLeft = ref(0);
  const width = ref(0);
  const height = ref(0);

  const viewRowsPx = computed(() => Math.max(0, (height.value || FALLBACK_HEIGHT_PX) - COLUMN_HEADER_HEIGHT_PX));
  const viewColsPx = computed(() => Math.max(0, (width.value || FALLBACK_WIDTH_PX) - ROW_HEADER_WIDTH_PX));
  const rows = computed(() => axis(extentRows.value, DEFAULT_ROW_HEIGHT_PX, scrollTop.value, viewRowsPx.value, GRID_OVERSCAN_ROWS));
  const cols = computed(() => axis(extentCols.value, DEFAULT_COL_WIDTH_PX, scrollLeft.value, viewColsPx.value, GRID_OVERSCAN_COLS));

  /** The spacer: the scroll extent, headers included. */
  const canvasWidth = computed(() => ROW_HEADER_WIDTH_PX + cols.value.extentPx);
  const canvasHeight = computed(() => rows.value.extentPx);

  /** A row's top inside the row canvas — `row × height` until the extent is
   * capped, then placed relative to the scroll position by the proportional
   * mapping. */
  function rowTop(row: number): number {
    return scrollTop.value + row * DEFAULT_ROW_HEIGHT_PX - rows.value.virtual;
  }

  /** A column's left edge inside a row (past the row header). */
  function colLeft(col: number): number {
    return ROW_HEADER_WIDTH_PX + scrollLeft.value + col * DEFAULT_COL_WIDTH_PX - cols.value.virtual;
  }

  function measure(): void {
    const element = scroller.value;
    if (element) {
      width.value = element.clientWidth;
      height.value = element.clientHeight;
    }
  }

  /** Reads the scroll position, growing the extent when the viewport comes
   * within one screen of its end. */
  function onScroll(): void {
    const element = scroller.value;
    if (!element) {
      return;
    }
    scrollTop.value = element.scrollTop;
    scrollLeft.value = element.scrollLeft;
    if (scrollTop.value + 2 * viewRowsPx.value >= canvasHeight.value && extentRows.value < MAX_ROWS) {
      extentRows.value = Math.min(MAX_ROWS, extentRows.value + EXTENT_ROW_MARGIN);
    }
    if (scrollLeft.value + 2 * viewColsPx.value >= cols.value.extentPx && extentCols.value < MAX_COLS) {
      extentCols.value = Math.min(MAX_COLS, extentCols.value + EXTENT_COL_MARGIN);
    }
  }

  /** The column headers in the window. */
  const visibleCols = computed(() => {
    const out: { index: number; letters: string; left: number }[] = [];
    for (let col = cols.value.first; col <= cols.value.last; col += 1) {
      out.push({ index: col, letters: colToLetters(col), left: colLeft(col) });
    }
    return out;
  });

  /** Every row in the window, with its stored cells in the window shown
   * through the General display. Empty cells render nothing — the row draws
   * the gridlines — so a cell element exists only where a value does. */
  const visibleRows = computed(() => {
    const current = sheet();
    const firstCol = cols.value.first;
    const lastCol = cols.value.last;
    const out: { index: number; top: number; cells: { col: number; ref: string; left: number; text: string; align: string }[] }[] = [];
    for (let row = rows.value.first; row <= rows.value.last; row += 1) {
      const cells = findRow(current, row)?.cells ?? [];
      const shown = [];
      for (let i = firstCellFrom(cells, firstCol); i < cells.length && cells[i]!.col <= lastCol; i += 1) {
        const cell = cells[i]!;
        shown.push({
          col: cell.col,
          ref: formatA1({ row, col: cell.col }),
          left: colLeft(cell.col),
          text: generalText(cell.value),
          align: cellAlignment(cell.value),
        });
      }
      out.push({ index: row, top: rowTop(row), cells: shown });
    }
    return out;
  });

  let observer: ResizeObserver | null = null;
  onMounted(() => {
    measure();
    if (typeof ResizeObserver === "function" && scroller.value) {
      observer = new ResizeObserver(measure);
      observer.observe(scroller.value);
    }
  });
  onBeforeUnmount(() => observer?.disconnect());

  return {
    extentRows,
    extentCols,
    canvasWidth,
    canvasHeight,
    firstRow: computed(() => rows.value.first),
    lastRow: computed(() => rows.value.last),
    firstCol: computed(() => cols.value.first),
    lastCol: computed(() => cols.value.last),
    rowTop,
    colLeft,
    visibleCols,
    visibleRows,
    onScroll,
  };
}
