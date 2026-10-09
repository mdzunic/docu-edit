// Grid geometry (SPEC-117 "Geometry"): column widths in px at 100% zoom,
// row heights in pt, the xlsx character-width ↔ px conversion and the
// GridMetrics every renderer, hit test and scroll computation reads.
//
// - Widths come from the workbook default font's maximum digit width (MDW),
//   looked up in a pinned table — never measured: a column is as wide on a
//   machine without Calibri as on one with it.
// - Queries answer integer px: colWidth = round(widthPx), rowHeight =
//   round(heightPt × 4/3) (15 pt → 20 px); a hidden row or column is 0.
// - Custom sizes are sorted index runs with prefix sums of
//   (size − default) × length, so every query is a binary search — O(log n)
//   in the number of custom runs — and nothing is ever allocated per row of
//   the 1,048,576-row grid. Zoom is a view transform applied outside: these
//   numbers never see it.
import { findRow } from "./address";
import { MAX_COLS, MAX_ROWS, type Sheet } from "./model";

/** Padding on each side of a cell's text; + the 1 px gridline = Excel's 5 px. */
export const CELL_PADDING_PX = 2;
export const INDENT_PX_PER_LEVEL = 9;
/** The default column for Calibri 11 (and Aptos Narrow 11): 8.43 characters. */
export const DEFAULT_COL_WIDTH_PX = 64;
/** `sheetFormatPr@defaultRowHeight` when the element is absent. */
export const DEFAULT_ROW_HEIGHT_PT = 15;
/** LibreOffice's defaults when a table declares none: 0.8889 in and 12.8 pt. */
export const ODS_DEFAULT_COL_WIDTH_PX = 85.33;
export const ODS_DEFAULT_ROW_HEIGHT_PT = 12.8;

/** MDW ratios by family (case-insensitive); a metric twin shares its
 * original's row. Aptos Narrow shares Calibri's because Excel's switch to it
 * kept the 8.43-character, 64 px default column. */
const MDW_RATIOS: ReadonlyArray<readonly [readonly string[], number]> = [
  [["calibri", "carlito", "aptos narrow"], 0.5068],
  [["arial", "helvetica", "liberation sans", "arimo"], 0.5562],
  [["times new roman", "liberation serif", "tinos"], 0.5],
  [["cambria", "caladea"], 0.5557],
  [["courier new", "liberation mono", "cousine"], 0.6001],
];
const OTHER_FAMILY_RATIO = 0.5068;

/** The maximum digit width in px of `family` at `sizePt`, from the pinned
 * table: round(ratio × sizePt × 96 / 72), at least 1. */
export function maxDigitWidthPx(family: string, sizePt: number): number {
  const key = family.trim().toLowerCase();
  const ratio = MDW_RATIOS.find(([names]) => names.includes(key))?.[1] ?? OTHER_FAMILY_RATIO;
  return Math.max(1, Math.round((ratio * sizePt * 96) / 72));
}

/** An xlsx `<col width>` (characters) → px: trunc(((256·w + trunc(128/mdw)) / 256) · mdw). */
export function xlsxWidthToPx(width: number, mdw: number): number {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** px → an xlsx `<col width>`, Excel's own inverse:
 * trunc((trunc((px − 5)/mdw · 100 + 0.5)/100 · mdw + 5)/mdw · 256)/256. */
export function pxToXlsxWidth(px: number, mdw: number): number {
  const characters = Math.trunc(((px - 5) / mdw) * 100 + 0.5) / 100;
  return Math.trunc(((characters * mdw + 5) / mdw) * 256) / 256;
}

/** `sheetFormatPr@baseColWidth` (characters, default 8) → the default column
 * in px: ceil8(trunc(((256·b + trunc(128/mdw)) / 256) · mdw) + 5). */
export function baseColWidthToPx(baseColWidth: number, mdw: number): number {
  const px = Math.trunc(((256 * baseColWidth + Math.trunc(128 / mdw)) / 256) * mdw) + 5;
  return Math.ceil(px / 8) * 8;
}

/** A height in pt as the grid's integer px. */
export function ptToPx(pt: number): number {
  return Math.round((pt * 4) / 3);
}

export interface GridMetrics {
  colLeft(c: number): number;
  colWidth(c: number): number;
  colAt(x: number): number;
  rowTop(r: number): number;
  rowHeight(r: number): number;
  rowAt(y: number): number;
  /** The width of the first `cols` columns (= colLeft(cols)). */
  totalWidth(cols: number): number;
  /** The height of the first `rows` rows (= rowTop(rows)). */
  totalHeight(rows: number): number;
}

/** One axis: sorted, non-overlapping runs of items whose size differs from
 * the default, with prefix sums of their deltas. */
interface Axis {
  size: number; // the default item size, ≥ 1
  count: number; // items on the axis
  runs: number;
  starts: Int32Array;
  ends: Int32Array; // inclusive
  sizes: Int32Array;
  /** prefix[k] = Σ (sizes[j] − size) × length(j) over runs j < k. */
  prefix: Float64Array;
  /** offsets[k] = the leading edge of run k's first item. */
  offsets: Float64Array;
}

interface Run {
  start: number;
  end: number;
  size: number;
}

/** Appends an item run, merging it into the previous one when adjacent and
 * the same size; skips default-sized items. */
function pushRun(runs: Run[], start: number, end: number, size: number, defaultSize: number): void {
  if (size === defaultSize || end < start) {
    return;
  }
  const last = runs.at(-1);
  if (last && last.end + 1 === start && last.size === size) {
    last.end = end;
    return;
  }
  runs.push({ start, end, size });
}

function buildAxis(runs: readonly Run[], defaultSize: number, count: number): Axis {
  const n = runs.length;
  const axis: Axis = {
    size: defaultSize,
    count,
    runs: n,
    starts: new Int32Array(n),
    ends: new Int32Array(n),
    sizes: new Int32Array(n),
    prefix: new Float64Array(n + 1),
    offsets: new Float64Array(n),
  };
  for (let k = 0; k < n; k += 1) {
    const run = runs[k]!;
    axis.starts[k] = run.start;
    axis.ends[k] = run.end;
    axis.sizes[k] = run.size;
    axis.offsets[k] = run.start * defaultSize + axis.prefix[k]!;
    axis.prefix[k + 1] = axis.prefix[k]! + (run.size - defaultSize) * (run.end - run.start + 1);
  }
  return axis;
}

/** The last run starting at or before item `i`, or -1. */
function runAtOrBefore(axis: Axis, i: number): number {
  let lo = 0;
  let hi = axis.runs - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (axis.starts[mid]! <= i) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}

function offsetOf(axis: Axis, i: number): number {
  const k = runAtOrBefore(axis, i);
  if (k < 0) {
    return i * axis.size;
  }
  if (i > axis.ends[k]!) {
    return i * axis.size + axis.prefix[k + 1]!;
  }
  return i * axis.size + axis.prefix[k]! + (i - axis.starts[k]!) * (axis.sizes[k]! - axis.size);
}

function sizeOf(axis: Axis, i: number): number {
  const k = runAtOrBefore(axis, i);
  return k >= 0 && i <= axis.ends[k]! ? axis.sizes[k]! : axis.size;
}

/** The item whose extent holds `pos` — never a zero-sized one — clamped to
 * the axis. */
function indexAt(axis: Axis, pos: number): number {
  const at = Math.max(0, pos);
  let lo = 0;
  let hi = axis.runs - 1;
  let k = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (axis.offsets[mid]! <= at) {
      k = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  let index: number;
  if (k < 0) {
    index = Math.floor(at / axis.size);
  } else {
    const runSize = axis.sizes[k]!;
    const runEnd = axis.offsets[k]! + runSize * (axis.ends[k]! - axis.starts[k]! + 1);
    index =
      at < runEnd ? axis.starts[k]! + Math.floor((at - axis.offsets[k]!) / runSize) : axis.ends[k]! + 1 + Math.floor((at - runEnd) / axis.size);
  }
  return Math.min(index, axis.count - 1);
}

/** The column runs of a sheet: its spans' widths (hidden → 0). */
function columnRuns(sheet: Sheet, defaultSize: number): Run[] {
  const runs: Run[] = [];
  for (const span of sheet.columns ?? []) {
    const size = span.hidden ? 0 : Math.round(span.widthPx ?? sheet.defaultColWidthPx ?? DEFAULT_COL_WIDTH_PX);
    pushRun(runs, Math.max(0, span.min), Math.min(MAX_COLS - 1, span.max), size, defaultSize);
  }
  return runs;
}

/** The row runs of a sheet: one pass over the stored rows, skipping every
 * row with no size property. */
function rowRuns(sheet: Sheet, defaultSize: number): Run[] {
  const runs: Run[] = [];
  for (const row of sheet.rows) {
    if (row.hidden) {
      pushRun(runs, row.index, row.index, 0, defaultSize);
    } else if (row.heightPt !== undefined) {
      pushRun(runs, row.index, row.index, ptToPx(row.heightPt), defaultSize);
    }
  }
  return runs;
}

/** The default column width of a sheet in integer px (≥ 1). */
export function defaultColumnPx(sheet: Pick<Sheet, "defaultColWidthPx">): number {
  return Math.max(1, Math.round(sheet.defaultColWidthPx ?? DEFAULT_COL_WIDTH_PX));
}

/** The default row height of a sheet in integer px (≥ 1). */
export function defaultRowPx(sheet: Pick<Sheet, "defaultRowHeightPt">): number {
  return Math.max(1, ptToPx(sheet.defaultRowHeightPt ?? DEFAULT_ROW_HEIGHT_PT));
}

/** The sheet's GridMetrics. Built in one pass over its column spans and
 * stored rows; rebuild it when the sheet's `columns` or `rows` change. */
export function buildGridMetrics(sheet: Sheet): GridMetrics {
  const colDefault = defaultColumnPx(sheet);
  const rowDefault = defaultRowPx(sheet);
  const cols = buildAxis(columnRuns(sheet, colDefault), colDefault, MAX_COLS);
  const rows = buildAxis(rowRuns(sheet, rowDefault), rowDefault, MAX_ROWS);
  return {
    colLeft: (c) => offsetOf(cols, c),
    colWidth: (c) => sizeOf(cols, c),
    colAt: (x) => indexAt(cols, x),
    rowTop: (r) => offsetOf(rows, r),
    rowHeight: (r) => sizeOf(rows, r),
    rowAt: (y) => indexAt(rows, y),
    totalWidth: (count) => offsetOf(cols, count),
    totalHeight: (count) => offsetOf(rows, count),
  };
}

/** A run of hidden rows or columns, inclusive. */
export interface HiddenRun {
  start: number;
  end: number;
}

/** The hidden column runs touching [first − 1, last + 1] — where a header
 * boundary marker belongs ("Hidden columns C–D"). */
export function hiddenColumnRuns(sheet: Sheet, first: number, last: number): HiddenRun[] {
  const out: HiddenRun[] = [];
  for (const span of sheet.columns ?? []) {
    if (!span.hidden || span.max < first - 1 || span.min > last + 1) {
      continue;
    }
    const previous = out.at(-1);
    if (previous && previous.end + 1 === span.min) {
      previous.end = span.max;
    } else {
      out.push({ start: span.min, end: span.max });
    }
  }
  return out;
}

/** The hidden row runs touching [first − 1, last + 1]: a binary search to
 * the window, then a walk over the stored rows inside it. */
export function hiddenRowRuns(sheet: Sheet, first: number, last: number): HiddenRun[] {
  const rows = sheet.rows;
  let lo = 0;
  let hi = rows.length;
  const from = Math.max(0, first - 1);
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (rows[mid]!.index < from) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  // A run that started above the window but reaches into it.
  let begin = lo;
  if (lo < rows.length && rows[lo]!.hidden) {
    while (begin > 0 && rows[begin - 1]!.hidden && rows[begin - 1]!.index + 1 === rows[begin]!.index) {
      begin -= 1;
    }
  }
  const out: HiddenRun[] = [];
  for (let i = begin; i < rows.length && rows[i]!.index <= last + 1; i += 1) {
    const row = rows[i]!;
    if (!row.hidden) {
      continue;
    }
    const previous = out.at(-1);
    if (previous && previous.end + 1 === row.index) {
      previous.end = row.index;
    } else {
      out.push({ start: row.index, end: row.index });
    }
  }
  // Extend a run that continues below the window.
  const tail = out.at(-1);
  if (tail) {
    let next = findRow(sheet, tail.end + 1);
    while (next?.hidden) {
      tail.end = next.index;
      next = findRow(sheet, tail.end + 1);
    }
  }
  return out;
}
