// The range-class engine behind SPEC-123's formatting commands ("Range
// classes"): rewrites the style ids a range shows — position by position for
// an ordinary range, through `Row.styleId` for whole rows, through
// ColumnSpans for whole columns, and through every span, styled row and
// stored cell for the whole sheet — without ever materialising a position
// whose winner (cell → row → column → 0, styles.ts's effectiveStyleId) does
// not change. formatting.ts drives it with a StylePatch, Clear's removal and
// the format painter's ids; THIS FILE owns the model invariants every one of
// those commands keeps:
//
// - Copy-on-write (model.ts invariant 1): applying a range returns a NEW
//   Sheet with a new `rows` array (and a new `columns` array when a span
//   changed), new Row objects only for rows it changed and new Cell objects
//   only for cells it changed; every untouched row and cell keeps its
//   identity. A range that changes nothing returns the sheet itself. The
//   caller works on its own Sheet values and writes the model once, at the
//   end — so a refused command never touches it.
// - Sparse (invariant 2): an empty position gets a style-only cell only when
//   its new style differs from what it resolves to without a cell (its row's
//   style, else its column's, else 0); a cell left holding nothing is
//   removed, a row left holding nothing is removed, a span is created only
//   where a whole-column rewrite changes the default of a column no span
//   covers, and a span left holding nothing is dropped.
// - A cell the engine writes carries its styleId explicitly whenever it is
//   non-zero, and `styleId: 0` when it is 0 but its row/column fallback is
//   not, so it means the same thing under any reading of an absent id; a
//   cell whose look does not change keeps its object, and an absent id it
//   already inherits stays absent.
// - Materialisation is counted as it happens — new Row objects and
//   style-only cells alike — and passing MAX_FORMAT_MATERIALISED throws
//   TooManyPositions before the cells are built, so a command that would
//   explode (a million-row selection under a styled column) is refused.
//
// Range classes (SPEC-123's table):
// - Ordinary: every position of the range; stored cells are rewritten, and
//   an empty position whose rewritten fallback differs from its fallback
//   gets a style-only cell (≤ area).
// - Whole rows (`left = 0`, `right = MAX_COLS − 1`): a row with a style gets
//   its style rewritten; a row without one gets `styleId` = the rewrite of
//   0, and — because the row now wins over the columns — a cell for every
//   empty position of a styled span whose rewritten style differs from the
//   row's. When the rewrite of 0 is 0 itself the row stays unstyled and only
//   the styled spans' positions that change get cells (the sparser
//   equivalent). Every stored cell is rewritten.
// - Whole columns (`top = 0`, `bottom = MAX_ROWS − 1`): spans are split at
//   the range edges and rewritten, gaps the rewrite of 0 changes become new
//   spans; every stored cell is rewritten; a styled row still wins over the
//   column, so each (styled row × target column) whose rewritten row style
//   differs gets a cell.
// - Whole sheet (both): every span (gaps filled when the rewrite of 0
//   changes), every row with a style and every stored cell are rewritten;
//   nothing is materialised, since no position's winner changes.
// A `storedOnly` rewrite (the border presets' neighbour clearing) touches
// only what exists — stored cells, styled rows, existing spans — and never
// creates a cell, row or span.
import { MAX_COLS, MAX_ROWS, type Cell, type ColumnSpan, type Row, type Sheet } from "./model";
import type { GridRange } from "./address";
import { isEmptyRow } from "./cells";
import { columnSpanAt } from "./styles";

/** The most cells and rows one formatting command may materialise. */
export const MAX_FORMAT_MATERIALISED = 1_000_000;

/** Thrown when a command would materialise more than MAX_FORMAT_MATERIALISED. */
export class TooManyPositions extends Error {
  constructor() {
    super("Too many positions to format one by one");
    this.name = "TooManyPositions";
  }
}

/** The running count of what one command materialises. */
export class MaterialiseBudget {
  private created = 0;

  /** Reserves `n` more cells or rows; throws TooManyPositions past the cap. */
  take(n: number): void {
    if (this.created + n > MAX_FORMAT_MATERIALISED) {
      throw new TooManyPositions();
    }
    this.created += n;
  }

  get used(): number {
    return this.created;
  }
}

/** How one command rewrites the style ids of a range. */
export interface RestyleOp {
  /** The style id a cell, row or span whose current style is `base` gets —
   * `base` itself for no change. */
  next(base: number): number;
  /** A stored cell's own extra change (a font patch reaching rich-text runs);
   * returns the cell itself when nothing changes. */
  cell?(cell: Cell): Cell;
  /** Touch only what exists (see the module header). */
  storedOnly?: boolean;
}

export type RangeClass = "ordinary" | "rows" | "columns" | "sheet";

/** Which of SPEC-123's range classes a range is. */
export function rangeClass(range: GridRange): RangeClass {
  const rows = range.left === 0 && range.right === MAX_COLS - 1;
  const columns = range.top === 0 && range.bottom === MAX_ROWS - 1;
  return rows && columns ? "sheet" : rows ? "rows" : columns ? "columns" : "ordinary";
}

// --- lookups -----------------------------------------------------------------

/** The first stored row at or after `index` (an index into `rows`). */
export function firstRowAtOrAfter(rows: readonly Row[], index: number): number {
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (rows[mid]!.index < index) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** The first cell at or after `col` (an index into `cells`). */
export function firstCellAtOrAfter(cells: readonly Cell[], col: number): number {
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

/** A run of columns sharing one column default. */
export interface ColumnRun {
  min: number;
  max: number;
  /** The span's styleId (undefined: no span, or a span without a style). */
  styleId: number | undefined;
}

/** The runs of [left, right] by column default — spans and the gaps between
 * them — in column order; O(spans), never per column. */
export function columnRuns(columns: readonly ColumnSpan[], left: number, right: number): ColumnRun[] {
  const runs: ColumnRun[] = [];
  let col = left;
  for (const span of columns) {
    if (span.max < col) {
      continue;
    }
    if (span.min > right) {
      break;
    }
    if (span.min > col) {
      runs.push({ min: col, max: span.min - 1, styleId: undefined });
    }
    const max = Math.min(span.max, right);
    runs.push({ min: Math.max(span.min, col), max, styleId: span.styleId });
    col = max + 1;
    if (col > right) {
      break;
    }
  }
  if (col <= right) {
    runs.push({ min: col, max: right, styleId: undefined });
  }
  return runs;
}

/** Whether a cell holds nothing but its style: a null value and no other
 * field (a formula, runs, preserved content all count as content). */
export function holdsOnlyStyle(cell: Cell): boolean {
  if (cell.value !== null) {
    return false;
  }
  const record = cell as unknown as Record<string, unknown>;
  return Object.keys(record).every((key) => key === "col" || key === "value" || key === "styleId" || record[key] === undefined);
}

/** Whether a span holds anything but its range: a width, `hidden`, a style
 * or an attribute the writers re-emit. */
function spanHoldsSomething(span: ColumnSpan): boolean {
  const record = span as unknown as Record<string, unknown>;
  return Object.keys(record).some((key) => key !== "min" && key !== "max" && record[key] !== undefined);
}

// --- one row -------------------------------------------------------------------

/** A run of empty positions that each get a style-only cell. */
interface Fill {
  min: number;
  max: number;
  styleId: number;
}

/** What one row's positions in [left, right] become. */
interface RowRewrite {
  left: number;
  right: number;
  /** A stored cell's new style id. */
  cellWant(cell: Cell): number;
  /** What a position without a cell shows after the command. */
  fallbackAfter(col: number): number;
  /** Runs (ascending, inside [left, right]) whose empty positions get cells. */
  fills: readonly Fill[];
  cellHook?: (cell: Cell) => Cell;
  /** Set the row's own styleId (undefined removes it). */
  rowStyle?: { value: number | undefined };
}

/** A stored cell with its new style id (see the module header): the cell
 * itself when nothing changes, undefined when it is left holding nothing. */
export function writeCell(cell: Cell, want: number, fallbackAfter: number, hook?: (cell: Cell) => Cell): Cell | undefined {
  let next = hook ? hook(cell) : cell;
  const own = next.styleId;
  const target = (own === undefined && want === fallbackAfter) || (want === 0 && fallbackAfter === 0) ? undefined : want;
  if (target !== own) {
    next = next === cell ? { ...cell } : next;
    if (target === undefined) {
      delete next.styleId;
    } else {
      next.styleId = target;
    }
  }
  if (holdsOnlyStyle(next) && (next.styleId === undefined || next.styleId === fallbackAfter)) {
    return undefined;
  }
  return next;
}

/** One row rewritten: the same object when nothing changes, undefined when
 * it ends up storing nothing (or never existed and still holds nothing). A
 * new Row and every style-only cell are taken from `budget` first. */
function rewriteRow(row: Row | undefined, index: number, rewrite: RowRewrite, budget: MaterialiseBudget): Row | undefined {
  const cells = row?.cells ?? [];
  const start = firstCellAtOrAfter(cells, rewrite.left);
  let end = start;
  while (end < cells.length && cells[end]!.col <= rewrite.right) {
    end += 1;
  }
  // The fills' empty positions, counted before anything is built.
  let fillCount = 0;
  if (rewrite.fills.length > 0) {
    let at = start;
    for (const fill of rewrite.fills) {
      while (at < end && cells[at]!.col < fill.min) {
        at += 1;
      }
      let stored = 0;
      let probe = at;
      while (probe < end && cells[probe]!.col <= fill.max) {
        stored += 1;
        probe += 1;
      }
      fillCount += fill.max - fill.min + 1 - stored;
    }
  }
  const styleChanges = rewrite.rowStyle !== undefined && rewrite.rowStyle.value !== row?.styleId;
  const creates = row === undefined && (fillCount > 0 || (styleChanges && rewrite.rowStyle?.value !== undefined));
  if (fillCount > 0 || creates) {
    budget.take(fillCount + (creates ? 1 : 0));
  }
  let changed = styleChanges || fillCount > 0;
  const middle: Cell[] = [];
  let fillIndex = 0;
  let col = rewrite.left;
  const emitFills = (upTo: number): void => {
    // Style-only cells for the fill positions before column `upTo`.
    while (fillIndex < rewrite.fills.length) {
      const fill = rewrite.fills[fillIndex]!;
      const from = Math.max(fill.min, col);
      const to = Math.min(fill.max, upTo - 1);
      for (let c = from; c <= to; c += 1) {
        middle.push({ col: c, value: null, styleId: fill.styleId });
      }
      if (fill.max >= upTo) {
        return;
      }
      fillIndex += 1;
    }
  };
  for (let i = start; i < end; i += 1) {
    const cell = cells[i]!;
    emitFills(cell.col);
    col = cell.col + 1;
    const next = writeCell(cell, rewrite.cellWant(cell), rewrite.fallbackAfter(cell.col), rewrite.cellHook);
    if (next !== cell) {
      changed = true;
    }
    if (next) {
      middle.push(next);
    }
  }
  emitFills(rewrite.right + 1);
  if (!changed) {
    return row;
  }
  const nextCells = start === 0 && end === cells.length ? middle : [...cells.slice(0, start), ...middle, ...cells.slice(end)];
  const next: Row = row ? { ...row, cells: nextCells } : { index, cells: nextCells };
  if (rewrite.rowStyle !== undefined) {
    if (rewrite.rowStyle.value === undefined) {
      delete next.styleId;
    } else {
      next.styleId = rewrite.rowStyle.value;
    }
  }
  return isEmptyRow(next) ? undefined : next;
}

// --- spans -----------------------------------------------------------------------

/** The spans with [left, right] rewritten (see "Whole columns"): the same
 * array when nothing changes. */
function rewriteSpans(columns: ColumnSpan[], left: number, right: number, op: RestyleOp): ColumnSpan[] {
  const out: ColumnSpan[] = [];
  let changed = false;
  let col = left;
  const gap = (min: number, max: number): void => {
    if (op.storedOnly || min > max) {
      return;
    }
    const want = op.next(0);
    if (want !== 0) {
      out.push({ min, max, styleId: want });
      changed = true;
    }
  };
  for (const span of columns) {
    if (span.max < left || span.min > right) {
      if (span.min > right && col <= right) {
        gap(col, right);
        col = right + 1;
      }
      out.push(span);
      continue;
    }
    const base = span.styleId ?? 0;
    const want = op.next(base);
    if (span.min > col) {
      gap(col, span.min - 1);
    }
    col = Math.min(span.max, right) + 1;
    if (want === base) {
      out.push(span);
      continue;
    }
    changed = true;
    if (span.min < left) {
      out.push({ ...span, max: left - 1 });
    }
    const inside: ColumnSpan = { ...span, min: Math.max(span.min, left), max: Math.min(span.max, right) };
    if (want === 0) {
      delete inside.styleId;
    } else {
      inside.styleId = want;
    }
    if (spanHoldsSomething(inside)) {
      out.push(inside);
    }
    if (span.max > right) {
      out.push({ ...span, min: right + 1 });
    }
  }
  if (col <= right) {
    gap(col, right);
  }
  return changed ? out : columns;
}

// --- ranges -----------------------------------------------------------------------

/** `sheet` with rows replaced (undefined: removed) or added, by index —
 * untouched rows keep their identity. */
function withRows(sheet: Sheet, replaced: ReadonlyMap<number, Row | undefined>): Row[] {
  const keys = [...replaced.keys()].toSorted((a, b) => a - b);
  const rows: Row[] = [];
  let k = 0;
  for (const row of sheet.rows) {
    while (k < keys.length && keys[k]! < row.index) {
      const added = replaced.get(keys[k]!);
      if (added) {
        rows.push(added);
      }
      k += 1;
    }
    if (k < keys.length && keys[k] === row.index) {
      const next = replaced.get(row.index);
      if (next) {
        rows.push(next);
      }
      k += 1;
    } else {
      rows.push(row);
    }
  }
  for (; k < keys.length; k += 1) {
    const added = replaced.get(keys[k]!);
    if (added) {
      rows.push(added);
    }
  }
  return rows;
}

/** Collects a range's row results. */
class RowResults {
  readonly replaced = new Map<number, Row | undefined>();

  record(old: Row | undefined, index: number, next: Row | undefined): void {
    if (next !== old) {
      this.replaced.set(index, next);
    }
  }
}

/** Calls `stored(row)` for every stored row in [top, bottom] and
 * `unstored(from, to)` for every stretch of rows between them, in order. */
function walkRows(sheet: Sheet, top: number, bottom: number, stored: (row: Row) => void, unstored: (from: number, to: number) => void): void {
  const rows = sheet.rows;
  let expected = top;
  for (let r = firstRowAtOrAfter(rows, top); r < rows.length && rows[r]!.index <= bottom; r += 1) {
    const row = rows[r]!;
    if (row.index > expected) {
      unstored(expected, row.index - 1);
    }
    stored(row);
    expected = row.index + 1;
  }
  if (expected <= bottom) {
    unstored(expected, bottom);
  }
}

/** One range rewritten by `op` (see the module header): a new Sheet, or
 * `sheet` itself when nothing changes. */
export function restyleRange(sheet: Sheet, range: GridRange, op: RestyleOp, budget: MaterialiseBudget): Sheet {
  switch (rangeClass(range)) {
    case "ordinary":
      return ordinary(sheet, range, op, budget);
    case "rows":
      return wholeRows(sheet, range, op, budget);
    case "columns":
      return wholeColumns(sheet, range, op, budget, false);
    case "sheet":
      return wholeColumns(sheet, range, op, budget, true);
  }
}

function finish(sheet: Sheet, results: RowResults, columns: ColumnSpan[] = sheet.columns): Sheet {
  if (results.replaced.size === 0 && columns === sheet.columns) {
    return sheet;
  }
  const next: Sheet = { ...sheet, rows: results.replaced.size === 0 ? sheet.rows : withRows(sheet, results.replaced) };
  if (columns !== sheet.columns) {
    next.columns = columns;
  }
  return next;
}

/** The fills of a run list: each run whose rewritten default differs from
 * `shown(run)` — what its positions show without a cell after the command. */
function fillsOf(runs: readonly ColumnRun[], op: RestyleOp, shown: (run: ColumnRun) => number): Fill[] {
  const fills: Fill[] = [];
  for (const run of runs) {
    const want = op.next(run.styleId ?? 0);
    if (want !== shown(run)) {
      fills.push({ min: run.min, max: run.max, styleId: want });
    }
  }
  return fills;
}

function ordinary(sheet: Sheet, range: GridRange, op: RestyleOp, budget: MaterialiseBudget): Sheet {
  const results = new RowResults();
  const columns = sheet.columns;
  const runs = columnRuns(columns, range.left, range.right);
  const spanStyle = (col: number): number => columnSpanAt(columns, col)?.styleId ?? 0;
  // An unstyled row's empty positions show their column's default (worked
  // out on first need, so nothing is interned for rows the range lacks).
  let fills: Fill[] | undefined;
  const unstyledFills = (): Fill[] => (fills ??= op.storedOnly ? [] : fillsOf(runs, op, (run) => run.styleId ?? 0));
  const unstyled = (cell: Cell): number => op.next(cell.styleId ?? spanStyle(cell.col));
  walkRows(
    sheet,
    range.top,
    range.bottom,
    (row) => {
      const rowStyle = row.styleId;
      let rewrite: RowRewrite;
      if (rowStyle !== undefined) {
        const want = op.next(rowStyle);
        rewrite = {
          left: range.left,
          right: range.right,
          cellWant: (cell) => op.next(cell.styleId ?? rowStyle),
          fallbackAfter: () => rowStyle,
          fills: op.storedOnly || want === rowStyle ? [] : [{ min: range.left, max: range.right, styleId: want }],
        };
      } else {
        rewrite = { left: range.left, right: range.right, cellWant: unstyled, fallbackAfter: spanStyle, fills: unstyledFills() };
      }
      if (op.cell) {
        rewrite.cellHook = op.cell;
      }
      results.record(row, row.index, rewriteRow(row, row.index, rewrite, budget));
    },
    (from, to) => {
      const stretch = unstyledFills();
      if (stretch.length === 0) {
        return;
      }
      const perRow = stretch.reduce((sum, fill) => sum + fill.max - fill.min + 1, 0) + 1;
      budget.take(perRow * (to - from + 1));
      for (let index = from; index <= to; index += 1) {
        results.replaced.set(index, newRow(index, stretch, undefined));
      }
    },
  );
  return finish(sheet, results);
}

/** A row that did not exist: `styleId` (when set) and a style-only cell for
 * every position of `fills` — already taken from the budget. */
function newRow(index: number, fills: readonly Fill[], styleId: number | undefined): Row {
  const cells: Cell[] = [];
  for (const fill of fills) {
    for (let col = fill.min; col <= fill.max; col += 1) {
      cells.push({ col, value: null, styleId: fill.styleId });
    }
  }
  const row: Row = { index, cells };
  if (styleId !== undefined) {
    row.styleId = styleId;
  }
  return row;
}

function wholeRows(sheet: Sheet, range: GridRange, op: RestyleOp, budget: MaterialiseBudget): Sheet {
  const results = new RowResults();
  const columns = sheet.columns;
  const spanStyle = (col: number): number => columnSpanAt(columns, col)?.styleId ?? 0;
  // A row without a style of its own: with the rewrite of 0 as its style
  // (when that changes anything), the styled spans' positions get cells
  // unless their rewrite equals the row's new style. Worked out on first
  // need, so nothing is interned for rows the range lacks.
  let plan: { style: number | undefined; fills: Fill[] } | undefined;
  const unstyled = (): { style: number | undefined; fills: Fill[] } => {
    if (!plan) {
      const zero = op.storedOnly ? 0 : op.next(0);
      const styledRuns = columnRuns(columns, 0, MAX_COLS - 1).filter((run) => run.styleId !== undefined);
      plan = {
        style: zero === 0 ? undefined : zero,
        fills: op.storedOnly ? [] : fillsOf(styledRuns, op, (run) => (zero === 0 ? run.styleId! : zero)),
      };
    }
    return plan;
  };
  walkRows(
    sheet,
    range.top,
    range.bottom,
    (row) => {
      const rowStyle = row.styleId;
      let rewrite: RowRewrite;
      if (rowStyle !== undefined) {
        const want = op.next(rowStyle);
        rewrite = {
          left: 0,
          right: MAX_COLS - 1,
          cellWant: (cell) => op.next(cell.styleId ?? rowStyle),
          fallbackAfter: () => want,
          fills: [],
          rowStyle: { value: want },
        };
      } else {
        const { style, fills } = unstyled();
        rewrite = {
          left: 0,
          right: MAX_COLS - 1,
          cellWant: (cell) => op.next(cell.styleId ?? spanStyle(cell.col)),
          fallbackAfter: (col) => style ?? spanStyle(col),
          fills,
        };
        if (style !== undefined) {
          rewrite.rowStyle = { value: style };
        }
      }
      if (op.cell) {
        rewrite.cellHook = op.cell;
      }
      results.record(row, row.index, rewriteRow(row, row.index, rewrite, budget));
    },
    (from, to) => {
      if (op.storedOnly) {
        return;
      }
      const { style, fills } = unstyled();
      if (style === undefined && fills.length === 0) {
        return;
      }
      const perRow = fills.reduce((sum, fill) => sum + fill.max - fill.min + 1, 0) + 1;
      budget.take(perRow * (to - from + 1));
      for (let index = from; index <= to; index += 1) {
        results.replaced.set(index, newRow(index, fills, style));
      }
    },
  );
  return finish(sheet, results);
}

function wholeColumns(sheet: Sheet, range: GridRange, op: RestyleOp, budget: MaterialiseBudget, wholeSheet: boolean): Sheet {
  const results = new RowResults();
  const before = sheet.columns;
  const after = rewriteSpans(before, range.left, range.right, op);
  const oldSpan = (col: number): number => columnSpanAt(before, col)?.styleId ?? 0;
  const newSpan = (col: number): number => columnSpanAt(after, col)?.styleId ?? 0;
  for (const row of sheet.rows) {
    const rowStyle = row.styleId;
    let rewrite: RowRewrite;
    if (rowStyle === undefined) {
      rewrite = {
        left: range.left,
        right: range.right,
        cellWant: (cell) => op.next(cell.styleId ?? oldSpan(cell.col)),
        fallbackAfter: newSpan,
        fills: [],
      };
    } else if (wholeSheet) {
      // The whole sheet rewrites the row's own style: no position's winner
      // changes, so nothing is materialised.
      const want = op.next(rowStyle);
      rewrite = {
        left: range.left,
        right: range.right,
        cellWant: (cell) => op.next(cell.styleId ?? rowStyle),
        fallbackAfter: () => want,
        fills: [],
        rowStyle: { value: want },
      };
    } else {
      // The row still wins over the column: its positions in the columns
      // need cells carrying the rewritten row style.
      const want = op.next(rowStyle);
      rewrite = {
        left: range.left,
        right: range.right,
        cellWant: (cell) => op.next(cell.styleId ?? rowStyle),
        fallbackAfter: () => rowStyle,
        fills: op.storedOnly || want === rowStyle ? [] : [{ min: range.left, max: range.right, styleId: want }],
      };
    }
    if (op.cell) {
      rewrite.cellHook = op.cell;
    }
    results.record(row, row.index, rewriteRow(row, row.index, rewrite, budget));
  }
  return finish(sheet, results, after);
}

// --- explicit positions (the format painter's tiling) ---------------------------

/** Every position of `range` set to `idAt(row, col)` (explicit and sparse,
 * as above) — the painter tiling a pattern over an ordinary target. The
 * caller bounds the range's area. */
export function restylePositions(sheet: Sheet, range: GridRange, idAt: (row: number, col: number) => number, budget: MaterialiseBudget): Sheet {
  const results = new RowResults();
  const columns = sheet.columns;
  const spanStyle = (col: number): number => columnSpanAt(columns, col)?.styleId ?? 0;
  const fillsFor = (rowIndex: number, fallback: (col: number) => number): Fill[] => {
    const fills: Fill[] = [];
    for (let col = range.left; col <= range.right; col += 1) {
      const want = idAt(rowIndex, col);
      if (want === fallback(col)) {
        continue;
      }
      const last = fills.at(-1);
      if (last && last.max === col - 1 && last.styleId === want) {
        last.max = col;
      } else {
        fills.push({ min: col, max: col, styleId: want });
      }
    }
    return fills;
  };
  walkRows(
    sheet,
    range.top,
    range.bottom,
    (row) => {
      const rowStyle = row.styleId;
      const fallback = rowStyle === undefined ? spanStyle : (): number => rowStyle;
      const rewrite: RowRewrite = {
        left: range.left,
        right: range.right,
        cellWant: (cell) => idAt(row.index, cell.col),
        fallbackAfter: fallback,
        fills: fillsFor(row.index, fallback),
      };
      results.record(row, row.index, rewriteRow(row, row.index, rewrite, budget));
    },
    (from, to) => {
      for (let index = from; index <= to; index += 1) {
        const fills = fillsFor(index, spanStyle);
        if (fills.length === 0) {
          continue;
        }
        budget.take(fills.reduce((sum, fill) => sum + fill.max - fill.min + 1, 0) + 1);
        results.replaced.set(index, newRow(index, fills, undefined));
      }
    },
  );
  return finish(sheet, results);
}
