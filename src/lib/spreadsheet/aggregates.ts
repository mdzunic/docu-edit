// Status-bar aggregates (SPEC-117 "Aggregates"): Average, Count and Sum of
// the selection, over stored cells only.
//
// - The union of the selection's ranges is cut into row bands, each with its
//   merged column spans, so every stored cell is visited once even where
//   ranges overlap; hidden rows and columns are excluded.
// - `count` = cells whose value is not null; `numericCount`/`sum` take
//   numbers only (booleans and numeric-looking strings are counted, never
//   summed; dates are numbers); any counted error sets `hasError`.
// - A band is reached by a binary search of the sheet's sparse rows and a
//   span by a binary search of each row's sorted cells, so a whole-column or
//   whole-sheet selection costs the stored cells it holds — never the
//   17-billion-cell grid. A selection over 100,000 stored cells (by
//   selectionStoredCellEstimate) is the caller's to defer through the idle
//   scheduler.
// - The texts appear only when the selection spans more than one cell and
//   count ≥ 1: "Count: n" always, "Average: …" and "Sum: …" only when
//   numericCount ≥ 1 and there is no error; numbers through generalText
//   (SPEC-118 formats them).
import type { GridRange } from "./address";
import { generalText } from "./general";
import { mergeAt, type MergeIndex } from "./merges";
import { isCellError, MAX_COLS, MAX_ROWS, type Cell, type Row, type Sheet } from "./model";
import type { NavContext } from "./selection";

export interface SelectionAggregates {
  count: number;
  numericCount: number;
  sum: number;
  hasError: boolean;
}

/** The status bar's texts; a field is absent when it is not shown. */
export interface AggregateTexts {
  average?: string;
  count?: string;
  sum?: string;
}

/** Rows top..bottom of the union, with the disjoint, ascending column spans
 * the union covers on each of them. */
interface RowBand {
  top: number;
  bottom: number;
  spans: Array<readonly [number, number]>;
}

/** A range normalised and intersected with the grid; null when it lies
 * wholly outside it. */
function clampToGrid(range: GridRange): GridRange | null {
  const top = Math.max(0, Math.min(range.top, range.bottom));
  const bottom = Math.min(MAX_ROWS - 1, Math.max(range.top, range.bottom));
  const left = Math.max(0, Math.min(range.left, range.right));
  const right = Math.min(MAX_COLS - 1, Math.max(range.left, range.right));
  return top <= bottom && left <= right ? { top, left, bottom, right } : null;
}

/** Sorted column intervals with overlapping and adjacent ones joined. */
function joinSpans(spans: Array<readonly [number, number]>): Array<readonly [number, number]> {
  const sorted = spans.toSorted((a, b) => a[0] - b[0]);
  const joined: Array<[number, number]> = [];
  for (const [left, right] of sorted) {
    const last = joined.at(-1);
    if (last && left <= last[1] + 1) {
      last[1] = Math.max(last[1], right);
    } else {
      joined.push([left, right]);
    }
  }
  return joined;
}

function sameSpans(a: ReadonlyArray<readonly [number, number]>, b: ReadonlyArray<readonly [number, number]>): boolean {
  return a.length === b.length && a.every((span, i) => span[0] === b[i]![0] && span[1] === b[i]![1]);
}

/** The union of `ranges` as disjoint row bands, top to bottom. */
function rowBands(ranges: readonly GridRange[]): RowBand[] {
  const valid = ranges.map(clampToGrid).filter((range): range is GridRange => range !== null);
  const cuts = [...new Set(valid.flatMap((range) => [range.top, range.bottom + 1]))].toSorted((a, b) => a - b);
  const bands: RowBand[] = [];
  for (let k = 0; k + 1 < cuts.length; k += 1) {
    const top = cuts[k]!;
    const bottom = cuts[k + 1]! - 1;
    const covering = valid.filter((range) => range.top <= top && range.bottom >= top);
    if (covering.length === 0) {
      continue;
    }
    const spans = joinSpans(covering.map((range) => [range.left, range.right] as const));
    const previous = bands.at(-1);
    if (previous && previous.bottom + 1 === top && sameSpans(previous.spans, spans)) {
      previous.bottom = bottom;
    } else {
      bands.push({ top, bottom, spans });
    }
  }
  return bands;
}

/** The first stored row at or below `index` — a binary search. */
function firstRowAtOrAfter(rows: readonly Row[], index: number): number {
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

/** The first stored cell at or right of `col` — a binary search. */
function firstCellAtOrAfter(cells: readonly Cell[], col: number): number {
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

/** Count, numeric count, sum and the error flag over the stored cells in
 * the union of `ranges` — each cell once, hidden rows and columns excluded. */
export function selectionAggregates(
  sheet: Sheet,
  ranges: readonly GridRange[],
  nav: Pick<NavContext, "isRowHidden" | "isColHidden">,
): SelectionAggregates {
  const totals: SelectionAggregates = { count: 0, numericCount: 0, sum: 0, hasError: false };
  const rows = sheet.rows;
  for (const band of rowBands(ranges)) {
    for (let i = firstRowAtOrAfter(rows, band.top); i < rows.length && rows[i]!.index <= band.bottom; i += 1) {
      const row = rows[i]!;
      const cells = row.cells;
      if (cells.length === 0 || nav.isRowHidden(row.index)) {
        continue;
      }
      for (const [left, right] of band.spans) {
        for (let j = firstCellAtOrAfter(cells, left); j < cells.length && cells[j]!.col <= right; j += 1) {
          const { col, value } = cells[j]!;
          if (value === null || nav.isColHidden(col)) {
            continue;
          }
          totals.count += 1;
          if (typeof value === "number") {
            totals.numericCount += 1;
            totals.sum += value;
          } else if (isCellError(value)) {
            totals.hasError = true;
          }
        }
      }
    }
  }
  return totals;
}

/** A cheap upper bound of the stored cells selectionAggregates would visit
 * for `ranges` — one pass over the stored rows the selection crosses,
 * without touching a cell — so the caller can defer a large computation to
 * idle time. Stops counting once the bound exceeds `cap`. */
export function selectionStoredCellEstimate(sheet: Sheet, ranges: readonly GridRange[], cap = Number.POSITIVE_INFINITY): number {
  const rows = sheet.rows;
  let estimate = 0;
  for (const band of rowBands(ranges)) {
    const width = band.spans.reduce((total, [left, right]) => total + right - left + 1, 0);
    for (let i = firstRowAtOrAfter(rows, band.top); i < rows.length && rows[i]!.index <= band.bottom; i += 1) {
      estimate += Math.min(rows[i]!.cells.length, width);
      if (estimate > cap) {
        return estimate;
      }
    }
  }
  return estimate;
}

/** Whether a selection's ranges span more than one cell — the condition for
 * showing aggregates at all. With `merges`, a single range that is exactly
 * one merge counts as one (merged) cell. */
export function selectionSpansMultipleCells(ranges: readonly GridRange[], merges?: MergeIndex): boolean {
  if (ranges.length !== 1) {
    return ranges.length > 1;
  }
  const range = ranges[0]!;
  if (range.top === range.bottom && range.left === range.right) {
    return false;
  }
  const merge = merges ? mergeAt(merges, range.top, range.left) : undefined;
  return !(merge && merge.top === range.top && merge.left === range.left && merge.bottom === range.bottom && merge.right === range.right);
}

/** The status bar's texts for `agg`: nothing unless the selection spans more
 * than one cell and count ≥ 1; then "Count: n", plus "Average: …" and
 * "Sum: …" when numericCount ≥ 1 and no counted cell is an error. */
export function formatAggregates(agg: SelectionAggregates, isMultiCell: boolean): AggregateTexts {
  if (!isMultiCell || agg.count < 1) {
    return {};
  }
  const texts: AggregateTexts = { count: `Count: ${generalText(agg.count)}` };
  if (agg.numericCount >= 1 && !agg.hasError) {
    texts.average = `Average: ${generalText(agg.sum / agg.numericCount)}`;
    texts.sum = `Sum: ${generalText(agg.sum)}`;
  }
  return texts;
}
