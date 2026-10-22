// Merged ranges (SPEC-117 "Merges"): the index every renderer, the selection
// model and the overflow pass query merges through. A sheet's `merges` are
// plain GridRanges (model.ts); this module sorts them by top row and keeps a
// running maximum of their bottom rows, so "which merges touch this window"
// is a binary search plus a walk back over the candidates — O(log n + k) —
// never a scan of every merge per cell. paneMerges is that query as the
// grid's cell layer asks it, pane by pane.
import type { GridRange } from "./address";

export interface MergeIndex {
  /** The merges, sorted by top row (then left column). */
  ranges: GridRange[];
  /** maxBottom[i] = the largest `bottom` among ranges[0..i] — non-decreasing,
   * so a walk back from a candidate stops the moment it falls above a row. */
  maxBottom: number[];
}

const EMPTY_INDEX: MergeIndex = { ranges: [], maxBottom: [] };

export function buildMergeIndex(merges: readonly GridRange[] | undefined): MergeIndex {
  if (!merges || merges.length === 0) {
    return EMPTY_INDEX;
  }
  const ranges = merges.toSorted((a, b) => a.top - b.top || a.left - b.left);
  const maxBottom: number[] = [];
  let running = -1;
  for (const range of ranges) {
    running = Math.max(running, range.bottom);
    maxBottom.push(running);
  }
  return { ranges, maxBottom };
}

/** Whether two inclusive ranges share a cell. */
export function rangesIntersect(a: GridRange, b: GridRange): boolean {
  return a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
}

/** The index of the last merge whose top is at or above `row`, or -1. */
function lastStartingAtOrAbove(index: MergeIndex, row: number): number {
  let lo = 0;
  let hi = index.ranges.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (index.ranges[mid]!.top <= row) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}

/** Every merge that shares a cell with `range`, in index order (by top). */
export function mergesIn(index: MergeIndex, range: GridRange): GridRange[] {
  const out: GridRange[] = [];
  for (let i = lastStartingAtOrAbove(index, range.bottom); i >= 0 && index.maxBottom[i]! >= range.top; i -= 1) {
    const merge = index.ranges[i]!;
    if (rangesIntersect(merge, range)) {
      out.push(merge);
    }
  }
  return out.toReversed();
}

/** The merge covering (row, col), or undefined. */
export function mergeAt(index: MergeIndex, row: number, col: number): GridRange | undefined {
  for (let i = lastStartingAtOrAbove(index, row); i >= 0 && index.maxBottom[i]! >= row; i -= 1) {
    const merge = index.ranges[i]!;
    if (merge.bottom >= row && merge.left <= col && merge.right >= col) {
      return merge;
    }
  }
  return undefined;
}

/** A merge one pane of the grid draws, and the row whose element hosts it. */
export interface PaneMerge<T> {
  range: GridRange;
  host: T;
}

/** The merges one pane of the grid draws (SPEC-117 "Merges", "Frozen panes
 * and hidden rows and columns"): each merge sharing a cell with the pane's
 * `window` that shows some of itself there, hosted by the first of the
 * pane's `rendered` rows (its rows of non-zero height, ascending by index)
 * inside the merge — so an anchor row that is hidden, or scrolled above the
 * window, hands the merge to the first row that shows it. A merge whose rows
 * or columns inside the window are all hidden is not drawn there. */
export function paneMerges<T extends { index: number }>(
  index: MergeIndex,
  window: GridRange,
  rendered: readonly T[],
  colLeft: (col: number) => number,
): PaneMerge<T>[] {
  const out: PaneMerge<T>[] = [];
  for (const range of mergesIn(index, window)) {
    let lo = 0;
    let hi = rendered.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (rendered[mid]!.index < range.top) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }
    const host = rendered[lo];
    const shown = colLeft(Math.min(range.right, window.right) + 1) - colLeft(Math.max(range.left, window.left));
    if (host !== undefined && host.index <= range.bottom && shown > 0) {
      out.push({ range, host });
    }
  }
  return out;
}

/** The merges a reader keeps, in source order (SPEC-117 "Model additions"):
 * ranges of two or more cells that overlap no earlier kept merge. A later
 * overlapping or single-cell merge is not modelled — it stays in the source.
 * (A quadratic check: real sheets carry hundreds of merges, not millions.) */
export function acceptMerges(candidates: readonly GridRange[]): GridRange[] {
  const kept: GridRange[] = [];
  for (const range of candidates) {
    if (range.top === range.bottom && range.left === range.right) {
      continue;
    }
    if (kept.some((other) => rangesIntersect(other, range))) {
      continue;
    }
    kept.push(range);
  }
  return kept;
}
