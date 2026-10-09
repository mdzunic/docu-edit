// The read-only selection model (SPEC-117 "Selection", "Keyboard map",
// "Pointer"): what is selected on a sheet, and every move the keyboard and
// the pointer make on it, as pure functions of the selection and a
// NavContext. A selection is plain JSON — it is persisted as view state
// (Sheet.selection) — so every function returns fresh objects, never one its
// input or the sheet's merges hold.
//
// Rules every function keeps:
// - `active` is inside `ranges[activeRangeIndex]`, visible, and — when in a
//   merge — the merge's top-left. A merge is visible when it has a visible
//   row and a visible column (its top-left cell may itself be hidden: the
//   merge then shows at its visible extent, as the grid renders it).
// - Ranges are merge-snapped: expanded until no merge crosses their edge.
// - Moving collapses to one cell (a merge's whole range); from a merge a
//   step leaves from the merge's far edge, entering one lands on its
//   top-left; hidden rows and columns are skipped; the grid's edges clamp.
// - Extending keeps `active` as the anchor and moves the range's corner
//   farthest from it on each axis (on a tie, the bottom/right edge). A step
//   a merge swallows is retried one row or column further, so Shift+arrow
//   never stalls against a merge.
// - Searches are sparse: a column is walked over the sheet's stored rows and
//   a row over its sorted cells (binary searches, then monotone cursors),
//   never position by position across the 1,048,576 × 16,384 grid.
import { findRow, getCell, usedRange, type CellRef, type GridRange } from "./address";
import { mergeAt, mergesIn, type MergeIndex } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type Row, type Sheet } from "./model";

export interface SheetSelection {
  active: CellRef;
  ranges: GridRange[];
  activeRangeIndex: number;
}

export type Direction = "up" | "down" | "left" | "right";

export interface NavContext {
  sheet: Sheet;
  merges: MergeIndex;
  isRowHidden(r: number): boolean;
  isColHidden(c: number): boolean;
  pageRows: number; // visible rows/cols for PageUp/PageDown
  pageCols: number;
}

/** A key press as the keyboard map reads it. `mod` is the platform's
 * primary modifier (Cmd on macOS, Ctrl elsewhere — the caller computes it
 * from core/platform.ts); `ctrl` is the literal Control key, which
 * Ctrl+Space reads on every platform. */
export interface NavKey {
  key: string; // KeyboardEvent.key
  shift: boolean;
  mod: boolean;
  alt: boolean;
  ctrl: boolean;
}

const LAST_ROW = MAX_ROWS - 1;
const LAST_COL = MAX_COLS - 1;
const WHOLE_SHEET: Readonly<GridRange> = { top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL };

type Step = 1 | -1;

/** One axis of the grid: its last index and which of its lines are hidden. */
interface Axis {
  last: number;
  hidden(index: number): boolean;
}

function rowAxis(nav: NavContext): Axis {
  return { last: LAST_ROW, hidden: (r) => nav.isRowHidden(r) };
}

function colAxis(nav: NavContext): Axis {
  return { last: LAST_COL, hidden: (c) => nav.isColHidden(c) };
}

// --- plain data helpers ---------------------------------------------------

/** `value` as an index in 0..last: truncated and clamped (NaN → 0). */
function clampIndex(value: number, last: number): number {
  if (Number.isNaN(value)) {
    return 0;
  }
  return Math.min(last, Math.max(0, Math.trunc(value)));
}

function clampRef(ref: CellRef): CellRef {
  return { row: clampIndex(ref.row, LAST_ROW), col: clampIndex(ref.col, LAST_COL) };
}

function copyRange(range: GridRange): GridRange {
  return { top: range.top, left: range.left, bottom: range.bottom, right: range.right };
}

function sameRange(a: GridRange, b: GridRange): boolean {
  return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

function cloneSelection(sel: SheetSelection): SheetSelection {
  return { active: { row: sel.active.row, col: sel.active.col }, ranges: sel.ranges.map(copyRange), activeRangeIndex: sel.activeRangeIndex };
}

function hasValue(value: CellValue | undefined): boolean {
  return value !== null && value !== undefined;
}

/** The first index whose position is at or after `target` (positions ascending). */
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

/** The row's cell at `col` — a binary search of its sorted cells. */
function cellInRow(row: Row, col: number): Cell | undefined {
  const cells = row.cells;
  const cell = cells[lowerBound(cells.length, (i) => cells[i]!.col, col)];
  return cell?.col === col ? cell : undefined;
}

/** Whether (row, col) lies inside `range` (inclusive). */
export function rangeContains(range: GridRange, row: number, col: number): boolean {
  return row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
}

/** Whether any range of the selection holds (row, col). */
export function isCellSelected(sel: SheetSelection, row: number, col: number): boolean {
  return sel.ranges.some((range) => rangeContains(range, row, col));
}

/** Whole rows top..bottom (either order), clamped to the grid. */
export function wholeRows(top: number, bottom: number): GridRange {
  const a = clampIndex(top, LAST_ROW);
  const b = clampIndex(bottom, LAST_ROW);
  return { top: Math.min(a, b), left: 0, bottom: Math.max(a, b), right: LAST_COL };
}

/** Whole columns left..right (either order), clamped to the grid. */
export function wholeColumns(left: number, right: number): GridRange {
  const a = clampIndex(left, LAST_COL);
  const b = clampIndex(right, LAST_COL);
  return { top: 0, left: Math.min(a, b), bottom: LAST_ROW, right: Math.max(a, b) };
}

// --- visibility and merges -------------------------------------------------

/** The first visible index after `from` in direction `step`, or -1 past the
 * grid's edge. */
function nextVisible(from: number, step: Step, axis: Axis): number {
  for (let t = from + step; t >= 0 && t <= axis.last; t += step) {
    if (!axis.hidden(t)) {
      return t;
    }
  }
  return -1;
}

/** The first visible index walking from `near` to `far` (inclusive), or -1. */
function firstVisibleBetween(near: number, far: number, step: Step, axis: Axis): number {
  for (let t = near; step > 0 ? t <= far : t >= far; t += step) {
    if (!axis.hidden(t)) {
      return t;
    }
  }
  return -1;
}

/** `index` when visible, else the nearest visible index (the following one
 * on a tie); `index` itself when the whole axis is hidden. */
function nearestVisible(index: number, axis: Axis): number {
  if (!axis.hidden(index)) {
    return index;
  }
  for (let d = 1; index + d <= axis.last || index - d >= 0; d += 1) {
    if (index + d <= axis.last && !axis.hidden(index + d)) {
      return index + d;
    }
    if (index - d >= 0 && !axis.hidden(index - d)) {
      return index - d;
    }
  }
  return index;
}

/** The line a walk leaves a unit along: `preferred` when visible, else the
 * unit's first visible line between lo and hi (else `preferred`). */
function lineIndex(preferred: number, lo: number, hi: number, axis: Axis): number {
  if (!axis.hidden(preferred)) {
    return preferred;
  }
  const first = firstVisibleBetween(lo, hi, 1, axis);
  return first >= 0 ? first : preferred;
}

/** The unit a cell belongs to: the merge covering it (a fresh copy) or the
 * cell itself. */
function unitOf(ref: CellRef, merges: MergeIndex): GridRange {
  const merge = mergeAt(merges, ref.row, ref.col);
  return merge ? copyRange(merge) : { top: ref.row, left: ref.col, bottom: ref.row, right: ref.col };
}

/** Where selecting `ref` lands: the top-left of the merge covering it, else
 * the cell. */
function landing(ref: CellRef, merges: MergeIndex): CellRef {
  const merge = mergeAt(merges, ref.row, ref.col);
  return merge ? { row: merge.top, col: merge.left } : { row: ref.row, col: ref.col };
}

/** Whether a unit (a cell or a merge) shows: it has a visible row and a
 * visible column. */
function unitVisible(unit: GridRange, nav: NavContext): boolean {
  return firstVisibleBetween(unit.top, unit.bottom, 1, rowAxis(nav)) >= 0 && firstVisibleBetween(unit.left, unit.right, 1, colAxis(nav)) >= 0;
}

/** The visible unit at `ref`, else the one at the nearest visible cell. */
function visibleUnitNear(ref: CellRef, nav: NavContext): GridRange {
  const unit = unitOf(ref, nav.merges);
  if (unitVisible(unit, nav)) {
    return unit;
  }
  return unitOf({ row: nearestVisible(ref.row, rowAxis(nav)), col: nearestVisible(ref.col, colAxis(nav)) }, nav.merges);
}

/** A one-unit selection: the unit's top-left active, the unit its range. */
function onUnit(unit: GridRange): SheetSelection {
  return { active: { row: unit.top, col: unit.left }, ranges: [unit], activeRangeIndex: 0 };
}

/** The first visible cell of the sheet (A1 unless row 1 or column A is hidden). */
function firstVisibleCell(nav: NavContext): CellRef {
  return {
    row: Math.max(0, firstVisibleBetween(0, LAST_ROW, 1, rowAxis(nav))),
    col: Math.max(0, firstVisibleBetween(0, LAST_COL, 1, colAxis(nav))),
  };
}

/** The last visible index at or before `index` (else the first after it). */
function lastVisibleAtOrBefore(index: number, axis: Axis): number {
  for (let t = index; t >= 0; t -= 1) {
    if (!axis.hidden(t)) {
      return t;
    }
  }
  return Math.max(0, firstVisibleBetween(index, axis.last, 1, axis));
}

/** The bottom-right of the used range, moved up/left off hidden lines; the
 * first visible cell when the sheet stores nothing (Mod+End). */
function lastUsedCell(nav: NavContext): CellRef {
  const used = usedRange(nav.sheet);
  if (!used) {
    return firstVisibleCell(nav);
  }
  return { row: lastVisibleAtOrBefore(used.bottom, rowAxis(nav)), col: lastVisibleAtOrBefore(used.right, colAxis(nav)) };
}

/** `range` expanded until no merge crosses its edge — repeated to a fixed
 * point, since taking in one merge can reach another. A fresh object. */
export function snapToMerges(range: GridRange, merges: MergeIndex): GridRange {
  const snapped = copyRange(range);
  for (let grown = true; grown; ) {
    grown = false;
    for (const merge of mergesIn(merges, snapped)) {
      if (merge.top < snapped.top) {
        snapped.top = merge.top;
        grown = true;
      }
      if (merge.left < snapped.left) {
        snapped.left = merge.left;
        grown = true;
      }
      if (merge.bottom > snapped.bottom) {
        snapped.bottom = merge.bottom;
        grown = true;
      }
      if (merge.right > snapped.right) {
        snapped.right = merge.right;
        grown = true;
      }
    }
  }
  return snapped;
}

// --- selections of one cell -----------------------------------------------

/** One cell (clamped to the grid) as the whole selection. Knows nothing of
 * merges or hidden lines — `selectCell` is the snapped form. */
export function singleCell(ref: CellRef): SheetSelection {
  const at = clampRef(ref);
  return { active: at, ranges: [{ top: at.row, left: at.col, bottom: at.row, right: at.col }], activeRangeIndex: 0 };
}

/** A pointer press on `ref`: the cell clamped to the grid, moved to the
 * nearest visible cell when hidden, and snapped to its merge (the merge's
 * range selected, its top-left active). */
export function selectCell(ref: CellRef, nav: NavContext): SheetSelection {
  return onUnit(visibleUnitNear(clampRef(ref), nav));
}

/** The whole sheet, A1 active (the `sheet-select-all` corner). */
export function selectAll(): SheetSelection {
  return { active: { row: 0, col: 0 }, ranges: [copyRange(WHOLE_SHEET)], activeRangeIndex: 0 };
}

/** The whole sheet with the first visible cell active (Mod+A's second press). */
function wholeSheet(nav: NavContext): SheetSelection {
  return { ...selectAll(), active: landing(firstVisibleCell(nav), nav.merges) };
}

// --- moving -----------------------------------------------------------------

function stepCount(steps: number): number {
  return Number.isFinite(steps) ? Math.max(1, Math.trunc(steps)) : 1;
}

/** Moves the active cell `steps` units in `dir` and collapses the selection
 * to it. Each step leaves the current unit from its far edge, skips hidden
 * rows/columns and lands on the next unit's top-left; the walk keeps to the
 * line it started on and stops at the grid's edge. */
export function moveActive(sel: SheetSelection, dir: Direction, nav: NavContext, steps = 1): SheetSelection {
  const start = clampRef(sel.active);
  const vertical = dir === "up" || dir === "down";
  const step: Step = dir === "down" || dir === "right" ? 1 : -1;
  const axis = vertical ? rowAxis(nav) : colAxis(nav);
  let unit = unitOf(start, nav.merges);
  const line = vertical ? lineIndex(start.col, unit.left, unit.right, colAxis(nav)) : lineIndex(start.row, unit.top, unit.bottom, rowAxis(nav));
  for (let i = 0; i < stepCount(steps); i += 1) {
    const edge = vertical ? (step > 0 ? unit.bottom : unit.top) : step > 0 ? unit.right : unit.left;
    const next = nextVisible(edge, step, axis);
    if (next < 0) {
      break;
    }
    unit = unitOf(vertical ? { row: next, col: line } : { row: line, col: next }, nav.merges);
  }
  return onUnit(unit);
}

// --- extending --------------------------------------------------------------

/** The selection's active range (a copy), or the active cell's unit when
 * the selection holds no valid one. */
function activeRange(sel: SheetSelection, nav: NavContext): GridRange {
  const range = sel.ranges[sel.activeRangeIndex] as GridRange | undefined;
  return range ? copyRange(range) : unitOf(clampRef(sel.active), nav.merges);
}

/** `sel` with its active range replaced by `range`; `active` and every
 * other range kept (as copies). */
function withActiveRange(sel: SheetSelection, range: GridRange): SheetSelection {
  const active = { row: sel.active.row, col: sel.active.col };
  if (!(sel.ranges[sel.activeRangeIndex] as GridRange | undefined)) {
    return { active, ranges: [range], activeRangeIndex: 0 };
  }
  return { active, ranges: sel.ranges.map((other, i) => (i === sel.activeRangeIndex ? range : copyRange(other))), activeRangeIndex: sel.activeRangeIndex };
}

/** Which edge of `range` moves when extending along one axis: the one
 * farther from the anchor's unit; a tie moves the bottom/right edge. */
function movesHighEdge(lo: number, hi: number, anchorLo: number, anchorHi: number): boolean {
  return hi - anchorHi >= anchorLo - lo;
}

/** The corner of the active range farthest from the anchor on each axis —
 * where Mod+Shift+arrow searches from, and Shift+Home keeps its row. */
function movingCorner(sel: SheetSelection, nav: NavContext): CellRef {
  const range = activeRange(sel, nav);
  const anchor = unitOf(clampRef(sel.active), nav.merges);
  return {
    row: movesHighEdge(range.top, range.bottom, anchor.top, anchor.bottom) ? range.bottom : range.top,
    col: movesHighEdge(range.left, range.right, anchor.left, anchor.right) ? range.right : range.left,
  };
}

/** The range spanning the anchor unit, the edge opposite the moving one and
 * the moving edge's new position `to`, along one axis, snapped. */
function resized(range: GridRange, anchor: GridRange, vertical: boolean, to: number, nav: NavContext): GridRange {
  const lo = vertical ? range.top : range.left;
  const hi = vertical ? range.bottom : range.right;
  const anchorLo = vertical ? anchor.top : anchor.left;
  const anchorHi = vertical ? anchor.bottom : anchor.right;
  const fixed = movesHighEdge(lo, hi, anchorLo, anchorHi) ? lo : hi;
  const first = Math.min(fixed, to, anchorLo);
  const last = Math.max(fixed, to, anchorHi);
  return snapToMerges(vertical ? { ...range, top: first, bottom: last } : { ...range, left: first, right: last }, nav.merges);
}

/** One Shift+arrow step, or null when the moving edge is at the grid's edge. */
function extendOnce(sel: SheetSelection, dir: Direction, nav: NavContext): SheetSelection | null {
  const range = activeRange(sel, nav);
  const anchor = unitOf(clampRef(sel.active), nav.merges);
  const vertical = dir === "up" || dir === "down";
  const step: Step = dir === "down" || dir === "right" ? 1 : -1;
  const axis = vertical ? rowAxis(nav) : colAxis(nav);
  const lo = vertical ? range.top : range.left;
  const hi = vertical ? range.bottom : range.right;
  let moving = movesHighEdge(lo, hi, vertical ? anchor.top : anchor.left, vertical ? anchor.bottom : anchor.right) ? hi : lo;
  for (;;) {
    moving = nextVisible(moving, step, axis);
    if (moving < 0) {
      return null;
    }
    const next = resized(range, anchor, vertical, moving, nav);
    if (!sameRange(next, range)) {
      return withActiveRange(sel, next);
    }
  }
}

/** Shift+arrow: keeps `active` as the anchor and moves the active range's
 * corner farthest from it `steps` visible rows/columns in `dir`, snapping to
 * merges; other ranges are kept. */
export function extendBy(sel: SheetSelection, dir: Direction, nav: NavContext, steps = 1): SheetSelection {
  let current = withActiveRange(sel, activeRange(sel, nav));
  for (let i = 0; i < stepCount(steps); i += 1) {
    const next = extendOnce(current, dir, nav);
    if (!next) {
      break;
    }
    current = next;
  }
  return current;
}

/** Shift+press and drag: the active range becomes the rectangle from the
 * anchor (the active cell's unit) to `target` (clamped), snapped. */
export function extendTo(sel: SheetSelection, target: CellRef, nav: NavContext): SheetSelection {
  const to = clampRef(target);
  const anchor = unitOf(clampRef(sel.active), nav.merges);
  const range = {
    top: Math.min(anchor.top, to.row),
    left: Math.min(anchor.left, to.col),
    bottom: Math.max(anchor.bottom, to.row),
    right: Math.max(anchor.right, to.col),
  };
  return withActiveRange(sel, snapToMerges(range, nav.merges));
}

/** Mod+Shift+arrow: the moving corner's edge goes to edgeInDirection of
 * that corner. */
function extendToEdge(sel: SheetSelection, dir: Direction, nav: NavContext): SheetSelection {
  const range = activeRange(sel, nav);
  const anchor = unitOf(clampRef(sel.active), nav.merges);
  const target = edgeInDirection(movingCorner(sel, nav), dir, nav);
  const vertical = dir === "up" || dir === "down";
  return withActiveRange(sel, resized(range, anchor, vertical, vertical ? target.row : target.col, nav));
}

// --- the data-region edge (Mod+arrow) ---------------------------------------

/** One line of stored cells as a walk sees it — a column's stored rows or a
 * row's stored cells, `position(i)` ascending. */
interface StoredLine {
  length: number;
  position(i: number): number;
  value(i: number): CellValue | undefined;
}

function columnLine(sheet: Sheet, col: number): StoredLine {
  const rows = sheet.rows;
  return { length: rows.length, position: (i) => rows[i]!.index, value: (i) => cellInRow(rows[i]!, col)?.value };
}

function rowLine(sheet: Sheet, row: number): StoredLine {
  const cells = findRow(sheet, row)?.cells ?? [];
  return { length: cells.length, position: (i) => cells[i]!.col, value: (i) => cells[i]!.value };
}

interface LineCursor {
  /** The value stored at `t` (undefined when nothing is). */
  valueAt(t: number): CellValue | undefined;
  /** The nearest position beyond `after` storing a value, or -1. */
  nextValue(after: number): number;
}

/** A monotone cursor over a stored line for a walk from `start` in
 * direction `step`: each query is at or beyond the previous one, so after
 * one binary search a walk costs O(entries passed) — a column walk steps
 * over stored rows, never over row indices. */
function lineCursor(line: StoredLine, step: Step, start: number): LineCursor {
  const position = (i: number): number => line.position(i);
  let i = step > 0 ? lowerBound(line.length, position, start) : lowerBound(line.length, position, start + 1) - 1;
  const inside = (): boolean => i >= 0 && i < line.length;
  const beyond = (a: number, b: number): boolean => (step > 0 ? a > b : a < b);
  return {
    valueAt(t) {
      while (inside() && beyond(t, line.position(i))) {
        i += step;
      }
      return inside() && line.position(i) === t ? line.value(i) : undefined;
    },
    nextValue(after) {
      while (inside() && !beyond(line.position(i), after)) {
        i += step;
      }
      for (; inside(); i += step) {
        if (hasValue(line.value(i))) {
          return line.position(i);
        }
      }
      return -1;
    },
  };
}

/** A merge crossing a walk's line, as positions on the walk axis (`near` is
 * met first). */
interface LineMerge {
  near: number;
  far: number;
  range: GridRange;
}

/** The merges crossing the line from `start` onward, in walk order. */
function lineMerges(merges: MergeIndex, vertical: boolean, line: number, start: number, step: Step): LineMerge[] {
  const lo = step > 0 ? start : 0;
  const hi = step > 0 ? (vertical ? LAST_ROW : LAST_COL) : start;
  const window = vertical ? { top: lo, left: line, bottom: hi, right: line } : { top: line, left: lo, bottom: line, right: hi };
  return mergesIn(merges, window)
    .map((range): LineMerge => {
      const first = vertical ? range.top : range.left;
      const last = vertical ? range.bottom : range.right;
      return step > 0 ? { near: first, far: last, range } : { near: last, far: first, range };
    })
    .toSorted((a, b) => (a.near - b.near) * step);
}

/** Mod+arrow's target from `from`: if the current cell and the next visible
 * one both hold values, the last filled cell before an empty one; otherwise
 * the next filled cell; if there is none, the grid's edge (the last visible
 * row or column). "Filled" is a stored value other than null; a merge counts
 * as one cell holding its top-left's value, is left from its far edge and
 * is returned as its top-left. Hidden rows and columns are skipped. */
export function edgeInDirection(from: CellRef, dir: Direction, nav: NavContext): CellRef {
  const start = clampRef(from);
  const vertical = dir === "up" || dir === "down";
  const step: Step = dir === "down" || dir === "right" ? 1 : -1;
  const walk = vertical ? rowAxis(nav) : colAxis(nav);
  const unit = unitOf(start, nav.merges);
  const line = vertical ? lineIndex(start.col, unit.left, unit.right, colAxis(nav)) : lineIndex(start.row, unit.top, unit.bottom, rowAxis(nav));
  const t0 = vertical ? start.row : start.col;
  const beyond = (a: number, b: number): boolean => (step > 0 ? a > b : a < b);
  const cursor = lineCursor(vertical ? columnLine(nav.sheet, line) : rowLine(nav.sheet, line), step, t0);
  const merges = lineMerges(nav.merges, vertical, line, t0, step);
  let passed = 0; // merges[0..passed) lie wholly behind the walk
  const mergeOn = (t: number): LineMerge | undefined => {
    while (passed < merges.length && beyond(t, merges[passed]!.far)) {
      passed += 1;
    }
    const merge = merges[passed];
    return merge && !beyond(merge.near, t) ? merge : undefined;
  };
  const mergeFilled = (merge: LineMerge): boolean => hasValue(getCell(nav.sheet, merge.range.top, merge.range.left)?.value);
  const filled = (t: number): boolean => {
    const merge = mergeOn(t);
    return merge ? mergeFilled(merge) : hasValue(cursor.valueAt(t));
  };
  const farEdge = (t: number): number => mergeOn(t)?.far ?? t;
  const cellAt = (t: number): CellRef => landing(vertical ? { row: t, col: line } : { row: line, col: t }, nav.merges);

  const startFilled = filled(t0);
  const startFar = farEdge(t0);
  const next = nextVisible(startFar, step, walk);
  if (next < 0) {
    return cellAt(t0);
  }
  if (startFilled && filled(next)) {
    for (let current = next; ; ) {
      const following = nextVisible(farEdge(current), step, walk);
      if (following < 0 || !filled(following)) {
        return cellAt(current);
      }
      current = following;
    }
  }
  // The nearest filled merge ahead (met at its first visible position)…
  let target = -1;
  for (let k = passed; k < merges.length; k += 1) {
    const merge = merges[k]!;
    if (!beyond(merge.near, startFar) || !mergeFilled(merge)) {
      continue;
    }
    const entry = firstVisibleBetween(merge.near, merge.far, step, walk);
    if (entry >= 0) {
      target = entry;
      break;
    }
  }
  // …or a nearer stored value outside every merge, on a visible line.
  for (let t = cursor.nextValue(startFar); t >= 0 && (target < 0 || beyond(target, t)); t = cursor.nextValue(t)) {
    const merge = mergeOn(t);
    if (merge) {
      t = merge.far;
      continue;
    }
    if (!walk.hidden(t)) {
      target = t;
      break;
    }
  }
  if (target >= 0) {
    return cellAt(target);
  }
  for (let t = step > 0 ? walk.last : 0; beyond(t, startFar); t -= step) {
    if (!walk.hidden(t)) {
      return cellAt(t);
    }
  }
  return cellAt(t0);
}

// --- the current region (Mod+A) ---------------------------------------------

/** Whether the row stores a value in columns lo..hi. */
function rowHasValue(row: Row, lo: number, hi: number): boolean {
  const cells = row.cells;
  for (let j = lowerBound(cells.length, (i) => cells[i]!.col, lo); j < cells.length && cells[j]!.col <= hi; j += 1) {
    if (cells[j]!.value !== null) {
      return true;
    }
  }
  return false;
}

/** Whether any stored row in top..bottom holds a value in column `col`. */
function columnHasValue(rows: readonly Row[], col: number, top: number, bottom: number): boolean {
  for (let p = lowerBound(rows.length, (i) => rows[i]!.index, top); p < rows.length && rows[p]!.index <= bottom; p += 1) {
    if (hasValue(cellInRow(rows[p]!, col)?.value)) {
      return true;
    }
  }
  return false;
}

/** Grows `region` row by row in direction `step` while the row just outside
 * holds a value within its columns ± 1 (the diagonal) — walking the stored
 * rows, so a gap ends the growth without a lookup. */
function growRows(rows: readonly Row[], region: GridRange, used: GridRange, step: Step): boolean {
  const lo = region.left - 1;
  const hi = region.right + 1;
  let target = step > 0 ? region.bottom + 1 : region.top - 1;
  let p = step > 0 ? lowerBound(rows.length, (i) => rows[i]!.index, target) : lowerBound(rows.length, (i) => rows[i]!.index, target + 1) - 1;
  let grew = false;
  while (target >= used.top && target <= used.bottom) {
    const row = rows[p] as Row | undefined;
    if (!row || row.index !== target || !rowHasValue(row, lo, hi)) {
      break;
    }
    if (step > 0) {
      region.bottom = target;
    } else {
      region.top = target;
    }
    grew = true;
    target += step;
    p += step;
  }
  return grew;
}

/** Grows `region` column by column in direction `step` while the column
 * just outside holds a value within its rows ± 1. */
function growColumns(rows: readonly Row[], region: GridRange, used: GridRange, step: Step): boolean {
  let grew = false;
  for (let target = step > 0 ? region.right + 1 : region.left - 1; target >= used.left && target <= used.right; target += step) {
    if (!columnHasValue(rows, target, region.top - 1, region.bottom + 1)) {
      break;
    }
    if (step > 0) {
      region.right = target;
    } else {
      region.left = target;
    }
    grew = true;
  }
  return grew;
}

/** The current region around `at`: from the cell (its merge), each side
 * grows while the row or column just outside — extended one cell diagonally
 * — holds a value, until nothing grows (merge-snapped, within usedRange). */
export function currentRegion(at: CellRef, nav: NavContext): GridRange {
  const region = unitOf(clampRef(at), nav.merges);
  const used = usedRange(nav.sheet);
  if (!used) {
    return region;
  }
  const rows = nav.sheet.rows;
  for (let grown = true; grown; ) {
    grown = growRows(rows, region, used, -1);
    grown = growRows(rows, region, used, 1) || grown;
    grown = growColumns(rows, region, used, -1) || grown;
    grown = growColumns(rows, region, used, 1) || grown;
    const snapped = snapToMerges(region, nav.merges);
    if (!sameRange(snapped, region)) {
      Object.assign(region, snapped);
      grown = true;
    }
  }
  return region;
}

/** Mod+A: the active cell's current region, or the whole sheet when the
 * selection already is that region (or the whole sheet), or the active cell
 * is isolated. */
function selectRegionOrAll(sel: SheetSelection, nav: NavContext): SheetSelection {
  const active = landing(clampRef(sel.active), nav.merges);
  const region = currentRegion(active, nav);
  const only = sel.ranges.length === 1 ? sel.ranges[0] : undefined;
  if (sameRange(region, unitOf(active, nav.merges)) || (only && (sameRange(only, region) || sameRange(only, WHOLE_SHEET)))) {
    return wholeSheet(nav);
  }
  return { active, ranges: [region], activeRangeIndex: 0 };
}

// --- cycling (Tab / Enter) --------------------------------------------------

/** The next visible cell of `range` after `from` (from its start when null)
 * in row order (along each row, then down) or column order, backwards when
 * `back`. Merge members are skipped — a merge is met at its top-left — and
 * hidden cells too; null when the walk leaves the range. */
function nextInRange(range: GridRange, from: CellRef | null, order: "row" | "column", back: boolean, nav: NavContext): CellRef | null {
  const byRow = order === "row";
  const step: Step = back ? -1 : 1;
  const major = byRow ? rowAxis(nav) : colAxis(nav);
  const minor = byRow ? colAxis(nav) : rowAxis(nav);
  const majorLo = byRow ? range.top : range.left;
  const majorHi = byRow ? range.bottom : range.right;
  const minorLo = byRow ? range.left : range.top;
  const minorHi = byRow ? range.right : range.bottom;
  // No visible line on either axis: no visible cell (merges inside are inside).
  if (firstVisibleBetween(majorLo, majorHi, 1, major) < 0 || firstVisibleBetween(minorLo, minorHi, 1, minor) < 0) {
    return null;
  }
  const at = (a: number, b: number): CellRef => (byRow ? { row: a, col: b } : { row: b, col: a });
  const firstLine = (merge: GridRange): number => (byRow ? merge.top : merge.left);
  const spanLo = (merge: GridRange): number => (byRow ? merge.left : merge.top);
  const spanHi = (merge: GridRange): number => (byRow ? merge.right : merge.bottom);
  const restart = back ? minorHi : minorLo;
  let a = from ? (byRow ? from.row : from.col) : back ? majorHi : majorLo;
  let b = from ? (byRow ? from.col : from.row) + step : restart;
  for (; a >= majorLo && a <= majorHi; a += step, b = restart) {
    if (major.hidden(a)) {
      // Only a visible merge whose top-left sits on this hidden line shows here.
      const lo = back ? minorLo : b;
      const hi = back ? b : minorHi;
      let best: GridRange | null = null;
      if (lo <= hi) {
        for (const merge of mergesIn(nav.merges, byRow ? { top: a, left: lo, bottom: a, right: hi } : { top: lo, left: a, bottom: hi, right: a })) {
          const s = spanLo(merge);
          if (firstLine(merge) === a && s >= lo && s <= hi && (!best || (spanLo(best) - s) * step > 0) && unitVisible(merge, nav)) {
            best = merge;
          }
        }
      }
      if (best) {
        return { row: best.top, col: best.left };
      }
      continue;
    }
    while (b >= minorLo && b <= minorHi) {
      const cell = at(a, b);
      const merge = mergeAt(nav.merges, cell.row, cell.col);
      if (!merge) {
        if (!minor.hidden(b)) {
          return cell;
        }
        b += step;
        continue;
      }
      if (merge.top === cell.row && merge.left === cell.col && unitVisible(merge, nav)) {
        return cell;
      }
      // A member: going back along the merge's first line its top-left comes
      // next; otherwise jump past the merge.
      b = back && firstLine(merge) === a && b !== spanLo(merge) ? spanLo(merge) : back ? spanLo(merge) - 1 : spanHi(merge) + 1;
    }
  }
  return null;
}

/** Tab/Enter in a selection: in a multi-cell selection `active` moves to the
 * next cell of its range in row order (Tab) or column order (Enter), on to
 * the next range at the end (the last wrapping to the first), skipping
 * hidden cells and merge members; `back` reverses (Shift). With one cell (or
 * one merged cell) selected it is moveActive right/down (left/up). */
export function cycleWithin(sel: SheetSelection, order: "row" | "column", back: boolean, nav: NavContext): SheetSelection {
  const active = landing(clampRef(sel.active), nav.merges);
  const ranges = sel.ranges;
  const only = ranges.length === 1 ? ranges[0] : undefined;
  if (ranges.length === 0 || (only && sameRange(only, unitOf(active, nav.merges)))) {
    const dir: Direction = order === "row" ? (back ? "left" : "right") : back ? "up" : "down";
    return moveActive(sel, dir, nav);
  }
  const count = ranges.length;
  let index = (ranges[sel.activeRangeIndex] as GridRange | undefined) ? sel.activeRangeIndex : 0;
  const current = ranges[index]!;
  let next = nextInRange(current, rangeContains(current, active.row, active.col) ? active : null, order, back, nav);
  for (let tries = 0; !next && tries < count; tries += 1) {
    index = (index + (back ? count - 1 : 1)) % count;
    next = nextInRange(ranges[index]!, null, order, back, nav);
  }
  return next ? { active: next, ranges: ranges.map(copyRange), activeRangeIndex: index } : cloneSelection(sel);
}

// --- ranges, rows and columns -----------------------------------------------

/** Mod+press: the pressed cell (its merge; the nearest visible cell when
 * hidden) appended as a new range, which becomes the active one. */
export function addRange(sel: SheetSelection, ref: CellRef, nav: NavContext): SheetSelection {
  const unit = visibleUnitNear(clampRef(ref), nav);
  const ranges = [...sel.ranges.map(copyRange), unit];
  return { active: { row: unit.top, col: unit.left }, ranges, activeRangeIndex: ranges.length - 1 };
}

/** Shift+Space: the active range becomes its whole rows (snapped); the
 * active cell and the other ranges are kept. */
export function selectRows(sel: SheetSelection, nav: NavContext): SheetSelection {
  const range = activeRange(sel, nav);
  return withActiveRange(sel, snapToMerges(wholeRows(range.top, range.bottom), nav.merges));
}

/** Ctrl+Space: the active range becomes its whole columns (snapped); the
 * active cell and the other ranges are kept. */
export function selectColumns(sel: SheetSelection, nav: NavContext): SheetSelection {
  const range = activeRange(sel, nav);
  return withActiveRange(sel, snapToMerges(wholeColumns(range.left, range.right), nav.merges));
}

/** A row/column header press (`from === to`) or drag (`from` the pressed
 * header, `to` the current one): the whole rows or columns between them,
 * snapped, with the first visible cell of the `from` row (column) active.
 * With `base` (Mod+press) the band is appended to base's ranges and becomes
 * the active range; Shift+press is the caller passing the anchor's row or
 * column as `from`. */
export function headerSelection(axis: "row" | "column", from: number, to: number, nav: NavContext, base?: SheetSelection): SheetSelection {
  const byRow = axis === "row";
  const band = snapToMerges(byRow ? wholeRows(from, to) : wholeColumns(from, to), nav.merges);
  const along = byRow ? rowAxis(nav) : colAxis(nav);
  const across = byRow ? colAxis(nav) : rowAxis(nav);
  let line = clampIndex(from, along.last);
  if (along.hidden(line)) {
    const first = firstVisibleBetween(byRow ? band.top : band.left, byRow ? band.bottom : band.right, 1, along);
    line = first >= 0 ? first : line;
  }
  const crossing = Math.max(0, firstVisibleBetween(0, across.last, 1, across));
  const active = landing(byRow ? { row: line, col: crossing } : { row: crossing, col: line }, nav.merges);
  const ranges = base ? [...base.ranges.map(copyRange), band] : [band];
  return { active, ranges, activeRangeIndex: ranges.length - 1 };
}

// --- validation -------------------------------------------------------------

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/** A cell reference read from untrusted data, clamped; null when unreadable. */
function readRef(value: unknown): CellRef | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const { row, col } = value as Record<string, unknown>;
  return isFiniteNumber(row) && isFiniteNumber(col) ? clampRef({ row, col }) : null;
}

/** A range read from untrusted data, normalised and intersected with the
 * grid; null when unreadable or wholly outside the grid. */
function readRange(value: unknown): GridRange | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const { top, left, bottom, right } = value as Record<string, unknown>;
  if (!isFiniteNumber(top) || !isFiniteNumber(left) || !isFiniteNumber(bottom) || !isFiniteNumber(right)) {
    return null;
  }
  const t = Math.trunc(Math.min(top, bottom));
  const b = Math.trunc(Math.max(top, bottom));
  const l = Math.trunc(Math.min(left, right));
  const r = Math.trunc(Math.max(left, right));
  if (b < 0 || t > LAST_ROW || r < 0 || l > LAST_COL) {
    return null;
  }
  return { top: Math.max(0, t), left: Math.max(0, l), bottom: Math.min(LAST_ROW, b), right: Math.min(LAST_COL, r) };
}

/** A selection read from a file or stored view state, made valid for the
 * sheet: every range normalised, clamped and snapped (unreadable ones and
 * ones wholly outside the grid dropped); `active` landed on its merge's
 * top-left and kept when visible inside a range (the requested one first),
 * else the first visible cell of the requested range (else of any range).
 * Undefined, or nothing usable, selects the first visible cell (A1). */
export function normaliseSelection(sel: SheetSelection | undefined, nav: NavContext): SheetSelection {
  const raw = sel as Partial<Record<keyof SheetSelection, unknown>> | undefined;
  const ranges: GridRange[] = [];
  let requestedIndex = -1;
  const source: readonly unknown[] = Array.isArray(raw?.ranges) ? (raw.ranges as unknown[]) : [];
  source.forEach((candidate, i) => {
    const range = readRange(candidate);
    if (!range) {
      return;
    }
    if (i === raw?.activeRangeIndex) {
      requestedIndex = ranges.length;
    }
    ranges.push(snapToMerges(range, nav.merges));
  });
  const requested = readRef(raw?.active);
  if (ranges.length === 0) {
    return requested ? selectCell(requested, nav) : selectCell(firstVisibleCell(nav), nav);
  }
  const order = [...ranges.keys()];
  if (requestedIndex > 0) {
    order.splice(requestedIndex, 1);
    order.unshift(requestedIndex);
  }
  if (requested) {
    const active = landing(requested, nav.merges);
    if (unitVisible(unitOf(active, nav.merges), nav)) {
      const holder = order.find((i) => rangeContains(ranges[i]!, active.row, active.col));
      if (holder !== undefined) {
        return { active, ranges, activeRangeIndex: holder };
      }
    }
  }
  for (const i of order) {
    const first = nextInRange(ranges[i]!, null, "row", false, nav);
    if (first) {
      return { active: first, ranges, activeRangeIndex: i };
    }
  }
  return selectCell(requested ?? firstVisibleCell(nav), nav);
}

// --- the keyboard map -------------------------------------------------------

/** The direction of an arrow key, or null for any other key. */
function arrowDirection(name: string): Direction | null {
  switch (name) {
    case "ArrowUp":
      return "up";
    case "ArrowDown":
      return "down";
    case "ArrowLeft":
      return "left";
    case "ArrowRight":
      return "right";
    default:
      return null;
  }
}

/** The keyboard map (SPEC-117 "Keyboard map") applied to `sel`; null for a
 * key it does not handle — the caller then leaves the event alone (no
 * preventDefault). Ctrl+PageUp/PageDown (sheet switching) belong to the
 * view and are null here. Beyond the table, Shift extends Home, Mod+Home and
 * Mod+End as it does the arrows. */
export function navigateByKey(sel: SheetSelection, key: NavKey, nav: NavContext): SheetSelection | null {
  const { shift, mod, alt, ctrl } = key;
  const arrow = arrowDirection(key.key);
  if (arrow) {
    if (alt || (ctrl && !mod)) {
      return null;
    }
    if (mod) {
      return shift ? extendToEdge(sel, arrow, nav) : selectCell(edgeInDirection(sel.active, arrow, nav), nav);
    }
    return shift ? extendBy(sel, arrow, nav) : moveActive(sel, arrow, nav);
  }
  switch (key.key) {
    case "PageDown":
    case "PageUp": {
      if (ctrl || mod) {
        return null;
      }
      const forward = key.key === "PageDown";
      const dir: Direction = alt ? (forward ? "right" : "left") : forward ? "down" : "up";
      const steps = alt ? nav.pageCols : nav.pageRows;
      return shift ? extendBy(sel, dir, nav, steps) : moveActive(sel, dir, nav, steps);
    }
    case "Home": {
      if (alt || (ctrl && !mod)) {
        return null;
      }
      const firstCol = Math.max(0, firstVisibleBetween(0, LAST_COL, 1, colAxis(nav)));
      if (mod) {
        const target = firstVisibleCell(nav);
        return shift ? extendTo(sel, target, nav) : selectCell(target, nav);
      }
      if (shift) {
        return extendTo(sel, { row: movingCorner(sel, nav).row, col: firstCol }, nav);
      }
      const active = clampRef(sel.active);
      const unit = unitOf(active, nav.merges);
      return selectCell({ row: lineIndex(active.row, unit.top, unit.bottom, rowAxis(nav)), col: firstCol }, nav);
    }
    case "End": {
      if (!mod || alt) {
        return null;
      }
      const target = lastUsedCell(nav);
      return shift ? extendTo(sel, target, nav) : selectCell(target, nav);
    }
    case "a":
    case "A":
      return mod && !shift && !alt ? selectRegionOrAll(sel, nav) : null;
    case " ":
    case "Spacebar":
      if (alt) {
        return null;
      }
      if (ctrl && !shift) {
        return selectColumns(sel, nav);
      }
      return shift && !ctrl && !mod ? selectRows(sel, nav) : null;
    case "Tab":
      return mod || ctrl || alt ? null : cycleWithin(sel, "row", shift, nav);
    case "Enter":
      return mod || ctrl || alt ? null : cycleWithin(sel, "column", shift, nav);
    default:
      return null;
  }
}
