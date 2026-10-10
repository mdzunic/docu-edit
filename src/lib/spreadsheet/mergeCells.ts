// Merging and unmerging cells (SPEC-124 "Merges") — the commands behind the
// toolbar's Merge split button, the format painter's merges and Clear
// formats' unmerge. Copy-on-write like every command, the sheet marked
// dirty, a FormatResult back; the view wraps each in ONE undo step ("Merge
// cells", "Unmerge cells").
//
// mergeCells(model, sheetId, ranges, mode):
// 1. Each range is expanded until it contains every merge it intersects (a
//    range partially overlapping a merge EXTENDS to cover it — never
//    refused).
// 2. Merges inside the expanded range are removed — before "across" splits
//    it into rows, so a merge spanning several of them goes too.
// 3. Kept value: the upper-left cell (per row for "across"); every other
//    cell's value and formula are cleared, its style kept (Excel draws the
//    merge's borders from them). mergeDiscardsValues tells the view whether
//    any would be lost, so it can warn first.
// 4. "centre" adds `{ alignment: { horizontal: "center" } }` to each merge's
//    upper-left cell through SPEC-123's applyStylePatch; "across" merges each
//    row of a multi-column range separately; a single cell is a no-op.
// 5. A merge partly over an array formula's range is refused with
//    PARTIAL_ARRAY_MESSAGE; a locked cell of a protected sheet, with
//    SPEC-122's protected message.
// The ranges are taken in turn, as though merged one after another: a range
// meeting an earlier one's merge extends to cover it like any other. So no
// two merges ever overlap — Excel repairs a file whose merges do.
// unmergeCells removes every merge intersecting the ranges; content and
// styles stay.
import type { GridRange } from "./address";
import { firstLockedCell, isEmptyRow, PROTECTED_CELL_MESSAGE } from "./cells";
import { applyStylePatch, patternTarget, type FormatResult, type StylePattern } from "./formatting";
import { rangesIntersect } from "./merges";
import { dropValueBound, type Cell, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { firstRowAtOrAfter } from "./restyle";
import { holdsContent, NOT_A_GRID_MESSAGE, PARTIAL_ARRAY_MESSAGE } from "./structure";

export type MergeMode = "centre" | "across" | "cells";

const contains = (outer: GridRange, inner: GridRange): boolean =>
  outer.top <= inner.top && outer.left <= inner.left && outer.bottom >= inner.bottom && outer.right >= inner.right;

const isSingleCell = (range: GridRange): boolean => range.top === range.bottom && range.left === range.right;

/** `range` grown until it contains every merge it intersects. */
export function expandToMerges(range: GridRange, merges: readonly GridRange[]): GridRange {
  let current = range;
  let grown = true;
  while (grown) {
    grown = false;
    for (const merge of merges) {
      if (rangesIntersect(current, merge) && !contains(current, merge)) {
        current = {
          top: Math.min(current.top, merge.top),
          left: Math.min(current.left, merge.left),
          bottom: Math.max(current.bottom, merge.bottom),
          right: Math.max(current.right, merge.right),
        };
        grown = true;
      }
    }
  }
  return current;
}

/** What a merge command does on a sheet. */
export interface MergePlan {
  /** The merges it makes (steps 1 and 4): each range expanded, split per row
   * for "across", single cells dropped. */
  targets: GridRange[];
  /** The sheet's merges afterwards: those inside an expanded range removed
   * (step 2), the targets added. */
  merges: GridRange[];
}

/** mergeCells' plan on `sheet` (see the module header). */
export function planMerges(sheet: Sheet, ranges: readonly GridRange[], mode: MergeMode): MergePlan {
  let merges = sheet.merges;
  let targets: GridRange[] = [];
  for (const range of ranges) {
    const expanded = expandToMerges(range, merges);
    const pieces =
      mode === "across" ? Array.from({ length: expanded.bottom - expanded.top + 1 }, (_, k) => ({ ...expanded, top: expanded.top + k, bottom: expanded.top + k })) : [{ ...expanded }];
    const made = pieces.filter((piece) => !isSingleCell(piece));
    const outside = (merge: GridRange): boolean => !contains(expanded, merge);
    merges = [...merges.filter(outside), ...made];
    targets = [...targets.filter(outside), ...made];
  }
  return { targets, merges };
}

/** Calls `visit` for every stored cell in `range`. */
function forEachCell(sheet: Sheet, range: GridRange, visit: (row: Row, cell: Cell) => void): void {
  for (let r = firstRowAtOrAfter(sheet.rows, range.top); r < sheet.rows.length && sheet.rows[r]!.index <= range.bottom; r += 1) {
    const row = sheet.rows[r]!;
    for (const cell of row.cells) {
      if (cell.col >= range.left && cell.col <= range.right) {
        visit(row, cell);
      }
    }
  }
}

/** Whether merging `targets` would clear a value or formula — the warning
 * "Merging keeps only the upper-left value". */
export function targetsDiscardValues(sheet: Sheet, targets: readonly GridRange[]): boolean {
  return targets.some((target) => {
    let found = false;
    forEachCell(sheet, target, (row, cell) => {
      found ||= holdsContent(cell) && !(row.index === target.top && cell.col === target.left);
    });
    return found;
  });
}

/** Whether mergeCells(model, sheetId, ranges, mode) would clear a value or
 * formula (see the module header). */
export function mergeDiscardsValues(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], mode: MergeMode): boolean {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  return sheet ? targetsDiscardValues(sheet, planMerges(sheet, ranges, mode).targets) : false;
}

/** A cell with its value and formula cleared (style kept); undefined when
 * nothing would be left. */
function clearedCell(cell: Cell): Cell | undefined {
  const next: Cell = { ...cell, value: null };
  dropValueBound(next);
  delete next.formula;
  delete next.formulaUnparsed;
  delete next.arrayRange;
  delete next.dynamicArray;
  delete next.cm;
  delete next.dataTableXml;
  const record = next as unknown as Record<string, unknown>;
  const holds = Object.keys(record).some((key) => key !== "col" && key !== "value" && record[key] !== undefined);
  return holds ? next : undefined;
}

/** `row` with the values and formulas of its cells in `target` cleared, the
 * upper-left one's kept (step 3); null when nothing is left of it. */
function clearedRow(row: Row, target: GridRange): Row | null {
  let changed = false;
  const cells = row.cells.flatMap((cell): Cell[] => {
    if (cell.col < target.left || cell.col > target.right || !holdsContent(cell) || (row.index === target.top && cell.col === target.left)) {
      return [cell];
    }
    changed = true;
    const cleared = clearedCell(cell);
    return cleared ? [cleared] : [];
  });
  if (!changed) {
    return row;
  }
  const next = { ...row, cells };
  return isEmptyRow(next) ? null : next;
}

/** Carries out `plan` on the sheet at `index` (steps 2 and 3), no refusals —
 * the caller checked them. */
export function placeMerges(model: SpreadsheetModel, index: number, plan: MergePlan): void {
  const sheet = model.sheets[index]!;
  // sheet.rows by position, each replaced as a target clears its cells.
  const rows: (Row | null)[] = [...sheet.rows];
  let cleared = false;
  for (const target of plan.targets) {
    for (let r = firstRowAtOrAfter(sheet.rows, target.top); r < rows.length && sheet.rows[r]!.index <= target.bottom; r += 1) {
      const row = rows[r] ?? null;
      const next = row && clearedRow(row, target);
      if (next !== row) {
        rows[r] = next;
        cleared = true;
      }
    }
  }
  const kept = cleared ? rows.filter((row): row is Row => row !== null) : sheet.rows;
  model.sheets = model.sheets.map((each, i) => (i === index ? { ...sheet, rows: kept, merges: plan.merges, dirty: true as const } : each));
}

/** Why `plan` is refused on the sheet at `index`, or null (step 5): a merge
 * partly over an array formula's range, or — on a protected sheet — a locked
 * cell among those it merges or unmerges. */
export function mergeRefusal(model: SpreadsheetModel, index: number, plan: MergePlan): string | null {
  const sheet = model.sheets[index]!;
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      const array = cell.arrayRange;
      if (array && plan.targets.some((t) => rangesIntersect(t, array) && !contains(t, array))) {
        return PARTIAL_ARRAY_MESSAGE;
      }
    }
  }
  if (!sheet.protection) {
    return null;
  }
  const kept = new Set(plan.merges);
  const unmerged = sheet.merges.filter((merge) => !kept.has(merge));
  return firstLockedCell(model, index, [...plan.targets, ...unmerged]) ? PROTECTED_CELL_MESSAGE : null;
}

/** Merges `ranges` (see the module header). */
export function mergeCells(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], mode: MergeMode): FormatResult {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  const plan = planMerges(sheet, ranges, mode);
  if (plan.targets.length === 0 && plan.merges.length === sheet.merges.length) {
    return { ok: true };
  }
  const refusal = mergeRefusal(model, index, plan);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  placeMerges(model, index, plan);
  if (mode === "centre") {
    return applyStylePatch(
      model,
      sheetId,
      plan.targets.map((t) => ({ top: t.top, left: t.left, bottom: t.top, right: t.left })),
      { alignment: { horizontal: "center" } },
    );
  }
  return { ok: true };
}

/** The source's merges tiled over a pattern's target — every copy that lies
 * wholly inside it (the format painter, SPEC-124). */
export function patternMerges(target: GridRange, pattern: StylePattern): GridRange[] {
  const range = patternTarget(target, pattern);
  const out: GridRange[] = [];
  if (!pattern.merges || pattern.axis !== undefined) {
    return out;
  }
  for (let top = range.top; top <= range.bottom; top += pattern.rows) {
    for (let left = range.left; left <= range.right; left += pattern.cols) {
      for (const m of pattern.merges) {
        const placed = { top: top + m.top, left: left + m.left, bottom: top + m.bottom, right: left + m.right };
        if (contains(range, placed)) {
          out.push(placed);
        }
      }
    }
  }
  return out;
}

/** The painter's merges (see patternMerges): the target's own merges inside
 * it are removed — as is one reaching into it that a placed merge would
 * overlap — and the source's placed, their non-anchor values cleared (the
 * view warns first). Refused as mergeCells refuses. */
export function applyPatternMerges(model: SpreadsheetModel, sheetId: string, target: GridRange, pattern: StylePattern): FormatResult {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || pattern.axis !== undefined) {
    return { ok: true };
  }
  const range = patternTarget(target, pattern);
  const targets = patternMerges(target, pattern);
  const kept = sheet.merges.filter((merge) => !contains(range, merge) && !targets.some((t) => rangesIntersect(t, merge)));
  if (targets.length === 0 && kept.length === sheet.merges.length) {
    return { ok: true };
  }
  const plan = { targets, merges: [...kept, ...targets] };
  const refusal = mergeRefusal(model, index, plan);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  placeMerges(model, index, plan);
  return { ok: true };
}

/** Removes every merge intersecting `ranges`; content and styles stay. */
export function unmergeCells(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[]): FormatResult {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  if (sheet.protection && !sheet.protection.allowFormatCells) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  const merges = sheet.merges.filter((merge) => !ranges.some((range) => rangesIntersect(range, merge)));
  if (merges.length !== sheet.merges.length) {
    model.sheets = model.sheets.map((each, i) => (i === index ? { ...sheet, merges, dirty: true as const } : each));
  }
  return { ok: true };
}
