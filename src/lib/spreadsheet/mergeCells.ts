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
// 2. Merges inside the expanded range are removed.
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

/** The merges a command makes on `sheet` (steps 1 and 4): expanded, split
 * per row for "across", single cells dropped. */
export function mergeTargets(sheet: Sheet, ranges: readonly GridRange[], mode: MergeMode): GridRange[] {
  const targets: GridRange[] = [];
  for (const range of ranges) {
    const expanded = expandToMerges(range, sheet.merges);
    const pieces =
      mode === "across" ? Array.from({ length: expanded.bottom - expanded.top + 1 }, (_, k) => ({ ...expanded, top: expanded.top + k, bottom: expanded.top + k })) : [expanded];
    for (const piece of pieces) {
      if (!isSingleCell(piece) && !targets.some((done) => rangesIntersect(done, piece))) {
        targets.push(piece);
      }
    }
  }
  return targets;
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
  return sheet ? targetsDiscardValues(sheet, mergeTargets(sheet, ranges, mode)) : false;
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

/** Merges `targets` on the sheet at `index` (steps 2 and 3), no refusals —
 * the caller checked them. */
export function placeMerges(model: SpreadsheetModel, index: number, targets: readonly GridRange[]): void {
  const sheet = model.sheets[index]!;
  if (targets.length === 0) {
    return;
  }
  const inTarget = (rowIndex: number, col: number): GridRange | undefined =>
    targets.find((t) => rowIndex >= t.top && rowIndex <= t.bottom && col >= t.left && col <= t.right);
  const top = Math.min(...targets.map((t) => t.top));
  const bottom = Math.max(...targets.map((t) => t.bottom));
  const rows = sheet.rows.flatMap((row): Row[] => {
    if (row.index < top || row.index > bottom || !row.cells.some(holdsContent)) {
      return [row];
    }
    let changed = false;
    const cells = row.cells.flatMap((cell): Cell[] => {
      const t = inTarget(row.index, cell.col);
      if (!t || !holdsContent(cell) || (row.index === t.top && cell.col === t.left)) {
        return [cell];
      }
      changed = true;
      const cleared = clearedCell(cell);
      return cleared ? [cleared] : [];
    });
    if (!changed) {
      return [row];
    }
    const next = { ...row, cells };
    return isEmptyRow(next) ? [] : [next];
  });
  const merges = [...sheet.merges.filter((merge) => !targets.some((t) => contains(t, merge))), ...targets];
  model.sheets = model.sheets.map((each, i) => (i === index ? { ...sheet, rows, merges, dirty: true as const } : each));
}

/** Why merging `targets` on `sheet` is refused, or null (step 5). */
export function mergeRefusal(model: SpreadsheetModel, index: number, targets: readonly GridRange[]): string | null {
  const sheet = model.sheets[index]!;
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      const array = cell.arrayRange;
      if (array && targets.some((t) => rangesIntersect(t, array) && !contains(t, array))) {
        return PARTIAL_ARRAY_MESSAGE;
      }
    }
  }
  return firstLockedCell(model, index, [...targets]) ? PROTECTED_CELL_MESSAGE : null;
}

/** Merges `ranges` (see the module header). */
export function mergeCells(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], mode: MergeMode): FormatResult {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  const targets = mergeTargets(sheet, ranges, mode);
  if (targets.length === 0) {
    return { ok: true };
  }
  const refusal = mergeRefusal(model, index, targets);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  placeMerges(model, index, targets);
  if (mode === "centre") {
    return applyStylePatch(
      model,
      sheetId,
      targets.map((t) => ({ top: t.top, left: t.left, bottom: t.top, right: t.left })),
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
 * it are removed and the source's placed, their non-anchor values cleared
 * (the view warns first). Refused as mergeCells refuses. */
export function applyPatternMerges(model: SpreadsheetModel, sheetId: string, target: GridRange, pattern: StylePattern): FormatResult {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || pattern.axis !== undefined) {
    return { ok: true };
  }
  const range = patternTarget(target, pattern);
  const targets = patternMerges(target, pattern);
  const kept = sheet.merges.filter((merge) => !contains(range, merge));
  if (targets.length === 0 && kept.length === sheet.merges.length) {
    return { ok: true };
  }
  const refusal = mergeRefusal(model, index, targets);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  model.sheets = model.sheets.map((each, i) => (i === index ? { ...sheet, merges: kept, dirty: true as const } : each));
  placeMerges(model, index, targets);
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
