// Cut and move (SPEC-125 "Cut and move") — THIS FILE is the normative source
// for RangeMove, repointFormula, moveMapper, the MOVE_ADJUSTERS registry and
// moveRange: what pasting a cut (and a drag-move) does to the workbook.
//
// | A reference anywhere in the workbook, moved formulas included | becomes |
// | entirely inside the source (cell or range) | moved by the offset, onto the destination sheet (prefix added, changed or dropped to match the host) |
// | range partly overlapping the source | unchanged |
// | entirely inside the destination and not in the source | `#REF!` (its cells were overwritten) |
// | whole-row/column, 3-D, external, anything else | unchanged |
// Moved formulas are otherwise NOT shifted (Excel): a moved formula keeps
// pointing at the cells it pointed at — across sheets its references to its
// old sheet gain that sheet's prefix.
//
// moveRange checks the refusals before anything changes — a sheet that is
// not a worksheet, a destination past the grid's edge, more than
// MAX_COPY_CELLS positions, a merge or an array range crossing the source's
// or the destination's edge (one entirely inside the source moves with it),
// a locked cell of a protected sheet in either — then runs every adjuster in
// registry order and marks both sheets dirty. Source and destination may
// overlap (moving a block down one row): the move reads the source before it
// writes. Each adjuster replaces what it changes copy-on-write and keeps the
// identity of everything else.
//
// | Order | id | Moves |
// | 1 | cells | values, formulas, styles (each position's effective style), runs, preserved content and array anchors; the source left as unstyled blanks, the destination overwritten |
// | 2 | formulas | repointFormula over every sheet's formulas (not formulaUnparsed) |
// | 3 | merges | source merges move, destination merges are replaced |
// | 4 | retained | SPEC-124's rewriteRetainedReferences with moveMapper |
// Later specs that register a SPEC-124 ReferenceAdjuster for something that
// travels with cells or points at them — the autofilter (128), conditional
// formats (129), validations (130), chart references (132), defined names
// and tables (133), notes and hyperlinks (135), pivots (139) — register a
// MoveAdjuster with the same id here, with its row in move.spec.ts's
// registry contract; drawings (131) and page breaks (136) deliberately do
// not.
import type { CellRef, GridRange } from "./address";
import { forEachPosition, PASTE_MERGE_MESSAGE, PASTE_PAST_EDGE_MESSAGE, TOO_MANY_TO_CHANGE_MESSAGE } from "./clipboard";
import { firstLockedCell, PROTECTED_CELL_MESSAGE } from "./cells";
import { MAX_COPY_CELLS } from "./clipboardPayload";
import { rewriteReferences, type RefPart, type ReferenceNode } from "./formula/references";
import { rangesIntersect } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type Sheet, type SpreadsheetModel } from "./model";
import { rewriteRetainedReferences, type RetainedReferenceMapper } from "./retainedReferences";
import { contains, inheritedStyleId, SheetEdit } from "./sheetEdit";
import { resolveRetainedSheetName } from "./sourceNames";
import { NOT_A_GRID_MESSAGE, PARTIAL_ARRAY_MESSAGE } from "./structure";
import { LazyParts } from "./xlsx/package";

export interface RangeMove {
  fromSheetId: string;
  from: GridRange;
  toSheetId: string;
  toTop: number;
  toLeft: number;
}

export interface MoveAdjuster {
  id: string;
  move(model: SpreadsheetModel, move: RangeMove): void;
}

export type MoveResult = { ok: true } | { ok: false; message: string };

/** The rectangle a move writes. */
export function moveDestination(move: RangeMove): GridRange {
  return {
    top: move.toTop,
    left: move.toLeft,
    bottom: move.toTop + move.from.bottom - move.from.top,
    right: move.toLeft + move.from.right - move.from.left,
  };
}

const offsetOf = (move: RangeMove): { dRow: number; dCol: number } => ({ dRow: move.toTop - move.from.top, dCol: move.toLeft - move.from.left });

const sheetOf = (model: SpreadsheetModel, id: string): Sheet | undefined => model.sheets.find((sheet) => sheet.id === id);

function shifted(range: GridRange, dRow: number, dCol: number): GridRange {
  return { top: range.top + dRow, left: range.left + dCol, bottom: range.bottom + dRow, right: range.right + dCol };
}

/** Whether any array formula's range on `sheet` crosses `range`'s edge,
 * ignoring those entirely inside `moving` (which travel with the block). */
function arrayCrosses(sheet: Sheet, range: GridRange, moving?: GridRange): boolean {
  return sheet.rows.some((row) =>
    row.cells.some((cell) => cell.arrayRange !== undefined && rangesIntersect(cell.arrayRange, range) && !contains(range, cell.arrayRange) && !(moving && contains(moving, cell.arrayRange))),
  );
}

/** Why `move` cannot be made, or null (see the module header). */
export function moveRefusal(model: SpreadsheetModel, move: RangeMove): string | null {
  const from = sheetOf(model, move.fromSheetId);
  const to = sheetOf(model, move.toSheetId);
  if (!from || !to || from.kind !== "grid" || to.kind !== "grid") {
    return NOT_A_GRID_MESSAGE;
  }
  const dest = moveDestination(move);
  if (dest.top < 0 || dest.left < 0 || dest.bottom >= MAX_ROWS || dest.right >= MAX_COLS) {
    return PASTE_PAST_EDGE_MESSAGE;
  }
  if ((dest.bottom - dest.top + 1) * (dest.right - dest.left + 1) > MAX_COPY_CELLS) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  const same = from === to;
  if (
    from.merges.some((merge) => rangesIntersect(merge, move.from) && !contains(move.from, merge)) ||
    to.merges.some((merge) => rangesIntersect(merge, dest) && !contains(dest, merge) && !(same && contains(move.from, merge)))
  ) {
    return PASTE_MERGE_MESSAGE;
  }
  if (arrayCrosses(from, move.from) || arrayCrosses(to, dest, same ? move.from : undefined)) {
    return PARTIAL_ARRAY_MESSAGE;
  }
  const fromIndex = model.sheets.indexOf(from);
  const toIndex = model.sheets.indexOf(to);
  return firstLockedCell(model, fromIndex, [move.from]) || firstLockedCell(model, toIndex, [dest]) ? PROTECTED_CELL_MESSAGE : null;
}

// --- 1. cells ----------------------------------------------------------------------------------

const cellsAdjuster: MoveAdjuster = {
  id: "cells",
  move(model, move) {
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    const { dRow, dCol } = offsetOf(move);
    if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
      return;
    }
    // Read the whole source first: source and destination may overlap.
    const moved = new Map<number, Cell>();
    forEachPosition(from, move.from, (row, col, cell, effective) => {
      if (!cell && effective === 0) {
        return;
      }
      const next: Cell = cell ? { ...cell, col: col + dCol } : { col: col + dCol, value: null };
      if (next.styleId === undefined && effective !== 0) {
        next.styleId = effective;
      }
      if (next.arrayRange) {
        next.arrayRange = shifted(next.arrayRange, dRow, dCol);
      }
      moved.set((row + dRow) * MAX_COLS + col + dCol, next);
    });
    const fromEdit = new SheetEdit(from);
    const toEdit = from === to ? fromEdit : new SheetEdit(to);
    // The source becomes unstyled blanks.
    forEachPosition(from, move.from, (row, col, cell) => {
      if (inheritedStyleId(from, row, col) !== 0) {
        fromEdit.set(row, col, { col, value: null, styleId: 0 });
      } else if (cell) {
        fromEdit.set(row, col, undefined);
      }
    });
    // The destination is overwritten.
    const dest = moveDestination(move);
    for (let row = dest.top; row <= dest.bottom; row += 1) {
      for (let col = dest.left; col <= dest.right; col += 1) {
        const cell = moved.get(row * MAX_COLS + col);
        const inherited = inheritedStyleId(to, row, col);
        if (cell) {
          toEdit.set(row, col, cell.styleId === undefined && inherited !== 0 ? { ...cell, styleId: 0 } : cell);
        } else {
          toEdit.set(row, col, inherited !== 0 ? { col, value: null, styleId: 0 } : undefined);
        }
      }
    }
    const fromRows = fromEdit.rows();
    const toRows = toEdit === fromEdit ? fromRows : toEdit.rows();
    if (fromRows === from.rows && toRows === to.rows) {
      return;
    }
    model.sheets = model.sheets.map((sheet) => {
      if (sheet === from && fromRows !== from.rows) {
        return { ...sheet, rows: fromRows };
      }
      if (sheet === to && to !== from && toRows !== to.rows) {
        return { ...sheet, rows: toRows };
      }
      return sheet;
    });
  },
};

// --- 2. formulas -------------------------------------------------------------------------------

/** A reference's area (whole rows and columns are null: never moved). */
function areaOf(ref: ReferenceNode): GridRange | null {
  const end = ref.end ?? ref.start;
  if (ref.start.row === undefined || ref.start.col === undefined || end.row === undefined || end.col === undefined) {
    return null;
  }
  return {
    top: Math.min(ref.start.row, end.row),
    left: Math.min(ref.start.col, end.col),
    bottom: Math.max(ref.start.row, end.row),
    right: Math.max(ref.start.col, end.col),
  };
}

function movedPart(part: RefPart, dRow: number, dCol: number): RefPart {
  return { ...part, row: part.row! + dRow, col: part.col! + dCol };
}

/** `text` (canonical, no "=") hosted on `hostSheetId` with every reference
 * re-pointed for `move` (see the module header). `movedTo`: the formula
 * itself moved there with the cells (its host was `hostSheetId`). */
export function repointFormula(text: string, hostSheetId: string, move: RangeMove, sheetNameOf: (id: string) => string, movedTo?: string): string {
  const fromName = sheetNameOf(move.fromSheetId).toUpperCase();
  const toName = sheetNameOf(move.toSheetId).toUpperCase();
  const hostName = sheetNameOf(hostSheetId);
  const newHost = movedTo ?? hostSheetId;
  const crossSheet = move.fromSheetId !== move.toSheetId;
  const { dRow, dCol } = offsetOf(move);
  const dest = moveDestination(move);
  return rewriteReferences(text, (ref) => {
    const area = ref.sheetEnd === undefined ? areaOf(ref) : null;
    if (!area) {
      return undefined;
    }
    const on = (ref.sheet ?? hostName).toUpperCase();
    if (on === fromName && contains(move.from, area)) {
      const moved: ReferenceNode = { start: movedPart(ref.start, dRow, dCol) };
      if (ref.end) {
        moved.end = movedPart(ref.end, dRow, dCol);
      }
      if (!crossSheet) {
        if (ref.sheet !== undefined) {
          moved.sheet = ref.sheet;
        }
      } else if (move.toSheetId !== newHost) {
        moved.sheet = sheetNameOf(move.toSheetId);
      }
      return moved;
    }
    if (on === toName && contains(dest, area)) {
      return null;
    }
    if (movedTo !== undefined && movedTo !== hostSheetId && ref.sheet === undefined) {
      // A moved formula keeps pointing at its old sheet's cells.
      return { ...ref, sheet: hostName };
    }
    return undefined;
  });
}

const formulasAdjuster: MoveAdjuster = {
  id: "formulas",
  move(model, move) {
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    const dest = moveDestination(move);
    let sheetsChanged = false;
    const sheets = model.sheets.map((sheet) => {
      const involved = sheet.id === move.fromSheetId || sheet.id === move.toSheetId;
      let rowsChanged = false;
      const rows = sheet.rows.map((row) => {
        let cellsChanged = false;
        const cells = row.cells.map((cell) => {
          // Another sheet's formula reaches the moved cells only with a `!`.
          if (cell.formula === undefined || cell.formulaUnparsed || (!involved && !cell.formula.includes("!"))) {
            return cell;
          }
          const moved = sheet.id === move.toSheetId && row.index >= dest.top && row.index <= dest.bottom && cell.col >= dest.left && cell.col <= dest.right;
          const formula = moved ? repointFormula(cell.formula, move.fromSheetId, move, nameOf, move.toSheetId) : repointFormula(cell.formula, sheet.id, move, nameOf);
          if (formula === cell.formula) {
            return cell;
          }
          cellsChanged = true;
          return { ...cell, formula };
        });
        if (!cellsChanged) {
          return row;
        }
        rowsChanged = true;
        return { ...row, cells };
      });
      if (!rowsChanged) {
        return sheet;
      }
      sheetsChanged = true;
      return { ...sheet, rows, dirty: true as const };
    });
    if (sheetsChanged) {
      model.sheets = sheets;
    }
  },
};

// --- 3. merges ---------------------------------------------------------------------------------

const mergesAdjuster: MoveAdjuster = {
  id: "merges",
  move(model, move) {
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    const { dRow, dCol } = offsetOf(move);
    if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
      return;
    }
    const dest = moveDestination(move);
    const moving = from.merges.filter((merge) => contains(move.from, merge));
    const replaced = to.merges.filter((merge) => contains(dest, merge) && !(from === to && contains(move.from, merge)));
    if (moving.length === 0 && replaced.length === 0) {
      return;
    }
    const arrived = moving.map((merge) => shifted(merge, dRow, dCol));
    model.sheets = model.sheets.map((sheet) => {
      if (sheet === from && from === to) {
        return { ...sheet, merges: [...sheet.merges.filter((merge) => !moving.includes(merge) && !replaced.includes(merge)), ...arrived] };
      }
      if (sheet === from) {
        return moving.length > 0 ? { ...sheet, merges: sheet.merges.filter((merge) => !moving.includes(merge)) } : sheet;
      }
      if (sheet === to) {
        return { ...sheet, merges: [...sheet.merges.filter((merge) => !replaced.includes(merge)), ...arrived] };
      }
      return sheet;
    });
  },
};

// --- 4. retained -------------------------------------------------------------------------------

/** The mapper a move re-points retained text with (see
 * retainedReferences.ts): on the source sheet, a range or reference entirely
 * inside the source moves with the cells — on the same sheet only, since no
 * sheet name in retained text is ever rewritten; on the destination sheet,
 * one entirely inside the destination is dropped (a range) or becomes
 * `#REF!` (a reference); partial overlaps, whole rows and columns and
 * indices (drawings, page breaks) stay. Retained formulas name sheets as the
 * source file did: each resolves through SPEC-121's
 * resolveRetainedSheetName. */
export function moveMapper(model: SpreadsheetModel, move: RangeMove): RetainedReferenceMapper {
  const fromName = sheetOf(model, move.fromSheetId)?.name.toUpperCase();
  const toName = sheetOf(model, move.toSheetId)?.name.toUpperCase();
  const same = move.fromSheetId === move.toSheetId;
  const { dRow, dCol } = offsetOf(move);
  const dest = moveDestination(move);
  const parts = new LazyParts(model.opaque);
  return {
    range: (sheetName, range) => {
      const name = sheetName.toUpperCase();
      if (name === fromName && contains(move.from, range)) {
        return same ? shifted(range, dRow, dCol) : range;
      }
      return name === toName && contains(dest, range) ? null : range;
    },
    index: (_sheetName, _axis, index) => index,
    formula: (canonical, hostSheetId) =>
      rewriteReferences(canonical, (ref) => {
        const area = ref.sheetEnd === undefined ? areaOf(ref) : null;
        const sheetId = ref.sheet === undefined ? hostSheetId || undefined : resolveRetainedSheetName(model, parts, ref.sheet, hostSheetId || undefined)?.id;
        if (!area || sheetId === undefined) {
          return undefined;
        }
        if (sheetId === move.fromSheetId && contains(move.from, area)) {
          return same ? { ...ref, start: movedPart(ref.start, dRow, dCol), ...(ref.end ? { end: movedPart(ref.end, dRow, dCol) } : {}) } : undefined;
        }
        return sheetId === move.toSheetId && contains(dest, area) ? null : undefined;
      }),
  };
}

const retainedAdjuster: MoveAdjuster = {
  id: "retained",
  move(model, move) {
    if (model.opaque.length === 0) {
      return;
    }
    rewriteRetainedReferences(model, moveMapper(model, move));
  },
};

// --- the registry ------------------------------------------------------------------------------

export const MOVE_ADJUSTERS: MoveAdjuster[] = [cellsAdjuster, formulasAdjuster, mergesAdjuster, retainedAdjuster];

/** One move (see the module header). */
export function moveRange(model: SpreadsheetModel, move: RangeMove): MoveResult {
  const refusal = moveRefusal(model, move);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  const { dRow, dCol } = offsetOf(move);
  if (move.fromSheetId === move.toSheetId && dRow === 0 && dCol === 0) {
    return { ok: true };
  }
  for (const adjuster of MOVE_ADJUSTERS) {
    adjuster.move(model, move);
  }
  model.sheets = model.sheets.map((sheet) => ((sheet.id === move.fromSheetId || sheet.id === move.toSheetId) && !sheet.dirty ? { ...sheet, dirty: true as const } : sheet));
  return { ok: true };
}

/** The top-left a block lands on when moved from `anchor` to `at`, kept
 * inside the grid (a drag-move's snapped destination). */
export function snappedMove(from: GridRange, anchor: CellRef, at: CellRef): { toTop: number; toLeft: number } {
  const rows = from.bottom - from.top;
  const cols = from.right - from.left;
  return {
    toTop: Math.min(Math.max(0, from.top + at.row - anchor.row), MAX_ROWS - 1 - rows),
    toLeft: Math.min(Math.max(0, from.left + at.col - anchor.col), MAX_COLS - 1 - cols),
  };
}
