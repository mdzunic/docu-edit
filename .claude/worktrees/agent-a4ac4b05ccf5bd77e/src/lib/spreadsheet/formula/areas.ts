// Reading references and arrays (SPEC-120 "Ranges and areas", "Implicit
// intersection", "Array mode") — shared by the evaluator and the functions,
// so every reader of a range walks it the same way:
//
// - Iteration visits STORED cells only: a binary search to the area's first
//   row, then forward, and within each row a binary search to its first
//   column — `=SUM(A:A)` over 1,000 values reads 1,000 cells, never
//   1,048,576. Every value is read through `ctx.valueAt`, so a precedent that
//   is dirty in the running pass is evaluated first. A function that counts
//   empties computes them from the area's size.
// - Materialising a range (array mode, "array" arguments) is bounded by
//   MAX_ARRAY_CELLS (#NUM! beyond it) and needs a single area (#VALUE!).
// - Broadcasting: a 1×1 array or a scalar fills any shape; a single row
//   repeats down and a single column across; otherwise the larger extent
//   wins in each dimension and a position outside an operand is #N/A.
// - Implicit intersection: a single cell is its value; one column (row)
//   yields the cell in the host's row (column), else #VALUE!; a 2-D range or
//   several areas are #VALUE!; an array is its top-left element.
import type { CellRef } from "../address";
import { MAX_COLS, MAX_ROWS, type Cell, type CellError, type Row, type Sheet, type SpreadsheetModel } from "../model";
import type { EvalContext } from "./evaluate";
import type { RefPart, ReferenceNode } from "./references";
import {
  EMPTY,
  isArray,
  isError,
  isRange,
  MAX_ARRAY_CELLS,
  NA_ERROR,
  NUM_ERROR,
  REF_ERROR,
  VALUE_ERROR,
  type Area,
  type ArrayValue,
  type EvalValue,
  type RangeValue,
  type ScalarValue,
} from "./values";

export interface Shape {
  rows: number;
  cols: number;
}

// --- sheets ------------------------------------------------------------------------

const sheetIndexes = new WeakMap<readonly Sheet[], { byId: Map<string, number>; byName: Map<string, number> }>();

function indexesOf(model: SpreadsheetModel): { byId: Map<string, number>; byName: Map<string, number> } {
  let built = sheetIndexes.get(model.sheets);
  if (!built) {
    built = { byId: new Map(), byName: new Map() };
    model.sheets.forEach((sheet, index) => {
      built!.byId.set(sheet.id, index);
      const key = sheet.name.toUpperCase();
      if (!built!.byName.has(key)) {
        built!.byName.set(key, index);
      }
    });
    sheetIndexes.set(model.sheets, built);
  }
  return built;
}

/** The workbook position of a sheet id; -1 when there is none. */
export function sheetIndexOf(model: SpreadsheetModel, sheetId: string): number {
  return indexesOf(model).byId.get(sheetId) ?? -1;
}

/** The sheet with an id. */
export function sheetById(model: SpreadsheetModel, sheetId: string): Sheet | undefined {
  const index = sheetIndexOf(model, sheetId);
  return index < 0 ? undefined : model.sheets[index];
}

/** The workbook position of a sheet name, case-insensitively; -1 when
 * there is none. */
export function sheetIndexByName(model: SpreadsheetModel, name: string): number {
  return indexesOf(model).byName.get(name.toUpperCase()) ?? -1;
}

// --- references --------------------------------------------------------------------

/** A part moved by the host's offset from the origin (absolute parts stay);
 * null when it leaves the grid — or, `wrap` (SPEC-133, a defined name's
 * evaluation), wrapped modulo the grid. */
function movePart(part: RefPart, dRow: number, dCol: number, wrap = false): { row?: number; col?: number } | null {
  const moved: { row?: number; col?: number } = {};
  if (part.row !== undefined) {
    let row = part.rowAbs ? part.row : part.row + dRow;
    if (wrap) {
      row = ((row % MAX_ROWS) + MAX_ROWS) % MAX_ROWS;
    }
    if (row < 0 || row >= MAX_ROWS) {
      return null;
    }
    moved.row = row;
  }
  if (part.col !== undefined) {
    let col = part.colAbs ? part.col : part.col + dCol;
    if (wrap) {
      col = ((col % MAX_COLS) + MAX_COLS) % MAX_COLS;
    }
    if (col < 0 || col >= MAX_COLS) {
      return null;
    }
    moved.col = col;
  }
  return moved;
}

/** The sheet ids a reference's prefix names (the host's when it has none):
 * one, or every grid sheet between a 3-D reference's ends in workbook
 * order, either direction. #REF! for an unknown or foreign sheet. */
function referencedSheets(model: SpreadsheetModel, hostSheetId: string, ref: ReferenceNode): string[] | CellError {
  if (ref.sheet === undefined) {
    return [hostSheetId];
  }
  const first = sheetIndexByName(model, ref.sheet);
  if (first < 0) {
    return REF_ERROR;
  }
  if (ref.sheetEnd === undefined) {
    const sheet = model.sheets[first]!;
    return sheet.kind === "grid" ? [sheet.id] : REF_ERROR;
  }
  const last = sheetIndexByName(model, ref.sheetEnd);
  if (last < 0) {
    return REF_ERROR;
  }
  const ids: string[] = [];
  for (let index = Math.min(first, last); index <= Math.max(first, last); index += 1) {
    const sheet = model.sheets[index]!;
    if (sheet.kind === "grid") {
      ids.push(sheet.id);
    }
  }
  return ids.length > 0 ? ids : REF_ERROR;
}

/** A reference node at a host: its areas (one per sheet of a 3-D
 * reference), or #REF! (an unknown sheet, a part shifted off the grid). */
export function resolveReference(
  ref: ReferenceNode,
  ctx: { model: SpreadsheetModel; host: { sheetId: string; row: number; col: number }; origin: CellRef; wrapReferences?: true },
): RangeValue | CellError {
  const dRow = ctx.host.row - ctx.origin.row;
  const dCol = ctx.host.col - ctx.origin.col;
  const wrap = ctx.wrapReferences === true;
  const start = movePart(ref.start, dRow, dCol, wrap);
  const end = ref.end ? movePart(ref.end, dRow, dCol, wrap) : start;
  if (!start || !end) {
    return REF_ERROR;
  }
  const sheets = referencedSheets(ctx.model, ctx.host.sheetId, ref);
  if (isError(sheets)) {
    return sheets;
  }
  const wholeColumns = start.row === undefined || end.row === undefined;
  const wholeRows = start.col === undefined || end.col === undefined;
  const top = wholeColumns ? 0 : Math.min(start.row!, end.row!);
  const bottom = wholeColumns ? MAX_ROWS - 1 : Math.max(start.row!, end.row!);
  const left = wholeRows ? 0 : Math.min(start.col!, end.col!);
  const right = wholeRows ? MAX_COLS - 1 : Math.max(start.col!, end.col!);
  return { kind: "range", areas: sheets.map((sheetId) => ({ sheetId, top, left, bottom, right })) };
}

// --- areas -------------------------------------------------------------------------

export function areaRows(area: Area): number {
  return area.bottom - area.top + 1;
}

export function areaCols(area: Area): number {
  return area.right - area.left + 1;
}

export function areaSize(area: Area): number {
  return areaRows(area) * areaCols(area);
}

/** The index of the first row at or after `row` (rows.length when none). */
function firstRowAtOrAfter(rows: readonly Row[], row: number): number {
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (rows[mid]!.index < row) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** The index of the first cell at or after `col` (cells.length when none). */
function firstCellAtOrAfter(cells: Row["cells"], col: number): number {
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

/** Visits the stored cells of an area, row-major, each value read through
 * `ctx.valueAt` (the stored Cell beside it, for SUBTOTAL's look at its
 * formula); `visit` returns false to stop. Returns false when stopped. */
export function forEachStoredCell(
  ctx: Pick<EvalContext, "model" | "valueAt">,
  area: Area,
  visit: (value: ScalarValue, row: number, col: number, cell: Cell) => boolean | void,
): boolean {
  const sheet = sheetById(ctx.model, area.sheetId);
  if (!sheet) {
    return true;
  }
  const rows = sheet.rows;
  for (let r = firstRowAtOrAfter(rows, area.top); r < rows.length; r += 1) {
    const row = rows[r]!;
    if (row.index > area.bottom) {
      break;
    }
    const cells = row.cells;
    for (let c = firstCellAtOrAfter(cells, area.left); c < cells.length; c += 1) {
      const col = cells[c]!.col;
      if (col > area.right) {
        break;
      }
      if (visit(ctx.valueAt(area.sheetId, row.index, col), row.index, col, cells[c]!) === false) {
        return false;
      }
    }
  }
  return true;
}

/** Visits every scalar a value holds, row-major: a range's stored cells
 * (areas in order), an array's elements, or the scalar itself. `direct` is
 * true only for a scalar given directly. Returns false when stopped. */
export function forEachScalar(
  ctx: Pick<EvalContext, "model" | "valueAt">,
  value: EvalValue,
  visit: (value: ScalarValue, direct: boolean) => boolean | void,
): boolean {
  if (isRange(value)) {
    for (const area of value.areas) {
      if (!forEachStoredCell(ctx, area, (cell) => visit(cell, false))) {
        return false;
      }
    }
    return true;
  }
  if (isArray(value)) {
    for (const element of value.data) {
      if (visit(element, false) === false) {
        return false;
      }
    }
    return true;
  }
  return visit(value, true) !== false;
}

// --- arrays ------------------------------------------------------------------------

export function makeArray(rows: number, cols: number, data: ScalarValue[]): ArrayValue {
  return { kind: "array", rows, cols, data };
}

/** A scalar or array's shape (a scalar is 1×1). */
export function shapeOf(value: ScalarValue | ArrayValue): Shape {
  return isArray(value) ? { rows: value.rows, cols: value.cols } : { rows: 1, cols: 1 };
}

/** The shape two operands broadcast to (see the module header). */
export function broadcastShape(a: Shape, b: Shape): Shape {
  const rows = a.rows === 1 ? b.rows : b.rows === 1 ? a.rows : Math.max(a.rows, b.rows);
  const cols = a.cols === 1 ? b.cols : b.cols === 1 ? a.cols : Math.max(a.cols, b.cols);
  return { rows, cols };
}

/** The element of an operand at (row, col) of a broadcast shape: a single
 * row repeats down, a single column across, outside the operand is #N/A. */
export function elementAt(value: ScalarValue | ArrayValue, row: number, col: number): ScalarValue {
  if (!isArray(value)) {
    return value;
  }
  const r = value.rows === 1 ? 0 : row;
  const c = value.cols === 1 ? 0 : col;
  if (r >= value.rows || c >= value.cols) {
    return NA_ERROR;
  }
  return value.data[r * value.cols + c]!;
}

/** An array's top-left element. */
export function topLeft(value: ArrayValue): ScalarValue {
  return value.data[0] ?? EMPTY;
}

/** A single-area range as an array — empties as EMPTY — or #VALUE! for
 * several areas, #NUM! past MAX_ARRAY_CELLS. */
export function materialize(ctx: Pick<EvalContext, "model" | "valueAt">, range: RangeValue): ArrayValue | CellError {
  if (range.areas.length !== 1) {
    return VALUE_ERROR;
  }
  const area = range.areas[0]!;
  const rows = areaRows(area);
  const cols = areaCols(area);
  if (rows * cols > MAX_ARRAY_CELLS) {
    return NUM_ERROR;
  }
  const data = Array.from<ScalarValue>({ length: rows * cols }).fill(EMPTY);
  forEachStoredCell(ctx, area, (value, row, col) => {
    data[(row - area.top) * cols + (col - area.left)] = value;
  });
  return makeArray(rows, cols, data);
}

/** Any value as an array: a range materialised, a scalar as 1×1. */
export function toArray(ctx: Pick<EvalContext, "model" | "valueAt">, value: EvalValue): ArrayValue | CellError {
  if (isRange(value)) {
    return materialize(ctx, value);
  }
  if (isArray(value)) {
    return value;
  }
  return makeArray(1, 1, [value]);
}

/** Implicit intersection of a range with the formula's host cell (see the
 * module header). */
export function intersect(ctx: Pick<EvalContext, "valueAt" | "host">, range: RangeValue): ScalarValue {
  if (range.areas.length !== 1) {
    return VALUE_ERROR;
  }
  const area = range.areas[0]!;
  const { row, col } = ctx.host;
  if (area.top === area.bottom && area.left === area.right) {
    return ctx.valueAt(area.sheetId, area.top, area.left);
  }
  if (area.left === area.right) {
    return row >= area.top && row <= area.bottom ? ctx.valueAt(area.sheetId, row, area.left) : VALUE_ERROR;
  }
  if (area.top === area.bottom) {
    return col >= area.left && col <= area.right ? ctx.valueAt(area.sheetId, area.top, col) : VALUE_ERROR;
  }
  return VALUE_ERROR;
}

/** Any value reduced to one scalar the way a single-value position reads
 * it outside array mode: a range intersected, an array's top-left. */
export function toScalar(ctx: Pick<EvalContext, "valueAt" | "host">, value: EvalValue): ScalarValue {
  if (isRange(value)) {
    return intersect(ctx, value);
  }
  return isArray(value) ? topLeft(value) : value;
}

/** The first error among a value's scalars, row-major (none: undefined). */
export function firstError(ctx: Pick<EvalContext, "model" | "valueAt">, value: EvalValue): CellError | undefined {
  let found: CellError | undefined;
  forEachScalar(ctx, value, (scalar) => {
    if (isError(scalar)) {
      found = scalar;
      return false;
    }
    return true;
  });
  return found;
}
