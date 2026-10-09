// Lookup & Reference (SPEC-120 "The core functions"): VLOOKUP HLOOKUP INDEX
// MATCH ROW COLUMN ROWS COLUMNS OFFSET INDIRECT SINGLE.
//
// - Exact mode (MATCH 0, VLOOKUP/HLOOKUP FALSE or an empty slot) is a linear
//   scan for the first match — wildcards for a text lookup value,
//   case-insensitive unless `caseSensitive`; empty cells never match.
//   Approximate mode (MATCH 1/−1, VLOOKUP/HLOOKUP TRUE or omitted) bisects
//   the vector, skipping elements of another type than the lookup value by
//   probing to the nearest same-type element below, then above — Excel's
//   documented answer on sorted data, whatever the bisection finds on
//   unsorted data (where Excel documents none).
// - INDEX given a range returns a REFERENCE (row or column 0 selects the
//   whole column or row; a one-row source indexes columns with a lone second
//   argument); out of bounds is #REF!. OFFSET returns a reference (#REF! off
//   the grid or for a height/width ≤ 0). INDIRECT accepts exactly one
//   reference or name token, optionally sheet-qualified (#REF! otherwise; a
//   second argument of FALSE — R1C1 text — is #REF!). OFFSET and INDIRECT
//   record their result areas as dynamic dependencies.
// - ROW()/COLUMN() of a multi-cell range in array mode return the vertical/
//   horizontal array of indices; SINGLE applies implicit intersection in any
//   mode (the wire form of Excel's `@`).
import { MAX_COLS, MAX_ROWS, type CellError } from "../../model";
import { areaCols, areaRows, forEachStoredCell, intersect, makeArray, resolveReference, topLeft } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { tokenize } from "../tokens";
import {
  collatorFor,
  isArray,
  isEmpty,
  isError,
  isRange,
  NA_ERROR,
  REF_ERROR,
  roundTo15,
  toLogical,
  VALUE_ERROR,
  type Area,
  type ArrayValue,
  type EvalValue,
  type RangeValue,
  type ScalarValue,
} from "../values";
import { arg, integerArg, numberOr, optional, scalarArg, singleArea } from "./common";
import { lookupMatcher } from "./criteria";

/** A table a lookup reads: a range's one area or an array, with a reader of
 * its cells by offset. */
interface Table {
  rows: number;
  cols: number;
  at(row: number, col: number): ScalarValue;
  /** Visits the non-empty cells of one column (or row), in order, until
   * `visit` returns true; the offset found, or -1. */
  scan(line: "col" | "row", index: number, visit: (value: ScalarValue) => boolean): number;
}

function tableOf(value: EvalValue, ctx: EvalContext): Table | CellError {
  if (isError(value)) {
    return value;
  }
  if (isRange(value)) {
    if (value.areas.length !== 1) {
      return REF_ERROR;
    }
    const area = value.areas[0]!;
    return {
      rows: areaRows(area),
      cols: areaCols(area),
      at: (row, col) => ctx.valueAt(area.sheetId, area.top + row, area.left + col),
      scan(line, index, visit) {
        const sub: Area =
          line === "col"
            ? { ...area, left: area.left + index, right: area.left + index }
            : { ...area, top: area.top + index, bottom: area.top + index };
        let found = -1;
        forEachStoredCell(ctx, sub, (cell, row, col) => {
          if (!isEmpty(cell) && visit(cell)) {
            found = line === "col" ? row - area.top : col - area.left;
            return false;
          }
          return true;
        });
        return found;
      },
    };
  }
  const array: ArrayValue = isArray(value) ? value : makeArray(1, 1, [value]);
  const at = (row: number, col: number): ScalarValue => array.data[row * array.cols + col]!;
  return {
    rows: array.rows,
    cols: array.cols,
    at,
    scan(line, index, visit) {
      const length = line === "col" ? array.rows : array.cols;
      for (let i = 0; i < length; i += 1) {
        const cell = line === "col" ? at(i, index) : at(index, i);
        if (!isEmpty(cell) && visit(cell)) {
          return i;
        }
      }
      return -1;
    },
  };
}

type Kind = "number" | "string" | "boolean";

function kindOf(value: ScalarValue): Kind | null {
  const type = typeof value;
  return type === "number" || type === "string" || type === "boolean" ? (type as Kind) : null;
}

/** Bisects a vector for the last element ≤ the lookup value (`descending`:
 * ≥), probing past elements of another type (see the module header). */
function bisect(length: number, read: (i: number) => ScalarValue, lookup: ScalarValue, descending: boolean, ctx: EvalContext): number {
  const kind = kindOf(lookup);
  if (kind === null) {
    return -1;
  }
  const collator = collatorFor(ctx.settings);
  const compare = (value: ScalarValue): number => {
    if (kind === "number") {
      return Math.sign(roundTo15(value as number) - roundTo15(lookup as number));
    }
    if (kind === "string") {
      return Math.sign(collator.compare(value as string, lookup as string));
    }
    return Number(value) - Number(lookup);
  };
  let lo = 0;
  let hi = length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    let probe = mid;
    while (probe >= lo && kindOf(read(probe)) !== kind) {
      probe -= 1;
    }
    if (probe < lo) {
      probe = mid + 1;
      while (probe <= hi && kindOf(read(probe)) !== kind) {
        probe += 1;
      }
      if (probe > hi) {
        break;
      }
    }
    const order = compare(read(probe));
    if (order === 0 || (descending ? order > 0 : order < 0)) {
      found = probe;
      lo = probe + 1;
    } else {
      hi = probe - 1;
    }
  }
  return found;
}

/** VLOOKUP/HLOOKUP: the lookup value searched along the table's first column
 * (row), the result read from the `index`th column (row). */
function tableLookup(name: "VLOOKUP" | "HLOOKUP"): FunctionSpec {
  const vertical = name === "VLOOKUP";
  return {
    name,
    category: "Lookup & Reference",
    minArgs: 3,
    maxArgs: 4,
    argKinds: ["value", "any", "value", "value"],
    summary: vertical
      ? "Looks a value up in the first column of a table and returns a value from the same row."
      : "Looks a value up in the first row of a table and returns a value from the same column.",
    args: [
      arg("lookup_value", "The value to look for."),
      arg("table_array", "The table to search."),
      arg(vertical ? "col_index_num" : "row_index_num", vertical ? "The column to return, from 1." : "The row to return, from 1."),
      optional("range_lookup", "TRUE or omitted: the closest match in sorted data; FALSE: an exact match."),
    ],
    impl(args, ctx) {
      const lookup = scalarArg(args, 0);
      const table = tableOf(args.get(1), ctx);
      if (isError(table)) {
        return table;
      }
      const index = integerArg(args, 2, ctx);
      if (isError(index)) {
        return index;
      }
      const approximate = args.length > 3 ? toLogical(scalarArg(args, 3)) : true;
      if (isError(approximate)) {
        return approximate;
      }
      if (index < 1) {
        return VALUE_ERROR;
      }
      if (index > (vertical ? table.cols : table.rows)) {
        return REF_ERROR;
      }
      let found: number;
      if (approximate) {
        const length = vertical ? table.rows : table.cols;
        found = bisect(length, (i) => (vertical ? table.at(i, 0) : table.at(0, i)), lookup, false, ctx);
      } else {
        const matches = lookupMatcher(lookup, ctx.settings);
        found = table.scan(vertical ? "col" : "row", 0, matches);
      }
      if (found < 0) {
        return NA_ERROR;
      }
      return vertical ? table.at(found, index - 1) : table.at(index - 1, found);
    },
  };
}

/** ROW/COLUMN: the index of a reference's first row (column), or every index
 * of a multi-cell range in array mode; the host's without an argument. */
function position(name: "ROW" | "COLUMN"): FunctionSpec {
  const rows = name === "ROW";
  return {
    name,
    category: "Lookup & Reference",
    minArgs: 0,
    maxArgs: 1,
    argKinds: ["range"],
    summary: rows ? "Returns the row number of a reference." : "Returns the column number of a reference.",
    args: [optional("reference", "The cell or range (default: the formula's own cell).")],
    impl(args, ctx) {
      if (args.length === 0 || args.isMissing(0)) {
        return (rows ? ctx.host.row : ctx.host.col) + 1;
      }
      const area = singleArea(args.get(0));
      if (isError(area)) {
        return area;
      }
      const first = rows ? area.top : area.left;
      const count = rows ? areaRows(area) : areaCols(area);
      if (!ctx.arrayMode || count === 1) {
        return first + 1;
      }
      const indices = Array.from({ length: count }, (_, i) => first + i + 1);
      return rows ? makeArray(count, 1, indices) : makeArray(1, count, indices);
    },
  };
}

/** ROWS/COLUMNS of a reference, an array or a scalar. */
function extent(name: "ROWS" | "COLUMNS"): FunctionSpec {
  const rows = name === "ROWS";
  return {
    name,
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["any"],
    summary: rows ? "Returns the number of rows in a reference or array." : "Returns the number of columns in a reference or array.",
    args: [arg("array", "The reference or array.")],
    impl(args) {
      const value = args.get(0);
      if (isError(value)) {
        return value;
      }
      if (isRange(value)) {
        if (value.areas.length !== 1) {
          return REF_ERROR;
        }
        return rows ? areaRows(value.areas[0]!) : areaCols(value.areas[0]!);
      }
      if (isArray(value)) {
        return rows ? value.rows : value.cols;
      }
      return 1;
    },
  };
}

/** INDEX over a range: the reference to a cell, a whole row or a whole
 * column of one of its areas. */
function indexRange(value: RangeValue, row: number, col: number | undefined, areaNumber: number): RangeValue | CellError {
  if (areaNumber < 1) {
    return VALUE_ERROR;
  }
  const area = value.areas[areaNumber - 1];
  if (!area) {
    return REF_ERROR;
  }
  const height = areaRows(area);
  const width = areaCols(area);
  let r = row;
  let c = col;
  if (c === undefined) {
    if (height === 1) {
      c = r;
      r = 1;
    } else {
      c = width === 1 ? 1 : 0;
    }
  }
  if (r < 0 || c < 0) {
    return VALUE_ERROR;
  }
  if (r > height || c > width) {
    return REF_ERROR;
  }
  const top = r === 0 ? area.top : area.top + r - 1;
  const bottom = r === 0 ? area.bottom : top;
  const left = c === 0 ? area.left : area.left + c - 1;
  const right = c === 0 ? area.right : left;
  return { kind: "range", areas: [{ sheetId: area.sheetId, top, left, bottom, right }] };
}

/** INDEX over an array: an element, or a whole row or column as an array. */
function indexArray(array: ArrayValue, row: number, col: number | undefined): EvalValue {
  let r = row;
  let c = col;
  if (c === undefined) {
    if (array.rows === 1) {
      c = r;
      r = 1;
    } else {
      c = array.cols === 1 ? 1 : 0;
    }
  }
  if (r < 0 || c < 0) {
    return VALUE_ERROR;
  }
  if (r > array.rows || c > array.cols) {
    return REF_ERROR;
  }
  if (r > 0 && c > 0) {
    return array.data[(r - 1) * array.cols + (c - 1)]!;
  }
  if (r > 0) {
    return makeArray(1, array.cols, array.data.slice((r - 1) * array.cols, r * array.cols));
  }
  if (c > 0) {
    return makeArray(
      array.rows,
      1,
      Array.from({ length: array.rows }, (_, i) => array.data[i * array.cols + c! - 1]!),
    );
  }
  return array;
}

/** A dimension argument of OFFSET: the default when omitted or empty. */
function dimension(args: ArgList, index: number, ctx: EvalContext, fallback: number): number | CellError {
  return args.isMissing(index) ? fallback : integerArg(args, index, ctx);
}

export const LOOKUP_FUNCTIONS: FunctionSpec[] = [
  tableLookup("VLOOKUP"),
  tableLookup("HLOOKUP"),
  {
    name: "INDEX",
    category: "Lookup & Reference",
    minArgs: 2,
    maxArgs: 4,
    argKinds: ["any", "value", "value", "value"],
    summary: "Returns the cell (a reference) or array element at a row and column.",
    args: [
      arg("array", "The range or array."),
      arg("row_num", "The row, from 1; 0 for the whole column."),
      optional("column_num", "The column, from 1; 0 for the whole row."),
      optional("area_num", "Which area of a multi-area reference (default 1)."),
    ],
    impl(args, ctx) {
      const source = args.get(0);
      if (isError(source)) {
        return source;
      }
      const row = integerArg(args, 1, ctx);
      if (isError(row)) {
        return row;
      }
      let col: number | undefined;
      if (args.length > 2 && !args.isMissing(2)) {
        const given = integerArg(args, 2, ctx);
        if (isError(given)) {
          return given;
        }
        col = given;
      }
      const areaNumber = numberOr(args, 3, ctx, 1);
      if (isError(areaNumber)) {
        return areaNumber;
      }
      if (isRange(source)) {
        return indexRange(source, row, col, Math.trunc(areaNumber));
      }
      return indexArray(isArray(source) ? source : makeArray(1, 1, [source]), row, col);
    },
  },
  {
    name: "MATCH",
    category: "Lookup & Reference",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "any", "value"],
    summary: "Returns the position of a value in a row or column.",
    args: [
      arg("lookup_value", "The value to find."),
      arg("lookup_array", "One row or one column to search."),
      optional("match_type", "1 or omitted: the largest value ≤ it (sorted ascending); 0: an exact match; -1: the smallest value ≥ it (sorted descending)."),
    ],
    impl(args, ctx) {
      const lookup = scalarArg(args, 0);
      const table = tableOf(args.get(1), ctx);
      if (isError(table)) {
        return table;
      }
      const type = numberOr(args, 2, ctx, 1);
      if (isError(type)) {
        return type;
      }
      if (table.rows !== 1 && table.cols !== 1) {
        return NA_ERROR;
      }
      const vertical = table.cols === 1 && table.rows > 1;
      const length = vertical ? table.rows : table.cols;
      let found: number;
      if (type === 0) {
        found = table.scan(vertical ? "col" : "row", 0, lookupMatcher(lookup, ctx.settings));
      } else {
        found = bisect(length, (i) => (vertical ? table.at(i, 0) : table.at(0, i)), lookup, type < 0, ctx);
      }
      return found < 0 ? NA_ERROR : found + 1;
    },
  },
  position("ROW"),
  position("COLUMN"),
  extent("ROWS"),
  extent("COLUMNS"),
  {
    name: "OFFSET",
    category: "Lookup & Reference",
    minArgs: 3,
    maxArgs: 5,
    argKinds: ["range", "value", "value", "value", "value"],
    volatile: true,
    summary: "Returns a reference offset from a starting reference by rows and columns.",
    args: [
      arg("reference", "The reference to start from."),
      arg("rows", "Rows down (negative: up)."),
      arg("cols", "Columns right (negative: left)."),
      optional("height", "The height of the result (default the reference's)."),
      optional("width", "The width of the result (default the reference's)."),
    ],
    impl(args, ctx) {
      const area = singleArea(args.get(0));
      if (isError(area)) {
        return area;
      }
      const values: number[] = [];
      for (const [index, fallback] of [
        [1, 0],
        [2, 0],
        [3, areaRows(area)],
        [4, areaCols(area)],
      ] as const) {
        const n = index < 3 ? integerArg(args, index, ctx) : dimension(args, index, ctx, fallback);
        if (isError(n)) {
          return n;
        }
        values.push(n);
      }
      const [rows, cols, height, width] = values as [number, number, number, number];
      const top = area.top + rows;
      const left = area.left + cols;
      const bottom = top + height - 1;
      const right = left + width - 1;
      if (height <= 0 || width <= 0 || top < 0 || left < 0 || bottom >= MAX_ROWS || right >= MAX_COLS) {
        return REF_ERROR;
      }
      const result: Area = { sheetId: area.sheetId, top, left, bottom, right };
      ctx.recordDependency(result);
      return { kind: "range", areas: [result] };
    },
  },
  {
    name: "INDIRECT",
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    volatile: true,
    summary: "Returns the reference a text names.",
    args: [arg("ref_text", "A reference as text, such as \"B2\" or \"Sheet2!A1:C3\"."), optional("a1", "TRUE or omitted: A1 style (R1C1 is not supported).")],
    impl(args, ctx) {
      const text = scalarArg(args, 0);
      if (typeof text !== "string") {
        return REF_ERROR;
      }
      if (args.length > 1) {
        const a1 = toLogical(scalarArg(args, 1));
        if (isError(a1) || !a1) {
          return REF_ERROR;
        }
      }
      const tokens = tokenize(text.trim()).filter((token) => token.kind !== "end");
      const only = tokens[0];
      if (tokens.length !== 1 || only?.kind !== "ref" || !only.ref) {
        return REF_ERROR;
      }
      // The text names absolute positions: no shift from the host.
      const resolved = resolveReference(only.ref, { model: ctx.model, host: ctx.host, origin: { row: ctx.host.row, col: ctx.host.col } });
      if (isError(resolved)) {
        return resolved;
      }
      for (const area of resolved.areas) {
        ctx.recordDependency(area);
      }
      return resolved;
    },
  },
  {
    name: "SINGLE",
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["any"],
    summary: "Returns one value by implicit intersection (Excel's @ operator).",
    args: [arg("value", "The reference, array or value to reduce.")],
    impl(args, ctx) {
      const value = args.get(0);
      if (isRange(value)) {
        return intersect(ctx, value);
      }
      return isArray(value) ? topLeft(value) : value;
    },
  },
];
