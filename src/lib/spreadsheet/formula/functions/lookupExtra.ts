// Lookup & Reference, SPEC-127's library: XLOOKUP XMATCH LOOKUP ADDRESS AREAS
// TRANSPOSE HYPERLINK FORMULATEXT.
//
// - XLOOKUP/XMATCH search one row or one column. Match modes: 0 exact (no
//   wildcards, case-insensitive), −1 exact or the next smaller item, 1 exact
//   or the next larger, 2 wildcards (`*`, `?`, `~`). Search modes: 1 first to
//   last, −1 last to first, 2 and −2 a binary search of ascending and
//   descending data through SPEC-120's approximate-match bisection (largest
//   value ≤ the target; what unsorted data gives is untested). Wildcards
//   with a binary search, or any other mode, is #VALUE!. XLOOKUP's
//   return_array must be as long as the lookup_array along the search
//   (#VALUE! otherwise); not found is `if_not_found` when given, else #N/A.
//   One matched cell of a range is returned as a reference to it; a matched
//   row or column of a two-dimensional return_array is an array, so a plain
//   cell shows its first element (SPEC-120's single-cell rule) and an array
//   formula all of it.
// - LOOKUP's vector form bisects lookup_vector (largest value ≤ the target)
//   and reads the same position of result_vector (default: the vector
//   itself); its array form searches the first row of an array wider than
//   tall, else the first column, and returns from the last row or column.
// - ADDRESS writes `$A$1`, `A$1`, `$A1` or `A1` (abs_num 1–4), or in R1C1
//   `R2C3`, `R2C[3]`, `R[2]C3`, `R[2]C[3]`, after `sheet_text!` quoted as
//   SPEC-119 quotes a sheet name.
// - AREAS counts the areas of a (union) reference; TRANSPOSE swaps an
//   array's rows and columns (a legacy array function); HYPERLINK returns
//   its friendly_name when given, else its link (SPEC-135 makes it a link);
//   FORMULATEXT is "=" + the formula of a reference's top-left cell, #N/A for
//   a constant.
import { colToLetters, getCell } from "../../address";
import { MAX_COLS, MAX_ROWS, type CellError } from "../../model";
import { makeArray, sheetById, toArray } from "../areas";
import type { EvalContext } from "../evaluate";
import { quoteSheetName } from "../print";
import type { ArgList, FunctionSpec } from "../registry";
import {
  compareScalars,
  isEmpty,
  isError,
  isRange,
  NA_ERROR,
  toLogical,
  toText,
  VALUE_ERROR,
  type ArrayValue,
  type EvalValue,
  type ScalarValue,
} from "../values";
import { arg, integerArg, optional, scalarArg, singleArea } from "./common";
import { lookupMatcher } from "./criteria";
import { bisect, tableOf, type Table } from "./lookup";

/** A lookup vector: one row or one column of a table, its length and reader. */
interface Vector {
  length: number;
  vertical: boolean;
  at(i: number): ScalarValue;
}

/** A table's one row or column as a vector; null for a 2-D table. */
function vectorOf(table: Table): Vector | null {
  if (table.rows !== 1 && table.cols !== 1) {
    return null;
  }
  const vertical = table.cols === 1;
  return { length: vertical ? table.rows : table.cols, vertical, at: (i) => (vertical ? table.at(i, 0) : table.at(0, i)) };
}

const KIND_ORDER = (value: ScalarValue): string | null =>
  typeof value === "number" || typeof value === "string" || typeof value === "boolean" ? typeof value : null;

/** The position XLOOKUP/XMATCH find, or -1 (see the module header); #VALUE!
 * for a mode they do not know. */
function findPosition(lookup: ScalarValue, vector: Vector, matchMode: number, searchMode: number, ctx: EvalContext): number | CellError {
  if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) {
    return VALUE_ERROR;
  }
  const same = (value: ScalarValue): boolean => !isEmpty(value) && !isError(value) && compareScalars(value, lookup, ctx.settings) === 0;
  if (searchMode === 2 || searchMode === -2) {
    if (matchMode === 2) {
      return VALUE_ERROR;
    }
    const descending = searchMode === -2;
    const found = bisect(vector.length, vector.at, lookup, descending, ctx);
    if (found >= 0 && same(vector.at(found))) {
      return found;
    }
    // `found` is the last item on the target's near side: the next item
    // past it is on the far side.
    if (matchMode === 0) {
      return -1;
    }
    const wantsSmaller = matchMode === -1;
    if (wantsSmaller !== descending) {
      return found;
    }
    const next = found + 1;
    return next < vector.length && KIND_ORDER(vector.at(next)) === KIND_ORDER(lookup) ? next : -1;
  }
  const order = Array.from({ length: vector.length }, (_, i) => (searchMode === 1 ? i : vector.length - 1 - i));
  if (matchMode === 2) {
    const matches = lookupMatcher(lookup, ctx.settings);
    return order.find((i) => {
      const value = vector.at(i);
      return !isEmpty(value) && matches(value);
    }) ?? -1;
  }
  let best = -1;
  let bestValue: ScalarValue | undefined;
  for (const i of order) {
    const value = vector.at(i);
    if (isEmpty(value) || isError(value) || KIND_ORDER(value) !== KIND_ORDER(lookup)) {
      continue;
    }
    const cmp = compareScalars(value, lookup, ctx.settings);
    if (cmp === 0) {
      return i;
    }
    if (matchMode === 0 || isError(cmp)) {
      continue;
    }
    // −1: the largest item below the target; 1: the smallest above it.
    const onSide = matchMode === -1 ? cmp < 0 : cmp > 0;
    if (!onSide) {
      continue;
    }
    if (bestValue === undefined) {
      best = i;
      bestValue = value;
      continue;
    }
    const versus = compareScalars(value, bestValue, ctx.settings);
    if (!isError(versus) && (matchMode === -1 ? versus > 0 : versus < 0)) {
      best = i;
      bestValue = value;
    }
  }
  return best;
}

/** A mode argument: its default when omitted or empty. */
function mode(args: ArgList, index: number, ctx: EvalContext, fallback: number): number | CellError {
  return args.isMissing(index) ? fallback : integerArg(args, index, ctx);
}

/** XLOOKUP's result at a position of the search: the matched cell (a
 * reference for a range), or the matched row/column as an array. */
function resultAt(source: EvalValue, table: Table, vertical: boolean, position: number): EvalValue {
  if (isRange(source) && source.areas.length === 1) {
    const area = source.areas[0]!;
    const single = vertical ? table.cols === 1 : table.rows === 1;
    if (single) {
      const row = vertical ? area.top + position : area.top;
      const col = vertical ? area.left : area.left + position;
      return { kind: "range", areas: [{ sheetId: area.sheetId, top: row, left: col, bottom: row, right: col }] };
    }
  }
  if (vertical) {
    return makeArray(1, table.cols, Array.from({ length: table.cols }, (_, c) => table.at(position, c)));
  }
  return makeArray(table.rows, 1, Array.from({ length: table.rows }, (_, r) => table.at(r, position)));
}

/** An R1C1 or A1 coordinate: `R2`/`R[2]`, `$A`/`A`. */
function addressText(row: number, col: number, absNum: number, a1: boolean): string {
  const rowAbsolute = absNum === 1 || absNum === 2;
  const colAbsolute = absNum === 1 || absNum === 3;
  if (!a1) {
    return `${rowAbsolute ? `R${row}` : `R[${row}]`}${colAbsolute ? `C${col}` : `C[${col}]`}`;
  }
  return `${colAbsolute ? "$" : ""}${colToLetters(col - 1)}${rowAbsolute ? "$" : ""}${row}`;
}

export const LOOKUP_EXTRA_FUNCTIONS: FunctionSpec[] = [
  {
    name: "XLOOKUP",
    category: "Lookup & Reference",
    minArgs: 3,
    maxArgs: 6,
    argKinds: ["value", "any", "any", "any", "value", "value"],
    summary: "Searches a row or column for a value and returns the matching item of another.",
    args: [
      arg("lookup_value", "The value to search for."),
      arg("lookup_array", "The row or column to search."),
      arg("return_array", "The row, column or range to return from, as long as the lookup_array."),
      optional("if_not_found", "What to return when nothing matches (default #N/A)."),
      optional("match_mode", "0 exact (default), -1 exact or next smaller, 1 exact or next larger, 2 wildcards."),
      optional("search_mode", "1 first to last (default), -1 last to first, 2 binary ascending, -2 binary descending."),
    ],
    impl(args, ctx) {
      const lookup = scalarArg(args, 0);
      if (isError(lookup)) {
        return lookup;
      }
      const lookupTable = tableOf(args.get(1), ctx);
      if (isError(lookupTable)) {
        return lookupTable;
      }
      const source = args.get(2);
      const returnTable = tableOf(source, ctx);
      if (isError(returnTable)) {
        return returnTable;
      }
      const matchMode = mode(args, 4, ctx, 0);
      if (isError(matchMode)) {
        return matchMode;
      }
      const searchMode = mode(args, 5, ctx, 1);
      if (isError(searchMode)) {
        return searchMode;
      }
      const vector = vectorOf(lookupTable);
      if (!vector) {
        return VALUE_ERROR;
      }
      // A one-cell lookup_array searches whichever way return_array is one long.
      const vertical = vector.length === 1 ? returnTable.rows === 1 : vector.vertical;
      if ((vertical ? returnTable.rows : returnTable.cols) !== vector.length) {
        return VALUE_ERROR;
      }
      const position = findPosition(lookup, vector, matchMode, searchMode, ctx);
      if (isError(position)) {
        return position;
      }
      if (position < 0) {
        return args.length > 3 && !args.isMissing(3) ? args.get(3) : NA_ERROR;
      }
      return resultAt(source, returnTable, vertical, position);
    },
  },
  {
    name: "XMATCH",
    category: "Lookup & Reference",
    minArgs: 2,
    maxArgs: 4,
    argKinds: ["value", "any", "value", "value"],
    summary: "Returns the position of a value in a row or column, searching in either direction.",
    args: [
      arg("lookup_value", "The value to search for."),
      arg("lookup_array", "The row or column to search."),
      optional("match_mode", "0 exact (default), -1 exact or next smaller, 1 exact or next larger, 2 wildcards."),
      optional("search_mode", "1 first to last (default), -1 last to first, 2 binary ascending, -2 binary descending."),
    ],
    impl(args, ctx) {
      const lookup = scalarArg(args, 0);
      const table = tableOf(args.get(1), ctx);
      if (isError(table)) {
        return table;
      }
      const matchMode = mode(args, 2, ctx, 0);
      if (isError(matchMode)) {
        return matchMode;
      }
      const searchMode = mode(args, 3, ctx, 1);
      if (isError(searchMode)) {
        return searchMode;
      }
      const vector = vectorOf(table);
      if (!vector) {
        return VALUE_ERROR;
      }
      const position = findPosition(lookup, vector, matchMode, searchMode, ctx);
      if (isError(position)) {
        return position;
      }
      return position < 0 ? NA_ERROR : position + 1;
    },
  },
  {
    name: "LOOKUP",
    category: "Lookup & Reference",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "any", "any"],
    summary: "Looks a value up in a sorted row or column and returns the matching item of another.",
    args: [
      arg("lookup_value", "The value to search for."),
      arg("lookup_vector", "A sorted row or column to search, or an array (the array form)."),
      optional("result_vector", "The row or column to return from (default the searched one)."),
    ],
    impl(args, ctx) {
      const lookup = scalarArg(args, 0);
      const table = tableOf(args.get(1), ctx);
      if (isError(table)) {
        return table;
      }
      const hasResult = args.length > 2 && !args.isMissing(2);
      let search: Vector;
      let result: (i: number) => ScalarValue;
      let resultLength: number;
      const vector = vectorOf(table);
      if (vector) {
        search = vector;
        result = vector.at;
        resultLength = vector.length;
      } else {
        if (hasResult) {
          return NA_ERROR;
        }
        // The array form: a wider-than-tall array searches its first row.
        const byRow = table.cols > table.rows;
        search = {
          vertical: !byRow,
          length: byRow ? table.cols : table.rows,
          at: (i) => (byRow ? table.at(0, i) : table.at(i, 0)),
        };
        result = (i) => (byRow ? table.at(table.rows - 1, i) : table.at(i, table.cols - 1));
        resultLength = search.length;
      }
      if (hasResult) {
        const resultTable = tableOf(args.get(2), ctx);
        if (isError(resultTable)) {
          return resultTable;
        }
        const resultVector = vectorOf(resultTable);
        if (!resultVector) {
          return NA_ERROR;
        }
        result = resultVector.at;
        resultLength = resultVector.length;
      }
      const found = bisect(search.length, search.at, lookup, false, ctx);
      if (found < 0) {
        return NA_ERROR;
      }
      return found < resultLength ? result(found) : NA_ERROR;
    },
  },
  {
    name: "ADDRESS",
    category: "Lookup & Reference",
    minArgs: 2,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns a cell address as text from a row and a column number.",
    args: [
      arg("row_num", "The row number, from 1."),
      arg("column_num", "The column number, from 1."),
      optional("abs_num", "1 or omitted: $A$1; 2: A$1; 3: $A1; 4: A1."),
      optional("a1", "TRUE or omitted: the A1 style; FALSE: the R1C1 style."),
      optional("sheet_text", "A sheet name to put before the address."),
    ],
    impl(args, ctx) {
      const row = integerArg(args, 0, ctx);
      if (isError(row)) {
        return row;
      }
      const col = integerArg(args, 1, ctx);
      if (isError(col)) {
        return col;
      }
      const absNum = args.isMissing(2) ? 1 : integerArg(args, 2, ctx);
      if (isError(absNum)) {
        return absNum;
      }
      const a1 = args.isMissing(3) ? true : toLogical(scalarArg(args, 3));
      if (isError(a1)) {
        return a1;
      }
      if (row < 1 || row > MAX_ROWS || col < 1 || col > MAX_COLS || absNum < 1 || absNum > 4) {
        return VALUE_ERROR;
      }
      const address = addressText(row, col, absNum, a1);
      if (args.length < 5 || args.isMissing(4)) {
        return address;
      }
      const sheet = toText(scalarArg(args, 4));
      if (isError(sheet)) {
        return sheet;
      }
      return `${quoteSheetName(sheet)}!${address}`;
    },
  },
  {
    name: "AREAS",
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["range"],
    summary: "Returns the number of areas in a reference.",
    args: [arg("reference", "A reference; a union such as (A1:B2,D4) has several areas.")],
    impl(args) {
      const value = args.get(0);
      return isRange(value) ? value.areas.length : VALUE_ERROR;
    },
  },
  {
    name: "TRANSPOSE",
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["array"],
    summary: "Swaps the rows and columns of an array or range.",
    args: [arg("array", "The array or range to transpose.")],
    impl(args, ctx) {
      const array = toArray(ctx, args.get(0));
      if (isError(array)) {
        return array;
      }
      return transpose(array);
    },
  },
  {
    name: "HYPERLINK",
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns a link's friendly name, or the link itself when it has none.",
    args: [arg("link_location", "The address the link goes to."), optional("friendly_name", "The text or value shown instead of the address.")],
    impl(args) {
      return args.length > 1 && !args.isMissing(1) ? scalarArg(args, 1) : scalarArg(args, 0);
    },
  },
  {
    name: "FORMULATEXT",
    category: "Lookup & Reference",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["range"],
    summary: "Returns the formula of a cell as text.",
    args: [arg("reference", "The cell whose formula to show (a range's top-left cell).")],
    impl(args, ctx) {
      const area = singleArea(args.get(0));
      if (isError(area)) {
        return area;
      }
      const sheet = sheetById(ctx.model, area.sheetId);
      const formula = sheet ? getCell(sheet, area.top, area.left)?.formula : undefined;
      return formula === undefined ? NA_ERROR : `=${formula}`;
    },
  },
];

/** An array with its rows and columns swapped. */
export function transpose(array: ArrayValue): ArrayValue {
  const data: ScalarValue[] = [];
  for (let c = 0; c < array.cols; c += 1) {
    for (let r = 0; r < array.rows; r += 1) {
      data.push(array.data[r * array.cols + c]!);
    }
  }
  return makeArray(array.cols, array.rows, data);
}
