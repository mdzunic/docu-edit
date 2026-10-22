// SUBTOTAL and AGGREGATE (SPEC-127 "SUBTOTAL and AGGREGATE") — the two
// visibility-sensitive functions:
//
// | function_num | 1/101 AVERAGE, 2/102 COUNT, 3/103 COUNTA, 4/104 MAX, 5/105 MIN, 6/106 PRODUCT,
// |              | 7/107 STDEV, 8/108 STDEVP, 9/109 SUM, 10/110 VAR, 11/111 VARP |
//
// - SUBTOTAL 1–11 skip filter-hidden rows (EvalContext.rowVisibility
//   "filtered", which SPEC-128 adds) and 101–111 every hidden row; any other
//   number is #VALUE!.
// - AGGREGATE adds 12 MEDIAN, 13 MODE.SNGL and the array form 14 LARGE, 15
//   SMALL, 16 PERCENTILE.INC, 17 QUARTILE.INC, 18 PERCENTILE.EXC, 19
//   QUARTILE.EXC (array, k). Options 0–3 ignore nested SUBTOTAL/AGGREGATE,
//   1, 3, 5 and 7 hidden rows, 2, 3, 6 and 7 error values; 4 ignores
//   nothing.
// - Both skip a cell whose formula calls SUBTOTAL or AGGREGATE (AGGREGATE
//   with options 0–3) — visibility.ts's memo. Hidden rows and nested cells
//   are skipped only inside references, never inside computed arrays;
//   hidden columns are never skipped.
// - Values are read as the aggregates read them (SPEC-120 "Coercion"):
//   numbers from references and arrays, the first error unless errors are
//   ignored; COUNT counts numbers and COUNTA every non-empty value.
import type { CellError } from "../../model";
import { forEachStoredCell } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { DIV0_ERROR, isArray, isEmpty, isError, isRange, toNumber, VALUE_ERROR, type EvalValue, type ScalarValue } from "../values";
import { callsVisibilitySensitive } from "../visibility";
import { arg, integerArg, repeating } from "./common";
import { kth, median, modeSingle, percentileExc, percentileInc, quartileExc, quartileInc, stdev, sum, variance } from "./stats";

/** What to leave out while reading. */
interface Skips {
  /** Rows of this visibility or worse: "filtered" skips filtered rows only,
   * "hidden" every row not visible, null none. */
  rows: "filtered" | "hidden" | null;
  nested: boolean;
  errors: boolean;
}

/** The values read: numbers, the non-empty count, the first error kept. */
interface Read {
  numbers: number[];
  nonEmpty: number;
  error?: CellError;
}

/** Reads every argument from `from` on (see the module header). */
function read(args: ArgList, from: number, ctx: EvalContext, skips: Skips): Read {
  const out: Read = { numbers: [], nonEmpty: 0 };
  const take = (value: ScalarValue, direct: boolean): boolean => {
    if (isError(value)) {
      if (skips.errors) {
        return true;
      }
      out.error ??= value;
      out.nonEmpty += 1;
      return true;
    }
    if (isEmpty(value)) {
      if (direct) {
        out.numbers.push(0);
        out.nonEmpty += 1;
      }
      return true;
    }
    out.nonEmpty += 1;
    if (typeof value === "number") {
      out.numbers.push(value);
    } else if (direct) {
      const n = toNumber(value, ctx.settings);
      if (isError(n)) {
        out.error ??= n;
      } else {
        out.numbers.push(n);
      }
    }
    return true;
  };
  for (let i = from; i < args.length; i += 1) {
    const value: EvalValue = args.get(i);
    if (isRange(value)) {
      for (const area of value.areas) {
        let lastRow = -1;
        let rowShows = true;
        forEachStoredCell(ctx, area, (cell, row, _col, stored) => {
          if (row !== lastRow) {
            lastRow = row;
            const visibility = skips.rows === null ? "visible" : ctx.rowVisibility(area.sheetId, row);
            rowShows = skips.rows === "filtered" ? visibility !== "filtered" : visibility === "visible";
          }
          if (!rowShows || (skips.nested && stored.formula !== undefined && callsVisibilitySensitive(stored.formula))) {
            return true;
          }
          return take(cell, false);
        });
      }
    } else if (isArray(value)) {
      value.data.forEach((element) => take(element, false));
    } else {
      take(value, true);
    }
  }
  return out;
}

/** SUBTOTAL's eleven functions (and AGGREGATE's first eleven) over what was
 * read. */
function basic(fn: number, values: Read): number | CellError {
  if (fn === 2) {
    return values.numbers.length;
  }
  if (fn === 3) {
    return values.nonEmpty;
  }
  if (values.error) {
    return values.error;
  }
  const xs = values.numbers;
  switch (fn) {
    case 1:
      return xs.length === 0 ? DIV0_ERROR : sum(xs) / xs.length;
    case 4:
      return xs.length === 0 ? 0 : Math.max(...xs);
    case 5:
      return xs.length === 0 ? 0 : Math.min(...xs);
    case 6:
      return xs.length === 0 ? 0 : xs.reduce((product, x) => product * x, 1);
    case 7:
      return stdev(xs, true);
    case 8:
      return stdev(xs, false);
    case 9:
      return sum(xs);
    case 10:
      return variance(xs, true);
    default:
      return variance(xs, false);
  }
}

/** A function_num or options argument, truncated. */
function whole(args: ArgList, index: number, ctx: EvalContext): number | CellError {
  return args.isMissing(index) ? 0 : integerArg(args, index, ctx);
}

/** AGGREGATE's k: a number however it was given. */
function kArg(value: EvalValue, ctx: EvalContext): number | CellError {
  let scalar: ScalarValue;
  if (isRange(value)) {
    const area = value.areas[0];
    scalar = value.areas.length === 1 && area && area.top === area.bottom && area.left === area.right ? ctx.valueAt(area.sheetId, area.top, area.left) : VALUE_ERROR;
  } else {
    scalar = isArray(value) ? (value.data[0] ?? VALUE_ERROR) : value;
  }
  return toNumber(scalar, ctx.settings);
}

export const SUBTOTAL_FUNCTIONS: FunctionSpec[] = [
  {
    name: "SUBTOTAL",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 255,
    argKinds: ["value", "range", "range"],
    repeat: { from: 2, count: 1 },
    visibilitySensitive: true,
    summary: "Returns a subtotal of references, leaving out other subtotals and hidden rows.",
    args: [
      arg("function_num", "1–11 (filtered rows skipped) or 101–111 (all hidden rows skipped): AVERAGE … VARP."),
      arg("ref1", "The first range to subtotal."),
      repeating("ref2", "More ranges to subtotal."),
    ],
    impl(args, ctx) {
      const fn = whole(args, 0, ctx);
      if (isError(fn)) {
        return fn;
      }
      const base = fn > 100 ? fn - 100 : fn;
      if (base < 1 || base > 11) {
        return VALUE_ERROR;
      }
      return basic(base, read(args, 1, ctx, { rows: fn > 100 ? "hidden" : "filtered", nested: true, errors: false }));
    },
  },
  {
    name: "AGGREGATE",
    category: "Math & Trig",
    minArgs: 3,
    maxArgs: 255,
    argKinds: ["value", "value", "array", "array"],
    repeat: { from: 3, count: 1 },
    visibilitySensitive: true,
    summary: "Returns an aggregate of references or an array, optionally ignoring errors and hidden rows.",
    args: [
      arg("function_num", "1–13 for AVERAGE … MODE.SNGL; 14–19 for LARGE … QUARTILE.EXC (the array form)."),
      arg("options", "0–3 ignore nested subtotals, plus 1/3/5/7 hidden rows and 2/3/6/7 errors; 4 nothing."),
      arg("ref1", "The first range, or the array of the array form."),
      repeating("ref2", "More ranges, or the k of the array form."),
    ],
    impl(args, ctx) {
      const fn = whole(args, 0, ctx);
      if (isError(fn)) {
        return fn;
      }
      const options = whole(args, 1, ctx);
      if (isError(options)) {
        return options;
      }
      if (fn < 1 || fn > 19 || options < 0 || options > 7) {
        return VALUE_ERROR;
      }
      const skips: Skips = {
        rows: options % 2 === 1 ? "hidden" : null,
        nested: options <= 3,
        errors: options === 2 || options === 3 || options === 6 || options === 7,
      };
      if (fn <= 11) {
        return basic(fn, read(args, 2, ctx, skips));
      }
      if (fn <= 13) {
        const values = read(args, 2, ctx, skips);
        if (values.error) {
          return values.error;
        }
        return fn === 12 ? median(values.numbers) : modeSingle(values.numbers);
      }
      // The array form: (array, k).
      if (args.length !== 4) {
        return VALUE_ERROR;
      }
      const k = kArg(args.get(3), ctx);
      if (isError(k)) {
        return k;
      }
      const values = read({ length: 3, get: (i) => args.get(i), isMissing: (i) => args.isMissing(i) }, 2, ctx, skips);
      if (values.error) {
        return values.error;
      }
      const xs = values.numbers;
      switch (fn) {
        case 14:
          return kth(xs, k, true);
        case 15:
          return kth(xs, k, false);
        case 16:
          return percentileInc(xs, k);
        case 17:
          return quartileInc(xs, k);
        case 18:
          return percentileExc(xs, k);
        default:
          return quartileExc(xs, k);
      }
    },
  },
];
