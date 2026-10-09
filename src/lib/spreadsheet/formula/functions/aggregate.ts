// Statistical aggregates (SPEC-120 "The core functions"): AVERAGE AVERAGEIF
// AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MIN MAX MINIFS MAXIFS.
//
// - AVERAGE, MIN and MAX read numbers only from references and arrays and
//   coerce arguments given directly (`=MAX("a")` is #VALUE!); AVERAGE* of
//   nothing is #DIV/0!, MIN/MAX/MINIFS/MAXIFS of nothing 0.
// - COUNT counts numbers in references and arrays, and direct arguments
//   that are numbers or coerce to one (errors never count). COUNTA counts
//   every non-empty cell — `""` results and errors included — and every
//   direct argument. COUNTBLANK counts empty cells and `""` results, the
//   empties from the area's size.
// - AVERAGEIF resizes its average range to the criteria range's shape from
//   its top-left cell; the *IFS family needs every range the same shape.
import type { CellError } from "../../model";
import { areaSize, forEachScalar, forEachStoredCell } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { DIV0_ERROR, isEmpty, isError, isRange, toNumber, type Area } from "../values";
import { arg, countMatches, criteriaPairs, eachNumber, forEachMatch, optional, repeating, resizedArea, scalarArg, singleArea, type CriteriaPair } from "./common";
import { parseCriterion } from "./criteria";

/** The numbers among the matched cells of `values`, or the first error. */
function matchedNumbers(ctx: EvalContext, values: Area, pairs: readonly CriteriaPair[], visit: (n: number) => void): CellError | undefined {
  let failure: CellError | undefined;
  forEachMatch(ctx, values, pairs, (value) => {
    if (isError(value)) {
      failure = value;
      return false;
    }
    if (typeof value === "number") {
      visit(value);
    }
    return true;
  });
  return failure;
}

/** The value area and criteria pairs of an *IFS call (value range first). */
function ifsArguments(args: ArgList, ctx: EvalContext): { values: Area; pairs: CriteriaPair[] } | CellError {
  const values = singleArea(args.get(0));
  if (isError(values)) {
    return values;
  }
  const pairs = criteriaPairs(args, 1, ctx, values);
  return isError(pairs) ? pairs : { values, pairs };
}

/** MIN or MAX over numbers fed by `feed`; 0 when there are none. */
function extreme(pick: (a: number, b: number) => number, feed: (visit: (n: number) => void) => CellError | undefined): number | CellError {
  let best: number | undefined;
  const failure = feed((n) => {
    best = best === undefined ? n : pick(best, n);
  });
  return failure ?? best ?? 0;
}

function minMax(name: "MIN" | "MAX", pick: (a: number, b: number) => number): FunctionSpec {
  return {
    name,
    category: "Statistical",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: name === "MIN" ? "Returns the smallest number." : "Returns the largest number.",
    args: [repeating("number", "A number, or a range or array of numbers.")],
    impl(args, ctx) {
      return extreme(pick, (visit) => eachNumber(args, ctx, visit));
    },
  };
}

function minMaxIfs(name: "MINIFS" | "MAXIFS", pick: (a: number, b: number) => number): FunctionSpec {
  return {
    name,
    category: "Statistical",
    minArgs: 3,
    maxArgs: 255,
    argKinds: ["range", "range", "value"],
    repeat: { from: 1, count: 2 },
    summary: name === "MINIFS" ? "Returns the smallest number among cells that meet every criterion." : "Returns the largest number among cells that meet every criterion.",
    args: [
      arg(name === "MINIFS" ? "min_range" : "max_range", "The cells whose numbers are compared."),
      repeating("criteria_range", "The cells a criterion is tested against."),
      repeating("criteria", "The criterion for that range."),
    ],
    impl(args, ctx) {
      const call = ifsArguments(args, ctx);
      return isError(call) ? call : extreme(pick, (visit) => matchedNumbers(ctx, call.values, call.pairs, visit));
    },
  };
}

/** sum / count, #DIV/0! for no numbers. */
function average(feed: (visit: (n: number) => void) => CellError | undefined): number | CellError {
  let sum = 0;
  let count = 0;
  const failure = feed((n) => {
    sum += n;
    count += 1;
  });
  if (failure) {
    return failure;
  }
  return count === 0 ? DIV0_ERROR : sum / count;
}

export const AGGREGATE_FUNCTIONS: FunctionSpec[] = [
  {
    name: "AVERAGE",
    category: "Statistical",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: "Returns the arithmetic mean of its arguments.",
    args: [repeating("number", "A number, or a range or array of numbers.")],
    impl(args, ctx) {
      return average((visit) => eachNumber(args, ctx, visit));
    },
  },
  {
    name: "AVERAGEIF",
    category: "Statistical",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["range", "value", "range"],
    summary: "Returns the mean of the cells that meet a criterion.",
    args: [
      arg("range", "The cells the criterion is tested against."),
      arg("criteria", "The criterion."),
      optional("average_range", "The cells to average, when not the range itself."),
    ],
    impl(args, ctx) {
      const range = singleArea(args.get(0));
      if (isError(range)) {
        return range;
      }
      let values = range;
      if (args.length > 2 && !args.isMissing(2)) {
        const given = singleArea(args.get(2));
        if (isError(given)) {
          return given;
        }
        values = resizedArea(given, range);
      }
      const pairs = [{ area: range, criterion: parseCriterion(scalarArg(args, 1), ctx.settings) }];
      return average((visit) => matchedNumbers(ctx, values, pairs, visit));
    },
  },
  {
    name: "AVERAGEIFS",
    category: "Statistical",
    minArgs: 3,
    maxArgs: 255,
    argKinds: ["range", "range", "value"],
    repeat: { from: 1, count: 2 },
    summary: "Returns the mean of the cells that meet every criterion.",
    args: [
      arg("average_range", "The cells to average."),
      repeating("criteria_range", "The cells a criterion is tested against."),
      repeating("criteria", "The criterion for that range."),
    ],
    impl(args, ctx) {
      const call = ifsArguments(args, ctx);
      return isError(call) ? call : average((visit) => matchedNumbers(ctx, call.values, call.pairs, visit));
    },
  },
  {
    name: "COUNT",
    category: "Statistical",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: "Counts the numbers among its arguments.",
    args: [repeating("value", "A value, or a range or array whose numbers are counted.")],
    impl(args, ctx) {
      let count = 0;
      for (let i = 0; i < args.length; i += 1) {
        forEachScalar(ctx, args.get(i), (value, direct) => {
          if (typeof value === "number" || (direct && !isError(value) && !isError(toNumber(value, ctx.settings)))) {
            count += 1;
          }
        });
      }
      return count;
    },
  },
  {
    name: "COUNTA",
    category: "Statistical",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: "Counts the values that are not empty.",
    args: [repeating("value", "A value, or a range or array whose non-empty values are counted.")],
    impl(args, ctx) {
      let count = 0;
      for (let i = 0; i < args.length; i += 1) {
        forEachScalar(ctx, args.get(i), (value, direct) => {
          if (direct || !isEmpty(value)) {
            count += 1;
          }
        });
      }
      return count;
    },
  },
  {
    name: "COUNTBLANK",
    category: "Statistical",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["range"],
    summary: "Counts the empty cells in a range.",
    args: [arg("range", "The cells to count.")],
    impl(args, ctx) {
      const value = args.get(0);
      if (!isRange(value)) {
        return isError(value) ? value : 0;
      }
      let blank = 0;
      for (const area of value.areas) {
        let filled = 0;
        forEachStoredCell(ctx, area, (cell) => {
          if (!isEmpty(cell) && cell !== "") {
            filled += 1;
          }
        });
        blank += areaSize(area) - filled;
      }
      return blank;
    },
  },
  {
    name: "COUNTIF",
    category: "Statistical",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["range", "value"],
    summary: "Counts the cells that meet a criterion.",
    args: [arg("range", "The cells to test."), arg("criteria", "The criterion: a value, or text such as \">5\" or \"b*\".")],
    impl(args, ctx) {
      const area = singleArea(args.get(0));
      if (isError(area)) {
        return area;
      }
      return countMatches(ctx, [{ area, criterion: parseCriterion(scalarArg(args, 1), ctx.settings) }]);
    },
  },
  {
    name: "COUNTIFS",
    category: "Statistical",
    minArgs: 2,
    maxArgs: 254,
    argKinds: ["range", "value"],
    repeat: { from: 0, count: 2 },
    summary: "Counts the positions where every criterion is met.",
    args: [repeating("criteria_range", "The cells a criterion is tested against."), repeating("criteria", "The criterion for that range.")],
    impl(args, ctx) {
      const pairs = criteriaPairs(args, 0, ctx);
      return isError(pairs) ? pairs : countMatches(ctx, pairs);
    },
  },
  minMax("MIN", Math.min),
  minMax("MAX", Math.max),
  minMaxIfs("MINIFS", Math.min),
  minMaxIfs("MAXIFS", Math.max),
];
