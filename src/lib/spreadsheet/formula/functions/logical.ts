// Logical (SPEC-120 "The core functions"): IF IFS IFERROR IFNA AND OR NOT XOR
// TRUE FALSE SWITCH CHOOSE.
//
// - IF, IFS, SWITCH, CHOOSE, IFERROR and IFNA are lazy: they evaluate only
//   the branch they take (`=IF(TRUE,1,1/0)` is 1). Their branches are "any"
//   arguments, so a branch may return a reference (`SUM(IF(…,A1:A3,B1:B3))`).
// - AND, OR and XOR read logicals and numbers from references and arrays
//   (text and empties skipped) and coerce direct arguments logically; with
//   nothing to read they are #VALUE!.
// - CHOOSE truncates its index (#VALUE! outside 1..n); SWITCH compares as
//   `=` does and is #N/A with no match and no default; IFS is #N/A when no
//   condition holds.
import type { CellError } from "../../model";
import { forEachScalar } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { compareScalars, isError, NA_ERROR, toLogical, toNumber, VALUE_ERROR, type ScalarValue } from "../values";
import { arg, optional, repeating, scalarArg } from "./common";

/** Folds the logicals AND/OR/XOR read (see the module header); #VALUE!
 * when there is nothing to read, else the first error met. */
function logicals(args: ArgList, ctx: EvalContext, start: boolean, step: (acc: boolean, b: boolean) => boolean): boolean | CellError {
  let acc = start;
  let read = 0;
  let failure: CellError | undefined;
  for (let i = 0; i < args.length; i += 1) {
    forEachScalar(ctx, args.get(i), (value, direct) => {
      if (isError(value)) {
        failure = value;
        return false;
      }
      if (typeof value === "boolean" || typeof value === "number") {
        acc = step(acc, value === true || (typeof value === "number" && value !== 0));
        read += 1;
      } else if (direct) {
        const b = toLogical(value);
        if (isError(b)) {
          failure = b;
          return false;
        }
        acc = step(acc, b);
        read += 1;
      }
      return true;
    });
    if (failure) {
      return failure;
    }
  }
  return read === 0 ? VALUE_ERROR : acc;
}

function logicalFold(name: string, summary: string, start: boolean, step: (acc: boolean, b: boolean) => boolean): FunctionSpec {
  return {
    name,
    category: "Logical",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any", "any"],
    repeat: { from: 1, count: 1 },
    summary,
    args: [arg("logical1", "A condition, or a range or array of logicals."), repeating("logical2", "More conditions, ranges or arrays.")],
    impl(args, ctx) {
      return logicals(args, ctx, start, step);
    },
  };
}

function constant(name: "TRUE" | "FALSE"): FunctionSpec {
  return {
    name,
    category: "Logical",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    summary: `Returns the logical value ${name}.`,
    args: [],
    impl() {
      return name === "TRUE";
    },
  };
}

/** A condition argument as a logical. */
function condition(args: ArgList, index: number): boolean | CellError {
  return toLogical(scalarArg(args, index));
}

function errorFallback(name: "IFERROR" | "IFNA", summary: string, caught: (error: CellError) => boolean): FunctionSpec {
  return {
    name,
    category: "Logical",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    lazy: true,
    acceptsErrors: true,
    summary,
    args: [arg("value", "The value to check."), arg(name === "IFERROR" ? "value_if_error" : "value_if_na", "What to return instead.")],
    impl(args) {
      const value = args.get(0);
      return isError(value) && caught(value) ? args.get(1) : value;
    },
  };
}

export const LOGICAL_FUNCTIONS: FunctionSpec[] = [
  {
    name: "IF",
    category: "Logical",
    minArgs: 1,
    maxArgs: 3,
    argKinds: ["value", "any", "any"],
    lazy: true,
    summary: "Returns one value when a condition is TRUE and another when it is FALSE.",
    args: [
      arg("logical_test", "The condition."),
      optional("value_if_true", "The result when it is TRUE."),
      optional("value_if_false", "The result when it is FALSE (default FALSE)."),
    ],
    impl(args) {
      const test = condition(args, 0);
      if (isError(test)) {
        return test;
      }
      if (test) {
        return args.length > 1 ? args.get(1) : true;
      }
      return args.length > 2 ? args.get(2) : false;
    },
  },
  {
    name: "IFS",
    category: "Logical",
    minArgs: 2,
    maxArgs: 254,
    argKinds: ["value", "any", "value", "any"],
    repeat: { from: 2, count: 2 },
    lazy: true,
    summary: "Returns the value of the first condition that is TRUE.",
    args: [
      arg("logical_test1", "The first condition."),
      arg("value_if_true1", "The result when it is TRUE."),
      repeating("logical_test2", "A further condition, tested in order."),
      repeating("value_if_true2", "The result when it is the first TRUE one."),
    ],
    impl(args) {
      if (args.length % 2 !== 0) {
        return VALUE_ERROR;
      }
      for (let i = 0; i < args.length; i += 2) {
        const test = condition(args, i);
        if (isError(test)) {
          return test;
        }
        if (test) {
          return args.get(i + 1);
        }
      }
      return NA_ERROR;
    },
  },
  errorFallback("IFERROR", "Returns a value, or another when the value is an error.", () => true),
  errorFallback("IFNA", "Returns a value, or another when the value is #N/A.", (error) => error.error === "#N/A"),
  logicalFold("AND", "Returns TRUE when every argument is TRUE.", true, (acc, b) => acc && b),
  logicalFold("OR", "Returns TRUE when any argument is TRUE.", false, (acc, b) => acc || b),
  {
    name: "NOT",
    category: "Logical",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Reverses a logical value.",
    args: [arg("logical", "The value to reverse.")],
    impl(args) {
      const b = condition(args, 0);
      return isError(b) ? b : !b;
    },
  },
  logicalFold("XOR", "Returns TRUE when an odd number of arguments are TRUE.", false, (acc, b) => acc !== b),
  constant("TRUE"),
  constant("FALSE"),
  {
    name: "SWITCH",
    category: "Logical",
    minArgs: 3,
    maxArgs: 254,
    argKinds: ["value", "value", "any", "value", "any"],
    repeat: { from: 3, count: 2 },
    lazy: true,
    summary: "Returns the result paired with the first value that equals an expression.",
    args: [
      arg("expression", "The value to match."),
      arg("value1", "The first value to compare it with."),
      arg("result1", "The result when it matches."),
      repeating("default_or_value2", "A further value to compare; a last unpaired argument is the default."),
      repeating("result2", "The result when that value matches."),
    ],
    impl(args, ctx) {
      const expression = scalarArg(args, 0);
      if (isError(expression)) {
        return expression;
      }
      const pairs = Math.floor((args.length - 1) / 2);
      for (let k = 0; k < pairs; k += 1) {
        const order = compareScalars(expression, scalarArg(args, 1 + 2 * k), ctx.settings);
        if (isError(order)) {
          return order;
        }
        if (order === 0) {
          return args.get(2 + 2 * k);
        }
      }
      return (args.length - 1) % 2 === 1 ? args.get(args.length - 1) : NA_ERROR;
    },
  },
  {
    name: "CHOOSE",
    category: "Lookup & Reference",
    minArgs: 2,
    maxArgs: 255,
    argKinds: ["value", "any", "any"],
    repeat: { from: 2, count: 1 },
    lazy: true,
    summary: "Returns one of its values by position.",
    args: [arg("index_num", "Which value to return, from 1."), arg("value1", "The first value to choose from."), repeating("value2", "A further value to choose from.")],
    impl(args, ctx) {
      const index = toNumber(scalarArg(args, 0) as ScalarValue, ctx.settings);
      if (isError(index)) {
        return index;
      }
      const position = Math.trunc(index);
      return position < 1 || position >= args.length ? VALUE_ERROR : args.get(position);
    },
  },
];
