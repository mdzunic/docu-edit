// Engineering: the error functions (SPEC-127 "Engineering") — ERF
// ERF.PRECISE ERFC ERFC.PRECISE, computed by numeric.ts's erf and erfc
// (Cody's rational approximations, netlib specfun CALERF).
//
// - ERF(lower_limit) is erf(lower_limit); with upper_limit it is the
//   integral between the two, erf(upper_limit) − erf(lower_limit).
// - Negative arguments are allowed throughout (Excel 2010 and later):
//   erf is odd, erfc(−x) = 2 − erfc(x).
// - A non-numeric argument is #VALUE! (the usual value coercion).
import type { CellError } from "../../model";
import type { FunctionSpec } from "../registry";
import { isError } from "../values";
import { arg, numberArg, optional } from "./common";
import { erf, erfc } from "./numeric";

/** A one-argument error function over a number `x`. */
function unary(name: string, summary: string, description: string, f: (x: number) => number): FunctionSpec {
  return {
    name,
    category: "Engineering",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary,
    args: [arg("x", description)],
    impl(args, ctx): number | CellError {
      const x = numberArg(args, 0, ctx);
      return isError(x) ? x : f(x);
    },
  };
}

export const ERF_FUNCTIONS: FunctionSpec[] = [
  {
    name: "ERF",
    category: "Engineering",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the error function integrated from 0, or between two limits.",
    args: [
      arg("lower_limit", "The lower limit, or the upper one when upper_limit is omitted."),
      optional("upper_limit", "The upper limit of the integral."),
    ],
    impl(args, ctx) {
      const lower = numberArg(args, 0, ctx);
      if (isError(lower)) {
        return lower;
      }
      if (args.length < 2) {
        return erf(lower);
      }
      const upper = numberArg(args, 1, ctx);
      return isError(upper) ? upper : erf(upper) - erf(lower);
    },
  },
  unary("ERF.PRECISE", "Returns the error function integrated from 0 to a number.", "The upper limit of the integral from 0.", erf),
  unary("ERFC", "Returns the complementary error function, integrated from a number to infinity.", "The lower limit of the integral to infinity.", erfc),
  unary("ERFC.PRECISE", "Returns the complementary error function of a number.", "The lower limit of the integral to infinity.", erfc),
];
