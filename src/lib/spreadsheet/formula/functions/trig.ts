// The trigonometric, inverse, reciprocal and hyperbolic functions (SPEC-127
// "Function pins"): SIN COS TAN ASIN ACOS ATAN ATAN2 SINH COSH TANH ASINH
// ACOSH ATANH COT CSC SEC ACOT, in radians, with the domain limits
// Microsoft's pages state: ASIN/ACOS need −1…1, ACOSH ≥ 1, ATANH strictly
// inside −1…1 (#NUM! otherwise); COT, CSC and SEC of |x| ≥ 2^27 are #NUM!
// and COT/CSC of 0 #DIV/0!; ATAN2(x_num, y_num) keeps Excel's argument
// order and ATAN2(0,0) is #DIV/0!. ACOT answers in 0…π. A non-finite result
// (SINH of 1000) is #NUM! by the evaluator's rule.
import type { CellError } from "../../model";
import type { FunctionSpec } from "../registry";
import { DIV0_ERROR, isError, NUM_ERROR } from "../values";
import { arg, numberArg } from "./common";

/** The reciprocal functions' bound: |x| must stay below 2^27. */
const RECIPROCAL_LIMIT = 2 ** 27;

function trig(name: string, summary: string, description: string, f: (x: number) => number | CellError): FunctionSpec {
  return {
    name,
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary,
    args: [arg("number", description)],
    impl(args, ctx) {
      const x = numberArg(args, 0, ctx);
      return isError(x) ? x : f(x);
    },
  };
}

/** 1 / f(x) for COT, CSC and SEC (see the module header). */
function reciprocal(x: number, f: (x: number) => number): number | CellError {
  if (Math.abs(x) >= RECIPROCAL_LIMIT) {
    return NUM_ERROR;
  }
  const value = f(x);
  return value === 0 ? DIV0_ERROR : 1 / value;
}

const ANGLE = "The angle in radians.";

export const TRIG_FUNCTIONS: FunctionSpec[] = [
  trig("SIN", "Returns the sine of an angle.", ANGLE, Math.sin),
  trig("COS", "Returns the cosine of an angle.", ANGLE, Math.cos),
  trig("TAN", "Returns the tangent of an angle.", ANGLE, Math.tan),
  trig("ASIN", "Returns the arcsine of a number, in radians.", "A number from -1 to 1.", (x) => (Math.abs(x) > 1 ? NUM_ERROR : Math.asin(x))),
  trig("ACOS", "Returns the arccosine of a number, in radians.", "A number from -1 to 1.", (x) => (Math.abs(x) > 1 ? NUM_ERROR : Math.acos(x))),
  trig("ATAN", "Returns the arctangent of a number, in radians.", "The tangent of the angle.", Math.atan),
  {
    name: "ATAN2",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the angle of the point (x, y) from the x-axis, in radians.",
    args: [arg("x_num", "The point's x coordinate."), arg("y_num", "The point's y coordinate.")],
    impl(args, ctx) {
      const x = numberArg(args, 0, ctx);
      if (isError(x)) {
        return x;
      }
      const y = numberArg(args, 1, ctx);
      if (isError(y)) {
        return y;
      }
      return x === 0 && y === 0 ? DIV0_ERROR : Math.atan2(y, x);
    },
  },
  trig("SINH", "Returns the hyperbolic sine of a number.", "Any real number.", Math.sinh),
  trig("COSH", "Returns the hyperbolic cosine of a number.", "Any real number.", Math.cosh),
  trig("TANH", "Returns the hyperbolic tangent of a number.", "Any real number.", Math.tanh),
  trig("ASINH", "Returns the inverse hyperbolic sine of a number.", "Any real number.", Math.asinh),
  trig("ACOSH", "Returns the inverse hyperbolic cosine of a number.", "A number of 1 or more.", (x) => (x < 1 ? NUM_ERROR : Math.acosh(x))),
  trig("ATANH", "Returns the inverse hyperbolic tangent of a number.", "A number strictly between -1 and 1.", (x) => (Math.abs(x) >= 1 ? NUM_ERROR : Math.atanh(x))),
  trig("COT", "Returns the cotangent of an angle.", "The angle in radians, below 2^27 in size.", (x) => reciprocal(x, Math.tan)),
  trig("CSC", "Returns the cosecant of an angle.", "The angle in radians, below 2^27 in size.", (x) => reciprocal(x, Math.sin)),
  trig("SEC", "Returns the secant of an angle.", "The angle in radians, below 2^27 in size.", (x) => reciprocal(x, Math.cos)),
  trig("ACOT", "Returns the arccotangent of a number, from 0 to pi.", "The cotangent of the angle.", (x) => Math.PI / 2 - Math.atan(x)),
];
