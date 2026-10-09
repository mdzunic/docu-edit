// Math & Trig (SPEC-120 "The core functions"): SUM SUMIF SUMIFS SUMPRODUCT
// PRODUCT ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD ABS SIGN SQRT POWER EXP LN LOG
// LOG10 PI CEILING FLOOR MROUND RAND RANDBETWEEN.
//
// - Rounding works on the 15-significant-digit decimal form of the input,
//   never the binary value: ROUND is half away from zero, so ROUND(1.005,2)
//   is 1.01; CEILING/FLOOR/MROUND round their quotient to 15 digits before
//   flooring it, so FLOOR(0.3,0.1) is 0.3. Their sign rules are Excel's
//   documented ones (CEILING(-2.5,2) = −2, CEILING(-2.5,-2) = −4,
//   FLOOR(2.5,-2) = #NUM!, MROUND(5,-2) = #NUM!); CEILING(x,0) is 0 and
//   FLOOR(x,0) #DIV/0!.
// - MOD(n,d) = n − d·INT(n/d) (the divisor's sign); INT floors, TRUNC
//   truncates. PRODUCT of no numbers is 0.
// - SUMIF resizes its sum range to the criteria range's shape from its
//   top-left cell; SUMIFS needs every range the same shape. SUMPRODUCT needs
//   equal dimensions and reads a non-number as 0.
// - RAND and RANDBETWEEN read ctx.random() (volatile):
//   RANDBETWEEN(a,b) = ceil(a) + floor(random()·(floor(b)−ceil(a)+1)), #NUM!
//   when ceil(a) > floor(b).
import type { CellError } from "../../model";
import { significand } from "../../numberFormat/decimal";
import { toArray } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { DIV0_ERROR, isError, NUM_ERROR, power, roundTo15, VALUE_ERROR, type ArrayValue } from "../values";
import { arg, criteriaPairs, eachNumber, forEachMatch, numberArg, numberOr, optional, repeating, resizedArea, scalarArg, singleArea } from "./common";
import { parseCriterion } from "./criteria";

/** A digit string incremented by one in its last place ("199" → "200"),
 * `carried` when it gained a digit ("99" → "100"). */
function increment(digits: string): { digits: string; carried: boolean } {
  const out = digits.split("");
  let i = out.length - 1;
  while (i >= 0 && out[i] === "9") {
    out[i] = "0";
    i -= 1;
  }
  if (i < 0) {
    return { digits: `1${out.join("")}`, carried: true };
  }
  out[i] = String(Number(out[i]) + 1);
  return { digits: out.join(""), carried: false };
}

/** x rounded to `places` decimals (negative: tens, hundreds…) on its
 * 15-significant-digit decimal form: "half" away from zero, "up" away from
 * zero, "down" toward zero. */
export function roundDecimal(x: number, places: number, mode: "half" | "up" | "down"): number {
  if (x === 0 || !Number.isFinite(x)) {
    return x === 0 ? 0 : x;
  }
  const sig = significand(x);
  const sign = sig.negative ? -1 : 1;
  const keep = sig.exponent + 1 + places;
  if (keep >= sig.digits.length) {
    return roundTo15(x);
  }
  if (keep <= 0) {
    const up = mode === "up" || (mode === "half" && keep === 0 && sig.digits.charCodeAt(0) >= 0x35);
    return up ? sign * 10 ** -places : 0;
  }
  const kept = sig.digits.slice(0, keep);
  const up = mode === "up" || (mode === "half" && sig.digits.charCodeAt(keep) >= 0x35);
  const rounded = up ? increment(kept) : { digits: kept, carried: false };
  return sign * Number(`0.${rounded.digits}e${sig.exponent + 1 + (rounded.carried ? 1 : 0)}`);
}

/** n and d as numbers, or the first error. */
function twoNumbers(args: ArgList, ctx: EvalContext): [number, number] | CellError {
  const n = numberArg(args, 0, ctx);
  if (isError(n)) {
    return n;
  }
  const d = numberArg(args, 1, ctx);
  return isError(d) ? d : [n, d];
}

function rounding(name: string, mode: "half" | "up" | "down", summary: string): FunctionSpec {
  return {
    name,
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary,
    args: [arg("number", "The number to round."), arg("num_digits", "The decimal places to keep; negative rounds left of the point.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      return isError(pair) ? pair : roundDecimal(pair[0], Math.trunc(pair[1]), mode);
    },
  };
}

/** One number in, one number out — `f` may answer an error. */
function unary(name: string, summary: string, description: string, f: (x: number) => number | CellError): FunctionSpec {
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

/** x to a multiple of `significance` through its 15-digit quotient. */
function toMultiple(x: number, significance: number, round: (q: number) => number): number {
  return roundTo15(round(roundTo15(x / significance)) * significance);
}

/** Rounds half away from zero. */
function halfAwayFromZero(q: number): number {
  return Math.sign(q) * Math.floor(Math.abs(q) + 0.5);
}

export const MATH_FUNCTIONS: FunctionSpec[] = [
  {
    name: "SUM",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: "Adds its arguments.",
    args: [repeating("number", "A number, or a range or array whose numbers are added.")],
    impl(args, ctx) {
      let total = 0;
      return eachNumber(args, ctx, (n) => {
        total += n;
      }) ?? total;
    },
  },
  {
    name: "SUMIF",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["range", "value", "range"],
    summary: "Adds the cells that meet a criterion.",
    args: [
      arg("range", "The cells the criterion is tested against."),
      arg("criteria", "The criterion: a value, or text such as \">5\" or \"b*\"."),
      optional("sum_range", "The cells to add, when not the range itself."),
    ],
    impl(args, ctx) {
      const range = singleArea(args.get(0));
      if (isError(range)) {
        return range;
      }
      let values = range;
      if (args.length > 2 && !args.isMissing(2)) {
        const sum = singleArea(args.get(2));
        if (isError(sum)) {
          return sum;
        }
        values = resizedArea(sum, range);
      }
      const pairs = [{ area: range, criterion: parseCriterion(scalarArg(args, 1), ctx.settings) }];
      let total = 0;
      let failure: CellError | undefined;
      forEachMatch(ctx, values, pairs, (value) => {
        if (isError(value)) {
          failure = value;
          return false;
        }
        if (typeof value === "number") {
          total += value;
        }
        return true;
      });
      return failure ?? total;
    },
  },
  {
    name: "SUMIFS",
    category: "Math & Trig",
    minArgs: 3,
    maxArgs: 255,
    argKinds: ["range", "range", "value"],
    repeat: { from: 1, count: 2 },
    summary: "Adds the cells that meet every criterion.",
    args: [
      arg("sum_range", "The cells to add."),
      repeating("criteria_range", "The cells a criterion is tested against."),
      repeating("criteria", "The criterion for that range."),
    ],
    impl(args, ctx) {
      const sum = singleArea(args.get(0));
      if (isError(sum)) {
        return sum;
      }
      const pairs = criteriaPairs(args, 1, ctx, sum);
      if (isError(pairs)) {
        return pairs;
      }
      let total = 0;
      let failure: CellError | undefined;
      forEachMatch(ctx, sum, pairs, (value) => {
        if (isError(value)) {
          failure = value;
          return false;
        }
        if (typeof value === "number") {
          total += value;
        }
        return true;
      });
      return failure ?? total;
    },
  },
  {
    name: "SUMPRODUCT",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["array"],
    repeat: { from: 0, count: 1 },
    summary: "Multiplies corresponding elements of arrays and adds the products.",
    args: [repeating("array", "An array or range, all of the same dimensions.")],
    impl(args, ctx) {
      const arrays: ArrayValue[] = [];
      for (let i = 0; i < args.length; i += 1) {
        const array = toArray(ctx, args.get(i));
        if (isError(array)) {
          return array;
        }
        if (arrays.length > 0 && (array.rows !== arrays[0]!.rows || array.cols !== arrays[0]!.cols)) {
          return VALUE_ERROR;
        }
        arrays.push(array);
      }
      for (const array of arrays) {
        const failure = array.data.find((element) => isError(element));
        if (failure !== undefined) {
          return failure;
        }
      }
      let total = 0;
      const size = arrays[0]!.data.length;
      for (let i = 0; i < size; i += 1) {
        let product = 1;
        for (const array of arrays) {
          const element = array.data[i];
          product *= typeof element === "number" ? element : 0;
        }
        total += product;
      }
      return total;
    },
  },
  {
    name: "PRODUCT",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any"],
    repeat: { from: 0, count: 1 },
    summary: "Multiplies its arguments.",
    args: [repeating("number", "A number, or a range or array whose numbers are multiplied.")],
    impl(args, ctx) {
      let product = 1;
      let count = 0;
      const failure = eachNumber(args, ctx, (n) => {
        product *= n;
        count += 1;
      });
      return failure ?? (count === 0 ? 0 : product);
    },
  },
  rounding("ROUND", "half", "Rounds a number to a number of digits, halves away from zero."),
  rounding("ROUNDUP", "up", "Rounds a number away from zero to a number of digits."),
  rounding("ROUNDDOWN", "down", "Rounds a number toward zero to a number of digits."),
  unary("INT", "Rounds a number down to the nearest integer.", "The number to round down.", (x) => Math.floor(x)),
  {
    name: "TRUNC",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Truncates a number to a number of digits.",
    args: [arg("number", "The number to truncate."), optional("num_digits", "The decimal places to keep (default 0).")],
    impl(args, ctx) {
      const x = numberArg(args, 0, ctx);
      if (isError(x)) {
        return x;
      }
      const places = numberOr(args, 1, ctx, 0);
      return isError(places) ? places : roundDecimal(x, Math.trunc(places), "down");
    },
  },
  {
    name: "MOD",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the remainder of a division, with the divisor's sign.",
    args: [arg("number", "The number to divide."), arg("divisor", "The number to divide by.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const [n, d] = pair;
      return d === 0 ? DIV0_ERROR : n - d * Math.floor(n / d);
    },
  },
  unary("ABS", "Returns the absolute value of a number.", "The number.", (x) => Math.abs(x)),
  unary("SIGN", "Returns 1, 0 or -1: the sign of a number.", "The number.", (x) => Math.sign(x)),
  unary("SQRT", "Returns a positive square root.", "The number; negative is #NUM!.", (x) => (x < 0 ? NUM_ERROR : Math.sqrt(x))),
  {
    name: "POWER",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns a number raised to a power.",
    args: [arg("number", "The base."), arg("power", "The exponent.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      return isError(pair) ? pair : power(pair[0], pair[1]);
    },
  },
  unary("EXP", "Returns e raised to a power.", "The exponent.", (x) => Math.exp(x)),
  unary("LN", "Returns the natural logarithm of a number.", "A positive number.", (x) => (x <= 0 ? NUM_ERROR : Math.log(x))),
  {
    name: "LOG",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the logarithm of a number to a base.",
    args: [arg("number", "A positive number."), optional("base", "The base (default 10).")],
    impl(args, ctx) {
      const x = numberArg(args, 0, ctx);
      if (isError(x)) {
        return x;
      }
      const base = numberOr(args, 1, ctx, 10);
      if (isError(base)) {
        return base;
      }
      if (x <= 0 || base <= 0) {
        return NUM_ERROR;
      }
      if (base === 1) {
        return DIV0_ERROR;
      }
      return base === 10 ? Math.log10(x) : base === 2 ? Math.log2(x) : Math.log(x) / Math.log(base);
    },
  },
  unary("LOG10", "Returns the base-10 logarithm of a number.", "A positive number.", (x) => (x <= 0 ? NUM_ERROR : Math.log10(x))),
  {
    name: "PI",
    category: "Math & Trig",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    summary: "Returns the value of pi.",
    args: [],
    impl() {
      return Math.PI;
    },
  },
  {
    name: "CEILING",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Rounds a number up to the nearest multiple of a significance.",
    args: [arg("number", "The number to round."), arg("significance", "The multiple to round to.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const [x, significance] = pair;
      if (significance === 0 || x === 0) {
        return 0;
      }
      if (x > 0 && significance < 0) {
        return NUM_ERROR;
      }
      return toMultiple(x, significance, Math.ceil);
    },
  },
  {
    name: "FLOOR",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Rounds a number down to the nearest multiple of a significance.",
    args: [arg("number", "The number to round."), arg("significance", "The multiple to round to.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const [x, significance] = pair;
      if (significance === 0) {
        return DIV0_ERROR;
      }
      if (x > 0 && significance < 0) {
        return NUM_ERROR;
      }
      return x === 0 ? 0 : toMultiple(x, significance, Math.floor);
    },
  },
  {
    name: "MROUND",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Rounds a number to the nearest multiple, halves away from zero.",
    args: [arg("number", "The number to round."), arg("multiple", "The multiple to round to.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const [x, multiple] = pair;
      if (multiple === 0 || x === 0) {
        return 0;
      }
      if (x * multiple < 0) {
        return NUM_ERROR;
      }
      return toMultiple(x, multiple, halfAwayFromZero);
    },
  },
  {
    name: "RAND",
    category: "Math & Trig",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    volatile: true,
    summary: "Returns a random number from 0 up to 1.",
    args: [],
    impl(_args, ctx) {
      return ctx.random();
    },
  },
  {
    name: "RANDBETWEEN",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    volatile: true,
    summary: "Returns a random integer between two numbers.",
    args: [arg("bottom", "The smallest integer it may return."), arg("top", "The largest integer it may return.")],
    impl(args, ctx) {
      const pair = twoNumbers(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const low = Math.ceil(pair[0]);
      const high = Math.floor(pair[1]);
      if (low > high) {
        return NUM_ERROR;
      }
      return low + Math.floor(ctx.random() * (high - low + 1));
    },
  },
];
