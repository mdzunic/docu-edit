// Math & Trig, SPEC-127's library: SUMSQ GCD LCM FACT COMBIN PERMUT QUOTIENT
// EVEN ODD DEGREES RADIANS CEILING.MATH FLOOR.MATH CEILING.PRECISE
// FLOOR.PRECISE ISO.CEILING ROMAN ARABIC BASE DECIMAL MMULT. (The
// trigonometric set is trig.ts's, SUBTOTAL and AGGREGATE subtotal.ts's.)
//
// - GCD/LCM read numbers as SUM does, truncated; a negative operand, or one
//   of 2^53 or more, is #NUM!. FACT above 170 is #NUM!; COMBIN and PERMUT
//   truncate and need 0 ≤ number_chosen ≤ number.
// - EVEN/ODD round away from zero (ODD(0) is 1).
// - CEILING.MATH/FLOOR.MATH round to a multiple of |significance| (default
//   1): up and down for a positive number; for a negative one CEILING.MATH
//   goes toward zero and FLOOR.MATH away from it unless `mode` is non-zero,
//   which reverses that. CEILING.PRECISE (and ISO.CEILING, which shares its
//   implementation) and FLOOR.PRECISE go up and down whatever the signs.
//   Each rounds its quotient at 15 digits first, as SPEC-120's CEILING does.
// - ROMAN supports the classic form only (form 0 or TRUE); forms 1–4 (or
//   FALSE) are valid calls this library does not compute
//   (UnsupportedCaseError). ARABIC reads any well-formed or lenient numeral,
//   a leading "-" negating it.
// - BASE/DECIMAL take radixes 2–36; MMULT multiplies two all-number arrays
//   whose inner dimensions agree (a legacy array function).
import type { CellError } from "../../model";
import { makeArray, toArray } from "../areas";
import type { EvalContext } from "../evaluate";
import { UnsupportedCaseError, type ArgList, type FunctionSpec } from "../registry";
import { DIV0_ERROR, isError, NUM_ERROR, roundTo15, VALUE_ERROR } from "../values";
import { arg, eachNumber, integerArg, numberArg, numberOr, optional, repeating, scalarArg, textArg } from "./common";

/** 2^53: the integers GCD and LCM take stop below it. */
const SAFE = 2 ** 53;

/** Every number GCD/LCM read, truncated; #NUM! for a negative or huge one. */
function integers(args: ArgList, ctx: EvalContext): number[] | CellError {
  const out: number[] = [];
  let bad = false;
  const failure = eachNumber(args, ctx, (n) => {
    const whole = Math.trunc(n);
    if (whole < 0 || whole >= SAFE) {
      bad = true;
    }
    out.push(whole);
  });
  return failure ?? (bad ? NUM_ERROR : out);
}

function gcd(a: number, b: number): number {
  let x = a;
  let y = b;
  while (y !== 0) {
    [x, y] = [y, x % y];
  }
  return x;
}

/** x to a multiple of `step` (> 0) by `round`, through the 15-digit
 * quotient. */
function toMultiple(x: number, step: number, round: (q: number) => number): number {
  return roundTo15(round(roundTo15(x / step)) * step);
}

/** CEILING.MATH (`up`) or FLOOR.MATH. */
function mathRounding(name: "CEILING.MATH" | "FLOOR.MATH", up: boolean, summary: string): FunctionSpec {
  return {
    name,
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary,
    args: [
      arg("number", "The number to round."),
      optional("significance", "The multiple to round to (default 1); its sign is ignored."),
      optional("mode", "For a negative number, non-zero rounds the other way."),
    ],
    impl(args, ctx) {
      const x = numberArg(args, 0, ctx);
      if (isError(x)) {
        return x;
      }
      const significance = args.isMissing(1) ? 1 : numberOr(args, 1, ctx, 1);
      if (isError(significance)) {
        return significance;
      }
      const mode = numberOr(args, 2, ctx, 0);
      if (isError(mode)) {
        return mode;
      }
      const step = Math.abs(significance);
      if (step === 0 || x === 0) {
        return 0;
      }
      // Toward +∞ for CEILING, −∞ for FLOOR; a negative number with a mode
      // goes the other way (away from zero for CEILING, toward it for FLOOR).
      const ceil = x < 0 && mode !== 0 ? !up : up;
      return toMultiple(x, step, ceil ? Math.ceil : Math.floor);
    },
  };
}

/** CEILING.PRECISE / FLOOR.PRECISE: toward +∞ / −∞ whatever the signs. */
function precise(name: string, up: boolean, summary: string): FunctionSpec {
  return {
    name,
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary,
    args: [arg("number", "The number to round."), optional("significance", "The multiple to round to (default 1); its sign is ignored.")],
    impl: up ? ceilingPrecise : floorPrecise,
  };
}

function preciseRound(args: ArgList, ctx: EvalContext, up: boolean): number | CellError {
  const x = numberArg(args, 0, ctx);
  if (isError(x)) {
    return x;
  }
  const significance = args.isMissing(1) ? 1 : numberOr(args, 1, ctx, 1);
  if (isError(significance)) {
    return significance;
  }
  const step = Math.abs(significance);
  return step === 0 || x === 0 ? 0 : toMultiple(x, step, up ? Math.ceil : Math.floor);
}

function ceilingPrecise(args: ArgList, ctx: EvalContext): number | CellError {
  return preciseRound(args, ctx, true);
}

function floorPrecise(args: ArgList, ctx: EvalContext): number | CellError {
  return preciseRound(args, ctx, false);
}

const ROMAN_DIGITS: readonly [number, string][] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

const ROMAN_VALUES: Readonly<Record<string, number>> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

const DIGITS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

/** A unary function of one number. */
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

/** Two truncated integers (COMBIN, PERMUT), 0 ≤ chosen ≤ number. */
function chooseArguments(args: ArgList, ctx: EvalContext): [number, number] | CellError {
  const n = integerArg(args, 0, ctx);
  if (isError(n)) {
    return n;
  }
  const k = integerArg(args, 1, ctx);
  if (isError(k)) {
    return k;
  }
  return n < 0 || k < 0 || n < k ? NUM_ERROR : [n, k];
}

const CHOOSE_ARGS = [arg("number", "The number of items."), arg("number_chosen", "The number of items in each choice.")];

export const MATH_EXTRA_FUNCTIONS: FunctionSpec[] = [
  {
    name: "SUMSQ",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any", "any"],
    repeat: { from: 1, count: 1 },
    summary: "Adds the squares of its arguments.",
    args: [arg("number1", "A number, or a range or array of numbers."), repeating("number2", "More numbers, ranges or arrays.")],
    impl(args, ctx) {
      let total = 0;
      return eachNumber(args, ctx, (n) => {
        total += n * n;
      }) ?? total;
    },
  },
  {
    name: "GCD",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any", "any"],
    repeat: { from: 1, count: 1 },
    summary: "Returns the greatest common divisor of integers.",
    args: [arg("number1", "A non-negative integer, or a range or array of them."), repeating("number2", "More integers, ranges or arrays.")],
    impl(args, ctx) {
      const values = integers(args, ctx);
      return isError(values) ? values : values.reduce(gcd, 0);
    },
  },
  {
    name: "LCM",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any", "any"],
    repeat: { from: 1, count: 1 },
    summary: "Returns the least common multiple of integers.",
    args: [arg("number1", "A non-negative integer, or a range or array of them."), repeating("number2", "More integers, ranges or arrays.")],
    impl(args, ctx) {
      const values = integers(args, ctx);
      if (isError(values)) {
        return values;
      }
      let lcm = 1;
      for (const n of values) {
        if (n === 0) {
          return 0;
        }
        lcm = (lcm / gcd(lcm, n)) * n;
        if (lcm >= SAFE) {
          return NUM_ERROR;
        }
      }
      return values.length === 0 ? 0 : lcm;
    },
  },
  unary("FACT", "Returns the factorial of a number.", "A number from 0 to 170, truncated.", (x) => {
    const n = Math.trunc(x);
    if (n < 0 || n > 170) {
      return NUM_ERROR;
    }
    let product = 1;
    for (let i = 2; i <= n; i += 1) {
      product *= i;
    }
    return product;
  }),
  {
    name: "COMBIN",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the number of combinations of items, order ignored.",
    args: CHOOSE_ARGS,
    impl(args, ctx) {
      const pair = chooseArguments(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const [n, chosen] = pair;
      const k = Math.min(chosen, n - chosen);
      let result = 1;
      for (let i = 1; i <= k; i += 1) {
        result = (result * (n - k + i)) / i;
      }
      // The running product is exact while it stays below 2^53.
      return result < SAFE ? Math.round(result) : result;
    },
  },
  {
    name: "PERMUT",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the number of permutations of items, order counted.",
    args: CHOOSE_ARGS,
    impl(args, ctx) {
      const pair = chooseArguments(args, ctx);
      if (isError(pair)) {
        return pair;
      }
      const [n, k] = pair;
      let result = 1;
      for (let i = 0; i < k; i += 1) {
        result *= n - i;
      }
      return result;
    },
  },
  {
    name: "QUOTIENT",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the integer part of a division.",
    args: [arg("numerator", "The number to divide."), arg("denominator", "The number to divide by.")],
    impl(args, ctx) {
      const n = numberArg(args, 0, ctx);
      if (isError(n)) {
        return n;
      }
      const d = numberArg(args, 1, ctx);
      if (isError(d)) {
        return d;
      }
      return d === 0 ? DIV0_ERROR : Math.trunc(roundTo15(n / d));
    },
  },
  unary("EVEN", "Rounds a number away from zero to the nearest even integer.", "The number to round.", (x) => {
    const magnitude = Math.ceil(roundTo15(Math.abs(x) / 2)) * 2;
    return x < 0 ? -magnitude : magnitude;
  }),
  unary("ODD", "Rounds a number away from zero to the nearest odd integer.", "The number to round.", (x) => {
    const magnitude = Math.ceil(roundTo15((Math.abs(x) + 1) / 2)) * 2 - 1;
    return x < 0 ? -magnitude : magnitude;
  }),
  unary("DEGREES", "Converts an angle in radians to degrees.", "The angle in radians.", (x) => (x * 180) / Math.PI),
  unary("RADIANS", "Converts an angle in degrees to radians.", "The angle in degrees.", (x) => (x * Math.PI) / 180),
  mathRounding("CEILING.MATH", true, "Rounds a number up to a multiple, with a mode for negative numbers."),
  mathRounding("FLOOR.MATH", false, "Rounds a number down to a multiple, with a mode for negative numbers."),
  precise("CEILING.PRECISE", true, "Rounds a number up to a multiple, whatever the signs."),
  precise("FLOOR.PRECISE", false, "Rounds a number down to a multiple, whatever the signs."),
  precise("ISO.CEILING", true, "Rounds a number up to a multiple, whatever the signs (ISO 8601)."),
  {
    name: "ROMAN",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Converts a number to a Roman numeral as text.",
    args: [arg("number", "A number from 0 to 3999."), optional("form", "0 or TRUE (default): the classic form; 1–4 simplify (not computed here).")],
    impl(args, ctx) {
      const n = integerArg(args, 0, ctx);
      if (isError(n)) {
        return n;
      }
      if (args.length > 1 && !args.isMissing(1)) {
        const form = scalarArg(args, 1);
        const kind = typeof form === "boolean" ? (form ? 0 : 4) : numberArg(args, 1, ctx);
        if (isError(kind)) {
          return kind;
        }
        const whole = Math.trunc(kind);
        if (whole < 0 || whole > 4) {
          return VALUE_ERROR;
        }
        if (whole !== 0) {
          throw new UnsupportedCaseError("Only ROMAN's classic form is computed");
        }
      }
      if (n < 0 || n > 3999) {
        return VALUE_ERROR;
      }
      let left = n;
      let out = "";
      for (const [value, letters] of ROMAN_DIGITS) {
        while (left >= value) {
          out += letters;
          left -= value;
        }
      }
      return out;
    },
  },
  {
    name: "ARABIC",
    category: "Math & Trig",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Converts a Roman numeral to a number.",
    args: [arg("text", "The Roman numeral, such as \"MMXXVI\".")],
    impl(args) {
      const raw = textArg(args, 0);
      if (isError(raw)) {
        return raw;
      }
      let text = raw.trim().toUpperCase();
      if (text.length > 255) {
        return VALUE_ERROR;
      }
      const negative = text.startsWith("-");
      if (negative) {
        text = text.slice(1);
      }
      let total = 0;
      for (let i = 0; i < text.length; i += 1) {
        const value = ROMAN_VALUES[text[i]!];
        if (value === undefined) {
          return VALUE_ERROR;
        }
        const next = i + 1 < text.length ? ROMAN_VALUES[text[i + 1]!] : undefined;
        total += next !== undefined && next > value ? -value : value;
      }
      return negative ? -total : total;
    },
  },
  {
    name: "BASE",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Converts a number to text in a radix from 2 to 36.",
    args: [
      arg("number", "A whole number from 0 to 2^53 − 1."),
      arg("radix", "The base, from 2 to 36."),
      optional("min_length", "The least number of digits, padded with zeros."),
    ],
    impl(args, ctx) {
      const n = integerArg(args, 0, ctx);
      if (isError(n)) {
        return n;
      }
      const radix = integerArg(args, 1, ctx);
      if (isError(radix)) {
        return radix;
      }
      const minLength = numberOr(args, 2, ctx, 0);
      if (isError(minLength)) {
        return minLength;
      }
      const pad = Math.trunc(minLength);
      if (n < 0 || n >= SAFE || radix < 2 || radix > 36 || pad < 0 || pad > 255) {
        return NUM_ERROR;
      }
      return n.toString(radix).toUpperCase().padStart(pad, "0");
    },
  },
  {
    name: "DECIMAL",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Converts text in a radix from 2 to 36 to a number.",
    args: [arg("text", "The digits, at most 255 of them."), arg("radix", "The base, from 2 to 36.")],
    impl(args, ctx) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      const radix = integerArg(args, 1, ctx);
      if (isError(radix)) {
        return radix;
      }
      if (radix < 2 || radix > 36 || text.length > 255) {
        return NUM_ERROR;
      }
      let total = 0;
      for (const ch of text.toUpperCase()) {
        const digit = DIGITS.indexOf(ch);
        if (digit < 0 || digit >= radix) {
          return NUM_ERROR;
        }
        total = total * radix + digit;
      }
      return total >= SAFE ? NUM_ERROR : total;
    },
  },
  {
    name: "MMULT",
    category: "Math & Trig",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["array", "array"],
    summary: "Returns the matrix product of two arrays.",
    args: [arg("array1", "The left array, as many columns as array2 has rows."), arg("array2", "The right array.")],
    impl(args, ctx) {
      const a = toArray(ctx, args.get(0));
      if (isError(a)) {
        return a;
      }
      const b = toArray(ctx, args.get(1));
      if (isError(b)) {
        return b;
      }
      if (a.cols !== b.rows) {
        return VALUE_ERROR;
      }
      for (const element of [...a.data, ...b.data]) {
        if (isError(element)) {
          return element;
        }
        if (typeof element !== "number") {
          return VALUE_ERROR;
        }
      }
      const data: number[] = [];
      for (let r = 0; r < a.rows; r += 1) {
        for (let c = 0; c < b.cols; c += 1) {
          let total = 0;
          for (let k = 0; k < a.cols; k += 1) {
            total += (a.data[r * a.cols + k] as number) * (b.data[k * b.cols + c] as number);
          }
          data.push(total);
        }
      }
      return makeArray(a.rows, b.cols, data);
    },
  },
];