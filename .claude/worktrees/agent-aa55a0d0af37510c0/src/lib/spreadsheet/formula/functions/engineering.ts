// Engineering (SPEC-127 "Function pins"): the twelve base conversions
// BIN2DEC BIN2HEX BIN2OCT DEC2BIN DEC2HEX DEC2OCT HEX2BIN HEX2DEC HEX2OCT
// OCT2BIN OCT2DEC OCT2HEX, and DELTA and GESTEP (ERF and its kin live in
// erf.ts).
//
// - A binary, octal or hexadecimal `number` is at most 10 characters in
//   10-character two's complement, as Microsoft documents: a ten-digit value
//   with its top bit set is negative (binary 1111111111 is −1). Text is read
//   as written, letters in either case; a number by its decimal digits
//   (BIN2DEC(1100100)); an empty cell or "" is 0; a logical is #VALUE!; more
//   than 10 characters or a digit outside the base (a sign, a point) #NUM!.
// - A decimal `number` is coerced as usual and truncated toward zero.
// - A result is upper case: a negative one as its ten-character two's
//   complement, `places` ignored; a non-negative one padded with zeros to
//   `places`. A result outside the target's range (DEC2BIN −512…511, DEC2OCT
//   −2^29…2^29−1, DEC2HEX −2^39…2^39−1, the same for X2BIN and X2OCT) is
//   #NUM!. `places` is truncated; below 1, above 10 or fewer than the digits
//   needed is #NUM!, non-numeric text #VALUE!; an omitted one means as few
//   characters as needed.
// - DELTA and GESTEP compare as the "=" and ">=" operators do (at 15
//   significant digits), number2 and step defaulting to 0.
import type { CellError } from "../../model";
import type { ArgList, FunctionSpec } from "../registry";
import type { EvalContext } from "../evaluate";
import { isError, NUM_ERROR, numberToText, roundTo15, VALUE_ERROR } from "../values";
import { arg, numberArg, numberOr, optional, scalarArg } from "./common";

type Radix = 2 | 8 | 10 | 16;

const BASE_NAMES: Readonly<Record<Radix, string>> = { 2: "BIN", 8: "OCT", 10: "DEC", 16: "HEX" };
const BASE_WORDS: Readonly<Record<Radix, string>> = { 2: "binary", 8: "octal", 10: "decimal", 16: "hexadecimal" };

/** The digits a binary, octal or hexadecimal `number` spells (see the module
 * header). */
function digitsArg(args: ArgList): string | CellError {
  const value = scalarArg(args, 0);
  if (isError(value)) {
    return value;
  }
  if (typeof value === "boolean") {
    return VALUE_ERROR;
  }
  if (typeof value === "number") {
    return numberToText(value);
  }
  return typeof value === "string" ? value : "";
}

/** The value of up to ten digits in `radix`, a ten-digit one with its top
 * bit set negative; #NUM! for more than ten characters or a stray digit. */
function parseSigned(text: string, radix: 2 | 8 | 16): number | CellError {
  if (text.length > 10) {
    return NUM_ERROR;
  }
  let value = 0;
  for (const ch of text) {
    const digit = Number.parseInt(ch, radix);
    if (Number.isNaN(digit)) {
      return NUM_ERROR;
    }
    value = value * radix + digit;
  }
  const full = radix ** 10;
  return value >= full / 2 ? value - full : value;
}

/** `places` (argument 1): undefined when omitted, else truncated and 1…10. */
function placesArg(args: ArgList, ctx: EvalContext): number | undefined | CellError {
  if (args.length < 2) {
    return undefined;
  }
  const places = numberArg(args, 1, ctx);
  if (isError(places)) {
    return places;
  }
  const whole = Math.trunc(places);
  return whole < 1 || whole > 10 ? NUM_ERROR : whole;
}

/** `value` in ten-character two's complement of `radix` (see the module
 * header); #NUM! outside its range or for too few places. */
function formatSigned(value: number, radix: 2 | 8 | 16, places: number | undefined): string | CellError {
  const full = radix ** 10;
  if (value < -full / 2 || value >= full / 2) {
    return NUM_ERROR;
  }
  if (value < 0) {
    return (full + value).toString(radix).toUpperCase();
  }
  const digits = value.toString(radix).toUpperCase();
  if (places === undefined) {
    return digits;
  }
  return digits.length > places ? NUM_ERROR : digits.padStart(places, "0");
}

/** One base conversion, from → to. */
function conversion(from: Radix, to: Radix): FunctionSpec {
  const numberMeta =
    from === 10
      ? arg("number", "The decimal integer to convert.")
      : arg("number", `Up to 10 ${BASE_WORDS[from]} digits; with all 10, a set top bit makes it negative.`);
  const toDecimal = to === 10;
  return {
    name: `${BASE_NAMES[from]}2${BASE_NAMES[to]}`,
    category: "Engineering",
    minArgs: 1,
    maxArgs: toDecimal ? 1 : 2,
    argKinds: toDecimal ? ["value"] : ["value", "value"],
    summary: `Converts a ${BASE_WORDS[from]} number to ${BASE_WORDS[to]}.`,
    args: toDecimal ? [numberMeta] : [numberMeta, optional("places", "The minimum number of characters, padded with leading zeros.")],
    impl(args, ctx) {
      let value: number | CellError;
      if (from === 10) {
        const n = numberArg(args, 0, ctx);
        value = isError(n) ? n : Math.trunc(n);
      } else {
        const digits = digitsArg(args);
        value = isError(digits) ? digits : parseSigned(digits, from);
      }
      if (to === 10) {
        return value;
      }
      const places = placesArg(args, ctx);
      if (isError(places)) {
        return places;
      }
      return isError(value) ? value : formatSigned(value, to, places);
    },
  };
}

/** DELTA or GESTEP: 1 when the comparison holds, else 0. */
function comparison(name: "DELTA" | "GESTEP", summary: string, holds: (a: number, b: number) => boolean): FunctionSpec {
  return {
    name,
    category: "Engineering",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary,
    args:
      name === "DELTA"
        ? [arg("number1", "The first number."), optional("number2", "The second number (default 0).")]
        : [arg("number", "The number to test."), optional("step", "The threshold (default 0).")],
    impl(args, ctx) {
      const a = numberArg(args, 0, ctx);
      if (isError(a)) {
        return a;
      }
      const b = numberOr(args, 1, ctx, 0);
      if (isError(b)) {
        return b;
      }
      return holds(roundTo15(a), roundTo15(b)) ? 1 : 0;
    },
  };
}

const RADIXES: readonly Radix[] = [2, 10, 16, 8];

export const ENGINEERING_FUNCTIONS: FunctionSpec[] = [
  ...RADIXES.flatMap((from) => RADIXES.filter((to) => to !== from).map((to) => conversion(from, to))),
  comparison("DELTA", "Returns 1 when two numbers are equal, else 0.", (a, b) => a === b),
  comparison("GESTEP", "Returns 1 when a number is at least a threshold, else 0.", (a, b) => a >= b),
];
