// The "General" display of a cell value at full column width (SPEC-116),
// matching Excel's General format: what the grid and the library preview
// show until number formats arrive. SPEC-118 subsumes this with width-aware
// shrinking and the locale's decimal separator; at full width in en-US its
// output equals this function's, so nothing produced here ever changes.
//
// - null → ""; a string → itself; true/false → "TRUE"/"FALSE"; an error →
//   its code.
// - A number: 0 → "0". If 1e-9 ≤ |n| < 1e11, the fixed-notation text of n
//   rounded to at most 15 significant digits, trailing zeros and a trailing
//   point removed; if that is longer than 11 characters (sign and point
//   included), decimals are rounded off until it fits. Otherwise — or when
//   no rounding fits without losing every significant digit — scientific:
//   the mantissa rounded so the whole text is at most 11 characters,
//   upper-case E, a signed exponent of at least two digits
//   (123456789012 → "1.23457E+11", 0.0000000001234 → "1.234E-10").
import type { CellValue } from "./model";

/** The General format's width at a standard column, in characters. */
const GENERAL_WIDTH = 11;

/** "1.2300" → "1.23", "4.000" → "4", "120" → "120". */
function trimFraction(text: string): string {
  if (!text.includes(".")) {
    return text;
  }
  return text.replace(/0+$/, "").replace(/\.$/, "");
}

/** A number as plain decimal text (no exponent) at up to 15 significant
 * digits, trailing zeros removed. */
function plainDecimal(n: number): string {
  const rounded = Number(n.toPrecision(15));
  const text = String(rounded);
  if (!/e/i.test(text)) {
    return text;
  }
  const exponent = Math.floor(Math.log10(Math.abs(rounded)));
  return trimFraction(rounded.toFixed(Math.min(100, Math.max(0, 14 - exponent))));
}

/** The fixed-notation General text, or null when it cannot fit in
 * GENERAL_WIDTH characters with a significant digit left. */
function fixedText(n: number): string | null {
  const text = plainDecimal(n);
  if (text.length <= GENERAL_WIDTH) {
    return text;
  }
  const negative = text.startsWith("-");
  const magnitude = negative ? text.slice(1) : text;
  const integerDigits = magnitude.split(".")[0]!.length;
  for (let decimals = GENERAL_WIDTH - (negative ? 1 : 0) - integerDigits - 1; decimals >= 0; decimals -= 1) {
    const rounded = trimFraction(Math.abs(n).toFixed(decimals));
    if (Number(rounded) === 0) {
      return null;
    }
    const candidate = negative ? `-${rounded}` : rounded;
    if (candidate.length <= GENERAL_WIDTH) {
      return candidate;
    }
  }
  return null;
}

/** The scientific General text: as many mantissa decimals as fit. */
function scientificText(n: number): string {
  const sign = n < 0 ? "-" : "";
  const magnitude = Math.abs(n);
  let text = "";
  for (let decimals = GENERAL_WIDTH - 1; decimals >= 0; decimals -= 1) {
    const [mantissa, exponentText] = magnitude.toExponential(decimals).split("e");
    const exponent = Number(exponentText);
    text = `${sign}${trimFraction(mantissa!)}E${exponent < 0 ? "-" : "+"}${String(Math.abs(exponent)).padStart(2, "0")}`;
    if (text.length <= GENERAL_WIDTH) {
      return text;
    }
  }
  return text;
}

function generalNumber(n: number): string {
  if (n === 0) {
    return "0";
  }
  if (!Number.isFinite(n)) {
    return "#NUM!";
  }
  const magnitude = Math.abs(n);
  if (magnitude >= 1e-9 && magnitude < 1e11) {
    const fixed = fixedText(n);
    if (fixed !== null) {
      return fixed;
    }
  }
  return scientificText(n);
}

/** A cell value as Excel's General format shows it in a full-width cell. */
export function generalText(value: CellValue): string {
  if (value === null) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "number") {
    return generalNumber(value);
  }
  return value.error;
}
