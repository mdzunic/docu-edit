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
//   included), decimals are rounded off — that text, half up — until it
//   fits. Otherwise — or when no rounding fits without losing every
//   significant digit — scientific: the mantissa rounded so the whole text
//   is at most 11 characters, upper-case E, a signed exponent of at least two
//   digits (123456789012 → "1.23457E+11", 0.0000000001234 → "1.234E-10").
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

/** A number's decimal significand at 15 significant digits (trailing zeros
 * removed) and the decimal exponent of its first digit — the text every
 * rounding below works on, as Excel's display does, rather than the binary
 * value (so 0.1234567895 rounds up to 0.12345679, not down). */
function significand(n: number): { negative: boolean; digits: string; exponent: number } {
  const [mantissa, exponent] = Math.abs(n).toExponential(14).split("e");
  return { negative: n < 0, digits: mantissa!.replace(".", "").replace(/0+$/, ""), exponent: Number(exponent) };
}

/** Rounds a digit string half-up to its first `keep` digits; `carried` when
 * the rounding gained a leading digit ("999" kept to 2 → "100"). */
function roundDigits(digits: string, keep: number): { digits: string; carried: boolean } {
  if (digits.length <= keep) {
    return { digits, carried: false };
  }
  const kept = digits.slice(0, keep).split("");
  if (digits.charCodeAt(keep) < 0x35 /* "5" */) {
    return { digits: kept.join(""), carried: false };
  }
  let i = kept.length - 1;
  while (i >= 0 && kept[i] === "9") {
    kept[i] = "0";
    i -= 1;
  }
  if (i < 0) {
    return { digits: `1${kept.join("")}`, carried: true };
  }
  kept[i] = String(Number(kept[i]) + 1);
  return { digits: kept.join(""), carried: false };
}

/** The fixed-notation General text, or null when it cannot fit in
 * GENERAL_WIDTH characters with a significant digit left. */
function fixedText(sig: ReturnType<typeof significand>): string | null {
  const sign = sig.negative ? "-" : "";
  let integer: string;
  let fraction: string;
  if (sig.exponent >= 0) {
    const padded = sig.digits.padEnd(sig.exponent + 1, "0");
    integer = padded.slice(0, sig.exponent + 1);
    fraction = padded.slice(sig.exponent + 1);
  } else {
    integer = "0";
    fraction = "0".repeat(-sig.exponent - 1) + sig.digits;
  }
  const join = (whole: string, decimals: string): string => `${sign}${whole}${decimals ? `.${decimals}` : ""}`;
  const plain = join(integer, fraction);
  if (plain.length <= GENERAL_WIDTH) {
    return plain;
  }
  // Round off decimals until it fits — down to none at all.
  for (let decimals = Math.max(0, GENERAL_WIDTH - sign.length - integer.length - 1); decimals >= 0; decimals -= 1) {
    const rounded = roundDigits(integer + fraction, integer.length + decimals);
    const integerLength = integer.length + (rounded.carried ? 1 : 0);
    const whole = rounded.digits.slice(0, integerLength);
    const decimalsText = rounded.digits.slice(integerLength).replace(/0+$/, "");
    if (/^0*$/.test(whole + decimalsText)) {
      return null;
    }
    const candidate = join(whole, decimalsText);
    if (candidate.length <= GENERAL_WIDTH) {
      return candidate;
    }
  }
  return null;
}

/** The scientific General text: as many mantissa decimals as fit. */
function scientificText(sig: ReturnType<typeof significand>): string {
  const sign = sig.negative ? "-" : "";
  let text = "";
  for (let decimals = GENERAL_WIDTH - 1; decimals >= 0; decimals -= 1) {
    const rounded = roundDigits(sig.digits, decimals + 1);
    const exponent = sig.exponent + (rounded.carried ? 1 : 0);
    const mantissa = trimFraction(`${rounded.digits[0]}.${rounded.digits.slice(1, decimals + 1)}`);
    text = `${sign}${mantissa}E${exponent < 0 ? "-" : "+"}${String(Math.abs(exponent)).padStart(2, "0")}`;
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
  const sig = significand(n);
  const magnitude = Math.abs(n);
  if (magnitude >= 1e-9 && magnitude < 1e11) {
    const fixed = fixedText(sig);
    if (fixed !== null) {
      return fixed;
    }
  }
  return scientificText(sig);
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
