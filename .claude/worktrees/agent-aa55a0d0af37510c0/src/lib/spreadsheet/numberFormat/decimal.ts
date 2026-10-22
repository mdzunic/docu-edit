// Decimal-text arithmetic for the format engine (SPEC-118 "Rendering
// numbers" steps 1–2, "General", "editText"). Every rounding here works on a
// number's 15-significant-digit decimal text — what Excel's display does —
// never on the binary value: 0.1 + 0.2 renders as 0.3, 2.675 rounds to 2.68
// and 1.005 to 1.01 (half away from zero on the decimal text).
//
// Also SPEC-116's General text at full width (the 11-character rule), moved
// here from general.ts unchanged so formatGeneral and generalText share it:
// 0 → "0"; if 1e-9 ≤ |n| < 1e11, the fixed-notation text of n rounded to at
// most 15 significant digits, trailing zeros and a trailing point removed,
// decimals rounded off until it is at most 11 characters (sign and point
// included); otherwise — or when no rounding keeps a significant digit —
// scientific with upper-case E and a signed exponent of at least two digits.

/** The General format's width at a standard column, in characters. */
export const GENERAL_WIDTH = 11;

/** A number's decimal significand at 15 significant digits (trailing zeros
 * removed; "" for zero) and the decimal exponent of its first digit. */
export function significand(n: number): { negative: boolean; digits: string; exponent: number } {
  const [mantissa, exponent] = Math.abs(n).toExponential(14).split("e");
  return { negative: n < 0, digits: mantissa!.replace(".", "").replace(/0+$/, ""), exponent: Number(exponent) };
}

/** Rounds a digit string half-up to its first `keep` digits; `carried` when
 * the rounding gained a leading digit ("999" kept to 2 → "100"). */
export function roundDigits(digits: string, keep: number): { digits: string; carried: boolean } {
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

/** "1.2300" → "1.23", "4.000" → "4", "120" → "120". */
export function trimFraction(text: string): string {
  if (!text.includes(".")) {
    return text;
  }
  return text.replace(/0+$/, "").replace(/\.$/, "");
}

/** |x| at 15 significant digits as its integer digits (no leading zeros —
 * "" when the integer part is 0) and fraction digits (no trailing zeros). */
export function decimalParts(x: number): { int: string; frac: string } {
  const sig = significand(x);
  if (sig.digits === "") {
    return { int: "", frac: "" };
  }
  if (sig.exponent >= 0) {
    const padded = sig.digits.padEnd(sig.exponent + 1, "0");
    return { int: padded.slice(0, sig.exponent + 1), frac: padded.slice(sig.exponent + 1) };
  }
  return { int: "", frac: "0".repeat(-sig.exponent - 1) + sig.digits };
}

/** |x| rounded half away from zero to `places` decimals: its integer digits
 * ("" for 0) and exactly `places` fraction digits. */
export function roundToPlaces(x: number, places: number): { int: string; frac: string } {
  const { int, frac } = decimalParts(x);
  if (frac.length <= places) {
    return { int, frac: frac.padEnd(places, "0") };
  }
  const rounded = roundDigits(int + frac, int.length + places);
  const intLength = int.length + (rounded.carried ? 1 : 0);
  return { int: rounded.digits.slice(0, intLength).replace(/^0+/, ""), frac: rounded.digits.slice(intLength) };
}

/** The scientific text of a significand with `decimals` mantissa decimals
 * (trailing zeros dropped): "1.23E+11", "1E-10". */
export function scientificCandidate(sig: ReturnType<typeof significand>, decimals: number, decimal = "."): string {
  const sign = sig.negative ? "-" : "";
  const rounded = roundDigits(sig.digits, decimals + 1);
  const exponent = sig.exponent + (rounded.carried ? 1 : 0);
  const mantissa = trimFraction(`${rounded.digits[0]}.${rounded.digits.slice(1, decimals + 1)}`).replace(".", decimal);
  return `${sign}${mantissa}E${exponent < 0 ? "-" : "+"}${String(Math.abs(exponent)).padStart(2, "0")}`;
}

/** The fixed-notation text of a significand rounded to `decimals` decimals
 * (trailing zeros dropped), or null when that rounding loses every
 * significant digit. */
export function fixedCandidate(sig: ReturnType<typeof significand>, decimals: number, decimal = "."): string | null {
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
  const rounded = roundDigits(integer + fraction, integer.length + decimals);
  const integerLength = integer.length + (rounded.carried ? 1 : 0);
  const whole = rounded.digits.slice(0, integerLength);
  const decimalsText = rounded.digits.slice(integerLength).replace(/0+$/, "");
  if (/^0*$/.test(whole + decimalsText)) {
    return null;
  }
  return `${sign}${whole}${decimalsText ? `${decimal}${decimalsText}` : ""}`;
}

/** The number of decimals a significand's plain fixed text carries. */
export function fixedDecimals(sig: ReturnType<typeof significand>): number {
  return Math.max(0, sig.digits.length - sig.exponent - 1);
}

/** The fixed-notation General text, or null when it cannot fit in
 * GENERAL_WIDTH characters with a significant digit left. */
function fixedText(sig: ReturnType<typeof significand>): string | null {
  const sign = sig.negative ? "-" : "";
  const plain = fixedCandidate(sig, fixedDecimals(sig));
  if (plain !== null && plain.length <= GENERAL_WIDTH) {
    return plain;
  }
  const integerLength = sig.exponent >= 0 ? sig.exponent + 1 : 1;
  // Round off decimals until it fits — down to none at all.
  for (let decimals = Math.max(0, GENERAL_WIDTH - sign.length - integerLength - 1); decimals >= 0; decimals -= 1) {
    const candidate = fixedCandidate(sig, decimals);
    if (candidate === null) {
      return null;
    }
    if (candidate.length <= GENERAL_WIDTH) {
      return candidate;
    }
  }
  return null;
}

/** The scientific General text: as many mantissa decimals as fit. */
function scientificText(sig: ReturnType<typeof significand>): string {
  let text = "";
  for (let decimals = GENERAL_WIDTH - 1; decimals >= 0; decimals -= 1) {
    text = scientificCandidate(sig, decimals);
    if (text.length <= GENERAL_WIDTH) {
      return text;
    }
  }
  return text;
}

/** A number as Excel's General format shows it in a full-width cell, with
 * `decimal` as the decimal separator. */
export function generalNumberText(n: number, decimal = "."): string {
  if (n === 0) {
    return "0";
  }
  if (!Number.isFinite(n)) {
    return "#NUM!";
  }
  const sig = significand(n);
  const magnitude = Math.abs(n);
  let text: string | null = null;
  if (magnitude >= 1e-9 && magnitude < 1e11) {
    text = fixedText(sig);
  }
  text ??= scientificText(sig);
  return decimal === "." ? text : text.replace(".", decimal);
}

/** A number in plain notation (SPEC-118 "editText"): up to 15 significant
 * digits, no grouping, `decimal` as the separator, scientific
 * ("1.23456789012346E+17") only when |n| ≥ 1e15 or 0 < |n| < 1e−9. */
export function plainNotation(n: number, decimal = "."): string {
  if (n === 0) {
    return "0";
  }
  if (!Number.isFinite(n)) {
    return "#NUM!";
  }
  const sig = significand(n);
  if (sig.exponent >= 15 || sig.exponent < -9) {
    return scientificCandidate(sig, sig.digits.length - 1, decimal);
  }
  return fixedCandidate(sig, fixedDecimals(sig), decimal) ?? "0";
}
