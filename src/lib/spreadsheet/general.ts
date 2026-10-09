// The "General" display of a cell value at full column width (SPEC-116),
// matching Excel's General format. SPEC-118 subsumed it: the rule lives in
// numberFormat/decimal.ts and formatGeneral (numberFormat/render.ts) adds
// width-aware shrinking and the locale's decimal separator; at full width in
// en-US its output equals what this function always returned, so this stays
// the one-line way to ask for it.
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
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { formatGeneral, type FormatContext } from "./numberFormat/render";

/** Full width, en-US: SPEC-116's General. */
const GENERAL_CONTEXT: FormatContext = { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" };

/** A cell value as Excel's General format shows it in a full-width cell. */
export function generalText(value: CellValue): string {
  return formatGeneral(value, GENERAL_CONTEXT).text;
}
