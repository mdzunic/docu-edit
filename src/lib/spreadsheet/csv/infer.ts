// CSV field type inference (SPEC-126 "Type inference") — THIS FILE is the
// normative source for what a field of a delimited-text file becomes. One
// rule, deliberately not Excel's: a field becomes a number, date, boolean or
// error only when that is lossless — when the typed value, written back the
// way csv/write.ts writes it (csvFieldText in the file's locale), is exactly
// the trimmed field. Everything else stays text, verbatim, and survives a
// save untouched. Each rule below keeps as text a value Excel would have
// changed, and each is a documented data-loss pattern: postcodes, phone and
// account numbers (leading zeros), card and order numbers (past a double's 15
// digits), identifiers that look like dates (`1-2`, `SEPT2`).
//
// inferCsvField(field), with t = field.trim():
// 1. "" → nothing stored.
// 2. Text when t starts with `=` (CSV carries data, not formulas — and
//    `=HYPERLINK(…)`/DDE payloads are the classic CSV attack) or `'` (CSV has
//    no quote-prefix convention); when t has a leading zero (`007`, `00`,
//    `-0123` — not `0`, `0.5`, `0,5`); or when t is only digits, a sign,
//    `.`, `,` and spaces and holds more than 15 digits.
// 3. SPEC-122's parseInput(t) in the dialect's locale; a formula or a text
//    result stays text.
// 4. The typed value is kept with the first format whose csvFieldText is t:
//    parseInput's own suggestion (absent = General; a locale-dependent
//    built-in 14/22 resolved in the dialect's locale, and kept as that
//    built-in), then the fidelity formats derived from t's shape — for a
//    number with n decimals `0` / `0.` + n zeros, the same with a `#,##0`
//    integer part when t holds the locale's group separator, each with `%`
//    when t ends in `%`; for a date or time `yyyy-mm-dd`, `yyyy-mm-dd hh:mm`,
//    `yyyy-mm-dd hh:mm:ss`, `hh:mm`, `hh:mm:ss`.
// 5. Otherwise text, verbatim (untrimmed).
import type { CellValue } from "../model";
import { formatCodeOf } from "../numberFormat/builtin";
import { parseInput } from "../numberFormat/input";
import type { SpreadsheetLocale } from "../numberFormat/locale";
import { parseFormatCode } from "../numberFormat/parse";
import { csvFieldText } from "./write";

export interface InferredCsvField {
  value: CellValue;
  /** The format that reproduces the field (en-US code syntax); absent =
   * General. */
  numberFormat?: string;
  /** With `numberFormat`: SPEC-118's locale-dependent built-in (14, 22)
   * parseInput suggested — the cell keeps the locale's short date. */
  builtinFormatId?: number;
}

const LEADING_ZERO = /^[+-]?0\d+([.,]\d*)?$/;
const NUMERIC_CHARACTERS = /^[+-]?[\d.,   ]+$/;
const DIGIT = /\d/;
const DIGITS = /\d/g;
const MAX_TYPED_DIGITS = 15;
const SPACE_LIKE = /[   ]/;
const DATE_TIME_FIDELITY_FORMATS = ["yyyy-mm-dd", "yyyy-mm-dd hh:mm", "yyyy-mm-dd hh:mm:ss", "hh:mm", "hh:mm:ss"];

/** Whether t could only ever be text to parseInput: no digit, and neither
 * TRUE/FALSE nor an error literal (every number, date and time has a
 * digit). A fast path — parseInput would agree. */
function plainlyText(t: string): boolean {
  if (DIGIT.test(t) || t.startsWith("#")) {
    return false;
  }
  const upper = t.toUpperCase();
  return upper !== "TRUE" && upper !== "FALSE";
}

function isDateCode(code: string | undefined): boolean {
  if (code === undefined) {
    return false;
  }
  try {
    return parseFormatCode(code).sections.some((section) => section.kind === "date");
  } catch {
    return false;
  }
}

/** The number fidelity formats for t (see the module header). */
function numberFidelityFormats(t: string, locale: SpreadsheetLocale): string[] {
  const percent = t.endsWith("%");
  const body = percent ? t.slice(0, -1) : t;
  const point = body.lastIndexOf(locale.decimal);
  let decimals = 0;
  if (point >= 0) {
    for (let at = point + locale.decimal.length; at < body.length && body[at]! >= "0" && body[at]! <= "9"; at += 1) {
      decimals += 1;
    }
  }
  const fraction = decimals === 0 ? "" : `.${"0".repeat(decimals)}`;
  const suffix = percent ? "%" : "";
  const formats = [`0${fraction}${suffix}`];
  const grouped = body.includes(locale.group) || (SPACE_LIKE.test(locale.group) && SPACE_LIKE.test(body));
  if (grouped) {
    formats.push(`#,##0${fraction}${suffix}`);
  }
  return formats;
}

/** What one field becomes (see the module header). */
export function inferCsvField(field: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): InferredCsvField {
  if (field === "") {
    return { value: null };
  }
  const t = field.trim();
  if (
    t.startsWith("=") ||
    t.startsWith("'") ||
    LEADING_ZERO.test(t) ||
    (NUMERIC_CHARACTERS.test(t) && (t.match(DIGITS)?.length ?? 0) > MAX_TYPED_DIGITS) ||
    plainlyText(t)
  ) {
    return { value: field };
  }
  const parsed = parseInput(t, locale, dateSystem);
  const value = parsed.value;
  if (parsed.formula !== undefined || value === null || typeof value === "string") {
    return { value: field };
  }
  const own = formatCodeOf(parsed, locale);
  if (csvFieldText(value, own, locale, dateSystem) === t) {
    const typed: InferredCsvField = { value };
    if (parsed.numberFormat !== undefined) {
      typed.numberFormat = parsed.numberFormat;
      if (parsed.builtinFormatId !== undefined) {
        typed.builtinFormatId = parsed.builtinFormatId;
      }
    }
    return typed;
  }
  if (typeof value === "number") {
    const candidates = isDateCode(own) ? DATE_TIME_FIDELITY_FORMATS : numberFidelityFormats(t, locale);
    for (const numberFormat of candidates) {
      if (csvFieldText(value, numberFormat, locale, dateSystem) === t) {
        return { value, numberFormat };
      }
    }
  }
  return { value: field };
}
