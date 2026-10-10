// Value input parsing (SPEC-122 "Input parsing") — THE normative source for
// what the text a person commits into a cell becomes: a value, a formula,
// and the number format the text's shape suggests.
//
// Every number, date, time and date-time VALUE comes from SPEC-120's
// valueFromText (formula/values.ts), the one text → number grammar, so a
// typed value and a text a formula coerces never disagree. parseInput adds
// only the steps around it and the classification of which shape matched,
// in this order:
//
// 1. "" → null.
// 2. An existing format of exactly "@" (Text) → the text verbatim, even
//    "=A1" or "'x".
// 3. "=" followed by anything → a formula: the rest, without "=", NOT
//    validated here (the caller runs SPEC-119's parseFormula). A lone "=" is
//    text, and a leading "+" or "-" never starts a formula.
// 4. A leading apostrophe → the text after it, with `quotePrefix`.
// 5. TRUE / FALSE and every ErrorCode literal, case-insensitively (spaces
//    around them are ignored, as valueFromText ignores them around numbers).
// 6. "1 1/2", "-0 3/16": an integer, spaces, a fraction (parsed here — no
//    other grammar reads it) → `# ?/?`, `# ??/??` or `# ???/???` by the
//    denominator's digits; a denominator of 0 or of more than three digits
//    is no fraction.
// 7. A number in valueFromText's number grammar (mirrored here only to tell
//    its shapes apart): digits after the 15th significant digit become 0
//    before valueFromText reads it (Excel's truncation); the format is the
//    locale's currency (`"$"#,##0`, de-DE `#,##0" €"`, `.00` with decimals)
//    when the currency symbol was typed, else `0%` / `0.00%` for a percent,
//    else `0.00E+00` for an exponent, else `#,##0` / `#,##0.00` when digits
//    were grouped, else none. A plain number (no %, no currency) typed into
//    a percent-formatted cell is Excel's automatic percent entry: magnitude
//    1 or more is divided by 100, a smaller one is already the fraction
//    ("10" → 0.1, "0.5" → 0.5).
// 8. Anything else valueFromText reads is a date, a time or both:
//    `yyyy-mm-ddThh:mm` and a date followed by a time → builtinFormatId 22;
//    a time → [18] h:mm AM/PM / [19] h:mm:ss AM/PM with AM/PM, else [46]
//    [h]:mm:ss from 24 hours up, else [20] h:mm / [21] h:mm:ss (fractional
//    seconds keep [21]: no built-in shows them with the hours); a numeric or
//    ISO date with a year → builtinFormatId 14; day and month only ("10/8",
//    "8-Oct", "Oct 8") → [16] d-mmm; a month and a four-digit year
//    ("Oct-2026") → [17] mmm-yy; a month-name date with a year → [15]
//    d-mmm-yy. A date outside the system's range is no value (text).
// 9. Anything else → the text as typed.
//
// `numberFormat` is a suggestion: the caller (cells.ts setCellEntry) applies
// it only to a cell whose format is General. 14 and 22 are SPEC-118's
// locale-dependent built-ins, returned as `builtinFormatId` with their en-US
// code; [15]–[21] and [46] are locale-independent codes, returned as codes.
// The current year of a yearless date comes from core now() inside
// valueFromText (tests pin it with setClock).
import { isMonthName, valueFromText, roundTo15 } from "../formula/values";
import { ERROR_CODES, type CellError, type CellValue, type ErrorCode } from "../model";
import { BUILTIN_NUMBER_FORMATS } from "./builtin";
import type { SpreadsheetLocale } from "./locale";
import { parseFormatCode } from "./parse";

export interface ParsedInput {
  value: CellValue;
  /** Canonical formula text without "=" — validated by the caller. */
  formula?: string;
  /** Applied only when the cell's format is General. */
  numberFormat?: string;
  /** With `numberFormat`, for SPEC-118's locale-dependent built-ins (14, 22). */
  builtinFormatId?: number;
  quotePrefix?: true;
}

export interface ParseInputOptions {
  /** The code the cell currently renders with (formatCodeOf its effective
   *  style); absent = General. */
  existingFormat?: string;
}

/** Excel keeps 15 significant digits of a typed number. */
export const INPUT_SIGNIFICANT_DIGITS = 15;

const SPACE_LIKE = new Set([" ", " ", " "]);
const ERROR_LITERALS: ReadonlyMap<string, ErrorCode> = new Map(ERROR_CODES.map((code) => [code.toUpperCase(), code]));
const FRACTION = /^([+-]?)(\d+) +(\d+)\/(\d+)$/;
const ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const ISO_DATE_TIME = /^(\d{4})-(\d{1,2})-(\d{1,2})T(.+)$/;
const NUMERIC_DATE = /^(\d{1,4})([/.-])(\d{1,2})(?:\2(\d{1,4}))?$/;
const TIME = /^(\d+)(?::(\d{1,2})(?::(\d{1,2})(\.\d+)?)?)?$/;
const MONTH_NAME_TEXT = /^[\s/.,\-\p{L}\d]+$/u;

/** Leading and trailing U+0020 removed (valueFromText's trimming). */
function trimSpaces(text: string): string {
  let start = 0;
  let end = text.length;
  while (start < end && text[start] === " ") {
    start += 1;
  }
  while (end > start && text[end - 1] === " ") {
    end -= 1;
  }
  return text.slice(start, end);
}

function isDigit(ch: string | undefined): boolean {
  return ch !== undefined && ch >= "0" && ch <= "9";
}

function digitsAt(text: string, i: number): string {
  let end = i;
  while (isDigit(text[end])) {
    end += 1;
  }
  return text.slice(i, end);
}

function isGroupChar(ch: string | undefined, locale: SpreadsheetLocale): boolean {
  if (ch === undefined) {
    return false;
  }
  return ch === locale.group || (SPACE_LIKE.has(locale.group) && SPACE_LIKE.has(ch));
}

/** Which parts of valueFromText's number grammar a text used. */
interface NumberShape {
  grouped: boolean;
  /** Digits after the decimal separator. */
  decimals: boolean;
  percent: boolean;
  currency: boolean;
  exponent: boolean;
  /** The mantissa's digit runs, [start, end) offsets in the text. */
  digitRuns: [number, number][];
}

/** The shape of a text valueFromText's number grammar accepts, or null — a
 * step-for-step mirror of formula/values.ts numberFromText that records what
 * matched instead of computing the value. */
function numberShape(text: string, locale: SpreadsheetLocale): NumberShape | null {
  let s = text;
  let offset = 0;
  let negative = false;
  if (s.startsWith("(") && s.endsWith(")") && s.length > 2) {
    negative = true;
    s = s.slice(1, -1);
    offset = 1;
  }
  let i = 0;
  let signed = false;
  const takeSign = (): void => {
    if (!signed && !negative && (s[i] === "-" || s[i] === "+")) {
      negative = s[i] === "-";
      signed = true;
      i += 1;
    }
  };
  takeSign();
  let currency = false;
  const symbol = locale.currencySymbol;
  if (locale.currencyPosition === "prefix" && s.startsWith(symbol, i)) {
    currency = true;
    i += symbol.length;
    takeSign();
  }
  const digitRuns: [number, number][] = [];
  const integer = digitsAt(s, i);
  digitRuns.push([offset + i, offset + i + integer.length]);
  i += integer.length;
  let grouped = false;
  if (integer.length > 0 && isGroupChar(s[i], locale)) {
    if (integer.length > 3) {
      return null;
    }
    while (isGroupChar(s[i], locale)) {
      const group = digitsAt(s, i + 1);
      if (group.length !== 3) {
        if (SPACE_LIKE.has(s[i]!) && group.length === 0) {
          break;
        }
        return null;
      }
      digitRuns.push([offset + i + 1, offset + i + 4]);
      grouped = true;
      i += 4;
    }
  }
  let fraction = "";
  if (s.startsWith(locale.decimal, i)) {
    i += locale.decimal.length;
    fraction = digitsAt(s, i);
    digitRuns.push([offset + i, offset + i + fraction.length]);
    i += fraction.length;
  }
  if (integer.length === 0 && fraction.length === 0) {
    return null;
  }
  let exponent = false;
  if ((s[i] === "E" || s[i] === "e") && i + 1 < s.length) {
    let j = i + 1;
    if (s[j] === "+" || s[j] === "-") {
      j += 1;
    }
    const digits = digitsAt(s, j);
    if (digits.length === 0) {
      return null;
    }
    exponent = true;
    i = j + digits.length;
  }
  let percent = false;
  if (!currency && s[i] === "%") {
    percent = true;
    i += 1;
  }
  if (!currency && !percent && locale.currencyPosition === "suffix") {
    let j = i;
    if (SPACE_LIKE.has(s[j] ?? "")) {
      j += 1;
    }
    if (s.startsWith(symbol, j) && j + symbol.length === s.length) {
      currency = true;
      i = s.length;
    }
  }
  if (i !== s.length) {
    return null;
  }
  return { grouped, decimals: fraction.length > 0, percent, currency, exponent, digitRuns };
}

/** `text` with every mantissa digit after the 15th significant one made 0
 * ("1234567890123456789" → "1234567890123450000"). */
function truncateDigits(text: string, runs: readonly [number, number][]): string {
  let significant = 0;
  let out = "";
  let last = 0;
  for (const [start, end] of runs) {
    out += text.slice(last, start);
    for (let i = start; i < end; i += 1) {
      const ch = text[i]!;
      if (significant === 0 && ch === "0") {
        out += ch;
        continue;
      }
      significant += 1;
      out += significant > INPUT_SIGNIFICANT_DIGITS ? "0" : ch;
    }
    last = end;
  }
  return out + text.slice(last);
}

/** The locale's currency format: its symbol before the number, or after it
 * following a space. */
export function currencyFormatFor(locale: SpreadsheetLocale, decimals: boolean): string {
  const number = decimals ? "#,##0.00" : "#,##0";
  return locale.currencyPosition === "prefix"
    ? `"${locale.currencySymbol}"${number}`
    : `${number}" ${locale.currencySymbol}"`;
}

/** Whether a format code shows numbers as percentages (a `%` in a non-date
 * section — editText's rule, so what the editor shows re-enters unchanged). */
export function isPercentFormat(code: string | undefined): boolean {
  if (code === undefined) {
    return false;
  }
  try {
    return parseFormatCode(code).sections.some((s) => s.kind !== "date" && s.tokens.some((token) => token.type === "percent"));
  } catch {
    return false;
  }
}

/** A built-in code as a suggestion: locale-dependent ids (14, 22) carry the
 * id with their en-US code, the others their code alone. */
function builtin(id: number): Pick<ParsedInput, "numberFormat" | "builtinFormatId"> {
  const numberFormat = BUILTIN_NUMBER_FORMATS[id]!;
  return id === 14 || id === 22 ? { numberFormat, builtinFormatId: id } : { numberFormat };
}

/** A time of day as valueFromText reads it (formula/values.ts timeFromText,
 * mirrored with its validation), or null. */
function timeShape(text: string, locale: SpreadsheetLocale): { meridiem: boolean; seconds: boolean; hours: number } | null {
  let body = text;
  let meridiem = false;
  const upper = body.toUpperCase();
  for (const marker of locale.amPm) {
    const at = upper.lastIndexOf(marker.toUpperCase());
    if (at > 0 && at + marker.length === upper.length) {
      meridiem = true;
      body = trimSpaces(body.slice(0, at));
      break;
    }
  }
  const match = TIME.exec(body);
  if (!match || (match[2] === undefined && !meridiem)) {
    return null;
  }
  const hours = Number(match[1]);
  const minutes = Number(match[2] ?? 0);
  const seconds = Number(match[3] ?? 0) + Number(match[4] ?? 0);
  if (minutes > 59 || seconds >= 60 || (meridiem && hours > 12)) {
    return null;
  }
  return { meridiem, seconds: match[3] !== undefined, hours };
}

/** The built-in a date-only text suggests (formula/values.ts numericDate
 * and monthNameDate's shapes — the month a name the locale or English
 * spells, so "10/8 1PM" is a date and a time), or null when the text has
 * neither shape. */
function dateFormatId(text: string, locale: SpreadsheetLocale): 14 | 15 | 16 | 17 | null {
  if (ISO_DATE.test(text)) {
    return 14;
  }
  const numeric = NUMERIC_DATE.exec(text);
  if (numeric) {
    return numeric[4] === undefined ? 16 : 14;
  }
  const parts = text.split(/[\s/.,-]+/).filter((part) => part.length > 0);
  if (parts.length < 2 || parts.length > 3 || !MONTH_NAME_TEXT.test(text)) {
    return null;
  }
  const index = parts.findIndex((part) => !/^\d+$/.test(part));
  if (index < 0 || parts.some((part, i) => i !== index && !/^\d+$/.test(part)) || !isMonthName(parts[index]!, locale)) {
    return null;
  }
  const numbers = parts.filter((_, i) => i !== index);
  if (index === 0 && numbers.length === 1) {
    return numbers[0]!.length === 4 ? 17 : 16;
  }
  return numbers.length === 2 ? 15 : 16;
}

/** The suggestion for a text valueFromText read as a date and/or time. */
function dateTimeSuggestion(text: string, locale: SpreadsheetLocale): Pick<ParsedInput, "numberFormat" | "builtinFormatId"> {
  if (ISO_DATE_TIME.test(text)) {
    return builtin(22);
  }
  const time = timeShape(text, locale);
  if (time) {
    if (time.meridiem) {
      return builtin(time.seconds ? 19 : 18);
    }
    if (time.hours >= 24) {
      return builtin(46);
    }
    return builtin(time.seconds ? 21 : 20);
  }
  const date = dateFormatId(text, locale);
  // Neither a time nor a date alone: valueFromText's remaining form, a date
  // followed by a time.
  return builtin(date ?? 22);
}

/** The suggestion for a number by the shape it was typed in. */
function numberSuggestion(shape: NumberShape, locale: SpreadsheetLocale): Pick<ParsedInput, "numberFormat"> {
  if (shape.currency) {
    return { numberFormat: currencyFormatFor(locale, shape.decimals) };
  }
  if (shape.percent) {
    return { numberFormat: shape.decimals ? "0.00%" : "0%" };
  }
  if (shape.exponent) {
    return { numberFormat: BUILTIN_NUMBER_FORMATS[11]! };
  }
  if (shape.grouped) {
    return { numberFormat: shape.decimals ? "#,##0.00" : "#,##0" };
  }
  return {};
}

/** TRUE / FALSE / an error code, case-insensitively, or undefined. */
function literalValue(text: string): boolean | CellError | undefined {
  const upper = text.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return upper === "TRUE";
  }
  const code = ERROR_LITERALS.get(upper);
  return code === undefined ? undefined : { error: code };
}

/** "1 1/2" → 1.5 with `# ?/?`; null for any other text. */
function fractionInput(text: string): ParsedInput | null {
  const match = FRACTION.exec(text);
  if (!match) {
    return null;
  }
  const [, sign, whole, numerator, denominator] = match as unknown as [string, string, string, string, string];
  const den = Number(denominator);
  if (den === 0 || denominator.length > 3) {
    return null;
  }
  const magnitude = Number(whole) + Number(numerator) / den;
  const value = sign === "-" && magnitude !== 0 ? -magnitude : magnitude;
  const marks = "?".repeat(denominator.length);
  return { value, numberFormat: `# ${marks}/${marks}` };
}

/** What a committed editor text becomes (see the module header). */
export function parseInput(
  text: string,
  locale: SpreadsheetLocale,
  dateSystem: "1900" | "1904",
  options: ParseInputOptions = {},
): ParsedInput {
  if (text.length === 0) {
    return { value: null };
  }
  if (options.existingFormat === "@") {
    return { value: text };
  }
  if (text.startsWith("=") && text.length > 1) {
    return { value: null, formula: text.slice(1) };
  }
  if (text.startsWith("'")) {
    return { value: text.slice(1), quotePrefix: true };
  }
  const trimmed = trimSpaces(text);
  const literal = literalValue(trimmed);
  if (literal !== undefined) {
    return { value: literal };
  }
  const fraction = fractionInput(trimmed);
  if (fraction) {
    return fraction;
  }
  const shape = numberShape(trimmed, locale);
  if (shape) {
    let value = valueFromText(truncateDigits(trimmed, shape.digitRuns), locale, dateSystem);
    if (value === null) {
      return { value: text }; // past the double range: text, as valueFromText says
    }
    if (!shape.percent && !shape.currency && Math.abs(value) >= 1 && isPercentFormat(options.existingFormat)) {
      value = roundTo15(value / 100);
    }
    return { value, ...numberSuggestion(shape, locale) };
  }
  const value = valueFromText(trimmed, locale, dateSystem);
  if (value === null) {
    return { value: text };
  }
  return { value, ...dateTimeSuggestion(trimmed, locale) };
}
