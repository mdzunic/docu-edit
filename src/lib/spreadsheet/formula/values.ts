// The calculation engine's value model (SPEC-120 "Values", "Coercion", "Text
// to number", "Comparison") — THE normative source for what a formula
// computes with and how one kind of value becomes another:
//
// - EvalValue: a scalar (number, string, boolean, CellError or the EMPTY
//   singleton), a RangeValue (a reference, read lazily — one Area per sheet of
//   a 3-D reference) or an ArrayValue (row-major). EMPTY is what reading an
//   empty cell or a `missing` argument produces; it never reaches the model
//   (toCellValue stores it as 0).
// - Coercion (toNumber / toText / toLogical) is Excel's "Converting values in
//   formulas" table: text becomes a number only through valueFromText, a
//   boolean is 1/0 and "TRUE"/"FALSE", EMPTY is 0, "" and FALSE at once, and
//   an error stays that error.
// - valueFromText is the ONE text → number/date/time grammar: a formula's
//   coercion uses it and SPEC-122's parseInput takes every value it
//   recognises from it, so a typed value and a coerced one never disagree.
//   It reads the user's SpreadsheetLocale (never the file's): separators,
//   currency, date order and month names. Booleans are never parsed here.
// - numberToText is the locale-independent text a number becomes inside a
//   formula (`=1/3&""`), SPEC-118's plain notation with "." as the point.
// - compareScalars orders number < text < logical, gives an EMPTY operand the
//   other side's type, compares numbers at 15 significant digits and text
//   through one Intl.Collator per locale (accent-sensitive, case-insensitive;
//   case-sensitive for an .ods whose settings say so).
import { now } from "../../core/clock";
import type { GridRange } from "../address";
import { ERROR_CODES, type CellError, type CellValue, type ErrorCode } from "../model";
import { plainNotation } from "../numberFormat/decimal";
import { localeNames, type SpreadsheetLocale } from "../numberFormat/locale";
import { dateToSerial } from "../serial";
import type { EvalSettings } from "./evaluate";

/** The most elements a range may materialise into, or an array hold. */
export const MAX_ARRAY_CELLS = 4_194_304;
/** The longest text a formula may produce (Excel's cell limit). */
export const MAX_TEXT_LENGTH = 32_767;

export interface EmptyValue {
  kind: "empty";
}
/** An empty cell or argument — the one EmptyValue. */
export const EMPTY: EmptyValue = Object.freeze({ kind: "empty" as const });

/** One rectangle of a reference, on one sheet. */
export interface Area extends GridRange {
  sheetId: string;
}
/** A reference, read lazily: one Area per sheet of a 3-D reference, several
 * for a union. */
export interface RangeValue {
  kind: "range";
  areas: Area[];
}
export type ScalarValue = number | string | boolean | CellError | EmptyValue;
/** A rows × cols array, row-major. */
export interface ArrayValue {
  kind: "array";
  rows: number;
  cols: number;
  data: ScalarValue[];
}
export type EvalValue = ScalarValue | RangeValue | ArrayValue;

// --- errors and kinds ------------------------------------------------------------

const ERROR_VALUES: ReadonlyMap<ErrorCode, CellError> = new Map(ERROR_CODES.map((code) => [code, Object.freeze({ error: code })]));

/** The (shared, frozen) error value of a code. */
export function errorOf(code: ErrorCode): CellError {
  return ERROR_VALUES.get(code)!;
}

export const NULL_ERROR = errorOf("#NULL!");
export const DIV0_ERROR = errorOf("#DIV/0!");
export const VALUE_ERROR = errorOf("#VALUE!");
export const REF_ERROR = errorOf("#REF!");
export const NAME_ERROR = errorOf("#NAME?");
export const NUM_ERROR = errorOf("#NUM!");
export const NA_ERROR = errorOf("#N/A");

/** Whether a value — or a function's interim result — is an error. */
export function isError(value: unknown): value is CellError {
  return typeof value === "object" && value !== null && "error" in value;
}

/** Whether a value is EMPTY. */
export function isEmpty(value: EvalValue): value is EmptyValue {
  return value === EMPTY || (typeof value === "object" && "kind" in value && value.kind === "empty");
}

export function isRange(value: EvalValue): value is RangeValue {
  return typeof value === "object" && "kind" in value && value.kind === "range";
}

export function isArray(value: EvalValue): value is ArrayValue {
  return typeof value === "object" && "kind" in value && value.kind === "array";
}

/** A cell's value as the engine reads it: no value → EMPTY. */
export function fromCellValue(value: CellValue): ScalarValue {
  return value === null ? EMPTY : value;
}

// --- coercion --------------------------------------------------------------------

/** A scalar as a number (see "Coercion"). */
export function toNumber(value: ScalarValue, settings: Pick<EvalSettings, "locale" | "dateSystem">): number | CellError {
  switch (typeof value) {
    case "number":
      return value;
    case "boolean":
      return value ? 1 : 0;
    case "string":
      return valueFromText(value, settings.locale, settings.dateSystem) ?? VALUE_ERROR;
    default:
      return isError(value) ? value : 0;
  }
}

/** A scalar as text (see "Coercion"). */
export function toText(value: ScalarValue): string | CellError {
  switch (typeof value) {
    case "string":
      return value;
    case "number":
      return numberToText(value);
    case "boolean":
      return value ? "TRUE" : "FALSE";
    default:
      return isError(value) ? value : "";
  }
}

/** A scalar as a logical (see "Coercion"): text only as "TRUE"/"FALSE". */
export function toLogical(value: ScalarValue): boolean | CellError {
  switch (typeof value) {
    case "boolean":
      return value;
    case "number":
      return value !== 0;
    case "string": {
      const upper = value.toUpperCase();
      return upper === "TRUE" ? true : upper === "FALSE" ? false : VALUE_ERROR;
    }
    default:
      return isError(value) ? value : false;
  }
}

/** A number as a formula's text: 15 significant digits, fixed notation for
 * 1e-9 ≤ |n| < 1e15 (trailing zeros dropped), scientific otherwise
 * ("1E+15"). Locale-independent, as in Excel. */
export function numberToText(n: number): string {
  return plainNotation(n);
}

/** n at 15 significant digits — Number(n.toPrecision(15)); 0 (and −0) is 0. */
export function roundTo15(n: number): number {
  if (n === 0) {
    return 0;
  }
  return Number.isFinite(n) ? Number(n.toPrecision(15)) : n;
}

/** x^y — the `^` operator and POWER (see "Errors"): 0^0 and a negative base
 * with a non-integer exponent are #NUM!, 0 to a negative power #DIV/0!, a
 * non-finite result #NUM!. */
export function power(x: number, y: number): number | CellError {
  if (x === 0 && y === 0) {
    return NUM_ERROR;
  }
  if (x === 0 && y < 0) {
    return DIV0_ERROR;
  }
  if (x < 0 && !Number.isInteger(y)) {
    return NUM_ERROR;
  }
  const result = x ** y;
  return Number.isFinite(result) ? result : NUM_ERROR;
}

/** What a scalar result stores as: EMPTY → 0, −0 → 0, a non-finite number
 * → #NUM!, text past MAX_TEXT_LENGTH → #VALUE!, an error as a fresh object. */
export function toCellValue(value: ScalarValue): CellValue {
  switch (typeof value) {
    case "number":
      return Number.isFinite(value) ? (value === 0 ? 0 : value) : { error: "#NUM!" };
    case "string":
      return value.length > MAX_TEXT_LENGTH ? { error: "#VALUE!" } : value;
    case "boolean":
      return value;
    default:
      return isError(value) ? { error: value.error } : 0;
  }
}

// --- text to number (valueFromText) ------------------------------------------------

const SPACE_LIKE = new Set([" ", " ", " "]);
const ENGLISH_MONTHS = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december",
];

function isDigit(text: string, i: number): boolean {
  const code = text.charCodeAt(i);
  return code >= 0x30 && code <= 0x39;
}

/** The run of digits starting at `i`. */
function digitsAt(text: string, i: number): string {
  let end = i;
  while (end < text.length && isDigit(text, end)) {
    end += 1;
  }
  return text.slice(i, end);
}

/** Leading and trailing U+0020 removed (and nothing else). */
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

/** Whether `ch` separates digit groups in `locale` (a space-like group
 * separator also accepts U+0020, U+00A0 and U+202F). */
function isGroupChar(ch: string | undefined, locale: SpreadsheetLocale): boolean {
  if (ch === undefined) {
    return false;
  }
  return ch === locale.group || (SPACE_LIKE.has(locale.group) && SPACE_LIKE.has(ch));
}

/** The number grammar of "Text to number" over trimmed text, or null. */
function numberFromText(text: string, locale: SpreadsheetLocale): number | null {
  let s = text;
  let negative = false;
  if (s.startsWith("(") && s.endsWith(")") && s.length > 2) {
    negative = true;
    s = s.slice(1, -1);
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
  // Integer digits, grouped only in threes after the first group.
  let integer = digitsAt(s, i);
  i += integer.length;
  if (integer.length > 0 && isGroupChar(s[i], locale)) {
    if (integer.length > 3) {
      return null;
    }
    while (isGroupChar(s[i], locale)) {
      const group = digitsAt(s, i + 1);
      if (group.length !== 3) {
        // A space-like separator before a suffix currency is not a group.
        if (SPACE_LIKE.has(s[i]!) && group.length === 0) {
          break;
        }
        return null;
      }
      integer += group;
      i += 4;
    }
  }
  let fraction = "";
  if (s.startsWith(locale.decimal, i)) {
    i += locale.decimal.length;
    fraction = digitsAt(s, i);
    i += fraction.length;
  }
  if (integer.length === 0 && fraction.length === 0) {
    return null;
  }
  let exponent = 0;
  if ((s[i] === "E" || s[i] === "e") && i + 1 < s.length) {
    let j = i + 1;
    let expNegative = false;
    if (s[j] === "+" || s[j] === "-") {
      expNegative = s[j] === "-";
      j += 1;
    }
    const digits = digitsAt(s, j);
    if (digits.length === 0) {
      return null;
    }
    exponent = Number(digits) * (expNegative ? -1 : 1);
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
  const value = Number(`${integer || "0"}.${fraction || "0"}e${exponent - (percent ? 2 : 0)}`);
  if (!Number.isFinite(value)) {
    return null;
  }
  return negative && value !== 0 ? -value : value;
}

const monthTables = new Map<string, ReadonlyMap<string, number>>();

/** Month names a date may spell, lower-case → 1–12: English names and their
 * three-letter abbreviations, plus the locale's own long, genitive and short
 * names and three-letter prefixes (an ambiguous prefix is left out). */
function monthTable(locale: SpreadsheetLocale): ReadonlyMap<string, number> {
  let table = monthTables.get(locale.tag);
  if (table) {
    return table;
  }
  const names = localeNames(locale.tag);
  const candidates = new Map<string, Set<number>>();
  const add = (name: string, month: number): void => {
    const key = name.toLowerCase().replace(/\.$/, "");
    if (key.length === 0) {
      return;
    }
    const months = candidates.get(key) ?? new Set<number>();
    months.add(month);
    candidates.set(key, months);
  };
  ENGLISH_MONTHS.forEach((name, index) => {
    add(name, index + 1);
    add(name.slice(0, 3), index + 1);
  });
  for (let index = 0; index < 12; index += 1) {
    for (const name of [names.monthsLong[index], names.monthsGenitive[index], names.monthsShort[index]]) {
      if (name) {
        add(name, index + 1);
        add(name.slice(0, 3), index + 1);
      }
    }
  }
  const built = new Map<string, number>();
  for (const [name, months] of candidates) {
    if (months.size === 1) {
      built.set(name, [...months][0]!);
    }
  }
  table = built;
  monthTables.set(locale.tag, table);
  return table;
}

/** A two-digit year as Excel reads it — 00–29 → 2000s, 30–99 → 1900s —
 * or four digits as written; anything else is no year. */
function fullYear(digits: string): number | null {
  if (digits.length === 4) {
    return Number(digits);
  }
  if (digits.length === 2) {
    const year = Number(digits);
    return year < 30 ? 2000 + year : 1900 + year;
  }
  return null;
}

/** Days in a month, February 1900 having Excel's 29 in the 1900 system. */
function daysInMonth(year: number, month: number, dateSystem: "1900" | "1904"): number {
  if (month === 2) {
    const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 || (year === 1900 && dateSystem === "1900");
    return leap ? 29 : 28;
  }
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}

/** The serial of a calendar date, or null for an impossible one or one
 * outside the system's range (before its first day, after 9999-12-31). */
function dateSerial(year: number | null, month: number, day: number, dateSystem: "1900" | "1904"): number | null {
  if (year === null || month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month, dateSystem) || year > 9999) {
    return null;
  }
  if (year < (dateSystem === "1900" ? 1900 : 1904)) {
    return null;
  }
  return dateToSerial(year, month, day, dateSystem);
}

function currentYear(): number {
  return now().getFullYear();
}

/** A numeric date in the locale's order — day, month and year (two or four
 * digits) separated by `/`, `-` or `.` (`6/1/2001`, `10-8-26`), or day and
 * month of the current year (`6/1`) — or ISO `yyyy-mm-dd`; else null. */
function numericDate(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): number | null {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (iso) {
    return dateSerial(Number(iso[1]), Number(iso[2]), Number(iso[3]), dateSystem);
  }
  const match = /^(\d{1,4})([/.-])(\d{1,2})(?:\2(\d{1,4}))?$/.exec(text);
  if (!match) {
    return null;
  }
  const [, a, , b, c] = match as unknown as [string, string, string, string, string | undefined];
  if (c === undefined) {
    if (a.length > 2) {
      return null;
    }
    const [month, day] = locale.dateOrder === "DMY" ? [b, a] : [a, b];
    return dateSerial(currentYear(), Number(month), Number(day), dateSystem);
  }
  switch (locale.dateOrder) {
    case "MDY":
      return a.length > 2 ? null : dateSerial(fullYear(c), Number(a), Number(b), dateSystem);
    case "DMY":
      return a.length > 2 ? null : dateSerial(fullYear(c), Number(b), Number(a), dateSystem);
    default:
      return c.length > 2 ? null : dateSerial(fullYear(a), Number(b), Number(c), dateSystem);
  }
}

/** A date spelling its month (`8-Oct-2026`, `Oct 8, 2026`, `8 October 2026`,
 * `8. Oktober 2026`, `Oct-2026`), or null. */
function monthNameDate(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): number | null {
  const parts = text.split(/[\s/.,-]+/).filter((part) => part.length > 0);
  if (parts.length < 2 || parts.length > 3 || !/^[\s/.,\-\p{L}\d]+$/u.test(text)) {
    return null;
  }
  const months = monthTable(locale);
  const index = parts.findIndex((part) => !/^\d+$/.test(part));
  if (index < 0 || parts.some((part, i) => i !== index && !/^\d+$/.test(part))) {
    return null;
  }
  const month = months.get(parts[index]!.toLowerCase());
  if (month === undefined) {
    return null;
  }
  const numbers = parts.filter((_, i) => i !== index);
  if (index === 0 && numbers.length === 1) {
    // "Oct-2026" (the 1st of that month) or "Oct 8" (this year).
    const only = numbers[0]!;
    if (only.length === 4) {
      return dateSerial(Number(only), month, 1, dateSystem);
    }
    return only.length > 2 ? null : dateSerial(currentYear(), month, Number(only), dateSystem);
  }
  if (index === 0 || index === 1) {
    const day = numbers[0]!;
    if (day.length > 2) {
      return null;
    }
    const year = numbers.length === 2 ? fullYear(numbers[1]!) : currentYear();
    return dateSerial(year, month, Number(day), dateSystem);
  }
  return null;
}

/** A date part, or null. */
function dateFromText(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): number | null {
  return numericDate(text, locale, dateSystem) ?? monthNameDate(text, locale, dateSystem);
}

/** A time of day (`13:45`, `13:45:30.5`, `1:45 PM`, `1 pm`, `36:00`) as a
 * day fraction, or null. Hours may pass 24 without AM/PM. */
function timeFromText(text: string, locale: SpreadsheetLocale): number | null {
  let body = text;
  let meridiem: "am" | "pm" | null = null;
  const upper = body.toUpperCase();
  for (const [index, marker] of locale.amPm.entries()) {
    const at = upper.lastIndexOf(marker.toUpperCase());
    if (at > 0 && at + marker.length === upper.length) {
      meridiem = index === 0 ? "am" : "pm";
      body = trimSpaces(body.slice(0, at));
      break;
    }
  }
  const match = /^(\d+)(?::(\d{1,2})(?::(\d{1,2})(\.\d+)?)?)?$/.exec(body);
  if (!match || (match[2] === undefined && meridiem === null)) {
    return null;
  }
  let hours = Number(match[1]);
  const minutes = Number(match[2] ?? 0);
  const seconds = Number(match[3] ?? 0) + Number(match[4] ?? 0);
  if (minutes > 59 || seconds >= 60) {
    return null;
  }
  if (meridiem !== null) {
    if (hours > 12) {
      return null;
    }
    hours = (hours % 12) + (meridiem === "pm" ? 12 : 0);
  }
  return (hours * 3600 + minutes * 60 + seconds) / 86_400;
}

/** A date, a time, or a date and a time separated by spaces (or ISO's T). */
function dateTimeFromText(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): number | null {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})T(.+)$/.exec(text);
  if (iso) {
    const date = dateSerial(Number(iso[1]), Number(iso[2]), Number(iso[3]), dateSystem);
    const time = timeFromText(iso[4]!, locale);
    return date === null || time === null ? null : date + time;
  }
  if (text.includes(":") || /[a-z]$/i.test(text)) {
    const time = timeFromText(text, locale);
    if (time !== null) {
      return time;
    }
    // A date, then the time after the last run of spaces before a digit.
    for (let at = text.lastIndexOf(" "); at > 0; at = text.lastIndexOf(" ", at - 1)) {
      const timePart = timeFromText(trimSpaces(text.slice(at + 1)), locale);
      if (timePart === null) {
        continue;
      }
      const date = dateFromText(trimSpaces(text.slice(0, at)), locale, dateSystem);
      if (date !== null) {
        return date + timePart;
      }
    }
  }
  return dateFromText(text, locale, dateSystem);
}

/** Text as a number, date or time in `locale` (see "Text to number"); null
 * for anything else — "", "TRUE" and every other text included. */
export function valueFromText(text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): number | null {
  const trimmed = trimSpaces(text);
  if (trimmed.length === 0) {
    return null;
  }
  return numberFromText(trimmed, locale) ?? dateTimeFromText(trimmed, locale, dateSystem);
}

// --- comparison ------------------------------------------------------------------

const collators = new Map<string, Intl.Collator>();

/** The collator text compares through: one per locale and case rule. */
export function collatorFor(settings: Pick<EvalSettings, "locale" | "caseSensitive">): Intl.Collator {
  const key = `${settings.locale.tag}|${settings.caseSensitive ? "variant" : "accent"}`;
  let collator = collators.get(key);
  if (!collator) {
    collator = new Intl.Collator(settings.locale.tag, { sensitivity: settings.caseSensitive ? "variant" : "accent" });
    collators.set(key, collator);
  }
  return collator;
}

/** number < text < logical. */
function typeRank(value: number | string | boolean): number {
  return typeof value === "number" ? 0 : typeof value === "string" ? 1 : 2;
}

/** EMPTY as the other side's type: 0, "" or FALSE. */
function blankLike(other: number | string | boolean): number | string | boolean {
  return typeof other === "number" ? 0 : typeof other === "string" ? "" : false;
}

/** Orders two scalars (see "Comparison"); an error operand is returned,
 * the left one first. */
export function compareScalars(a: ScalarValue, b: ScalarValue, settings: Pick<EvalSettings, "locale" | "caseSensitive">): -1 | 0 | 1 | CellError {
  if (isError(a)) {
    return a;
  }
  if (isError(b)) {
    return b;
  }
  const aEmpty = isEmpty(a);
  const bEmpty = isEmpty(b);
  if (aEmpty && bEmpty) {
    return 0;
  }
  const left = aEmpty ? blankLike(b as number | string | boolean) : (a as number | string | boolean);
  const right = bEmpty ? blankLike(left) : (b as number | string | boolean);
  const rank = typeRank(left) - typeRank(right);
  if (rank !== 0) {
    return rank < 0 ? -1 : 1;
  }
  if (typeof left === "number") {
    const x = roundTo15(left);
    const y = roundTo15(right as number);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof left === "string") {
    if (left === right) {
      return 0;
    }
    const order = collatorFor(settings).compare(left, right as string);
    return order < 0 ? -1 : order > 0 ? 1 : 0;
  }
  return left === right ? 0 : left ? 1 : -1;
}
