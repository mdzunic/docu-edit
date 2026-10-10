// Series detection (SPEC-125 "Series detection") — THIS FILE is the
// normative source for detectSeries: what the fill handle writes past a seed
// line (each column of a vertical fill, each row of a horizontal one).
//
// | Seed line (in fill order) | Continuation |
// | `5` | `5, 5, …` (toggle: `6, 7, …`) |
// | `1, 3` | `5, 7, …` |
// | `1, 2, 4` | least squares: `5.33333333333333, 6.83333333333333` |
// | `0.1, 0.2` | `0.3, 0.4` (15 significant digits) |
// | `2026-01-31` (date) | `2026-02-01, …` (toggle: copy) |
// | `2026-01-31, 2026-02-28` | month-end: `2026-03-31, 2026-04-30` |
// | `2026-01-15, 2026-02-15` | `2026-03-15, …` (EDATE from the first seed) |
// | `2024-03-01, 2025-03-01` | `2026-03-01, …` |
// | `2026-01-01, 2026-01-08` | weekly: `2026-01-15, …` |
// | `9:00` (time-only format) | `10:00, 11:00` |
// | `Mon` / `MONDAY` / `monday` | `Tue` / `TUESDAY` / `tuesday` …, wrapping |
// | `Jan, Apr` | `Jul, Oct, Jan` |
// | `Montag` (de-DE locale) | `Dienstag` |
// | `Q3 2026` / `Qtr 4` | `Q4 2026, Q1 2027` / `Qtr 1` |
// | `Item 9` / `Item 009` / `Item 1, Item 3` | `Item 10` / `Item 010` / `Item 5` |
// | `1st` | `2nd, 3rd, 4th, …, 11th, 12th, 13th, 21st` |
// | `=A1*2` at B1 | `=A2*2` at B2 (shifted by the distance) |
// | `TRUE`, `#N/A`, `abc`, `Item 1, Thing 2`, mixed kinds | the seed pattern repeats |
//
// - Numbers: one seed copies (the toggle counts up by 1); two or more follow
//   their least-squares linear trend, every result rounded to 15 significant
//   digits. Dates (a seed whose code SPEC-118 classifies as a date) try, in
//   order: equal year steps on the same month and day; equal month steps on
//   the same day of month, or all on month-ends (EDATE from the first seed,
//   clamped to month ends); equal whole-day steps; else the linear trend on
//   serials. One date counts days; one time-only value counts hours.
// - Weekday and month names, full and short, in English and in the user's
//   SPEC-118 locale, keep the first seed's case (all capitals, all lower or
//   capitalised). Name cycles, quarters, counters and ordinals need a
//   constant integer step between seeds, else the pattern repeats.
// - Formulas are shifted by their distance from the seed cell along the
//   fill's axis (SPEC-119's shiftFormula), in every mode.
// - `at(i)` is the value at index i of the line, counted from its first seed
//   (0-based); i < 0 extends it backwards (filling up or left). `byDefault`
//   says whether a plain drag fills the series (true) or copies (a single
//   number; a repeating pattern).
import type { CellValue } from "./model";
import { shiftFormula } from "./formula/references";
import { localeNames } from "./numberFormat/locale";
import type { SpreadsheetLocale } from "./numberFormat/locale";
import { parseFormatCode } from "./numberFormat/parse";
import { dateToSerial, serialToParts } from "./serial";

/** Days between the 1900 and 1904 date systems' serials of one date. */
export const DATE_SYSTEM_OFFSET = 1462;

/** Whether a format code shows a date ("date": a year, month, day, week or
 * quarter part), only a time of day or a duration ("time"), or neither —
 * SPEC-118's classification of the code's first section. */
export function dateKindOf(code: string | undefined): "date" | "time" | null {
  if (code === undefined || code === "" || code.toLowerCase() === "general") {
    return null;
  }
  let section;
  try {
    section = parseFormatCode(code).sections[0];
  } catch {
    return null;
  }
  if (!section || section.kind !== "date") {
    return null;
  }
  return section.tokens.some((token) => token.type === "date" && token.part !== "hour" && token.part !== "minute" && token.part !== "second") ? "date" : "time";
}

export interface SeriesSeed {
  value: CellValue;
  formula?: string;
  /** The code the seed renders with (formatCodeOf its effective style). */
  code?: string;
}

export type SeriesKind = "repeat" | "number" | "date" | "time" | "name" | "quarter" | "counter" | "ordinal";

export interface SeriesValue {
  value: CellValue;
  formula?: string;
}

export interface Series {
  kind: SeriesKind;
  byDefault: boolean;
  at(i: number): SeriesValue;
}

export interface SeriesOptions {
  /** The fill's axis: "row" for a vertical fill (formulas shift rows),
   * "col" for a horizontal one. Default "row". */
  axis?: "row" | "col";
  dateSystem?: "1900" | "1904";
}

const mod = (a: number, n: number): number => ((a % n) + n) % n;

/** A number at 15 significant digits (SPEC-120's roundTo15). */
function round15(n: number): number {
  return n === 0 || !Number.isFinite(n) ? n : Number(n.toPrecision(15));
}

/** The seed pattern repeated, formulas shifted by their distance. */
function repeatSeries(seed: readonly SeriesSeed[], axis: "row" | "col", kind: SeriesKind = "repeat", byDefault = false): Series {
  return {
    kind,
    byDefault,
    at(i) {
      const j = mod(i, seed.length);
      const source = seed[j]!;
      if (source.formula === undefined) {
        return { value: source.value };
      }
      const distance = i - j;
      return { value: source.value, formula: shiftFormula(source.formula, axis === "row" ? distance : 0, axis === "col" ? distance : 0) };
    },
  };
}

/** The least-squares line through (k, ys[k]): [intercept, slope]. */
function trend(ys: readonly number[]): [number, number] {
  const n = ys.length;
  const meanX = (n - 1) / 2;
  const meanY = ys.reduce((sum, y) => sum + y, 0) / n;
  let num = 0;
  let den = 0;
  ys.forEach((y, k) => {
    num += (k - meanX) * (y - meanY);
    den += (k - meanX) * (k - meanX);
  });
  const slope = den === 0 ? 0 : num / den;
  return [meanY - slope * meanX, slope];
}

/** Every consecutive difference of `values` equal; that step, else null. */
function constantStep(values: readonly number[]): number | null {
  const step = values[1]! - values[0]!;
  for (let k = 2; k < values.length; k += 1) {
    if (values[k]! - values[k - 1]! !== step) {
      return null;
    }
  }
  return step;
}

function numberSeries(values: readonly number[], seed: readonly SeriesSeed[], axis: "row" | "col"): Series {
  if (values.length === 1) {
    const start = values[0]!;
    return { kind: "number", byDefault: false, at: (i) => ({ value: round15(start + i) }) };
  }
  const [intercept, slope] = trend(values);
  return {
    kind: "number",
    byDefault: true,
    at: (i) => (i >= 0 && i < seed.length ? repeatSeries(seed, axis).at(i) : { value: round15(intercept + slope * i) }),
  };
}

const daysIn = (year: number, month: number): number => new Date(Date.UTC(year, month, 0)).getUTCDate();

function dateSeries(values: readonly number[], seed: readonly SeriesSeed[], axis: "row" | "col", system: "1900" | "1904"): Series {
  const series = (at: (i: number) => number): Series => ({
    kind: "date",
    byDefault: true,
    at: (i) => (i >= 0 && i < seed.length ? repeatSeries(seed, axis).at(i) : { value: at(i) }),
  });
  if (values.length === 1) {
    const start = values[0]!;
    return series((i) => start + i);
  }
  if (values.every(Number.isInteger)) {
    const parts = values.map((serial) => serialToParts(serial, system));
    const first = parts[0]!;
    const months = parts.map((p) => p.year * 12 + (p.month - 1));
    const sameDay = parts.every((p) => p.day === first.day);
    // Equal year steps on the same month and day.
    const yearStep = constantStep(parts.map((p) => p.year));
    if (sameDay && parts.every((p) => p.month === first.month) && yearStep !== null && yearStep !== 0) {
      return series((i) => {
        const year = first.year + i * yearStep;
        return dateToSerial(year, first.month, Math.min(first.day, daysIn(year, first.month)), system);
      });
    }
    // Equal month steps on the same day of month, or all on month-ends.
    const monthStep = constantStep(months);
    const monthEnds = parts.every((p) => p.day === daysIn(p.year, p.month));
    if (monthStep !== null && monthStep !== 0 && (sameDay || monthEnds)) {
      return series((i) => {
        const index = months[0]! + i * monthStep;
        const year = Math.floor(index / 12);
        const month = mod(index, 12) + 1;
        const last = daysIn(year, month);
        return dateToSerial(year, month, monthEnds && !sameDay ? last : Math.min(first.day, last), system);
      });
    }
  }
  // Equal whole-day steps.
  const dayStep = constantStep(values);
  if (dayStep !== null && Number.isInteger(dayStep)) {
    return series((i) => values[0]! + i * dayStep);
  }
  const [intercept, slope] = trend(values);
  return series((i) => round15(intercept + slope * i));
}

function timeSeries(values: readonly number[], seed: readonly SeriesSeed[], axis: "row" | "col"): Series {
  if (values.length === 1) {
    const start = values[0]!;
    return { kind: "time", byDefault: true, at: (i) => ({ value: round15(start + i / 24) }) };
  }
  const [intercept, slope] = trend(values);
  return {
    kind: "time",
    byDefault: true,
    at: (i) => (i >= 0 && i < seed.length ? repeatSeries(seed, axis).at(i) : { value: round15(intercept + slope * i) }),
  };
}

// --- text series ------------------------------------------------------------------------

type Casing = "upper" | "lower" | "title" | "as-is";

function casingOf(text: string): Casing {
  if (text.toUpperCase() === text && text.toLowerCase() !== text) {
    return "upper";
  }
  if (text.toLowerCase() === text && text.toUpperCase() !== text) {
    return "lower";
  }
  const head = text.slice(0, 1);
  const rest = text.slice(1);
  return head.toUpperCase() === head && rest.toLowerCase() === rest ? "title" : "as-is";
}

function applyCasing(text: string, casing: Casing): string {
  switch (casing) {
    case "upper":
      return text.toUpperCase();
    case "lower":
      return text.toLowerCase();
    case "title":
      return text.slice(0, 1).toUpperCase() + text.slice(1).toLowerCase();
    default:
      return text;
  }
}

const ENGLISH = {
  daysShort: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
  daysLong: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
  monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  monthsLong: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
};

/** The name cycles a seed may follow, in lookup order: English short and
 * long weekdays and months, then the locale's. */
function nameLists(locale: SpreadsheetLocale): (readonly string[])[] {
  const names = localeNames(locale.tag);
  return [ENGLISH.daysShort, ENGLISH.daysLong, ENGLISH.monthsShort, ENGLISH.monthsLong, names.daysShort, names.daysLong, names.monthsShort, names.monthsLong];
}

/** A text series over integer indices with a constant step; null when the
 * steps differ (or every seed is the same). */
function steppedSeries(indices: readonly number[], kind: SeriesKind, cycle: number | null, render: (index: number) => string): Series | null {
  let step = 1;
  if (indices.length > 1) {
    const steps = indices.slice(1).map((index, k) => (cycle === null ? index - indices[k]! : mod(index - indices[k]!, cycle)));
    if (steps.some((each) => each !== steps[0]) || steps[0] === 0) {
      return null;
    }
    step = steps[0]!;
  }
  return { kind, byDefault: true, at: (i) => ({ value: render(indices[0]! + i * step) }) };
}

function nameSeries(texts: readonly string[], locale: SpreadsheetLocale): Series | null {
  for (const list of nameLists(locale)) {
    const lower = list.map((name) => name.toLowerCase());
    const indices = texts.map((text) => lower.indexOf(text.toLowerCase()));
    if (indices.every((index) => index >= 0)) {
      const casing = casingOf(texts[0]!);
      const own = casing === "title" && list[indices[0]!] === texts[0] ? "as-is" : casing;
      return steppedSeries(indices, "name", list.length, (index) => applyCasing(list[mod(index, list.length)]!, own));
    }
  }
  return null;
}

const QUARTER = /^(Q|Qtr|Quarter)(\s*)([1-4])(?:(\s+)(\d{1,4}))?$/i;

function quarterSeries(texts: readonly string[]): Series | null {
  const matches = texts.map((text) => QUARTER.exec(text));
  const first = matches[0];
  if (!first || matches.some((m) => !m || m[1]!.toLowerCase() !== first[1]!.toLowerCase() || (m[5] === undefined) !== (first[5] === undefined))) {
    return null;
  }
  const withYear = first[5] !== undefined;
  const indices = matches.map((m) => (withYear ? Number(m![5]) * 4 : 0) + Number(m![3]) - 1);
  return steppedSeries(indices, "quarter", withYear ? null : 4, (index) =>
    withYear ? `${first[1]}${first[2]}${mod(index, 4) + 1}${first[4]}${Math.floor(index / 4)}` : `${first[1]}${first[2]}${mod(index, 4) + 1}`,
  );
}

const ORDINAL = /^(\d+)(st|nd|rd|th)$/i;

function ordinalSuffix(n: number): string {
  const tens = n % 100;
  if (tens >= 11 && tens <= 13) {
    return "th";
  }
  return ["th", "st", "nd", "rd"][n % 10] ?? "th";
}

function ordinalSeries(texts: readonly string[]): Series | null {
  const matches = texts.map((text) => ORDINAL.exec(text));
  if (matches.some((m) => !m)) {
    return null;
  }
  const upper = matches[0]![2]! === matches[0]![2]!.toUpperCase();
  return steppedSeries(
    matches.map((m) => Number(m![1])),
    "ordinal",
    null,
    (n) => `${Math.abs(n)}${upper ? ordinalSuffix(Math.abs(n)).toUpperCase() : ordinalSuffix(Math.abs(n))}`,
  );
}

const COUNTER = /^([\s\S]*\D)(\d+)$/;

function counterSeries(texts: readonly string[]): Series | null {
  const matches = texts.map((text) => COUNTER.exec(text));
  const first = matches[0];
  if (!first || matches.some((m) => !m || m[1] !== first[1])) {
    return null;
  }
  const digits = first[2]!;
  const width = digits.length > 1 && digits.startsWith("0") ? digits.length : 0;
  return steppedSeries(
    matches.map((m) => Number(m![2])),
    "counter",
    null,
    (n) => `${first[1]}${String(Math.abs(n)).padStart(width, "0")}`,
  );
}

/** The series a seed line continues (see the module header). */
export function detectSeries(seed: readonly SeriesSeed[], locale: SpreadsheetLocale, options: SeriesOptions = {}): Series {
  const axis = options.axis ?? "row";
  if (seed.length === 0 || seed.some((each) => each.formula !== undefined)) {
    return repeatSeries(seed.length === 0 ? [{ value: null }] : seed, axis);
  }
  const values = seed.map((each) => each.value);
  if (values.every((value) => typeof value === "number")) {
    const numbers = values as number[];
    const kinds = seed.map((each) => dateKindOf(each.code));
    if (kinds.every((kind) => kind === "date")) {
      return dateSeries(numbers, seed, axis, options.dateSystem ?? "1900");
    }
    if (kinds.every((kind) => kind === "time")) {
      return timeSeries(numbers, seed, axis);
    }
    return numberSeries(numbers, seed, axis);
  }
  if (values.every((value) => typeof value === "string" && value.length > 0)) {
    const texts = values as string[];
    const series = nameSeries(texts, locale) ?? quarterSeries(texts) ?? ordinalSeries(texts) ?? counterSeries(texts);
    if (series) {
      return series;
    }
  }
  return repeatSeries(seed, axis);
}
