// Date & Time, SPEC-127's library: DATEVALUE TIMEVALUE WEEKNUM ISOWEEKNUM
// WORKDAY WORKDAY.INTL NETWORKDAYS NETWORKDAYS.INTL YEARFRAC DATEDIF DAYS360,
// over dateTime.ts's calendar (Excel's fictitious 1900-02-29 in a 1900-system
// .xlsx, the proleptic calendar in an .ods).
//
// - DATEVALUE/TIMEVALUE parse their text through SPEC-122's parseInput in
//   the user's locale and keep the integer / fractional part; a text that is
//   no date or time is #VALUE!.
// - WEEKNUM: return types 1/17 (weeks from Sunday), 2/11 (Monday), 12–16
//   (Tuesday…Saturday) number the week holding 1 January as week 1; 21 is
//   ISOWEEKNUM, ISO 8601's week (the week of the year's first Thursday).
// - WORKDAY(.INTL) and NETWORKDAYS(.INTL) skip weekends and holidays.
//   Weekend codes: 1 Sat–Sun, 2 Sun–Mon, … 7 Fri–Sat, 11 Sun only, … 17 Sat
//   only, or a Monday-first mask of seven 0/1 ("0000011"); a mask of all
//   1s is #VALUE!, any other code #NUM!.
// - YEARFRAC: basis 0 is OpenFormula's US 30/360 (both dates the last of
//   February: the end becomes the 30th; the start the last of February: the
//   30th; an end on the 31st after a start on the 30th or 31st: the 30th; a
//   start on the 31st: the 30th); 1 actual/actual — a period of at most a
//   year divides by 366 when it lies in one leap year or spans a 29 February,
//   else 365, and a longer one by the average length of the calendar years
//   it touches; 2 actual/360; 3 actual/365; 4 European 30/360.
// - DATEDIF: units Y, M, D, MD, YM, YD (any case); a start after the end is
//   #NUM!. MD is computed as Excel does it — the day difference, or, when
//   the end's day is earlier, the end minus the start's day in the month
//   before the end (overflowing as DATE does, so 31 Jan → 1 Mar 2021 is −2).
// - DAYS360 FALSE (US, Microsoft's page): a start on the last day of its
//   month becomes the 30th; an end on the last day of its month becomes the
//   1st of the next month when the start is before the 30th, else the 30th.
//   TRUE (European): any 31st becomes the 30th.
import type { CellError } from "../../model";
import { parseInput } from "../../numberFormat/input";
import { forEachScalar } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { isEmpty, isError, NUM_ERROR, toNumber, VALUE_ERROR } from "../values";
import { arg, logicalArg, numberOr, optional, scalarArg, textArg } from "./common";
import { dateOf, maxSerial, monthLength, serialArg, serialOf, weekdayOf, type Dates } from "./dateTime";

/** A date/time text's serial through parseInput, or #VALUE!. */
function parsedDateTime(args: ArgList, ctx: EvalContext): number | CellError {
  const value = scalarArg(args, 0);
  if (typeof value !== "string") {
    return isError(value) ? value : VALUE_ERROR;
  }
  const text = textArg(args, 0);
  if (isError(text)) {
    return text;
  }
  const parsed = parseInput(text, ctx.settings.locale, ctx.settings.dateSystem);
  // A date or time is an input parseInput suggests a date or time format for
  // ("m/d/yyyy", "d-mmm-yy", "h:mm AM/PM"): one with a d, m, y, h or s
  // outside its quoted text, escapes and [brackets] — which no number,
  // percent, currency or fraction format has.
  const dateTimeFormat = /[dmyhs]/i.test((parsed.numberFormat ?? "").replace(/"[^"]*"|\\.|\[[^\]]*\]/g, ""));
  return typeof parsed.value === "number" && dateTimeFormat && parsed.formula === undefined ? parsed.value : VALUE_ERROR;
}

/** The Gregorian leap rule. */
function isLeap(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

/** Whether a calendar date is the last day of its month. */
function lastOfMonth(date: { year: number; month: number; day: number }, settings: Dates): boolean {
  return date.day === monthLength(date.year, date.month, settings);
}

/** ISO 8601's week number of a serial. */
function isoWeek(serial: number, settings: Dates): number {
  const day = Math.floor(serial);
  const monday0 = (weekdayOf(day, settings) + 6) % 7;
  const thursday = day - monday0 + 3;
  const year = dateOf(thursday, settings).year;
  return Math.floor((thursday - serialOf(year, 1, 1, settings)) / 7) + 1;
}

/** WEEKNUM's first day of the week by return type (0 = Sunday); 21 → ISO. */
const WEEK_START: ReadonlyMap<number, number> = new Map([
  [1, 0],
  [2, 1],
  [11, 1],
  [12, 2],
  [13, 3],
  [14, 4],
  [15, 5],
  [16, 6],
  [17, 0],
]);

/** The weekend days, indexed by weekday (0 = Sunday). */
type Weekend = readonly boolean[];

/** The weekend of a WORKDAY.INTL/NETWORKDAYS.INTL code (see the header). */
function weekendOf(args: ArgList, index: number, ctx: EvalContext): Weekend | CellError {
  if (args.length <= index || args.isMissing(index)) {
    return [true, false, false, false, false, false, true];
  }
  const value = scalarArg(args, index);
  if (typeof value === "string") {
    if (!/^[01]{7}$/.test(value)) {
      return VALUE_ERROR;
    }
    if (value === "1111111") {
      return VALUE_ERROR;
    }
    // Monday first: mask[0] is Monday … mask[6] Sunday.
    return [value[6] === "1", ...Array.from(value.slice(0, 6), (bit) => bit === "1")];
  }
  const code = toNumber(value, ctx.settings);
  if (isError(code)) {
    return code;
  }
  const n = Math.trunc(code);
  const weekend = [false, false, false, false, false, false, false];
  if (n >= 1 && n <= 7) {
    // 1: Saturday and Sunday; each code moves the pair one day on.
    weekend[(n + 5) % 7] = true;
    weekend[(n + 6) % 7] = true;
    return weekend;
  }
  if (n >= 11 && n <= 17) {
    weekend[(n - 11) % 7] = true;
    return weekend;
  }
  return NUM_ERROR;
}

/** The whole-day holidays a WORKDAY/NETWORKDAYS argument lists. */
function holidaysOf(args: ArgList, index: number, ctx: EvalContext): Set<number> | CellError {
  const days = new Set<number>();
  if (args.length <= index || args.isMissing(index)) {
    return days;
  }
  let failure: CellError | undefined;
  forEachScalar(ctx, args.get(index), (value) => {
    if (isEmpty(value)) {
      return true;
    }
    const n = typeof value === "boolean" ? VALUE_ERROR : toNumber(value, ctx.settings);
    if (isError(n)) {
      failure = n;
      return false;
    }
    if (n < 0) {
      failure = NUM_ERROR;
      return false;
    }
    days.add(Math.floor(n));
    return true;
  });
  return failure ?? days;
}

/** WORKDAY's date `days` working days from `start`. */
function workday(start: number, days: number, weekend: Weekend, holidays: ReadonlySet<number>, settings: Dates): number | CellError {
  const last = maxSerial(settings);
  let day = start;
  let left = Math.abs(days);
  const step = days < 0 ? -1 : 1;
  while (left > 0) {
    day += step;
    if (day < 0 || day > last) {
      return NUM_ERROR;
    }
    if (!weekend[weekdayOf(day, settings)] && !holidays.has(day)) {
      left -= 1;
    }
  }
  return day;
}

/** NETWORKDAYS's count of working days from `start` to `end`, both
 * included; negative when the end is before the start. */
function networkdays(start: number, end: number, weekend: Weekend, holidays: ReadonlySet<number>, settings: Dates): number {
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  const span = to - from + 1;
  const perWeek = weekend.filter((off) => !off).length;
  let count = Math.floor(span / 7) * perWeek;
  for (let day = from + Math.floor(span / 7) * 7; day <= to; day += 1) {
    if (!weekend[weekdayOf(day, settings)]) {
      count += 1;
    }
  }
  for (const holiday of holidays) {
    if (holiday >= from && holiday <= to && !weekend[weekdayOf(holiday, settings)]) {
      count -= 1;
    }
  }
  return start > end ? -count : count;
}

/** The working-day functions' shared arguments: start, the count or end,
 * the weekend and the holidays. */
function workdayArguments(
  args: ArgList,
  ctx: EvalContext,
  intl: boolean,
): { start: number; other: number; weekend: Weekend; holidays: Set<number> } | CellError {
  const start = serialArg(args, 0, ctx);
  if (isError(start)) {
    return start;
  }
  const other = numberOr(args, 1, ctx, 0);
  if (isError(other)) {
    return other;
  }
  const weekend = intl ? weekendOf(args, 2, ctx) : weekendOf(args, 99, ctx);
  if (isError(weekend)) {
    return weekend;
  }
  const holidays = holidaysOf(args, intl ? 3 : 2, ctx);
  if (isError(holidays)) {
    return holidays;
  }
  return { start: Math.floor(start), other, weekend, holidays };
}

/** 30/360 day count from the adjusted day numbers. */
function days360(a: { year: number; month: number }, d1: number, b: { year: number; month: number }, d2: number): number {
  return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1);
}

/** YEARFRAC's fraction for start ≤ end (see the module header). */
function yearFraction(start: number, end: number, basis: number, settings: Dates): number {
  const a = dateOf(start, settings);
  const b = dateOf(end, settings);
  const days = end - start;
  switch (basis) {
    case 0: {
      let d1 = a.day;
      let d2 = b.day;
      const aFebEnd = a.month === 2 && lastOfMonth(a, settings);
      const bFebEnd = b.month === 2 && lastOfMonth(b, settings);
      if (aFebEnd && bFebEnd) {
        d2 = 30;
      }
      if (aFebEnd) {
        d1 = 30;
      }
      if (d2 === 31 && d1 >= 30) {
        d2 = 30;
      }
      if (d1 === 31) {
        d1 = 30;
      }
      return days360(a, d1, b, d2) / 360;
    }
    case 1: {
      const withinAYear = a.year === b.year || (b.year === a.year + 1 && (a.month > b.month || (a.month === b.month && a.day >= b.day)));
      if (withinAYear) {
        const spansLeapDay =
          (isLeap(a.year) && start <= serialOf(a.year, 2, 29, settings) && serialOf(a.year, 2, 29, settings) <= end) ||
          (isLeap(b.year) && start <= serialOf(b.year, 2, 29, settings) && serialOf(b.year, 2, 29, settings) <= end);
        const leap = (a.year === b.year && isLeap(a.year)) || spansLeapDay;
        return days / (leap ? 366 : 365);
      }
      const years = b.year - a.year + 1;
      const total = serialOf(b.year + 1, 1, 1, settings) - serialOf(a.year, 1, 1, settings);
      return days / (total / years);
    }
    case 2:
      return days / 360;
    case 3:
      return days / 365;
    default:
      return days360(a, Math.min(a.day, 30), b, Math.min(b.day, 30)) / 360;
  }
}

/** DATEDIF in a unit (see the module header), start ≤ end. */
function dateDif(start: number, end: number, unit: string, settings: Dates): number | CellError {
  const a = dateOf(start, settings);
  const b = dateOf(end, settings);
  const earlierDay = b.day < a.day;
  const months = (b.year - a.year) * 12 + (b.month - a.month) - (earlierDay ? 1 : 0);
  switch (unit) {
    case "Y":
      return Math.floor(months / 12);
    case "M":
      return months;
    case "D":
      return end - start;
    case "YM":
      return months % 12;
    case "MD":
      return earlierDay ? end - serialOf(b.year, b.month - 1, a.day, settings) : b.day - a.day;
    case "YD": {
      let anniversary = serialOf(a.year, b.month, b.day, settings);
      if (anniversary < start) {
        anniversary = serialOf(a.year + 1, b.month, b.day, settings);
      }
      return anniversary - start;
    }
    default:
      return NUM_ERROR;
  }
}

/** Two serial arguments, truncated to whole days. */
function twoDates(args: ArgList, ctx: EvalContext): [number, number] | CellError {
  const start = serialArg(args, 0, ctx);
  if (isError(start)) {
    return start;
  }
  const end = serialArg(args, 1, ctx);
  return isError(end) ? end : [Math.floor(start), Math.floor(end)];
}

const HOLIDAYS = optional("holidays", "Dates to skip as well, as a range or array.");
const WEEKEND = optional("weekend", "1 Sat–Sun (default) … 7, 11–17 one day, or a Monday-first mask such as \"0000011\".");

export const DATE_EXTRA_FUNCTIONS: FunctionSpec[] = [
  {
    name: "DATEVALUE",
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Converts a date written as text to its serial number.",
    args: [arg("date_text", "The date as text, such as \"10/8/2026\" or \"8-Oct-2026\".")],
    impl(args, ctx) {
      const serial = parsedDateTime(args, ctx);
      return isError(serial) ? serial : Math.floor(serial);
    },
  },
  {
    name: "TIMEVALUE",
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Converts a time written as text to a fraction of a day.",
    args: [arg("time_text", "The time as text, such as \"6:45 PM\" or \"18:45\".")],
    impl(args, ctx) {
      const serial = parsedDateTime(args, ctx);
      // To the millisecond, the finest a time text states: a date's serial
      // beside it would otherwise leak its rounding into the fraction.
      return isError(serial) ? serial : (Math.round((serial - Math.floor(serial)) * 86_400_000) % 86_400_000) / 86_400_000;
    },
  },
  {
    name: "WEEKNUM",
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the week number of a date in its year.",
    args: [
      arg("serial_number", "The date."),
      optional("return_type", "1 weeks start Sunday (default); 2 or 11 Monday; 12–17 Tuesday to Sunday; 21 ISO 8601 weeks."),
    ],
    impl(args, ctx) {
      const serial = serialArg(args, 0, ctx);
      if (isError(serial)) {
        return serial;
      }
      const type = numberOr(args, 1, ctx, 1);
      if (isError(type)) {
        return type;
      }
      const kind = Math.trunc(type);
      if (kind === 21) {
        return isoWeek(serial, ctx.settings);
      }
      const first = WEEK_START.get(kind);
      if (first === undefined) {
        return NUM_ERROR;
      }
      const day = Math.floor(serial);
      const january1 = serialOf(dateOf(day, ctx.settings).year, 1, 1, ctx.settings);
      const lead = (weekdayOf(january1, ctx.settings) - first + 7) % 7;
      return Math.floor((day - january1 + lead) / 7) + 1;
    },
  },
  {
    name: "ISOWEEKNUM",
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns the ISO 8601 week number of a date.",
    args: [arg("date", "The date.")],
    impl(args, ctx) {
      const serial = serialArg(args, 0, ctx);
      return isError(serial) ? serial : isoWeek(serial, ctx.settings);
    },
  },
  {
    name: "WORKDAY",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "any"],
    summary: "Returns the date a number of working days before or after a date.",
    args: [arg("start_date", "The date to count from."), arg("days", "Working days to move (negative: back)."), HOLIDAYS],
    impl(args, ctx) {
      const call = workdayArguments(args, ctx, false);
      return isError(call) ? call : workday(call.start, Math.trunc(call.other), call.weekend, call.holidays, ctx.settings);
    },
  },
  {
    name: "WORKDAY.INTL",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 4,
    argKinds: ["value", "value", "value", "any"],
    summary: "Returns the date a number of working days away, with a chosen weekend.",
    args: [arg("start_date", "The date to count from."), arg("days", "Working days to move (negative: back)."), WEEKEND, HOLIDAYS],
    impl(args, ctx) {
      const call = workdayArguments(args, ctx, true);
      return isError(call) ? call : workday(call.start, Math.trunc(call.other), call.weekend, call.holidays, ctx.settings);
    },
  },
  {
    name: "NETWORKDAYS",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "any"],
    summary: "Counts the working days between two dates, both included.",
    args: [arg("start_date", "The first date."), arg("end_date", "The last date."), HOLIDAYS],
    impl(args, ctx) {
      const call = workdayArguments(args, ctx, false);
      if (isError(call)) {
        return call;
      }
      const end = serialArg(args, 1, ctx);
      return isError(end) ? end : networkdays(call.start, Math.floor(end), call.weekend, call.holidays, ctx.settings);
    },
  },
  {
    name: "NETWORKDAYS.INTL",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 4,
    argKinds: ["value", "value", "value", "any"],
    summary: "Counts the working days between two dates, with a chosen weekend.",
    args: [arg("start_date", "The first date."), arg("end_date", "The last date."), WEEKEND, HOLIDAYS],
    impl(args, ctx) {
      const call = workdayArguments(args, ctx, true);
      if (isError(call)) {
        return call;
      }
      const end = serialArg(args, 1, ctx);
      return isError(end) ? end : networkdays(call.start, Math.floor(end), call.weekend, call.holidays, ctx.settings);
    },
  },
  {
    name: "YEARFRAC",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns the fraction of a year between two dates.",
    args: [
      arg("start_date", "The first date."),
      arg("end_date", "The last date."),
      optional("basis", "0 US 30/360 (default), 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360."),
    ],
    impl(args, ctx) {
      const dates = twoDates(args, ctx);
      if (isError(dates)) {
        return dates;
      }
      const basis = numberOr(args, 2, ctx, 0);
      if (isError(basis)) {
        return basis;
      }
      const kind = Math.trunc(basis);
      if (kind < 0 || kind > 4) {
        return NUM_ERROR;
      }
      const [start, end] = dates;
      return yearFraction(Math.min(start, end), Math.max(start, end), kind, ctx.settings);
    },
  },
  {
    name: "DATEDIF",
    category: "Date & Time",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns the years, months or days between two dates.",
    args: [
      arg("start_date", "The earlier date."),
      arg("end_date", "The later date."),
      arg("unit", "\"Y\", \"M\" or \"D\", or \"MD\", \"YM\" or \"YD\" ignoring the larger units."),
    ],
    impl(args, ctx) {
      const dates = twoDates(args, ctx);
      if (isError(dates)) {
        return dates;
      }
      const unit = textArg(args, 2);
      if (isError(unit)) {
        return unit;
      }
      const [start, end] = dates;
      return start > end ? NUM_ERROR : dateDif(start, end, unit.toUpperCase(), ctx.settings);
    },
  },
  {
    name: "DAYS360",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns the days between two dates in a 360-day year.",
    args: [
      arg("start_date", "The first date."),
      arg("end_date", "The last date."),
      optional("method", "FALSE or omitted: the US (NASD) method; TRUE: the European method."),
    ],
    impl(args, ctx) {
      const dates = twoDates(args, ctx);
      if (isError(dates)) {
        return dates;
      }
      const european = args.length > 2 && !args.isMissing(2) ? logicalArg(args, 2) : false;
      if (isError(european)) {
        return european;
      }
      const a = dateOf(dates[0], ctx.settings);
      const b = dateOf(dates[1], ctx.settings);
      if (european) {
        return days360(a, Math.min(a.day, 30), b, Math.min(b.day, 30));
      }
      const d1 = lastOfMonth(a, ctx.settings) ? 30 : a.day;
      if (!lastOfMonth(b, ctx.settings)) {
        return days360(a, d1, b, b.day);
      }
      if (d1 < 30) {
        const next = b.month === 12 ? { year: b.year + 1, month: 1 } : { year: b.year, month: b.month + 1 };
        return days360(a, d1, next, 1);
      }
      return days360(a, d1, b, 30);
    },
  },
];
