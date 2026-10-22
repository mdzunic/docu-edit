// Date & Time (SPEC-120 "The core functions"): DATE TIME YEAR MONTH DAY HOUR
// MINUTE SECOND TODAY NOW WEEKDAY EDATE EOMONTH DAYS — in both date systems
// over SPEC-116's serial.ts.
//
// - In the 1900 system February 1900 has 29 days in every date function
//   (`=DATE(1900,2,29)` and `=EOMONTH(DATE(1900,2,1),0)` are 60) — except in
//   an `.ods` workbook (`dateQuirks: "proleptic"`, as SPEC-118 renders it),
//   where it has 28 and no function produces serial 60.
// - DATE adds 1900 to years 0–1899, carries month and day overflow, and is
//   #NUM! for a year outside 0–9999 or a negative serial. TIME truncates its
//   arguments and is the fraction of a day modulo 1 (#NUM! if negative).
//   HOUR/MINUTE/SECOND read the time of day rounded to the nearest second.
// - WEEKDAY is Excel's arithmetic, (⌊serial⌋ + 6) mod 7 in the 1900 system
//   (`WEEKDAY(1)` is 1, Sunday) and (⌊serial⌋ + 5) mod 7 in 1904, for return
//   types 1, 2, 3 and 11–17.
// - EDATE and EOMONTH clamp the day to the target month's end; DAYS is
//   INT(end) − INT(start).
// - TODAY and NOW read ctx.now() — core/clock.ts's now(), never Date.now() —
//   so a test pins them with setClock; NOW is local time to the millisecond.
import type { CellError } from "../../model";
import { civilFromDays, dateToSerial, daysFromCivil, serialToParts } from "../../serial";
import type { EvalContext, EvalSettings } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { isError, NUM_ERROR } from "../values";
import { arg, integerArg, numberArg, numberOr, optional } from "./common";

export type Dates = Pick<EvalSettings, "dateSystem" | "dateQuirks">;

/** The last serial: 9999-12-31. */
export function maxSerial(settings: Dates): number {
  return dateToSerial(9999, 12, 31, settings.dateSystem);
}

/** Whether February 1900 has Excel's fictitious 29th. */
export function lotus(settings: Dates): boolean {
  return settings.dateSystem === "1900" && settings.dateQuirks === "excel";
}

const DAY_1899_12_31 = daysFromCivil(1899, 12, 31);

/** The serial of a calendar date whose month and day may overflow, in the
 * workbook's calendar. */
export function serialOf(year: number, month: number, day: number, settings: Dates): number {
  const months = year * 12 + (month - 1);
  const y = Math.floor(months / 12);
  const m = months - y * 12 + 1;
  if (settings.dateSystem === "1900" && !lotus(settings)) {
    // Proleptic: real days from 1899-12-31, the fictitious serial 60 skipped.
    const days = daysFromCivil(y, m, day) - DAY_1899_12_31;
    return days >= 60 ? days + 1 : days;
  }
  return dateToSerial(y, m, 1, settings.dateSystem) + day - 1;
}

/** The calendar date of a whole-day serial (Excel's 1900-01-00 for 0). */
export function dateOf(serial: number, settings: Dates): { year: number; month: number; day: number } {
  const days = Math.floor(serial);
  if (lotus(settings) && days === 0) {
    return { year: 1900, month: 1, day: 0 };
  }
  if (settings.dateSystem === "1900" && !lotus(settings) && days >= 60) {
    return civilFromDays(DAY_1899_12_31 + days - 1);
  }
  const parts = serialToParts(days, settings.dateSystem);
  return { year: parts.year, month: parts.month, day: parts.day };
}

/** Days in a month in the workbook's calendar. */
export function monthLength(year: number, month: number, settings: Dates): number {
  if (month === 2 && year === 1900) {
    return lotus(settings) ? 29 : 28;
  }
  return civilFromDays(daysFromCivil(year, month + 1, 1) - 1).day;
}

/** A serial argument: a number (text coerced) from 0 to 9999-12-31. */
export function serialArg(args: ArgList, index: number, ctx: EvalContext): number | CellError {
  const serial = numberArg(args, index, ctx);
  if (isError(serial)) {
    return serial;
  }
  return serial < 0 || serial >= maxSerial(ctx.settings) + 1 ? NUM_ERROR : serial;
}

/** A date part of a serial argument. */
function datePart(name: "YEAR" | "MONTH" | "DAY", summary: string): FunctionSpec {
  return {
    name,
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary,
    args: [arg("serial_number", "The date.")],
    impl(args, ctx) {
      const serial = serialArg(args, 0, ctx);
      if (isError(serial)) {
        return serial;
      }
      const date = dateOf(serial, ctx.settings);
      return name === "YEAR" ? date.year : name === "MONTH" ? date.month : date.day;
    },
  };
}

/** A time part of a serial argument, rounded to the nearest second. */
function timePart(name: "HOUR" | "MINUTE" | "SECOND", summary: string): FunctionSpec {
  return {
    name,
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary,
    args: [arg("serial_number", "The time.")],
    impl(args, ctx) {
      const serial = serialArg(args, 0, ctx);
      if (isError(serial)) {
        return serial;
      }
      const seconds = Math.round((serial - Math.floor(serial)) * 86_400) % 86_400;
      return name === "HOUR" ? Math.floor(seconds / 3600) : name === "MINUTE" ? Math.floor(seconds / 60) % 60 : seconds % 60;
    },
  };
}

/** EDATE/EOMONTH: a start date moved by whole months. */
function monthShift(name: "EDATE" | "EOMONTH", summary: string): FunctionSpec {
  return {
    name,
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary,
    args: [arg("start_date", "The date to start from."), arg("months", "How many months later (negative: earlier).")],
    impl(args, ctx) {
      const start = serialArg(args, 0, ctx);
      if (isError(start)) {
        return start;
      }
      const months = integerArg(args, 1, ctx);
      if (isError(months)) {
        return months;
      }
      const date = dateOf(start, ctx.settings);
      const total = date.year * 12 + (date.month - 1) + months;
      const year = Math.floor(total / 12);
      const month = total - year * 12 + 1;
      if (year < 1900 || year > 9999) {
        return NUM_ERROR;
      }
      const last = monthLength(year, month, ctx.settings);
      const serial = serialOf(year, month, name === "EOMONTH" ? last : Math.min(date.day, last), ctx.settings);
      return serial < 0 ? NUM_ERROR : serial;
    },
  };
}

/** The day of the week of a serial, 0 = Sunday — Excel's arithmetic (see
 * the module header). */
export function weekdayOf(serial: number, settings: Dates): number {
  const days = Math.floor(serial);
  if (settings.dateSystem === "1904") {
    return (days + 5) % 7;
  }
  if (lotus(settings) || days >= 60) {
    return (days + 6) % 7;
  }
  return days % 7; // proleptic before March 1900: 1899-12-31 (0) was a Sunday
}

/** The workbook serial of a local date (`now`'s calendar day). */
function todaySerial(now: Date, settings: Dates): number {
  return serialOf(now.getFullYear(), now.getMonth() + 1, now.getDate(), settings);
}

export const DATE_TIME_FUNCTIONS: FunctionSpec[] = [
  {
    name: "DATE",
    category: "Date & Time",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns the serial number of a date.",
    args: [arg("year", "The year; 0–1899 counts from 1900."), arg("month", "The month; overflow carries into the year."), arg("day", "The day; overflow carries into the month.")],
    impl(args, ctx) {
      const parts: number[] = [];
      for (let i = 0; i < 3; i += 1) {
        const n = integerArg(args, i, ctx);
        if (isError(n)) {
          return n;
        }
        parts.push(n);
      }
      let [year, month, day] = parts as [number, number, number];
      if (year < 0 || year > 9999) {
        return NUM_ERROR;
      }
      if (year < 1900) {
        year += 1900;
      }
      const serial = serialOf(year, month, day, ctx.settings);
      return serial < 0 || serial > maxSerial(ctx.settings) ? NUM_ERROR : serial;
    },
  },
  {
    name: "TIME",
    category: "Date & Time",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns the serial number of a time of day.",
    args: [arg("hour", "The hour."), arg("minute", "The minute."), arg("second", "The second.")],
    impl(args, ctx) {
      let seconds = 0;
      for (const [index, scale] of [3600, 60, 1].entries()) {
        const n = integerArg(args, index, ctx);
        if (isError(n)) {
          return n;
        }
        seconds += n * scale;
      }
      return seconds < 0 ? NUM_ERROR : (seconds % 86_400) / 86_400;
    },
  },
  datePart("YEAR", "Returns the year of a date."),
  datePart("MONTH", "Returns the month of a date, 1 to 12."),
  datePart("DAY", "Returns the day of the month of a date."),
  timePart("HOUR", "Returns the hour of a time, 0 to 23."),
  timePart("MINUTE", "Returns the minute of a time, 0 to 59."),
  timePart("SECOND", "Returns the second of a time, 0 to 59."),
  {
    name: "TODAY",
    category: "Date & Time",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    volatile: true,
    summary: "Returns today's date.",
    args: [],
    impl(_args, ctx) {
      return todaySerial(ctx.now(), ctx.settings);
    },
  },
  {
    name: "NOW",
    category: "Date & Time",
    minArgs: 0,
    maxArgs: 0,
    argKinds: [],
    volatile: true,
    summary: "Returns the current date and time.",
    args: [],
    impl(_args, ctx) {
      const now = ctx.now();
      const ms = ((now.getHours() * 60 + now.getMinutes()) * 60 + now.getSeconds()) * 1000 + now.getMilliseconds();
      return todaySerial(now, ctx.settings) + ms / 86_400_000;
    },
  },
  {
    name: "WEEKDAY",
    category: "Date & Time",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the day of the week of a date as a number.",
    args: [arg("serial_number", "The date."), optional("return_type", "1: Sunday is 1 (default); 2: Monday is 1; 3: Monday is 0; 11–17: the week starts Monday…Sunday.")],
    impl(args, ctx) {
      const serial = serialArg(args, 0, ctx);
      if (isError(serial)) {
        return serial;
      }
      const type = numberOr(args, 1, ctx, 1);
      if (isError(type)) {
        return type;
      }
      const sunday0 = weekdayOf(serial, ctx.settings);
      const kind = Math.trunc(type);
      if (kind === 1 || kind === 17) {
        return sunday0 + 1;
      }
      if (kind === 2) {
        return ((sunday0 + 6) % 7) + 1;
      }
      if (kind === 3) {
        return (sunday0 + 6) % 7;
      }
      if (kind >= 11 && kind <= 16) {
        const first = kind - 10; // 11 → Monday (1) … 16 → Saturday (6)
        return ((sunday0 - first + 7) % 7) + 1;
      }
      return NUM_ERROR;
    },
  },
  monthShift("EDATE", "Returns the date a number of months before or after a date."),
  monthShift("EOMONTH", "Returns the last day of the month a number of months before or after a date."),
  {
    name: "DAYS",
    category: "Date & Time",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the number of days between two dates.",
    args: [arg("end_date", "The later date."), arg("start_date", "The earlier date.")],
    impl(args, ctx) {
      const end = serialArg(args, 0, ctx);
      if (isError(end)) {
        return end;
      }
      const start = serialArg(args, 1, ctx);
      return isError(start) ? start : Math.floor(end) - Math.floor(start);
    },
  },
];
