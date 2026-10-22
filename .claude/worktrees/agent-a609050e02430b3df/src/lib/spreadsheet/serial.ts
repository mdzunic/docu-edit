// Dates are serials (SPEC-116, the model's fourth invariant): a date or time
// is a number in the workbook's date system, and these are the conversions
// between that number and a calendar — the normative source every later
// spec (number formats, formulas, the writers) converts through.
//
// - 1900 system: serial 1 = 1900-01-01 … 59 = 1900-02-28, 60 = the
//   fictitious 1900-02-29 Excel inherited from Lotus, 61 = 1900-03-01, and
//   from there serial = days since 1899-12-30. Dates before 1900-01-01 (only
//   an .ods can hold them) count backwards from 1899-12-31 = 0, so the scale
//   stays monotonic: 1899-12-30 = −1.
// - 1904 system: serial 0 = 1904-01-01; serial = days since then.
// - ODS: `table:null-date` absent or 1899-12-30 → "1900"; 1904-01-01 →
//   "1904"; anything else → "1900" plus the model's `odsNullDate`, and then
//   conversions count days from that stored null date.
// - Times are fractions of a day; every conversion rounds to the nearest
//   millisecond.
//
// Calendar arithmetic is proleptic Gregorian on day numbers (no `Date`, so no
// time zone or DST ever leaks in).

const MS_PER_DAY = 86_400_000;

/** Days from 1970-01-01 to the given proleptic-Gregorian date (H. Hinnant's
 * days_from_civil). Out-of-range days/months roll over linearly. */
export function daysFromCivil(year: number, month: number, day: number): number {
  const y = month <= 2 ? year - 1 : year;
  const era = Math.floor(y / 400);
  const yoe = y - era * 400;
  const mp = (month + 9) % 12;
  const doy = Math.floor((153 * mp + 2) / 5) + day - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146_097 + doe - 719_468;
}

/** The inverse of daysFromCivil. */
export function civilFromDays(days: number): { year: number; month: number; day: number } {
  const z = days + 719_468;
  const era = Math.floor(z / 146_097);
  const doe = z - era * 146_097;
  const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36_524) - Math.floor(doe / 146_096)) / 365);
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const day = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const month = mp < 10 ? mp + 3 : mp - 9;
  return { year: yoe + era * 400 + (month <= 2 ? 1 : 0), month, day };
}

/** Day 0 of the 1900 scale's backward count (1900-01-01 is serial 1). */
const DAY_1899_12_31 = daysFromCivil(1899, 12, 31);
/** Serial 0 of the 1904 system. */
const DAY_1904_01_01 = daysFromCivil(1904, 1, 1);
/** The fictitious 1900-02-29. */
const LOTUS_LEAP_DAY = 60;

/** The serial of a calendar date (midnight) in `system`. 1900-02-29 — a day
 * that never was — is serial 60 in the 1900 system, as in Excel. */
export function dateToSerial(y: number, m: number, d: number, system: "1900" | "1904"): number {
  if (system === "1904") {
    return daysFromCivil(y, m, d) - DAY_1904_01_01;
  }
  if (y === 1900 && m === 2 && d === 29) {
    return LOTUS_LEAP_DAY;
  }
  const days = daysFromCivil(y, m, d) - DAY_1899_12_31;
  return days >= LOTUS_LEAP_DAY ? days + 1 : days;
}

/** Splits a serial into whole days and the millisecond of the day, rounding
 * to the nearest millisecond (a rounding that reaches midnight carries into
 * the next day). */
function splitSerial(serial: number): { days: number; msOfDay: number } {
  const totalMs = Math.round(serial * MS_PER_DAY);
  const days = Math.floor(totalMs / MS_PER_DAY);
  return { days, msOfDay: totalMs - days * MS_PER_DAY };
}

/** The calendar date and time a serial stands for in `system`. */
export function serialToParts(
  serial: number,
  system: "1900" | "1904",
): { year: number; month: number; day: number; hour: number; minute: number; second: number; ms: number } {
  const { days, msOfDay } = splitSerial(serial);
  let date: { year: number; month: number; day: number };
  if (system === "1904") {
    date = civilFromDays(DAY_1904_01_01 + days);
  } else if (days === LOTUS_LEAP_DAY) {
    date = { year: 1900, month: 2, day: 29 };
  } else {
    date = civilFromDays(DAY_1899_12_31 + (days > LOTUS_LEAP_DAY ? days - 1 : days));
  }
  return {
    ...date,
    hour: Math.floor(msOfDay / 3_600_000),
    minute: Math.floor(msOfDay / 60_000) % 60,
    second: Math.floor(msOfDay / 1000) % 60,
    ms: msOfDay % 1000,
  };
}

const ISO_DATE_TIME = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/;

/** Parses an ISO 8601 date or date-time into its calendar date and the
 * millisecond of the day. A zone designator is accepted and ignored (a cell
 * holds local wall-clock time). Throws RangeError for anything else. */
function parseIsoDateTime(iso: string): { year: number; month: number; day: number; msOfDay: number } {
  const match = ISO_DATE_TIME.exec(iso.trim());
  if (!match) {
    throw new RangeError(`Not an ISO 8601 date: ${iso}`);
  }
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4] ?? 0);
  const minute = Number(match[5] ?? 0);
  const second = Number(match[6] ?? 0);
  if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 24 || minute > 59 || second > 60) {
    throw new RangeError(`Not an ISO 8601 date: ${iso}`);
  }
  const fraction = match[7] ? Number(match[7]) : 0;
  const msOfDay = Math.round(((hour * 60 + minute) * 60 + second + fraction) * 1000);
  return { year: Number(match[1]), month, day, msOfDay };
}

/** The day number a custom ODS null date (an ISO date) counts from. */
function nullDateDays(nullDate: string): number {
  const parsed = parseIsoDateTime(nullDate);
  return daysFromCivil(parsed.year, parsed.month, parsed.day);
}

/** An ODS `office:date-value` (a date, optionally with a time —
 * `2026-10-08T13:45:00.5`) as a serial: in `system`, or counted from
 * `nullDate` when the file declared a custom one. */
export function odsDateValueToSerial(iso: string, system: "1900" | "1904", nullDate?: string): number {
  const { year, month, day, msOfDay } = parseIsoDateTime(iso);
  const days =
    nullDate !== undefined
      ? daysFromCivil(year, month, day) - nullDateDays(nullDate)
      : dateToSerial(year, month, day, system);
  return days + msOfDay / MS_PER_DAY;
}

const ISO_DURATION = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/;

/** An ODS `office:time-value` — an ISO 8601 duration, which may exceed 24
 * hours (`PT36H` → 1.5) — as a day fraction. Throws RangeError for anything
 * that is not a day/time duration. */
export function odsTimeValueToSerial(duration: string): number {
  const text = duration.trim();
  const match = ISO_DURATION.exec(text);
  if (!match || text === "P" || text.endsWith("T") || !/\d/.test(text)) {
    throw new RangeError(`Not an ISO 8601 duration: ${duration}`);
  }
  const days = Number(match[2] ?? 0);
  const hours = Number(match[3] ?? 0);
  const minutes = Number(match[4] ?? 0);
  const seconds = Number(match[5] ?? 0);
  const ms = Math.round((((days * 24 + hours) * 60 + minutes) * 60 + seconds) * 1000);
  return (match[1] ? -ms : ms) / MS_PER_DAY;
}

const pad = (value: number, width = 2): string => String(value).padStart(width, "0");

/** ".5" for 500 ms, "" for none — fractional seconds with trailing zeros
 * removed, the way the ODS examples write them. */
function fractionText(ms: number): string {
  return ms === 0 ? "" : `.${pad(ms, 3).replace(/0+$/, "")}`;
}

/** A serial as an ODS `office:date-value`: the date alone at midnight, else
 * date and time. The fictitious 1900-02-29 has no calendar date; it is
 * written as 1900-02-28, the day it repeats. */
export function serialToOdsDateValue(serial: number, system: "1900" | "1904", nullDate?: string): string {
  const { days, msOfDay } = splitSerial(serial);
  let date: { year: number; month: number; day: number };
  if (nullDate !== undefined) {
    date = civilFromDays(nullDateDays(nullDate) + days);
  } else if (system === "1904") {
    date = civilFromDays(DAY_1904_01_01 + days);
  } else {
    date = civilFromDays(DAY_1899_12_31 + (days >= LOTUS_LEAP_DAY ? days - 1 : days));
  }
  const year = date.year < 0 ? `-${pad(-date.year, 4)}` : pad(date.year, 4);
  const dateText = `${year}-${pad(date.month)}-${pad(date.day)}`;
  if (msOfDay === 0) {
    return dateText;
  }
  const hour = Math.floor(msOfDay / 3_600_000);
  const minute = Math.floor(msOfDay / 60_000) % 60;
  const second = Math.floor(msOfDay / 1000) % 60;
  return `${dateText}T${pad(hour)}:${pad(minute)}:${pad(second)}${fractionText(msOfDay % 1000)}`;
}

/** A day fraction as an ODS `office:time-value` duration, hours unbounded:
 * 1.5034722… → "PT36H05M00S". */
export function serialToOdsTimeValue(serial: number): string {
  const totalMs = Math.round(Math.abs(serial) * MS_PER_DAY);
  const hours = Math.floor(totalMs / 3_600_000);
  const minutes = Math.floor(totalMs / 60_000) % 60;
  const seconds = Math.floor(totalMs / 1000) % 60;
  const sign = serial < 0 && totalMs > 0 ? "-" : "";
  return `${sign}PT${pad(hours)}H${pad(minutes)}M${pad(seconds)}${fractionText(totalMs % 1000)}S`;
}
