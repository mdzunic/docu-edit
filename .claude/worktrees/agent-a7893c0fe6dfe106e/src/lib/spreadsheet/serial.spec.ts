// SPEC-116 "Dates are serials": the conversions every later spec converts
// through — the 1900 system's boundary days (59 = 1900-02-28, 60 = Excel's
// fictitious 1900-02-29, 61 = 1900-03-01, and the backward count before
// 1900-01-01 that keeps 1899-12-30 = −1), the 1904 system (serial 0 =
// 1904-01-01), times as day fractions rounded to the nearest millisecond,
// ODS `office:date-value` (with a time, fractional seconds and a zone
// designator) and `office:time-value` durations (past 24 hours too), and a
// custom `table:null-date` honoured both ways.
import { describe, expect, it } from "vitest";
import {
  dateToSerial,
  odsDateValueToSerial,
  odsTimeValueToSerial,
  serialToOdsDateValue,
  serialToOdsTimeValue,
  serialToParts,
} from "./serial";

const MS_PER_DAY = 86_400_000;

interface Day {
  year: number;
  month: number;
  day: number;
}

/** serialToParts' shape for a date at midnight. */
function midnight(year: number, month: number, day: number) {
  return { year, month, day, hour: 0, minute: 0, second: 0, ms: 0 };
}

/** Proleptic-Gregorian month lengths, written out independently of
 * serial.ts's day arithmetic. */
function daysInMonth(year: number, month: number): number {
  if (month === 2) {
    return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 ? 29 : 28;
  }
  return [4, 6, 9, 11].includes(month) ? 30 : 31;
}

/** Every day from `from` to `to` inclusive. With `lotusLeapDay`, February
 * 1900 runs to the 29th — the day the 1900 system counts that never was. */
function calendarDays(from: Day, to: Day, lotusLeapDay: boolean): Day[] {
  const days: Day[] = [];
  let { year, month, day } = from;
  for (;;) {
    days.push({ year, month, day });
    if (year === to.year && month === to.month && day === to.day) {
      return days;
    }
    const last = lotusLeapDay && year === 1900 && month === 2 ? 29 : daysInMonth(year, month);
    if (day < last) {
      day += 1;
    } else if (month < 12) {
      month += 1;
      day = 1;
    } else {
      year += 1;
      month = 1;
      day = 1;
    }
  }
}

/** "1900-02-29" for labels. */
function label({ year, month, day }: Day): string {
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

describe("the 1900 system", () => {
  it("counts serial 1 = 1900-01-01 … 59 = 1900-02-28, 60 = the fictitious 1900-02-29, 61 = 1900-03-01", () => {
    expect(dateToSerial(1900, 1, 1, "1900")).toBe(1);
    expect(dateToSerial(1900, 2, 28, "1900")).toBe(59);
    expect(dateToSerial(1900, 2, 29, "1900")).toBe(60);
    expect(dateToSerial(1900, 3, 1, "1900")).toBe(61);
  });

  it("counts backwards before 1900-01-01 from 1899-12-31 = 0, so 1899-12-30 = −1", () => {
    expect(dateToSerial(1899, 12, 31, "1900")).toBe(0);
    expect(dateToSerial(1899, 12, 30, "1900")).toBe(-1);
    expect(dateToSerial(1899, 12, 1, "1900")).toBe(-30);
  });

  it("is days since 1899-12-30 from 1900-03-01 on (well-known anchors)", () => {
    expect(dateToSerial(1904, 1, 1, "1900")).toBe(1462);
    expect(dateToSerial(1970, 1, 1, "1900")).toBe(25_569);
    expect(dateToSerial(2000, 1, 1, "1900")).toBe(36_526);
    expect(dateToSerial(2024, 1, 1, "1900")).toBe(45_292);
    expect(dateToSerial(2026, 10, 8, "1900")).toBe(46_303);
  });

  it("splits serials 59, 60 and 61 into 1900-02-28, the fictitious 1900-02-29 and 1900-03-01, at midnight", () => {
    expect(serialToParts(59, "1900")).toEqual(midnight(1900, 2, 28));
    expect(serialToParts(60, "1900")).toEqual({ year: 1900, month: 2, day: 29, hour: 0, minute: 0, second: 0, ms: 0 });
    expect(serialToParts(61, "1900")).toEqual(midnight(1900, 3, 1));
  });

  it("splits the serials around 1900-01-01, the backward count included", () => {
    expect(serialToParts(1, "1900")).toEqual(midnight(1900, 1, 1));
    expect(serialToParts(0, "1900")).toEqual(midnight(1899, 12, 31));
    expect(serialToParts(-1, "1900")).toEqual(midnight(1899, 12, 30));
  });

  it("round-trips every day from 1899-12-01 to 1901-12-31 on consecutive serials, the fictitious 1900-02-29 included", () => {
    const days = calendarDays({ year: 1899, month: 12, day: 1 }, { year: 1901, month: 12, day: 31 }, true);
    expect(days).toHaveLength(31 + 366 + 365);
    let expected = -30;
    for (const day of days) {
      const serial = dateToSerial(day.year, day.month, day.day, "1900");
      expect(serial, label(day)).toBe(expected);
      expect(serialToParts(serial, "1900"), label(day)).toEqual(midnight(day.year, day.month, day.day));
      expected += 1;
    }
    expect(dateToSerial(1901, 12, 31, "1900")).toBe(731);
  });
});

describe("the 1904 system", () => {
  it("starts at serial 0 = 1904-01-01", () => {
    expect(dateToSerial(1904, 1, 1, "1904")).toBe(0);
    expect(serialToParts(0, "1904")).toEqual(midnight(1904, 1, 1));
  });

  it("is days since 1904-01-01 — 1,462 less than the 1900 system from 1900-03-01 on", () => {
    expect(dateToSerial(1904, 2, 29, "1904")).toBe(59);
    expect(dateToSerial(1904, 3, 1, "1904")).toBe(60);
    expect(dateToSerial(1903, 12, 31, "1904")).toBe(-1);
    expect(dateToSerial(2000, 1, 1, "1904")).toBe(35_064);
    for (const [y, m, d] of [
      [1900, 3, 1],
      [1950, 6, 15],
      [2000, 2, 29],
      [2026, 10, 8],
      [2099, 12, 31],
    ] as const) {
      expect(dateToSerial(y, m, d, "1904"), `${y}-${m}-${d}`).toBe(dateToSerial(y, m, d, "1900") - 1462);
    }
  });

  it("round-trips every day from 1903-12-01 to 1905-12-31 on consecutive serials", () => {
    const days = calendarDays({ year: 1903, month: 12, day: 1 }, { year: 1905, month: 12, day: 31 }, false);
    let expected = -31;
    for (const day of days) {
      const serial = dateToSerial(day.year, day.month, day.day, "1904");
      expect(serial, label(day)).toBe(expected);
      expect(serialToParts(serial, "1904"), label(day)).toEqual(midnight(day.year, day.month, day.day));
      expected += 1;
    }
    expect(expected).toBe(-31 + 31 + 366 + 365);
  });

  it("round-trips a sample of dates across centuries", () => {
    for (const day of [
      { year: 1900, month: 1, day: 1 },
      { year: 1999, month: 12, day: 31 },
      { year: 2000, month: 2, day: 29 },
      { year: 2026, month: 10, day: 8 },
      { year: 2100, month: 3, day: 1 },
      { year: 9999, month: 12, day: 31 },
    ]) {
      const serial = dateToSerial(day.year, day.month, day.day, "1904");
      expect(serialToParts(serial, "1904"), label(day)).toEqual(midnight(day.year, day.month, day.day));
    }
  });
});

describe("times are day fractions, rounded to the nearest millisecond", () => {
  it("reads a serial's fraction as the time of day", () => {
    expect(serialToParts(45_366.75, "1900")).toEqual({ year: 2024, month: 3, day: 15, hour: 18, minute: 0, second: 0, ms: 0 });
    expect(serialToParts(45_366 + 49_500_500 / MS_PER_DAY, "1900")).toEqual({
      year: 2024,
      month: 3,
      day: 15,
      hour: 13,
      minute: 45,
      second: 0,
      ms: 500,
    });
    expect(serialToParts(0.75, "1904")).toEqual({ year: 1904, month: 1, day: 1, hour: 18, minute: 0, second: 0, ms: 0 });
  });

  it("rounds to the nearest millisecond", () => {
    const noon = { year: 2024, month: 3, day: 15, hour: 12, minute: 0, second: 0 };
    expect(serialToParts(45_366 + (MS_PER_DAY / 2 + 0.4) / MS_PER_DAY, "1900")).toEqual({ ...noon, ms: 0 });
    expect(serialToParts(45_366 + (MS_PER_DAY / 2 + 0.6) / MS_PER_DAY, "1900")).toEqual({ ...noon, ms: 1 });
    expect(serialToParts(45_366 + (MS_PER_DAY / 2 - 0.4) / MS_PER_DAY, "1900")).toEqual({ ...noon, ms: 0 });
  });

  it("carries a rounding that reaches midnight into the next day", () => {
    expect(serialToParts(45_366 + (MS_PER_DAY - 0.4) / MS_PER_DAY, "1900")).toEqual(midnight(2024, 3, 16));
    expect(serialToParts(59 + (MS_PER_DAY - 0.4) / MS_PER_DAY, "1900")).toEqual(midnight(1900, 2, 29));
  });

  it("keeps the scale monotonic before 1900-01-01: −0.5 is noon on 1899-12-30", () => {
    expect(serialToParts(-0.5, "1900")).toEqual({ year: 1899, month: 12, day: 30, hour: 12, minute: 0, second: 0, ms: 0 });
    expect(serialToParts(-0.75, "1900")).toEqual({ year: 1899, month: 12, day: 30, hour: 6, minute: 0, second: 0, ms: 0 });
  });
});

describe("ODS office:date-value → serial", () => {
  it("converts a date to the system's serial", () => {
    expect(odsDateValueToSerial("2026-10-08", "1900")).toBe(dateToSerial(2026, 10, 8, "1900"));
    expect(odsDateValueToSerial("2026-10-08", "1900")).toBe(46_303);
    expect(odsDateValueToSerial("1900-02-28", "1900")).toBe(59);
    expect(odsDateValueToSerial("1900-03-01", "1900")).toBe(61);
    expect(odsDateValueToSerial("1899-12-30", "1900")).toBe(-1);
  });

  it("adds a time as an exact day fraction, fractional seconds included", () => {
    expect(odsDateValueToSerial("2026-10-08T13:45:00.5", "1900")).toBe(
      dateToSerial(2026, 10, 8, "1900") + (13 * 3600 + 45 * 60 + 0.5) / 86_400,
    );
    expect(odsDateValueToSerial("2026-10-08T18:00:00", "1900")).toBe(46_303.75);
    expect(odsDateValueToSerial("2026-10-08T00:00:00", "1900")).toBe(46_303);
    expect(odsDateValueToSerial("1899-12-30T12:00:00", "1900")).toBe(-0.5);
  });

  it("ignores a zone designator", () => {
    const local = odsDateValueToSerial("2026-10-08T13:45:00", "1900");
    expect(odsDateValueToSerial("2026-10-08T13:45:00Z", "1900")).toBe(local);
    expect(odsDateValueToSerial("2026-10-08T13:45:00+02:00", "1900")).toBe(local);
    expect(odsDateValueToSerial("2026-10-08T13:45:00-05:00", "1900")).toBe(local);
  });

  it("converts in the 1904 system", () => {
    expect(odsDateValueToSerial("1904-01-01", "1904")).toBe(0);
    expect(odsDateValueToSerial("2026-10-08", "1904")).toBe(dateToSerial(2026, 10, 8, "1904"));
    expect(odsDateValueToSerial("2026-10-08T18:00:00", "1904")).toBe(dateToSerial(2026, 10, 8, "1904") + 0.75);
  });

  it("counts days from a custom table:null-date", () => {
    expect(odsDateValueToSerial("1900-01-01", "1900", "1900-01-01")).toBe(0);
    expect(odsDateValueToSerial("1900-01-02", "1900", "1900-01-01")).toBe(1);
    // Real calendar days from the null date — no fictitious 1900-02-29.
    expect(odsDateValueToSerial("1900-03-01", "1900", "1900-01-01")).toBe(59);
    expect(odsDateValueToSerial("1899-12-31", "1900", "1900-01-01")).toBe(-1);
    expect(odsDateValueToSerial("1900-01-02T06:00:00", "1900", "1900-01-01")).toBe(1.25);
    expect(odsDateValueToSerial("1970-01-02", "1900", "1970-01-01")).toBe(1);
  });

  it.each([
    "",
    "2026",
    "2026-10",
    "08/10/2026",
    "2026/10/08",
    "2026-13-01",
    "2026-00-10",
    "2026-10-32",
    "2026-10-08T25:00:00",
    "2026-10-08T12:60:00",
    "2026-10-08 13:45",
    "not a date",
  ])("throws a RangeError for %j", (text) => {
    expect(() => odsDateValueToSerial(text, "1900")).toThrow(RangeError);
  });

  it("throws a RangeError for an invalid null date", () => {
    expect(() => odsDateValueToSerial("2026-10-08", "1900", "garbage")).toThrow(RangeError);
  });
});

describe("ODS office:time-value → serial", () => {
  it("converts ISO 8601 durations to day fractions, hours past 24 included", () => {
    expect(odsTimeValueToSerial("PT12H30M00S")).toBe(0.5208333333333334);
    expect(odsTimeValueToSerial("PT12H30M00S")).toBe(750 / 1440);
    expect(odsTimeValueToSerial("PT36H05M00S")).toBe(2165 / 1440);
    expect(odsTimeValueToSerial("PT36H")).toBe(1.5);
    expect(odsTimeValueToSerial("PT00H00M01.5S")).toBe(1.5 / 86_400);
    expect(odsTimeValueToSerial("-PT06H")).toBe(-0.25);
    expect(odsTimeValueToSerial("P1DT12H")).toBe(1.5);
    expect(odsTimeValueToSerial("PT0S")).toBe(0);
  });

  it("rounds to the nearest millisecond", () => {
    expect(odsTimeValueToSerial("PT00H00M00.0004S")).toBe(0);
    expect(odsTimeValueToSerial("PT00H00M00.0006S")).toBe(1 / MS_PER_DAY);
  });

  it.each(["", "P", "PT", "-P", "12:00", "P1Y", "P1M", "PTH", "T12H", "PT12H30M00"])("throws a RangeError for %j", (text) => {
    expect(() => odsTimeValueToSerial(text)).toThrow(RangeError);
  });
});

describe("serial → ODS office:date-value", () => {
  it("writes the date alone at midnight", () => {
    expect(serialToOdsDateValue(46_303, "1900")).toBe("2026-10-08");
    expect(serialToOdsDateValue(1, "1900")).toBe("1900-01-01");
    expect(serialToOdsDateValue(61, "1900")).toBe("1900-03-01");
    expect(serialToOdsDateValue(0, "1900")).toBe("1899-12-31");
    expect(serialToOdsDateValue(-1, "1900")).toBe("1899-12-30");
  });

  it("writes the fictitious 1900-02-29 (serial 60) as 1900-02-28, the day it repeats", () => {
    expect(serialToOdsDateValue(59, "1900")).toBe("1900-02-28");
    expect(serialToOdsDateValue(60, "1900")).toBe("1900-02-28");
  });

  it("adds the time, fractional seconds with trailing zeros trimmed", () => {
    expect(serialToOdsDateValue(46_303 + (13 * 3600 + 45 * 60 + 0.5) / 86_400, "1900")).toBe("2026-10-08T13:45:00.5");
    expect(serialToOdsDateValue(46_303.75, "1900")).toBe("2026-10-08T18:00:00");
    expect(serialToOdsDateValue(46_303 + 1250 / MS_PER_DAY, "1900")).toBe("2026-10-08T00:00:01.25");
    expect(serialToOdsDateValue(46_303 + 5 / MS_PER_DAY, "1900")).toBe("2026-10-08T00:00:00.005");
    expect(serialToOdsDateValue(-0.5, "1900")).toBe("1899-12-30T12:00:00");
  });

  it("inverts the 1904 system", () => {
    expect(serialToOdsDateValue(0, "1904")).toBe("1904-01-01");
    expect(serialToOdsDateValue(46_303 - 1462, "1904")).toBe("2026-10-08");
    expect(serialToOdsDateValue(-1, "1904")).toBe("1903-12-31");
  });

  it("inverts a custom null date", () => {
    expect(serialToOdsDateValue(0, "1900", "1900-01-01")).toBe("1900-01-01");
    expect(serialToOdsDateValue(1, "1900", "1900-01-01")).toBe("1900-01-02");
    expect(serialToOdsDateValue(59, "1900", "1900-01-01")).toBe("1900-03-01");
    expect(serialToOdsDateValue(1.25, "1900", "1900-01-01")).toBe("1900-01-02T06:00:00");
  });
});

describe("serial → ODS office:time-value", () => {
  it("writes a day fraction as a duration, hours unbounded", () => {
    expect(serialToOdsTimeValue(2165 / 1440)).toBe("PT36H05M00S");
    expect(serialToOdsTimeValue(750 / 1440)).toBe("PT12H30M00S");
    expect(serialToOdsTimeValue(1.5 / 86_400)).toBe("PT00H00M01.5S");
    expect(serialToOdsTimeValue(1.5)).toBe("PT36H00M00S");
    expect(serialToOdsTimeValue(0)).toBe("PT00H00M00S");
  });

  it("writes a negative duration with a leading minus", () => {
    expect(serialToOdsTimeValue(-0.25)).toBe("-PT06H00M00S");
    expect(serialToOdsTimeValue(-2165 / 1440)).toBe("-PT36H05M00S");
  });
});

describe("the ODS conversions are inverses", () => {
  it.each([
    { text: "2026-10-08", system: "1900" },
    { text: "2026-10-08T13:45:00.5", system: "1900" },
    { text: "2026-10-08T23:59:59.999", system: "1900" },
    { text: "1900-01-01", system: "1900" },
    { text: "1900-02-28", system: "1900" },
    { text: "1900-03-01T06:00:00", system: "1900" },
    { text: "1899-12-30", system: "1900" },
    { text: "1899-12-30T12:00:00", system: "1900" },
    { text: "1850-06-15T06:30:15.25", system: "1900" },
    { text: "9999-12-31T23:59:59.999", system: "1900" },
    { text: "1904-01-01", system: "1904" },
    { text: "1903-12-31T18:00:00", system: "1904" },
    { text: "2026-10-08T13:45:00.5", system: "1904" },
    { text: "1900-01-02", system: "1900", nullDate: "1900-01-01" },
    { text: "1899-06-01T01:02:03.004", system: "1900", nullDate: "1900-01-01" },
    { text: "2026-10-08T13:45:00.5", system: "1900", nullDate: "1970-01-01" },
  ] satisfies { text: string; system: "1900" | "1904"; nullDate?: string }[])(
    "office:date-value $text ($system, null date $nullDate) survives serial and back",
    ({ text, system, nullDate }) => {
      expect(serialToOdsDateValue(odsDateValueToSerial(text, system, nullDate), system, nullDate)).toBe(text);
    },
  );

  it("a millisecond-exact serial survives office:date-value and back", () => {
    for (const serial of [46_303, 46_303 + 49_500_500 / MS_PER_DAY, 61 + 21_600_000 / MS_PER_DAY, -1 + 43_200_000 / MS_PER_DAY]) {
      expect(odsDateValueToSerial(serialToOdsDateValue(serial, "1900"), "1900")).toBe(serial);
      expect(odsDateValueToSerial(serialToOdsDateValue(serial, "1904"), "1904")).toBe(serial);
      expect(odsDateValueToSerial(serialToOdsDateValue(serial, "1900", "1900-01-01"), "1900", "1900-01-01")).toBe(serial);
    }
  });

  it("office:time-value and the day fraction convert both ways", () => {
    for (const text of ["PT36H05M00S", "PT12H30M00S", "PT00H00M01.5S", "PT00H00M00.001S", "-PT06H00M00S", "PT1000H00M00S"]) {
      expect(serialToOdsTimeValue(odsTimeValueToSerial(text)), text).toBe(text);
    }
    for (const serial of [2165 / 1440, 750 / 1440, 1.5 / 86_400, -0.25, 1.5, 1 / MS_PER_DAY]) {
      expect(odsTimeValueToSerial(serialToOdsTimeValue(serial))).toBe(serial);
    }
  });
});
