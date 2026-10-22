// SPEC-120 "The core functions" (dateTime.ts): DATE TIME YEAR MONTH DAY HOUR
// MINUTE SECOND TODAY NOW WEEKDAY EDATE EOMONTH DAYS in both date systems.
// In an xlsx workbook's 1900 system February 1900 has Excel's fictitious
// 29th in every function (=DATE(1900,2,29) and =EOMONTH(DATE(1900,2,1),0)
// are 60); in an .ods workbook (dateQuirks "proleptic") it has 28 days and
// no function produces serial 60; in the 1904 system serial 0 is 1904-01-01.
// DATE adds 1900 to years 0–1899, carries month and day overflow and is
// #NUM! for a year outside 0–9999 or a negative serial; TIME truncates and is
// the day fraction modulo 1 (#NUM! when negative); HOUR/MINUTE/SECOND round
// to the nearest second; WEEKDAY's return types 1, 2, 3 and 11–17; EDATE and
// EOMONTH clamp to the month's end; DAYS is INT(end) − INT(start), text
// coerced; TODAY and NOW read the injected ctx.now — never the wall clock
// (2026-10-08 09:30 here, a day before this spec was written).
//
// Sources ("Where expected values come from"): "docs" rows are Microsoft's
// function-reference examples (serials computed by hand from the dates they
// name: 2008-01-01 is 39448, 2011-01-01 is 40544); "pinned" rows follow from
// the rule SPEC-120 states.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../../address";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellError, CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

const err = (code: ErrorCode): CellError => ({ error: code });

/** 2026-10-08 09:30 local, a Thursday: serial 46303 + 9.5/24. */
const NOW = new Date(2026, 9, 8, 9, 30);

interface Where {
  format?: "xlsx" | "ods";
  dateSystem?: "1900" | "1904";
  now?: Date;
}

/** What `formula`, in Z100 of Sheet1 beside `cells`, stores — an xlsx
 * workbook in the 1900 system with the clock at NOW unless `where` says. */
function value(formula: string, cells: Cells = {}, where: Where = {}) {
  const host = "Z100";
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: { formula } } }, { format: where.format, dateSystem: where.dateSystem });
  const at = parseA1(host)!;
  const now = where.now ?? NOW;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, { ...storedValueContext(model), now: () => now });
}

/** Whole numbers and errors exactly; a fraction to `digits` decimal places
 * (12 for a time of day; 9 — still finer than a millisecond, 1.2e-8 — for a
 * serial near 46303, whose double carries only ~11 decimal digits). */
function expectValue(actual: unknown, expected: CellValue, digits = 12): void {
  if (typeof expected === "number" && !Number.isInteger(expected)) {
    expect(typeof actual).toBe("number");
    expect(actual as number).toBeCloseTo(expected, digits);
  } else {
    expect(actual).toEqual(expected);
  }
}

type Row = [formula: string, expected: CellValue];

describe("DATE", () => {
  it.each<Row>([
    ["DATE(2011,1,1)", 40544], // docs
    ["DATE(2008,1,2)", 39449], // docs
    ["DATE(108,1,2)", 39449], // conformance row: 1900 + 108
    ["DATE(0,1,1)", 1], // docs: years 0–1899 count from 1900
    ["DATE(2008,14,2)", 39846], // conformance row: 2009-02-02
    ["DATE(2008,-3,2)", 39327], // docs: 2007-09-02
    ["DATE(2008,1,35)", 39482], // docs: 2008-02-04
    ["DATE(2008,1,-15)", 39432], // docs: 2007-12-16
    ["DATE(1900,1,1)", 1],
    ["DATE(1900,2,28)", 59],
    ["DATE(1900,2,29)", 60], // conformance row: Excel's fictitious leap day
    ["DATE(1900,3,1)", 61],
    ["DATE(1900,1,0)", 0], // pinned: not negative, so not #NUM!
    ["DATE(2026,10,8)", 46303],
    ["DATE(9999,12,31)", 2958465],
    ['DATE("2008","1","2")', 39449], // numeric text coerced
    ["DATE(-1,1,1)", err("#NUM!")], // docs: year < 0
    ["DATE(10000,1,1)", err("#NUM!")], // docs: year ≥ 10000
    ["DATE(1900,1,-1)", err("#NUM!")], // pinned: a negative serial
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("TIME", () => {
  it.each<Row>([
    ["TIME(12,0,0)", 0.5], // docs
    ["TIME(16,48,10)", 60490 / 86400], // docs: 0.7001157
    ["TIME(27,0,0)", 0.125], // docs: = TIME(3,0,0)
    ["TIME(0,750,0)", 45000 / 86400], // docs: = TIME(12,30,0)
    ["TIME(0,0,2000)", 2000 / 86400], // docs: 0.023148
    ["TIME(24,0,0)", 0], // pinned: modulo 1
    ["TIME(1.9,30.9,0.9)", 5400 / 86400], // pinned: arguments truncated
    ["TIME(1,-30,0)", 1800 / 86400], // pinned: the total is not negative
    ["TIME(-1,0,0)", err("#NUM!")], // pinned
    ["TIME(0,0,-1)", err("#NUM!")], // pinned
  ])("=%s is %j", (formula, expected) => {
    expectValue(value(formula), expected);
  });
});

describe("YEAR, MONTH and DAY", () => {
  it.each<Row>([
    ["YEAR(DATE(2008,7,5))", 2008], // docs
    ["YEAR(DATE(2010,7,5))", 2010], // docs
    ['MONTH("15-Apr-2011")', 4], // docs, the date as text
    ['DAY("15-Apr-2011")', 15], // docs
    ['YEAR("2008-07-05")', 2008],
    ["YEAR(0)", 1900], // Excel's 1900-01-00
    ["MONTH(0)", 1],
    ["DAY(0)", 0],
    ["YEAR(60)", 1900], // Excel's 1900-02-29
    ["MONTH(60)", 2],
    ["DAY(60)", 29],
    ["YEAR(2958465)", 9999],
    ["MONTH(2958465)", 12],
    ["DAY(2958465)", 31],
    ["DAY(46303.75)", 8], // the time of day is ignored
    ["YEAR(-1)", err("#NUM!")],
    ["YEAR(2958466)", err("#NUM!")], // after 9999-12-31
    ['YEAR("x")', err("#VALUE!")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("HOUR, MINUTE and SECOND", () => {
  it.each<Row>([
    ["HOUR(0.75)", 18], // docs
    ["HOUR(DATE(2011,7,18)+TIME(7,45,0))", 7], // docs: 7/18/2011 7:45
    ["MINUTE(DATE(2011,7,18)+TIME(7,45,0))", 45],
    ["HOUR(DATE(2012,4,21))", 0], // docs: 4/21/2012
    ['MINUTE("12:45:00 PM")', 45], // docs
    ['SECOND("4:48:18 PM")', 18], // docs
    ['SECOND("4:48 PM")', 0], // docs
    ['HOUR("4:48:18 PM")', 16],
    ["HOUR(1.75)", 18], // the date is ignored
    ["SECOND(10.6/86400)", 11], // pinned: the nearest second
    ["SECOND(10.4/86400)", 10], // pinned
    ["MINUTE(59.6/86400)", 1], // pinned: 0:00:59.6 rounds to 0:01:00
    ["SECOND(59.6/86400)", 0], // pinned
    ["HOUR(-0.5)", err("#NUM!")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("TODAY and NOW", () => {
  it("read the injected clock, not the wall clock", () => {
    expect(value("TODAY()")).toBe(46303);
    expectValue(value("NOW()"), 46303 + 9.5 / 24, 9);
    expectValue(value("NOW()-TODAY()"), 9.5 / 24, 9);
    expect(value("YEAR(TODAY())*10000+MONTH(TODAY())*100+DAY(TODAY())")).toBe(20261008);
    expect(value("WEEKDAY(TODAY())")).toBe(5);
    expect(value("HOUR(NOW())*100+MINUTE(NOW())")).toBe(930);
  });

  it("NOW is local time to the millisecond", () => {
    const now = new Date(2026, 9, 8, 9, 30, 15, 250);
    expectValue(value("NOW()", {}, { now }), 46303 + (9 * 3600 + 30 * 60 + 15.25) / 86400, 9);
  });

  it("follow the workbook's date system", () => {
    expect(value("TODAY()", {}, { dateSystem: "1904" })).toBe(46303 - 1462);
    expect(value("TODAY()", {}, { format: "ods" })).toBe(46303);
  });
});

describe("WEEKDAY", () => {
  // docs: 2008-02-14, a Thursday, is 5, 4 and 3 for types 1 (default), 2, 3.
  it.each<[type: number, expected: number]>([
    [1, 5],
    [2, 4],
    [3, 3],
    [11, 4], // Monday is 1
    [12, 3], // Tuesday is 1
    [13, 2], // Wednesday is 1
    [14, 1], // Thursday is 1
    [15, 7], // Friday is 1
    [16, 6], // Saturday is 1
    [17, 5], // Sunday is 1
  ])("return type %d gives Thursday 2008-02-14 the number %d", (type, expected) => {
    expect(value(`WEEKDAY(DATE(2008,2,14),${type})`)).toBe(expected);
  });

  it.each([0, 4, 10, 18])("return type %d is #NUM!", (type) => {
    // docs: a return_type out of the listed range is #NUM!
    expect(value(`WEEKDAY(DATE(2008,2,14),${type})`)).toEqual(err("#NUM!"));
  });

  it.each<Row>([
    ["WEEKDAY(DATE(2008,2,14))", 5], // docs
    ["WEEKDAY(1)", 1], // conformance row (Excel): 1900-01-01 is "Sunday"
    ["WEEKDAY(46303)", 5], // conformance row (pinned)
    ["WEEKDAY(0)", 7], // pinned: (0 + 6) mod 7 + 1
    ['WEEKDAY("2026-10-08")', 5],
    ["WEEKDAY(-1)", err("#NUM!")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("EDATE and EOMONTH", () => {
  it.each<Row>([
    ["EDATE(DATE(2011,1,15),1)", 40589], // docs: 15-Feb-11
    ["EDATE(DATE(2011,1,15),-1)", 40527], // docs: 15-Dec-10
    ["EDATE(DATE(2011,1,15),2)", 40617], // docs: 15-Mar-11
    ["EDATE(DATE(2011,1,15),1.9)", 40589], // docs: months truncated
    ["EDATE(DATE(2011,1,31),1)", 40602], // clamped: 2011-02-28
    ["EDATE(DATE(2012,1,31),1)", 40968], // clamped: 2012-02-29
    ["EDATE(DATE(2011,3,31),-1)", 40602],
    ['EDATE("2011-01-15",1)', 40589],
    ['EDATE("x",1)', err("#VALUE!")], // docs: start_date not a date
    ["EOMONTH(DATE(2011,1,1),1)", 40602], // docs: 2/28/2011
    ["EOMONTH(DATE(2011,1,1),-3)", 40482], // docs: 10/31/2010
    ["EOMONTH(DATE(2011,1,15),1.9)", 40602], // docs: months truncated
    ["EOMONTH(DATE(1900,2,1),0)", 60], // acceptance criterion: the fictitious 29th
    ["EDATE(DATE(1900,1,31),1)", 60], // pinned: February 1900 has 29 days everywhere
    ["EOMONTH(-1,0)", err("#NUM!")], // docs: start_date not valid
    ["EOMONTH(DATE(9999,12,1),1)", err("#NUM!")], // docs: the result is no valid date
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("DAYS", () => {
  it.each<Row>([
    ['DAYS("15-MAR-2021","1-FEB-2021")', 42], // docs
    ["DAYS(DATE(2021,12,31),DATE(2021,1,1))", 364], // docs
    ["DAYS(DATE(2021,1,1),DATE(2021,12,31))", -364],
    ["DAYS(10.9,1.2)", 9], // pinned: INT(end) − INT(start)
    ['DAYS("2021-03-15 18:00","2021-03-15")', 0], // docs: text gives the integer date
    ['DAYS("x",1)', err("#VALUE!")], // docs: text that is no date
    ["DAYS(-1,1)", err("#NUM!")], // docs: a number outside the valid dates
    ["DAYS(2958466,1)", err("#NUM!")], // docs: after 9999-12-31
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("an .ods workbook: February 1900 has 28 days", () => {
  it.each<Row>([
    ["DATE(1900,2,28)", 59],
    ["DATE(1900,2,29)", 61], // = 1900-03-01
    ["DATE(1900,3,1)", 61],
    ["EOMONTH(DATE(1900,2,1),0)", 59],
    ["DAY(EOMONTH(DATE(1900,2,1),0))", 28],
    ["EDATE(DATE(1900,1,31),1)", 59],
    ["WEEKDAY(DATE(1900,1,1))", 2], // the real 1900-01-01, a Monday
    // Pinned: serial 0 is 1899-12-31 in the proleptic calendar, as SPEC-118
    // renders it (render.spec.ts: 0 with "yyyy-mm-dd" → "1899-12-31") — not
    // Excel's 1900-01-00, which only the "excel" quirks have.
    ["YEAR(0)", 1899],
    ["MONTH(0)", 12],
    ["DAY(0)", 31],
    ["WEEKDAY(0)", 1], // 1899-12-31 was a Sunday
    ["DATE(2026,10,8)", 46303], // from March 1900 the serials agree with Excel's
    ["WEEKDAY(46303)", 5],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula, {}, { format: "ods" })).toEqual(expected);
  });

  it("no date function produces serial 60 (the xlsx workbook's DATE does)", () => {
    for (let d = 1; d <= 70; d += 1) {
      expect(value(`DATE(1900,1,${d})`, {}, { format: "ods" })).not.toBe(60);
    }
    for (let d = 1; d <= 31; d += 1) {
      expect(value(`EDATE(DATE(1900,1,${d}),1)`, {}, { format: "ods" })).not.toBe(60);
      expect(value(`EOMONTH(DATE(1900,1,${d}),1)`, {}, { format: "ods" })).not.toBe(60);
    }
    expect(value("DATE(1900,1,60)")).toBe(60);
  });
});

describe("the 1904 date system", () => {
  it.each<Row>([
    ["DATE(1904,1,1)", 0],
    ["DATE(2008,1,2)", 39449 - 1462],
    ["DATE(1903,12,31)", err("#NUM!")], // before serial 0
    ["YEAR(0)", 1904],
    ["MONTH(0)", 1],
    ["DAY(0)", 1],
    ["WEEKDAY(0)", 6], // 1904-01-01 was a Friday: (0 + 5) mod 7 + 1
    ["EOMONTH(DATE(2011,1,1),1)", 40602 - 1462],
    ["DAYS(DATE(2021,12,31),DATE(2021,1,1))", 364],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula, {}, { dateSystem: "1904" })).toEqual(expected);
  });
});
