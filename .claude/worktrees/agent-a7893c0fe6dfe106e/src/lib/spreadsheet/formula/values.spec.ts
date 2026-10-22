// SPEC-120 "Values", "Coercion", "Text to number" and "Comparison": the
// bounds and the EMPTY singleton; toNumber / toText / toLogical cell by cell
// against the "Coercion" table; valueFromText in all six SPEC-118 locales —
// sign and enclosing parentheses, group separators only between three-digit
// groups (a space-like group accepting U+0020, U+00A0 and U+202F), the
// decimal fraction, the exponent, a trailing %, the locale's currency at its
// position, dates in the locale's order with / - . and month names (English
// and the locale's own), a missing year read from core now(), two-digit
// years, ISO yyyy-mm-dd everywhere, impossible dates, the fictitious
// 1900-02-29, times with AM/PM and past 24 hours, a date and a time added —
// plus the forms SPEC-122's parseInput takes its values from; numberToText's
// notation boundaries; roundTo15; compareScalars (type order, EMPTY, errors
// left first, 15-digit numbers, the locale's accent-sensitive collator and
// caseSensitive); toCellValue; and power's error cases.
//
// Expected serials are computed by hand, never by running the engine: a
// 1900-system serial counts days from 1899-12-30 for dates from 1900-03-01
// (2026-10-08 is 46303, the conformance table's pin; 9999-12-31 is Excel's
// last date, 2958465), 60 is the fictitious 1900-02-29, and a 1904-system
// serial is 1462 less (1904-01-01 is 0).
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { setClock } from "../../core/clock";
import type { CellError, CellValue } from "../model";
import { SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocale, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import {
  compareScalars,
  EMPTY,
  MAX_ARRAY_CELLS,
  MAX_TEXT_LENGTH,
  numberToText,
  power,
  roundTo15,
  toCellValue,
  toLogical,
  toNumber,
  toText,
  valueFromText,
  type ScalarValue,
} from "./values";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const DE_DE = SPREADSHEET_LOCALES["de-DE"];

/** The settings toNumber reads. */
const numberSettings = (locale: SpreadsheetLocale, dateSystem: "1900" | "1904" = "1900"): { locale: SpreadsheetLocale; dateSystem: "1900" | "1904" } => ({
  locale,
  dateSystem,
});

/** The settings compareScalars reads. */
const compareSettings = (locale: SpreadsheetLocale, caseSensitive = false): { locale: SpreadsheetLocale; caseSensitive: boolean } => ({
  locale,
  caseSensitive,
});

const err = (code: CellError["error"]): CellError => ({ error: code });

/** The day fraction of a time of day. */
const time = (h: number, m: number, s = 0): number => (h * 3600 + m * 60 + s) / 86_400;

// Serials by hand (see the header).
const OCT_8_2026 = 46303;
const OCT_1_2026 = 46296;
const JAN_5_2026 = 46027;
const JAN_13_2026 = 46035;
const JAN_15_2026 = 46037;
const MAR_14_2026 = 46095;
const MAY_1_2026 = 46143;
const DEC_8_2026 = 46364;
const DEC_24_2026 = 46380;
const OCT_8_2027 = 46668;
const JAN_1_2000 = 36526;
const DEC_31_1999 = 36525;
const JAN_1_2029 = 47119;
const JAN_1_1930 = 10959;
const FEB_29_2024 = 45351;
const DEC_31_9999 = 2958465;
/** 46303 − 1462. */
const OCT_8_2026_1904 = 44841;

/** `actual` is `expected`: exactly for null and integers, to a relative 1e-13
 * otherwise — the spec pins the value, not the floating-point steps that
 * build a day fraction. */
function expectNumber(actual: number | null, expected: number | null): void {
  if (expected === null || Number.isInteger(expected)) {
    expect(actual).toBe(expected);
    return;
  }
  const close = actual !== null && Math.abs(actual - expected) <= Math.abs(expected) * 1e-13;
  expect(close ? expected : actual).toBe(expected);
}

const parse = (tag: SpreadsheetLocaleTag, text: string, dateSystem: "1900" | "1904" = "1900"): number | null =>
  valueFromText(text, SPREADSHEET_LOCALES[tag], dateSystem);

// A missing year is the current year, read from core now(): Thursday
// 2026-10-08 09:30, local time.
beforeEach(() => {
  setClock(new Date(2026, 9, 8, 9, 30));
});

afterEach(() => {
  setClock(null);
});

describe("the value model's constants", () => {
  it("bounds arrays at 4,194,304 elements and text at 32,767 characters", () => {
    expect(MAX_ARRAY_CELLS).toBe(4_194_304);
    expect(MAX_TEXT_LENGTH).toBe(32_767);
  });

  it("has one EMPTY value of kind \"empty\"", () => {
    expect(EMPTY).toEqual({ kind: "empty" });
    expect(EMPTY.kind).toBe("empty");
  });
});

describe("toNumber (the Coercion table's number column)", () => {
  const US = numberSettings(EN_US);

  it.each<[string, ScalarValue, number | CellError]>([
    ["a number is itself", 42, 42],
    ["a negative number is itself", -1.5, -1.5],
    ["zero is itself", 0, 0],
    ["text through valueFromText", "1", 1],
    ["currency text through valueFromText", "$4.00", 4],
    ["date text through valueFromText", "6/1/2001", 37043],
    ["time text through valueFromText", "12:00", 0.5],
    ["percent text through valueFromText", "10%", 0.1],
    ["other text is #VALUE!", "abc", err("#VALUE!")],
    ["\"\" is #VALUE!", "", err("#VALUE!")],
    ["\"TRUE\" is #VALUE! — booleans are never parsed from text", "TRUE", err("#VALUE!")],
    ["\"FALSE\" is #VALUE!", "FALSE", err("#VALUE!")],
    ["\"1,5\" is #VALUE! in en-US", "1,5", err("#VALUE!")],
    ["TRUE is 1", true, 1],
    ["FALSE is 0", false, 0],
    ["EMPTY is 0", EMPTY, 0],
    ["an error is that error", err("#DIV/0!"), err("#DIV/0!")],
    ["#N/A is #N/A", err("#N/A"), err("#N/A")],
  ])("%s", (_label, value, expected) => {
    expect(toNumber(value, US)).toEqual(expected);
  });

  it("reads text in the given locale and date system", () => {
    expect(toNumber("1,5", numberSettings(DE_DE))).toBe(1.5);
    expect(toNumber("1.000", numberSettings(DE_DE))).toBe(1000);
    expect(toNumber("2026-10-08", numberSettings(EN_US, "1904"))).toBe(OCT_8_2026_1904);
  });
});

describe("toText (the Coercion table's text column)", () => {
  it.each<[string, ScalarValue, string | CellError]>([
    ["a number through numberToText", 1 / 3, "0.333333333333333"],
    ["a decimal number", 1.5, "1.5"],
    ["a large number in scientific notation", 1e15, "1E+15"],
    ["a negative number", -2, "-2"],
    ["zero", 0, "0"],
    ["text is itself", "abc", "abc"],
    ["\"\" is itself", "", ""],
    ["text that looks boolean is itself", "true", "true"],
    ["TRUE is \"TRUE\"", true, "TRUE"],
    ["FALSE is \"FALSE\"", false, "FALSE"],
    ["EMPTY is \"\"", EMPTY, ""],
    ["an error is that error", err("#VALUE!"), err("#VALUE!")],
  ])("%s", (_label, value, expected) => {
    expect(toText(value)).toEqual(expected);
  });
});

describe("toLogical (the Coercion table's logical column)", () => {
  it.each<[string, ScalarValue, boolean | CellError]>([
    ["0 is FALSE", 0, false],
    ["1 is TRUE", 1, true],
    ["a negative number is TRUE", -0.5, true],
    ["a tiny number is TRUE", 1e-300, true],
    ["2 is TRUE", 2, true],
    ["\"TRUE\" is TRUE", "TRUE", true],
    ["\"true\" is TRUE (any case)", "true", true],
    ["\"tRuE\" is TRUE (any case)", "tRuE", true],
    ["\"FALSE\" is FALSE", "FALSE", false],
    ["\"False\" is FALSE (any case)", "False", false],
    ["other text is #VALUE!", "yes", err("#VALUE!")],
    ["numeric text is #VALUE!", "1", err("#VALUE!")],
    ["\"0\" is #VALUE!", "0", err("#VALUE!")],
    ["\"\" is #VALUE!", "", err("#VALUE!")],
    ["TRUE is itself", true, true],
    ["FALSE is itself", false, false],
    ["EMPTY is FALSE", EMPTY, false],
    ["an error is that error", err("#N/A"), err("#N/A")],
  ])("%s", (_label, value, expected) => {
    expect(toLogical(value)).toEqual(expected);
  });
});

describe("valueFromText — the conformance table's pinned rows", () => {
  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["en-US", "1,5", null],
    ["de-DE", "1,5", 1.5],
    ["en-US", "1,000", 1000],
    ["en-US", "2026-10-08", OCT_8_2026],
    ["en-US", "12:00", 0.5],
    ["en-US", "(5)", -5],
    ["en-US", "10%", 0.1],
    // ="1"+"$4.00" is 5 (Excel).
    ["en-US", "$4.00", 4],
    // =SQRT("8+1"), ="TRUE"+0, =1+"" and ="a"*1 are #VALUE! (Excel).
    ["en-US", "8+1", null],
    ["en-US", "TRUE", null],
    ["en-US", "", null],
    ["en-US", "a", null],
  ])("%s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it("makes \"6/1/2001\" − \"5/1/2001\" 31 days in en-US (Excel)", () => {
    expect(parse("en-US", "6/1/2001")! - parse("en-US", "5/1/2001")!).toBe(31);
  });
});

describe("valueFromText — numbers", () => {
  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    // Sign.
    ["en-US", "42", 42],
    ["en-US", "-1.5", -1.5],
    ["en-US", "+7", 7],
    ["en-US", "007", 7],
    ["de-DE", "-1,5", -1.5],
    // Enclosing parentheses: negative.
    ["en-US", "(12)", -12],
    ["en-US", "(1,234.5)", -1234.5],
    ["de-DE", "(1.234,5)", -1234.5],
    ["en-US", "(10%)", -0.1],
    ["en-US", "($1,200.50)", -1200.5],
    ["de-DE", "(4,00 €)", -4],
    // Only U+0020 is trimmed.
    ["en-US", "  5  ", 5],
    ["en-US", " 5", null],
    // A decimal fraction.
    ["en-US", "0.25", 0.25],
    ["en-US", ".5", 0.5],
    ["de-DE", ",5", 0.5],
    ["fr-FR", "3,75", 3.75],
    // An exponent.
    ["en-US", "1.5E3", 1500],
    ["en-US", "1e-4", 0.0001],
    ["en-US", "2E+2", 200],
    ["en-US", "-1.5e2", -150],
    ["de-DE", "1,5E3", 1500],
    // A trailing %.
    ["en-US", "12.5%", 0.125],
    ["en-US", "-10%", -0.1],
    ["en-US", "1,000%", 10],
    ["de-DE", "12,5%", 0.125],
  ])("%s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["en-US", "1,234,567", 1234567],
    ["en-US", "1,234,567.5", 1234567.5],
    ["en-US", "1,234.5", 1234.5],
    ["en-GB", "1,000", 1000],
    ["de-DE", "1.000", 1000],
    ["de-DE", "1.234.567,5", 1234567.5],
    ["es-ES", "1.000", 1000],
    ["es-ES", "1.234,5", 1234.5],
    // Not between groups of exactly three digits: no number.
    ["en-US", "1,0000", null],
    ["en-US", "12,34", null],
    ["en-US", "1,234,56", null],
    ["en-US", ",000", null],
    ["en-US", "1,000,", null],
    ["en-US", "1,,000", null],
    ["en-US", "1.234,5", null],
    ["en-GB", "1,5", null],
    // A space groups digits only where the locale's group is space-like.
    ["en-US", "1 000", null],
    ["de-DE", "1 000", null],
  ])("groups digits by the locale's separator in threes: %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["fr-FR", "1 000", 1000],
    ["fr-FR", "1 000", 1000],
    ["fr-FR", "1 000", 1000],
    ["fr-FR", "1 234 567,5", 1234567.5],
    ["fr-FR", "1 234,5", 1234.5],
    ["fr-FR", "1 234,50 €", 1234.5],
    ["cs-CZ", "1 000", 1000],
    ["cs-CZ", "1 000", 1000],
    ["cs-CZ", "1 000", 1000],
    ["cs-CZ", "1 234,5", 1234.5],
    ["fr-FR", "1 0000", null],
    ["cs-CZ", "12 34", null],
  ])("accepts U+0020, U+00A0 and U+202F for a space-like group: %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["en-US", "$1,200", 1200],
    ["en-US", "$1,200.50", 1200.5],
    ["en-US", "-$5", -5],
    ["en-US", "($5)", -5],
    ["en-GB", "£4.00", 4],
    ["en-GB", "£1,200.50", 1200.5],
    ["de-DE", "4,00 €", 4],
    ["de-DE", "4,00€", 4],
    ["de-DE", "1.200 €", 1200],
    ["de-DE", "-4,00 €", -4],
    ["fr-FR", "4,00 €", 4],
    ["fr-FR", "1 200 €", 1200],
    ["es-ES", "4,00 €", 4],
    ["es-ES", "1.200,50 €", 1200.5],
    ["cs-CZ", "4,00 Kč", 4],
    ["cs-CZ", "1 200 Kč", 1200],
    // Only the locale's own symbol.
    ["en-US", "€4.00", null],
    ["en-GB", "$4.00", null],
    ["de-DE", "$4,00", null],
  ])("reads the locale's currency symbol at its position: %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });
});

describe("valueFromText — dates", () => {
  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["en-US", "10/8/2026", OCT_8_2026],
    ["en-US", "10-8-2026", OCT_8_2026],
    ["en-US", "10.8.2026", OCT_8_2026],
    ["en-US", "1/5/2026", JAN_5_2026],
    ["en-GB", "8/10/2026", OCT_8_2026],
    ["en-GB", "8-10-2026", OCT_8_2026],
    ["en-GB", "8.10.2026", OCT_8_2026],
    ["en-GB", "1/5/2026", MAY_1_2026],
    ["en-GB", "13/1/2026", JAN_13_2026],
    ["de-DE", "8.10.2026", OCT_8_2026],
    ["de-DE", "8/10/2026", OCT_8_2026],
    ["de-DE", "8-10-2026", OCT_8_2026],
    ["fr-FR", "8/10/2026", OCT_8_2026],
    ["fr-FR", "8-10-2026", OCT_8_2026],
    ["es-ES", "8/10/2026", OCT_8_2026],
    ["es-ES", "8.10.2026", OCT_8_2026],
    ["cs-CZ", "8.10.2026", OCT_8_2026],
    ["cs-CZ", "8/10/2026", OCT_8_2026],
  ])("reads day, month and year in the locale's order with / - or .: %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("reads ISO yyyy-mm-dd in %s, in both date systems", (tag) => {
    expect(parse(tag, "2026-10-08")).toBe(OCT_8_2026);
    expect(parse(tag, "2026-10-08", "1904")).toBe(OCT_8_2026_1904);
  });

  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["en-US", "10/8", OCT_8_2026],
    ["en-US", "1/5", JAN_5_2026],
    ["en-US", "Oct 8", OCT_8_2026],
    ["en-US", "8-Oct", OCT_8_2026],
    ["en-GB", "8/10", OCT_8_2026],
    ["en-GB", "1/5", MAY_1_2026],
    ["de-DE", "8.10", OCT_8_2026],
    // "." groups thousands in de-DE, so "1.5" is no number — it is 1 May.
    ["de-DE", "1.5", MAY_1_2026],
    ["cs-CZ", "8.10", OCT_8_2026],
    ["fr-FR", "8/10", OCT_8_2026],
    ["es-ES", "8/10", OCT_8_2026],
  ])("takes a missing year from core now(): %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it("reads the year from the clock, not a constant", () => {
    setClock(new Date(2027, 0, 15, 12, 0));
    expect(parse("en-US", "10/8")).toBe(OCT_8_2027);
    expect(parse("en-US", "8-Oct")).toBe(OCT_8_2027);
    expect(parse("en-US", "10/8/2026")).toBe(OCT_8_2026);
  });

  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    ["en-US", "10/8/26", OCT_8_2026],
    ["en-US", "10-8-26", OCT_8_2026],
    ["en-US", "1/1/00", JAN_1_2000],
    ["en-US", "1/1/29", JAN_1_2029],
    ["en-US", "1/1/30", JAN_1_1930],
    ["en-US", "12/31/99", DEC_31_1999],
    ["en-US", "8-Oct-26", OCT_8_2026],
    ["en-GB", "8/10/26", OCT_8_2026],
    ["de-DE", "08.10.26", OCT_8_2026],
    ["de-DE", "1.1.29", JAN_1_2029],
    ["de-DE", "1.1.30", JAN_1_1930],
  ])("reads two-digit years 00–29 as 20xx and 30–99 as 19xx: %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });

  it.each<[SpreadsheetLocaleTag, string]>([
    ["en-US", "2/30/2026"],
    ["en-US", "2/29/2026"],
    ["en-US", "4/31/2026"],
    ["en-US", "13/1/2026"],
    ["en-US", "0/1/2026"],
    ["en-US", "1/0/2026"],
    ["en-US", "1/32/2026"],
    ["en-US", "2026-02-30"],
    ["en-US", "2026-13-01"],
    ["en-GB", "30/2/2026"],
    ["en-GB", "31/4/2026"],
    ["de-DE", "30.2.2026"],
    ["de-DE", "2026-02-29"],
  ])("finds no match in an impossible date: %s %j → null", (tag, text) => {
    expect(parse(tag, text)).toBeNull();
  });

  it("accepts 29 February in a leap year", () => {
    expect(parse("en-US", "2/29/2024")).toBe(FEB_29_2024);
    expect(parse("de-DE", "29.2.2024")).toBe(FEB_29_2024);
  });

  it.each<[SpreadsheetLocaleTag, string, number]>([
    ["en-US", "1/1/1900", 1],
    ["en-US", "2/28/1900", 59],
    ["en-US", "2/29/1900", 60],
    ["en-US", "3/1/1900", 61],
    ["en-US", "1900-02-29", 60],
    ["de-DE", "29.2.1900", 60],
    ["en-US", "12/31/9999", DEC_31_9999],
  ])("counts the fictitious 1900-02-29 as serial 60 in the 1900 system: %s %j → %s", (tag, text, expected) => {
    expect(parse(tag, text)).toBe(expected);
  });

  it.each<[SpreadsheetLocaleTag, string, number]>([
    ["en-US", "1/1/1904", 0],
    ["en-US", "1/2/1904", 1],
    ["en-US", "1904-01-01", 0],
    ["en-US", "10/8/2026", OCT_8_2026_1904],
    ["de-DE", "8.10.2026", OCT_8_2026_1904],
    ["en-US", "8-Oct-2026", OCT_8_2026_1904],
  ])("returns 1904-system serials in a 1904 workbook: %s %j → %s", (tag, text, expected) => {
    expect(parse(tag, text, "1904")).toBe(expected);
  });

  it.each<[SpreadsheetLocaleTag, string, number | null]>([
    // English names and three-letter abbreviations (SPEC-122's en-US rows).
    ["en-US", "8-Oct-2026", OCT_8_2026],
    ["en-US", "Oct 8, 2026", OCT_8_2026],
    ["en-US", "8 October 2026", OCT_8_2026],
    ["en-US", "October 8, 2026", OCT_8_2026],
    ["en-US", "24-Dec-2026", DEC_24_2026],
    ["en-US", "Oct-2026", OCT_1_2026],
    ["en-US", "Oct 2026", OCT_1_2026],
    ["en-GB", "8 October 2026", OCT_8_2026],
    ["en-GB", "8-Oct-2026", OCT_8_2026],
    // English in every locale.
    ["de-DE", "8 October 2026", OCT_8_2026],
    ["fr-FR", "8-Oct-2026", OCT_8_2026],
    ["cs-CZ", "8 October 2026", OCT_8_2026],
    // The locale's own names and three-letter abbreviations.
    ["de-DE", "8. Oktober 2026", OCT_8_2026],
    ["de-DE", "8-Okt-2026", OCT_8_2026],
    ["de-DE", "14. März 2026", MAR_14_2026],
    ["de-DE", "24. Dezember 2026", DEC_24_2026],
    ["fr-FR", "8 octobre 2026", OCT_8_2026],
    ["fr-FR", "8-déc-2026", DEC_8_2026],
    ["es-ES", "8 octubre 2026", OCT_8_2026],
    ["es-ES", "15-ene-2026", JAN_15_2026],
    ["es-ES", "24-dic-2026", DEC_24_2026],
    ["cs-CZ", "8. října 2026", OCT_8_2026],
    ["cs-CZ", "8. říjen 2026", OCT_8_2026],
    ["cs-CZ", "8-říj-2026", OCT_8_2026],
    // Another locale's names are no month.
    ["en-US", "8 Oktober 2026", null],
  ])("reads a month name in place of the month number: %s %j → %s", (tag, text, expected) => {
    expectNumber(parse(tag, text), expected);
  });
});

describe("valueFromText — times", () => {
  it.each<[string, number]>([
    ["12:00", 0.5],
    ["9:30", time(9, 30)],
    ["13:45", time(13, 45)],
    ["0:00", 0],
    ["13:45:30", time(13, 45, 30)],
    ["13:45:30.5", time(13, 45, 30.5)],
    ["13:45:30.250", time(13, 45, 30.25)],
    // AM/PM in any case.
    ["1:45 PM", time(13, 45)],
    ["1:45:30 pm", time(13, 45, 30)],
    ["1 PM", time(13, 0)],
    ["9:30 Am", time(9, 30)],
    ["11:59 pM", time(23, 59)],
    ["12:00 AM", 0],
    ["12:00 PM", 0.5],
    ["12 am", 0],
    ["12:30:15 AM", time(0, 30, 15)],
    // Hours past 24 without AM/PM.
    ["24:00", 1],
    ["36:00", 1.5],
    ["25:30:00", 25.5 / 24],
    ["100:00", 100 / 24],
  ])("en-US %j → %s", (text, expected) => {
    expectNumber(parse("en-US", text), expected);
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("reads h:mm, AM/PM and hours past 24 the same in %s", (tag) => {
    expectNumber(parse(tag, "12:00"), 0.5);
    expectNumber(parse(tag, "13:45:30"), time(13, 45, 30));
    expectNumber(parse(tag, "1:45 PM"), time(13, 45));
    expectNumber(parse(tag, "36:00"), 1.5);
  });

  it.each<[SpreadsheetLocaleTag, string, number, "1900" | "1904"]>([
    ["en-US", "10/8/2026 13:45", OCT_8_2026 + time(13, 45), "1900"],
    ["en-US", "10/8/2026 1:45 PM", OCT_8_2026 + time(13, 45), "1900"],
    ["en-US", "10/8/2026   12:00", OCT_8_2026 + 0.5, "1900"],
    ["en-US", "Oct 8, 2026 1:45 PM", OCT_8_2026 + time(13, 45), "1900"],
    ["en-US", "8-Oct-2026 12:00", OCT_8_2026 + 0.5, "1900"],
    ["en-US", "2026-10-08 12:00", OCT_8_2026 + 0.5, "1900"],
    // SPEC-122's ISO date-time row.
    ["en-US", "2026-10-08T13:45", OCT_8_2026 + time(13, 45), "1900"],
    ["en-GB", "8/10/2026 13:45", OCT_8_2026 + time(13, 45), "1900"],
    ["de-DE", "8.10.2026 13:45", OCT_8_2026 + time(13, 45), "1900"],
    ["cs-CZ", "8.10.2026 13:45:30", OCT_8_2026 + time(13, 45, 30), "1900"],
    ["en-US", "2026-10-08 12:00", OCT_8_2026_1904 + 0.5, "1904"],
  ])("adds a date and a time separated by spaces: %s %j → %s (%s)", (tag, text, expected, dateSystem) => {
    expectNumber(parse(tag, text, dateSystem), expected);
  });
});

describe("valueFromText — no match", () => {
  it.each(SPREADSHEET_LOCALE_TAGS)("returns null in %s for \"\", blanks, booleans, error literals and other text", (tag) => {
    for (const text of ["", "   ", "TRUE", "FALSE", "true", "false", "False", "abc", "#N/A", "#div/0!", "8+1"]) {
      expect(parse(tag, text), JSON.stringify(text)).toBeNull();
    }
  });
});

describe("numberToText", () => {
  it.each<[string, number, string]>([
    ["0", 0, "0"],
    ["-0", -0, "0"],
    ["42", 42, "42"],
    ["100", 100, "100"],
    ["-1.5", -1.5, "-1.5"],
    ["1/3", 1 / 3, "0.333333333333333"],
    ["-1/3", -1 / 3, "-0.333333333333333"],
    ["2/3", 2 / 3, "0.666666666666667"],
    ["0.1+0.2", 0.1 + 0.2, "0.3"],
    ["1234567.89", 1234567.89, "1234567.89"],
    ["999999999999999", 999999999999999, "999999999999999"],
    ["1e14", 1e14, "100000000000000"],
    ["1e15", 1e15, "1E+15"],
    ["-1e15", -1e15, "-1E+15"],
    ["2^53", 2 ** 53, "9.00719925474099E+15"],
    ["123456789012345678", Number("123456789012345678"), "1.23456789012346E+17"],
    ["1.5e20", 1.5e20, "1.5E+20"],
    ["1e100", 1e100, "1E+100"],
    ["0.0001", 0.0001, "0.0001"],
    ["1e-9", 1e-9, "0.000000001"],
    ["1.5e-9", 1.5e-9, "0.0000000015"],
    ["-1e-9", -1e-9, "-0.000000001"],
    ["1e-10", 1e-10, "1E-10"],
    ["-1e-10", -1e-10, "-1E-10"],
    ["1.25e-12", 1.25e-12, "1.25E-12"],
    ["1e-300", 1e-300, "1E-300"],
  ])("%s → %j", (_label, n, expected) => {
    expect(numberToText(n)).toBe(expected);
  });
});

describe("roundTo15", () => {
  it.each<[string, number, number]>([
    ["0.1+0.2", 0.1 + 0.2, 0.3],
    ["1/3", 1 / 3, 0.333333333333333],
    ["-1/3", -1 / 3, -0.333333333333333],
    ["2/3", 2 / 3, 0.666666666666667],
    ["123456789.123456789", Number("123456789.123456789"), 123456789.123457],
    ["1e20 + 12345", 1e20 + 12345, 1e20],
    ["1 + 2^-52", 1 + 2 ** -52, 1],
    ["42", 42, 42],
  ])("%s → %s", (_label, n, expected) => {
    expect(roundTo15(n)).toBe(expected);
  });

  it("keeps 0 as 0 and makes −0 0", () => {
    expect(Object.is(roundTo15(0), 0)).toBe(true);
    expect(Object.is(roundTo15(-0), 0)).toBe(true);
  });

  it("is Number(n.toPrecision(15))", () => {
    for (const n of [Math.PI, -Math.E, 1 / 7, 1e-20 / 3, 2 ** 53 + 2, 0.1 * 3, 1e300 / 7, -123.456e-7]) {
      expect(roundTo15(n), String(n)).toBe(Number(n.toPrecision(15)));
    }
  });
});

describe("compareScalars", () => {
  const US = compareSettings(EN_US);

  it.each<[ScalarValue, ScalarValue, -1 | 0 | 1]>([
    // number < text < logical, whatever the values (=1<"1", ="z"<TRUE: Excel).
    [1, "1", -1],
    ["1", 1, 1],
    [1e300, "", -1],
    ["z", true, -1],
    ["", false, -1],
    [true, "a", 1],
    [false, 0, 1],
    [0, true, -1],
    // FALSE < TRUE (=FALSE<TRUE: Excel).
    [false, true, -1],
    [true, false, 1],
    [true, true, 0],
    [false, false, 0],
  ])("orders types number < text < logical: %j vs %j → %s", (a, b, expected) => {
    expect(compareScalars(a, b, US)).toBe(expected);
  });

  it.each<[ScalarValue, ScalarValue, -1 | 0 | 1]>([
    // =A1=0, =A1="", =A1=FALSE with A1 empty are all TRUE (Excel).
    [EMPTY, 0, 0],
    [0, EMPTY, 0],
    [EMPTY, "", 0],
    ["", EMPTY, 0],
    [EMPTY, false, 0],
    [false, EMPTY, 0],
    [EMPTY, EMPTY, 0],
    [EMPTY, 1, -1],
    [EMPTY, -1, 1],
    [1, EMPTY, 1],
    [EMPTY, "a", -1],
    ["a", EMPTY, 1],
    [EMPTY, true, -1],
    [true, EMPTY, 1],
  ])("gives an EMPTY operand the other side's type: %j vs %j → %s", (a, b, expected) => {
    expect(compareScalars(a, b, US)).toBe(expected);
  });

  it.each<[ScalarValue, ScalarValue, CellError]>([
    [err("#N/A"), err("#DIV/0!"), err("#N/A")],
    [err("#DIV/0!"), err("#N/A"), err("#DIV/0!")],
    [1, err("#REF!"), err("#REF!")],
    [err("#VALUE!"), 1, err("#VALUE!")],
    [EMPTY, err("#NUM!"), err("#NUM!")],
    [err("#NULL!"), EMPTY, err("#NULL!")],
    ["a", err("#NAME?"), err("#NAME?")],
  ])("returns an error operand, the left one first: %j vs %j → %j", (a, b, expected) => {
    expect(compareScalars(a, b, US)).toEqual(expected);
  });

  it.each<[string, number, number, -1 | 0 | 1]>([
    // =0.1+0.2=0.3 is TRUE (Excel).
    ["0.1+0.2 vs 0.3", 0.1 + 0.2, 0.3, 0],
    ["0.3 vs 0.1+0.2", 0.3, 0.1 + 0.2, 0],
    ["1/3 vs 0.333333333333333", 1 / 3, 0.333333333333333, 0],
    ["1 vs 2", 1, 2, -1],
    ["2 vs 1", 2, 1, 1],
    ["-1 vs 1", -1, 1, -1],
    ["-0 vs 0", -0, 0, 0],
    // Significant digits, not decimal places.
    ["1e-20 vs 0", 1e-20, 0, 1],
  ])("compares numbers at 15 significant digits: %s → %s", (_label, a, b, expected) => {
    expect(compareScalars(a, b, US)).toBe(expected);
  });

  it.each<[string, string, -1 | 0 | 1]>([
    // ="A"="a" and ="a"<"B" are TRUE (Excel).
    ["a", "A", 0],
    ["A", "a", 0],
    ["abc", "ABC", 0],
    ["a", "B", -1],
    ["B", "a", 1],
    ["apple", "banana", -1],
    ["same", "same", 0],
  ])("compares text case-insensitively through the collator: %j vs %j → %s", (a, b, expected) => {
    expect(compareScalars(a, b, US)).toBe(expected);
  });

  it("is accent-sensitive", () => {
    const forward = compareScalars("é", "e", US);
    expect(forward).not.toBe(0);
    expect(compareScalars("e", "é", US)).toBe(forward === 1 ? -1 : 1);
    expect(compareScalars("ä", "a", compareSettings(DE_DE))).not.toBe(0);
  });

  it("collates in the locale's own order (Czech files \"ch\" after \"h\")", () => {
    expect(compareScalars("ch", "h", US)).toBe(-1);
    expect(compareScalars("ch", "h", compareSettings(SPREADSHEET_LOCALES["cs-CZ"]))).toBe(1);
  });

  it("is case-sensitive when the settings say so (sensitivity \"variant\")", () => {
    const sensitive = compareSettings(EN_US, true);
    const order = compareScalars("a", "A", sensitive);
    expect(order).not.toBe(0);
    expect(order).toBe(Math.sign(new Intl.Collator("en-US", { sensitivity: "variant" }).compare("a", "A")));
    expect(compareScalars("A", "a", sensitive)).toBe(order === 1 ? -1 : 1);
    expect(compareScalars("a", "a", sensitive)).toBe(0);
    expect(compareScalars("é", "e", sensitive)).not.toBe(0);
    expect(compareScalars("a", "B", sensitive)).toBe(-1);
    expect(compareScalars(EMPTY, "", sensitive)).toBe(0);
  });
});

describe("toCellValue", () => {
  it("stores EMPTY as 0 and −0 as 0", () => {
    expect(Object.is(toCellValue(EMPTY), 0)).toBe(true);
    expect(Object.is(toCellValue(-0), 0)).toBe(true);
  });

  it.each<[string, ScalarValue, CellValue]>([
    ["Infinity", Infinity, err("#NUM!")],
    ["-Infinity", -Infinity, err("#NUM!")],
    ["NaN", NaN, err("#NUM!")],
    ["32,768 characters", "x".repeat(MAX_TEXT_LENGTH + 1), err("#VALUE!")],
    ["a number", 1.5, 1.5],
    ["text", "a", "a"],
    ["\"\"", "", ""],
    ["TRUE", true, true],
    ["FALSE", false, false],
    ["an error", err("#N/A"), err("#N/A")],
  ])("stores %s as %j", (_label, value, expected) => {
    expect(toCellValue(value)).toEqual(expected);
  });

  it("keeps text of exactly 32,767 characters", () => {
    const text = "x".repeat(MAX_TEXT_LENGTH);
    expect(toCellValue(text)).toBe(text);
  });

  it("copies an error to a fresh object", () => {
    const error = err("#DIV/0!");
    const stored = toCellValue(error);
    expect(stored).toEqual({ error: "#DIV/0!" });
    expect(stored).not.toBe(error);
  });
});

describe("power", () => {
  it.each<[number, number, number | CellError]>([
    // =0^0 is #NUM! (Excel).
    [0, 0, err("#NUM!")],
    [0, -1, err("#DIV/0!")],
    // A negative base with a non-integer exponent.
    [-8, 1 / 3, err("#NUM!")],
    [-1, 0.5, err("#NUM!")],
    [-2, 2, 4],
    [-2, 3, -8],
    [-8, 2, 64],
    [2, -1, 0.5],
    [0, 2, 0],
    [4, 0.5, 2],
    [2, 10, 1024],
    // A non-finite result.
    [10, 400, err("#NUM!")],
  ])("%s ^ %s → %j", (x, y, expected) => {
    expect(power(x, y)).toEqual(expected);
  });
});
