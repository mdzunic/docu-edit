// SPEC-122 "Input parsing": parseInput's rules table in all six SPEC-118
// locales (each locale's own separators, currency and date order), the
// 15-significant-digit truncation, format steering (an existing "@" format,
// a percent-formatted cell), the system's date range, and the current year
// read from core now() (pinned with setClock, restored after each test).
//
// Expected serials are computed by hand: 2026-10-08 is 46303 in the 1900
// system (1462 less in 1904), 2026-10-01 is 46296, 13:45 is 49500/86400 of
// a day.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { setClock } from "../../core/clock";
import type { CellValue } from "../model";
import { SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocale, type SpreadsheetLocaleTag } from "./locale";
import { parseInput, type ParsedInput } from "./input";

const OCT_8_2026 = 46303;
const OCT_1_2026 = 46296;
const T_13_45 = 49_500 / 86_400;
const T_13_45_30 = 49_530 / 86_400;

beforeEach(() => {
  setClock(new Date(2026, 9, 8, 9, 30));
});

afterEach(() => {
  setClock(null);
});

/** The currency formats each locale's typed currency suggests. */
const CURRENCY: Record<SpreadsheetLocaleTag, [string, string]> = {
  "en-US": ['"$"#,##0', '"$"#,##0.00'],
  "en-GB": ['"£"#,##0', '"£"#,##0.00'],
  "de-DE": ['#,##0" €"', '#,##0.00" €"'],
  "fr-FR": ['#,##0" €"', '#,##0.00" €"'],
  "es-ES": ['#,##0" €"', '#,##0.00" €"'],
  "cs-CZ": ['#,##0" Kč"', '#,##0.00" Kč"'],
};

type Expected = { value: CellValue; numberFormat?: string; builtinFormatId?: number; formula?: string; quotePrefix?: true };

/** The rules table as `locale` spells each input. */
function rules(locale: SpreadsheetLocale): [string, Expected][] {
  const d = locale.decimal;
  // What a person types for a space-like group: a plain space.
  const g = locale.group === " " || locale.group === " " ? " " : locale.group;
  const money = (amount: string, sign = ""): string =>
    locale.currencyPosition === "prefix" ? `${sign}${locale.currencySymbol}${amount}` : `${sign}${amount} ${locale.currencySymbol}`;
  const sep = locale.dateSeparator;
  const date = (month: number, day: number, year: string, separator = sep): string =>
    locale.dateOrder === "MDY" ? `${month}${separator}${day}${separator}${year}` : `${day}${separator}${month}${separator}${year}`;
  const [whole, cents] = CURRENCY[locale.tag];
  return [
    ["", { value: null }],
    ["=A2*2", { value: null, formula: "A2*2" }],
    ["=SUM(B1:B2)", { value: null, formula: "SUM(B1:B2)" }],
    ["'0012", { value: "0012", quotePrefix: true }],
    ["'=A1", { value: "=A1", quotePrefix: true }],
    ["42", { value: 42 }],
    [`-1${d}5`, { value: -1.5 }],
    ["+7", { value: 7 }],
    ["007", { value: 7 }],
    [`1${g}234${g}567`, { value: 1_234_567, numberFormat: "#,##0" }],
    [`1${g}234${d}5`, { value: 1234.5, numberFormat: "#,##0.00" }],
    ["(12)", { value: -12 }],
    ["10%", { value: 0.1, numberFormat: "0%" }],
    [`12${d}5%`, { value: 0.125, numberFormat: "0.00%" }],
    [money(`1${g}200`), { value: 1200, numberFormat: whole }],
    [money(`1${g}200${d}50`), { value: 1200.5, numberFormat: cents }],
    [money("5", "-"), { value: -5, numberFormat: whole }],
    [`(${money("5")})`, { value: -5, numberFormat: whole }],
    [`1${d}5E3`, { value: 1500, numberFormat: "0.00E+00" }],
    ["1e-4", { value: 0.0001, numberFormat: "0.00E+00" }],
    ["1 1/2", { value: 1.5, numberFormat: "# ?/?" }],
    ["0 3/16", { value: 0.1875, numberFormat: "# ??/??" }],
    [date(10, 8, "2026"), { value: OCT_8_2026, numberFormat: "m/d/yyyy", builtinFormatId: 14 }],
    [date(10, 8, "26", "-"), { value: OCT_8_2026, numberFormat: "m/d/yyyy", builtinFormatId: 14 }],
    ["2026-10-08", { value: OCT_8_2026, numberFormat: "m/d/yyyy", builtinFormatId: 14 }],
    ["8-Oct-2026", { value: OCT_8_2026, numberFormat: "d-mmm-yy" }],
    ["Oct 8, 2026", { value: OCT_8_2026, numberFormat: "d-mmm-yy" }],
    ["8 October 2026", { value: OCT_8_2026, numberFormat: "d-mmm-yy" }],
    ["8-Oct", { value: OCT_8_2026, numberFormat: "d-mmm" }],
    ["Oct 8", { value: OCT_8_2026, numberFormat: "d-mmm" }],
    ["Oct-2026", { value: OCT_1_2026, numberFormat: "mmm-yy" }],
    ["Oct 2026", { value: OCT_1_2026, numberFormat: "mmm-yy" }],
    ["13:45", { value: T_13_45, numberFormat: "h:mm" }],
    ["13:45:30", { value: T_13_45_30, numberFormat: "h:mm:ss" }],
    ["1:45 PM", { value: T_13_45, numberFormat: "h:mm AM/PM" }],
    ["1:45:30 pm", { value: T_13_45_30, numberFormat: "h:mm:ss AM/PM" }],
    ["1 PM", { value: 13 / 24, numberFormat: "h:mm AM/PM" }],
    ["36:00", { value: 1.5, numberFormat: "[h]:mm:ss" }],
    ["25:30:00", { value: 1.0625, numberFormat: "[h]:mm:ss" }],
    [`${date(10, 8, "2026")} 13:45`, { value: OCT_8_2026 + T_13_45, numberFormat: "m/d/yyyy h:mm", builtinFormatId: 22 }],
    ["2026-10-08T13:45", { value: OCT_8_2026 + T_13_45, numberFormat: "m/d/yyyy h:mm", builtinFormatId: 22 }],
    ["TRUE", { value: true }],
    ["false", { value: false }],
    ["#N/A", { value: { error: "#N/A" } }],
    ["#div/0!", { value: { error: "#DIV/0!" } }],
    ["#getting_data", { value: { error: "#GETTING_DATA" } }],
    ["Revenue", { value: "Revenue" }],
  ];
}

function expectParsed(actual: ParsedInput, expected: Expected): void {
  const { value, ...rest } = actual;
  if (typeof expected.value === "number" && !Number.isInteger(expected.value)) {
    expect(typeof value).toBe("number");
    expect(value as number).toBeCloseTo(expected.value, 10);
  } else {
    expect(value).toEqual(expected.value);
  }
  const { value: _expectedValue, ...expectedRest } = expected;
  expect(rest).toEqual(expectedRest);
}

describe("parseInput — the rules table in every locale", () => {
  const cases = SPREADSHEET_LOCALE_TAGS.flatMap((tag) => rules(SPREADSHEET_LOCALES[tag]).map(([text, expected]) => [tag, text, expected] as const));
  it.each(cases)("%s %j", (tag, text, expected) => {
    expectParsed(parseInput(text, SPREADSHEET_LOCALES[tag], "1900"), expected);
  });

  it("accepts every space-like group separator where the locale's group is one", () => {
    for (const tag of ["fr-FR", "cs-CZ"] as const) {
      for (const space of [" ", " ", " "]) {
        expect(parseInput(`1${space}234${space}567`, SPREADSHEET_LOCALES[tag], "1900")).toEqual({ value: 1_234_567, numberFormat: "#,##0" });
      }
    }
  });

  it("reads the locale's own month names", () => {
    expect(parseInput("8. Oktober 2026", SPREADSHEET_LOCALES["de-DE"], "1900")).toEqual({ value: OCT_8_2026, numberFormat: "d-mmm-yy" });
    expect(parseInput("8 octobre 2026", SPREADSHEET_LOCALES["fr-FR"], "1900")).toEqual({ value: OCT_8_2026, numberFormat: "d-mmm-yy" });
  });

  it("takes a yearless numeric date in the current year as d-mmm", () => {
    expect(parseInput("10/8", SPREADSHEET_LOCALES["en-US"], "1900")).toEqual({ value: OCT_8_2026, numberFormat: "d-mmm" });
    expect(parseInput("8.10", SPREADSHEET_LOCALES["de-DE"], "1900")).toEqual({ value: OCT_8_2026, numberFormat: "d-mmm" });
  });

  it("reads a month-name date with a time as a date and time", () => {
    expect(parseInput("Oct 8, 2026 1:45 PM", SPREADSHEET_LOCALES["en-US"], "1900").builtinFormatId).toBe(22);
  });
});

describe("parseInput — what is not a value", () => {
  const EN_US = SPREADSHEET_LOCALES["en-US"];

  it.each([
    "+A1",
    "-A1",
    "=",
    "1,0000",
    "1.234,5",
    "$4,00",
    "Oct",
    "1/1/1899",
    "1 1/0",
    "1 1/1000",
    "12:75",
    "   ",
  ])("%j is text as typed", (text) => {
    expect(parseInput(text, EN_US, "1900")).toEqual({ value: text });
  });

  it("never starts a formula with a leading sign", () => {
    expect(parseInput("+1", EN_US, "1900")).toEqual({ value: 1 });
    expect(parseInput("-A1+1", EN_US, "1900")).toEqual({ value: "-A1+1" });
  });

  it("keeps a lone apostrophe as an empty quote-prefixed text", () => {
    expect(parseInput("'", EN_US, "1900")).toEqual({ value: "", quotePrefix: true });
  });

  it("ignores spaces around a literal and a number, keeping text as typed", () => {
    expect(parseInput(" true ", EN_US, "1900")).toEqual({ value: true });
    expect(parseInput("  42 ", EN_US, "1900")).toEqual({ value: 42 });
    expect(parseInput(" hello ", EN_US, "1900")).toEqual({ value: " hello " });
  });
});

describe("parseInput — dates and the date system", () => {
  const EN_US = SPREADSHEET_LOCALES["en-US"];

  it("keeps dates inside the system's range and makes the rest text", () => {
    expect(parseInput("2/29/1900", EN_US, "1900")).toEqual({ value: 60, numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(parseInput("12/31/9999", EN_US, "1900")).toEqual({ value: 2_958_465, numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(parseInput("12/31/1899", EN_US, "1900")).toEqual({ value: "12/31/1899" });
    expect(parseInput("1/1/1903", EN_US, "1904")).toEqual({ value: "1/1/1903" });
  });

  it("gives serials in the workbook's date system", () => {
    expect(parseInput("10/8/2026", EN_US, "1904")).toEqual({ value: OCT_8_2026 - 1462, numberFormat: "m/d/yyyy", builtinFormatId: 14 });
  });

  it("reads the current year from core now()", () => {
    setClock(new Date(2027, 0, 15, 12, 0));
    expect(parseInput("8-Oct", EN_US, "1900")).toEqual({ value: OCT_8_2026 + 365, numberFormat: "d-mmm" });
    expect(parseInput("Oct 8", EN_US, "1900")).toEqual({ value: OCT_8_2026 + 365, numberFormat: "d-mmm" });
    expect(parseInput("10/8", EN_US, "1900")).toEqual({ value: OCT_8_2026 + 365, numberFormat: "d-mmm" });
  });
});

describe("parseInput — 15 significant digits", () => {
  const EN_US = SPREADSHEET_LOCALES["en-US"];

  it("turns every digit after the 15th significant one into 0, as Excel does", () => {
    expect(parseInput("1234567890123456789", EN_US, "1900")).toEqual({ value: 1_234_567_890_123_450_000 });
    expect(parseInput("1,234,567,890,123,456,789", EN_US, "1900")).toEqual({ value: 1_234_567_890_123_450_000, numberFormat: "#,##0" });
    expect(parseInput("0.0001234567890123456789", EN_US, "1900")).toEqual({ value: 0.000_123_456_789_012_345 });
    expect(parseInput("12345678901234567.5", EN_US, "1900")).toEqual({ value: 12_345_678_901_234_500 });
    expect(parseInput("-1234567890123456789%", EN_US, "1900")).toEqual({ value: -12_345_678_901_234_500, numberFormat: "0%" });
    expect(parseInput("1.2345678901234567E5", EN_US, "1900")).toEqual({ value: 123_456.789_012_345, numberFormat: "0.00E+00" });
  });

  it("keeps 15 digits and fewer as typed", () => {
    expect(parseInput("123456789012345", EN_US, "1900")).toEqual({ value: 123_456_789_012_345 });
    expect(parseInput("0.000000000000000000001", EN_US, "1900")).toEqual({ value: 1e-21 });
  });

  it("reads 1,234 in de-DE as 1.234 and truncates there too", () => {
    expect(parseInput("1,234", SPREADSHEET_LOCALES["de-DE"], "1900")).toEqual({ value: 1.234 });
    expect(parseInput("1.234.567.890.123.456.789", SPREADSHEET_LOCALES["de-DE"], "1900")).toEqual({
      value: 1_234_567_890_123_450_000,
      numberFormat: "#,##0",
    });
  });
});

describe("parseInput — format steering", () => {
  const EN_US = SPREADSHEET_LOCALES["en-US"];

  it("stores text verbatim in a Text-formatted cell", () => {
    for (const text of ["=A1", "42", "'x", "TRUE", "10/8/2026", " 7 "]) {
      expect(parseInput(text, EN_US, "1900", { existingFormat: "@" })).toEqual({ value: text });
    }
    expect(parseInput("", EN_US, "1900", { existingFormat: "@" })).toEqual({ value: null });
  });

  it("divides a plain number by 100 in a percent-formatted cell (automatic percent entry)", () => {
    for (const code of ["0%", "0.00%", "#,##0.0%;[Red]-#,##0.0%"]) {
      expect(parseInput("10", EN_US, "1900", { existingFormat: code })).toEqual({ value: 0.1 });
      expect(parseInput("-12.5", EN_US, "1900", { existingFormat: code })).toEqual({ value: -0.125 });
      expect(parseInput("1,000", EN_US, "1900", { existingFormat: code })).toEqual({ value: 10, numberFormat: "#,##0" });
      // Below 1 it already is the fraction; a typed % or currency is kept.
      expect(parseInput("0.5", EN_US, "1900", { existingFormat: code })).toEqual({ value: 0.5 });
      expect(parseInput("10%", EN_US, "1900", { existingFormat: code })).toEqual({ value: 0.1, numberFormat: "0%" });
      expect(parseInput("$5", EN_US, "1900", { existingFormat: code })).toEqual({ value: 5, numberFormat: '"$"#,##0' });
    }
    expect(parseInput("1.1", EN_US, "1900", { existingFormat: "0%" })).toEqual({ value: 0.011 });
  });

  it("leaves numbers alone in any other format", () => {
    for (const code of ["General", "0.00", "m/d/yyyy", '"%"0']) {
      expect(parseInput("10", EN_US, "1900", { existingFormat: code })).toEqual({ value: 10 });
    }
  });

  it("still suggests a format whatever the cell has (the caller applies it to General only)", () => {
    expect(parseInput("10/8/2026", EN_US, "1900", { existingFormat: "0.00" })).toEqual({
      value: OCT_8_2026,
      numberFormat: "m/d/yyyy",
      builtinFormatId: 14,
    });
  });
});
