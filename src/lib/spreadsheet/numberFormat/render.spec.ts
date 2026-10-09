// SPEC-118 "Sections", "Rendering numbers", "Rendering dates and times",
// "Text, booleans and errors", "General" (the non-General fit) and
// "editText": formatValue as a table of (value, code, locale, width) → the
// text, colour, fill and overflow of every example in the Reference — the
// sign rule, the six locales' separators, the "########" overflow and its
// width, booleans and error codes in English in de-DE and fr-FR, formula
// text never localised, a code that is invalid or unsupported rendering as
// General — then alignHint, plainText and editText (what the formula bar
// shows).
import { describe, expect, it } from "vitest";
import type { CellStyle, CellValue } from "../model";
import { dateToSerial } from "../serial";
import { INDEXED_COLORS } from "../styles";
import { localeNames, SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "./locale";
import { editText, formatGeneral, formatValue, plainText, type FormatContext, type FormattedValue } from "./render";

const RED = "#ff0000";
const BLUE = "#0000ff";
const FIGURE_SPACE = "\u2007";

/** 2026-10-08 13:45:00 — a Thursday; the fixture's A2:A4. */
const OCT_8_2026_1345 = 46303.5729166667;
const OCT_8_2026 = 46303;

/** A 1900-system serial. */
function serial(year: number, month: number, day: number): number {
  return dateToSerial(year, month, day, "1900");
}

/** Where a value renders: the locale (default en-US) and the context's
 * other fields. */
interface Where {
  locale?: SpreadsheetLocaleTag;
  widthPx?: number;
  digitWidthPx?: number;
  dateSystem?: "1900" | "1904";
  dateQuirks?: "excel" | "proleptic";
  palette?: readonly string[];
}

function context({ locale = "en-US", dateSystem = "1900", ...rest }: Where = {}): FormatContext {
  return { locale: SPREADSHEET_LOCALES[locale], dateSystem, ...rest };
}

/** The parts of a FormattedValue the table pins (alignHint has its own block). */
interface Shown {
  text: string;
  color?: string;
  fill?: { char: string; index: number };
  overflow?: "hash";
}

function shown(fv: FormattedValue): Shown {
  return { text: fv.text, color: fv.color, fill: fv.fill, overflow: fv.overflow };
}

/** `#` × n, the overflow fill. */
function hash(n: number): Shown {
  return { text: "#".repeat(n), overflow: "hash" };
}

/** (value, code, expected text or shown parts, where). */
type Row = [value: CellValue, code: string | undefined, expected: string | Shown, where?: Where];

function check(value: CellValue, code: string | undefined, expected: string | Shown, where?: Where): void {
  const want: Shown = typeof expected === "string" ? { text: expected } : expected;
  expect(shown(formatValue(value, code, context(where)))).toEqual(want);
}

const TITLE = "%j through %j → %j";

// "Sections"
describe("sections and the sign rule", () => {
  it.each<Row>([
    // Without conditions: one section → every number, with the automatic "-".
    [5, "0.0", "5.0"],
    [-5, "0.0", "-5.0"],
    // Two → (≥ 0; < 0); the classic negative section writes its own sign.
    [3, "0;(0)", "3"],
    [0, "0;(0)", "0"],
    [-3, "0;(0)", "(3)"],
    [-1, "0.00;0.00", "1.00"],
    [-5, "0;[Red]0", { text: "5", color: RED }],
    [-1234.5, "#,##0.00_);[Red]\\(#,##0.00\\)", { text: "(1,234.50)", color: RED }],
    [-1, "0.00;[Red]-0.00", { text: "-1.00", color: RED }],
    // Three → (> 0; < 0; = 0).
    [5, '0;(0);"zero"', "5"],
    [-5, '0;(0);"zero"', "(5)"],
    [0, '0;(0);"zero"', "zero"],
    // A fourth section is for text …
    ["x", '0;0;0;"t: "@', "t: x"],
    // … as is, with fewer, a final section containing @ — then the numeric
    // sections are the ones before it.
    ["abc", '0.00;"text: "@', "text: abc"],
    [-5, '0.00;"text: "@', "-5.00"],
    [-5, "0;0;@", "5"],
    // Conditions: the first that holds wins; the next unconditioned numeric
    // section is the else; every section but an unconditioned code's second
    // gets the automatic "-".
    [150, "[>100][Red]0;[Blue]0", { text: "150", color: RED }],
    [5, "[>100][Red]0;[Blue]0", { text: "5", color: BLUE }],
    [-5, "[>100][Red]0;[Blue]0", { text: "-5", color: BLUE }],
    [-150, '[<-100]"low";[>100]"high";0', "-low"],
    [150, '[<-100]"low";[>100]"high";0', "high"],
    [5, '[<-100]"low";[>100]"high";0', "5"],
    [-5, '[<-100]"low";[>100]"high";0', "-5"],
    // No section for the value → the # overflow.
    [5, '[>100]"big";[<-100]"small"', hash(8), { widthPx: 64 }],
    [5, '[>100]"big";[<-100]"small"', hash(3), { widthPx: 26 }],
    // An empty section renders "" — `;;;` hides everything.
    [5, ";;;", ""],
    [-5, ";;;", ""],
    [0, ";;;", ""],
    ["Bob", ";;;", ""],
    [-5, "0;;0", ""],
    // A section of only literals renders the literals.
    [5, '"pos";"neg";"zero"', "pos"],
    [-5, '"pos";"neg";"zero"', "neg"],
    [0, '"pos";"neg";"zero"', "zero"],
  ])(TITLE, check);
});

// "Rendering numbers"
describe("rendering numbers", () => {
  describe("1. normalising: %, scaling commas, 15 significant digits", () => {
    it.each<Row>([
      [0.1 + 0.2, "0.0000000000000000", "0.3000000000000000"],
      [0.1 + 0.2, undefined, "0.3"],
      [0.125, "0%", "13%"],
      [0.125, "0.00%", "12.50%"],
      [-0.125, "0.0%", "-12.5%"],
      [0.5, "0%%", "5000%%"],
      [1234567, '#,##0,"K"', "1,235K"],
      [999, "#,##0,", "1"],
      [2500, "0,", "3"],
      [1500000, "0,,", "2"],
      [123456789, "#,##0,,", "123"],
      // A comma directly after the last integer placeholder, before the point.
      [1234567, "0,.00", "1234.57"],
    ])(TITLE, check);
  });

  describe("2. rounding half away from zero on the decimal text", () => {
    it.each<Row>([
      [2.675, "0.00", "2.68"],
      [1.005, "0.00", "1.01"],
      [-0.001, "0.00", "-0.00"],
      [-2.675, "0.00", "-2.68"],
      [1.25, "0.0", "1.3"],
      [0.5, "0", "1"],
      [-0.5, "0", "-1"],
      [2.5, "0", "3"],
      [9.995, "0.00", "10.00"],
    ])(TITLE, check);
  });

  describe("3. the integer part", () => {
    it.each<Row>([
      [123456789, "000-00-0000", "123-45-6789"],
      [12, "0000", "0012"],
      [5, "???", `${FIGURE_SPACE}${FIGURE_SPACE}5`],
      [12345, "0", "12345"],
      // Extra leading digits precede the first placeholder.
      [12345, '"x"0-0', "x1234-5"],
      [1234567.891, "#,##0.00", "1,234,567.89"],
      [123, "#,##0", "123"],
      [0, "#,##0", "0"],
      [-1234567, "#,##0", "-1,234,567"],
      // An integer part of 0 under # only renders nothing.
      [0.5, "#.##", ".5"],
      [0, "#", ""],
    ])(TITLE, check);
  });

  describe("4. the fraction part", () => {
    it.each<Row>([
      [1.5, "0.000", "1.500"],
      [1.5, "0.0#", "1.5"],
      [1.5, "0.###", "1.5"],
      [1.5, "0.0?", `1.5${FIGURE_SPACE}`],
      [1.5, "0.???", `1.5${FIGURE_SPACE}${FIGURE_SPACE}`],
      // The point shows whenever the code has one.
      [5, "#.##", "5."],
      [5, "0.#", "5."],
    ])(TITLE, check);
  });

  describe("5. scientific", () => {
    it.each<Row>([
      [12345, "##0.0E+0", "12.3E+3"],
      [1234567, "##0.0E+0", "1.2E+6"],
      [0.00012345, "##0.0E+0", "123.5E-6"],
      [12345, "0.00E+00", "1.23E+04"],
      [-12345, "0.00E+00", "-1.23E+04"],
      [0.00012345, "0.00E+00", "1.23E-04"],
      [0, "0.00E+00", "0.00E+00"],
      [9.999, "0.00E+00", "1.00E+01"],
      // The exponent is padded to its 0 count.
      [12345, "0.0E+000", "1.2E+004"],
      [12345, "0E+0", "1E+4"],
      // E+ always signs, E- only negatives; the case is the code's.
      [12345, "0.00E-00", "1.23E04"],
      [0.00012345, "0.00E-00", "1.23E-04"],
      [12345, "0.00e+00", "1.23e+04"],
    ])(TITLE, check);
  });

  describe("6. fractions", () => {
    it.each<Row>([
      // A fixed denominator, never reduced.
      [0.3, "?/8", "2/8"],
      [0.5, "?/8", "4/8"],
      [1.5, "# ?/8", "1 4/8"],
      [0.25, "??/100", "25/100"],
      // Bounded by 10^k − 1: the best rational approximation.
      [Math.PI, "# ??/??", "3 14/99"],
      [1.5, "# ?/?", "1 1/2"],
      [1.5, "?/?", "3/2"],
      [-1.5, "# ?/?", "-1 1/2"],
      [0.333, "?/?", "1/3"],
      [0.5, "# ?/?", " 1/2"],
      // 0.3875 lies midway between 3/8 and 2/5: a tie goes to the smaller
      // denominator.
      [0.3875, "?/?", "2/5"],
      [1.3875, "# ?/?", "1 2/5"],
      // A zero fraction with an integer part: numerator, slash and
      // denominator become spaces of the same length; exact 0 → "0" plus them.
      [2, "# ?/?", "2    "],
      [2, "# ??/??", "2      "],
      [0, "# ?/?", "0    "],
    ])(TITLE, check);
  });

  describe("7. General inside a numeric section", () => {
    it.each<Row>([
      [-5, "[Red]General", { text: "-5", color: RED }],
      [1234.5, "[Red]General", { text: "1234.5", color: RED }],
      [123456789012, "[Red]General", { text: "1.23457E+11", color: RED }],
      [5, '"Total: "General', "Total: 5"],
      [-5, "General;[Red]-General", { text: "-5", color: RED }],
      [1234.5, "General", "1234.5"],
    ])(TITLE, check);
  });

  describe("8. _x, *x, \\x and quoted text", () => {
    it.each<Row>([
      [5, "0_)", "5 "],
      [5, "_(0_)", " 5 "],
      [5, "0\\x", "5x"],
      [5, '0" units"', "5 units"],
      [5, "*-0", { text: "5", fill: { char: "-", index: 0 } }],
      [5, "0*x", { text: "5", fill: { char: "x", index: 1 } }],
      [5, '"$"*-0', { text: "$5", fill: { char: "-", index: 1 } }],
      // The last fill in a section wins.
      [5, "*a0*b", { text: "5", fill: { char: "b", index: 1 } }],
      // Accounting (built-in 44): `_("$"* #,##0.00_)` and its zero section.
      [1234.5, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', { text: " $1,234.50 ", fill: { char: " ", index: 2 } }],
      [-1234.5, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', { text: " $(1,234.50)", fill: { char: " ", index: 2 } }],
      [0, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', { text: ` $-${FIGURE_SPACE}${FIGURE_SPACE} `, fill: { char: " ", index: 2 } }],
    ])(TITLE, check);
  });

  describe("9. colours", () => {
    it.each<[name: string, hex: string]>([
      ["Black", "#000000"],
      ["Blue", "#0000ff"],
      ["Cyan", "#00ffff"],
      ["Green", "#00ff00"],
      ["Magenta", "#ff00ff"],
      ["Red", "#ff0000"],
      ["White", "#ffffff"],
      ["Yellow", "#ffff00"],
    ])("[%s] → %s", (name, hex) => {
      check(5, `[${name}]0`, { text: "5", color: hex });
      check(5, `[${name.toLowerCase()}]0`, { text: "5", color: hex });
    });

    it("[ColorN] → palette[N + 7], the default palette INDEXED_COLORS", () => {
      check(5, "[Color1]0", { text: "5", color: "#000000" });
      check(5, "[Color10]0", { text: "5", color: "#008000" });
      check(5, "[Color56]0", { text: "5", color: "#333333" });
      for (let n = 1; n <= 56; n += 1) {
        expect(formatValue(5, `[Color${n}]0`, context()).color).toBe(INDEXED_COLORS[n + 7]);
      }
    });

    it("[ColorN] → palette[N + 7] of a custom palette", () => {
      const palette = Array.from({ length: 64 }, (_, i) => `#0000${i.toString(16).padStart(2, "0")}`);
      check(5, "[Color3]0", { text: "5", color: "#00000a" }, { palette });
      check(5, "[Color56]0", { text: "5", color: "#00003f" }, { palette });
      // The named colours are fixed, whatever the palette.
      check(5, "[Red]0", { text: "5", color: RED }, { palette });
    });
  });
});

// "Locales" + "Codes are stored and parsed in the invariant syntax"
describe("the six locales' separators", () => {
  it.each<[tag: SpreadsheetLocaleTag, grouped: string, general: string, scientific: string, percent: string]>([
    ["en-US", "1,234,567.89", "1234.5", "1.23E+04", "12.5%"],
    ["en-GB", "1,234,567.89", "1234.5", "1.23E+04", "12.5%"],
    ["de-DE", "1.234.567,89", "1234,5", "1,23E+04", "12,5%"],
    ["fr-FR", "1\u202f234\u202f567,89", "1234,5", "1,23E+04", "12,5%"],
    ["es-ES", "1.234.567,89", "1234,5", "1,23E+04", "12,5%"],
    ["cs-CZ", "1\u00a0234\u00a0567,89", "1234,5", "1,23E+04", "12,5%"],
  ])("%s: #,##0.00 → %j, General → %j, 0.00E+00 → %j, 0.0%% → %j", (tag, grouped, general, scientific, percent) => {
    check(1234567.891, "#,##0.00", grouped, { locale: tag });
    check(1234.5, undefined, general, { locale: tag });
    check(1234.5, "General", general, { locale: tag });
    check(12345, "0.00E+00", scientific, { locale: tag });
    check(0.125, "0.0%", percent, { locale: tag });
  });

  it.each<Row>([
    // The fixture's B5 and the LibreOffice red negative, in de-DE.
    [1234.5, "#,##0.00", "1.234,50", { locale: "de-DE" }],
    [-1, "0.00;[Red]-0.00", { text: "-1,00", color: RED }, { locale: "de-DE" }],
    // `/` and `:` are literal: no date-separator substitution.
    [OCT_8_2026_1345, "m/d/yyyy h:mm", "10/8/2026 13:45", { locale: "de-DE" }],
    [OCT_8_2026_1345, "dd.mm.yyyy hh:mm", "08.10.2026 13:45", { locale: "fr-FR" }],
    [1.5, "# ?/?", "1 1/2", { locale: "fr-FR" }],
  ])(TITLE, check);
});

// "Rendering dates and times"
describe("rendering dates and times", () => {
  describe("rounding to the smallest shown unit, carrying into the date", () => {
    it.each<Row>([
      [0.999999, "h:mm", "0:00"],
      [46303.9999999, "yyyy-mm-dd hh:mm", "2026-10-09 00:00"],
      [59.6 / 86_400, "h:mm:ss", "0:01:00"],
      [59.96 / 86_400, "h:mm:ss.0", "0:01:00.0"],
      // Fractional seconds: .0/.00/.000 after the seconds.
      [0.5 + 0.5 / 86_400, "hh:mm:ss.00", "12:00:00.50"],
      [1.234 / 86_400, "ss.000", "01.234"],
      [62.3 / 86_400, "mmss.0", "0102.3"],
    ])(TITLE, check);
  });

  describe("the tokens", () => {
    it.each<Row>([
      [OCT_8_2026_1345, "yyyy-mm-dd hh:mm:ss", "2026-10-08 13:45:00"],
      [OCT_8_2026_1345, "m/d/yyyy", "10/8/2026"],
      [OCT_8_2026_1345, "d/m/yy", "8/10/26"],
      [OCT_8_2026_1345, "y", "26"],
      [OCT_8_2026_1345, "yyy", "2026"],
      [OCT_8_2026_1345, "mmm", "Oct"],
      [OCT_8_2026_1345, "mmmm", "October"],
      [OCT_8_2026_1345, "mmmmm", "O"],
      [OCT_8_2026_1345, "mmmmmm", "October"],
      [OCT_8_2026_1345, "d", "8"],
      [OCT_8_2026_1345, "dd", "08"],
      [OCT_8_2026_1345, "ddd", "Thu"],
      [OCT_8_2026_1345, "dddd", "Thursday"],
      [OCT_8_2026_1345, "ddddd", "Thursday"],
      [OCT_8_2026_1345, "h", "13"],
      [OCT_8_2026_1345, "hh:mm", "13:45"],
      [0.25 + 5 / 86_400, "h:m:s", "6:0:5"],
      [0.25 + 5 / 86_400, "hh:mm:ss", "06:00:05"],
      // The fixture's ids 14 and 22 and the [$-F800] long date.
      [OCT_8_2026_1345, "m/d/yyyy h:mm", "10/8/2026 13:45"],
      [OCT_8_2026_1345, "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", "Thursday, October 8, 2026"],
    ])(TITLE, check);
  });

  describe("elapsed [h], [m], [s]", () => {
    it.each<Row>([
      [1.5, "[h]:mm:ss", "36:00:00"],
      [1.5, "[hh]:mm:ss", "36:00:00"],
      [1.5, "[mm]:ss", "2160:00"],
      [1.5, "[ss]", "129600"],
      [0.25, "[hh]:mm", "06:00"],
      [0.25, "[h]:mm", "6:00"],
      [1.99, "[h]", "47"],
      [1 + 1 / 24 + 30 / 1440, "[h]:mm", "25:30"],
      // A negative value under an elapsed token → hash, in either calendar.
      [-0.25, "[h]:mm", hash(8), { widthPx: 64 }],
      [-0.25, "[h]:mm", hash(8), { widthPx: 64, dateQuirks: "proleptic" }],
    ])(TITLE, check);
  });

  describe("AM/PM and A/P keep the code's case, in every locale", () => {
    it.each<Row>([
      [OCT_8_2026_1345, "h:mm AM/PM", "1:45 PM"],
      [OCT_8_2026_1345, "h:mm am/pm", "1:45 pm"],
      [OCT_8_2026_1345, "h:mm A/P", "1:45 P"],
      [OCT_8_2026_1345, "h:mm a/p", "1:45 p"],
      [0.25, "hh:mm AM/PM", "06:00 AM"],
      [0, "h:mm AM/PM", "12:00 AM"],
      [0.5, "h:mm AM/PM", "12:00 PM"],
      [OCT_8_2026_1345, "h:mm AM/PM", "1:45 PM", { locale: "de-DE" }],
      [OCT_8_2026_1345, "h:mm am/pm", "1:45 pm", { locale: "fr-FR" }],
      [OCT_8_2026_1345, "h:mm:ss A/P", "1:45:00 P", { locale: "cs-CZ" }],
    ])(TITLE, check);
  });

  describe("Excel's calendar (dateQuirks excel, the default)", () => {
    it.each<Row>([
      [60, "d-mmm-yyyy", "29-Feb-1900"],
      [59, "d-mmm-yyyy", "28-Feb-1900"],
      [61, "d-mmm-yyyy", "1-Mar-1900"],
      [0, "yyyy-mm-dd", "1900-01-00"],
      [0, "m/d/yyyy", "1/0/1900"],
      [0, "d-mmm-yyyy", "0-Jan-1900"],
      // Weekdays: serial 1 is a Sunday, serial 60 a Wednesday.
      [1, "dddd", "Sunday"],
      [60, "dddd", "Wednesday"],
      [61, "dddd", "Thursday"],
      [OCT_8_2026, "dddd", "Thursday"],
      [2958465, "yyyy-mm-dd", "9999-12-31"],
      // Negative or past 9999 → "########".
      [-1, "yyyy-mm-dd", hash(8)],
      [2958466, "yyyy-mm-dd", hash(8)],
      [-1, "yyyy-mm-dd", hash(8), { widthPx: 64 }],
      [-0.5, "h:mm", hash(8), { widthPx: 64 }],
      [-1, "dddd", hash(3), { widthPx: 26 }],
    ])(TITLE, check);
  });

  describe("the proleptic Gregorian calendar of an ods-born workbook", () => {
    const proleptic: Where = { dateQuirks: "proleptic" };
    it.each<Row>([
      [-6, "yyyy-mm-dd", "1899-12-25", proleptic],
      [-6, "dddd", "Monday", proleptic],
      [-1, "yyyy-mm-dd", "1899-12-30", proleptic],
      // days = floor(total / 86 400); the non-negative remainder is the time.
      [-6.25, "yyyy-mm-dd hh:mm", "1899-12-24 18:00", proleptic],
      [0, "yyyy-mm-dd", "1899-12-31", proleptic],
      [1, "dddd", "Monday", proleptic],
      [OCT_8_2026, "dddd, yyyy-mm-dd", "Thursday, 2026-10-08", proleptic],
      // A year below 1 → hash.
      [serial(1, 1, 1), "yyyy-mm-dd", "0001-01-01", proleptic],
      [serial(1, 1, 1) - 1, "yyyy-mm-dd", hash(8), proleptic],
      [2958466, "yyyy-mm-dd", hash(8), proleptic],
    ])(TITLE, check);
  });

  describe("the 1904 date system", () => {
    const system1904: Where = { dateSystem: "1904" };
    it.each<Row>([
      [0, "yyyy-mm-dd", "1904-01-01", system1904],
      [0, "dddd", "Friday", system1904],
      [44841, "yyyy-mm-dd", "2026-10-08", system1904],
      [44841, "dddd", "Thursday", system1904],
      [-1, "yyyy-mm-dd", hash(8), system1904],
    ])(TITLE, check);
  });

  describe("month and day names", () => {
    it.each<Row>([
      // The genitive month when the section also shows a day number.
      [OCT_8_2026, "d. mmmm yyyy", "8. října 2026", { locale: "cs-CZ" }],
      [OCT_8_2026, "mmmm yyyy", "říjen 2026", { locale: "cs-CZ" }],
      [OCT_8_2026, "dd mmmm", "08 října", { locale: "cs-CZ" }],
      [OCT_8_2026, "mmmmm", "ř", { locale: "cs-CZ" }],
      [OCT_8_2026, "dddd d. mmmm yyyy", "Donnerstag 8. Oktober 2026", { locale: "de-DE" }],
      [OCT_8_2026, "ddd mmm", "Do Okt", { locale: "de-DE" }],
      [serial(2026, 3, 1), "mmm mmmm", "Mär März", { locale: "de-DE" }],
      // A section LCID of one of the six takes that locale's names …
      [OCT_8_2026, "[$-407]mmmm", "Oktober"],
      [OCT_8_2026, "[$-407]dddd", "Donnerstag"],
      [OCT_8_2026, "[$-409]mmmm dddd", "October Thursday", { locale: "de-DE" }],
      [OCT_8_2026, "[$-405]d. mmmm", "8. října"],
      // … any other LCID the user's.
      [OCT_8_2026, "[$-411]mmmm", "Oktober", { locale: "de-DE" }],
    ])(TITLE, check);

    it("takes 40C, C0A/40A and 809's names from those locales", () => {
      const sept = serial(2026, 9, 15);
      check(OCT_8_2026, "[$-40C]dddd", localeNames("fr-FR").daysLong[4]!);
      check(OCT_8_2026, "[$-C0A]mmmm", localeNames("es-ES").monthsLong[9]!);
      check(OCT_8_2026, "[$-40A]mmmm", localeNames("es-ES").monthsLong[9]!);
      check(sept, "[$-809]mmm", localeNames("en-GB").monthsShort[8]!);
      check(sept, "mmm", "Sep");
    });
  });

  describe("[$-F800]/[$-x-sysdate] → the locale's longDate, [$-F400]/[$-x-systime] → its longTime", () => {
    it.each<Row>([
      [OCT_8_2026_1345, "[$-F800]dddd, mmmm dd, yyyy", "Thursday, October 8, 2026"],
      [OCT_8_2026_1345, "[$-F800]dddd, mmmm dd, yyyy", "Donnerstag, 8. Oktober 2026", { locale: "de-DE" }],
      [OCT_8_2026_1345, "[$-x-sysdate]dddd, mmmm dd, yyyy", "Donnerstag, 8. Oktober 2026", { locale: "de-DE" }],
      [OCT_8_2026_1345, "[$-F800]dddd, mmmm dd, yyyy", "08 October 2026", { locale: "en-GB" }],
      [OCT_8_2026_1345, "[$-F800]dddd, mmmm dd, yyyy", "jeudi 8 octobre 2026", { locale: "fr-FR" }],
      [OCT_8_2026_1345, "[$-F800]dddd, mmmm dd, yyyy", "jueves, 8 de octubre de 2026", { locale: "es-ES" }],
      [OCT_8_2026_1345, "[$-F800]dddd, mmmm dd, yyyy", "čtvrtek 8. října 2026", { locale: "cs-CZ" }],
      [OCT_8_2026_1345, "[$-F400]h:mm:ss AM/PM", "1:45:00 PM"],
      [OCT_8_2026_1345, "[$-F400]h:mm:ss AM/PM", "13:45:00", { locale: "de-DE" }],
      [OCT_8_2026_1345, "[$-x-systime]h:mm:ss AM/PM", "13:45:00", { locale: "en-GB" }],
    ])(TITLE, check);
  });

  describe("Q, QQ and WW", () => {
    it.each<Row>([
      [OCT_8_2026, "Q yyyy", "Q4 2026"],
      [OCT_8_2026, "QQ", "Q4"],
      [serial(2026, 2, 1), "Q", "Q1"],
      [serial(2026, 6, 30), "Q", "Q2"],
      [OCT_8_2026, "WW", "41"],
      // ISO 8601: 2027-01-01 is in 2026's week 53, 2025-12-29 in 2026's week 1.
      [serial(2027, 1, 1), "WW", "53"],
      [serial(2025, 12, 29), "WW", "1"],
    ])(TITLE, check);
  });
});

// "Text, booleans and errors"
describe("text", () => {
  it.each<Row>([
    ["Bob", "@", "Bob"],
    ["Bob", '"Name: "@', "Name: Bob"],
    ["Bob", '0;-0;0;"Name: "@', "Name: Bob"],
    // No text section → the string as itself.
    ["Bob", "0.00", "Bob"],
    ["Bob", "yyyy-mm-dd", "Bob"],
    ["", "0.00", ""],
    ["x", "[Blue]@", { text: "x", color: BLUE }],
    // A number under a code whose only section is a text section → General.
    [5, "@", "5"],
    [5, '"Name: "@', "5"],
    [1234.5, "@", "1234,5", { locale: "de-DE" }],
    // null → "".
    [null, "0.00", ""],
    [null, "@", ""],
    [null, undefined, ""],
  ])(TITLE, check);
});

describe("booleans and error codes: English and centred in every locale", () => {
  const codes = [undefined, "General", "0.00", "#,##0", "0%", "yyyy-mm-dd", "@", '"x"', ";;;"];

  it.each<SpreadsheetLocaleTag>(["de-DE", "fr-FR", ...SPREADSHEET_LOCALE_TAGS])("%s", (tag) => {
    const where = context({ locale: tag });
    for (const code of codes) {
      expect(formatValue(true, code, where), String(code)).toEqual({ text: "TRUE", alignHint: "center" });
      expect(formatValue(false, code, where), String(code)).toEqual({ text: "FALSE", alignHint: "center" });
      expect(formatValue({ error: "#DIV/0!" }, code, where), String(code)).toEqual({ text: "#DIV/0!", alignHint: "center" });
      expect(formatValue({ error: "#N/A" }, code, where), String(code)).toEqual({ text: "#N/A", alignHint: "center" });
    }
    const red = formatValue(true, "[Red]0", where);
    expect([red.text, red.alignHint]).toEqual(["TRUE", "center"]);
  });
});

describe("formula text is never localised", () => {
  it.each<SpreadsheetLocaleTag>(["de-DE", "fr-FR"])("%s", (tag) => {
    const where = context({ locale: tag });
    for (const text of ["=SUM(A1,B1)", "=1.5", "=ROUND(2.675,2)", '=IF(A1>0,"a","b")']) {
      for (const code of [undefined, "General", "@", "0.00"]) {
        expect(formatValue(text, code, where).text, `${text} under ${String(code)}`).toBe(text);
      }
      expect(editText({ value: text }, where)).toBe(text);
    }
    // editText shows a formula cell's value, never its formula.
    const formulaCell = { value: 1.5, formula: "A1*1.5" };
    expect(editText(formulaCell, where)).toBe("1,5");
  });
});

// "General" — the non-General fit
describe("widths: the # overflow of a non-General number or date", () => {
  it.each<Row>([
    [12345, "0", hash(3), { widthPx: 26 }],
    [123456.789, "0.00", hash(3), { widthPx: 30 }],
    [OCT_8_2026, "yyyy-mm-dd", hash(3), { widthPx: 30 }],
    [OCT_8_2026, "yyyy-mm-dd", hash(8), { widthPx: 64 }],
    [123456789, "0", hash(6), { widthPx: 64, digitWidthPx: 9 }],
    // max(1, …).
    [5, "0", hash(1), { widthPx: 10 }],
    // Fits when its estimate equals the available width: 4 digits × 7 = 33 − 5.
    [1234, "0", "1234", { widthPx: 33 }],
    // "1234.5" is 5.5 digit widths (the point is half): 44 px at MDW 8.
    [1234.5, "0.0", "1234.5", { widthPx: 49, digitWidthPx: 8 }],
    [1234.5, "0.0", hash(5), { widthPx: 48, digitWidthPx: 8 }],
    // General inside a section renders at full width in place, then the
    // section's text is fitted like any non-General number: no shrinking.
    [123456789012, "[Red]General", { ...hash(8), color: RED }, { widthPx: 64 }],
    [12345, "[Red]General", { text: "12345", color: RED }, { widthPx: 64 }],
    // A fill counts one fill character toward the fit.
    [1234, "0*-", hash(4), { widthPx: 33 }],
    [123, "0*-", { text: "123", fill: { char: "-", index: 3 } }, { widthPx: 33 }],
    // Text never shrinks or hashes.
    ["a long text that does not fit", "@", "a long text that does not fit", { widthPx: 26 }],
    ["a long text that does not fit", undefined, "a long text that does not fit", { widthPx: 26 }],
    ["a long text that does not fit", "0.00", "a long text that does not fit", { widthPx: 26 }],
    ["Bob", '"Name: "@', "Name: Bob", { widthPx: 10 }],
    // With no widthPx nothing is fitted.
    [123456789.123, "0.00", "123456789.12"],
    [OCT_8_2026, "dddd, mmmm d, yyyy hh:mm:ss", "Thursday, October 8, 2026 00:00:00"],
  ])(TITLE, check);

  it("hashes to max(1, floor((width − 5) / MDW)) characters", () => {
    for (const widthPx of [10, 20, 26, 30, 33, 64, 100]) {
      for (const digitWidthPx of [5, 7, 9]) {
        const fv = formatValue(1e20, "0", context({ widthPx, digitWidthPx }));
        expect(fv.overflow).toBe("hash");
        expect(fv.text, `${widthPx} px, MDW ${digitWidthPx}`).toBe("#".repeat(Math.max(1, Math.floor((widthPx - 5) / digitWidthPx))));
      }
    }
  });

  it("defaults the MDW to 7", () => {
    expect(formatValue(1e20, "0", context({ widthPx: 64 })).text).toBe("########");
    expect(formatValue(1234, "0", context({ widthPx: 33 })).text).toBe("1234");
  });
});

describe("formatValue never throws: an invalid or unsupported code renders as General", () => {
  it.each([
    '0"abc',
    "[Red",
    "0;0;0;0;0",
    "[Foo]0",
    "0@",
    "0.00E+",
    // Unsupported calendars and numeral systems.
    "[DBNum1]General",
    "[NatNum1]0",
    "B1yyyy-mm-dd",
    "ggge",
    "[$-1010409]d/m/yyyy",
  ])("%j", (code) => {
    for (const tag of ["en-US", "de-DE"] as const) {
      const where = context({ locale: tag });
      for (const value of [1234.5, -0.125, 46303, 0]) {
        expect(() => formatValue(value, code, where)).not.toThrow();
        expect(formatValue(value, code, where)).toEqual(formatGeneral(value, where));
      }
      expect(formatValue("text", code, where)).toEqual({ text: "text", alignHint: "left" });
    }
    expect(formatValue(1234.5, code, context({ locale: "de-DE" })).text).toBe("1234,5");
  });
});

describe("alignHint", () => {
  it.each<[value: CellValue, code: string | undefined, hint: FormattedValue["alignHint"]]>([
    [5, "0.00", "right"],
    [5, undefined, "right"],
    [5, "# ?/?", "right"],
    [5, "0.00E+00", "right"],
    [5, '"pos";"neg"', "right"],
    [OCT_8_2026, "yyyy-mm-dd", "right"],
    [0.5, "h:mm", "right"],
    ["Bob", "@", "left"],
    ["Bob", "0.00", "left"],
    ["Bob", undefined, "left"],
    [true, "0.00", "center"],
    [{ error: "#REF!" }, undefined, "center"],
  ])("%j through %j is %s-aligned", (value, code, hint) => {
    expect(formatValue(value, code, context()).alignHint).toBe(hint);
  });
});

describe("plainText", () => {
  it("is the text with one fill character inserted at the fill's index", () => {
    expect(plainText({ text: "5", fill: { char: "x", index: 1 } })).toBe("5x");
    expect(plainText({ text: "$5", fill: { char: "-", index: 1 } })).toBe("$-5");
    expect(plainText(formatValue(5, '"$"*-0', context()))).toBe("$-5");
    expect(plainText(formatValue(5, "*-0", context()))).toBe("-5");
    expect(plainText(formatValue(1234.5, '_("$"* #,##0.00_)', context()))).toBe(" $ 1,234.50 ");
  });

  it("is the text itself without a fill", () => {
    expect(plainText({ text: "1,234.50" })).toBe("1,234.50");
    expect(plainText(formatValue(1234.5, "#,##0.00", context()))).toBe("1,234.50");
  });
});

// "editText"
describe("editText", () => {
  type EditRow = [label: string, value: CellValue, style: CellStyle | undefined, expected: string, where?: Where];

  it.each<EditRow>([
    // Dates: the locale's short date, + the long time when the time part
    // rounds to a non-zero second.
    ["id 14, en-US", OCT_8_2026_1345, { builtinFormatId: 14 }, "10/8/2026 1:45:00 PM"],
    ["id 14, de-DE", OCT_8_2026_1345, { builtinFormatId: 14 }, "08.10.2026 13:45:00", { locale: "de-DE" }],
    ["id 14, en-GB", OCT_8_2026_1345, { builtinFormatId: 14 }, "08/10/2026 13:45:00", { locale: "en-GB" }],
    ["id 14, fr-FR", OCT_8_2026_1345, { builtinFormatId: 14 }, "08/10/2026 13:45:00", { locale: "fr-FR" }],
    ["id 14, es-ES", OCT_8_2026_1345, { builtinFormatId: 14 }, "08/10/2026 13:45:00", { locale: "es-ES" }],
    ["id 14, cs-CZ", OCT_8_2026_1345, { builtinFormatId: 14 }, "08.10.2026 13:45:00", { locale: "cs-CZ" }],
    ["id 22", OCT_8_2026_1345, { builtinFormatId: 22 }, "10/8/2026 1:45:00 PM"],
    ["the [$-F800] long date", OCT_8_2026_1345, { numberFormat: "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" }, "10/8/2026 1:45:00 PM"],
    ["a date without a time part", OCT_8_2026, { builtinFormatId: 14 }, "10/8/2026"],
    ["a date without a time part, de-DE", OCT_8_2026, { builtinFormatId: 14 }, "08.10.2026", { locale: "de-DE" }],
    ["a time part rounding to 0 s", 46303.000001, { numberFormat: "yyyy-mm-dd" }, "10/8/2026"],
    ["a time part rounding to 1 s", 46303.00001, { numberFormat: "yyyy-mm-dd" }, "10/8/2026 12:00:01 AM"],
    // Time-only codes (elapsed included): the long time below 1, else the
    // short date and the long time.
    ["a time below 1", 0.5729166667, { numberFormat: "h:mm" }, "1:45:00 PM"],
    ["a time below 1, de-DE", 0.5729166667, { numberFormat: "h:mm" }, "13:45:00", { locale: "de-DE" }],
    ["an elapsed time below 1", 0.25, { numberFormat: "[h]:mm:ss" }, "6:00:00 AM"],
    ["[h]:mm:ss on 1.5", 1.5, { numberFormat: "[h]:mm:ss" }, "1/1/1900 12:00:00 PM"],
    ["id 46 on 1.5", 1.5, { numberFormat: "[h]:mm:ss" }, "1/1/1900 12:00:00 PM"],
    ["a time of 1 or more", 2.25, { numberFormat: "h:mm" }, "1/2/1900 6:00:00 AM"],
    // Hash cases → plain notation.
    ["a negative xlsx date", -1, { numberFormat: "yyyy-mm-dd" }, "-1"],
    ["a negative xlsx date, de-DE", -1.5, { numberFormat: "yyyy-mm-dd" }, "-1,5", { locale: "de-DE" }],
    ["a date past 9999", 2958466, { numberFormat: "yyyy-mm-dd" }, "2958466"],
    ["a negative elapsed time", -0.25, { numberFormat: "[h]:mm:ss" }, "-0.25"],
    // An ods-born negative serial is a real date.
    ["a proleptic negative date", -6, { numberFormat: "yyyy-mm-dd" }, "12/25/1899", { dateQuirks: "proleptic" }],
    // Percentages: ×100 in plain notation + "%".
    ["0.125 under 0.00%", 0.125, { numberFormat: "0.00%" }, "12.5%"],
    ["0.125 under id 10", 0.125, { numberFormat: "0.00%" }, "12.5%"],
    ["0.125 under 0%, de-DE", 0.125, { numberFormat: "0%" }, "12,5%", { locale: "de-DE" }],
    ["0.07 under 0%", 0.07, { numberFormat: "0%" }, "7%"],
    ["0.1 + 0.2 under 0%", 0.1 + 0.2, { numberFormat: "0%" }, "30%"],
    ["a quoted % is not a percentage", 0.125, { numberFormat: '0"%"' }, "0.125"],
    // Any other number: plain notation.
    ["no grouping, every digit", 1234567.891, { numberFormat: "#,##0.00" }, "1234567.891"],
    ["no grouping, de-DE", 1234567.891, { numberFormat: "#,##0.00" }, "1234567,891", { locale: "de-DE" }],
    ["the accounting negative", -1234.5, { numberFormat: "#,##0.00_);[Red]\\(#,##0.00\\)" }, "-1234.5"],
    ["General, no style", OCT_8_2026_1345, undefined, "46303.5729166667"],
    ["15 significant digits", 1 / 3, undefined, "0.333333333333333"],
    ["no General shrinking", 123456789012, {}, "123456789012"],
    ["just below 1e15", 999999999999999, undefined, "999999999999999"],
    ["1e15", 1e15, undefined, "1E+15"],
    ["above 1e15", 123_456_789_012_345_680, undefined, "1.23456789012346E+17"],
    ["1e-9", 1e-9, undefined, "0.000000001"],
    ["below 1e-9", 1e-10, undefined, "1E-10"],
    ["0", 0, { numberFormat: "0.00" }, "0"],
    ["fr-FR decimal", 1234.5, { numberFormat: "#,##0.00" }, "1234,5", { locale: "fr-FR" }],
    // Booleans, errors, null, text.
    ["TRUE", true, { numberFormat: "0.00" }, "TRUE"],
    ["FALSE", false, undefined, "FALSE"],
    ["TRUE, de-DE", true, undefined, "TRUE", { locale: "de-DE" }],
    ["an error", { error: "#N/A" }, undefined, "#N/A"],
    ["an error, fr-FR", { error: "#DIV/0!" }, { numberFormat: "0.00" }, "#DIV/0!", { locale: "fr-FR" }],
    ["no value", null, { numberFormat: "0.00" }, ""],
    ["text", "007", undefined, "007"],
    ["text under a plain style", "007", {}, "007"],
    ["quote-prefixed text", "007", { quotePrefix: true }, "'007"],
    ["quote-prefixed text under @", "1/2", { quotePrefix: true, numberFormat: "@" }, "'1/2"],
  ])("%s → %j", (_label, value, style, expected, where) => {
    expect(editText({ value }, { ...context(where), style })).toBe(expected);
  });

  it("shows the formula bar's e2e values for the fixture's A4", () => {
    const style: CellStyle = { numberFormat: "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" };
    expect(editText({ value: OCT_8_2026_1345 }, { ...context(), style })).toBe("10/8/2026 1:45:00 PM");
    expect(editText({ value: OCT_8_2026_1345 }, { ...context({ locale: "de-DE" }), style })).toBe("08.10.2026 13:45:00");
  });
});
