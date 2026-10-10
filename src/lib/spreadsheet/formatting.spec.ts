// SPEC-123 cell formatting — the unit tests of its formatting layer. The
// blocks here cover numberCodes.ts: the number-format presets in all six
// locales with the toolbar's currency/percent/comma buttons and the
// number-format shortcuts, every "Decimal stepping" example plus the
// accounting and General cases, and the Format Cells code builders' round
// trip over every category's whole option set in every locale.
import { describe, expect, it } from "vitest";
import type { CellValue } from "./model";
import {
  buildNumberCode,
  commaButtonFormat,
  currencyButtonFormat,
  CURRENCY_SYMBOLS,
  currencySymbolPosition,
  dateTypeCodes,
  describeNumberCode,
  FRACTION_TYPE_CODES,
  isValidFormatCode,
  LOCALE_CURRENCY,
  NUMBER_CATEGORIES,
  NUMBER_FORMAT_PRESETS,
  numberFormatChoiceOf,
  percentButtonFormat,
  presetIdOf,
  presetNumberFormat,
  shortcutNumberFormat,
  stepDecimals,
  TIME_TYPE_CODES,
  type NegativeStyle,
  type NumberCodeSpec,
  type NumberFormatChoice,
  type NumberPresetId,
  type NumberShortcutId,
} from "./numberCodes";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "./numberFormat/builtin";
import { SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "./numberFormat/locale";
import { FormatCodeError, formatCodeSections } from "./numberFormat/parse";
import { formatValue, type FormatContext } from "./numberFormat/render";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const DE_DE = SPREADSHEET_LOCALES["de-DE"];

const ID_41 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
const ID_42 = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)';
const ID_43 = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const ID_44 = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
const EURO_ACCOUNTING = '_-* #,##0.00 "€"_-;-* #,##0.00 "€"_-;_-* "-"?? "€"_-;_-@_-';
const NO_SYMBOL_ACCOUNTING = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-';

const NEGATIVES: readonly NegativeStyle[] = ["minus", "red", "parens", "redParens"];
const DECIMALS = Array.from({ length: 31 }, (_, decimals) => decimals);

type PresetTable = Record<NumberPresetId, NumberFormatChoice>;

/** The rows every locale shares; `rest` adds the locale's own. */
function presets(rest: Pick<PresetTable, "currency" | "accounting" | "long-date" | "time">): PresetTable {
  return {
    general: { numberFormat: null },
    number: { numberFormat: "0.00" },
    "short-date": { numberFormat: "m/d/yyyy", builtinFormatId: 14 },
    percentage: { numberFormat: "0.00%" },
    fraction: { numberFormat: "# ?/?" },
    scientific: { numberFormat: "0.00E+00" },
    text: { numberFormat: "@" },
    ...rest,
  };
}

/** The Reference's preset table. */
const PRESETS: Record<SpreadsheetLocaleTag, PresetTable> = {
  "en-US": presets({
    currency: { numberFormat: '"$"#,##0.00' },
    accounting: { numberFormat: ID_44, builtinFormatId: 44 },
    "long-date": { numberFormat: "dddd, mmmm d, yyyy" },
    time: { numberFormat: "h:mm:ss AM/PM" },
  }),
  "en-GB": presets({
    currency: { numberFormat: '"£"#,##0.00' },
    accounting: { numberFormat: '_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-' },
    "long-date": { numberFormat: "dddd, d mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
  "de-DE": presets({
    currency: { numberFormat: '#,##0.00 "€"' },
    accounting: { numberFormat: EURO_ACCOUNTING },
    "long-date": { numberFormat: "dddd, d. mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
  "fr-FR": presets({
    currency: { numberFormat: '#,##0.00 "€"' },
    accounting: { numberFormat: EURO_ACCOUNTING },
    "long-date": { numberFormat: "dddd d mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
  "es-ES": presets({
    currency: { numberFormat: '#,##0.00 "€"' },
    accounting: { numberFormat: EURO_ACCOUNTING },
    "long-date": { numberFormat: 'dddd, d "de" mmmm "de" yyyy' },
    time: { numberFormat: "h:mm:ss" },
  }),
  "cs-CZ": presets({
    currency: { numberFormat: '#,##0.00 "Kč"' },
    accounting: { numberFormat: '_-* #,##0.00 "Kč"_-;-* #,##0.00 "Kč"_-;_-* "-"?? "Kč"_-;_-@_-' },
    "long-date": { numberFormat: "dddd d. mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
};

const PRESET_IDS = NUMBER_FORMAT_PRESETS.map((preset) => preset.id);

function context(tag: SpreadsheetLocaleTag): FormatContext {
  return { locale: SPREADSHEET_LOCALES[tag], dateSystem: "1900" };
}

describe("SPEC-123 number-format presets per locale", () => {
  it("lists the dropdown's eleven entries in order", () => {
    expect(NUMBER_FORMAT_PRESETS).toEqual([
      { id: "general", label: "General" },
      { id: "number", label: "Number" },
      { id: "currency", label: "Currency" },
      { id: "accounting", label: "Accounting" },
      { id: "short-date", label: "Short date" },
      { id: "long-date", label: "Long date" },
      { id: "time", label: "Time" },
      { id: "percentage", label: "Percentage" },
      { id: "fraction", label: "Fraction" },
      { id: "scientific", label: "Scientific" },
      { id: "text", label: "Text" },
    ]);
  });

  it("pins each locale's currency: symbol, side and negatives", () => {
    expect(LOCALE_CURRENCY).toEqual({
      "en-US": { symbol: "$", position: "before", negativeParens: true },
      "en-GB": { symbol: "£", position: "before", negativeParens: false },
      "de-DE": { symbol: "€", position: "after", negativeParens: false },
      "fr-FR": { symbol: "€", position: "after", negativeParens: false },
      "es-ES": { symbol: "€", position: "after", negativeParens: false },
      "cs-CZ": { symbol: "Kč", position: "after", negativeParens: false },
    });
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: every preset's code and built-in id", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    for (const id of PRESET_IDS) {
      expect(presetNumberFormat(id, locale), id).toStrictEqual(PRESETS[tag][id]);
      const code = presetNumberFormat(id, locale).numberFormat;
      if (code !== null) {
        expect(isValidFormatCode(code), code).toBe(true);
      }
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: the currency, percent and comma buttons", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    expect(currencyButtonFormat(locale)).toStrictEqual(PRESETS[tag].accounting);
    expect(percentButtonFormat()).toStrictEqual({ numberFormat: "0%" });
    expect(commaButtonFormat(locale)).toStrictEqual(
      tag === "en-US" ? { numberFormat: ID_43, builtinFormatId: 43 } : { numberFormat: NO_SYMBOL_ACCOUNTING },
    );
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: the number-format shortcuts", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const expected: Record<NumberShortcutId, NumberFormatChoice> = {
      general: { numberFormat: null },
      number: { numberFormat: "#,##0.00" },
      time: { numberFormat: "h:mm AM/PM" },
      date: { numberFormat: "d-mmm-yy" },
      currency: tag === "en-US" ? { numberFormat: '"$"#,##0.00_);\\("$"#,##0.00\\)', builtinFormatId: 7 } : PRESETS[tag].currency,
      percent: { numberFormat: "0%" },
      scientific: { numberFormat: "0.00E+00" },
    };
    for (const [id, choice] of Object.entries(expected)) {
      expect(shortcutNumberFormat(id as NumberShortcutId, locale), id).toStrictEqual(choice);
    }
  });

  it("numberFormatChoiceOf: General is null; a locale-dependent built-in's code carries its id", () => {
    for (const general of [undefined, null, "", "  ", "General", "general", "GENERAL"]) {
      expect(numberFormatChoiceOf(general), String(general)).toStrictEqual({ numberFormat: null });
    }
    for (const id of LOCALE_DEPENDENT_BUILTINS) {
      const code = BUILTIN_NUMBER_FORMATS[id]!;
      expect(numberFormatChoiceOf(code), code).toStrictEqual({ numberFormat: code, builtinFormatId: id });
    }
    for (const id of [1, 2, 3, 4, 9, 10, 11, 12, 15, 18, 19, 21, 37, 45, 48, 49]) {
      const code = BUILTIN_NUMBER_FORMATS[id]!;
      expect(numberFormatChoiceOf(code), code).toStrictEqual({ numberFormat: code });
    }
    expect(numberFormatChoiceOf('"€"#,##0.0')).toStrictEqual({ numberFormat: '"€"#,##0.0' });
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: presetIdOf names the preset a style's format is exactly", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    for (const id of PRESET_IDS) {
      const { numberFormat, builtinFormatId } = presetNumberFormat(id, locale);
      const style = numberFormat === null ? {} : builtinFormatId === undefined ? { numberFormat } : { numberFormat, builtinFormatId };
      expect(presetIdOf(style, locale), id).toBe(id);
    }
    expect(presetIdOf({ numberFormat: "General" }, locale)).toBe("general");
    expect(presetIdOf({ numberFormat: "0.000" }, locale)).toBeNull();
    // Short date is the locale-dependent id 14, never the bare code.
    expect(presetIdOf({ numberFormat: "m/d/yyyy" }, locale)).toBeNull();
  });

  it("presetIdOf: Excel's id 44 is Accounting only where it is the locale's Accounting", () => {
    const id44 = { numberFormat: ID_44, builtinFormatId: 44 };
    expect(presetIdOf(id44, EN_US)).toBe("accounting");
    expect(presetIdOf(id44, SPREADSHEET_LOCALES["en-GB"])).toBeNull();
    expect(presetIdOf(id44, DE_DE)).toBeNull();
  });

  it("isValidFormatCode is parseFormatCode not throwing", () => {
    for (const code of ["0.00", '"$"#,##0.00_);("$"#,##0.00)', "General", "@", "yyyy-mm-dd", "[Red]0.0", ID_44]) {
      expect(isValidFormatCode(code), code).toBe(true);
    }
    for (const code of ['"abc', "0;0;0;0;0", "[Purple]0", "0@", "0E+", "[Red"]) {
      expect(isValidFormatCode(code), code).toBe(false);
    }
  });

  it("renders through SPEC-118: en-US Currency of 1234.5 is $1,234.50, 0.0% of 0.125 is 12.5%", () => {
    expect(formatValue(1234.5, presetNumberFormat("currency", EN_US).numberFormat!, context("en-US")).text).toBe("$1,234.50");
    expect(formatValue(0.125, "0.0%", context("en-US")).text).toBe("12.5%");
    expect(formatValue(1234.5, presetNumberFormat("currency", SPREADSHEET_LOCALES["en-GB"]).numberFormat!, context("en-GB")).text).toBe("£1,234.50");
    expect(formatValue(1234.5, presetNumberFormat("currency", DE_DE).numberFormat!, context("de-DE")).text).toBe("1.234,50 €");
    expect(formatValue(-1234.5, presetNumberFormat("currency", SPREADSHEET_LOCALES["cs-CZ"]).numberFormat!, context("cs-CZ")).text).toBe(
      "-1 234,50 Kč",
    );
  });
});

describe("SPEC-123 stepDecimals (every Decimal stepping example)", () => {
  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ["0", 1, "0.0"],
    ["0.0", -1, "0"],
    ["#,##0.00", 1, "#,##0.000"],
    ["0%", 1, "0.0%"],
    ["0.00E+00", 1, "0.000E+00"],
    ['"$"#,##0.00_);("$"#,##0.00)', 1, '"$"#,##0.000_);("$"#,##0.000)'],
    ["0", -1, "0"],
    ["yyyy-mm-dd", 1, "yyyy-mm-dd"],
    ["yyyy-mm-dd", -1, "yyyy-mm-dd"],
  ])("the Reference's example: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1234.5, dir)).toBe(stepped);
  });

  it.each<[value: CellValue, up: string, down: string]>([
    [1234.5, "0.00", "0"],
    [0.125, "0.0000", "0.00"],
    [-2.75, "0.000", "0.0"],
    [5, "0.0", "0"],
    [0, "0.0", "0"],
    [123456789012, "0.000000", "0.0000"], // "1.23457E+11": the mantissa's 5
    [0.0000000001234, "0.0000", "0.00"], // "1.234E-10"
    ["12.5", "0.0", "0"],
    [true, "0.0", "0"],
    [null, "0.0", "0"],
  ])("General on %j: +1 is %s, −1 is %s (the decimals generalText shows)", (value, up, down) => {
    for (const general of [undefined, "", "General", "general", "GENERAL"]) {
      expect(stepDecimals(general, value, 1)).toBe(up);
      expect(stepDecimals(general, value, -1)).toBe(down);
    }
  });

  it("steps id 44 down to id 42 through the one-decimal code, and back up", () => {
    const oneDecimal = '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)';
    expect(stepDecimals(ID_44, null, -1)).toBe(oneDecimal);
    expect(stepDecimals(oneDecimal, null, -1)).toBe(ID_42);
    expect(stepDecimals(ID_42, null, 1)).toBe(oneDecimal);
    expect(stepDecimals(oneDecimal, null, 1)).toBe(ID_44);
    expect(stepDecimals(ID_42, null, -1)).toBe(ID_42);
    expect(stepDecimals(ID_43, null, -1)).toBe('_(* #,##0.0_);_(* \\(#,##0.0\\);_(* "-"?_);_(@_)');
    expect(stepDecimals(ID_41, null, 1)).toBe('_(* #,##0.0_);_(* \\(#,##0.0\\);_(* "-"?_);_(@_)');
  });

  it("steps the symbol-after accounting template, its dash and symbol untouched", () => {
    const one = '_-* #,##0.0 "€"_-;-* #,##0.0 "€"_-;_-* "-"? "€"_-;_-@_-';
    const none = '_-* #,##0 "€"_-;-* #,##0 "€"_-;_-* "-" "€"_-;_-@_-';
    expect(stepDecimals(EURO_ACCOUNTING, 1, -1)).toBe(one);
    expect(stepDecimals(one, 1, -1)).toBe(none);
    expect(stepDecimals(none, 1, 1)).toBe(one);
    expect(stepDecimals(EURO_ACCOUNTING, 1, 1)).toBe('_-* #,##0.000 "€"_-;-* #,##0.000 "€"_-;_-* "-"??? "€"_-;_-@_-');
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: steps every code the builders make to the one with a decimal more or less", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const mismatches: string[] = [];
    for (const spec of everyOptionSet(tag)) {
      if (spec.decimals === undefined || spec.decimals === 30) {
        continue;
      }
      const fewer = buildNumberCode(spec, locale)!;
      const more = buildNumberCode({ ...spec, decimals: spec.decimals + 1 }, locale)!;
      if (stepDecimals(fewer, 0, 1) !== more) {
        mismatches.push(`${fewer} +1 → ${stepDecimals(fewer, 0, 1)}, not ${more}`);
      }
      if (stepDecimals(more, 0, -1) !== fewer) {
        mismatches.push(`${more} −1 → ${stepDecimals(more, 0, -1)}, not ${fewer}`);
      }
    }
    expect(mismatches).toEqual([]);
  });

  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ['"0.0" 0', 1, '"0.0" 0.0'],
    ['0.00 "0.00"', -1, '0.0 "0.00"'],
    ["\\0#_0*00", 1, "\\0#_0*00.0"],
    ['"$"#,##0.00_);\\("$"#,##0.00\\)', -1, '"$"#,##0.0_);\\("$"#,##0.0\\)'],
    ["[>=100]0.0;[Red]0", 1, "[>=100]0.00;[Red]0.0"],
    ["[Color10]0.00", -1, "[Color10]0.0"],
    ["[$€-407] #,##0.00", -1, "[$€-407] #,##0.0"],
    ["#,##0,", 1, "#,##0.0,"],
    ['0.00;-0.00;"zero";@', 1, '0.000;-0.000;"zero";@'],
    ["[<1]0.00;yyyy-mm-dd", 1, "[<1]0.000;yyyy-mm-dd"],
  ])("edits only placeholders and points: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1, dir)).toBe(stepped);
  });

  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ["##0.0E+0", -1, "##0E+0"],
    ["##0.0E+0", 1, "##0.00E+0"],
    ["0E+00", 1, "0.0E+00"],
    ["0.0E+00", -1, "0E+00"],
  ])("counts a scientific code's mantissa only: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1, dir)).toBe(stepped);
  });

  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ["0.", 1, "0.0"],
    ["0.", -1, "0."], // no decimal placeholder to remove
    [".00", -1, ".0"],
    [".0", -1, "0"], // the last decimal stays as the integer placeholder
    ["???", 1, "???.0"], // `?` pads only an accounting zero section's dash
    ['"-"??', 1, '"-"??.0'], // … in the third section
    ['0.00;"-"', 1, '0.000;"-"'],
    ['0.00;-0.00;"-"', 1, '0.000;-0.000;"-"?'],
    ["0.0#", 1, "0.0#0"],
  ])("edge: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1, dir)).toBe(stepped);
  });

  it.each(["yyyy-mm-dd", "h:mm:ss", "[h]:mm:ss", "mmss.0", "# ?/?", "# ??/100", "@", '"text"', "General;-General"])(
    "leaves %s as it is (no number section)",
    (code) => {
      expect(stepDecimals(code, 1, 1)).toBe(code);
      expect(stepDecimals(code, 1, -1)).toBe(code);
    },
  );

  it.each(['"abc', "0;0;0;0;0", "0@", "[Purple]0.0", "0.0E+"])("returns the invalid code %s unchanged", (code) => {
    expect(stepDecimals(code, 1, 1)).toBe(code);
    expect(stepDecimals(code, 1, -1)).toBe(code);
  });

  it("edits at formatCodeSections' offsets: each section's kind and its tokens where they are written", () => {
    const sections = formatCodeSections('"$"0.0;[Red]\\(0\\);yyyy;@');
    expect(sections.map((section) => section.kind)).toEqual(["number", "number", "date", "text"]);
    expect(sections[0]!.tokens.map(({ token, offset }) => [token.type, offset])).toEqual([
      ["literal", 0],
      ["digit", 3],
      ["point", 4],
      ["digit", 5],
    ]);
    expect(sections[1]!.tokens.map(({ offset }) => offset)).toEqual([12, 14, 15]); // [Red] is no token
    expect(sections[2]!.tokens).toEqual([{ token: { type: "date", part: "year", length: 4 }, offset: 18 }]);
    expect(formatCodeSections("0.0E")[0]!.tokens.map(({ token }) => token.type)).toEqual(["digit", "point", "digit"]); // the era `E` left out
    expect(() => formatCodeSections('"abc')).toThrow(FormatCodeError);
  });
});

/** Every category's whole option set in `tag`, written out here rather than
 * taken from numberCodes.ts — Currency with its None included. */
function everyOptionSet(tag: SpreadsheetLocaleTag): NumberCodeSpec[] {
  const locale = SPREADSHEET_LOCALES[tag];
  const symbols = ["", ...CURRENCY_SYMBOLS];
  return [
    { category: "general" },
    ...DECIMALS.flatMap((decimals) =>
      [false, true].flatMap((thousands) => NEGATIVES.map((negative): NumberCodeSpec => ({ category: "number", decimals, thousands, negative }))),
    ),
    ...DECIMALS.flatMap((decimals) =>
      NEGATIVES.flatMap((negative) => symbols.map((symbol): NumberCodeSpec => ({ category: "currency", decimals, negative, symbol }))),
    ),
    ...DECIMALS.flatMap((decimals) => symbols.map((symbol): NumberCodeSpec => ({ category: "accounting", decimals, symbol }))),
    ...dateTypeCodes(locale).map((type): NumberCodeSpec => ({ category: "date", type })),
    ...TIME_TYPE_CODES.map((type): NumberCodeSpec => ({ category: "time", type })),
    ...DECIMALS.map((decimals): NumberCodeSpec => ({ category: "percentage", decimals })),
    ...FRACTION_TYPE_CODES.map((type): NumberCodeSpec => ({ category: "fraction", type })),
    ...DECIMALS.map((decimals): NumberCodeSpec => ({ category: "scientific", decimals })),
    { category: "text" },
  ];
}

/** en-US's Number code for these options. */
function numberCode(decimals: number, thousands: boolean, negative: NegativeStyle): string | undefined {
  return buildNumberCode({ category: "number", decimals, thousands, negative }, EN_US);
}

/** A two-decimal Currency code for these options. */
function currencyCode(negative: NegativeStyle, symbol: string, locale = EN_US): string | undefined {
  return buildNumberCode({ category: "currency", decimals: 2, negative, symbol }, locale);
}

/** An Accounting code for these options. */
function accountingCode(decimals: number, symbol: string, tag: SpreadsheetLocaleTag = "en-US"): string | undefined {
  return buildNumberCode({ category: "accounting", decimals, symbol }, SPREADSHEET_LOCALES[tag]);
}

describe("SPEC-123 buildNumberCode / describeNumberCode round trip", () => {
  it("lists the categories, symbols and types of the Number tab", () => {
    expect(NUMBER_CATEGORIES.map((category) => category.label)).toEqual([
      "General",
      "Number",
      "Currency",
      "Accounting",
      "Date",
      "Time",
      "Percentage",
      "Fraction",
      "Scientific",
      "Text",
      "Custom",
    ]);
    expect(NUMBER_CATEGORIES.map((category) => category.id)).toEqual([
      "general",
      "number",
      "currency",
      "accounting",
      "date",
      "time",
      "percentage",
      "fraction",
      "scientific",
      "text",
      "custom",
    ]);
    expect(CURRENCY_SYMBOLS).toEqual(["$", "€", "£", "¥", "Kč", "CHF"]);
    expect(TIME_TYPE_CODES).toEqual(["h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "mm:ss", "[h]:mm:ss", "mmss.0"]);
    expect(FRACTION_TYPE_CODES).toEqual(["# ?/?", "# ??/??", "# ???/???", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10", "# ??/100"]);
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(dateTypeCodes(SPREADSHEET_LOCALES[tag]), tag).toEqual([
        "m/d/yyyy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "m/d/yyyy h:mm",
        PRESETS[tag]["long-date"].numberFormat,
        "yyyy-mm-dd",
        "mmmm yyyy",
      ]);
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: each symbol's side — the locale's own on its side, $ £ ¥ before, € Kč CHF after", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const sides = CURRENCY_SYMBOLS.map((symbol) => currencySymbolPosition(symbol, locale));
    expect(sides).toEqual(["before", "after", "before", "before", "after", "after"]);
    expect(currencySymbolPosition(LOCALE_CURRENCY[tag].symbol, locale)).toBe(LOCALE_CURRENCY[tag].position);
    expect(currencySymbolPosition("¤", locale)).toBe(LOCALE_CURRENCY[tag].position);
  });

  it("builds Number with each negative style, and Currency with the symbol in every section", () => {
    expect(numberCode(2, false, "minus")).toBe("0.00");
    expect(numberCode(0, false, "minus")).toBe("0");
    expect(numberCode(0, true, "minus")).toBe("#,##0");
    expect(numberCode(2, true, "red")).toBe("#,##0.00;[Red]#,##0.00");
    expect(numberCode(1, false, "parens")).toBe("0.0_);(0.0)");
    expect(numberCode(3, true, "redParens")).toBe("#,##0.000_);[Red](#,##0.000)");
    expect(currencyCode("parens", "$")).toBe('"$"#,##0.00_);("$"#,##0.00)');
    expect(currencyCode("red", "$")).toBe('"$"#,##0.00;[Red]"$"#,##0.00');
    expect(currencyCode("redParens", "€")).toBe('#,##0.00 "€"_);[Red](#,##0.00 "€")');
    expect(currencyCode("minus", "CHF", DE_DE)).toBe('#,##0.00 "CHF"');
    expect(currencyCode("minus", "¥", DE_DE)).toBe('"¥"#,##0.00');
    expect(currencyCode("minus", "")).toBe("#,##0.00");
    expect(buildNumberCode({ category: "currency" }, DE_DE)).toBe('#,##0.00 "€"'); // the locale's symbol, 2 decimals
  });

  it("builds Accounting from the locale's template: en-US's are Excel's ids 41–44 exactly", () => {
    expect(accountingCode(2, "$")).toBe(ID_44);
    expect(accountingCode(0, "$")).toBe(ID_42);
    expect(accountingCode(2, "")).toBe(ID_43);
    expect(accountingCode(0, "")).toBe(ID_41);
    expect(accountingCode(2, "€")).toBe('_(* #,##0.00 "€"_);_(* \\(#,##0.00 "€"\\);_(* "-"?? "€"_);_(@_)');
    expect(accountingCode(2, "€", "en-GB")).toBe(EURO_ACCOUNTING);
    expect(accountingCode(1, "$", "de-DE")).toBe('_-"$"* #,##0.0_-;-"$"* #,##0.0_-;_-"$"* "-"?_-;_-@_-');
    expect(accountingCode(0, "Kč", "cs-CZ")).toBe('_-* #,##0 "Kč"_-;-* #,##0 "Kč"_-;_-* "-" "Kč"_-;_-@_-');
    expect(accountingCode(2, "", "fr-FR")).toBe(NO_SYMBOL_ACCOUNTING);
  });

  it("builds the other categories, clamps decimals and falls back to a category's first type", () => {
    expect(buildNumberCode({ category: "general" }, EN_US)).toBeUndefined();
    expect(buildNumberCode({ category: "percentage", decimals: 0 }, EN_US)).toBe("0%");
    expect(buildNumberCode({ category: "percentage", decimals: 1 }, EN_US)).toBe("0.0%");
    expect(buildNumberCode({ category: "scientific", decimals: 0 }, EN_US)).toBe("0E+00");
    expect(buildNumberCode({ category: "scientific" }, EN_US)).toBe("0.00E+00");
    expect(buildNumberCode({ category: "text" }, EN_US)).toBe("@");
    expect(buildNumberCode({ category: "number", decimals: 31 }, EN_US)).toBe(`0.${"0".repeat(30)}`);
    expect(buildNumberCode({ category: "number", decimals: -3 }, EN_US)).toBe("0");
    expect(buildNumberCode({ category: "number", decimals: 2.7 }, EN_US)).toBe("0.00");
    expect(buildNumberCode({ category: "fraction", type: "# ?/8" }, EN_US)).toBe("# ?/8");
    expect(buildNumberCode({ category: "fraction", type: "# ?/3" }, EN_US)).toBe("# ?/?");
    expect(buildNumberCode({ category: "date" }, DE_DE)).toBe("m/d/yyyy");
    expect(buildNumberCode({ category: "date", type: "dddd, d. mmmm yyyy" }, DE_DE)).toBe("dddd, d. mmmm yyyy");
    expect(buildNumberCode({ category: "date", type: "dddd, d. mmmm yyyy" }, EN_US)).toBe("m/d/yyyy");
    expect(buildNumberCode({ category: "time", type: "hh:mm" }, EN_US)).toBe("h:mm AM/PM");
    expect(buildNumberCode({ category: "custom", type: '0.0 "kg"' }, EN_US)).toBe('0.0 "kg"');
    for (const general of [undefined, "", "General", "general"]) {
      expect(buildNumberCode({ category: "custom", type: general }, EN_US), String(general)).toBeUndefined();
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: every option set builds a valid code that describes back to it", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const specs = everyOptionSet(tag);
    const codes = specs.map((spec) => buildNumberCode(spec, locale));
    const described = codes.map((code) => describeNumberCode(code, locale));
    expect(described.map((spec) => buildNumberCode(spec, locale))).toEqual(codes);
    // The one alias: Currency with no symbol is Number with thousands.
    const expected = specs.map(
      (spec): NumberCodeSpec =>
        spec.category === "currency" && spec.symbol === ""
          ? { category: "number", decimals: spec.decimals!, thousands: true, negative: spec.negative! }
          : spec,
    );
    expect(described).toStrictEqual(expected);
    expect(codes.filter((code) => code !== undefined && !isValidFormatCode(code))).toEqual([]);
  });

  it("describes General as General and every code no option set builds as Custom", () => {
    for (const general of [undefined, "", "General", "general"]) {
      expect(describeNumberCode(general, EN_US), String(general)).toStrictEqual({ category: "general" });
    }
    const custom = [
      "0.0#",
      "##0.0E+0",
      BUILTIN_NUMBER_FORMATS[7]!, // Excel escapes the parentheses the builder writes bare
      BUILTIN_NUMBER_FORMATS[39]!,
      "[Red]0.00",
      '"€"#,##0.00', // € goes after the number
      "0.00 %",
      "yyyy-mm-dd hh:mm",
      `0.${"0".repeat(31)}`,
      "dddd, d mmmm yyyy", // en-GB's long date, not en-US's
    ];
    for (const code of custom) {
      expect(describeNumberCode(code, EN_US), code).toStrictEqual({ category: "custom", type: code });
      expect(buildNumberCode(describeNumberCode(code, EN_US), EN_US)).toBe(code);
    }
  });

  it("describes Excel's built-ins and the presets by their options", () => {
    expect(describeNumberCode(ID_44, EN_US)).toStrictEqual({ category: "accounting", decimals: 2, symbol: "$" });
    expect(describeNumberCode(ID_41, EN_US)).toStrictEqual({ category: "accounting", decimals: 0, symbol: "" });
    expect(describeNumberCode("#,##0.00", EN_US)).toStrictEqual({ category: "number", decimals: 2, thousands: true, negative: "minus" });
    expect(describeNumberCode("0%", DE_DE)).toStrictEqual({ category: "percentage", decimals: 0 });
    expect(describeNumberCode(EURO_ACCOUNTING, DE_DE)).toStrictEqual({ category: "accounting", decimals: 2, symbol: "€" });
    expect(describeNumberCode("h:mm:ss", DE_DE)).toStrictEqual({ category: "time", type: "h:mm:ss" });
    expect(describeNumberCode("@", DE_DE)).toStrictEqual({ category: "text" });
  });

  it("hands out a fresh spec each time", () => {
    const spec = describeNumberCode("0.00", EN_US);
    spec.decimals = 5;
    expect(describeNumberCode("0.00", EN_US)).toStrictEqual({ category: "number", decimals: 2, thousands: false, negative: "minus" });
  });
});
