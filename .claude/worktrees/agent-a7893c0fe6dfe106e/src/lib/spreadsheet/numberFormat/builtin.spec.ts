// SPEC-118 "Built-in formats": BUILTIN_NUMBER_FORMATS holds ids 0–22 and
// 37–49 in en-US spellings (the CJK ids 23–36 and 50–81 have no entry);
// LOCALE_DEPENDENT_BUILTINS is 5–8, 14, 22, 41–44; builtinFormatCode gives
// each of those in all six locales (14 → shortDate, 22 → shortDateTime, the
// currency ids with the locale's quoted symbol — before the number, or in
// the suffix templates after it) and every other id unchanged;
// formatCodeOf prefers a style's builtinFormatId. The acceptance criterion:
// id 14 renders "10/8/2026" in en-US and "08.10.2026" in de-DE, and ids
// 23–36 and 50–81 render as General.
import { describe, expect, it } from "vitest";
import { BUILTIN_NUMBER_FORMATS, builtinFormatCode, formatCodeOf, LOCALE_DEPENDENT_BUILTINS } from "./builtin";
import { SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "./locale";
import { formatValue, type FormatContext } from "./render";

/** The Reference's table, en-US spellings. */
const TABLE: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const DEPENDENT = [5, 6, 7, 8, 14, 22, 41, 42, 43, 44];

/** The currency ids of a prefix locale: the table with `"$"` → the quoted symbol. */
function prefixCurrency(symbol: string): Record<number, string> {
  const S = `"${symbol}"`;
  return {
    5: `${S}#,##0_);\\(${S}#,##0\\)`,
    6: `${S}#,##0_);[Red]\\(${S}#,##0\\)`,
    7: `${S}#,##0.00_);\\(${S}#,##0.00\\)`,
    8: `${S}#,##0.00_);[Red]\\(${S}#,##0.00\\)`,
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: `_(${S}* #,##0_);_(${S}* \\(#,##0\\);_(${S}* "-"_);_(@_)`,
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: `_(${S}* #,##0.00_);_(${S}* \\(#,##0.00\\);_(${S}* "-"??_);_(@_)`,
  };
}

/** The currency ids of a suffix locale: the Reference's templates, S = the
 * quoted symbol. */
function suffixCurrency(symbol: string): Record<number, string> {
  const S = `"${symbol}"`;
  return {
    5: `#,##0 ${S};\\-#,##0 ${S}`,
    6: `#,##0 ${S};[Red]\\-#,##0 ${S}`,
    7: `#,##0.00 ${S};\\-#,##0.00 ${S}`,
    8: `#,##0.00 ${S};[Red]\\-#,##0.00 ${S}`,
    41: '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-',
    42: `_-* #,##0 ${S}_-;\\-* #,##0 ${S}_-;_-* "-" ${S}_-;_-@_-`,
    43: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
    44: `_-* #,##0.00 ${S}_-;\\-* #,##0.00 ${S}_-;_-* "-"?? ${S}_-;_-@_-`,
  };
}

/** Every locale-dependent id's code in each of the six locales. */
const LOCALIZED: Record<SpreadsheetLocaleTag, Record<number, string>> = {
  "en-US": { ...prefixCurrency("$"), 14: "m/d/yyyy", 22: "m/d/yyyy h:mm" },
  "en-GB": { ...prefixCurrency("£"), 14: "dd/mm/yyyy", 22: "dd/mm/yyyy hh:mm" },
  "de-DE": { ...suffixCurrency("€"), 14: "dd.mm.yyyy", 22: "dd.mm.yyyy hh:mm" },
  "fr-FR": { ...suffixCurrency("€"), 14: "dd/mm/yyyy", 22: "dd/mm/yyyy hh:mm" },
  "es-ES": { ...suffixCurrency("€"), 14: "dd/mm/yyyy", 22: "dd/mm/yyyy h:mm" },
  "cs-CZ": { ...suffixCurrency("Kč"), 14: "dd.mm.yyyy", 22: "dd.mm.yyyy h:mm" },
};

function range(from: number, to: number): number[] {
  return Array.from({ length: to - from + 1 }, (_, i) => from + i);
}

function context(tag: SpreadsheetLocaleTag): FormatContext {
  return { locale: SPREADSHEET_LOCALES[tag], dateSystem: "1900" };
}

/** 2026-10-08 13:45:00. */
const OCT_8_2026_1345 = 46303.5729166667;

describe("BUILTIN_NUMBER_FORMATS", () => {
  it("is the Reference's table of ids 0–22 and 37–49, en-US spellings", () => {
    expect(BUILTIN_NUMBER_FORMATS).toEqual(TABLE);
  });

  it.each(Object.entries(TABLE))("id %s is %j", (id, code) => {
    expect(BUILTIN_NUMBER_FORMATS[Number(id)]).toBe(code);
  });

  it.each([...range(23, 36), ...range(50, 81)])("has no entry for the CJK id %i", (id) => {
    expect(BUILTIN_NUMBER_FORMATS[id]).toBeUndefined();
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(builtinFormatCode(id, SPREADSHEET_LOCALES[tag])).toBeUndefined();
    }
  });
});

describe("LOCALE_DEPENDENT_BUILTINS", () => {
  it("is 5–8, 14, 22 and 41–44", () => {
    expect(LOCALE_DEPENDENT_BUILTINS.toSorted((a, b) => a - b)).toEqual(DEPENDENT);
  });
});

describe("builtinFormatCode", () => {
  it.each(SPREADSHEET_LOCALE_TAGS)("gives every locale-dependent id in %s", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    for (const id of DEPENDENT) {
      expect(builtinFormatCode(id, locale), `id ${id}`).toBe(LOCALIZED[tag][id]);
    }
    expect(builtinFormatCode(14, locale)).toBe(locale.shortDate);
    expect(builtinFormatCode(22, locale)).toBe(locale.shortDateTime);
  });

  it("gives the table itself in en-US", () => {
    for (const [id, code] of Object.entries(TABLE)) {
      expect(builtinFormatCode(Number(id), SPREADSHEET_LOCALES["en-US"]), `id ${id}`).toBe(code);
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("gives every locale-independent id unchanged in %s", (tag) => {
    for (const [id, code] of Object.entries(TABLE)) {
      if (!DEPENDENT.includes(Number(id))) {
        expect(builtinFormatCode(Number(id), SPREADSHEET_LOCALES[tag]), `id ${id}`).toBe(code);
      }
    }
  });

  it("has nothing for an id outside the table", () => {
    for (const id of [-1, 50, 82, 163, 164, 1000]) {
      expect(builtinFormatCode(id, SPREADSHEET_LOCALES["en-US"]), `id ${id}`).toBeUndefined();
    }
  });
});

describe("formatCodeOf", () => {
  const deDE = SPREADSHEET_LOCALES["de-DE"];

  it("is builtinFormatCode(builtinFormatId) when the style has one — over its numberFormat", () => {
    expect(formatCodeOf({ builtinFormatId: 14, numberFormat: "m/d/yyyy" }, deDE)).toBe("dd.mm.yyyy");
    expect(formatCodeOf({ builtinFormatId: 14 }, SPREADSHEET_LOCALES["en-GB"])).toBe("dd/mm/yyyy");
    expect(formatCodeOf({ builtinFormatId: 22 }, SPREADSHEET_LOCALES["cs-CZ"])).toBe("dd.mm.yyyy h:mm");
    expect(formatCodeOf({ builtinFormatId: 7, numberFormat: TABLE[7] }, deDE)).toBe('#,##0.00 "€";\\-#,##0.00 "€"');
    expect(formatCodeOf({ builtinFormatId: 3 }, deDE)).toBe("#,##0");
  });

  it("is the style's numberFormat otherwise; undefined (General) when it has neither", () => {
    expect(formatCodeOf({ numberFormat: "0.00" }, deDE)).toBe("0.00");
    expect(formatCodeOf({ numberFormat: "m/d/yyyy" }, deDE)).toBe("m/d/yyyy");
    expect(formatCodeOf({}, deDE)).toBeUndefined();
  });
});

describe("rendering the built-ins", () => {
  it.each<[tag: SpreadsheetLocaleTag, id14: string, id22: string]>([
    ["en-US", "10/8/2026", "10/8/2026 13:45"],
    ["en-GB", "08/10/2026", "08/10/2026 13:45"],
    ["de-DE", "08.10.2026", "08.10.2026 13:45"],
    ["fr-FR", "08/10/2026", "08/10/2026 13:45"],
    ["es-ES", "08/10/2026", "08/10/2026 13:45"],
    ["cs-CZ", "08.10.2026", "08.10.2026 13:45"],
  ])("a style on id 14 / 22 renders in %s as %j / %j", (tag, id14, id22) => {
    const locale = SPREADSHEET_LOCALES[tag];
    expect(formatValue(OCT_8_2026_1345, formatCodeOf({ builtinFormatId: 14, numberFormat: "m/d/yyyy" }, locale), context(tag)).text).toBe(id14);
    expect(formatValue(OCT_8_2026_1345, formatCodeOf({ builtinFormatId: 22, numberFormat: "m/d/yyyy h:mm" }, locale), context(tag)).text).toBe(id22);
  });

  it.each<[tag: SpreadsheetLocaleTag, id: number, value: number, text: string, red: boolean]>([
    ["en-US", 7, 1234.5, "$1,234.50 ", false],
    ["en-US", 8, -1234.5, "($1,234.50)", true],
    ["en-US", 5, 1234.5, "$1,235 ", false],
    ["en-GB", 8, -1234.5, "(£1,234.50)", true],
    ["de-DE", 7, 1234.5, "1.234,50 €", false],
    ["de-DE", 8, -1234.5, "-1.234,50 €", true],
    ["fr-FR", 7, 1234.5, "1\u202f234,50 €", false],
    ["es-ES", 6, -1234.5, "-1.235 €", true],
    ["cs-CZ", 7, -1234.5, "-1\u00a0234,50 Kč", false],
  ])("%s, id %i on %d → %j", (tag, id, value, text, red) => {
    const fv = formatValue(value, formatCodeOf({ builtinFormatId: id }, SPREADSHEET_LOCALES[tag]), context(tag));
    expect(fv.text).toBe(text);
    expect(fv.color).toBe(red ? "#ff0000" : undefined);
  });

  it.each([...range(23, 36), ...range(50, 81)])("id %i renders as General", (id) => {
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      const locale = SPREADSHEET_LOCALES[tag];
      const general = tag === "en-US" || tag === "en-GB" ? "1234.5" : "1234,5";
      expect(formatValue(1234.5, builtinFormatCode(id, locale), context(tag)).text).toBe(general);
      expect(formatValue(1234.5, formatCodeOf({ builtinFormatId: id }, locale), context(tag)).text).toBe(general);
    }
  });
});
