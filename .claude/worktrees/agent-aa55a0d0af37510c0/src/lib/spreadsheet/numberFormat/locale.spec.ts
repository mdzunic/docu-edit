// SPEC-118 "Locales": the six SpreadsheetLocales pinned field by field
// (LCID, decimal and group characters, list separator, date order and
// separators, the short/long date and time codes, AM/PM, currency symbol,
// position and code), DEFAULT_SPREADSHEET_LOCALE = "en-US", the LCID → tag
// map, and localeNames — the en-US and de-DE name tables as fixed strings
// (Node full-ICU; CLDR's names, so German March is "Mär"), the cs-CZ
// genitive ("8. října 2026" next to "říjen 2026"), built once per tag with
// Intl.DateTimeFormat(tag, { timeZone: "UTC", … }).formatToParts.
import { describe, expect, it, vi } from "vitest";
import {
  DEFAULT_SPREADSHEET_LOCALE,
  isSpreadsheetLocaleTag,
  localeNames,
  localeTagOfLcid,
  SPREADSHEET_LOCALE_NAMES,
  SPREADSHEET_LOCALE_TAGS,
  SPREADSHEET_LOCALES,
  spreadsheetLocaleFor,
  type SpreadsheetLocale,
  type SpreadsheetLocaleTag,
} from "./locale";
import { formatValue, type FormatContext } from "./render";

/** The Reference's table, row by row. */
const TABLE: Record<SpreadsheetLocaleTag, SpreadsheetLocale> = {
  "en-US": {
    tag: "en-US",
    lcid: 0x409,
    decimal: ".",
    group: ",",
    listSeparator: ",",
    dateOrder: "MDY",
    dateSeparator: "/",
    timeSeparator: ":",
    shortDate: "m/d/yyyy",
    longDate: "dddd, mmmm d, yyyy",
    shortTime: "h:mm AM/PM",
    longTime: "h:mm:ss AM/PM",
    shortDateTime: "m/d/yyyy h:mm",
    amPm: ["AM", "PM"],
    currencySymbol: "$",
    currencyPosition: "prefix",
    currencyCode: "USD",
  },
  "en-GB": {
    tag: "en-GB",
    lcid: 0x809,
    decimal: ".",
    group: ",",
    listSeparator: ",",
    dateOrder: "DMY",
    dateSeparator: "/",
    timeSeparator: ":",
    shortDate: "dd/mm/yyyy",
    longDate: "dd mmmm yyyy",
    shortTime: "hh:mm",
    longTime: "hh:mm:ss",
    shortDateTime: "dd/mm/yyyy hh:mm",
    amPm: ["AM", "PM"],
    currencySymbol: "£",
    currencyPosition: "prefix",
    currencyCode: "GBP",
  },
  "de-DE": {
    tag: "de-DE",
    lcid: 0x407,
    decimal: ",",
    group: ".",
    listSeparator: ";",
    dateOrder: "DMY",
    dateSeparator: ".",
    timeSeparator: ":",
    shortDate: "dd.mm.yyyy",
    longDate: "dddd, d. mmmm yyyy",
    shortTime: "hh:mm",
    longTime: "hh:mm:ss",
    shortDateTime: "dd.mm.yyyy hh:mm",
    amPm: ["AM", "PM"],
    currencySymbol: "€",
    currencyPosition: "suffix",
    currencyCode: "EUR",
  },
  "fr-FR": {
    tag: "fr-FR",
    lcid: 0x40c,
    decimal: ",",
    group: "\u202f",
    listSeparator: ";",
    dateOrder: "DMY",
    dateSeparator: "/",
    timeSeparator: ":",
    shortDate: "dd/mm/yyyy",
    longDate: "dddd d mmmm yyyy",
    shortTime: "hh:mm",
    longTime: "hh:mm:ss",
    shortDateTime: "dd/mm/yyyy hh:mm",
    amPm: ["AM", "PM"],
    currencySymbol: "€",
    currencyPosition: "suffix",
    currencyCode: "EUR",
  },
  "es-ES": {
    tag: "es-ES",
    lcid: 0xc0a,
    decimal: ",",
    group: ".",
    listSeparator: ";",
    dateOrder: "DMY",
    dateSeparator: "/",
    timeSeparator: ":",
    shortDate: "dd/mm/yyyy",
    longDate: 'dddd, d "de" mmmm "de" yyyy',
    shortTime: "h:mm",
    longTime: "h:mm:ss",
    shortDateTime: "dd/mm/yyyy h:mm",
    amPm: ["AM", "PM"],
    currencySymbol: "€",
    currencyPosition: "suffix",
    currencyCode: "EUR",
  },
  "cs-CZ": {
    tag: "cs-CZ",
    lcid: 0x405,
    decimal: ",",
    group: "\u00a0",
    listSeparator: ";",
    dateOrder: "DMY",
    dateSeparator: ".",
    timeSeparator: ":",
    shortDate: "dd.mm.yyyy",
    longDate: "dddd d. mmmm yyyy",
    shortTime: "h:mm",
    longTime: "h:mm:ss",
    shortDateTime: "dd.mm.yyyy h:mm",
    amPm: ["AM", "PM"],
    currencySymbol: "Kč",
    currencyPosition: "suffix",
    currencyCode: "CZK",
  },
};

function context(tag: SpreadsheetLocaleTag): FormatContext {
  return { locale: SPREADSHEET_LOCALES[tag], dateSystem: "1900" };
}

/** 2026-10-08, a Thursday. */
const OCT_8_2026 = 46303;

describe("the six locales", () => {
  it("are en-US, en-GB, de-DE, fr-FR, es-ES and cs-CZ, in that order", () => {
    expect(SPREADSHEET_LOCALE_TAGS).toEqual(["en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "cs-CZ"]);
    expect(Object.keys(SPREADSHEET_LOCALES).toSorted()).toEqual(SPREADSHEET_LOCALE_TAGS.toSorted());
  });

  it("default to en-US", () => {
    expect(DEFAULT_SPREADSHEET_LOCALE).toBe("en-US");
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("pin every field of %s", (tag) => {
    expect(SPREADSHEET_LOCALES[tag]).toEqual(TABLE[tag]);
  });

  it.each<[tag: SpreadsheetLocaleTag, field: keyof SpreadsheetLocale]>(
    SPREADSHEET_LOCALE_TAGS.flatMap((tag) => (Object.keys(TABLE[tag]) as Array<keyof SpreadsheetLocale>).map((field): [SpreadsheetLocaleTag, keyof SpreadsheetLocale] => [tag, field])),
  )("%s.%s", (tag, field) => {
    expect(SPREADSHEET_LOCALES[tag][field]).toEqual(TABLE[tag][field]);
  });

  it("group with U+202F (fr-FR) and U+00A0 (cs-CZ)", () => {
    expect(SPREADSHEET_LOCALES["fr-FR"].group).toHaveLength(1);
    expect(SPREADSHEET_LOCALES["fr-FR"].group.codePointAt(0)).toBe(0x202f);
    expect(SPREADSHEET_LOCALES["cs-CZ"].group).toHaveLength(1);
    expect(SPREADSHEET_LOCALES["cs-CZ"].group.codePointAt(0)).toBe(0x00a0);
  });

  it("all write AM/PM as AM and PM, and ':' between time parts", () => {
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(SPREADSHEET_LOCALES[tag].amPm, tag).toEqual(["AM", "PM"]);
      expect(SPREADSHEET_LOCALES[tag].timeSeparator, tag).toBe(":");
    }
  });

  it("name themselves in their own language", () => {
    expect(SPREADSHEET_LOCALE_NAMES["en-US"]).toBe("English (United States)");
    expect(SPREADSHEET_LOCALE_NAMES["de-DE"]).toBe("Deutsch (Deutschland)");
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(SPREADSHEET_LOCALE_NAMES[tag], tag).not.toBe("");
    }
  });

  it("render the preference's live samples", () => {
    const sample = (tag: SpreadsheetLocaleTag): string =>
      `${formatValue(1234.56, "#,##0.00", context(tag)).text} · ${formatValue(OCT_8_2026, SPREADSHEET_LOCALES[tag].shortDate, context(tag)).text}`;
    expect(sample("en-US")).toBe("1,234.56 · 10/8/2026");
    expect(sample("de-DE")).toBe("1.234,56 · 08.10.2026");
    expect(sample("fr-FR")).toBe("1\u202f234,56 · 08/10/2026");
    expect(sample("cs-CZ")).toBe("1\u00a0234,56 · 08.10.2026");
  });
});

describe("tags and LCIDs", () => {
  it.each<[lcid: number, tag: SpreadsheetLocaleTag]>([
    [0x409, "en-US"],
    [0x809, "en-GB"],
    [0x407, "de-DE"],
    [0x40c, "fr-FR"],
    [0xc0a, "es-ES"],
    [0x40a, "es-ES"],
    [0x405, "cs-CZ"],
  ])("LCID %d is %s", (lcid, tag) => {
    expect(localeTagOfLcid(lcid)).toBe(tag);
  });

  it("maps each locale's own LCID back to it", () => {
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(localeTagOfLcid(SPREADSHEET_LOCALES[tag].lcid)).toBe(tag);
    }
  });

  it("knows no other LCID", () => {
    for (const lcid of [0, 0x411, 0x804, 0x807, 0xf800, 0xf400, 0x10409]) {
      expect(localeTagOfLcid(lcid), lcid.toString(16)).toBeUndefined();
    }
  });

  it("validates tags exactly and falls back to the default for an unknown one", () => {
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(isSpreadsheetLocaleTag(tag)).toBe(true);
      expect(spreadsheetLocaleFor(tag)).toBe(SPREADSHEET_LOCALES[tag]);
    }
    for (const value of ["de", "en-us", "en_US", "", "pt-BR", 42, null, undefined]) {
      expect(isSpreadsheetLocaleTag(value), String(value)).toBe(false);
    }
    expect(spreadsheetLocaleFor("pt-BR")).toBe(SPREADSHEET_LOCALES["en-US"]);
  });
});

describe("localeNames", () => {
  it("pins en-US's names", () => {
    const names = localeNames("en-US");
    const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    expect(names.monthsLong).toEqual(months);
    expect(names.monthsGenitive).toEqual(months);
    expect(names.monthsShort).toEqual(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]);
    expect(names.daysLong).toEqual(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]);
    expect(names.daysShort).toEqual(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]);
  });

  it("pins de-DE's names (CLDR's: March is Mär, not Windows' Mrz)", () => {
    const names = localeNames("de-DE");
    const months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    expect(names.monthsLong).toEqual(months);
    expect(names.monthsGenitive).toEqual(months);
    expect(names.monthsShort).toEqual(["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]);
    expect(names.daysLong).toEqual(["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]);
    expect(names.daysShort).toEqual(["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"]);
  });

  it("keeps cs-CZ's genitive month apart from its nominative", () => {
    const names = localeNames("cs-CZ");
    expect(names.monthsLong[9]).toBe("říjen");
    expect(names.monthsGenitive[9]).toBe("října");
    expect(formatValue(OCT_8_2026, "d. mmmm yyyy", context("cs-CZ")).text).toBe("8. října 2026");
    expect(formatValue(OCT_8_2026, "mmmm yyyy", context("cs-CZ")).text).toBe("říjen 2026");
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("gives %s 12 months and 7 days (Sunday first), none empty", (tag) => {
    const names = localeNames(tag);
    for (const list of [names.monthsLong, names.monthsGenitive, names.monthsShort]) {
      expect(list).toHaveLength(12);
      expect(list.every((name) => name.length > 0)).toBe(true);
    }
    for (const list of [names.daysLong, names.daysShort]) {
      expect(list).toHaveLength(7);
      expect(list.every((name) => name.length > 0)).toBe(true);
    }
    // Sunday first: the first day name is Intl's for a Sunday (2026-10-11).
    const sunday = new Intl.DateTimeFormat(tag, { timeZone: "UTC", weekday: "long" }).format(new Date(Date.UTC(2026, 9, 11)));
    expect(names.daysLong[0]).toBe(sunday);
  });

  it("builds each tag's names once", () => {
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(localeNames(tag)).toBe(localeNames(tag));
    }
    expect(localeNames("de-DE")).not.toBe(localeNames("en-US"));
  });

  it("builds them with Intl.DateTimeFormat in UTC, constructing no formatter on a second call", async () => {
    vi.resetModules();
    const fresh = await import("./locale");
    const Original = Intl.DateTimeFormat;
    const spy = vi.spyOn(Intl, "DateTimeFormat").mockImplementation(function (locales?: Intl.LocalesArgument, options?: Intl.DateTimeFormatOptions) {
      return new Original(locales, options);
    });
    try {
      const names = fresh.localeNames("de-DE");
      expect(names.monthsShort[2]).toBe("Mär");
      const calls = spy.mock.calls.length;
      expect(calls).toBeGreaterThan(0);
      for (const [locales, options] of spy.mock.calls) {
        expect(locales).toBe("de-DE");
        expect(options?.timeZone).toBe("UTC");
      }
      expect(fresh.localeNames("de-DE")).toBe(names);
      expect(spy.mock.calls.length).toBe(calls);
    } finally {
      spy.mockRestore();
    }
    expect(Intl.DateTimeFormat).toBe(Original);
  });
});
