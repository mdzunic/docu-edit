// The spreadsheet's number-display locales (SPEC-118 "Locales"): the six
// locales a person may choose to see values in — never a property of a file
// (SPEC-066's rule: the locale is a per-user preference, resolved in the view
// by src/views/spreadsheet/useSpreadsheetLocale.ts and handed in; nothing in
// this directory reads preferences). A locale changes how values look and
// (SPEC-122) how typed values are read, never formula text and never what is
// saved.
//
// Each locale pins its separators, its date order, the codes it formats
// dates and times with (the invariant format-code syntax: `.` and `,` mean
// what they mean in any code) and its currency. Month and day names are
// CLDR's, read from Intl.DateTimeFormat(tag, { timeZone: "UTC", … })
// .formatToParts once per tag — not Windows' (German `mmm` for March is
// "Mär" where German Excel shows "Mrz"; accepted and stated).

export type SpreadsheetLocaleTag = "en-US" | "en-GB" | "de-DE" | "fr-FR" | "es-ES" | "cs-CZ";

/** The six locales, in the order every picker lists them. */
export const SPREADSHEET_LOCALE_TAGS: readonly SpreadsheetLocaleTag[] = ["en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "cs-CZ"];

/** The locale nobody chose: en-US, deliberately (never detected from the
 * browser). */
export const DEFAULT_SPREADSHEET_LOCALE: SpreadsheetLocaleTag = "en-US";

export interface SpreadsheetLocale {
  tag: SpreadsheetLocaleTag;
  lcid: number;
  decimal: string;
  group: string;
  /** CSV's field separator (SPEC-126) — never a formula's argument separator. */
  listSeparator: string;
  dateOrder: "MDY" | "DMY" | "YMD";
  dateSeparator: string;
  timeSeparator: ":";
  // Format codes, in the invariant syntax.
  shortDate: string;
  longDate: string;
  shortTime: string;
  longTime: string;
  shortDateTime: string;
  /** "AM", "PM" for all six (SPEC-122's input accepts either case). */
  amPm: [string, string];
  currencySymbol: string;
  /** suffix = after a space. */
  currencyPosition: "prefix" | "suffix";
  currencyCode: string;
}

const NARROW_NO_BREAK_SPACE = "\u202f";
const NO_BREAK_SPACE = "\u00a0";

export const SPREADSHEET_LOCALES: Readonly<Record<SpreadsheetLocaleTag, SpreadsheetLocale>> = {
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
    group: NARROW_NO_BREAK_SPACE,
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
    group: NO_BREAK_SPACE,
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

/** Each locale's name in its own language, for the pickers. */
export const SPREADSHEET_LOCALE_NAMES: Readonly<Record<SpreadsheetLocaleTag, string>> = {
  "en-US": "English (United States)",
  "en-GB": "English (United Kingdom)",
  "de-DE": "Deutsch (Deutschland)",
  "fr-FR": "Français (France)",
  "es-ES": "Español (España)",
  "cs-CZ": "Čeština (Česko)",
};

/** Whether a value is one of the six tags. */
export function isSpreadsheetLocaleTag(value: unknown): value is SpreadsheetLocaleTag {
  return typeof value === "string" && (SPREADSHEET_LOCALE_TAGS as readonly string[]).includes(value);
}

/** The locale a tag names; an unknown tag is the default locale. */
export function spreadsheetLocaleFor(tag: string): SpreadsheetLocale {
  return SPREADSHEET_LOCALES[isSpreadsheetLocaleTag(tag) ? tag : DEFAULT_SPREADSHEET_LOCALE];
}

/** The locale an Excel LCID names (its language-and-region word; the sort
 * and calendar bytes above 0xFFFF are the caller's to refuse). es-ES has two:
 * 0xC0A (modern sort) and 0x40A (traditional sort). */
export function localeTagOfLcid(lcid: number): SpreadsheetLocaleTag | undefined {
  switch (lcid) {
    case 0x409:
      return "en-US";
    case 0x809:
      return "en-GB";
    case 0x407:
      return "de-DE";
    case 0x40c:
      return "fr-FR";
    case 0xc0a:
    case 0x40a:
      return "es-ES";
    case 0x405:
      return "cs-CZ";
    default:
      return undefined;
  }
}

/** The month and day names a locale's codes render (`mmm`, `mmmm`, `ddd`,
 * `dddd`), Sunday first. `monthsGenitive` is the form a month takes next to
 * a day number ("8. října 2026"); most languages spell it as the nominative. */
export interface LocaleNames {
  monthsLong: readonly string[];
  monthsGenitive: readonly string[];
  monthsShort: readonly string[];
  daysLong: readonly string[];
  daysShort: readonly string[];
}

const namesByTag = new Map<SpreadsheetLocaleTag, LocaleNames>();

/** One part of a date as Intl formats it for `tag`. */
/** The `type` part of each date, from ONE formatter for the option set (a
 * formatter costs far more to build than to use: the five tables built per
 * formatter took ~17 ms of a large workbook's first paint built per name). */
function nameParts(tag: SpreadsheetLocaleTag, options: Intl.DateTimeFormatOptions, dates: readonly Date[], type: "month" | "weekday"): string[] {
  const format = new Intl.DateTimeFormat(tag, { timeZone: "UTC", ...options });
  return dates.map((date) => format.formatToParts(date).find((part) => part.type === type)?.value ?? "");
}

/** A locale's name tables, built once per tag from Intl.DateTimeFormat over
 * 2001-mm-15 (months; the genitive is the month part of a "day month"
 * format) and 2001-01-07 + d (2001-01-07 is a Sunday). */
export function localeNames(tag: SpreadsheetLocaleTag): LocaleNames {
  let names = namesByTag.get(tag);
  if (names) {
    return names;
  }
  const months = Array.from({ length: 12 }, (_, month) => new Date(Date.UTC(2001, month, 15)));
  const days = Array.from({ length: 7 }, (_, day) => new Date(Date.UTC(2001, 0, 7 + day)));
  names = {
    monthsLong: nameParts(tag, { month: "long" }, months, "month"),
    monthsGenitive: nameParts(tag, { month: "long", day: "numeric" }, months, "month"),
    monthsShort: nameParts(tag, { month: "short" }, months, "month"),
    daysLong: nameParts(tag, { weekday: "long" }, days, "weekday"),
    daysShort: nameParts(tag, { weekday: "short" }, days, "weekday"),
  };
  namesByTag.set(tag, names);
  return names;
}
