// Number-format presets, decimal stepping and the Format Cells code builders
// (SPEC-123 "Number-format presets", "Decimal stepping", "Code builders") —
// the codes behind the sheet toolbar's number-format dropdown and its
// currency, percent, comma and decimal buttons, the number-format shortcuts
// and Format Cells' Number tab. Pure: codes in, codes out.
//
// - Codes are SPEC-118's invariant syntax (`.` decimal, `,` group); the
//   locale picks which code (its currency symbol and side, its long date,
//   its clock), never how it is spelled. A choice whose code is one of
//   Excel's locale-dependent built-ins (LOCALE_DEPENDENT_BUILTINS) carries
//   that id too (numberFormatChoiceOf), so it renders in the reader's locale
//   and the xlsx writer writes the id; General is `numberFormat: null`.
// - LOCALE_CURRENCY holds each locale's symbol, its side and whether its
//   negatives go in parentheses — here rather than widening SPEC-118's
//   SpreadsheetLocale. The parentheses pick the Accounting template: en-US
//   has Excel's id 41–44 shape (`_(` pads, negatives in `\(…\)`), the others
//   the `_-` shape with a leading minus; the symbol is quoted before the
//   number, or after it behind a space.
// - stepDecimals edits a code at the source offsets of SPEC-118's tokens
//   (formatCodeSections), never by a pattern over its text, so quoted
//   literals, escapes, `_x`, `*x` and brackets keep every character: each
//   number section (a scientific one's mantissa only) gains or loses one `0`
//   after its last decimal placeholder, inserting or removing the point;
//   text, date/time, fraction and General sections are left alone. An
//   accounting zero section — the third section, a quoted `"-"` padded with
//   a `?` per decimal — gains or loses a `?` instead, so stepping id 44 down
//   twice is id 42 exactly, and a code the builders make steps to the one
//   they make with a decimal more or less. General steps from the decimals
//   generalText shows. An invalid code, or one with nothing to step, comes
//   back as the identical string.
// - buildNumberCode turns the Number tab's options into a code;
//   describeNumberCode looks a code up in a per-locale map of every code the
//   option sets build (the first in category order wins), so
//   build(describe(code)) === code for every built code and any other code
//   opens on Custom. The one alias: Currency with no symbol builds Number
//   with thousands' code, so it describes as Number.
import { generalText } from "./general";
import type { CellStyle, CellValue } from "./model";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "./numberFormat/builtin";
import type { SpreadsheetLocale, SpreadsheetLocaleTag } from "./numberFormat/locale";
import { FormatCodeError, formatCodeSections, parseFormatCode, type FormatCodeSection, type ScannedToken } from "./numberFormat/parse";

/** A number format as a style patch sets it: the invariant code (null =
 * General) and, for a locale-dependent Excel built-in
 * (LOCALE_DEPENDENT_BUILTINS), its id. */
export interface NumberFormatChoice {
  numberFormat: string | null;
  builtinFormatId?: number;
}

/** A locale's currency (SPEC-123's LOCALE_CURRENCY). */
export interface LocaleCurrency {
  symbol: string;
  position: "before" | "after";
  /** Negatives in parentheses rather than behind a minus. */
  negativeParens: boolean;
}

export const LOCALE_CURRENCY: Readonly<Record<SpreadsheetLocaleTag, LocaleCurrency>> = {
  "en-US": { symbol: "$", position: "before", negativeParens: true },
  "en-GB": { symbol: "£", position: "before", negativeParens: false },
  "de-DE": { symbol: "€", position: "after", negativeParens: false },
  "fr-FR": { symbol: "€", position: "after", negativeParens: false },
  "es-ES": { symbol: "€", position: "after", negativeParens: false },
  "cs-CZ": { symbol: "Kč", position: "after", negativeParens: false },
};

/** Each locale's Long date — SPEC-123's table, not SpreadsheetLocale.longDate
 * (whose en-GB spelling has no weekday). */
const LONG_DATE_CODES: Readonly<Record<SpreadsheetLocaleTag, string>> = {
  "en-US": "dddd, mmmm d, yyyy",
  "en-GB": "dddd, d mmmm yyyy",
  "de-DE": "dddd, d. mmmm yyyy",
  "fr-FR": "dddd d mmmm yyyy",
  "es-ES": 'dddd, d "de" mmmm "de" yyyy',
  "cs-CZ": "dddd d. mmmm yyyy",
};

/** Whether a code renders as General: absent, blank or "General" in any
 * case (render.ts's rule). */
function isGeneralCode(code: string | null | undefined): boolean {
  return code === undefined || code === null || code.trim() === "" || code.trim().toLowerCase() === "general";
}

/** A code's decimal part: nothing for no decimals, else the point and
 * `decimals` zeros. */
function decimalPart(decimals: number): string {
  return decimals > 0 ? `.${"0".repeat(decimals)}` : "";
}

/** A code as a style patch sets it: undefined, null, blank or "General" →
 * General; a code equal to a locale-dependent built-in's → with its id;
 * anything else as it is. */
export function numberFormatChoiceOf(code: string | undefined | null): NumberFormatChoice {
  if (code === undefined || code === null || isGeneralCode(code)) {
    return { numberFormat: null };
  }
  const id = LOCALE_DEPENDENT_BUILTINS.find((candidate) => BUILTIN_NUMBER_FORMATS[candidate] === code);
  return id === undefined ? { numberFormat: code } : { numberFormat: code, builtinFormatId: id };
}

/** Whether SPEC-118's grammar accepts a code (parseFormatCode does not
 * throw). */
export function isValidFormatCode(code: string): boolean {
  try {
    parseFormatCode(code);
    return true;
  } catch {
    return false;
  }
}

// --- presets, toolbar buttons and shortcuts -------------------------------------------

export type NumberPresetId =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "short-date"
  | "long-date"
  | "time"
  | "percentage"
  | "fraction"
  | "scientific"
  | "text";

/** The number-format dropdown's entries in order (`sheet-number-format-<id>`). */
export const NUMBER_FORMAT_PRESETS: readonly { id: NumberPresetId; label: string }[] = [
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
];

/** The code a preset applies in `locale` (undefined = General). Currency and
 * Accounting are the builders' two-decimal codes with the locale's symbol —
 * en-US's Accounting is id 44's code exactly. */
function presetCode(id: NumberPresetId, locale: SpreadsheetLocale): string | undefined {
  switch (id) {
    case "general":
      return undefined;
    case "number":
      return BUILTIN_NUMBER_FORMATS[2];
    case "currency":
      return buildNumberCode({ category: "currency", decimals: 2, negative: "minus" }, locale);
    case "accounting":
      return buildNumberCode({ category: "accounting", decimals: 2 }, locale);
    case "short-date":
      return BUILTIN_NUMBER_FORMATS[14]; // rendered as the locale's short date
    case "long-date":
      return LONG_DATE_CODES[locale.tag];
    case "time":
      return BUILTIN_NUMBER_FORMATS[locale.tag === "en-US" ? 19 : 21];
    case "percentage":
      return BUILTIN_NUMBER_FORMATS[10];
    case "fraction":
      return BUILTIN_NUMBER_FORMATS[12];
    case "scientific":
      return BUILTIN_NUMBER_FORMATS[11];
    case "text":
      return BUILTIN_NUMBER_FORMATS[49];
  }
}

/** What a number-format preset applies in `locale`. */
export function presetNumberFormat(id: NumberPresetId, locale: SpreadsheetLocale): NumberFormatChoice {
  return numberFormatChoiceOf(presetCode(id, locale));
}

/** The preset whose choice is exactly the style's format (General when the
 * style has none), else null — what the dropdown marks as current. */
export function presetIdOf(style: Pick<CellStyle, "numberFormat" | "builtinFormatId">, locale: SpreadsheetLocale): NumberPresetId | null {
  if (style.builtinFormatId === undefined && isGeneralCode(style.numberFormat)) {
    return "general";
  }
  const preset = NUMBER_FORMAT_PRESETS.find(({ id }) => {
    const choice = presetNumberFormat(id, locale);
    return choice.numberFormat === style.numberFormat && choice.builtinFormatId === style.builtinFormatId;
  });
  return preset?.id ?? null;
}

/** `sheet-currency`: the locale's Accounting. */
export function currencyButtonFormat(locale: SpreadsheetLocale): NumberFormatChoice {
  return presetNumberFormat("accounting", locale);
}

/** `sheet-percent`: id 9, `0%`. */
export function percentButtonFormat(): NumberFormatChoice {
  return numberFormatChoiceOf(BUILTIN_NUMBER_FORMATS[9]);
}

/** `sheet-comma`: the locale's Accounting without the symbol (and its
 * space) — id 43 in en-US. */
export function commaButtonFormat(locale: SpreadsheetLocale): NumberFormatChoice {
  return numberFormatChoiceOf(buildNumberCode({ category: "accounting", decimals: 2, symbol: "" }, locale));
}

export type NumberShortcutId = "general" | "number" | "time" | "date" | "currency" | "percent" | "scientific";

/** What a number-format shortcut (SPEC-123 "Shortcuts") applies in
 * `locale`. */
export function shortcutNumberFormat(id: NumberShortcutId, locale: SpreadsheetLocale): NumberFormatChoice {
  switch (id) {
    case "general":
      return { numberFormat: null };
    case "number":
      return numberFormatChoiceOf(BUILTIN_NUMBER_FORMATS[4]);
    case "time":
      return numberFormatChoiceOf(BUILTIN_NUMBER_FORMATS[18]);
    case "date":
      return numberFormatChoiceOf(BUILTIN_NUMBER_FORMATS[15]);
    case "currency":
      return locale.tag === "en-US" ? numberFormatChoiceOf(BUILTIN_NUMBER_FORMATS[7]) : presetNumberFormat("currency", locale);
    case "percent":
      return percentButtonFormat();
    case "scientific":
      return numberFormatChoiceOf(BUILTIN_NUMBER_FORMATS[11]);
  }
}

// --- decimal stepping --------------------------------------------------------------

/** One edit of a code: `remove` characters at `at`, then `insert` there. */
interface CodeEdit {
  at: number;
  remove: number;
  insert: string;
}

const QUOTED_DASH = '"-"';
/** The section zero shows in a code of three or four sections — where an
 * accounting format puts its dash. */
const ZERO_SECTION = 2;

/** Whether a token is the quoted `"-"` of an accounting zero section. */
function isQuotedDash(code: string, { token, offset }: ScannedToken): boolean {
  return token.type === "literal" && token.text === "-" && code.startsWith(QUOTED_DASH, offset);
}

/** Whether a token is a digit placeholder (`0`, `#`, `?`). */
function isPlaceholder({ token }: ScannedToken): boolean {
  return token.type === "digit";
}

/** The decimals General shows for a value: the digits after the point of its
 * generalText — a scientific text's mantissa — and 0 for anything but a
 * number. */
function generalDecimals(value: CellValue): number {
  if (typeof value !== "number") {
    return 0;
  }
  const mantissa = generalText(value).split("E")[0]!;
  const point = mantissa.indexOf(".");
  return point < 0 ? 0 : mantissa.length - point - 1;
}

/** The edits stepping section `index` of `code` makes (see the module
 * header). */
function sectionEdits(code: string, { kind, tokens }: FormatCodeSection, index: number, dir: 1 | -1): CodeEdit[] {
  const dash = index === ZERO_SECTION ? tokens.findLast((entry) => isQuotedDash(code, entry)) : undefined;
  if (kind === "literal") {
    // An accounting zero section with no decimals (id 42's `_("$"* "-"_)`)
    // gains its first `?` right after the dash.
    return dir > 0 && dash !== undefined ? [{ at: dash.offset + QUOTED_DASH.length, remove: 0, insert: "?" }] : [];
  }
  if (kind !== "number" && kind !== "scientific") {
    return [];
  }
  const exponent = tokens.findIndex(({ token }) => token.type === "exponent");
  const mantissa = exponent < 0 ? tokens : tokens.slice(0, exponent);
  const pointIndex = mantissa.findIndex(({ token }) => token.type === "point");
  const point = pointIndex < 0 ? undefined : mantissa[pointIndex]!;
  const integers = (point === undefined ? mantissa : mantissa.slice(0, pointIndex)).filter(isPlaceholder);
  const decimals = point === undefined ? [] : mantissa.slice(pointIndex + 1).filter(isPlaceholder);
  const padsOnly = integers.every(({ token }) => token.type === "digit" && token.placeholder === "?");
  if (kind === "number" && dash !== undefined && point === undefined && padsOnly) {
    // An accounting zero section: its `?`s pad where the decimals would be.
    const last = integers.at(-1)!;
    return dir > 0 ? [{ at: last.offset + 1, remove: 0, insert: "?" }] : [{ at: last.offset, remove: 1, insert: "" }];
  }
  if (dir > 0) {
    if (point !== undefined) {
      return [{ at: (decimals.at(-1) ?? point).offset + 1, remove: 0, insert: "0" }];
    }
    const last = integers.at(-1);
    return last === undefined ? [] : [{ at: last.offset + 1, remove: 0, insert: ".0" }];
  }
  const last = decimals.at(-1);
  if (point === undefined || last === undefined) {
    return [];
  }
  if (decimals.length > 1) {
    return [{ at: last.offset, remove: 1, insert: "" }];
  }
  // The only decimal goes with its point — or, with no integer placeholder
  // to keep the section a number (`.0`), becomes one.
  const removePoint = { at: point.offset, remove: 1, insert: "" };
  return integers.length > 0 ? [removePoint, { at: last.offset, remove: 1, insert: "" }] : [removePoint];
}

/** Increase (+1) or Decrease (−1) Decimal on a code (see the module
 * header). `code` is the one the active cell renders with (formatCodeOf in
 * the user's locale), so a locale-dependent built-in steps as it shows — and
 * a result equal to `code` is nothing to apply; `value` is the cell's, read
 * only when the code is General (undefined, blank or "General"): +1 → `0.`
 * and one zero more than its generalText shows, −1 → one fewer (`0` at
 * none). */
export function stepDecimals(code: string | undefined, value: CellValue, dir: 1 | -1): string {
  if (code === undefined || isGeneralCode(code)) {
    return `0${decimalPart(generalDecimals(value) + dir)}`;
  }
  let sections: FormatCodeSection[];
  try {
    sections = formatCodeSections(code);
  } catch (error) {
    if (error instanceof FormatCodeError) {
      return code;
    }
    throw error;
  }
  // Applied from the end, so every earlier offset still holds.
  const edits = sections.flatMap((section, index) => sectionEdits(code, section, index, dir)).toSorted((a, b) => b.at - a.at);
  return edits.reduce((text, edit) => text.slice(0, edit.at) + edit.insert + text.slice(edit.at + edit.remove), code);
}

// --- Format Cells' code builders ---------------------------------------------------

export type NumberCategory =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "date"
  | "time"
  | "percentage"
  | "fraction"
  | "scientific"
  | "text"
  | "custom";

export type NegativeStyle = "minus" | "red" | "parens" | "redParens";

/** The Number tab's options (`format-cells-<field>`); each category reads
 * only its own. */
export interface NumberCodeSpec {
  category: NumberCategory;
  /** 0–30, default 2: number, currency, accounting, percentage, scientific. */
  decimals?: number;
  /** Number only (Currency always groups); default off. */
  thousands?: boolean;
  /** Number, currency; default "minus". */
  negative?: NegativeStyle;
  /** Currency, accounting: "" = None, else one of CURRENCY_SYMBOLS; absent
   * = the locale's. */
  symbol?: string;
  /** Date, time, fraction: the chosen type's code (one of the category's
   * list, else its first); custom: the code itself. */
  type?: string;
}

/** The category listbox in order (`format-cells-category`). */
export const NUMBER_CATEGORIES: readonly { id: NumberCategory; label: string }[] = [
  { id: "general", label: "General" },
  { id: "number", label: "Number" },
  { id: "currency", label: "Currency" },
  { id: "accounting", label: "Accounting" },
  { id: "date", label: "Date" },
  { id: "time", label: "Time" },
  { id: "percentage", label: "Percentage" },
  { id: "fraction", label: "Fraction" },
  { id: "scientific", label: "Scientific" },
  { id: "text", label: "Text" },
  { id: "custom", label: "Custom" },
];

/** The symbols Currency and Accounting offer besides None ("") — every
 * locale's own symbol among them. */
export const CURRENCY_SYMBOLS: readonly string[] = ["$", "€", "£", "¥", "Kč", "CHF"];

const SYMBOL_POSITIONS: ReadonlyMap<string, "before" | "after"> = new Map<string, "before" | "after">([
  ["$", "before"],
  ["£", "before"],
  ["¥", "before"],
  ["€", "after"],
  ["Kč", "after"],
  ["CHF", "after"],
]);

/** Which side of the number a symbol goes in `locale`: the locale's own
 * symbol on LOCALE_CURRENCY's side; `$ £ ¥` before, `€ Kč CHF` after; any
 * other on the locale's side. */
export function currencySymbolPosition(symbol: string, locale: SpreadsheetLocale): "before" | "after" {
  const own = LOCALE_CURRENCY[locale.tag];
  return symbol === own.symbol ? own.position : (SYMBOL_POSITIONS.get(symbol) ?? own.position);
}

/** The Date category's types in `locale`: ids 14, 15, 16, 17 and 22, the
 * locale's Long date, ISO's `yyyy-mm-dd` and `mmmm yyyy`. */
export function dateTypeCodes(locale: SpreadsheetLocale): readonly string[] {
  return [...[14, 15, 16, 17, 22].map((id) => BUILTIN_NUMBER_FORMATS[id]!), LONG_DATE_CODES[locale.tag], "yyyy-mm-dd", "mmmm yyyy"];
}

/** The Time category's types: ids 18–21 and 45–47. */
export const TIME_TYPE_CODES: readonly string[] = [18, 19, 20, 21, 45, 46, 47].map((id) => BUILTIN_NUMBER_FORMATS[id]!);

/** The Fraction category's types. */
export const FRACTION_TYPE_CODES: readonly string[] = ["# ?/?", "# ??/??", "# ???/???", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10", "# ??/100"];

const MAX_DECIMALS = 30;
const NEGATIVE_STYLES: readonly NegativeStyle[] = ["minus", "red", "parens", "redParens"];

/** A spec's decimals: whole, clamped to 0–30, 2 when absent. */
function decimalsOf(spec: NumberCodeSpec): number {
  const decimals = spec.decimals ?? 2;
  return Number.isFinite(decimals) ? Math.min(MAX_DECIMALS, Math.max(0, Math.trunc(decimals))) : 2;
}

/** A currency symbol as a quoted literal. */
function quoted(symbol: string): string {
  return `"${symbol.replaceAll('"', '\\"')}"`;
}

/** Section `c` with the negative style's second section. */
function withNegative(c: string, negative: NegativeStyle): string {
  switch (negative) {
    case "minus":
      return c;
    case "red":
      return `${c};[Red]${c}`;
    case "parens":
      return `${c}_);(${c})`;
    case "redParens":
      return `${c}_);[Red](${c})`;
  }
}

/** Accounting in `locale` (see the module header): its template with the
 * symbol — none for "" — on its side, `decimals` zeros, and as many `?`
 * after the zero section's dash. */
function accountingCode(decimals: number, symbol: string, locale: SpreadsheetLocale): string {
  const number = `#,##0${decimalPart(decimals)}`;
  const pads = "?".repeat(decimals);
  const after = symbol !== "" && currencySymbolPosition(symbol, locale) === "after";
  const lead = symbol === "" || after ? "" : quoted(symbol);
  const tail = after ? ` ${quoted(symbol)}` : "";
  if (LOCALE_CURRENCY[locale.tag].negativeParens) {
    return `_(${lead}* ${number}${tail}_);_(${lead}* \\(${number}${tail}\\);_(${lead}* "-"${pads}${tail}_);_(@_)`;
  }
  return `_-${lead}* ${number}${tail}_-;-${lead}* ${number}${tail}_-;_-${lead}* "-"${pads}${tail}_-;_-@_-`;
}

/** `type` when it is one of `codes`, else the first. */
function typeIn(codes: readonly string[], type: string | undefined): string {
  return type !== undefined && codes.includes(type) ? type : codes[0]!;
}

/** The code the Number tab's options make in `locale` (undefined = General;
 * see "Code builders"). */
export function buildNumberCode(spec: NumberCodeSpec, locale: SpreadsheetLocale): string | undefined {
  const decimals = decimalsOf(spec);
  const symbol = spec.symbol ?? LOCALE_CURRENCY[locale.tag].symbol;
  switch (spec.category) {
    case "general":
      return undefined;
    case "number":
      return withNegative(`${spec.thousands === true ? "#,##0" : "0"}${decimalPart(decimals)}`, spec.negative ?? "minus");
    case "currency": {
      const number = `#,##0${decimalPart(decimals)}`;
      const section =
        symbol === "" ? number : currencySymbolPosition(symbol, locale) === "after" ? `${number} ${quoted(symbol)}` : `${quoted(symbol)}${number}`;
      return withNegative(section, spec.negative ?? "minus");
    }
    case "accounting":
      return accountingCode(decimals, symbol, locale);
    case "date":
      return typeIn(dateTypeCodes(locale), spec.type);
    case "time":
      return typeIn(TIME_TYPE_CODES, spec.type);
    case "percentage":
      return `0${decimalPart(decimals)}%`;
    case "fraction":
      return typeIn(FRACTION_TYPE_CODES, spec.type);
    case "scientific":
      return `0${decimalPart(decimals)}E+00`;
    case "text":
      return "@";
    case "custom":
      return spec.type === undefined || isGeneralCode(spec.type) ? undefined : spec.type;
  }
}

/** Every option set the Number tab offers in `locale`, in category order —
 * Currency without its None, which is Number with thousands. */
function optionSets(locale: SpreadsheetLocale): NumberCodeSpec[] {
  const decimalsRange = Array.from({ length: MAX_DECIMALS + 1 }, (_, decimals) => decimals);
  const specs: NumberCodeSpec[] = [];
  for (const decimals of decimalsRange) {
    for (const thousands of [false, true]) {
      for (const negative of NEGATIVE_STYLES) {
        specs.push({ category: "number", decimals, thousands, negative });
      }
    }
  }
  for (const decimals of decimalsRange) {
    for (const negative of NEGATIVE_STYLES) {
      for (const symbol of CURRENCY_SYMBOLS) {
        specs.push({ category: "currency", decimals, negative, symbol });
      }
    }
  }
  for (const decimals of decimalsRange) {
    for (const symbol of ["", ...CURRENCY_SYMBOLS]) {
      specs.push({ category: "accounting", decimals, symbol });
    }
  }
  specs.push(...dateTypeCodes(locale).map((type): NumberCodeSpec => ({ category: "date", type })));
  specs.push(...TIME_TYPE_CODES.map((type): NumberCodeSpec => ({ category: "time", type })));
  specs.push(...decimalsRange.map((decimals): NumberCodeSpec => ({ category: "percentage", decimals })));
  specs.push(...FRACTION_TYPE_CODES.map((type): NumberCodeSpec => ({ category: "fraction", type })));
  specs.push(...decimalsRange.map((decimals): NumberCodeSpec => ({ category: "scientific", decimals })));
  specs.push({ category: "text" });
  return specs;
}

const builtCodesByLocale = new Map<SpreadsheetLocaleTag, ReadonlyMap<string, NumberCodeSpec>>();

/** Every code the option sets build in `locale` → the first spec building
 * it; built once per locale. */
function builtCodes(locale: SpreadsheetLocale): ReadonlyMap<string, NumberCodeSpec> {
  let codes = builtCodesByLocale.get(locale.tag);
  if (codes === undefined) {
    const map = new Map<string, NumberCodeSpec>();
    for (const spec of optionSets(locale)) {
      const code = buildNumberCode(spec, locale)!;
      if (!map.has(code)) {
        map.set(code, spec);
      }
    }
    codes = map;
    builtCodesByLocale.set(locale.tag, codes);
  }
  return codes;
}

/** The Number tab's options a code opens on in `locale`: General for none
 * (or blank, or "General"), the option set that builds it, else Custom with
 * the code. `code` is the style's stored one — ids 14 and 22 are in the
 * Date list by their stored spelling, whatever the locale shows. */
export function describeNumberCode(code: string | undefined, locale: SpreadsheetLocale): NumberCodeSpec {
  if (code === undefined || isGeneralCode(code)) {
    return { category: "general" };
  }
  const spec = builtCodes(locale).get(code);
  return spec === undefined ? { category: "custom", type: code } : { ...spec };
}
