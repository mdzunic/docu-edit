// Excel's built-in number formats (SPEC-118 "Built-in formats"): the codes a
// `numFmtId` below 164 stands for when the file writes no `<numFmt>` for it.
// The table holds what en-US Excel displays — ECMA-376 §18.8.30 lists 14 as
// `mm-dd-yy`, but Excel renders and SpreadsheetML consumers expect the
// locale's short date. Ten ids follow the reader's locale (5–8, 14, 22,
// 41–44): a style keeps such an id in `builtinFormatId` and renders it in the
// current locale. 23–36 and 50–81 (the CJK built-ins) have no entry and
// render as General.
import type { CellStyle } from "../model";
import type { SpreadsheetLocale } from "./locale";

export const BUILTIN_NUMBER_FORMATS: Readonly<Record<number, string>> = {
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

/** The built-ins whose meaning follows the reader's locale. */
export const LOCALE_DEPENDENT_BUILTINS: readonly number[] = [5, 6, 7, 8, 14, 22, 41, 42, 43, 44];

/** The currency built-ins of a locale that writes its symbol after the
 * number (after a space); `S` is the quoted symbol. */
const SUFFIX_CURRENCY_TEMPLATES: Readonly<Record<number, string>> = {
  5: "#,##0 S;\\-#,##0 S",
  6: "#,##0 S;[Red]\\-#,##0 S",
  7: "#,##0.00 S;\\-#,##0.00 S",
  8: "#,##0.00 S;[Red]\\-#,##0.00 S",
  41: '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-',
  42: '_-* #,##0 S_-;\\-* #,##0 S_-;_-* "-" S_-;_-@_-',
  43: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
  44: '_-* #,##0.00 S_-;\\-* #,##0.00 S_-;_-* "-"?? S_-;_-@_-',
};

/** The code a built-in id stands for in `locale` (undefined for an id with
 * no entry): 14 and 22 are the locale's short date (and time); the currency
 * ids carry the locale's symbol, before the number or — for a suffix locale
 * — after it. Every other id is locale-independent. */
export function builtinFormatCode(id: number, locale: SpreadsheetLocale): string | undefined {
  if (id === 14) {
    return locale.shortDate;
  }
  if (id === 22) {
    return locale.shortDateTime;
  }
  const code = BUILTIN_NUMBER_FORMATS[id];
  if (code === undefined || !LOCALE_DEPENDENT_BUILTINS.includes(id)) {
    return code;
  }
  const symbol = `"${locale.currencySymbol}"`;
  if (locale.currencyPosition === "prefix") {
    return code.replaceAll('"$"', symbol);
  }
  return SUFFIX_CURRENCY_TEMPLATES[id]!.replaceAll("S", symbol);
}

/** The code a cell style renders with in `locale`: its locale-dependent
 * built-in when it has one, else its `numberFormat` (undefined = General). */
export function formatCodeOf(style: Pick<CellStyle, "numberFormat" | "builtinFormatId">, locale: SpreadsheetLocale): string | undefined {
  if (style.builtinFormatId !== undefined) {
    const code = builtinFormatCode(style.builtinFormatId, locale);
    if (code !== undefined) {
      return code;
    }
  }
  return style.numberFormat;
}
