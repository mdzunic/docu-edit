// Text, SPEC-127's library: CHAR CODE UNICHAR UNICODE CLEAN FIXED DOLLAR
// NUMBERVALUE T TEXTBEFORE TEXTAFTER VALUETOTEXT ARRAYTOTEXT. (TEXTSPLIT is
// not registered: it spills.)
//
// - CHAR(1…255) and CODE map through Windows-1252, as Excel for Windows and
//   the web do (WHATWG's index: 0x80 is €, 0x81 stays U+0081); CODE of a
//   character with no Windows-1252 byte is 63 ("?"). UNICHAR/UNICODE work in
//   code points, a surrogate pair whole: UNICHAR of 0 or above U+10FFFF is
//   #VALUE!, of a surrogate #N/A. CLEAN removes U+0000–U+001F.
// - FIXED(number, [decimals=2], [no_commas]) and DOLLAR(number,
//   [decimals=2]) round half away from zero, then render through SPEC-118's
//   formatValue in the user's locale as TEXT does. DOLLAR's codes, decimals
//   substituted: en-US `$#,##0.00;($#,##0.00)`, en-GB `£#,##0.00;-£#,##0.00`,
//   de-DE/fr-FR/es-ES `#,##0.00 €;-#,##0.00 €`, cs-CZ
//   `#,##0.00 Kč;-#,##0.00 Kč`.
// - NUMBERVALUE reads a number with the given (default: the locale's)
//   decimal and group separators — spaces anywhere ignored, group separators
//   only before the decimal one, each trailing % a division by 100, "" is 0.
// - TEXTBEFORE/TEXTAFTER take all six of Excel's arguments: a delimiter or
//   an array of them, the instance (negative: from the end; 0 or past the
//   text's length #VALUE!), case-insensitive match_mode 1, match_end 1
//   (the text's end — its start, from the end — counts as a delimiter) and
//   if_not_found (default #N/A). An empty delimiter matches at once.
// - VALUETOTEXT and ARRAYTOTEXT render a value as a General-formatted cell
//   shows it (concise, format 0); strict (1) quotes text as a formula would.
//   ARRAYTOTEXT's concise form joins with ", "; its strict form is `{…}`
//   with "," between columns and ";" between rows.
import type { CellError } from "../../model";
import { formatGeneral, formatValue, plainText } from "../../numberFormat/render";
import type { SpreadsheetLocale } from "../../numberFormat/locale";
import { forEachScalar, toArray } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { isEmpty, isError, MAX_TEXT_LENGTH, NA_ERROR, toText, VALUE_ERROR, type ScalarValue } from "../values";
import { arg, integerArg, logicalArg, numberArg, optional, scalarArg, textArg } from "./common";
import { roundDecimal } from "./math";

/** Windows-1252's 0x80–0x9F (WHATWG index); every other byte is its code
 * point. */
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f, 0x90, 0x2018,
  0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

const BYTE_OF: ReadonlyMap<number, number> = new Map(CP1252_HIGH.map((codePoint, i) => [codePoint, 0x80 + i]));

/** The character of a Windows-1252 byte. */
function cp1252Char(byte: number): string {
  return String.fromCodePoint(byte >= 0x80 && byte <= 0x9f ? CP1252_HIGH[byte - 0x80]! : byte);
}

/** The Windows-1252 byte of a code point, 63 ("?") when it has none. */
function cp1252Byte(codePoint: number): number {
  if (codePoint < 0x80 || (codePoint >= 0xa0 && codePoint <= 0xff)) {
    return codePoint;
  }
  return BYTE_OF.get(codePoint) ?? 63;
}

/** A text result, #VALUE! past MAX_TEXT_LENGTH. */
function fitted(text: string): string | CellError {
  return text.length > MAX_TEXT_LENGTH ? VALUE_ERROR : text;
}

/** A `#,##0.00`-shaped code with `decimals` places (none at or below 0). */
function placesCode(grouped: boolean, decimals: number): string {
  return (grouped ? "#,##0" : "0") + (decimals > 0 ? `.${"0".repeat(decimals)}` : "");
}

/** DOLLAR's format code for a locale (see the module header). */
function dollarCode(locale: SpreadsheetLocale, decimals: number): string {
  const body = placesCode(true, decimals);
  switch (locale.tag) {
    case "en-US":
      return `"$"${body};("$"${body})`;
    case "en-GB":
      return `"£"${body};-"£"${body}`;
    case "cs-CZ":
      return `${body}" Kč";-${body}" Kč"`;
    default:
      return `${body}" €";-${body}" €"`;
  }
}

/** A number rounded half away from zero to `decimals` and rendered through
 * a format code in the locale — FIXED and DOLLAR. */
function formattedNumber(args: ArgList, ctx: EvalContext, code: (decimals: number) => string): string | CellError {
  const x = numberArg(args, 0, ctx);
  if (isError(x)) {
    return x;
  }
  const decimals = args.length > 1 && !args.isMissing(1) ? integerArg(args, 1, ctx) : 2;
  if (isError(decimals)) {
    return decimals;
  }
  if (decimals > 127) {
    return VALUE_ERROR;
  }
  const rounded = roundDecimal(x, decimals, "half");
  return plainText(formatValue(rounded, code(decimals), { locale: ctx.settings.locale, dateSystem: ctx.settings.dateSystem }));
}

/** VALUETOTEXT's rendering of one value: General (concise) or a formula's
 * spelling of it (strict). */
function valueText(value: ScalarValue, strict: boolean, ctx: EvalContext): string {
  if (isEmpty(value)) {
    return "";
  }
  if (typeof value === "string") {
    return strict ? `"${value.replaceAll('"', '""')}"` : value;
  }
  if (isError(value)) {
    return value.error;
  }
  return formatGeneral(value, { locale: ctx.settings.locale, dateSystem: ctx.settings.dateSystem }).text;
}

/** A format argument of VALUETOTEXT/ARRAYTOTEXT: 0 concise, 1 strict. */
function strictArg(args: ArgList, index: number, ctx: EvalContext): boolean | CellError {
  if (args.length <= index || args.isMissing(index)) {
    return false;
  }
  const given = scalarArg(args, index);
  if (isError(given)) {
    return given;
  }
  const n = integerArg(args, index, ctx);
  if (isError(n)) {
    return n;
  }
  return n === 0 ? false : n === 1 ? true : VALUE_ERROR;
}

/** One delimiter occurrence: where it starts and ends. */
interface Occurrence {
  start: number;
  end: number;
}

/** Every non-overlapping occurrence of any delimiter, left to right (the
 * longest delimiter where several start together). */
function occurrences(text: string, delimiters: readonly string[], caseInsensitive: boolean): Occurrence[] {
  const haystack = caseInsensitive ? text.toLowerCase() : text;
  const needles = delimiters.map((delimiter) => (caseInsensitive ? delimiter.toLowerCase() : delimiter)).filter((d) => d.length > 0);
  const found: Occurrence[] = [];
  let at = 0;
  while (at < haystack.length) {
    let best: Occurrence | null = null;
    for (const needle of needles) {
      const start = haystack.indexOf(needle, at);
      if (start >= 0 && (best === null || start < best.start || (start === best.start && start + needle.length > best.end))) {
        best = { start, end: start + needle.length };
      }
    }
    if (best === null) {
      break;
    }
    found.push(best);
    at = best.end;
  }
  return found;
}

/** TEXTBEFORE (`before`) or TEXTAFTER over its six arguments. */
function textAround(args: ArgList, ctx: EvalContext, before: boolean): EvalResult {
  const text = textArg(args, 0);
  if (isError(text)) {
    return text;
  }
  const delimiters: string[] = [];
  let failure: CellError | undefined;
  forEachScalar(ctx, args.get(1), (value) => {
    const piece = toText(value);
    if (isError(piece)) {
      failure = piece;
      return false;
    }
    delimiters.push(piece);
    return true;
  });
  if (failure) {
    return failure;
  }
  const options: number[] = [];
  for (const [index, fallback] of [
    [2, 1],
    [3, 0],
    [4, 0],
  ] as const) {
    const n = args.length > index && !args.isMissing(index) ? integerArg(args, index, ctx) : fallback;
    if (isError(n)) {
      return n;
    }
    options.push(n);
  }
  const [instance, matchMode, matchEnd] = options as [number, number, number];
  if (instance === 0 || Math.abs(instance) > Math.max(text.length, 1) || (matchMode !== 0 && matchMode !== 1) || (matchEnd !== 0 && matchEnd !== 1)) {
    return VALUE_ERROR;
  }
  const notFound = (): EvalResult => (args.length > 5 && !args.isMissing(5) ? scalarArg(args, 5) : NA_ERROR);
  if (delimiters.length === 0 || delimiters.some((delimiter) => delimiter.length === 0)) {
    // An empty delimiter matches at once: at the start searching forwards,
    // at the end searching backwards.
    const atStart = instance > 0;
    return before ? (atStart ? "" : text) : atStart ? text : "";
  }
  const found = occurrences(text, delimiters, matchMode === 1);
  const count = Math.abs(instance);
  if (instance > 0) {
    const hit = found[count - 1];
    if (hit) {
      return before ? text.slice(0, hit.start) : text.slice(hit.end);
    }
    // match_end: the end of the text is one more delimiter.
    return matchEnd === 1 && found.length === count - 1 ? (before ? text : "") : notFound();
  }
  const hit = found[found.length - count];
  if (hit && found.length - count >= 0) {
    return before ? text.slice(0, hit.start) : text.slice(hit.end);
  }
  // match_end from the end: the start of the text is one more delimiter.
  return matchEnd === 1 && found.length === count - 1 ? (before ? "" : text) : notFound();
}

type EvalResult = ScalarValue | CellError;

/** The text of a NUMBERVALUE argument, its separators applied, as a
 * number; #VALUE! when it is none. */
function numberValue(raw: string, decimal: string, group: string): number | CellError {
  let text = raw.replace(/\s+/g, "");
  if (text.length === 0) {
    return 0;
  }
  let percents = 0;
  while (text.endsWith("%")) {
    percents += 1;
    text = text.slice(0, -1);
  }
  const point = text.indexOf(decimal);
  if (point >= 0 && text.indexOf(decimal, point + 1) >= 0) {
    return VALUE_ERROR;
  }
  const whole = point >= 0 ? text.slice(0, point) : text;
  const fraction = point >= 0 ? text.slice(point + 1) : "";
  if (group.length > 0 && fraction.includes(group)) {
    return VALUE_ERROR;
  }
  const digits = `${group.length > 0 ? whole.split(group).join("") : whole}${point >= 0 ? `.${fraction}` : ""}`;
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(digits)) {
    return VALUE_ERROR;
  }
  const value = Number(digits) / 100 ** percents;
  return Number.isFinite(value) ? value : VALUE_ERROR;
}

/** A separator argument: its first character, the fallback when omitted. */
function separatorArg(args: ArgList, index: number, fallback: string): string | CellError {
  if (args.length <= index || args.isMissing(index)) {
    return fallback;
  }
  const text = textArg(args, index);
  if (isError(text)) {
    return text;
  }
  return text.length === 0 ? VALUE_ERROR : Array.from(text)[0]!;
}

const TEXT_ARG = arg("text", "The text to search.");

export const TEXT_EXTRA_FUNCTIONS: FunctionSpec[] = [
  {
    name: "CHAR",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns the character with a Windows-1252 code from 1 to 255.",
    args: [arg("number", "The character code, from 1 to 255.")],
    impl(args, ctx) {
      const n = integerArg(args, 0, ctx);
      if (isError(n)) {
        return n;
      }
      return n < 1 || n > 255 ? VALUE_ERROR : cp1252Char(n);
    },
  },
  {
    name: "CODE",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns the Windows-1252 code of the first character of a text.",
    args: [arg("text", "The text whose first character to read.")],
    impl(args) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      return text.length === 0 ? VALUE_ERROR : cp1252Byte(text.codePointAt(0)!);
    },
  },
  {
    name: "UNICHAR",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns the character with a Unicode code point.",
    args: [arg("number", "The code point, from 1 to 1114111.")],
    impl(args, ctx) {
      const n = integerArg(args, 0, ctx);
      if (isError(n)) {
        return n;
      }
      if (n < 1 || n > 0x10ffff) {
        return VALUE_ERROR;
      }
      return n >= 0xd800 && n <= 0xdfff ? NA_ERROR : String.fromCodePoint(n);
    },
  },
  {
    name: "UNICODE",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns the Unicode code point of the first character of a text.",
    args: [arg("text", "The text whose first character to read.")],
    impl(args) {
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      return text.length === 0 ? VALUE_ERROR : text.codePointAt(0)!;
    },
  },
  {
    name: "CLEAN",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Removes the control characters U+0000 to U+001F from a text.",
    args: [arg("text", "The text to clean.")],
    impl(args) {
      const text = textArg(args, 0);
      // eslint-disable-next-line no-control-regex
      return isError(text) ? text : text.replace(/[\u0000-\u001f]/g, "");
    },
  },
  {
    name: "FIXED",
    category: "Text",
    minArgs: 1,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Rounds a number and returns it as text, with or without digit grouping.",
    args: [
      arg("number", "The number to round and convert."),
      optional("decimals", "The decimal places (default 2); negative rounds left of the point."),
      optional("no_commas", "TRUE leaves out the thousands separators."),
    ],
    impl(args, ctx) {
      const noCommas = args.length > 2 && !args.isMissing(2) ? logicalArg(args, 2) : false;
      if (isError(noCommas)) {
        return noCommas;
      }
      return formattedNumber(args, ctx, (decimals) => placesCode(!noCommas, decimals));
    },
  },
  {
    name: "DOLLAR",
    category: "Text",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Rounds a number and returns it as currency text in the locale's format.",
    args: [arg("number", "The number to convert."), optional("decimals", "The decimal places (default 2); negative rounds left of the point.")],
    impl(args, ctx) {
      return formattedNumber(args, ctx, (decimals) => dollarCode(ctx.settings.locale, decimals));
    },
  },
  {
    name: "NUMBERVALUE",
    category: "Text",
    minArgs: 1,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Converts text to a number with the given decimal and group separators.",
    args: [
      arg("text", "The text to convert."),
      optional("decimal_separator", "The decimal separator (default the locale's)."),
      optional("group_separator", "The digit group separator (default the locale's)."),
    ],
    impl(args, ctx) {
      const value = scalarArg(args, 0);
      if (typeof value === "number") {
        return value;
      }
      const text = textArg(args, 0);
      if (isError(text)) {
        return text;
      }
      const decimal = separatorArg(args, 1, ctx.settings.locale.decimal);
      if (isError(decimal)) {
        return decimal;
      }
      const group = separatorArg(args, 2, ctx.settings.locale.group);
      if (isError(group)) {
        return group;
      }
      return decimal === group ? VALUE_ERROR : numberValue(text, decimal, group);
    },
  },
  {
    name: "T",
    category: "Text",
    minArgs: 1,
    maxArgs: 1,
    argKinds: ["value"],
    summary: "Returns a value when it is text, and empty text otherwise.",
    args: [arg("value", "The value to test.")],
    impl(args) {
      const value = scalarArg(args, 0);
      return typeof value === "string" ? value : "";
    },
  },
  {
    name: "TEXTBEFORE",
    category: "Text",
    minArgs: 2,
    maxArgs: 6,
    argKinds: ["value", "any", "value", "value", "value", "value"],
    summary: "Returns the text that comes before a delimiter.",
    args: [
      TEXT_ARG,
      arg("delimiter", "The text marking the point, or an array of such texts."),
      optional("instance_num", "Which occurrence (default 1); negative counts from the end."),
      optional("match_mode", "0 matches case (default); 1 ignores case."),
      optional("match_end", "1 treats the end of the text as a delimiter; 0 does not (default)."),
      optional("if_not_found", "What to return when there is no match (default #N/A)."),
    ],
    impl(args, ctx) {
      return textAround(args, ctx, true);
    },
  },
  {
    name: "TEXTAFTER",
    category: "Text",
    minArgs: 2,
    maxArgs: 6,
    argKinds: ["value", "any", "value", "value", "value", "value"],
    summary: "Returns the text that comes after a delimiter.",
    args: [
      TEXT_ARG,
      arg("delimiter", "The text marking the point, or an array of such texts."),
      optional("instance_num", "Which occurrence (default 1); negative counts from the end."),
      optional("match_mode", "0 matches case (default); 1 ignores case."),
      optional("match_end", "1 treats the end of the text as a delimiter; 0 does not (default)."),
      optional("if_not_found", "What to return when there is no match (default #N/A)."),
    ],
    impl(args, ctx) {
      return textAround(args, ctx, false);
    },
  },
  {
    name: "VALUETOTEXT",
    category: "Text",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["value", "value"],
    acceptsErrors: true,
    summary: "Returns any value as text.",
    args: [arg("value", "The value to convert."), optional("format", "0 concise, as a cell shows it (default); 1 strict, text quoted.")],
    impl(args, ctx) {
      const strict = strictArg(args, 1, ctx);
      if (isError(strict)) {
        return strict;
      }
      return fitted(valueText(scalarArg(args, 0), strict, ctx));
    },
  },
  {
    name: "ARRAYTOTEXT",
    category: "Text",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["any", "value"],
    acceptsErrors: true,
    summary: "Returns an array or range of values as one text.",
    args: [arg("array", "The array or range to convert."), optional("format", "0 concise, joined with commas (default); 1 strict, as an array constant.")],
    impl(args, ctx) {
      const strict = strictArg(args, 1, ctx);
      if (isError(strict)) {
        return strict;
      }
      const array = toArray(ctx, args.get(0));
      if (isError(array)) {
        return array;
      }
      const texts = array.data.map((value) => valueText(value, strict, ctx));
      if (!strict) {
        return fitted(texts.join(", "));
      }
      const rows: string[] = [];
      for (let r = 0; r < array.rows; r += 1) {
        rows.push(texts.slice(r * array.cols, (r + 1) * array.cols).join(","));
      }
      return fitted(`{${rows.join(";")}}`);
    },
  },
];
