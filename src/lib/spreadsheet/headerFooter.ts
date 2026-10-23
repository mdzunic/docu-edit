// Header and footer code text (SPEC-136 "Header and footer codes") — THIS
// FILE is the normative home of the code language a sheet's headers and
// footers are stored in (model.ts HeaderFooterSections): Excel's own, so an
// .xlsx header is read and written as it stands and an .ods one converts to
// it (ods/pageSetup.ts, through ODS_FIELD_CODES).
//
// - parseHeaderFooter splits a header on `&L`/`&C`/`&R` (text before the
//   first section code is the centre); formatHeaderFooter is the inverse for
//   every string in Excel's own form (`&L…&C…&R…`, each section once, in that
//   order) — `formatHeaderFooter(parseHeaderFooter(t)) === t`.
// - tokenizeHeaderFooter reads `&&`, `&P` (`&P+n`, `&P-n`), `&N`, `&D`, `&T`,
//   `&F`, `&Z`, `&A`, `&G`, `&"family,style"`, a run of digits (the font size
//   in pt), `&K` + six hex digits or a theme form (`&K01+000`), `&B`, `&I`,
//   `&U`, `&E`, `&S`, `&X`, `&Y`, `&O`, `&H`; any other `&` + character — and a
//   trailing `&` — is literal text, kept as written.
// - renderHeaderFooterSection turns a section into styled runs for one
//   page: the fields filled in, `&D`/`&T` the locale's short date and time
//   (SPEC-118), `&Z` an empty run, `&G` nothing (its picture stays in the
//   retained `legacyDrawingHF` part, never rendered), `&X`/`&Y`/`&O`/`&H` and
//   a theme `&K` no effect (plain text in the default colour). Toggles flip
//   for the rest of the section, as in Excel.
import { dateToSerial } from "./serial";
import type { HeaderFooterSections } from "./model";
import type { SpreadsheetLocale } from "./numberFormat/locale";
import { formatValue } from "./numberFormat/render";

export type HfToggle = "bold" | "italic" | "underline" | "doubleUnderline" | "strike" | "superscript" | "subscript" | "outline" | "shadow";

export type HfToken =
  | { kind: "text"; text: string }
  | { kind: "page"; offset: number }
  | { kind: "pages" }
  | { kind: "date" }
  | { kind: "time" }
  | { kind: "file" }
  | { kind: "path" }
  | { kind: "sheet" }
  | { kind: "picture" }
  | { kind: "font"; family: string; style: string }
  | { kind: "size"; pt: number }
  | { kind: "color"; rgb: string }
  | { kind: "themeColor"; code: string }
  | { kind: "toggle"; toggle: HfToggle };

/** One styled run of a rendered section. */
export interface HfRun {
  text: string;
  bold?: true;
  italic?: true;
  underline?: "single" | "double";
  strike?: true;
  /** Absent: the default font. */
  fontFamily?: string;
  fontSizePt?: number;
  /** "#rrggbb"; absent: the default colour. */
  color?: string;
}

export interface HeaderFooterContext {
  page: number;
  pages: number;
  date: Date;
  sheetName: string;
  fileName: string;
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
}

const SECTION_CODES: Readonly<Record<string, keyof HeaderFooterSections>> = { L: "left", C: "center", R: "right" };
const SECTION_LETTERS: readonly [keyof HeaderFooterSections, string][] = [
  ["left", "L"],
  ["center", "C"],
  ["right", "R"],
];

const TOGGLES: Readonly<Record<string, HfToggle>> = {
  B: "bold",
  I: "italic",
  U: "underline",
  E: "doubleUnderline",
  S: "strike",
  X: "superscript",
  Y: "subscript",
  O: "outline",
  H: "shadow",
};

const SIMPLE_FIELDS: Readonly<Record<string, HfToken>> = {
  N: { kind: "pages" },
  D: { kind: "date" },
  T: { kind: "time" },
  F: { kind: "file" },
  Z: { kind: "path" },
  A: { kind: "sheet" },
  G: { kind: "picture" },
};

/** The ODS text fields a header region maps onto codes (SPEC-136 "ods wire
 * mapping"), element local name → code. */
export const ODS_FIELD_CODES: Readonly<Record<string, string>> = {
  "page-number": "&P",
  "page-count": "&N",
  date: "&D",
  time: "&T",
  "sheet-name": "&A",
  "file-name": "&F",
};

/** Splits a header or footer into its sections (see the module header).
 * A section repeated appends to the first; an empty one is absent. */
export function parseHeaderFooter(text: string): HeaderFooterSections {
  const sections: HeaderFooterSections = {};
  let current: keyof HeaderFooterSections = "center";
  let start = 0;
  let i = 0;
  const flush = (end: number): void => {
    const piece = text.slice(start, end);
    if (piece !== "") {
      sections[current] = (sections[current] ?? "") + piece;
    }
  };
  while (i < text.length) {
    if (text[i] !== "&") {
      i += 1;
      continue;
    }
    const next = text[i + 1];
    const section = next === undefined ? undefined : SECTION_CODES[next];
    if (section !== undefined) {
      flush(i);
      current = section;
      i += 2;
      start = i;
      continue;
    }
    // `&&`, `&"…"` and every other code: skip the pair (a quoted font name
    // may hold an `&L` that is no section code).
    if (next === '"') {
      const close = text.indexOf('"', i + 2);
      i = close < 0 ? text.length : close + 1;
    } else {
      i += next === undefined ? 1 : 2;
    }
  }
  flush(text.length);
  return sections;
}

/** The inverse of parseHeaderFooter: `&L…&C…&R…`, absent sections omitted. */
export function formatHeaderFooter(sections: HeaderFooterSections): string {
  let text = "";
  for (const [key, letter] of SECTION_LETTERS) {
    const value = sections[key];
    if (value !== undefined && value !== "") {
      text += `&${letter}${value}`;
    }
  }
  return text;
}

/** Whether a header or footer has any text. */
export function hasSections(sections: HeaderFooterSections | undefined): boolean {
  return sections !== undefined && SECTION_LETTERS.some(([key]) => (sections[key] ?? "") !== "");
}

const HEX6 = /^[0-9A-Fa-f]{6}/;
const THEME_COLOR = /^\d\d[+-]\d\d\d/;

/** A section's codes (see the module header). Adjacent literal text joins
 * into one token. */
export function tokenizeHeaderFooter(section: string): HfToken[] {
  const tokens: HfToken[] = [];
  const text = (value: string): void => {
    const last = tokens.at(-1);
    if (last?.kind === "text") {
      last.text += value;
    } else if (value !== "") {
      tokens.push({ kind: "text", text: value });
    }
  };
  let i = 0;
  while (i < section.length) {
    const amp = section.indexOf("&", i);
    if (amp < 0) {
      text(section.slice(i));
      break;
    }
    text(section.slice(i, amp));
    const next = section[amp + 1];
    if (next === undefined) {
      text("&");
      break;
    }
    if (next === "&") {
      text("&");
      i = amp + 2;
      continue;
    }
    if (next === "P") {
      const offset = /^([+-])(\d+)/.exec(section.slice(amp + 2));
      tokens.push({ kind: "page", offset: offset ? (offset[1] === "-" ? -1 : 1) * Number(offset[2]) : 0 });
      i = amp + 2 + (offset ? offset[0].length : 0);
      continue;
    }
    const field = SIMPLE_FIELDS[next];
    if (field !== undefined) {
      tokens.push({ ...field });
      i = amp + 2;
      continue;
    }
    const toggle = TOGGLES[next];
    if (toggle !== undefined) {
      tokens.push({ kind: "toggle", toggle });
      i = amp + 2;
      continue;
    }
    if (next === '"') {
      const close = section.indexOf('"', amp + 2);
      if (close < 0) {
        text(section.slice(amp));
        break;
      }
      const spec = section.slice(amp + 2, close);
      const comma = spec.indexOf(",");
      tokens.push({ kind: "font", family: comma < 0 ? spec : spec.slice(0, comma), style: comma < 0 ? "" : spec.slice(comma + 1) });
      i = close + 1;
      continue;
    }
    if (next >= "0" && next <= "9") {
      const digits = /^\d+/.exec(section.slice(amp + 1))![0];
      tokens.push({ kind: "size", pt: Number(digits) });
      i = amp + 1 + digits.length;
      continue;
    }
    if (next === "K") {
      const rest = section.slice(amp + 2);
      const hex = HEX6.exec(rest);
      if (hex) {
        tokens.push({ kind: "color", rgb: `#${hex[0].toLowerCase()}` });
        i = amp + 8;
        continue;
      }
      const theme = THEME_COLOR.exec(rest);
      if (theme) {
        tokens.push({ kind: "themeColor", code: theme[0] });
        i = amp + 2 + theme[0].length;
        continue;
      }
    }
    // Anything else is literal, kept as written.
    text(section.slice(amp, amp + 2));
    i = amp + 2;
  }
  return tokens;
}

/** A Date as a serial (date and time of day) in `system`. */
function serialOf(date: Date, system: "1900" | "1904"): number {
  const day = dateToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate(), system);
  const seconds = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
  return day + seconds / 86_400;
}

interface RunStyle {
  bold: boolean;
  italic: boolean;
  underline: "single" | "double" | undefined;
  strike: boolean;
  fontFamily: string | undefined;
  fontSizePt: number | undefined;
  color: string | undefined;
}

function styledRun(text: string, style: RunStyle): HfRun {
  const run: HfRun = { text };
  if (style.bold) {
    run.bold = true;
  }
  if (style.italic) {
    run.italic = true;
  }
  if (style.underline !== undefined) {
    run.underline = style.underline;
  }
  if (style.strike) {
    run.strike = true;
  }
  if (style.fontFamily !== undefined) {
    run.fontFamily = style.fontFamily;
  }
  if (style.fontSizePt !== undefined) {
    run.fontSizePt = style.fontSizePt;
  }
  if (style.color !== undefined) {
    run.color = style.color;
  }
  return run;
}

/** A font token's effect: the family (`-` keeps the current one) and the
 * style words (Bold, Italic — `Regular` clears both). */
function applyFont(token: Extract<HfToken, { kind: "font" }>, style: RunStyle): void {
  if (token.family !== "" && token.family !== "-") {
    style.fontFamily = token.family;
  }
  const words = token.style.toLowerCase();
  if (words !== "") {
    style.bold = /\bbold\b/.test(words);
    style.italic = /\b(?:italic|oblique)\b/.test(words);
  }
}

/** A section (code text, or its tokens) rendered for one page (see the
 * module header). Adjacent runs of one style join. */
export function renderHeaderFooterSection(section: string | readonly HfToken[], ctx: HeaderFooterContext): HfRun[] {
  const tokens = typeof section === "string" ? tokenizeHeaderFooter(section) : section;
  const style: RunStyle = { bold: false, italic: false, underline: undefined, strike: false, fontFamily: undefined, fontSizePt: undefined, color: undefined };
  const runs: HfRun[] = [];
  const emit = (text: string): void => {
    const run = styledRun(text, style);
    const last = runs.at(-1);
    if (last !== undefined && sameStyle(last, run)) {
      last.text += text;
    } else {
      runs.push(run);
    }
  };
  const formatCtx = { locale: ctx.locale, dateSystem: ctx.dateSystem };
  for (const token of tokens) {
    switch (token.kind) {
      case "text":
        emit(token.text);
        break;
      case "page":
        emit(String(ctx.page + token.offset));
        break;
      case "pages":
        emit(String(ctx.pages));
        break;
      case "date":
        emit(formatValue(Math.floor(serialOf(ctx.date, ctx.dateSystem)), ctx.locale.shortDate, formatCtx).text);
        break;
      case "time":
        emit(formatValue(serialOf(ctx.date, ctx.dateSystem) % 1, ctx.locale.shortTime, formatCtx).text);
        break;
      case "file":
        emit(ctx.fileName);
        break;
      case "path":
        emit("");
        break;
      case "sheet":
        emit(ctx.sheetName);
        break;
      case "picture":
        break;
      case "font":
        applyFont(token, style);
        break;
      case "size":
        style.fontSizePt = token.pt;
        break;
      case "color":
        style.color = token.rgb;
        break;
      case "themeColor":
        style.color = undefined;
        break;
      case "toggle":
        applyToggle(token.toggle, style);
        break;
    }
  }
  return runs;
}

function applyToggle(toggle: HfToggle, style: RunStyle): void {
  switch (toggle) {
    case "bold":
      style.bold = !style.bold;
      break;
    case "italic":
      style.italic = !style.italic;
      break;
    case "underline":
      style.underline = style.underline === "single" ? undefined : "single";
      break;
    case "doubleUnderline":
      style.underline = style.underline === "double" ? undefined : "double";
      break;
    case "strike":
      style.strike = !style.strike;
      break;
    default:
      // &X, &Y, &O, &H: kept, rendered as plain text.
      break;
  }
}

function sameStyle(a: HfRun, b: HfRun): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.fontFamily === b.fontFamily &&
    a.fontSizePt === b.fontSizePt &&
    a.color === b.color
  );
}

/** Literal text as code text: every `&` doubled. */
export function escapeHeaderFooterText(text: string): string {
  return text.replaceAll("&", "&&");
}

/** A section as plain text with its codes named — what a field shows when
 * nothing renders it (never used to write). */
export function headerFooterPlainText(section: string, ctx: HeaderFooterContext): string {
  return renderHeaderFooterSection(section, ctx)
    .map((run) => run.text)
    .join("");
}
