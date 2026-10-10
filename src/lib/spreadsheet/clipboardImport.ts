// External HTML tables on the clipboard (SPEC-125 "External sources") —
// THIS FILE is the normative source for how a table another program put on
// the clipboard becomes a PasteBlock (paste.ts applies it). Whitelist
// import, as SPEC-016's htmlPaste.ts: only what the table below names is
// read; everything else is dropped.
//
// | Producer (detected by) | Value | Number format | Style | Formula |
// | Excel (`xmlns:x="urn:schemas-microsoft-com:office:excel"`) | `x:num="…"` → number (bare `x:num` → parseInput of the text); `x:bool`; `x:err`; `x:str` → string; else parseInput | `mso-number-format` from the cell's class rule or inline style, CSS escapes decoded (`\0022` → `"`, `\X` → `X`), Excel's named formats mapped | `<style>` rules (`td {…}`, `.xl65 {…}`) read as data, plus inline CSS | not read |
// | Google Sheets (`data-sheets-root` / `google-sheets-html-origin`) | `data-sheets-value` JSON: `"1"` the type (2 string, 3 number, 4 boolean), `"2"`/`"3"`/`"4"` the value | `data-sheets-numberformat` `"2"` | inline CSS | `data-sheets-formula` (R1C1, converted to A1 at each destination cell by formula/r1c1.ts; unconvertible → the value) |
// | LibreOffice (`sdval`/`sdnum`) | `sdval` → number (a BOOLEAN code → boolean); else parseInput | `sdnum="lang;index;code"` → `code` when non-empty | inline CSS and `<font>`/`<b>`/`<i>`/`<u>` | not available |
// | Any other `<table>` (Numbers, web pages) | parseInput of the text | from parseInput | inline CSS, `<b>`/`<i>`/`<u>`/`<s>` | not available |
//
// - CSS whitelist in: font-weight (≥ 600 or bold), font-style:italic,
//   text-decoration (underline, line-through), text-underline-style:double,
//   color, background/background-color, font-family (first family, quotes
//   stripped), font-size (pt; px × 0.75), text-align, vertical-align
//   (top/middle/bottom), white-space:normal (Excel's wrap), mso-rotate, and
//   border* (.5pt → thin, 1.0pt → medium, ≥ 1.5pt → thick; 1px/2px/3px
//   likewise; double, dotted, dashed; hairline → hair). Colours through
//   core/colors.ts parseColorInput plus the CSS basic names; windowtext is
//   automatic. A value that is not one of these shapes — a url(), an
//   expression — is never interpreted, only ignored.
// - The text is the cell's text with `<br>` as \n, whitespace collapsed as
//   a browser shows it; script, style and template content is no text.
// - colspan/rowspan become merges, laid out by HTML's table algorithm.
// - Parsing is DOMParser("text/html") on a detached document, so nothing in
//   the markup executes or loads. Input over MAX_IMPORT_HTML_LENGTH, or with
//   more than MAX_COPY_CELLS cells, is refused before parsing; a table whose
//   area (spans included) passes MAX_COPY_CELLS is refused as it is laid out.
// - What attributes give is held to the payload's bounds: text XML-legal and
//   ≤ 32,767 characters, format codes ≤ 255 printable characters, a Sheets
//   formula ≤ 8,192 characters (else its value pastes).
// - Numbers from attributes (x:num, sdval, data-sheets-value) are serials of
//   the 1900 date system; a date-formatted one is moved by 1462 days into a
//   1904 workbook here, so the block is in the destination's date system.
import { parseColorInput } from "../core/colors";
import { hasIllegalXmlScalars, sanitizeXmlText } from "../core/xmlText";
import type { GridRange } from "./address";
import { TOO_MANY_TO_CHANGE_MESSAGE } from "./clipboard";
import { isPrintable, MAX_CLIPBOARD_FORMULA_LENGTH, MAX_COPY_CELLS } from "./clipboardPayload";
import { MAX_TEXT_LENGTH } from "./formula/values";
import { MAX_COLS, MAX_ROWS, toCellError, type CellBorderEdge, type CellBorders, type CellStyle, type CellValue, type SheetBorderStyle } from "./model";
import { BUILTIN_NUMBER_FORMATS } from "./numberFormat/builtin";
import { parseFormatCode } from "./numberFormat/parse";
import { parsedField, type PasteBlock, type PasteCell, type PasteSource, type TextContext } from "./paste";
import { dateKindOf, DATE_SYSTEM_OFFSET } from "./series";
import { styleKey } from "./styles";

/** 64 MiB of HTML. */
export const MAX_IMPORT_HTML_LENGTH = 64 * 1024 * 1024;
const MAX_COLSPAN = 1000;
const MAX_ROWSPAN = 65534;
/** The payload's bound on a number-format code (clipboardPayload.ts). */
const MAX_FORMAT_CODE_LENGTH = 255;

/** Text an attribute gives a cell, under the bounds typed and parsed text
 * keeps (paste.ts parsedField): XML-legal, at most 32,767 characters. */
function attributeText(text: string): string {
  return sanitizeXmlText(text).slice(0, MAX_TEXT_LENGTH);
}

/** parseInput's locale and date system, and the view's canonical formula
 * form (paste.ts TextContext). */
export type ImportContext = TextContext;

type Declarations = Map<string, string>;

// --- CSS as data -------------------------------------------------------------------------

/** `a:b; c:"d;e"` → { a: "b", c: "\"d;e\"" }: property names lower-cased,
 * values trimmed, `!important` dropped; quotes and parentheses respected. */
export function cssDeclarations(text: string): Declarations {
  const out: Declarations = new Map();
  let start = 0;
  let quote = "";
  let depth = 0;
  const flush = (end: number): void => {
    const part = text.slice(start, end);
    const colon = part.indexOf(":");
    if (colon > 0) {
      const name = part.slice(0, colon).trim().toLowerCase();
      const value = part
        .slice(colon + 1)
        .replace(/!\s*important\s*$/i, "")
        .trim();
      if (name && value) {
        out.set(name, value);
      }
    }
    start = end + 1;
  };
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]!;
    if (quote) {
      if (ch === "\\") {
        i += 1;
      } else if (ch === quote) {
        quote = "";
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === "(") {
      depth += 1;
    } else if (ch === ")") {
      depth = Math.max(0, depth - 1);
    } else if (ch === ";" && depth === 0) {
      flush(i);
    }
  }
  flush(text.length);
  return out;
}

/** A `<style>` sheet's `td {…}` and `.class {…}` rules, read as data: a
 * rule is the text since the last `{`, `}` or `@` before a `{`, and a body
 * that reaches its `}` with no `{` (an at-rule's own block is no rule).
 * Found in one linear pass — a regex over hostile, brace-less text
 * backtracks quadratically. */
export function cssRules(text: string): { base: Declarations; classes: Map<string, Declarations> } {
  const base: Declarations = new Map();
  const classes = new Map<string, Declarations>();
  const clean = text.replace(/<!--|-->/g, " ").replace(/\/\*[\s\S]*?\*\//g, " ");
  const rules: [selectors: string, body: string][] = [];
  let selectorStart = 0;
  for (let i = 0; i < clean.length; i += 1) {
    const ch = clean[i];
    if (ch === "@" || ch === "}") {
      selectorStart = i + 1;
    } else if (ch === "{") {
      let end = i + 1;
      while (end < clean.length && clean[end] !== "{" && clean[end] !== "}") {
        end += 1;
      }
      if (clean[end] === "}" && i > selectorStart) {
        rules.push([clean.slice(selectorStart, i), clean.slice(i + 1, end)]);
        i = end;
      }
      selectorStart = i + 1;
    }
  }
  for (const [selectors, body] of rules) {
    const declarations = cssDeclarations(body);
    for (const selector of selectors.split(",")) {
      const name = selector.trim().toLowerCase();
      const target = name === "td" ? base : /^\.[\w-]+$/.test(name) ? classes.get(name.slice(1)) ?? new Map<string, string>() : null;
      if (!target) {
        continue;
      }
      for (const [key, value] of declarations) {
        target.set(key, value);
      }
      if (name !== "td") {
        classes.set(name.slice(1), target);
      }
    }
  }
  return { base, classes };
}

/** A CSS string's text: outer quotes removed, escapes decoded as Excel
 * writes them (`\0022` → `"`, `\X` → `X`). Excel's hex escape is always
 * exactly four digits with no terminating space — not CSS's one to six plus
 * an optional space — so `\0022Dec\0022` is `"Dec"` and `0\0022 kg\0022`
 * keeps its space. */
export function decodeCssString(value: string): string {
  let text = value.trim();
  if (text.length >= 2 && (text[0] === '"' || text[0] === "'") && text.at(-1) === text[0]) {
    text = text.slice(1, -1);
  }
  return text.replace(/\\([0-9a-fA-F]{4})|\\([\s\S])/g, (_, hex: string | undefined, char: string | undefined) =>
    hex !== undefined ? String.fromCharCode(Number.parseInt(hex, 16)) : (char ?? ""),
  );
}

const CSS_BASIC_COLORS: Readonly<Record<string, string>> = {
  black: "#000000",
  silver: "#c0c0c0",
  gray: "#808080",
  grey: "#808080",
  white: "#ffffff",
  maroon: "#800000",
  red: "#ff0000",
  purple: "#800080",
  fuchsia: "#ff00ff",
  green: "#008000",
  lime: "#00ff00",
  olive: "#808000",
  yellow: "#ffff00",
  navy: "#000080",
  blue: "#0000ff",
  teal: "#008080",
  aqua: "#00ffff",
};

/** A CSS colour as `#rrggbb`; undefined for automatic (windowtext) or
 * anything that is not a colour. */
export function cssColor(value: string): string | undefined {
  const text = value.trim().toLowerCase();
  return CSS_BASIC_COLORS[text] ?? parseColorInput(text);
}

/** A CSS number with an optional pt/px unit. Its digit runs never overlap
 * (unlike `\d*\.?\d+`, which backtracks quadratically over a long run of
 * digits that does not match). */
const CSS_LENGTH = /^(\d+(?:\.\d+)?|\.\d+)(pt|px)?$/i;

/** A CSS length in points: `pt`, `px` (× 0.75), or a bare number of pt. */
function lengthPt(value: string): number | undefined {
  const match = CSS_LENGTH.exec(value.trim());
  if (!match) {
    return undefined;
  }
  const amount = Number(match[1]);
  return match[2]?.toLowerCase() === "px" ? amount * 0.75 : amount;
}

const BORDER_KEYWORDS: Readonly<Record<string, SheetBorderStyle>> = { double: "double", dotted: "dotted", dashed: "dashed", hairline: "hair" };

/** One `border-*` value as an edge; null for none, undefined when it is
 * not one. */
function cssEdge(value: string): CellBorderEdge | null | undefined {
  const tokens = value.trim().toLowerCase().split(/\s+/);
  if (tokens.includes("none") || tokens.includes("hidden")) {
    return null;
  }
  let style: SheetBorderStyle | undefined;
  let width: number | undefined;
  let pixels = false;
  let color: string | undefined;
  for (const token of tokens) {
    const keyword = BORDER_KEYWORDS[token];
    if (keyword) {
      style = keyword;
    } else if (token === "solid") {
      style ??= "thin";
    } else if (CSS_LENGTH.test(token)) {
      width = lengthPt(token);
      pixels = token.endsWith("px");
    } else if (token === "thin" || token === "medium" || token === "thick") {
      width = token === "thin" ? 0.5 : token === "medium" ? 1 : 1.5;
    } else {
      color = cssColor(token) ?? color;
    }
  }
  if (style === undefined) {
    return undefined;
  }
  if (style === "thin" && width !== undefined) {
    const px = pixels ? width / 0.75 : undefined;
    style = px !== undefined ? (px < 1.5 ? "thin" : px < 2.5 ? "medium" : "thick") : width < 0.9 ? "thin" : width < 1.5 ? "medium" : "thick";
  }
  return color === undefined ? { style } : { style, color };
}

/** Tag-given formatting (LibreOffice's `<font>`, a web table's `<b>`). */
interface TagFormat {
  bold?: true;
  italic?: true;
  underline?: true;
  strike?: true;
  family?: string;
  sizePt?: number;
  color?: string;
}

const FONT_SIZES_PT = [0, 7.5, 10, 12, 13.5, 18, 24, 36];
/** A font family: printable, and nothing that could leave a CSS value. */
function cssFamily(value: string): string | undefined {
  const first = decodeCssString(value.split(",")[0] ?? "").trim();
  return first.length <= 255 && isPrintable(first) && !/[()\\<>{};]/.test(first) ? first : undefined;
}

/** The whitelisted CSS and tag formats as a cell style (see the header). */
export function importedStyle(css: Declarations, tags: TagFormat = {}): CellStyle {
  const style: CellStyle = {};
  const font: NonNullable<CellStyle["font"]> = {};
  const weight = css.get("font-weight")?.toLowerCase();
  if (tags.bold || weight === "bold" || weight === "bolder" || (weight !== undefined && Number(weight) >= 600)) {
    font.bold = true;
  }
  if (tags.italic || /^(italic|oblique)$/i.test(css.get("font-style") ?? "")) {
    font.italic = true;
  }
  const decoration = (css.get("text-decoration") ?? css.get("text-decoration-line") ?? "").toLowerCase();
  if (tags.underline || decoration.includes("underline")) {
    font.underline = css.get("text-underline-style")?.toLowerCase() === "double" ? "double" : "single";
  }
  if (tags.strike || decoration.includes("line-through")) {
    font.strike = true;
  }
  const color = css.has("color") ? cssColor(css.get("color")!) : tags.color;
  if (color !== undefined) {
    font.color = color;
  }
  const family = css.has("font-family") ? cssFamily(css.get("font-family")!) : tags.family;
  if (family !== undefined) {
    font.family = family;
  }
  const size = css.has("font-size") ? lengthPt(css.get("font-size")!) : tags.sizePt;
  if (size !== undefined && size >= 1 && size <= 409) {
    font.sizePt = Math.round(size * 2) / 2;
  }
  if (Object.keys(font).length > 0) {
    style.font = font;
  }
  const background = css.get("background-color") ?? css.get("background")?.split(/\s+/)[0];
  const fill = background === undefined ? undefined : cssColor(background);
  if (fill !== undefined) {
    style.fill = { kind: "solid", color: fill };
  }
  const alignment: NonNullable<CellStyle["alignment"]> = {};
  const horizontal = css.get("text-align")?.toLowerCase();
  if (horizontal === "left" || horizontal === "center" || horizontal === "right" || horizontal === "justify") {
    alignment.horizontal = horizontal;
  }
  const vertical = css.get("vertical-align")?.toLowerCase();
  if (vertical === "top" || vertical === "middle") {
    alignment.vertical = vertical === "middle" ? "center" : "top";
  }
  if (css.get("white-space")?.toLowerCase() === "normal") {
    alignment.wrap = true;
  }
  const rotate = css.get("mso-rotate")?.toLowerCase();
  if (rotate === "vertical") {
    alignment.textRotation = 255;
  } else if (rotate !== undefined && /^-?\d+$/.test(rotate) && Math.abs(Number(rotate)) <= 90 && Number(rotate) !== 0) {
    alignment.textRotation = Number(rotate) > 0 ? Number(rotate) : 90 - Number(rotate);
  }
  if (Object.keys(alignment).length > 0) {
    style.alignment = alignment;
  }
  const borders: CellBorders = {};
  const all = css.has("border") ? cssEdge(css.get("border")!) : undefined;
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const own = css.get(`border-${side}`);
    const edge = own !== undefined ? cssEdge(own) : all;
    if (edge) {
      borders[side] = edge;
    }
  }
  if (Object.keys(borders).length > 0) {
    style.borders = borders;
  }
  return style;
}

// --- number formats -------------------------------------------------------------------------

/** Excel's named mso-number-formats. */
const EXCEL_NAMED_FORMATS: Readonly<Record<string, { numberFormat: string; builtinFormatId?: number } | null>> = {
  "general": null,
  "general number": null,
  "standard": { numberFormat: "#,##0.00" },
  "fixed": { numberFormat: "0.00" },
  "percent": { numberFormat: "0.00%" },
  "scientific": { numberFormat: "0.00E+00" },
  "currency": { numberFormat: '"$"#,##0.00' },
  "short date": { numberFormat: BUILTIN_NUMBER_FORMATS[14]!, builtinFormatId: 14 },
  "general date": { numberFormat: BUILTIN_NUMBER_FORMATS[22]!, builtinFormatId: 22 },
  "medium date": { numberFormat: "d-mmm-yy" },
  "long date": { numberFormat: "dddd, mmmm d, yyyy" },
  "short time": { numberFormat: "h:mm" },
  "medium time": { numberFormat: "h:mm AM/PM" },
  "long time": { numberFormat: "h:mm:ss AM/PM" },
};

type Format = { numberFormat: string; builtinFormatId?: number };

/** A producer's format code as a style's format; undefined for General or
 * a code the payload's bounds (≤ 255 printable characters) or SPEC-118's
 * grammar reject. */
function formatOf(code: string | undefined): Format | undefined {
  if (code === undefined || code.trim() === "") {
    return undefined;
  }
  const named = EXCEL_NAMED_FORMATS[code.trim().toLowerCase()];
  if (named !== undefined) {
    return named ?? undefined;
  }
  if (code.length > MAX_FORMAT_CODE_LENGTH || !isPrintable(code) || hasIllegalXmlScalars(code)) {
    return undefined;
  }
  try {
    parseFormatCode(code);
  } catch {
    return undefined;
  }
  return { numberFormat: code };
}

// --- the table -----------------------------------------------------------------------------

const SKIPPED_TEXT = new Set(["SCRIPT", "STYLE", "TEMPLATE", "NOSCRIPT", "TITLE", "HEAD"]);

/** A cell's text: `<br>` as \n, other whitespace collapsed and trimmed per
 * line as a browser renders it; no script or style content. */
export function cellText(element: Element): string {
  let raw = "";
  const walk = (node: Node): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) {
        raw += (child as Text).data.replace(/[ \t\n\r\f]+/g, " ");
      } else if (child.nodeType === 1) {
        const tag = (child as Element).tagName.toUpperCase();
        if (tag === "BR") {
          raw += "\n";
        } else if (!SKIPPED_TEXT.has(tag)) {
          walk(child);
        }
      }
    }
  };
  walk(element);
  return raw
    .split("\n")
    .map((line) => line.replace(/ +/g, " ").trim())
    .join("\n")
    .replace(/^\n+|\n+$/g, "")
    .replaceAll(" ", " ");
}

/** Whether every piece of text in `cell` sits inside one of `tags`. */
function allTextWithin(cell: Element, tags: ReadonlySet<string>): boolean {
  let found = false;
  let outside = false;
  const walk = (node: Node, inside: boolean): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3) {
        if ((child as Text).data.trim() !== "") {
          found = true;
          outside ||= !inside;
        }
      } else if (child.nodeType === 1 && !SKIPPED_TEXT.has((child as Element).tagName.toUpperCase())) {
        walk(child, inside || tags.has((child as Element).tagName.toUpperCase()));
      }
    }
  };
  walk(cell, false);
  return found && !outside;
}

const BOLD_TAGS = new Set(["B", "STRONG"]);
const ITALIC_TAGS = new Set(["I", "EM"]);
const UNDERLINE_TAGS = new Set(["U"]);
const STRIKE_TAGS = new Set(["S", "STRIKE", "DEL"]);

/** Formatting the cell's tags give all of its text. */
function tagFormat(cell: Element, withFont: boolean): TagFormat {
  const format: TagFormat = {};
  if (allTextWithin(cell, BOLD_TAGS)) {
    format.bold = true;
  }
  if (allTextWithin(cell, ITALIC_TAGS)) {
    format.italic = true;
  }
  if (allTextWithin(cell, UNDERLINE_TAGS)) {
    format.underline = true;
  }
  if (allTextWithin(cell, STRIKE_TAGS)) {
    format.strike = true;
  }
  const font = withFont ? cell.querySelector("font") : null;
  if (font && allTextWithin(cell, new Set(["FONT"]))) {
    const face = font.getAttribute("face");
    const family = face ? cssFamily(face) : undefined;
    if (family) {
      format.family = family;
    }
    const size = Number(font.getAttribute("size"));
    if (Number.isInteger(size) && size >= 1 && size <= 7) {
      format.sizePt = FONT_SIZES_PT[size];
    }
    const color = font.getAttribute("color");
    const hex = color ? cssColor(color) : undefined;
    if (hex) {
      format.color = hex;
    }
  }
  return format;
}

function detectSource(html: string, doc: Document): PasteSource {
  if (html.includes("urn:schemas-microsoft-com:office:excel")) {
    return "excel";
  }
  if (doc.querySelector("google-sheets-html-origin, [data-sheets-root]")) {
    return "sheets";
  }
  if (doc.querySelector("td[sdval], td[sdnum], th[sdval], th[sdnum]")) {
    return "libreoffice";
  }
  return "html";
}

interface Placed {
  element: HTMLTableCellElement;
  row: number;
  col: number;
  rowspan: number;
  colspan: number;
}

/** The table's cells placed by HTML's table algorithm (spans occupying the
 * positions they cover); null past MAX_COPY_CELLS. The table's area is
 * checked as each cell is placed, before its span is marked, so the
 * positions marked never outnumber MAX_COPY_CELLS — a hostile
 * `rowspan=20000 colspan=1000` is refused, never walked. */
function placeCells(table: HTMLTableElement): { cells: Placed[]; rows: number; cols: number } | null {
  const taken = new Set<number>();
  const cells: Placed[] = [];
  let rows = 0;
  let cols = 0;
  let row = 0;
  for (const tr of Array.from(table.rows)) {
    let col = 0;
    for (const element of Array.from(tr.cells)) {
      while (taken.has(row * MAX_COLS + col)) {
        col += 1;
      }
      const colspan = Math.min(Math.max(1, element.colSpan || 1), MAX_COLSPAN);
      const rowspan = Math.min(Math.max(1, element.rowSpan || 1), MAX_ROWSPAN);
      const bottom = row + rowspan;
      const right = col + colspan;
      if (bottom > MAX_ROWS || right > MAX_COLS || cells.length >= MAX_COPY_CELLS || Math.max(rows, bottom) * Math.max(cols, right) > MAX_COPY_CELLS) {
        return null;
      }
      cells.push({ element, row, col, rowspan, colspan });
      if (rowspan > 1 || colspan > 1) {
        for (let r = row; r < bottom; r += 1) {
          for (let c = col; c < right; c += 1) {
            taken.add(r * MAX_COLS + c);
          }
        }
      }
      col = right;
      cols = Math.max(cols, col);
      rows = Math.max(rows, bottom);
    }
    row += 1;
    rows = Math.max(rows, row);
  }
  return rows * cols > MAX_COPY_CELLS ? null : { cells, rows, cols };
}

function jsonRecord(text: string | null): Record<string, unknown> | null {
  if (!text) {
    return null;
  }
  try {
    const value: unknown = JSON.parse(text);
    return typeof value === "object" && value !== null && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
  } catch {
    return null;
  }
}

/** What one producer's cell holds: its value, format and (Sheets) R1C1
 * formula; `fromAttribute`: the value is a 1900-system serial. */
interface Reading {
  value: CellValue;
  formula?: string;
  r1c1?: string;
  format?: Format;
  quotePrefix?: true;
  fromAttribute?: true;
}

function fromText(text: string, ctx: ImportContext): Reading {
  const field = parsedField(text, ctx);
  const reading: Reading = { value: field.value };
  if (field.formula !== undefined) {
    reading.formula = field.formula;
  }
  if (field.numberFormat !== undefined) {
    reading.format = field.builtinFormatId !== undefined ? { numberFormat: field.numberFormat, builtinFormatId: field.builtinFormatId } : { numberFormat: field.numberFormat };
  }
  if (field.quotePrefix) {
    reading.quotePrefix = true;
  }
  return reading;
}

function readExcel(cell: Element, text: string, css: Declarations, ctx: ImportContext): Reading {
  const format = formatOf(css.has("mso-number-format") ? decodeCssString(css.get("mso-number-format")!) : undefined);
  const withFormat = (reading: Reading): Reading => (format ? { ...reading, format } : reading);
  const num = cell.getAttribute("x:num");
  if (num !== null) {
    const value = Number(num);
    return withFormat(num.trim() !== "" && Number.isFinite(value) ? { value, fromAttribute: true } : fromText(text, ctx));
  }
  const bool = cell.getAttribute("x:bool");
  if (bool !== null) {
    return { value: (bool || text).trim().toUpperCase() === "TRUE" };
  }
  const error = cell.getAttribute("x:err");
  if (error !== null) {
    return { value: toCellError(error || text) };
  }
  const str = cell.getAttribute("x:str");
  if (str !== null) {
    return withFormat({ value: attributeText(str !== "" ? str : text) });
  }
  const reading = fromText(text, ctx);
  return format && reading.formula === undefined ? { ...reading, format } : reading;
}

function readSheets(cell: Element, text: string, ctx: ImportContext): Reading {
  const data = jsonRecord(cell.getAttribute("data-sheets-value"));
  const numberFormat = jsonRecord(cell.getAttribute("data-sheets-numberformat"));
  const code = numberFormat?.["2"];
  const format = typeof code === "string" ? formatOf(code) : undefined;
  let reading: Reading;
  const type = data?.["1"];
  if (type === 3 && typeof data?.["3"] === "number" && Number.isFinite(data["3"])) {
    reading = { value: data["3"], fromAttribute: true };
  } else if (type === 2 && typeof data?.["2"] === "string") {
    reading = { value: attributeText(data["2"]) };
  } else if (type === 4 && (typeof data?.["4"] === "boolean" || typeof data?.["4"] === "number")) {
    reading = { value: Boolean(data["4"]) };
  } else {
    reading = fromText(text, ctx);
    delete reading.formula;
  }
  if (format) {
    reading.format = format;
  }
  // A formula past the payload's length bound, or holding what XML cannot,
  // is no formula: the value pastes.
  const formula = cell.getAttribute("data-sheets-formula");
  if (formula && formula.startsWith("=") && formula.length <= MAX_CLIPBOARD_FORMULA_LENGTH + 1 && !hasIllegalXmlScalars(formula)) {
    reading.r1c1 = formula;
  }
  return reading;
}

function readLibreOffice(cell: Element, text: string, ctx: ImportContext): Reading {
  const sdnum = cell.getAttribute("sdnum") ?? "";
  const first = sdnum.indexOf(";");
  const second = first < 0 ? -1 : sdnum.indexOf(";", first + 1);
  const code = second < 0 ? "" : sdnum.slice(second + 1);
  const sdval = cell.getAttribute("sdval");
  if (sdval !== null && sdval.trim() !== "" && Number.isFinite(Number(sdval))) {
    if (code.trim().toUpperCase() === "BOOLEAN") {
      return { value: Number(sdval) !== 0 };
    }
    const format = formatOf(code);
    return format ? { value: Number(sdval), format, fromAttribute: true } : { value: Number(sdval), fromAttribute: true };
  }
  const reading = fromText(text, ctx);
  const format = code.trim().toUpperCase() === "BOOLEAN" ? undefined : formatOf(code);
  return format && reading.formula === undefined ? { ...reading, format } : reading;
}

/** Whether the HTML has more `<td>`/`<th>` tags than MAX_COPY_CELLS —
 * counted, stopping past the cap, never collected (64 MiB of them would be
 * some 16 million match strings). */
function moreCellTagsThanCap(html: string): boolean {
  const tag = /<t[dh][\s>/]/gi;
  for (let count = 0; tag.exec(html) !== null; count += 1) {
    if (count >= MAX_COPY_CELLS) {
      return true;
    }
  }
  return false;
}

/** An external clipboard table as a PasteBlock (see the module header):
 * null when the HTML holds no table, a refusal message when it is too
 * large. */
export function importClipboardHtml(html: string, ctx: ImportContext): PasteBlock | string | null {
  if (html.length > MAX_IMPORT_HTML_LENGTH || moreCellTagsThanCap(html)) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  if (typeof DOMParser === "undefined") {
    return null;
  }
  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) {
    return null;
  }
  const placed = placeCells(table);
  if (!placed) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  const source = detectSource(html, doc);
  const rules = source === "excel" ? cssRules(Array.from(doc.querySelectorAll("style"), (style) => style.textContent ?? "").join("\n")) : null;
  const styles: CellStyle[] = [];
  const styleIndex = new Map<string, number>();
  const cells: PasteCell[] = [];
  const merges: GridRange[] = [];
  for (const { element, row, col, rowspan, colspan } of placed.cells) {
    if (rowspan > 1 || colspan > 1) {
      merges.push({ top: row, left: col, bottom: row + rowspan - 1, right: col + colspan - 1 });
    }
    const inline = cssDeclarations(element.getAttribute("style") ?? "");
    let css = inline;
    if (rules) {
      css = new Map(rules.base);
      for (const name of Array.from(element.classList)) {
        for (const [key, value] of rules.classes.get(name.toLowerCase()) ?? []) {
          css.set(key, value);
        }
      }
      for (const [key, value] of inline) {
        css.set(key, value);
      }
    }
    const text = cellText(element);
    const reading =
      source === "excel" ? readExcel(element, text, css, ctx) : source === "sheets" ? readSheets(element, text, ctx) : source === "libreoffice" ? readLibreOffice(element, text, ctx) : fromText(text, ctx);
    const style = importedStyle(css, source === "excel" || source === "sheets" ? {} : tagFormat(element, source === "libreoffice"));
    if (reading.format) {
      style.numberFormat = reading.format.numberFormat;
      if (reading.format.builtinFormatId !== undefined) {
        style.builtinFormatId = reading.format.builtinFormatId;
      }
    }
    if (reading.quotePrefix) {
      style.quotePrefix = true;
    }
    let value = reading.value;
    if (reading.fromAttribute && ctx.dateSystem === "1904" && typeof value === "number" && dateKindOf(reading.format?.numberFormat) === "date") {
      value -= DATE_SYSTEM_OFFSET;
    }
    const hasStyle = Object.keys(style).length > 0;
    if (value === null && reading.formula === undefined && reading.r1c1 === undefined && !hasStyle) {
      continue;
    }
    const cell: PasteCell = { r: row, c: col, v: value };
    if (reading.formula !== undefined) {
      cell.f = reading.formula;
    }
    if (reading.r1c1 !== undefined) {
      cell.r1c1 = reading.r1c1;
    }
    if (hasStyle) {
      const key = styleKey(style);
      let index = styleIndex.get(key);
      if (index === undefined) {
        index = styles.length;
        styles.push(style);
        styleIndex.set(key, index);
      }
      cell.s = index;
    }
    cells.push(cell);
  }
  return {
    source,
    workbookId: "",
    sheetName: "",
    dateSystem: ctx.dateSystem,
    origin: { row: 0, col: 0 },
    rows: placed.rows,
    cols: placed.cols,
    cells,
    styles,
    merges,
  };
}
