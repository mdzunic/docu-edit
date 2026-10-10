// The xlsx cell-style table (SPEC-117 "Reading xlsx styles"): `xl/styles.xml`
// → SpreadsheetModel.styles, read with DOMParser (core/xml.ts — the part is
// small). SPEC-117 is the normative source for the mapping:
// - `fonts/font`: `b`, `i`, `strike` (`val` absent or true), `u@val` (absent →
//   single, `none` → absent), `vertAlign@val` (`baseline` → absent), `sz`,
//   `color`, `name` → family and `scheme` minor/major → `family: { theme }`
//   (the scheme wins over the name — it is what Excel renders). A rich-text
//   run's `<rPr>` maps through the same XlsxFontReader (`rFont` is its name).
// - `fills/fill`: `patternFill@patternType` none/absent → no fill; `solid` →
//   a solid fill of `fgColor` (black when automatic); any other pattern with
//   its fg/bg; `gradientFill` (`linear` by default, or `path`) with its
//   degree, path focus and stops.
// - `borders/border`: `left`, `right`, `top`, `bottom`, `diagonal` (`start`/
//   `end` read as left/right) with `@style` (absent/none → no edge) and
//   colour; `diagonalUp`/`diagonalDown` → the diagonal's up/down.
// - `cellXfs/xf` → one CellStyle per xf, IN ORDER and never deduplicated
//   (`styles[i].sourceXf === i`, so `<c s="7">` is styleId 7), each holding
//   COPIES of its font, fill and border (plain data, no shared objects), its
//   `alignment` and `protection` children and `quotePrefix`. The `apply*`
//   flags are ignored (Excel writes cellXfs fully resolved) and `xfId` (the
//   named style) is SPEC-123's.
// - `<colors><indexedColors>` → the 64-entry override palette.
// - SPEC-118 "Reading xlsx number formats": `numFmts/numFmt@numFmtId` +
//   `@formatCode` (attribute-decoded) are read into a map; each
//   `xf@numFmtId` (absent → 0) resolves: a map entry → `numberFormat` (any
//   id, Excel writes 5–8 and 41–44 explicitly); else a locale-dependent
//   built-in → `builtinFormatId` and the en-US code; else another built-in →
//   its code; 0, unknown and CJK ids → absent. A code equal to `General` is
//   stored absent; `applyNumberFormat` is ignored like every `apply*`.
// cellStyleXfs/cellStyles (SPEC-123), dxfs (SPEC-129), tableStyles
// (SPEC-133), extLst and colors/mruColors stay unread in the retained part.
// Colours map through themeColor.ts, the only module that interprets them.
//
// The style writer (SPEC-122 "Styles" › xlsx, normative there): one
// XlsxStyleWriter per save turns a model style id into a cellXfs index.
// - A style with `sourceXf` writes that index. Otherwise its font, fill,
//   border and number format each reuse the first source entry the reader
//   above parses to an equal value (compared as styleKey JSON), else an
//   earlier append, else are appended; an absent font is font 0, an absent
//   fill or border the first entry that reads as none.
// - `numFmtId`: `builtinFormatId` when set (no numFmt entry); else the first
//   built-in id 0–49 outside LOCALE_DEPENDENT_BUILTINS whose code equals
//   `numberFormat` (and that the file does not redefine); else an existing
//   numFmt with the same code; else a new numFmt with id max(163, highest
//   numFmt id) + 1.
// - The xf: `<xf numFmtId fontId fillId borderId xfId>` (xfId the
//   cellStyleXfs index of `namedStyle` through cellStyles; absent or
//   "Normal" → 0), `quotePrefix="1"`, then `applyNumberFormat`/`applyFont`/
//   `applyFill`/`applyBorder`/`applyAlignment`/`applyProtection="1"` where
//   the xf differs from that parent xf; children `<alignment>` and
//   `<protection>` (`locked="0"` only when false, `hidden="1"` only when
//   true). Equal appended xfs are one entry.
// - XML: fonts in Excel's child order (b, i, strike, u, vertAlign, sz, color,
//   name, family, scheme — a theme font writes its theme typeface, family 2
//   and `<scheme val>`); solid fills `<patternFill patternType="solid">
//   <fgColor …/><bgColor indexed="64"/>`; border edges in the order left,
//   right, top, bottom, diagonal, an automatic colour as `indexed="64"`;
//   every colour through themeColor.ts's modelColorToXlsxAttrs.
// - Appends are spliced before each collection's close tag with its `count`
//   set to the new total; an absent collection is inserted in CT_Stylesheet
//   order (numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles,
//   dxfs, tableStyles, colors, extLst); the part is rewritten only when
//   something was appended — every source entry and every other element
//   stays byte-verbatim.
// - From scratch (no source part) the base is a minimal styles part whose
//   font 0 is built from `model.styles[0]` and whose xf 0 IS style 0, so
//   unstyled cells keep the workbook's own default font in either direction;
//   a blank workbook's base is SPEC-121's FROM_SCRATCH_STYLES_XML.
import { DEFAULT_THEME, isThemeFont, resolveFont, type DocumentTheme, type ModelColor, type ModelFont } from "../../core/theme";
import { parseXml } from "../../core/xml";
import { scanXml } from "../../core/xmlScanner";
import type { CellAlignment, CellBorderEdge, CellBorders, CellFill, CellFont, CellStyle, SheetBorderStyle, SpreadsheetModel } from "../model";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "../numberFormat/builtin";
import { defaultStyles, INDEXED_COLORS, styleKey } from "../styles";
import { applySplices, escapeAttributeValue, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { elementAttributes, modelColorToXlsxAttrs, xlsxColorFrom, xlsxColorFromElement, xlsxRgbToHex, type AttributeGetter } from "./themeColor";

/** SpreadsheetML's main namespace, transitional and strict. */
export const SPREADSHEETML_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main",
]);

/** Whether a namespace URI is SpreadsheetML's main namespace. */
export function isSpreadsheetMl(ns: string | null | undefined): boolean {
  return ns !== null && ns !== undefined && SPREADSHEETML_NAMESPACES.has(ns);
}

/** The workbook default font when there is no styles part (or it names
 * none): Calibri 11, Excel's Normal style before Office 2023. */
export const DEFAULT_XLSX_FONT: Readonly<{ family: string; sizePt: number }> = { family: "Calibri", sizePt: 11 };

/** What readXlsxStyles hands back. */
export interface XlsxStyleTable {
  /** One style per cellXfs xf, in order — never empty (`[{}]` without a
   * styles part). */
  styles: CellStyle[];
  /** The file's `<colors><indexedColors>` palette, lowercase #rrggbb. */
  indexedColors?: string[];
  /** The Normal style's font, its family resolved against the theme — the
   * font the maximum digit width (geometry.ts) is looked up for. */
  defaultFont: { family: string; sizePt: number };
}

function isOn(value: string | null | undefined): boolean {
  return value === "1" || value === "true";
}

function parseNumber(text: string | null | undefined): number | undefined {
  if (text === null || text === undefined || text.trim() === "") {
    return undefined;
  }
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

function parseIndex(text: string | null | undefined): number | undefined {
  const value = parseNumber(text);
  return value !== undefined && Number.isInteger(value) && value >= 0 ? value : undefined;
}

const UNDERLINES: ReadonlySet<string> = new Set(["single", "double", "singleAccounting", "doubleAccounting"]);

/** Builds a CellFont from the children of a `<font>` (styles.xml) or a
 * rich-text run's `<rPr>` — one mapping for both. Feed each child element's
 * local name and attributes with `add`, then `build`. */
export class XlsxFontReader {
  private name: string | undefined;
  private scheme: "major" | "minor" | undefined;
  private sizePt: number | undefined;
  private bold = false;
  private italic = false;
  private underline: CellFont["underline"];
  private strike = false;
  private color: ModelColor | undefined;
  private vertAlign: CellFont["vertAlign"];

  constructor(private readonly indexedColors?: readonly string[]) {}

  /** One property element. `condense`, `extend`, `outline`, `shadow`,
   * `family` and `charset` are not modelled. */
  add(local: string, attr: AttributeGetter): void {
    const val = attr("val");
    switch (local) {
      case "b":
        this.bold = val === undefined || isOn(val);
        break;
      case "i":
        this.italic = val === undefined || isOn(val);
        break;
      case "strike":
        this.strike = val === undefined || isOn(val);
        break;
      case "u":
        this.underline = val === undefined ? "single" : UNDERLINES.has(val) ? (val as CellFont["underline"]) : undefined;
        break;
      case "vertAlign":
        this.vertAlign = val === "superscript" || val === "subscript" ? val : undefined;
        break;
      case "sz": {
        const size = parseNumber(val);
        this.sizePt = size !== undefined && size > 0 ? size : undefined;
        break;
      }
      case "color":
        this.color = xlsxColorFrom(attr, this.indexedColors);
        break;
      case "name":
      case "rFont":
        this.name = val !== undefined && val !== "" ? val : undefined;
        break;
      case "scheme":
        this.scheme = val === "major" || val === "minor" ? val : undefined;
        break;
      default:
        break;
    }
  }

  /** The font, with only the properties present. */
  build(): CellFont {
    const font: CellFont = {};
    const family: ModelFont | undefined = this.scheme !== undefined ? { theme: this.scheme } : this.name;
    if (family !== undefined) {
      font.family = family;
    }
    if (this.sizePt !== undefined) {
      font.sizePt = this.sizePt;
    }
    if (this.bold) {
      font.bold = true;
    }
    if (this.italic) {
      font.italic = true;
    }
    if (this.underline !== undefined) {
      font.underline = this.underline;
    }
    if (this.strike) {
      font.strike = true;
    }
    if (this.color !== undefined) {
      font.color = this.color;
    }
    if (this.vertAlign !== undefined) {
      font.vertAlign = this.vertAlign;
    }
    return font;
  }
}

// --- DOM helpers --------------------------------------------------------------

function childrenNamed(parent: Element, local: string): Element[] {
  return Array.from(parent.children).filter((child) => child.localName === local && isSpreadsheetMl(child.namespaceURI));
}

function childNamed(parent: Element | undefined, local: string): Element | undefined {
  return parent ? childrenNamed(parent, local)[0] : undefined;
}

function colorChild(parent: Element, local: string, indexedColors: readonly string[] | undefined): ModelColor | undefined {
  const element = childNamed(parent, local);
  return element ? xlsxColorFromElement(element, indexedColors) : undefined;
}

// --- fonts, fills, borders ----------------------------------------------------

function readFont(element: Element, indexedColors: readonly string[] | undefined): CellFont {
  const reader = new XlsxFontReader(indexedColors);
  for (const child of Array.from(element.children)) {
    if (isSpreadsheetMl(child.namespaceURI)) {
      reader.add(child.localName, elementAttributes(child));
    }
  }
  return reader.build();
}

type GradientFill = Extract<CellFill, { kind: "gradient" }>;
type PatternFill = Extract<CellFill, { kind: "pattern" }>;

const GRADIENT_FOCUS = ["left", "right", "top", "bottom"] as const;

function readGradient(element: Element, indexedColors: readonly string[] | undefined): GradientFill {
  const fill: GradientFill = { kind: "gradient", type: element.getAttribute("type") === "path" ? "path" : "linear", stops: [] };
  const degree = parseNumber(element.getAttribute("degree"));
  if (degree !== undefined) {
    fill.degree = degree;
  }
  for (const side of GRADIENT_FOCUS) {
    const value = parseNumber(element.getAttribute(side));
    if (value !== undefined) {
      fill[side] = value;
    }
  }
  for (const stop of childrenNamed(element, "stop")) {
    // An automatic stop colour has no meaning of its own; black, like a
    // solid fill's automatic foreground.
    fill.stops.push({
      position: parseNumber(stop.getAttribute("position")) ?? 0,
      color: colorChild(stop, "color", indexedColors) ?? "#000000",
    });
  }
  return fill;
}

function readFill(element: Element, indexedColors: readonly string[] | undefined): CellFill | undefined {
  const pattern = childNamed(element, "patternFill");
  if (pattern) {
    const type = pattern.getAttribute("patternType");
    if (type === null || type === "" || type === "none") {
      return undefined;
    }
    const fg = colorChild(pattern, "fgColor", indexedColors);
    if (type === "solid") {
      return { kind: "solid", color: fg ?? "#000000" };
    }
    const fill: PatternFill = { kind: "pattern", pattern: type };
    if (fg !== undefined) {
      fill.fg = fg;
    }
    const bg = colorChild(pattern, "bgColor", indexedColors);
    if (bg !== undefined) {
      fill.bg = bg;
    }
    return fill;
  }
  const gradient = childNamed(element, "gradientFill");
  return gradient ? readGradient(gradient, indexedColors) : undefined;
}

const BORDER_STYLES: ReadonlySet<string> = new Set<SheetBorderStyle>([
  "thin",
  "medium",
  "thick",
  "dashed",
  "dotted",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
]);

/** Border child element → the edge it fills (left-to-right sheets: `start`
 * is the left edge, `end` the right). */
const BORDER_EDGES: ReadonlyMap<string, keyof CellBorders> = new Map<string, keyof CellBorders>([
  ["left", "left"],
  ["start", "left"],
  ["right", "right"],
  ["end", "right"],
  ["top", "top"],
  ["bottom", "bottom"],
  ["diagonal", "diagonal"],
]);

function readBorder(element: Element, indexedColors: readonly string[] | undefined): CellBorders | undefined {
  const borders: CellBorders = {};
  let any = false;
  for (const child of Array.from(element.children)) {
    const side = isSpreadsheetMl(child.namespaceURI) ? BORDER_EDGES.get(child.localName) : undefined;
    const style = child.getAttribute("style");
    if (side === undefined || style === null || !BORDER_STYLES.has(style)) {
      continue; // absent, none or unknown → no edge
    }
    const edge: CellBorderEdge = { style: style as SheetBorderStyle };
    const color = colorChild(child, "color", indexedColors);
    if (color !== undefined) {
      edge.color = color;
    }
    borders[side] = edge;
    any = true;
  }
  if (borders.diagonal) {
    if (isOn(element.getAttribute("diagonalUp"))) {
      borders.diagonal.up = true;
    }
    if (isOn(element.getAttribute("diagonalDown"))) {
      borders.diagonal.down = true;
    }
  }
  return any ? borders : undefined;
}

// --- cellXfs ------------------------------------------------------------------

const HORIZONTAL: ReadonlySet<string> = new Set(["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]);
const VERTICAL: ReadonlySet<string> = new Set(["top", "center", "bottom", "justify", "distributed"]);

function readAlignment(element: Element | undefined): CellAlignment | undefined {
  if (!element) {
    return undefined;
  }
  const alignment: CellAlignment = {};
  const horizontal = element.getAttribute("horizontal");
  if (horizontal !== null && HORIZONTAL.has(horizontal)) {
    alignment.horizontal = horizontal as CellAlignment["horizontal"];
  }
  const vertical = element.getAttribute("vertical");
  if (vertical !== null && VERTICAL.has(vertical)) {
    alignment.vertical = vertical as CellAlignment["vertical"];
  }
  if (isOn(element.getAttribute("wrapText"))) {
    alignment.wrap = true;
  }
  const indent = parseIndex(element.getAttribute("indent"));
  if (indent !== undefined && indent > 0) {
    alignment.indent = indent;
  }
  const rotation = parseIndex(element.getAttribute("textRotation"));
  if (rotation !== undefined && rotation > 0 && (rotation <= 180 || rotation === 255)) {
    alignment.textRotation = rotation;
  }
  if (isOn(element.getAttribute("shrinkToFit"))) {
    alignment.shrinkToFit = true;
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined;
}

function readProtection(element: Element | undefined): CellStyle["protection"] {
  if (!element) {
    return undefined;
  }
  const protection: NonNullable<CellStyle["protection"]> = {};
  const locked = element.getAttribute("locked");
  if (locked === "0" || locked === "false") {
    protection.locked = false;
  }
  if (isOn(element.getAttribute("hidden"))) {
    protection.hidden = true;
  }
  return Object.keys(protection).length > 0 ? protection : undefined;
}

/** A fresh copy of a parsed font/fill/border: every style owns its objects. */
function copy<T>(value: T): T {
  return structuredClone(value);
}

/** `<numFmts>`: each `numFmt@numFmtId` → its `formatCode` (attribute-decoded). */
function readNumFmts(root: Element): Map<number, string> {
  const codes = new Map<number, string>();
  const numFmts = childNamed(root, "numFmts");
  for (const numFmt of numFmts ? childrenNamed(numFmts, "numFmt") : []) {
    const id = parseIndex(numFmt.getAttribute("numFmtId"));
    const code = numFmt.getAttribute("formatCode");
    if (id !== undefined && code !== null) {
      codes.set(id, code);
    }
  }
  return codes;
}

/** An xf's number format (SPEC-118): the file's own entry for the id, else
 * the built-in — a locale-dependent one also as `builtinFormatId`; General,
 * unknown and CJK ids leave both absent. */
export function xfNumberFormat(numFmtId: number, numFmts: ReadonlyMap<number, string>): Pick<CellStyle, "numberFormat" | "builtinFormatId"> {
  const explicit = numFmts.get(numFmtId);
  const code = explicit ?? BUILTIN_NUMBER_FORMATS[numFmtId];
  if (code === undefined || /^general$/i.test(code.trim())) {
    return {};
  }
  if (explicit === undefined && LOCALE_DEPENDENT_BUILTINS.includes(numFmtId)) {
    return { numberFormat: code, builtinFormatId: numFmtId };
  }
  return { numberFormat: code };
}

function readXf(
  xf: Element,
  index: number,
  fonts: readonly CellFont[],
  fills: readonly (CellFill | undefined)[],
  borders: readonly (CellBorders | undefined)[],
  numFmts: ReadonlyMap<number, string>,
): CellStyle {
  const style: CellStyle = {};
  // A missing id is the first record, as every producer means it.
  const font = fonts[parseIndex(xf.getAttribute("fontId")) ?? 0];
  if (font !== undefined) {
    style.font = copy(font);
  }
  const fill = fills[parseIndex(xf.getAttribute("fillId")) ?? 0];
  if (fill !== undefined) {
    style.fill = copy(fill);
  }
  const border = borders[parseIndex(xf.getAttribute("borderId")) ?? 0];
  if (border !== undefined) {
    style.borders = copy(border);
  }
  const alignment = readAlignment(childNamed(xf, "alignment"));
  if (alignment !== undefined) {
    style.alignment = alignment;
  }
  const protection = readProtection(childNamed(xf, "protection"));
  if (protection !== undefined) {
    style.protection = protection;
  }
  Object.assign(style, xfNumberFormat(parseIndex(xf.getAttribute("numFmtId")) ?? 0, numFmts));
  if (isOn(xf.getAttribute("quotePrefix"))) {
    style.quotePrefix = true;
  }
  style.sourceXf = index;
  return style;
}

/** `<colors><indexedColors>`: up to 64 `rgbColor@rgb` entries, lowercase;
 * an unreadable entry keeps the legacy colour at its position. */
function readIndexedColors(root: Element): string[] | undefined {
  const palette = childNamed(childNamed(root, "colors"), "indexedColors");
  if (!palette) {
    return undefined;
  }
  const colors = childrenNamed(palette, "rgbColor")
    .slice(0, 64)
    .map((entry, index) => xlsxRgbToHex(entry.getAttribute("rgb") ?? "") ?? INDEXED_COLORS[index]!);
  return colors.length > 0 ? colors : undefined;
}

function defaultTable(): XlsxStyleTable {
  return { styles: defaultStyles(), defaultFont: { ...DEFAULT_XLSX_FONT } };
}

/** Reads `xl/styles.xml` (undefined: the workbook has no styles part) into
 * the model's style table, the override palette and the default font, whose
 * family is resolved against `theme` (DEFAULT_THEME when the workbook has
 * none). A part that is not well-formed XML reads as no part at all, as
 * core/themeDrawingMl.ts's readThemePart treats a broken theme. */
export function readXlsxStyles(xml: string | undefined, theme: DocumentTheme | undefined): XlsxStyleTable {
  if (xml === undefined) {
    return defaultTable();
  }
  let root: Element;
  try {
    root = parseXml(xml).documentElement;
  } catch {
    return defaultTable();
  }
  if (root.localName !== "styleSheet" || !isSpreadsheetMl(root.namespaceURI)) {
    return defaultTable();
  }
  const indexedColors = readIndexedColors(root);
  const fontsElement = childNamed(root, "fonts");
  const fonts = fontsElement ? childrenNamed(fontsElement, "font").map((font) => readFont(font, indexedColors)) : [];
  const fillsElement = childNamed(root, "fills");
  const fills = fillsElement ? childrenNamed(fillsElement, "fill").map((fill) => readFill(fill, indexedColors)) : [];
  const bordersElement = childNamed(root, "borders");
  const borders = bordersElement ? childrenNamed(bordersElement, "border").map((border) => readBorder(border, indexedColors)) : [];
  const xfsElement = childNamed(root, "cellXfs");
  const xfs = xfsElement ? childrenNamed(xfsElement, "xf") : [];
  const numFmts = readNumFmts(root);
  const styles = xfs.map((xf, index) => readXf(xf, index, fonts, fills, borders, numFmts));

  // The Normal style's font: cellXfs[0]'s (fonts[0] when there is no xf).
  const normalFont = fonts[xfs[0] ? (parseIndex(xfs[0].getAttribute("fontId")) ?? 0) : 0];
  const family = normalFont?.family !== undefined ? resolveFont(normalFont.family, theme ?? DEFAULT_THEME) : DEFAULT_XLSX_FONT.family;
  const table: XlsxStyleTable = {
    styles: styles.length > 0 ? styles : defaultStyles(),
    defaultFont: { family, sizePt: normalFont?.sizePt ?? DEFAULT_XLSX_FONT.sizePt },
  };
  if (indexedColors !== undefined) {
    table.indexedColors = indexedColors;
  }
  return table;
}

// --- the style writer (SPEC-122) ------------------------------------------------

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** A number as attribute text: the shortest round-trip form. */
function numberText(value: number): string {
  return Object.is(value, -0) ? "0" : String(value);
}

/** A CellFont as SpreadsheetML font properties (styles.xml `<font>` content,
 * or a rich run's `<rPr>` content), in Excel's child order: b, i, strike, u,
 * vertAlign, sz, color, name, family, scheme. A theme font writes its theme
 * typeface, `<family val="2"/>` and `<scheme val>`; `nameElement` is "name"
 * in styles.xml and "rFont" in a run; `prefix` is the part's element prefix. */
export function xlsxFontProperties(font: CellFont, theme: DocumentTheme | undefined, nameElement: "name" | "rFont" = "name", prefix = ""): string {
  const p = prefix;
  let out = "";
  if (font.bold) {
    out += `<${p}b/>`;
  }
  if (font.italic) {
    out += `<${p}i/>`;
  }
  if (font.strike) {
    out += `<${p}strike/>`;
  }
  if (font.underline !== undefined) {
    out += font.underline === "single" ? `<${p}u/>` : `<${p}u val="${font.underline}"/>`;
  }
  if (font.vertAlign !== undefined) {
    out += `<${p}vertAlign val="${font.vertAlign}"/>`;
  }
  if (font.sizePt !== undefined) {
    out += `<${p}sz val="${numberText(font.sizePt)}"/>`;
  }
  if (font.color !== undefined) {
    out += `<${p}color${modelColorToXlsxAttrs(font.color, theme)}/>`;
  }
  if (font.family !== undefined) {
    out += `<${p}${nameElement} val="${escapeAttributeValue(resolveFont(font.family, theme ?? DEFAULT_THEME))}"/>`;
    if (isThemeFont(font.family)) {
      out += `<${p}family val="2"/><${p}scheme val="${font.family.theme}"/>`;
    }
  }
  return out;
}

function colorElement(local: string, color: ModelColor, theme: DocumentTheme | undefined, p: string): string {
  return `<${p}${local}${modelColorToXlsxAttrs(color, theme)}/>`;
}

/** A `<fill>` (undefined: `patternType="none"`). */
function fillXml(fill: CellFill | undefined, theme: DocumentTheme | undefined, p: string): string {
  if (fill === undefined) {
    return `<${p}fill><${p}patternFill patternType="none"/></${p}fill>`;
  }
  if (fill.kind === "solid") {
    return `<${p}fill><${p}patternFill patternType="solid">${colorElement("fgColor", fill.color, theme, p)}<${p}bgColor indexed="64"/></${p}patternFill></${p}fill>`;
  }
  if (fill.kind === "pattern") {
    const colors = (fill.fg !== undefined ? colorElement("fgColor", fill.fg, theme, p) : "") + (fill.bg !== undefined ? colorElement("bgColor", fill.bg, theme, p) : "");
    const type = `patternType="${escapeAttributeValue(fill.pattern)}"`;
    return `<${p}fill>${colors === "" ? `<${p}patternFill ${type}/>` : `<${p}patternFill ${type}>${colors}</${p}patternFill>`}</${p}fill>`;
  }
  let attrs = fill.type === "path" ? ` type="path"` : "";
  if (fill.degree !== undefined) {
    attrs += ` degree="${numberText(fill.degree)}"`;
  }
  for (const side of GRADIENT_FOCUS) {
    const value = fill[side];
    if (value !== undefined) {
      attrs += ` ${side}="${numberText(value)}"`;
    }
  }
  const stops = fill.stops.map((stop) => `<${p}stop position="${numberText(stop.position)}">${colorElement("color", stop.color, theme, p)}</${p}stop>`).join("");
  return `<${p}fill><${p}gradientFill${attrs}>${stops}</${p}gradientFill></${p}fill>`;
}

const BORDER_ORDER = ["left", "right", "top", "bottom", "diagonal"] as const;

/** A `<border>` (undefined: every edge empty). */
function borderXml(borders: CellBorders | undefined, theme: DocumentTheme | undefined, p: string): string {
  const diagonal = borders?.diagonal;
  const attrs = (diagonal?.up ? ` diagonalUp="1"` : "") + (diagonal?.down ? ` diagonalDown="1"` : "");
  const edges = BORDER_ORDER.map((side) => {
    const edge = borders?.[side];
    if (edge === undefined) {
      return `<${p}${side}/>`;
    }
    const color = edge.color !== undefined ? colorElement("color", edge.color, theme, p) : `<${p}color indexed="64"/>`;
    return `<${p}${side} style="${edge.style}">${color}</${p}${side}>`;
  }).join("");
  return `<${p}border${attrs}>${edges}</${p}border>`;
}

function alignmentXml(alignment: CellAlignment | undefined, p: string): string {
  if (alignment === undefined) {
    return "";
  }
  let attrs = "";
  if (alignment.horizontal !== undefined) {
    attrs += ` horizontal="${alignment.horizontal}"`;
  }
  if (alignment.vertical !== undefined) {
    attrs += ` vertical="${alignment.vertical}"`;
  }
  if (alignment.textRotation !== undefined) {
    attrs += ` textRotation="${alignment.textRotation}"`;
  }
  if (alignment.wrap) {
    attrs += ` wrapText="1"`;
  }
  if (alignment.indent !== undefined) {
    attrs += ` indent="${alignment.indent}"`;
  }
  if (alignment.shrinkToFit) {
    attrs += ` shrinkToFit="1"`;
  }
  return attrs === "" ? "" : `<${p}alignment${attrs}/>`;
}

function protectionXml(protection: CellStyle["protection"], p: string): string {
  const attrs = (protection?.locked === false ? ` locked="0"` : "") + (protection?.hidden ? ` hidden="1"` : "");
  return attrs === "" ? "" : `<${p}protection${attrs}/>`;
}

/** The collections the writer appends to. */
type StyleCollection = "numFmts" | "fonts" | "fills" | "borders" | "cellXfs";

/** CT_Stylesheet's children, in schema order. */
const STYLESHEET_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];

/** A level-2 element of the source part, located by the scanner. */
interface ElementSpan {
  local: string;
  name: string;
  start: number;
  /** End of the start tag (of the element, when self-closing). */
  end: number;
  selfClosing: boolean;
  /** Start of the close tag (-1 when self-closing). */
  closeStart: number;
}

/** A cellStyleXfs entry, as the apply flags compare against it. */
interface ParentXf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment: string;
  protection: string;
}

const NO_PARENT: ParentXf = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: styleKey({}), protection: styleKey({}) };

/** The default font a from-scratch font 0 falls back to: the blank
 * workbook's (SPEC-121 "The blank workbook"). */
const BLANK_DEFAULT_FONT: CellFont = { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } };

/** A from-scratch styles part before any xf: font 0 from the default style,
 * the two reserved fills, one empty border, the Normal style's xf. */
function fromScratchBase(defaultStyle: CellStyle | undefined, theme: DocumentTheme | undefined): string {
  const font = defaultStyle?.font !== undefined && Object.keys(defaultStyle.font).length > 0 ? defaultStyle.font : BLANK_DEFAULT_FONT;
  return (
    `${XML_DECLARATION}<styleSheet xmlns="${MAIN_NS}">` +
    `<fonts count="1"><font>${xlsxFontProperties(font, theme)}</font></fonts>` +
    `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="0"></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `<dxfs count="0"/><tableStyles count="0"/></styleSheet>`
  );
}

function intAttr(element: Element, name: string): number {
  return parseIndex(element.getAttribute(name)) ?? 0;
}

/** The workbook's cellXfs table for one save (see the module header). */
export class XlsxStyleWriter {
  private readonly text: string;
  private readonly fromSource: boolean;
  private readonly prefix: string;
  private readonly spans = new Map<string, ElementSpan>();
  private readonly order: ElementSpan[] = [];
  private rootCloseStart = -1;
  private readonly counts: Record<StyleCollection, number> = { numFmts: 0, fonts: 0, fills: 0, borders: 0, cellXfs: 0 };
  private readonly appends: Record<StyleCollection, string[]> = { numFmts: [], fonts: [], fills: [], borders: [], cellXfs: [] };
  /** styleKey → index, per collection (the first source entry, then appends). */
  private readonly fontIds = new Map<string, number>();
  private readonly fillIds = new Map<string, number>();
  private readonly borderIds = new Map<string, number>();
  private readonly numFmtCodes = new Map<number, string>();
  private readonly numFmtByCode = new Map<string, number>();
  private highestNumFmt = 0;
  private readonly parents: ParentXf[] = [];
  private readonly namedXf = new Map<string, number>();
  private readonly xfByXml = new Map<string, number>();
  private readonly xfByStyle = new Map<number, number>();
  private readonly baseAppends: number;

  /** `source` is the retained `xl/styles.xml` text — undefined builds the
   * from-scratch base from `model.styles[0]` (xf 0 is style 0). */
  constructor(
    source: string | undefined,
    private readonly model: Pick<SpreadsheetModel, "styles" | "theme" | "indexedColors">,
  ) {
    this.fromSource = source !== undefined;
    this.text = source ?? fromScratchBase(model.styles[0], model.theme);
    let prefix = "";
    let depth = 0;
    let current: ElementSpan | null = null;
    scanXml(this.text, (event) => {
      if (event.type === "open") {
        const level = depth + 1;
        if (level === 1) {
          prefix = event.name.includes(":") ? `${event.name.slice(0, event.name.indexOf(":"))}:` : "";
        } else if (level === 2) {
          const span: ElementSpan = { local: event.local, name: event.name, start: event.start, end: event.end, selfClosing: event.selfClosing, closeStart: -1 };
          if (isSpreadsheetMl(event.ns)) {
            this.order.push(span);
            if (!this.spans.has(event.local)) {
              this.spans.set(event.local, span);
            }
          }
          current = event.selfClosing ? null : span;
        }
        if (!event.selfClosing) {
          depth = level;
        }
      } else if (event.type === "close") {
        if (depth === 2 && current !== null) {
          current.closeStart = event.start;
          current = null;
        } else if (depth === 1) {
          this.rootCloseStart = event.start;
        }
        depth -= 1;
      }
    });
    this.prefix = prefix;
    this.readSource();
    if (!this.fromSource) {
      // Font 0 was built from the default style's font: it IS that font,
      // whatever a re-read of its XML would compare as.
      const font = model.styles[0]?.font;
      if (font !== undefined && Object.keys(font).length > 0) {
        this.fontIds.set(styleKey({ font }), 0);
      }
      this.xfFor(0);
    }
    this.baseAppends = this.appendCount();
  }

  /** The source collections, parsed by the reader's own mappings. */
  private readSource(): void {
    let root: Element;
    try {
      root = parseXml(this.text).documentElement;
    } catch {
      return;
    }
    const colors = this.model.indexedColors;
    const fontsElement = childNamed(root, "fonts");
    const fonts = fontsElement ? childrenNamed(fontsElement, "font") : [];
    this.counts.fonts = fonts.length;
    fonts.forEach((element, index) => remember(this.fontIds, styleKey({ font: readFont(element, colors) }), index));
    const fillsElement = childNamed(root, "fills");
    const fills = fillsElement ? childrenNamed(fillsElement, "fill") : [];
    this.counts.fills = fills.length;
    fills.forEach((element, index) => remember(this.fillIds, styleKey({ fill: readFill(element, colors) }), index));
    const bordersElement = childNamed(root, "borders");
    const borders = bordersElement ? childrenNamed(bordersElement, "border") : [];
    this.counts.borders = borders.length;
    borders.forEach((element, index) => remember(this.borderIds, styleKey({ borders: readBorder(element, colors) }), index));
    const numFmtsElement = childNamed(root, "numFmts");
    this.counts.numFmts = numFmtsElement ? childrenNamed(numFmtsElement, "numFmt").length : 0;
    for (const [id, code] of readNumFmts(root)) {
      this.numFmtCodes.set(id, code);
      if (!this.numFmtByCode.has(code)) {
        this.numFmtByCode.set(code, id);
      }
      this.highestNumFmt = Math.max(this.highestNumFmt, id);
    }
    const cellXfs = childNamed(root, "cellXfs");
    this.counts.cellXfs = cellXfs ? childrenNamed(cellXfs, "xf").length : 0;
    const styleXfs = childNamed(root, "cellStyleXfs");
    for (const xf of styleXfs ? childrenNamed(styleXfs, "xf") : []) {
      this.parents.push({
        numFmtId: intAttr(xf, "numFmtId"),
        fontId: intAttr(xf, "fontId"),
        fillId: intAttr(xf, "fillId"),
        borderId: intAttr(xf, "borderId"),
        alignment: styleKey({ alignment: readAlignment(childNamed(xf, "alignment")) }),
        protection: styleKey({ protection: readProtection(childNamed(xf, "protection")) }),
      });
    }
    const cellStyles = childNamed(root, "cellStyles");
    for (const cellStyle of cellStyles ? childrenNamed(cellStyles, "cellStyle") : []) {
      const name = cellStyle.getAttribute("name");
      if (name !== null && !this.namedXf.has(name)) {
        this.namedXf.set(name, intAttr(cellStyle, "xfId"));
      }
    }
  }

  private appendCount(): number {
    return Object.values(this.appends).reduce((sum, list) => sum + list.length, 0);
  }

  /** Appends an entry's XML to a collection; its new index. */
  private append(collection: StyleCollection, xml: string): number {
    const index = this.counts[collection] + this.appends[collection].length;
    this.appends[collection].push(xml);
    return index;
  }

  private entryId(ids: Map<string, number>, collection: StyleCollection, key: string, xml: () => string): number {
    const found = ids.get(key);
    if (found !== undefined) {
      return found;
    }
    const index = this.append(collection, xml());
    ids.set(key, index);
    return index;
  }

  private numFmtIdFor(style: CellStyle): number {
    if (style.builtinFormatId !== undefined) {
      return style.builtinFormatId;
    }
    const code = style.numberFormat;
    if (code === undefined || /^general$/i.test(code.trim())) {
      return 0;
    }
    for (let id = 0; id <= 49; id += 1) {
      if (!LOCALE_DEPENDENT_BUILTINS.includes(id) && BUILTIN_NUMBER_FORMATS[id] === code && (this.numFmtCodes.get(id) ?? code) === code) {
        return id;
      }
    }
    const existing = this.numFmtByCode.get(code);
    if (existing !== undefined) {
      return existing;
    }
    const id = Math.max(163, this.highestNumFmt) + 1;
    this.append("numFmts", `<${this.prefix}numFmt numFmtId="${id}" formatCode="${escapeAttributeValue(code)}"/>`);
    this.numFmtCodes.set(id, code);
    this.numFmtByCode.set(code, id);
    this.highestNumFmt = id;
    return id;
  }

  /** The cellXfs index a model style is written with (see the module
   * header). An id outside `model.styles` is the default style's. */
  xfFor(styleId: number): number {
    const cached = this.xfByStyle.get(styleId);
    if (cached !== undefined) {
      return cached;
    }
    const styles = this.model.styles;
    const style = (Number.isInteger(styleId) && styleId >= 0 ? styles[styleId] : undefined) ?? styles[0] ?? {};
    const xf = this.fromSource && style.sourceXf !== undefined ? style.sourceXf : this.xfIndexOf(style);
    this.xfByStyle.set(styleId, xf);
    return xf;
  }

  private xfIndexOf(style: CellStyle): number {
    const p = this.prefix;
    const theme = this.model.theme;
    const numFmtId = this.numFmtIdFor(style);
    const fontId =
      style.font === undefined ? 0 : this.entryId(this.fontIds, "fonts", styleKey({ font: style.font }), () => `<${p}font>${xlsxFontProperties(style.font!, theme, "name", p)}</${p}font>`);
    const fillId = this.entryId(this.fillIds, "fills", styleKey({ fill: style.fill }), () => fillXml(style.fill, theme, p));
    const borderId = this.entryId(this.borderIds, "borders", styleKey({ borders: style.borders }), () => borderXml(style.borders, theme, p));
    const xfId = style.namedStyle === undefined || style.namedStyle === "Normal" ? 0 : (this.namedXf.get(style.namedStyle) ?? 0);
    const parent = this.parents[xfId] ?? NO_PARENT;
    let attrs = `numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="${xfId}"`;
    if (style.quotePrefix) {
      attrs += ` quotePrefix="1"`;
    }
    if (numFmtId !== parent.numFmtId) {
      attrs += ` applyNumberFormat="1"`;
    }
    if (fontId !== parent.fontId) {
      attrs += ` applyFont="1"`;
    }
    if (fillId !== parent.fillId) {
      attrs += ` applyFill="1"`;
    }
    if (borderId !== parent.borderId) {
      attrs += ` applyBorder="1"`;
    }
    if (styleKey({ alignment: style.alignment }) !== parent.alignment) {
      attrs += ` applyAlignment="1"`;
    }
    if (styleKey({ protection: style.protection }) !== parent.protection) {
      attrs += ` applyProtection="1"`;
    }
    const children = alignmentXml(style.alignment, p) + protectionXml(style.protection, p);
    const xml = children === "" ? `<${p}xf ${attrs}/>` : `<${p}xf ${attrs}>${children}</${p}xf>`;
    const known = this.xfByXml.get(xml);
    if (known !== undefined) {
      return known;
    }
    const index = this.append("cellXfs", xml);
    this.xfByXml.set(xml, index);
    return index;
  }

  /** Whether anything was appended to the part (beyond a from-scratch
   * base's own xf 0). */
  get changed(): boolean {
    return this.appendCount() > this.baseAppends;
  }

  /** The part text: the source verbatim when nothing was appended, else
   * with the appends spliced in (see the module header). */
  xml(): string {
    if (this.appendCount() === 0) {
      return this.text;
    }
    const splices: Splice[] = [];
    const p = this.prefix;
    for (const collection of ["numFmts", "fonts", "fills", "borders", "cellXfs"] as const) {
      const items = this.appends[collection];
      if (items.length === 0) {
        continue;
      }
      const total = String(this.counts[collection] + items.length);
      const span = this.spans.get(collection);
      if (span === undefined) {
        const rank = STYLESHEET_ORDER.indexOf(collection);
        const later = this.order.find((element) => STYLESHEET_ORDER.indexOf(element.local) > rank);
        const at = later?.start ?? this.rootCloseStart;
        splices.push({ start: at, end: at, text: `<${p}${collection} count="${total}">${items.join("")}</${p}${collection}>` });
        continue;
      }
      const tag = setAttribute(startTagAt(this.text, span.start), "count", total);
      if (span.selfClosing) {
        splices.push({ start: span.start, end: span.end, text: `${tag.replace(/\s*\/>$/, ">")}${items.join("")}</${span.name}>` });
      } else {
        splices.push({ start: span.start, end: span.end, text: tag });
        splices.push({ start: span.closeStart, end: span.closeStart, text: items.join("") });
      }
    }
    return applySplices(this.text, splices);
  }
}

/** Records `key → index` unless an earlier entry already holds the key. */
function remember(ids: Map<string, number>, key: string, index: number): void {
  if (!ids.has(key)) {
    ids.set(key, index);
  }
}
