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
import { DEFAULT_THEME, resolveFont, type DocumentTheme, type ModelColor, type ModelFont } from "../../core/theme";
import { parseXml } from "../../core/xml";
import type { CellAlignment, CellBorderEdge, CellBorders, CellFill, CellFont, CellStyle, SheetBorderStyle } from "../model";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "../numberFormat/builtin";
import { defaultStyles, INDEXED_COLORS } from "../styles";
import { elementAttributes, xlsxColorFrom, xlsxColorFromElement, xlsxRgbToHex, type AttributeGetter } from "./themeColor";

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
