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
//   flags are ignored (Excel writes cellXfs fully resolved); `xfId` names the
//   style's `namedStyle` (SPEC-123, below).
// - `<colors><indexedColors>` → the 64-entry override palette.
// - SPEC-118 "Reading xlsx number formats": `numFmts/numFmt@numFmtId` +
//   `@formatCode` (attribute-decoded) are read into a map; each
//   `xf@numFmtId` (absent → 0) resolves: a map entry → `numberFormat` (any
//   id, Excel writes 5–8 and 41–44 explicitly); else a locale-dependent
//   built-in → `builtinFormatId` and the en-US code; else another built-in →
//   its code; 0, unknown and CJK ids → absent. A code equal to `General` is
//   stored absent; `applyNumberFormat` is ignored like every `apply*`.
// - SPEC-123 "Named cell styles" › xlsx: `cellStyles/cellStyle` → one
//   NamedCellStyle per element, in file order — `name`, `builtinId`,
//   `hidden` ("1"/"true"), `sourceXfId` = `xfId`; the first cellStyle per
//   xfId wins (a later one naming the same xfId is skipped), `customBuiltin`
//   is not modelled. Its definition is `cellStyleXfs/xf[xfId]` decoded by the
//   same xf mapping as cellXfs (never quotePrefix): here an ABSENT `applyX`
//   means the aspect is included and "0"/"false" excludes it (numberFormat,
//   font, fill, borders, alignment, protection), `includes` lists the
//   included aspects in STYLE_ASPECTS order and `style` holds those of them
//   with a value; an xfId with no cellStyleXfs entry is an empty definition.
//   A cellXfs style's `namedStyle` is the name of the named style whose
//   `sourceXfId` is its `xfId` — absent when the xf has no xfId, the xfId
//   names no cellStyle, or it names Normal (builtinId 0; the one named
//   "Normal" when no cellStyle has builtinId 0).
// - SPEC-129 "xlsx wire mapping" › dxfs: `dxfs/dxf` → SpreadsheetModel.dxfs
//   by index (readXlsxDxfs), each keeping its `<dxf>` verbatim as
//   `sourceXml` (alignment and protection children are not modelled). A dxf
//   is a PARTIAL style, so a font property is tri-state: `<b/>`, `<b
//   val="1"/>` → `bold: true`, `<b val="0"/>` → `bold: false` (likewise `i`,
//   `strike`), absent → unset; `<u/>` → "single", `<u val="none"/>` → "none";
//   `color` as everywhere. Excel's fill quirk: a `patternFill` with no
//   `patternType`, or `solid`, is a solid fill of `bgColor ?? fgColor` (no
//   colour: no fill); `none` is no fill; other patterns and `gradientFill`
//   as above. `border` edges left/start, right/end, top, bottom (diagonals
//   and inner edges are not modelled). `numFmt@formatCode` → `numberFormat`
//   (a built-in id without a code → the built-in's code).
// tableStyles (SPEC-133), extLst and colors/mruColors stay unread in the
// retained part.
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
//   cellStyleXfs index of `namedStyle`, below), `quotePrefix="1"`, then
//   `applyNumberFormat`/`applyFont`/`applyFill`/`applyBorder`/
//   `applyAlignment`/`applyProtection="1"` where the xf differs from that
//   parent xf; children `<alignment>` and `<protection>` (`locked="0"` only
//   when false, `hidden="1"` only when true). Equal appended xfs are one
//   entry.
// - SPEC-123, the xfId: an absent or "Normal" namedStyle (or one whose
//   definition has builtinId 0) → Normal's cellStyleXfs index (the xfId of
//   the source's builtinId-0 cellStyle, else of its "Normal" one, else 0).
//   Otherwise the workbook's definition (styles.ts workbookNamedStyle): its
//   `sourceXfId` when the source part holds that cellStyleXfs entry; else
//   the source cellStyle of its name; else — a built-in first applied by
//   the user, any definition in a from-scratch save, or a built-in the
//   model does not list yet (its pinned definition) — one appended
//   cellStyleXfs `xf` plus a `cellStyle` under the definition's own name (a
//   built-in a file names in its language, "Gut", keeps it), once per
//   name. The style xf's
//   numFmt, font, fill and border reuse-or-append as a cellXfs entry's do
//   (the font is the definition's keys over Normal's — `styles[0].font`, or
//   a from-scratch base's font 0 when that is empty — so it is complete);
//   an excluded aspect writes Normal's id (numFmt 0, font 0, the first none
//   fill, the first empty border) and `applyX="0"`, an included one no
//   flag; alignment/protection children are written as cellXfs writes them.
//   The cellStyle is `<cellStyle name xfId>` + `builtinId` (BUILTIN_CELL_
//   STYLES' id for one of its names, else the definition's own, read from a
//   file) + `hidden="1"`, never `customBuiltin`. A source without
//   cellStyleXfs entries gains Normal's `<xf numFmtId="0" fontId="0"
//   fillId="0" borderId="0"/>` first, one without cellStyles `<cellStyle
//   name="Normal" xfId="0" builtinId="0"/>` first; an index past the part's
//   entries that a source cellStyle or cellXfs xf already names is filled
//   with an every-aspect-excluded xf before an append lands on it (neither
//   changes meaning); every appended style xf is a parent the apply flags
//   compare against. A name with no definition and no source cellStyle →
//   Normal's.
// - XML: fonts in Excel's child order (b, i, strike, u, vertAlign, sz, color,
//   name, family, scheme — a theme font writes its theme typeface, family 2
//   and `<scheme val>`); solid fills `<patternFill patternType="solid">
//   <fgColor …/><bgColor indexed="64"/>`; border edges in the order left,
//   right, top, bottom, diagonal, an automatic colour as `indexed="64"`;
//   every colour through themeColor.ts's modelColorToXlsxAttrs.
// - SPEC-129, dxfs: `model.dxfs` is append-only, so the source's `<dxf>`s
//   are never rewritten and the entries past the source's count are
//   appended — every one up to the highest dxfId a serialised base rule
//   wrote (dxfFor), so an untouched workbook whose x14 rules interned their
//   inline dxfs at read stays byte-identical (dxfXml, children in CT_Dxf
//   order font, numFmt, fill, border): a font's
//   set properties (`<b val="0"/>` for an explicit false, `<u val="none"/>`),
//   a solid fill as Excel writes a dxf's — `<patternFill><bgColor/>` with no
//   patternType — other fills as cellXfs writes them, the set border edges;
//   a numFmt's id is the one `<numFmts>` gives its code, else the next free
//   id ≥ 164 (shared with the numFmts appends; a dxf's numFmt is not added to
//   `<numFmts>`). A from-scratch part starts with `<dxfs count="0"/>`.
// - Appends are spliced before each collection's close tag with its `count`
//   set to the new total; an absent collection is inserted in CT_Stylesheet
//   order (numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, cellStyles,
//   dxfs, tableStyles, colors, extLst); the part is rewritten only when
//   something was appended — every source entry (numFmts, fonts, fills,
//   borders, cellStyleXfs, cellXfs, cellStyles, dxfs: each keeps its index)
//   and every other element stays byte-verbatim.
// - From scratch (no source part) the base is a minimal styles part whose
//   font 0 is built from `model.styles[0]` and whose xf 0 IS style 0, so
//   unstyled cells keep the workbook's own default font in either direction;
//   a blank workbook's base is SPEC-121's FROM_SCRATCH_STYLES_XML. The base
//   is this module's own text, so its tables are seeded from what it wrote
//   (fromScratchTables) rather than parsed: the from-scratch path never
//   needs DOMParser, and runs as written in plain Node (SPEC-140's
//   generation API, which has no polyfill on its write path).
import { DEFAULT_THEME, isThemeFont, resolveFont, type DocumentTheme, type ModelColor, type ModelFont } from "../../core/theme";
import { parseXml } from "../../core/xml";
import { scanAttr, scanXml, XmlScanError } from "../../core/xmlScanner";
import { DXF_BORDER_EDGES, type Dxf, type DxfFont } from "../conditionalFormats";
import type { CellAlignment, CellBorderEdge, CellBorders, CellFill, CellFont, CellStyle, SheetBorderStyle, SpreadsheetModel } from "../model";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "../numberFormat/builtin";
import { builtinCellStyle, defaultStyles, INDEXED_COLORS, styleKey, workbookNamedStyle, type NamedCellStyle, type StyleAspect } from "../styles";
import { applySplices, escapeAttributeValue, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { childNodes, elementTree, nodeAttr, type XmlNode } from "./elementTree";
import { elementAttributes, modelColorToXlsxAttrs, scannedAttributes, xlsxColorFrom, xlsxColorFromElement, xlsxRgbToHex, type AttributeGetter } from "./themeColor";

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
  /** SPEC-123: the named cell styles, one per `cellStyles/cellStyle` (the
   * first per xfId) in file order — absent when the part names none. */
  namedStyles?: NamedCellStyle[];
  /** SPEC-129: the differential styles, one per `dxfs/dxf` in order. */
  dxfs: Dxf[];
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

/** The records an xf points into, parsed once per part. */
interface XfRecords {
  fonts: readonly CellFont[];
  fills: readonly (CellFill | undefined)[];
  borders: readonly (CellBorders | undefined)[];
  numFmts: ReadonlyMap<number, string>;
}

/** What an xf formats — its font, fill and border (copies), its alignment
 * and protection children and its number format — decoded alike for a
 * cellXfs and a cellStyleXfs entry. */
function readXfFormatting(xf: Element, { fonts, fills, borders, numFmts }: XfRecords): CellStyle {
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
  return style;
}

/** One cellXfs entry: its formatting, `quotePrefix`, the named style its
 * `xfId` names (`names`: xfId → name, Normal left out) and its index. */
function readXf(xf: Element, index: number, records: XfRecords, names: ReadonlyMap<number, string>): CellStyle {
  const style = readXfFormatting(xf, records);
  if (isOn(xf.getAttribute("quotePrefix"))) {
    style.quotePrefix = true;
  }
  const xfId = parseIndex(xf.getAttribute("xfId"));
  const named = xfId !== undefined ? names.get(xfId) : undefined;
  if (named !== undefined) {
    style.namedStyle = named;
  }
  style.sourceXf = index;
  return style;
}

// --- named cell styles (SPEC-123) ----------------------------------------------

/** Each aspect's `applyX` flag, in STYLE_ASPECTS order. */
const APPLY_FLAGS: readonly (readonly [StyleAspect, string])[] = [
  ["numberFormat", "applyNumberFormat"],
  ["font", "applyFont"],
  ["fill", "applyFill"],
  ["borders", "applyBorder"],
  ["alignment", "applyAlignment"],
  ["protection", "applyProtection"],
];

/** The CellStyle fields each aspect covers. */
const ASPECT_FIELDS: Readonly<Record<StyleAspect, readonly (keyof CellStyle)[]>> = {
  numberFormat: ["numberFormat", "builtinFormatId"],
  font: ["font"],
  fill: ["fill"],
  borders: ["borders"],
  alignment: ["alignment"],
  protection: ["protection"],
};

/** A cellStyleXfs entry as a named style's definition: an absent `applyX`
 * includes the aspect, "0"/"false" excludes it; `style` keeps the included
 * aspects that have a value. */
function readStyleXf(xf: Element, records: XfRecords): Pick<NamedCellStyle, "style" | "includes"> {
  const decoded = readXfFormatting(xf, records);
  const style: CellStyle = {};
  const includes: StyleAspect[] = [];
  const keep = <K extends keyof CellStyle>(field: K): void => {
    if (decoded[field] !== undefined) {
      style[field] = decoded[field];
    }
  };
  for (const [aspect, flag] of APPLY_FLAGS) {
    const value = xf.getAttribute(flag);
    if (value === "0" || value === "false") {
      continue;
    }
    includes.push(aspect);
    ASPECT_FIELDS[aspect].forEach(keep);
  }
  return { style, includes };
}

/** `cellStyles/cellStyle` → the named styles, in file order, the first per
 * xfId winning; each defined by its cellStyleXfs entry (an empty
 * definition when the part has none at that index). A cellStyle without a
 * name or an xfId names nothing a cell could use, and is skipped. */
function readNamedStyles(root: Element, records: XfRecords): NamedCellStyle[] {
  const styleXfsElement = childNamed(root, "cellStyleXfs");
  const styleXfs = styleXfsElement ? childrenNamed(styleXfsElement, "xf") : [];
  const cellStyles = childNamed(root, "cellStyles");
  const named: NamedCellStyle[] = [];
  const claimed = new Set<number>();
  for (const element of cellStyles ? childrenNamed(cellStyles, "cellStyle") : []) {
    const name = element.getAttribute("name");
    const xfId = parseIndex(element.getAttribute("xfId"));
    if (name === null || xfId === undefined || claimed.has(xfId)) {
      continue;
    }
    claimed.add(xfId);
    const xf = styleXfs[xfId];
    const { style, includes } = xf ? readStyleXf(xf, records) : { style: {}, includes: [] };
    const builtinId = parseIndex(element.getAttribute("builtinId"));
    named.push({
      name,
      ...(builtinId !== undefined ? { builtinId } : {}),
      style,
      includes,
      ...(isOn(element.getAttribute("hidden")) ? { hidden: true as const } : {}),
      sourceXfId: xfId,
    });
  }
  return named;
}

/** xfId → the name a cellXfs xf pointing there gets as its `namedStyle`;
 * Normal (builtinId 0, or — when none has builtinId 0 — the style named
 * "Normal") is the absent namedStyle and left out. */
function namedStyleNames(named: readonly NamedCellStyle[]): Map<number, string> {
  const builtinNormal = named.some((entry) => entry.builtinId === 0);
  const names = new Map<number, string>();
  for (const entry of named) {
    const normal = entry.builtinId === 0 || (!builtinNormal && entry.name === "Normal");
    if (!normal && entry.sourceXfId !== undefined) {
      names.set(entry.sourceXfId, entry.name);
    }
  }
  return names;
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

// --- differential styles (SPEC-129) ---------------------------------------------

const inMain = (ns: string | undefined): boolean => isSpreadsheetMl(ns);

/** A CT_BooleanProperty's value: absent `val` (or anything but an explicit
 * false) is true. */
function booleanProperty(node: XmlNode): boolean {
  const val = nodeAttr(node, "val");
  return !(val === "0" || val === "false");
}

function nodeColor(node: XmlNode | undefined, indexedColors: readonly string[] | undefined): ModelColor | undefined {
  return node ? xlsxColorFrom((name) => nodeAttr(node, name), indexedColors) : undefined;
}

function firstChild(node: XmlNode | undefined, local: string): XmlNode | undefined {
  return node ? childNodes(node, local, inMain)[0] : undefined;
}

/** A dxf `<font>`: the tri-state flags, the underline and the colour. */
function readDxfFont(node: XmlNode, indexedColors: readonly string[] | undefined): DxfFont | undefined {
  const font: DxfFont = {};
  for (const child of node.children) {
    if (!inMain(child.ns)) {
      continue;
    }
    switch (child.local) {
      case "b":
        font.bold = booleanProperty(child);
        break;
      case "i":
        font.italic = booleanProperty(child);
        break;
      case "strike":
        font.strike = booleanProperty(child);
        break;
      case "u": {
        const val = nodeAttr(child, "val");
        if (val === undefined) {
          font.underline = "single";
        } else if (val === "none" || UNDERLINES.has(val)) {
          font.underline = val as DxfFont["underline"];
        }
        break;
      }
      case "color": {
        const color = nodeColor(child, indexedColors);
        if (color !== undefined) {
          font.color = color;
        }
        break;
      }
      default:
        break;
    }
  }
  return Object.keys(font).length > 0 ? font : undefined;
}

/** A dxf `<fill>` (see the module header: Excel's solid-fill quirk). */
function readDxfFill(node: XmlNode, indexedColors: readonly string[] | undefined): CellFill | undefined {
  const pattern = firstChild(node, "patternFill");
  if (pattern) {
    const type = nodeAttr(pattern, "patternType");
    const fg = nodeColor(firstChild(pattern, "fgColor"), indexedColors);
    const bg = nodeColor(firstChild(pattern, "bgColor"), indexedColors);
    if (type === undefined || type === "" || type === "solid") {
      const color = bg ?? fg;
      return color === undefined ? undefined : { kind: "solid", color };
    }
    if (type === "none") {
      return undefined;
    }
    const fill: PatternFill = { kind: "pattern", pattern: type };
    if (fg !== undefined) {
      fill.fg = fg;
    }
    if (bg !== undefined) {
      fill.bg = bg;
    }
    return fill;
  }
  const gradient = firstChild(node, "gradientFill");
  if (!gradient) {
    return undefined;
  }
  const fill: GradientFill = { kind: "gradient", type: nodeAttr(gradient, "type") === "path" ? "path" : "linear", stops: [] };
  const degree = parseNumber(nodeAttr(gradient, "degree"));
  if (degree !== undefined) {
    fill.degree = degree;
  }
  for (const side of GRADIENT_FOCUS) {
    const value = parseNumber(nodeAttr(gradient, side));
    if (value !== undefined) {
      fill[side] = value;
    }
  }
  for (const stop of childNodes(gradient, "stop", inMain)) {
    fill.stops.push({ position: parseNumber(nodeAttr(stop, "position")) ?? 0, color: nodeColor(firstChild(stop, "color"), indexedColors) ?? "#000000" });
  }
  return fill;
}

/** A dxf `<border>`: its four outer edges. */
function readDxfBorders(node: XmlNode, indexedColors: readonly string[] | undefined): Dxf["borders"] {
  const borders: NonNullable<Dxf["borders"]> = {};
  for (const child of node.children) {
    const side = inMain(child.ns) ? BORDER_EDGES.get(child.local) : undefined;
    const style = nodeAttr(child, "style");
    if (side === undefined || side === "diagonal" || style === undefined || !BORDER_STYLES.has(style) || borders[side] !== undefined) {
      continue;
    }
    const edge: CellBorderEdge = { style: style as SheetBorderStyle };
    const color = nodeColor(firstChild(child, "color"), indexedColors);
    if (color !== undefined) {
      edge.color = color;
    }
    borders[side] = edge;
  }
  return Object.keys(borders).length > 0 ? borders : undefined;
}

/** One `<dxf>` (styles.xml) or `<x14:dxf>` (an x14 rule's inline one) →
 * the model's Dxf, without `sourceXml` (see the module header). */
export function readDxfNode(node: XmlNode, indexedColors?: readonly string[]): Dxf {
  const dxf: Dxf = {};
  const font = firstChild(node, "font");
  const fontValue = font ? readDxfFont(font, indexedColors) : undefined;
  if (fontValue !== undefined) {
    dxf.font = fontValue;
  }
  const fill = firstChild(node, "fill");
  const fillValue = fill ? readDxfFill(fill, indexedColors) : undefined;
  if (fillValue !== undefined) {
    dxf.fill = fillValue;
  }
  const border = firstChild(node, "border");
  const borders = border ? readDxfBorders(border, indexedColors) : undefined;
  if (borders !== undefined) {
    dxf.borders = borders;
  }
  const numFmt = firstChild(node, "numFmt");
  if (numFmt) {
    const code = nodeAttr(numFmt, "formatCode") ?? BUILTIN_NUMBER_FORMATS[parseIndex(nodeAttr(numFmt, "numFmtId")) ?? -1];
    if (code !== undefined) {
      dxf.numberFormat = code;
    }
  }
  return dxf;
}

/** `xl/styles.xml`'s `<dxfs>` → the model's differential styles, by index,
 * each with its `<dxf>` verbatim as `sourceXml` (see the module header); []
 * for a part without dxfs or one that is not well-formed. The part's own
 * `<indexedColors>` palette applies. */
export function readXlsxDxfs(xml: string | undefined): Dxf[] {
  const root = xml === undefined ? undefined : elementTree(xml)?.[0];
  if (xml === undefined || !root || root.local !== "styleSheet" || !inMain(root.ns)) {
    return [];
  }
  const palette = firstChild(firstChild(root, "colors"), "indexedColors");
  const indexed = palette
    ? childNodes(palette, "rgbColor", inMain)
        .slice(0, 64)
        .map((entry, index) => xlsxRgbToHex(nodeAttr(entry, "rgb") ?? "") ?? INDEXED_COLORS[index]!)
    : [];
  const indexedColors = indexed.length > 0 ? indexed : undefined;
  const dxfs = firstChild(root, "dxfs");
  return (dxfs ? childNodes(dxfs, "dxf", inMain) : []).map((node) => ({ ...readDxfNode(node, indexedColors), sourceXml: xml.slice(node.start, node.end) }));
}

/** A dxf's children as SpreadsheetML (see the module header), in CT_Dxf
 * order; `numFmtId` names the id a format code is written with. */
export function dxfXml(dxf: Dxf, theme: DocumentTheme | undefined, p: string, numFmtId: (code: string) => number): string {
  let out = "";
  const font = dxf.font;
  if (font !== undefined && Object.values(font).some((value) => value !== undefined)) {
    const flag = (local: string, value: boolean | undefined): string => (value === undefined ? "" : value ? `<${p}${local}/>` : `<${p}${local} val="0"/>`);
    let body = flag("b", font.bold) + flag("i", font.italic) + flag("strike", font.strike);
    if (font.underline !== undefined) {
      body += font.underline === "single" ? `<${p}u/>` : `<${p}u val="${font.underline}"/>`;
    }
    if (font.color !== undefined) {
      body += colorElement("color", font.color, theme, p);
    }
    out += `<${p}font>${body}</${p}font>`;
  }
  if (dxf.numberFormat !== undefined) {
    out += `<${p}numFmt numFmtId="${numFmtId(dxf.numberFormat)}" formatCode="${escapeAttributeValue(dxf.numberFormat)}"/>`;
  }
  if (dxf.fill !== undefined) {
    out +=
      dxf.fill.kind === "solid"
        ? `<${p}fill><${p}patternFill>${colorElement("bgColor", dxf.fill.color, theme, p)}</${p}patternFill></${p}fill>`
        : fillXml(dxf.fill, theme, p);
  }
  const borders = dxf.borders;
  if (borders !== undefined && DXF_BORDER_EDGES.some((edge) => borders[edge] !== undefined)) {
    const edges = (["left", "right", "top", "bottom"] as const)
      .map((side) => {
        const edge = borders[side];
        if (edge === undefined) {
          return "";
        }
        return edge.color === undefined
          ? `<${p}${side} style="${edge.style}"/>`
          : `<${p}${side} style="${edge.style}">${colorElement("color", edge.color, theme, p)}</${p}${side}>`;
      })
      .join("");
    out += `<${p}border>${edges}</${p}border>`;
  }
  return out;
}

function defaultTable(): XlsxStyleTable {
  return { styles: defaultStyles(), defaultFont: { ...DEFAULT_XLSX_FONT }, dxfs: [] };
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
  const records: XfRecords = { fonts, fills, borders, numFmts: readNumFmts(root) };
  const namedStyles = readNamedStyles(root, records);
  const names = namedStyleNames(namedStyles);
  const styles = xfs.map((xf, index) => readXf(xf, index, records, names));

  // The Normal style's font: cellXfs[0]'s (fonts[0] when there is no xf).
  const normalFont = fonts[xfs[0] ? (parseIndex(xfs[0].getAttribute("fontId")) ?? 0) : 0];
  const family = normalFont?.family !== undefined ? resolveFont(normalFont.family, theme ?? DEFAULT_THEME) : DEFAULT_XLSX_FONT.family;
  const table: XlsxStyleTable = {
    styles: styles.length > 0 ? styles : defaultStyles(),
    defaultFont: { family, sizePt: normalFont?.sizePt ?? DEFAULT_XLSX_FONT.sizePt },
    dxfs: readXlsxDxfs(xml),
  };
  if (indexedColors !== undefined) {
    table.indexedColors = indexedColors;
  }
  if (namedStyles.length > 0) {
    table.namedStyles = namedStyles;
  }
  return table;
}

/** The `<colors><indexedColors>` palette readXlsxStyles reads, read with the
 * scanner instead of DOMParser — what SPEC-137's phase 1 resolves
 * shared-string run colours against (it runs where DOMParser does not). The
 * same mapping (readIndexedColors): the root a SpreadsheetML `styleSheet`,
 * its first `colors` child, that one's first `indexedColors` child, up to 64
 * of its `rgbColor@rgb` children, an unreadable entry the legacy colour at
 * its position; undefined for no part, a part the scanner rejects (the whole
 * part is scanned, as DOMParser reads it whole), another root, or no
 * palette. A part only one of the two parsers accepts may still read
 * differently — the shell phase compares the two (xlsx/phases.ts). */
export function readXlsxIndexedColors(xml: string | undefined): string[] | undefined {
  if (xml === undefined) {
    return undefined;
  }
  const colors: string[] = [];
  let depth = 0;
  // 0: before the first <colors>; 1: inside it; 2: inside its first
  // <indexedColors>; 3: past them.
  let state = 0;
  try {
    let styleSheet = false;
    scanXml(xml, (event) => {
      if (event.type === "open") {
        const level = depth + 1;
        if (!event.selfClosing) {
          depth = level;
        }
        if (level === 1) {
          styleSheet = event.local === "styleSheet" && isSpreadsheetMl(event.ns);
          return styleSheet ? undefined : "stop";
        }
        if (!isSpreadsheetMl(event.ns)) {
          return undefined;
        }
        if (level === 2 && state === 0 && event.local === "colors") {
          state = event.selfClosing ? 3 : 1;
        } else if (level === 3 && state === 1 && event.local === "indexedColors") {
          state = event.selfClosing ? 3 : 2;
        } else if (level === 4 && state === 2 && event.local === "rgbColor" && colors.length < 64) {
          colors.push(xlsxRgbToHex(scanAttr(event.attrs, "rgb") ?? "") ?? INDEXED_COLORS[colors.length]!);
        }
      } else if (event.type === "close") {
        if ((depth === 3 && state === 2) || (depth === 2 && state === 1)) {
          state = 3;
        }
        depth -= 1;
      }
      return undefined;
    });
    if (!styleSheet) {
      return undefined;
    }
  } catch (error) {
    if (error instanceof XmlScanError) {
      return undefined;
    }
    throw error;
  }
  return colors.length > 0 ? colors : undefined;
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
type StyleCollection = "numFmts" | "fonts" | "fills" | "borders" | "cellStyleXfs" | "cellXfs" | "cellStyles" | "dxfs";

/** Those collections, in CT_Stylesheet order. */
const STYLE_COLLECTIONS: readonly StyleCollection[] = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs"];

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

/** A cellStyleXfs xf's flags excluding every aspect. */
const EXCLUDE_EVERY_ASPECT = APPLY_FLAGS.map(([, flag]) => ` ${flag}="0"`).join("");

/** The default font a from-scratch font 0 falls back to: the blank
 * workbook's (SPEC-121 "The blank workbook"). */
const BLANK_DEFAULT_FONT: CellFont = { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } };

/** The font a from-scratch base's font 0 is built from. */
function fromScratchFont(defaultStyle: CellStyle | undefined): CellFont {
  return defaultStyle?.font !== undefined && Object.keys(defaultStyle.font).length > 0 ? defaultStyle.font : BLANK_DEFAULT_FONT;
}

/** A from-scratch styles part before any xf: font 0 from the default style,
 * the two reserved fills, one empty border, the Normal style's xf. */
function fromScratchBase(defaultStyle: CellStyle | undefined, theme: DocumentTheme | undefined): string {
  const font = fromScratchFont(defaultStyle);
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

/** A `<font>` element's font, read with the scanner through the same
 * XlsxFontReader readFont uses (its SpreadsheetML children, in order). */
function scannedFont(xml: string): CellFont {
  const reader = new XlsxFontReader();
  let depth = 0;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      if (depth === 1 && isSpreadsheetMl(event.ns)) {
        reader.add(event.local, scannedAttributes(event.attrs));
      }
      if (!event.selfClosing) {
        depth += 1;
      }
    } else if (event.type === "close") {
      depth -= 1;
    }
  });
  return reader.build();
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
  private readonly counts: Record<StyleCollection, number> = { numFmts: 0, fonts: 0, fills: 0, borders: 0, cellStyleXfs: 0, cellXfs: 0, cellStyles: 0, dxfs: 0 };
  private readonly appends: Record<StyleCollection, string[]> = {
    numFmts: [],
    fonts: [],
    fills: [],
    borders: [],
    cellStyleXfs: [],
    cellXfs: [],
    cellStyles: [],
    dxfs: [],
  };
  /** SPEC-129: the ids given to dxf number formats `<numFmts>` lacks. */
  private readonly dxfNumFmts = new Map<string, number>();
  /** SPEC-129: the highest model dxfId a serialised rule wrote, plus one. */
  private dxfsUsed = 0;
  /** styleKey → index, per collection (the first source entry, then appends). */
  private readonly fontIds = new Map<string, number>();
  private readonly fillIds = new Map<string, number>();
  private readonly borderIds = new Map<string, number>();
  private readonly numFmtCodes = new Map<number, string>();
  private readonly numFmtByCode = new Map<string, number>();
  private highestNumFmt = 0;
  /** cellStyleXfs index → the entry, source and appended. */
  private readonly parents: ParentXf[] = [];
  /** cellStyle name → its xfId: the first source cellStyle of the name, then
   * the named styles appended this save. */
  private readonly namedXf = new Map<string, number>();
  /** Every cellStyleXfs index a source cellStyle or cellXfs xf names — the
   * ones past the part's entries an append must not land on. */
  private readonly sourceXfIds = new Set<number>();
  /** Normal's cellStyleXfs index. */
  private normalXfId = 0;
  private readonly xfByXml = new Map<string, number>();
  private readonly xfByStyle = new Map<number, number>();
  private readonly baseAppends: number;

  /** `source` is the retained `xl/styles.xml` text — undefined builds the
   * from-scratch base from `model.styles[0]` (xf 0 is style 0). */
  constructor(
    source: string | undefined,
    private readonly model: Pick<SpreadsheetModel, "styles" | "theme" | "indexedColors" | "namedStyles"> & Partial<Pick<SpreadsheetModel, "dxfs">>,
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
    if (this.fromSource) {
      this.readSource();
    } else {
      this.seedFromScratchBase(fromScratchFont(model.styles[0]));
    }
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
    const cellXfsElement = childNamed(root, "cellXfs");
    const cellXfs = cellXfsElement ? childrenNamed(cellXfsElement, "xf") : [];
    this.counts.cellXfs = cellXfs.length;
    for (const xf of cellXfs) {
      const xfId = parseIndex(xf.getAttribute("xfId"));
      if (xfId !== undefined) {
        this.sourceXfIds.add(xfId);
      }
    }
    const styleXfsElement = childNamed(root, "cellStyleXfs");
    const styleXfs = styleXfsElement ? childrenNamed(styleXfsElement, "xf") : [];
    this.counts.cellStyleXfs = styleXfs.length;
    for (const xf of styleXfs) {
      this.parents.push({
        numFmtId: intAttr(xf, "numFmtId"),
        fontId: intAttr(xf, "fontId"),
        fillId: intAttr(xf, "fillId"),
        borderId: intAttr(xf, "borderId"),
        alignment: styleKey({ alignment: readAlignment(childNamed(xf, "alignment")) }),
        protection: styleKey({ protection: readProtection(childNamed(xf, "protection")) }),
      });
    }
    const cellStylesElement = childNamed(root, "cellStyles");
    const cellStyles = cellStylesElement ? childrenNamed(cellStylesElement, "cellStyle") : [];
    this.counts.cellStyles = cellStyles.length;
    // SPEC-129: the source's dxfs keep their indices; their numFmt ids are
    // taken, like numFmts'.
    const dxfsElement = childNamed(root, "dxfs");
    const dxfs = dxfsElement ? childrenNamed(dxfsElement, "dxf") : [];
    this.counts.dxfs = dxfs.length;
    for (const dxf of dxfs) {
      this.highestNumFmt = Math.max(this.highestNumFmt, parseIndex(childNamed(dxf, "numFmt")?.getAttribute("numFmtId")) ?? 0);
    }
    let builtinNormal: number | undefined;
    for (const cellStyle of cellStyles) {
      const name = cellStyle.getAttribute("name");
      if (name !== null && !this.namedXf.has(name)) {
        this.namedXf.set(name, intAttr(cellStyle, "xfId"));
      }
      this.sourceXfIds.add(intAttr(cellStyle, "xfId"));
      if (builtinNormal === undefined && parseIndex(cellStyle.getAttribute("builtinId")) === 0) {
        builtinNormal = intAttr(cellStyle, "xfId");
      }
    }
    this.normalXfId = builtinNormal ?? this.namedXf.get("Normal") ?? 0;
  }

  /** The from-scratch base's tables (fromScratchBase), as readSource would
   * read them back — seeded from what the base wrote instead of parsed, so
   * the from-scratch path never touches DOMParser (see the module header).
   * Font 0 is keyed as its XML reads back, through the reader's own
   * XlsxFontReader over the scanner. */
  private seedFromScratchBase(font: CellFont): void {
    this.counts.fonts = 1;
    remember(this.fontIds, styleKey({ font: scannedFont(`<font xmlns="${MAIN_NS}">${xlsxFontProperties(font, this.model.theme)}</font>`) }), 0);
    this.counts.fills = 2;
    remember(this.fillIds, styleKey({}), 0);
    remember(this.fillIds, styleKey({ fill: { kind: "pattern", pattern: "gray125" } }), 1);
    this.counts.borders = 1;
    remember(this.borderIds, styleKey({}), 0);
    this.counts.cellStyleXfs = 1;
    this.parents.push({ ...NO_PARENT });
    this.counts.cellStyles = 1;
    this.namedXf.set("Normal", 0);
    this.sourceXfIds.add(0);
    this.normalXfId = 0;
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

  /** A font's fonts index: font 0 when absent, else the first equal entry,
   * else an append. */
  private fontIdFor(font: CellFont | undefined): number {
    const p = this.prefix;
    return font === undefined ? 0 : this.entryId(this.fontIds, "fonts", styleKey({ font }), () => `<${p}font>${xlsxFontProperties(font, this.model.theme, "name", p)}</${p}font>`);
  }

  /** A fill's fills index (none: the first entry that reads as none). */
  private fillIdFor(fill: CellFill | undefined): number {
    return this.entryId(this.fillIds, "fills", styleKey({ fill }), () => fillXml(fill, this.model.theme, this.prefix));
  }

  /** A border's borders index (none: the first entry that reads as empty). */
  private borderIdFor(borders: CellBorders | undefined): number {
    return this.entryId(this.borderIds, "borders", styleKey({ borders }), () => borderXml(borders, this.model.theme, this.prefix));
  }

  /** The cellStyleXfs index a cellXfs xf names for `namedStyle` (see the
   * module header). */
  private styleXfFor(namedStyle: string | undefined): number {
    if (namedStyle === undefined || namedStyle === "Normal") {
      return this.normalXfId;
    }
    const own = workbookNamedStyle(this.model, namedStyle);
    if (own?.builtinId === 0) {
      return this.normalXfId;
    }
    if (this.fromSource && own?.sourceXfId !== undefined && own.sourceXfId < this.counts.cellStyleXfs) {
      return own.sourceXfId;
    }
    // The definition's own name: a built-in the workbook names in its
    // language ("Gut" for Good) keeps it.
    const name = own?.name ?? namedStyle;
    const known = this.namedXf.get(name);
    if (known !== undefined) {
      return known;
    }
    const definition = own ?? builtinCellStyle(namedStyle);
    return definition === undefined ? this.normalXfId : this.appendNamedStyle(name, definition);
  }

  /** Appends a named style the part lacks: its cellStyleXfs xf (a parent
   * from now on) and its cellStyle; the xf's index. */
  private appendNamedStyle(name: string, definition: Pick<NamedCellStyle, "builtinId" | "style" | "includes" | "hidden">): number {
    const p = this.prefix;
    const { style, includes } = definition;
    const has = (aspect: StyleAspect): boolean => includes.includes(aspect);
    const numFmtId = has("numberFormat") ? this.numFmtIdFor(style) : 0;
    // A complete font: the definition's keys over Normal's — font 0's, which
    // a from-scratch base built from the blank workbook's when style 0 has none.
    const ownFont = this.model.styles[0]?.font;
    const normalFont = ownFont !== undefined && Object.keys(ownFont).length > 0 ? ownFont : this.fromSource ? undefined : BLANK_DEFAULT_FONT;
    const font: CellFont = { ...normalFont, ...style.font };
    const fontId = has("font") && Object.keys(font).length > 0 ? this.fontIdFor(font) : 0;
    const fillId = this.fillIdFor(has("fill") ? style.fill : undefined);
    const borderId = this.borderIdFor(has("borders") ? style.borders : undefined);
    const alignment = has("alignment") ? style.alignment : undefined;
    const protection = has("protection") ? style.protection : undefined;
    let attrs = `numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"`;
    for (const [aspect, flag] of APPLY_FLAGS) {
      if (!has(aspect)) {
        attrs += ` ${flag}="0"`;
      }
    }
    const children = alignmentXml(alignment, p) + protectionXml(protection, p);
    if (this.counts.cellStyleXfs + this.appends.cellStyleXfs.length === 0) {
      // Normal's xf first: every xfId="0" of the part already means Normal.
      this.parents[this.append("cellStyleXfs", `<${p}xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`)] = NO_PARENT;
    }
    // An index a source cellStyle or cellXfs xf names without an entry gets
    // an empty definition (every aspect excluded): neither changes meaning,
    // and no earlier cellStyle on the xfId shadows the appended one.
    while (this.sourceXfIds.has(this.counts.cellStyleXfs + this.appends.cellStyleXfs.length)) {
      this.parents[this.append("cellStyleXfs", `<${p}xf numFmtId="0" fontId="0" fillId="0" borderId="0"${EXCLUDE_EVERY_ASPECT}/>`)] = NO_PARENT;
    }
    const index = this.append("cellStyleXfs", children === "" ? `<${p}xf ${attrs}/>` : `<${p}xf ${attrs}>${children}</${p}xf>`);
    this.parents[index] = { numFmtId, fontId, fillId, borderId, alignment: styleKey({ alignment }), protection: styleKey({ protection }) };
    if (this.counts.cellStyles + this.appends.cellStyles.length === 0) {
      this.append("cellStyles", `<${p}cellStyle name="Normal" xfId="${this.normalXfId}" builtinId="0"/>`);
    }
    const builtinId = builtinCellStyle(name)?.builtinId ?? definition.builtinId;
    const hidden = definition.hidden ? ` hidden="1"` : "";
    this.append("cellStyles", `<${p}cellStyle name="${escapeAttributeValue(name)}" xfId="${index}"${builtinId !== undefined ? ` builtinId="${builtinId}"` : ""}${hidden}/>`);
    this.namedXf.set(name, index);
    return index;
  }

  private xfIndexOf(style: CellStyle): number {
    const p = this.prefix;
    // The parent first: a named style the part lacks appends its records
    // before the xf's own.
    const xfId = this.styleXfFor(style.namedStyle);
    const numFmtId = this.numFmtIdFor(style);
    const fontId = this.fontIdFor(style.font);
    const fillId = this.fillIdFor(style.fill);
    const borderId = this.borderIdFor(style.borders);
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
    return this.appendCount() > this.baseAppends || this.appendedDxfs().length > this.appends.dxfs.length;
  }

  /** SPEC-129: the id a dxf's number format is written with — the one
   * `<numFmts>` gives the code, else the next free id ≥ 164 (once per code
   * per save; never added to `<numFmts>`). */
  dxfNumFmtId(code: string): number {
    const known = this.numFmtByCode.get(code) ?? this.dxfNumFmts.get(code);
    if (known !== undefined) {
      return known;
    }
    const id = Math.max(163, this.highestNumFmt) + 1;
    this.highestNumFmt = id;
    this.dxfNumFmts.set(code, id);
    return id;
  }

  /** SPEC-129: the dxfId a serialised base rule writes for the model's dxf
   * `id` — the same index, the table being append-only — recording that the
   * part must hold it (see the module header). */
  dxfFor(id: number): number {
    this.dxfsUsed = Math.max(this.dxfsUsed, id + 1);
    return id;
  }

  /** The model's dxfs past the source part's, up to the last one a rule
   * wrote (dxfFor). */
  private appendedDxfs(): readonly Dxf[] {
    const dxfs = this.model.dxfs ?? [];
    return dxfs.slice(this.counts.dxfs, Math.min(dxfs.length, this.dxfsUsed));
  }

  /** The part text: the source verbatim when nothing was appended, else
   * with the appends spliced in (see the module header). */
  xml(): string {
    const p0 = this.prefix;
    for (const dxf of this.appendedDxfs().slice(this.appends.dxfs.length)) {
      this.append("dxfs", `<${p0}dxf>${dxfXml(dxf, this.model.theme, p0, (code) => this.dxfNumFmtId(code))}</${p0}dxf>`);
    }
    if (this.appendCount() === 0) {
      return this.text;
    }
    const splices: Splice[] = [];
    const p = this.prefix;
    for (const collection of STYLE_COLLECTIONS) {
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
