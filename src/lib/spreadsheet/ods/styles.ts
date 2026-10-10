// ODS styles (SPEC-117 "Reading ODS styles", "ODS borders", "Colours"): the
// cell styles an .ods workbook's cells, rows, columns and covered cells name,
// plus the column, row and table styles its layout reads (tableXml.ts, the
// reader). Sources: styles.xml (`office:font-face-decls`, and in
// `office:styles` the `style:default-style` per family and the common
// styles) and content.xml (`office:font-face-decls` and
// `office:automatic-styles`, which the reader slices out of its one
// forward-only scan — both are small next to the tables — and hands here
// to parse).
//
// - Flattening. A style's properties are its family's `style:default-style`,
//   then each `style:parent-style-name` ancestor from the root down, then
//   its own — later definitions winning PER PROPERTY. An `fo:border`
//   shorthand counts as its four sides, `fo:margin` as `fo:margin-left`, and
//   a percentage `fo:font-size` scales the inherited size.
// - The cell style table. `styles[0]` is the flattened "Default" style (the
//   table-cell default-style + the common "Default" style); every other
//   table-cell style name referenced — by a cell, a row or column default
//   (`table:default-cell-style-name`) or a covered cell — gets ONE entry, in
//   first-reference order, with `sourceStyleName` set. Names are never
//   deduplicated (two names are two entries, however alike); "Default" is
//   always 0.
// - SPEC-123 named styles. Every common table-cell style of `office:styles`
//   is one (`namedStyles`, document order): named by its
//   `style:display-name`, else its `style:name` with LibreOffice's `_xx_`
//   hex escapes decoded ("Heading_20_1" → "Heading 1"); its definition what
//   it and its ancestors set short of "Default" and the family
//   default-style (the decoder below, its data style included), `includes`
//   the aspects that sets. "Default" is Normal: styles[0]'s definition,
//   every aspect. A cell style's `namedStyle` is the display name of the
//   common style it is — a cell, row or column naming one directly — or of
//   an automatic style's `style:parent-style-name`; "Default" names none.
// - SPEC-118: `style:data-style-name` (inherited down the same chain, the
//   nearest definition winning) names a `number:*-style` in content.xml's
//   automatic styles or styles.xml's common styles; odsNumberStyleToCode
//   turns it into `numberFormat` (General → absent). A date, time or
//   percentage value in a style with no data style at all gets a derived
//   style with its implied format (impliedStyleId, impliedFormats.ts).
// - Mapping (the spec's table): `table-cell-properties` fo:background-color →
//   a solid fill (`transparent` → none); fo:border-* → borders by the pinned
//   "ODS borders" table; style:diagonal-bl-tr / -tl-br → the diagonal with
//   up / down; fo:wrap-option="wrap"; style:vertical-align top / middle /
//   bottom (automatic → absent); style:rotation-angle a (style:direction
//   "ttb" → 255): a ≤ 90 → a, a ≥ 270 → 90 + (360 − a), 90 < a ≤ 180 → 90,
//   else 180; style:shrink-to-fit; style:repeat-content="true" → horizontal
//   "fill"; style:cell-protect. `paragraph-properties` fo:text-align
//   start / end / center / justify → left / right / center / justify
//   (style:text-align-source="value-type" or no fo:text-align → general,
//   i.e. absent); fo:margin-left → indent = round(px / INDENT_PX_PER_LEVEL).
//   `text-properties` → the font (style:font-name through the font-face
//   declarations, fo:font-size, bold, italic, underline, line-through,
//   fo:color, style:text-position).
// - Colours are literal `#rrggbb` (lowercased); `transparent` → absent. ODS
//   has no theme: nothing here produces a theme reference.
// - Lengths are any ODF unit (in, cm, mm, pt, pc, px) at 96 dpi, converted
//   through exact rationals so "0.178in" is 12.816 pt, not 12.8159999….
// - Absent properties are omitted — a CellStyle never carries an
//   `undefined`-valued key (the model is plain JSON).
import { FO_NAMESPACE, NUMBER_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, SVG_NAMESPACE, TABLE_NAMESPACE } from "../../core/odfXml";
import { DEFAULT_THEME, resolveColor, resolveFont, type DocumentTheme, type ModelColor } from "../../core/theme";
import { escapeXmlAttribute, parseXml, XmlParseError } from "../../core/xml";
import { INDENT_PX_PER_LEVEL } from "../geometry";
import { ImpliedFormatStyles } from "../impliedFormats";
import { formatCodeOf } from "../numberFormat/builtin";
import type { SpreadsheetLocale } from "../numberFormat/locale";
import { odsNumberStyleToCode, odsNumberStyleXml } from "../numberFormat/odsNumberStyles";
import type {
  CellAlignment,
  CellBorderEdge,
  CellBorders,
  CellFill,
  CellFont,
  CellStyle,
  SheetBorderStyle,
  SpreadsheetModel,
} from "../model";
import { STYLE_ASPECTS, workbookNamedStyle, type NamedCellStyle } from "../styles";

/** OpenOffice's table extension namespace (`tableooo:tab-color`). */
export const TABLEOOO_NAMESPACE = "http://openoffice.org/2009/table";

/** The canonical prefix each namespace's attributes are keyed by, whatever
 * prefix the producer bound. Attributes in other namespaces are not read. */
const PREFIX_BY_NAMESPACE: ReadonlyMap<string, string> = new Map([
  [FO_NAMESPACE, "fo"],
  [STYLE_NAMESPACE, "style"],
  [TABLE_NAMESPACE, "table"],
  [SVG_NAMESPACE, "svg"],
  [TABLEOOO_NAMESPACE, "tableooo"],
]);

type PropertyKind = "table-cell" | "paragraph" | "text" | "table-column" | "table-row" | "table";
const PROPERTY_KINDS: ReadonlySet<string> = new Set<PropertyKind>([
  "table-cell",
  "paragraph",
  "text",
  "table-column",
  "table-row",
  "table",
]);

/** One `style:*-properties` element's attributes, keyed "prefix:local". */
type Props = Map<string, string>;
type PropertySets = Partial<Record<PropertyKind, Props>>;

/** A `style:style` or `style:default-style` as read. */
interface StyleDefinition {
  family: string;
  parent?: string;
  props: PropertySets;
  /** `style:data-style-name` (SPEC-118). */
  dataStyle?: string;
  /** `style:display-name` (SPEC-123: a common style's named-style name). */
  displayName?: string;
}

const EMPTY_PROPS: Props = new Map();

/** styles.xml as the reader needs it. */
export interface OdsStyleSheet {
  /** `style:default-style` by family. */
  defaults: Map<string, StyleDefinition>;
  /** The common styles (`office:styles`), keyed by styleKey(family, name). */
  common: Map<string, StyleDefinition>;
  /** `style:font-face@style:name` → its family (`svg:font-family`, unquoted). */
  fontFaces: Map<string, string>;
  /** The common data styles (`number:*-style`) by `style:name` (SPEC-118). */
  dataStyles: Map<string, Element>;
}

/** content.xml's style sections as the reader needs them. */
export interface OdsContentStyles {
  /** The automatic styles, keyed by styleKey(family, name). */
  automatic: Map<string, StyleDefinition>;
  fontFaces: Map<string, string>;
  /** The automatic data styles by `style:name` (SPEC-118). */
  dataStyles: Map<string, Element>;
}

/** A row style's layout properties, as tableXml.ts applies them. */
export interface OdsRowHeight {
  heightPt?: number;
  customHeight?: true;
}

/** A table style's properties, as the reader applies them. */
export interface OdsTableStyle {
  hidden: boolean;
  tabColor?: string;
}

function styleKey(family: string, name: string): string {
  return `${family}\u0000${name}`;
}

// --- parsing ------------------------------------------------------------------

/** Parses a small ODF part with the platform parser, refusing a document
 * type declaration the way core/xmlScanner.ts does for the parts it reads
 * (no producer writes one; refusing it closes the entity-expansion door).
 * Throws XmlParseError. */
export function parseOdsPart(xml: string): Document {
  if (/<!DOCTYPE/i.test(xml)) {
    throw new XmlParseError("A document type declaration (DTD) is not accepted");
  }
  return parseXml(xml);
}

function readProps(el: Element): Props {
  const props: Props = new Map();
  for (const attr of Array.from(el.attributes)) {
    const prefix = attr.namespaceURI === null ? undefined : PREFIX_BY_NAMESPACE.get(attr.namespaceURI);
    if (prefix !== undefined) {
      props.set(`${prefix}:${attr.localName}`, attr.value.trim());
    }
  }
  return props;
}

const BORDER_SIDES = ["top", "right", "bottom", "left"] as const;

/** Expands the shorthands a definition writes into the properties they set,
 * so flattening can merge per side: `fo:border` → `fo:border-{side}` (a side
 * the same element names wins), `fo:margin` → `fo:margin-left`. */
function expandShorthands(kind: PropertyKind, props: Props): Props {
  if (kind === "table-cell") {
    const all = props.get("fo:border");
    if (all !== undefined) {
      props.delete("fo:border");
      for (const side of BORDER_SIDES) {
        if (!props.has(`fo:border-${side}`)) {
          props.set(`fo:border-${side}`, all);
        }
      }
    }
  } else if (kind === "paragraph") {
    const margin = props.get("fo:margin");
    if (margin !== undefined) {
      props.delete("fo:margin");
      if (!props.has("fo:margin-left")) {
        props.set("fo:margin-left", margin);
      }
    }
  }
  return props;
}

function readDefinition(el: Element, family: string): StyleDefinition {
  const props: PropertySets = {};
  for (const child of Array.from(el.children)) {
    if (child.namespaceURI !== STYLE_NAMESPACE || !child.localName.endsWith("-properties")) {
      continue;
    }
    const kind = child.localName.slice(0, -"-properties".length);
    if (PROPERTY_KINDS.has(kind)) {
      props[kind as PropertyKind] = expandShorthands(kind as PropertyKind, readProps(child));
    }
  }
  const definition: StyleDefinition = { family, props };
  const parent = el.getAttributeNS(STYLE_NAMESPACE, "parent-style-name");
  if (parent) {
    definition.parent = parent;
  }
  const dataStyle = el.getAttributeNS(STYLE_NAMESPACE, "data-style-name");
  if (dataStyle) {
    definition.dataStyle = dataStyle;
  }
  const displayName = el.getAttributeNS(STYLE_NAMESPACE, "display-name");
  if (displayName) {
    definition.displayName = displayName;
  }
  return definition;
}

/** Reads every data style (`number:*-style`) child of `container` by name. */
function readDataStyles(container: Element, into: Map<string, Element>): void {
  for (const el of Array.from(container.children)) {
    const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
    if (el.namespaceURI === NUMBER_NAMESPACE && el.localName.endsWith("-style") && name) {
      into.set(name, el);
    }
  }
}

/** Reads every `style:style` child of `container` into `into`. */
function readStyles(container: Element, into: Map<string, StyleDefinition>): void {
  for (const el of Array.from(container.children)) {
    if (el.namespaceURI !== STYLE_NAMESPACE || el.localName !== "style") {
      continue;
    }
    const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
    const family = el.getAttributeNS(STYLE_NAMESPACE, "family");
    if (name && family) {
      into.set(styleKey(family, name), readDefinition(el, family));
    }
  }
}

/** The first family of a CSS-style font-family list, quotes stripped. */
function firstFamily(value: string | null | undefined): string | undefined {
  const first = value?.split(",")[0]?.trim().replace(/^(['"])(.*)\1$/, "$2").trim();
  return first ? first : undefined;
}

function readFontFaces(container: Element, into: Map<string, string>): void {
  for (const el of Array.from(container.children)) {
    if (el.namespaceURI !== STYLE_NAMESPACE || el.localName !== "font-face") {
      continue;
    }
    const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
    const family = firstFamily(el.getAttributeNS(SVG_NAMESPACE, "font-family"));
    if (name) {
      into.set(name, family ?? name);
    }
  }
}

/** A styles.xml with nothing in it — the part is optional. */
export function emptyOdsStyleSheet(): OdsStyleSheet {
  return { defaults: new Map(), common: new Map(), fontFaces: new Map(), dataStyles: new Map() };
}

/** Reads styles.xml: the default styles, the common styles and the font
 * faces. Throws XmlParseError for a malformed part. */
export function readOdsStyleSheet(xml: string): OdsStyleSheet {
  const sheet = emptyOdsStyleSheet();
  const root = parseOdsPart(xml).documentElement;
  for (const section of Array.from(root.children)) {
    if (section.namespaceURI !== OFFICE_NAMESPACE) {
      continue;
    }
    if (section.localName === "font-face-decls") {
      readFontFaces(section, sheet.fontFaces);
    } else if (section.localName === "styles") {
      readStyles(section, sheet.common);
      readDataStyles(section, sheet.dataStyles);
      for (const el of Array.from(section.children)) {
        const family = el.getAttributeNS(STYLE_NAMESPACE, "family");
        if (el.namespaceURI === STYLE_NAMESPACE && el.localName === "default-style" && family) {
          sheet.defaults.set(family, readDefinition(el, family));
        }
      }
    }
  }
  return sheet;
}

/** The namespace declarations of an element's scanned attributes, as source
 * text — what a slice of the document needs to parse on its own. */
export function namespaceDeclarations(attrs: readonly { name: string; value: string }[]): string {
  return attrs
    .filter((attr) => attr.name === "xmlns" || attr.name.startsWith("xmlns:"))
    .map((attr) => `${attr.name}="${escapeXmlAttribute(attr.value)}"`)
    .join(" ");
}

/** Reads content.xml's `office:font-face-decls` and
 * `office:automatic-styles`, given as the verbatim source slices of those
 * elements plus the document element's namespace declarations (they parse
 * wrapped in a root that declares them). Throws XmlParseError. */
export function readOdsContentStyles(slices: readonly string[], declarations: string): OdsContentStyles {
  const styles: OdsContentStyles = { automatic: new Map(), fontFaces: new Map(), dataStyles: new Map() };
  if (slices.length === 0) {
    return styles;
  }
  const root = parseOdsPart(`<ods-slice ${declarations}>${slices.join("")}</ods-slice>`).documentElement;
  for (const section of Array.from(root.children)) {
    if (section.namespaceURI !== OFFICE_NAMESPACE) {
      continue;
    }
    if (section.localName === "font-face-decls") {
      readFontFaces(section, styles.fontFaces);
    } else if (section.localName === "automatic-styles") {
      readStyles(section, styles.automatic);
      readDataStyles(section, styles.dataStyles);
    }
  }
  return styles;
}

// --- style names (SPEC-123) -----------------------------------------------------

/** Whether LibreOffice's style-name encoder (xmloff's
 * SvXMLUnitConverter::encodeStyleName) keeps a UTF-16 unit as it is: an
 * NCName character at its position, `_` (the escape) excepted. */
function literalNameUnit(unit: string, first: boolean): boolean {
  const c = unit.charCodeAt(0);
  if (c < 0xff) {
    return (
      (c >= 0x41 && c <= 0x5a) ||
      (c >= 0x61 && c <= 0x7a) ||
      (c >= 0xc0 && c <= 0xd6) ||
      (c >= 0xd8 && c <= 0xf6) ||
      c >= 0xf8 ||
      (!first && ((c >= 0x30 && c <= 0x39) || c === 0xb7 || c === 0x2d || c === 0x2e))
    );
  }
  if ((c >= 0xf900 && c <= 0xfffe) || (c >= 0x20dd && c <= 0x20e0)) {
    return false;
  }
  if ((c >= 0x2bb && c <= 0x2c1) || c === 0x559 || c === 0x6e5 || c === 0x6e6) {
    return true;
  }
  if (c === 0x387) {
    return !first;
  }
  if (/[\p{Lu}\p{Ll}\p{Lt}\p{Lo}\p{Nl}]/u.test(unit)) {
    return true;
  }
  return !first && /[\p{Mn}\p{Me}\p{Mc}\p{Lm}\p{Nd}]/u.test(unit);
}

/** A display name as the common style name LibreOffice writes for it: each
 * UTF-16 unit that is not an NCName character at its position (`_`
 * included) as `_<hex>_`, lowercase and without leading zeros —
 * "Heading 1" → "Heading_20_1", "A_B" → "A_5f_B". */
export function encodeOdsStyleName(name: string): string {
  let out = "";
  for (let i = 0; i < name.length; i += 1) {
    const unit = name[i]!;
    out += literalNameUnit(unit, i === 0) ? unit : `_${unit.charCodeAt(0).toString(16)}_`;
  }
  return out;
}

/** A common style name with LibreOffice's `_<hex>_` escapes decoded — the
 * display name of a style that writes none ("Heading_20_1" → "Heading 1").
 * An escape naming a control character stays as written. */
export function decodeOdsStyleName(name: string): string {
  return name.replace(/_([0-9a-fA-F]{1,4})_/g, (escape, hex: string) => {
    const code = Number.parseInt(hex, 16);
    return code >= 0x20 && code < 0xfffe ? String.fromCharCode(code) : escape;
  });
}

// --- lengths, angles and colours ---------------------------------------------

/** Each ODF length unit as an exact rational [numerator, denominator] of the
 * target unit, so a decimal source converts with one correctly rounded
 * division (2.54 cm and 25.4 mm to the inch: 72/2.54 = 3600/127). */
const PT_PER_UNIT: ReadonlyMap<string, readonly [number, number]> = new Map([
  ["pt", [1, 1]],
  ["in", [72, 1]],
  ["pc", [12, 1]],
  ["px", [3, 4]],
  ["cm", [3600, 127]],
  ["mm", [360, 127]],
]);
const PX_PER_UNIT: ReadonlyMap<string, readonly [number, number]> = new Map([
  ["px", [1, 1]],
  ["pt", [4, 3]],
  ["in", [96, 1]],
  ["pc", [16, 1]],
  ["cm", [4800, 127]],
  ["mm", [480, 127]],
]);

const LENGTH_PATTERN = /^([-+]?)(\d*)(?:\.(\d*))?\s*([a-z]*)$/i;

function convertLength(value: string | undefined, units: ReadonlyMap<string, readonly [number, number]>): number | undefined {
  const match = value === undefined ? null : LENGTH_PATTERN.exec(value.trim());
  if (!match) {
    return undefined;
  }
  const [, sign, whole = "", fraction = "", unit = ""] = match;
  if (whole.length + fraction.length === 0) {
    return undefined;
  }
  const digits = Number(`${whole}${fraction}` || "0");
  const negative = sign === "-" ? -1 : 1;
  if (unit === "") {
    return digits === 0 ? 0 : undefined; // a bare number is only a length when it is zero
  }
  const ratio = units.get(unit.toLowerCase());
  if (!ratio) {
    return undefined;
  }
  const scale = 10 ** fraction.length;
  const exact = Number.isSafeInteger(digits * ratio[0]) && Number.isSafeInteger(scale * ratio[1]);
  const result = exact
    ? (digits * ratio[0]) / (scale * ratio[1])
    : (Number(`${whole || "0"}.${fraction || "0"}`) * ratio[0]) / ratio[1];
  return Number.isFinite(result) ? negative * result : undefined;
}

/** An ODF length (in, cm, mm, pt, pc, px) in points, or undefined. */
export function odsLengthToPt(value: string | undefined): number | undefined {
  return convertLength(value, PT_PER_UNIT);
}

/** An ODF length in CSS px at 96 dpi, unrounded, or undefined. */
export function odsLengthToPx(value: string | undefined): number | undefined {
  return convertLength(value, PX_PER_UNIT);
}

/** An ODF colour as the model's literal: `#rrggbb` lowercased (`#rgb`
 * expanded); `transparent`, a keyword or garbage → undefined. */
export function odsColor(value: string | undefined): string | undefined {
  const text = value?.trim().toLowerCase();
  if (text === undefined) {
    return undefined;
  }
  if (/^#[0-9a-f]{6}$/.test(text)) {
    return text;
  }
  if (/^#[0-9a-f]{3}$/.test(text)) {
    return `#${text[1]}${text[1]}${text[2]}${text[2]}${text[3]}${text[3]}`;
  }
  return undefined;
}

const DEGREES_PER_UNIT: ReadonlyMap<string, number> = new Map([
  ["", 1],
  ["deg", 1],
  ["grad", 0.9],
  ["rad", 180 / Math.PI],
]);

/** `style:rotation-angle` (degrees, or with a deg/grad/rad unit) → the
 * model's 0..180 textRotation: a ≤ 90 → a; a ≥ 270 → 90 + (360 − a);
 * 90 < a ≤ 180 → 90; else 180. 0 → undefined (no rotation). */
export function odsTextRotation(value: string | undefined): number | undefined {
  const match = value === undefined ? null : /^([-+]?\d*\.?\d+)\s*(deg|grad|rad)?$/i.exec(value.trim());
  if (!match) {
    return undefined;
  }
  const degrees = Number(match[1]) * DEGREES_PER_UNIT.get((match[2] ?? "").toLowerCase())!;
  if (!Number.isFinite(degrees)) {
    return undefined;
  }
  const a = ((Math.round(degrees) % 360) + 360) % 360;
  if (a === 0) {
    return undefined;
  }
  if (a <= 90) {
    return a;
  }
  if (a >= 270) {
    return 90 + (360 - a);
  }
  return a <= 180 ? 90 : 180;
}

// --- borders (SPEC-117 "ODS borders") ----------------------------------------

/** The read table: line style → [w ≤ 1.0 pt, 1.0 < w ≤ 2.0 pt, w > 2.0 pt]. */
const BORDER_READ: ReadonlyMap<string, readonly [SheetBorderStyle, SheetBorderStyle, SheetBorderStyle]> = new Map([
  ["solid", ["thin", "medium", "thick"]],
  ["dashed", ["dashed", "mediumDashed", "mediumDashed"]],
  ["dotted", ["dotted", "dotted", "dotted"]],
  ["fine-dashed", ["hair", "hair", "hair"]],
  ["dash-dot", ["dashDot", "mediumDashDot", "mediumDashDot"]],
  ["dash-dot-dot", ["dashDotDot", "mediumDashDotDot", "mediumDashDotDot"]],
  ["double", ["double", "double", "double"]],
  ["double-thin", ["double", "double", "double"]],
  // CSS's 3D styles have no spreadsheet counterpart: drawn as solid lines.
  ["groove", ["thin", "medium", "thick"]],
  ["ridge", ["thin", "medium", "thick"]],
  ["inset", ["thin", "medium", "thick"]],
  ["outset", ["thin", "medium", "thick"]],
]);
const NO_LINE: ReadonlySet<string> = new Set(["none", "hidden"]);
/** CSS's width keywords, in pt. */
const WIDTH_KEYWORDS: ReadonlyMap<string, number> = new Map([
  ["thin", 0.75],
  ["medium", 1.5],
  ["thick", 2.25],
]);

/** One `fo:border*` / `style:diagonal-*` value — "<width> <style> <#color>",
 * in any order — as a border edge by the pinned read table; undefined for
 * `none`/`hidden`, a zero width or no line style. A literal colour is kept
 * as written (`#000000` included); none written → absent (automatic). */
export function odsBorderEdge(value: string | undefined): CellBorderEdge | undefined {
  if (value === undefined) {
    return undefined;
  }
  let width: number | undefined;
  let line: string | undefined;
  let color: string | undefined;
  for (const token of value.trim().toLowerCase().split(/\s+/)) {
    if (token.startsWith("#")) {
      color = odsColor(token);
    } else if (NO_LINE.has(token) || BORDER_READ.has(token)) {
      line = token;
    } else if (WIDTH_KEYWORDS.has(token)) {
      width = WIDTH_KEYWORDS.get(token);
    } else {
      width = odsLengthToPt(token) ?? width;
    }
  }
  const tiers = line === undefined ? undefined : BORDER_READ.get(line);
  if (!tiers || (width !== undefined && width <= 0)) {
    return undefined;
  }
  const w = width ?? 0.75;
  const edge: CellBorderEdge = { style: tiers[w <= 1 ? 0 : w <= 2 ? 1 : 2] };
  if (color !== undefined) {
    edge.color = color;
  }
  return edge;
}

/** The write table (pinned for SPEC-122): each style's width and ODF line
 * style. slantDashDot writes as mediumDashDot's line and so reads back as
 * mediumDashDot — the one stated loss. */
export const ODS_BORDER_WRITE: Readonly<Record<SheetBorderStyle, string>> = {
  thin: "0.74pt solid",
  medium: "1.75pt solid",
  thick: "2.5pt solid",
  dashed: "0.74pt dashed",
  mediumDashed: "1.75pt dashed",
  dotted: "0.74pt dotted",
  hair: "0.74pt fine-dashed",
  dashDot: "0.74pt dash-dot",
  mediumDashDot: "1.75pt dash-dot",
  slantDashDot: "1.75pt dash-dot",
  dashDotDot: "0.74pt dash-dot-dot",
  mediumDashDotDot: "1.75pt dash-dot-dot",
  double: "2.6pt double",
};

/** A border edge as an `fo:border*` value: the write table's width and line
 * style followed by the colour — `#000000` when automatic; a theme reference
 * (an edge that came from an xlsx) resolved against `theme`. */
export function odsBorderString(edge: CellBorderEdge, theme: DocumentTheme = DEFAULT_THEME): string {
  const color = edge.color === undefined ? "#000000" : resolveColor(edge.color, theme);
  return `${ODS_BORDER_WRITE[edge.style]} ${color}`;
}

// --- the property mapping -----------------------------------------------------

const HORIZONTAL: ReadonlyMap<string, NonNullable<CellAlignment["horizontal"]>> = new Map([
  ["start", "left"],
  ["left", "left"],
  ["end", "right"],
  ["right", "right"],
  ["center", "center"],
  ["justify", "justify"],
]);
const VERTICAL: ReadonlyMap<string, NonNullable<CellAlignment["vertical"]>> = new Map([
  ["top", "top"],
  ["middle", "center"],
  ["bottom", "bottom"],
]);

function bordersFrom(cell: Props): CellBorders | undefined {
  const borders: CellBorders = {};
  for (const side of BORDER_SIDES) {
    const edge = odsBorderEdge(cell.get(`fo:border-${side}`));
    if (edge) {
      borders[side] = edge;
    }
  }
  const up = odsBorderEdge(cell.get("style:diagonal-bl-tr"));
  const down = odsBorderEdge(cell.get("style:diagonal-tl-br"));
  const diagonal = up ?? down;
  if (diagonal) {
    borders.diagonal = { ...diagonal, ...(up ? { up: true as const } : {}), ...(down ? { down: true as const } : {}) };
  }
  return Object.keys(borders).length > 0 ? borders : undefined;
}

function alignmentFrom(cell: Props, paragraph: Props): CellAlignment | undefined {
  const alignment: CellAlignment = {};
  const horizontal =
    cell.get("style:repeat-content") === "true"
      ? "fill"
      : cell.get("style:text-align-source") === "value-type"
        ? undefined
        : HORIZONTAL.get(paragraph.get("fo:text-align") ?? "");
  if (horizontal !== undefined) {
    alignment.horizontal = horizontal;
  }
  const vertical = VERTICAL.get(cell.get("style:vertical-align") ?? "");
  if (vertical !== undefined) {
    alignment.vertical = vertical;
  }
  if (cell.get("fo:wrap-option") === "wrap") {
    alignment.wrap = true;
  }
  const marginPx = odsLengthToPx(paragraph.get("fo:margin-left"));
  const indent = marginPx === undefined ? 0 : Math.round(marginPx / INDENT_PX_PER_LEVEL);
  if (indent > 0) {
    alignment.indent = indent;
  }
  const rotation = cell.get("style:direction") === "ttb" ? 255 : odsTextRotation(cell.get("style:rotation-angle"));
  if (rotation !== undefined) {
    alignment.textRotation = rotation;
  }
  if (cell.get("style:shrink-to-fit") === "true") {
    alignment.shrinkToFit = true;
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined;
}

function protectionFrom(value: string | undefined): CellStyle["protection"] {
  const tokens = value === undefined ? [] : value.trim().split(/\s+/);
  if (tokens.includes("none")) {
    return { locked: false };
  }
  if (tokens.includes("hidden-and-protected") || tokens.includes("formula-hidden")) {
    return { hidden: true };
  }
  return undefined;
}

function vertAlignFrom(value: string | undefined): CellFont["vertAlign"] {
  const first = value?.trim().split(/\s+/)[0]?.toLowerCase() ?? "";
  if (first.startsWith("super")) {
    return "superscript";
  }
  if (first.startsWith("sub")) {
    return "subscript";
  }
  const percent = /^([-+]?\d*\.?\d+)%$/.exec(first);
  const shift = percent ? Number(percent[1]) : 0;
  return shift > 0 ? "superscript" : shift < 0 ? "subscript" : undefined;
}

/** `style:text-properties` → a CellFont (undefined when it sets nothing the
 * model holds). Font names resolve through the font-face declarations. */
function fontFrom(text: Props, fontFaces: ReadonlyMap<string, string>): CellFont | undefined {
  const font: CellFont = {};
  const name = text.get("style:font-name");
  const family = name !== undefined ? (fontFaces.get(name) ?? name) : firstFamily(text.get("fo:font-family"));
  if (family) {
    font.family = family;
  }
  const size = odsLengthToPt(text.get("fo:font-size"));
  if (size !== undefined && size > 0) {
    font.sizePt = size;
  }
  const weight = text.get("fo:font-weight");
  if (weight === "bold" || Number(weight) >= 600) {
    font.bold = true;
  }
  const fontStyle = text.get("fo:font-style");
  if (fontStyle === "italic" || fontStyle === "oblique") {
    font.italic = true;
  }
  const underline = text.get("style:text-underline-style");
  const underlineType = text.get("style:text-underline-type");
  if (underline !== undefined && underline !== "none" && underlineType !== "none") {
    font.underline = underlineType === "double" ? "double" : "single";
  }
  const strike = text.get("style:text-line-through-style");
  if (strike !== undefined && strike !== "none" && text.get("style:text-line-through-type") !== "none") {
    font.strike = true;
  }
  const color = odsColor(text.get("fo:color"));
  if (color !== undefined) {
    font.color = color;
  }
  const vertAlign = vertAlignFrom(text.get("style:text-position"));
  if (vertAlign !== undefined) {
    font.vertAlign = vertAlign;
  }
  return Object.keys(font).length > 0 ? font : undefined;
}

/** Flattened properties → a CellStyle (the spec's mapping table). */
function cellStyleFrom(sets: PropertySets, fontFaces: ReadonlyMap<string, string>): CellStyle {
  const style: CellStyle = {};
  const cell = sets["table-cell"] ?? EMPTY_PROPS;
  const font = fontFrom(sets.text ?? EMPTY_PROPS, fontFaces);
  if (font) {
    style.font = font;
  }
  const fill = odsColor(cell.get("fo:background-color"));
  if (fill !== undefined) {
    style.fill = { kind: "solid", color: fill };
  }
  const borders = bordersFrom(cell);
  if (borders) {
    style.borders = borders;
  }
  const alignment = alignmentFrom(cell, sets.paragraph ?? EMPTY_PROPS);
  if (alignment) {
    style.alignment = alignment;
  }
  const protection = protectionFrom(cell.get("style:cell-protect"));
  if (protection) {
    style.protection = protection;
  }
  return style;
}

/** Merges `source` over `target`, per property. A percentage font size
 * scales the size inherited so far (and is dropped when there is none). */
function mergeInto(target: PropertySets, source: PropertySets): void {
  for (const [kind, props] of Object.entries(source) as [PropertyKind, Props][]) {
    const into = (target[kind] ??= new Map());
    for (const [key, value] of props) {
      if (kind === "text" && key === "fo:font-size" && value.endsWith("%")) {
        const inherited = odsLengthToPt(into.get(key));
        const percent = Number(value.slice(0, -1));
        if (inherited !== undefined && Number.isFinite(percent)) {
          into.set(key, `${(inherited * percent) / 100}pt`);
        }
        continue;
      }
      into.set(key, value);
    }
  }
}

/** The text-property keys a run's font models (see textStyleModelled). */
const MODELLED_TEXT_PROPERTY =
  /^(?:style:font-|fo:font-|fo:color$|style:use-window-font-color$|style:text-position$|style:text-underline-|style:text-line-through-|fo:language$|fo:country$|fo:script$|style:language-|style:country-|style:script-|style:rfc-language-tag)/;

/** A parent chain longer than this is malformed (or a cycle). */
const MAX_STYLE_DEPTH = 64;

/** Every style an .ods workbook's cells, rows, columns and tables name:
 * styles.xml's default and common styles plus content.xml's automatic ones,
 * flattened on first use and cached. `styles` is the model's cell style
 * table — it grows as cellStyleId meets new names (first-reference order). */
export class OdsStyleBook {
  /** The cell styles, index 0 the flattened "Default" style. */
  readonly styles: CellStyle[];
  private readonly sheet: OdsStyleSheet;
  private readonly automatic: Map<string, StyleDefinition>;
  private readonly fontFaces: Map<string, string>;
  private readonly cellIds = new Map<string, number>();
  private readonly dataStyles: Map<string, Element>;
  /** The cell styles that name a data style (SPEC-118) — General or not. */
  private readonly formatted = new Set<number>();
  private readonly implied: ImpliedFormatStyles;
  private readonly textFonts = new Map<string, CellFont | undefined>();
  private readonly textModelled = new Map<string, boolean>();
  private readonly columnWidths = new Map<string | undefined, number | undefined>();
  private readonly rowHeights = new Map<string | undefined, OdsRowHeight>();

  constructor(sheet: OdsStyleSheet, content: OdsContentStyles = { automatic: new Map(), fontFaces: new Map(), dataStyles: new Map() }) {
    this.sheet = sheet;
    this.automatic = content.automatic;
    this.fontFaces = new Map([...sheet.fontFaces, ...content.fontFaces]);
    this.dataStyles = new Map([...sheet.dataStyles, ...content.dataStyles]);
    const base = cellStyleFrom(this.flatten("table-cell", "Default", { commonOnly: true }), this.fontFaces);
    this.applyDataStyle(base, 0, this.dataStyleOf("Default", true));
    if (sheet.common.has(styleKey("table-cell", "Default"))) {
      base.sourceStyleName = "Default";
    }
    this.styles = [base];
    this.cellIds.set("Default", 0);
    this.implied = new ImpliedFormatStyles(this.styles);
  }

  /** The index in `styles` of the table-cell style `name`, adding its
   * flattened entry on first reference. A name no part defines reads as the
   * default style's properties under its own name. */
  cellStyleId(name: string): number {
    let id = this.cellIds.get(name);
    if (id === undefined) {
      id = this.styles.length;
      const style = cellStyleFrom(this.flatten("table-cell", name), this.fontFaces);
      this.applyDataStyle(style, id, this.dataStyleOf(name, false));
      const named = this.namedStyleOf(name);
      if (named !== undefined) {
        style.namedStyle = named;
      }
      style.sourceStyleName = name;
      this.styles.push(style);
      this.cellIds.set(name, id);
    }
    return id;
  }

  /** SPEC-123: the workbook's named styles — every common table-cell style,
   * in document order (see the module header). */
  namedStyles(): NamedCellStyle[] {
    const named: NamedCellStyle[] = [];
    for (const [key, definition] of this.sheet.common) {
      if (definition.family !== "table-cell") {
        continue;
      }
      const name = key.slice(key.indexOf("\u0000") + 1);
      if (name === "Default") {
        const { sourceStyleName: _name, ...normal } = this.styles[0]!;
        named.push({ name: "Normal", style: structuredClone(normal), includes: [...STYLE_ASPECTS], sourceStyleName: name });
        continue;
      }
      const style = cellStyleFrom(this.flatten("table-cell", name, { withDefault: false, commonOnly: true, stopAt: "Default" }), this.fontFaces);
      const code = this.dataStyleCode(this.dataStyleOf(name, true, "Default"));
      if (code !== undefined) {
        style.numberFormat = code;
      }
      const includes = STYLE_ASPECTS.filter((aspect) => style[aspect] !== undefined);
      named.push({ name: this.displayName(name), style, includes, sourceStyleName: name });
    }
    return named;
  }

  /** A common style's named-style name: its display name, else its decoded
   * `style:name`. */
  private displayName(name: string): string {
    return this.sheet.common.get(styleKey("table-cell", name))?.displayName ?? decodeOdsStyleName(name);
  }

  /** The named style a table-cell style name shows: the first common style
   * met walking up from it — itself when a cell names a common style
   * directly, else its (automatic) style's parent — as its display name;
   * undefined at "Default", or with no common style on the way. */
  private namedStyleOf(name: string): string | undefined {
    const seen = new Set<string>();
    let current: string | undefined = name;
    // The name as a cell references it: an automatic style first.
    let automatic = this.automatic.get(styleKey("table-cell", name));
    while (current !== undefined && current !== "Default" && !seen.has(current) && seen.size < MAX_STYLE_DEPTH) {
      seen.add(current);
      if (automatic === undefined) {
        return this.sheet.common.has(styleKey("table-cell", current)) ? this.displayName(current) : undefined;
      }
      current = automatic.parent;
      automatic = current === undefined || this.sheet.common.has(styleKey("table-cell", current)) ? undefined : this.automatic.get(styleKey("table-cell", current));
    }
    return undefined;
  }

  /** The style a date, time or percentage value with an implied `code`
   * shows (SPEC-118 "Implied formats"): `baseId` when that style names a
   * data style, else a derived style carrying `code`. */
  impliedStyleId(baseId: number, code: string): number {
    return this.formatted.has(baseId) ? baseId : this.implied.styleFor(baseId, code);
  }

  /** The data style a table-cell style names, inherited down its parent
   * chain (the nearest definition wins), else the family default-style's;
   * undefined when none does. With `stopAt` (SPEC-123: a named style's own
   * definition) the walk ends below that ancestor and the default-style is
   * not consulted. */
  private dataStyleOf(name: string, commonOnly: boolean, stopAt?: string): string | undefined {
    const seen = new Set<string>();
    let current = commonOnly ? this.sheet.common.get(styleKey("table-cell", name)) : this.lookup("table-cell", name, false);
    seen.add(name);
    for (let depth = 0; current && depth < MAX_STYLE_DEPTH; depth += 1) {
      if (current.dataStyle !== undefined) {
        return current.dataStyle;
      }
      const parent = current.parent;
      if (parent === undefined || seen.has(parent) || parent === stopAt) {
        break;
      }
      seen.add(parent);
      current = this.lookup("table-cell", parent, true);
    }
    return stopAt === undefined ? this.sheet.defaults.get("table-cell")?.dataStyle : undefined;
  }

  /** Gives a cell style the number format its data style maps to. */
  private applyDataStyle(style: CellStyle, id: number, dataStyleName: string | undefined): void {
    if (dataStyleName === undefined || !this.dataStyles.has(dataStyleName)) {
      return;
    }
    this.formatted.add(id);
    const code = this.dataStyleCode(dataStyleName);
    if (code !== undefined) {
      style.numberFormat = code;
    }
  }

  /** The format code a data style maps to; undefined for General or a name
   * no part defines. */
  private dataStyleCode(dataStyleName: string | undefined): string | undefined {
    const element = dataStyleName === undefined ? undefined : this.dataStyles.get(dataStyleName);
    const code = element ? odsNumberStyleToCode(element, (name) => this.dataStyles.get(name)) : undefined;
    return code !== undefined && !/^general$/i.test(code.trim()) ? code : undefined;
  }

  /** A `text:span`'s own font: its text style's `style:text-properties`
   * (with the style's parents, never the cell's), or undefined. The object
   * is shared — copy it into the model. */
  textFont(name: string): CellFont | undefined {
    if (!this.textFonts.has(name)) {
      const text = this.flatten("text", name, { withDefault: false }).text;
      this.textFonts.set(name, text ? fontFrom(text, this.fontFaces) : undefined);
    }
    return this.textFonts.get(name);
  }

  /** SPEC-122: whether a `text:span`'s style holds only what a run's font
   * models — text properties naming the font, its size, weight, posture,
   * underline, line-through, colour and position (with their asian/complex
   * and language twins, which say nothing the model shows) — so the span is
   * re-created from `TextRun.font`. A name no part defines is modelled (it
   * shows nothing). Anything else (a highlight, letter spacing, a paragraph
   * property…) makes the span content the model does not hold. */
  textStyleModelled(name: string): boolean {
    let modelled = this.textModelled.get(name);
    if (modelled === undefined) {
      const sets = this.flatten("text", name, { withDefault: false });
      modelled =
        Object.keys(sets).every((kind) => kind === "text") &&
        [...(sets.text?.keys() ?? [])].every((key) => MODELLED_TEXT_PROPERTY.test(key));
      this.textModelled.set(name, modelled);
    }
    return modelled;
  }

  /** SPEC-122: whether styles.xml defines a common table-cell style `name`
   * (a generated style's `style:parent-style-name` must name one). */
  commonCellStyle(name: string): boolean {
    return this.sheet.common.has(styleKey("table-cell", name));
  }

  /** SPEC-122: the text styles a rich run can be written back in, so a run
   * read from the file keeps its span's style name — every text-family
   * style of content.xml's automatic styles, then of styles.xml's common
   * styles (a common one an automatic one shadows left out), in document
   * order, that holds only what a run's font models (textStyleModelled) and
   * gives a font, with that font. */
  runTextStyles(): { name: string; font: CellFont }[] {
    const found: { name: string; font: CellFont }[] = [];
    const seen = new Set<string>();
    for (const styles of [this.automatic, this.sheet.common]) {
      for (const [key, definition] of styles) {
        const name = key.slice(key.indexOf("\u0000") + 1);
        if (definition.family !== "text" || seen.has(name)) {
          continue;
        }
        seen.add(name);
        const font = this.textStyleModelled(name) ? this.textFont(name) : undefined;
        if (font) {
          found.push({ name, font });
        }
      }
    }
    return found;
  }

  /** A column style's `style:column-width` in px (unrounded) — the
   * table-column default-style's when the style sets none — or undefined. */
  columnWidthPx(name: string | undefined): number | undefined {
    if (!this.columnWidths.has(name)) {
      const props = this.flatten("table-column", name)["table-column"];
      this.columnWidths.set(name, odsLengthToPx(props?.get("style:column-width")));
    }
    return this.columnWidths.get(name);
  }

  /** A row style's height in pt and whether it is a custom one
   * (`style:use-optimal-row-height="false"`). */
  rowHeight(name: string | undefined): OdsRowHeight {
    let height = this.rowHeights.get(name);
    if (height === undefined) {
      const props = this.flatten("table-row", name)["table-row"] ?? EMPTY_PROPS;
      const heightPt = odsLengthToPt(props.get("style:row-height"));
      height = {};
      if (heightPt !== undefined && heightPt > 0) {
        height.heightPt = heightPt;
        if (props.get("style:use-optimal-row-height") === "false") {
          height.customHeight = true;
        }
      }
      this.rowHeights.set(name, height);
    }
    return height;
  }

  /** A table style's `table:display="false"` (hidden) and
   * `tableooo:tab-color`. */
  tableStyle(name: string | undefined): OdsTableStyle {
    const props = this.flatten("table", name).table ?? EMPTY_PROPS;
    const style: OdsTableStyle = { hidden: props.get("table:display") === "false" };
    const tabColor = odsColor(props.get("tableooo:tab-color"));
    if (tabColor !== undefined) {
      style.tabColor = tabColor;
    }
    return style;
  }

  private lookup(family: string, name: string, commonFirst: boolean): StyleDefinition | undefined {
    const key = styleKey(family, name);
    return commonFirst
      ? (this.sheet.common.get(key) ?? this.automatic.get(key))
      : (this.automatic.get(key) ?? this.sheet.common.get(key));
  }

  /** The family's default-style, then the named style's ancestors from the
   * root down, then the style itself — merged per property. `stopAt` ends
   * the chain below that ancestor (SPEC-123: a named style's definition
   * leaves out "Default"). */
  private flatten(
    family: string,
    name: string | undefined,
    options: { withDefault?: boolean; commonOnly?: boolean; stopAt?: string } = {},
  ): PropertySets {
    const chain: StyleDefinition[] = [];
    const seen = new Set<string>();
    let current =
      name === undefined
        ? undefined
        : options.commonOnly
          ? this.sheet.common.get(styleKey(family, name))
          : this.lookup(family, name, false);
    if (name !== undefined) {
      seen.add(name);
    }
    while (current && chain.length < MAX_STYLE_DEPTH) {
      chain.push(current);
      const parent = current.parent;
      if (parent === undefined || seen.has(parent) || parent === options.stopAt) {
        break;
      }
      seen.add(parent);
      current = this.lookup(family, parent, true);
    }
    const merged: PropertySets = {};
    const base = options.withDefault === false ? undefined : this.sheet.defaults.get(family);
    if (base) {
      mergeInto(merged, base.props);
    }
    for (let i = chain.length - 1; i >= 0; i -= 1) {
      mergeInto(merged, chain[i]!.props);
    }
    return merged;
  }
}

// --- generated styles (SPEC-121, extended by SPEC-122) ------------------------------------

/** A generated table style's properties: the master page (the source table
 * style's, else "Default"), whether the table shows, and its tab colour as
 * hex. */
export interface GeneratedTableStyle {
  masterPage: string;
  display: boolean;
  tabColor?: string;
}

/** What a generated cell or text style resolves against (SPEC-122): theme
 * colours and fonts, the user's locale (a `builtinFormatId` is written as
 * that built-in's code in it — ODS has no locale-dependent ids) and the
 * style's parent (see parentStyleName). */
export interface OdsStyleEnv {
  theme: DocumentTheme;
  locale: SpreadsheetLocale;
  /** SPEC-123: the common style a named style the workbook defines is
   * written under (OdsNamedStyles.parent); undefined for one it does not. */
  namedStyleParent?: (name: string) => string | undefined;
  /** Whether the target package defines a common table-cell style. */
  commonCellStyle?: (name: string) => boolean;
}

/** A generated cell style's `style:parent-style-name`: "Default" for no
 * `namedStyle` or Normal; the common style a named style the workbook
 * defines is written under (SPEC-123); for a name the workbook does not
 * define, the common style of that name when the package defines one,
 * else "Default". */
function parentStyleName(named: string | undefined, env: OdsStyleEnv): string {
  if (named === undefined || named === "Normal") {
    return "Default";
  }
  return env.namedStyleParent?.(named) ?? (env.commonCellStyle?.(named) ? named : "Default");
}

/** One generated style's property elements, as attribute lists. */
export interface OdsCellStyleProperties {
  /** `style:table-cell-properties` attributes, each with its leading space. */
  cell: string;
  /** `style:paragraph-properties` attributes. */
  paragraph: string;
  /** `style:text-properties` attributes. */
  text: string;
}

/** The ODF value of a font family: quoted when it holds a space or comma
 * (`'Liberation Sans'`), as LibreOffice writes `svg:font-family`. */
export function odsFontFamilyValue(family: string): string {
  return /[\s,]/.test(family) ? `'${family.replace(/'/g, "")}'` : family;
}

/** A number in a length or angle attribute: shortest round-trip text, never
 * an exponent (`-0` as `0`). */
function lengthNumber(n: number): string {
  if (Object.is(n, -0) || n === 0) {
    return "0";
  }
  const text = String(n);
  return /e/i.test(text) ? n.toFixed(6).replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "") : text;
}

/** A model colour as ODF's `#rrggbb` (theme references resolved). */
function hexColor(color: ModelColor, theme: DocumentTheme): string {
  return resolveColor(color, theme).toLowerCase();
}

/** A run or cell font as `style:text-properties` attributes — the inverse of
 * fontFrom: the family as `fo:font-family` (theme fonts resolved; `face`
 * writes `style:font-name` instead, for a style whose font-face the caller
 * declares), `fo:font-size` in pt, bold, italic, a solid single/double
 * underline (the accounting underlines as their plain twins), a solid
 * line-through, `fo:color` and `style:text-position` super/sub 58%. */
export function odsFontAttributes(font: CellFont | undefined, theme: DocumentTheme, options: { face?: boolean; omitBlack?: boolean } = {}): string {
  if (!font) {
    return "";
  }
  let out = "";
  if (font.family !== undefined) {
    const family = resolveFont(font.family, theme);
    out += options.face ? ` style:font-name="${escapeOdsAttribute(family)}"` : ` fo:font-family="${escapeOdsAttribute(odsFontFamilyValue(family))}"`;
  }
  if (font.sizePt !== undefined) {
    out += ` fo:font-size="${lengthNumber(font.sizePt)}pt"`;
  }
  if (font.bold) {
    out += ` fo:font-weight="bold"`;
  }
  if (font.italic) {
    out += ` fo:font-style="italic"`;
  }
  if (font.underline !== undefined) {
    const double = font.underline === "double" || font.underline === "doubleAccounting";
    out += ` style:text-underline-style="solid" style:text-underline-type="${double ? "double" : "single"}"`;
  }
  if (font.strike) {
    out += ` style:text-line-through-style="solid"`;
  }
  if (font.color !== undefined) {
    const color = hexColor(font.color, theme);
    if (!(options.omitBlack && color === "#000000")) {
      out += ` fo:color="${color}"`;
    }
  }
  if (font.vertAlign !== undefined) {
    out += ` style:text-position="${font.vertAlign === "superscript" ? "super" : "sub"} 58%"`;
  }
  return out;
}

/** A fill's one ODS colour: a solid fill's, a pattern's foreground (else its
 * background), a gradient's first stop — the documented approximation, ODS
 * cells having only a background colour. */
function fillColor(fill: CellFill | undefined, theme: DocumentTheme): string | undefined {
  if (!fill) {
    return undefined;
  }
  const color = fill.kind === "solid" ? fill.color : fill.kind === "pattern" ? (fill.fg ?? fill.bg) : fill.stops[0]?.color;
  return color === undefined ? undefined : hexColor(color, theme);
}

const HORIZONTAL_WRITE: Readonly<Record<string, string>> = {
  left: "start",
  center: "center",
  centerContinuous: "center",
  right: "end",
  justify: "justify",
  distributed: "justify",
};
const VERTICAL_WRITE: Readonly<Record<string, string>> = {
  top: "top",
  center: "middle",
  bottom: "bottom",
  justify: "top",
  distributed: "top",
};

/** The model's textRotation as `style:rotation-angle` (the inverse of
 * odsTextRotation): r ≤ 90 → r, 91–180 → 360 − (r − 90). */
function rotationAngle(rotation: number): number {
  return rotation <= 90 ? rotation : 360 - (rotation - 90);
}

/** A cell style's properties — the exact inverse of the reading table
 * (SPEC-122 "Styles" → ods), so a written style reads back equal: the fill's
 * colour, the four borders by the pinned write table and the diagonals,
 * horizontal alignment (`fo:text-align` + `style:text-align-source="fix"`;
 * fill → `style:repeat-content`; general → `value-type`), vertical, wrap,
 * rotation (255 → `style:direction="ttb"`), shrink, protection
 * (`style:cell-protect`), the indent as `fo:margin-left` in pt and the font.
 * centerContinuous and distributed write as center and justify, vertical
 * justify and distributed as top. */
export function odsCellStyleProperties(style: CellStyle, theme: DocumentTheme, options: { face?: boolean; omitBlack?: boolean } = {}): OdsCellStyleProperties {
  let cell = "";
  let paragraph = "";
  const alignment = style.alignment ?? {};
  const horizontal = alignment.horizontal;
  if (horizontal === "fill") {
    cell += ` style:repeat-content="true"`;
  } else if (horizontal === "general") {
    cell += ` style:text-align-source="value-type"`;
  } else if (horizontal !== undefined) {
    cell += ` style:text-align-source="fix" style:repeat-content="false"`;
    paragraph += ` fo:text-align="${HORIZONTAL_WRITE[horizontal]}"`;
  }
  const background = fillColor(style.fill, theme);
  if (background !== undefined) {
    cell += ` fo:background-color="${background}"`;
  }
  const borders = style.borders ?? {};
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const edge = borders[side];
    if (edge) {
      cell += ` fo:border-${side}="${odsBorderString(edge, theme).toLowerCase()}"`;
    }
  }
  const diagonal = borders.diagonal;
  if (diagonal) {
    const value = odsBorderString(diagonal, theme).toLowerCase();
    if (diagonal.down) {
      cell += ` style:diagonal-tl-br="${value}"`;
    }
    if (diagonal.up) {
      cell += ` style:diagonal-bl-tr="${value}"`;
    }
  }
  if (alignment.wrap) {
    cell += ` fo:wrap-option="wrap"`;
  }
  if (alignment.vertical !== undefined) {
    cell += ` style:vertical-align="${VERTICAL_WRITE[alignment.vertical]}"`;
  }
  if (alignment.textRotation === 255) {
    cell += ` style:direction="ttb"`;
  } else if (alignment.textRotation !== undefined && alignment.textRotation > 0 && alignment.textRotation <= 180) {
    cell += ` style:rotation-angle="${rotationAngle(alignment.textRotation)}"`;
  }
  if (alignment.shrinkToFit) {
    cell += ` style:shrink-to-fit="true"`;
  }
  const protection = style.protection;
  if (protection?.hidden) {
    cell += ` style:cell-protect="${protection.locked === false ? "formula-hidden" : "hidden-and-protected"}"`;
  } else if (protection?.locked === false) {
    cell += ` style:cell-protect="none"`;
  }
  if (alignment.indent !== undefined && alignment.indent > 0) {
    paragraph += ` fo:margin-left="${lengthNumber(alignment.indent * INDENT_PX_PER_LEVEL * 0.75)}pt"`;
  }
  return { cell, paragraph, text: odsFontAttributes(style.font, theme, options) };
}

/** The property elements of a style, in the schema's order. */
export function odsPropertyElements(properties: OdsCellStyleProperties): string {
  return (
    (properties.cell ? `<style:table-cell-properties${properties.cell}/>` : "") +
    (properties.paragraph ? `<style:paragraph-properties${properties.paragraph}/>` : "") +
    (properties.text ? `<style:text-properties${properties.text}/>` : "")
  );
}

/** The writers' automatic styles (SPEC-121 "Generated ODS styles", SPEC-122
 * "Styles"): `de-ta<n>` (family table: master page, `table:display`,
 * `style:writing-mode`, `tableooo:tab-color`), `de-co<n>` (a column width in
 * inches from px at 96 dpi, 4 decimals), `de-ro<n>` (a row height in pt and
 * `style:use-optimal-row-height`), and SPEC-122's `de-ce<n>` (family
 * table-cell: odsCellStyleProperties, `style:parent-style-name` the common
 * style of the style's `namedStyle` — parentStyleName, SPEC-123 — and
 * `style:data-style-name` = a `de-N<n>` for a code other than General),
 * `de-T<n>` (family text: a rich run's font) and `de-N<n>` (SPEC-118's
 * odsNumberStyleXml, its `de-N<n>P0`… sub-styles included; a
 * `builtinFormatId` written as that built-in's code in the user's locale) —
 * each family deduplicated by its properties. Each prefix's counter starts
 * past every `de-` name of that prefix already in content.xml or styles.xml
 * (a source holding `de-ce1` gets `de-ce2` next), so no generated name ever
 * collides; xml() is spliced before `</office:automatic-styles>`. */
export class OdsStyleRegistry {
  private readonly counters = new Map<string, number>();
  private readonly byKey = new Map<string, string>();
  private readonly elements: string[] = [];

  constructor(existingXml: readonly string[] = []) {
    for (const xml of existingXml) {
      for (const match of xml.matchAll(/style:name="de-([A-Za-z]+?)(\d+)(?:P\d+)?"/g)) {
        const prefix = match[1]!;
        this.counters.set(prefix, Math.max(this.counters.get(prefix) ?? 0, Number(match[2])));
      }
    }
  }

  private add(prefix: string, key: string, element: (name: string) => string): string {
    const known = this.byKey.get(`${prefix}|${key}`);
    if (known !== undefined) {
      return known;
    }
    const name = this.freshName(prefix);
    this.byKey.set(`${prefix}|${key}`, name);
    this.elements.push(element(name));
    return name;
  }

  /** The next `de-<prefix><n>` name, generating nothing — SPEC-123's
   * common data styles in styles.xml take theirs from the same counters, so
   * they never collide with a generated automatic one. */
  freshName(prefix: string): string {
    const next = (this.counters.get(prefix) ?? 0) + 1;
    this.counters.set(prefix, next);
    return `de-${prefix}${next}`;
  }

  tableStyle(style: GeneratedTableStyle): string {
    const key = `${style.masterPage}|${style.display}|${style.tabColor ?? ""}`;
    return this.add(
      "ta",
      key,
      (name) =>
        `<style:style style:name="${name}" style:family="table" style:master-page-name="${escapeOdsAttribute(style.masterPage)}">` +
        `<style:table-properties table:display="${style.display}" style:writing-mode="lr-tb"` +
        `${style.tabColor === undefined ? "" : ` tableooo:tab-color="${style.tabColor}"`}/></style:style>`,
    );
  }

  columnStyle(widthPx: number): string {
    const inches = (widthPx / 96).toFixed(4);
    return this.add(
      "co",
      inches,
      (name) => `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties style:column-width="${inches}in"/></style:style>`,
    );
  }

  rowStyle(heightPt: number, optimal: boolean): string {
    const height = `${Number(heightPt.toFixed(4))}pt`;
    return this.add(
      "ro",
      `${height}|${optimal}`,
      (name) =>
        `<style:style style:name="${name}" style:family="table-row"><style:table-row-properties style:row-height="${height}" style:use-optimal-row-height="${optimal}"/></style:style>`,
    );
  }

  /** SPEC-122: a `de-N<n>` data style for a format code (equal codes share
   * one name). */
  numberStyle(code: string): string {
    return this.add("N", code, (name) => odsNumberStyleXml(code, name));
  }

  /** SPEC-122: a `de-ce<n>` table-cell style with `style`'s properties (see
   * the class comment); equal property sets share one name. A style read
   * from the file keeps its own name — the caller's choice, not this one's. */
  cellStyle(style: CellStyle, env: OdsStyleEnv): string {
    const parent = parentStyleName(style.namedStyle, env);
    const code = formatCodeOf(style, env.locale);
    const dataStyle = code === undefined || code.trim() === "" || /^general$/i.test(code.trim()) ? undefined : this.numberStyle(code);
    const body = odsPropertyElements(odsCellStyleProperties(style, env.theme));
    const attributes =
      ` style:family="table-cell" style:parent-style-name="${escapeOdsAttribute(parent)}"` +
      (dataStyle === undefined ? "" : ` style:data-style-name="${dataStyle}"`);
    return this.add("ce", `${attributes}|${body}`, (name) =>
      body === "" ? `<style:style style:name="${name}"${attributes}/>` : `<style:style style:name="${name}"${attributes}>${body}</style:style>`,
    );
  }

  /** SPEC-122: makes an existing text style (a `runTextStyles` entry of the
   * retained package) the name of the property set its font writes, so a
   * run read from the file keeps its span's style and write → read → write
   * is stable. The first style adopted for a set wins; nothing is written. */
  adoptTextStyle(name: string, font: CellFont, theme: DocumentTheme): void {
    const attributes = odsFontAttributes(font, theme);
    const key = `T|${attributes}`;
    if (attributes !== "" && !this.byKey.has(key)) {
      this.byKey.set(key, name);
    }
  }

  /** SPEC-122: a `de-T<n>` text style for a rich run's font — an adopted
   * style of the file when one gives an equal font — or undefined for a
   * font that sets nothing (the run is plain text). */
  textStyle(font: CellFont | undefined, theme: DocumentTheme): string | undefined {
    const attributes = odsFontAttributes(font, theme);
    if (attributes === "") {
      return undefined;
    }
    return this.add(
      "T",
      attributes,
      (name) => `<style:style style:name="${name}" style:family="text"><style:text-properties${attributes}/></style:style>`,
    );
  }

  /** Every style generated so far, in order. */
  xml(): string {
    return this.elements.join("");
  }

  get size(): number {
    return this.elements.length;
  }
}

// --- named styles on write (SPEC-123) -------------------------------------------

/** A named style's definition restricted to the aspects it includes. */
function includedDefinition(named: NamedCellStyle): CellStyle {
  const definition = named.style;
  const included = new Set(named.includes);
  const style: CellStyle = {};
  if (included.has("numberFormat")) {
    if (definition.numberFormat !== undefined) {
      style.numberFormat = definition.numberFormat;
    }
    if (definition.builtinFormatId !== undefined) {
      style.builtinFormatId = definition.builtinFormatId;
    }
  }
  if (included.has("font") && definition.font) {
    style.font = definition.font;
  }
  if (included.has("fill") && definition.fill) {
    style.fill = definition.fill;
  }
  if (included.has("borders") && definition.borders) {
    style.borders = definition.borders;
  }
  if (included.has("alignment") && definition.alignment) {
    style.alignment = definition.alignment;
  }
  if (included.has("protection") && definition.protection) {
    style.protection = definition.protection;
  }
  return style;
}

/** Normal — the workbook's "Default" style, never written as a style of
 * its own. */
function isNormal(named: NamedCellStyle): boolean {
  return named.builtinId === 0 || named.sourceStyleName === "Default" || (named.sourceStyleName === undefined && named.name === "Normal");
}

/** The target styles.xml a save writes named styles into. */
export interface OdsNamedStyleTarget {
  /** Whether it defines a common table-cell style of this name. */
  defines(name: string): boolean;
  /** Whether there is one to append to (a package may lack styles.xml). */
  appendable: boolean;
}

/** One save's named styles (SPEC-123 "Named cell styles", ods write): the
 * common style each of `model.namedStyles` is written under — Normal as
 * "Default"; one whose `sourceStyleName` the target styles.xml defines,
 * under it — and `xml`, every other one as a common style for styles.xml's
 * `office:styles` ("" when there is none): `<style:style style:name
 * style:display-name style:family="table-cell"
 * style:parent-style-name="Default">` holding the property elements of its
 * definition restricted to the aspects it includes (the font keys it lacks
 * ODS inherits from Default, as applying takes Normal's), with, for a
 * number format other than General, `style:data-style-name` naming a
 * common data style written just before it (a `de-N<n>` from the
 * registry's counter; equal codes share one). Its `style:name` is its
 * `sourceStyleName` (a stale one, the target lacking it), else the display
 * name as LibreOffice encodes it ("Heading 1" → "Heading_20_1"), with `_2`,
 * `_3`… appended while the target or this save already uses that name.
 * With no styles.xml to append to, such a style is written under
 * "Default" (a generated style keeps its properties, not the name). */
export class OdsNamedStyles {
  readonly xml: string;
  private readonly model: Pick<SpreadsheetModel, "namedStyles">;
  private readonly written = new Map<NamedCellStyle, string>();
  private readonly parents = new Map<string, string | undefined>();

  constructor(model: Pick<SpreadsheetModel, "namedStyles" | "theme">, target: OdsNamedStyleTarget, registry: OdsStyleRegistry, locale: SpreadsheetLocale) {
    this.model = model;
    const theme = model.theme ?? DEFAULT_THEME;
    const used = new Set<string>();
    const dataStyles = new Map<string, string>();
    let xml = "";
    for (const named of model.namedStyles ?? []) {
      if (isNormal(named)) {
        this.written.set(named, "Default");
        continue;
      }
      if (named.sourceStyleName !== undefined && target.defines(named.sourceStyleName)) {
        this.written.set(named, named.sourceStyleName);
        continue;
      }
      if (!target.appendable) {
        this.written.set(named, "Default");
        continue;
      }
      const base = named.sourceStyleName ?? (encodeOdsStyleName(named.name) || "Style");
      let name = base;
      for (let n = 2; target.defines(name) || used.has(name); n += 1) {
        name = `${base}_${n}`;
      }
      used.add(name);
      this.written.set(named, name);
      const style = includedDefinition(named);
      const code = formatCodeOf(style, locale);
      let dataStyle: string | undefined;
      if (code !== undefined && code.trim() !== "" && !/^general$/i.test(code.trim())) {
        dataStyle = dataStyles.get(code);
        if (dataStyle === undefined) {
          dataStyle = registry.freshName("N");
          dataStyles.set(code, dataStyle);
          xml += odsNumberStyleXml(code, dataStyle);
        }
      }
      const body = odsPropertyElements(odsCellStyleProperties(style, theme));
      const attributes =
        ` style:name="${escapeOdsAttribute(name)}" style:display-name="${escapeOdsAttribute(named.name)}" style:family="table-cell" style:parent-style-name="Default"` +
        (dataStyle === undefined ? "" : ` style:data-style-name="${dataStyle}"`);
      xml += body === "" ? `<style:style${attributes}/>` : `<style:style${attributes}>${body}</style:style>`;
    }
    this.xml = xml;
  }

  /** The common style a named style the workbook defines is written under
   * (a generated style's parent); undefined for a name it does not define. */
  parent(name: string): string | undefined {
    if (!this.parents.has(name)) {
      const named = workbookNamedStyle(this.model, name);
      this.parents.set(name, named === undefined ? undefined : this.written.get(named));
    }
    return this.parents.get(name);
  }
}

function escapeOdsAttribute(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
