// The .ods half of SPEC-136's page setup ("ods wire mapping"): a table's
// automatic style names its `style:master-page-name`; the master page (in
// styles.xml) names its `style:page-layout-name`, whose
// `style:page-layout-properties` carry the paper, margins, orientation,
// scale (`style:scale-to`, `style:scale-to-X`/`-Y` in the `style` or `loext`
// namespace, `style:scale-to-pages`), `style:print` (`grid`, `headers`),
// `style:print-page-order`, `style:first-page-number` and
// `style:table-centering`; its header and footer are the master page's
// `style:header`/`style:footer` (`-left` = even pages, `-first` = the first
// page; `style:display="false"` = none), their regions code text through
// headerFooter.ts's ODS_FIELD_CODES. Print ranges are the start tag's
// `table:print-ranges` (`_xlnm.Print_Area`), print titles the rows inside
// `table:table-header-rows` and the columns inside `table:table-header-
// columns` (`_xlnm.Print_Titles`), manual breaks the rows and columns whose
// style has `fo:break-before="page"`.
//
// - Margins fold the header and footer in: shown, `headerCm` = the page's
//   `fo:margin-top` and `topCm` = that + the header's `fo:min-height` + its
//   `fo:margin-bottom`; hidden, `topCm` = `fo:margin-top` (the footer the
//   same, mirrored). Written back as `fo:margin-top` = `headerCm`, the
//   header's `fo:margin-bottom` = 0.25cm and `fo:min-height` =
//   max(0, topCm − headerCm − 0.25cm). A page layout stating no margin at all
//   reads as the defaults.
// - Writing never edits a master page another table may share: a table
//   whose setup differs from what its master page reads as gets a minted
//   `de-pm{n}` page layout and `de-mp{n}` master page (copies of its own,
//   their modelled attributes and changed headers spliced), appended to
//   styles.xml's automatic and master styles, and a `de-ta{n}` table style
//   pointing at them; every other byte of styles.xml stays. A header region
//   whose sections are unchanged keeps its bytes; a changed one is
//   regenerated from codes (an unmapped field — `text:title`,
//   `text:author-name` — then becomes the text it showed: a sanctioned delta).
import { parseXml } from "../../core/xml";
import { escapeHeaderFooterText, formatHeaderFooter, ODS_FIELD_CODES, tokenizeHeaderFooter } from "../headerFooter";
import type { DefinedName, HeaderFooterPair, HeaderFooterSections, Sheet, SheetHeaderFooter, SheetMargins, SheetPageSetup, SpreadsheetModel } from "../model";
import {
  DEFAULT_SHEET_MARGINS,
  normalizePageSetup,
  paperEntryOf,
  printAreaFormula,
  printAreaOf,
  printTitlesFormula,
  PRINT_AREA,
  PRINT_TITLES,
  referenceRanges,
  samePageSetup,
  type PrintTitles,
} from "../pageSetup";
import { RESERVED_PREFIX } from "../names";
import { attributeValue, escapeAttributeValue, setAttribute } from "../xmlSplice";
import { setTableAttribute, tableStyleNameOf } from "./package";
import { canonicalToOdsRangeAddress, odsRangeAddressToCanonical } from "./rangeAddress";

const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const LOEXT_NS = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const PREFIXES: Readonly<Record<string, string>> = {
  [STYLE_NS]: "style",
  "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0": "fo",
  "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0": "svg",
  [LOEXT_NS]: "loext",
  "urn:oasis:names:tc:opendocument:xmlns:table:1.0": "table",
};
const HEADER_GAP_CM = 0.25;

type Slot = "header" | "header-left" | "header-first" | "footer" | "footer-left" | "footer-first";
const SLOTS: readonly Slot[] = ["header", "header-left", "header-first", "footer", "footer-left", "footer-first"];
const SLOT_OF: Readonly<Record<Slot, [keyof Pick<SheetHeaderFooter, "odd" | "even" | "first">, keyof HeaderFooterPair]>> = {
  header: ["odd", "header"],
  "header-left": ["even", "header"],
  "header-first": ["first", "header"],
  footer: ["odd", "footer"],
  "footer-left": ["even", "footer"],
  "footer-first": ["first", "footer"],
};

interface PageLayoutInfo {
  props: Map<string, string>;
  header?: Map<string, string>;
  footer?: Map<string, string>;
}

interface RegionInfo {
  shown: boolean;
  sections: HeaderFooterSections;
}

interface MasterPageInfo {
  layout?: string;
  slots: Partial<Record<Slot, RegionInfo>>;
}

/** styles.xml's page layouts, master pages and text styles. */
export interface OdsPageStyles {
  layouts: Map<string, PageLayoutInfo>;
  masters: Map<string, MasterPageInfo>;
  textStyles: Map<string, { bold: boolean; italic: boolean }>;
}

// --- lengths ---------------------------------------------------------------------------------

const CM_PER: Readonly<Record<string, number>> = { cm: 1, mm: 0.1, in: 2.54, pt: 2.54 / 72, pc: 2.54 / 6, px: 2.54 / 96 };

/** An ODF length in cm, undefined when unreadable. */
export function odsLengthCm(text: string | undefined): number | undefined {
  const match = text === undefined ? null : /^\s*(-?[0-9]*\.?[0-9]+)\s*(cm|mm|in|pt|pc|px)\s*$/.exec(text);
  return match === null ? undefined : Number(match[1]) * CM_PER[match[2]!]!;
}

function round(value: number, places: number): number {
  const factor = 10 ** places;
  return Math.round(value * factor) / factor;
}

function cmText(cm: number): string {
  return `${round(cm, 4)}cm`;
}

// --- reading styles.xml --------------------------------------------------------------------------

function propsOf(element: Element | undefined): Map<string, string> {
  const props = new Map<string, string>();
  if (element === undefined) {
    return props;
  }
  for (const attr of Array.from(element.attributes)) {
    const prefix = attr.namespaceURI === null ? undefined : PREFIXES[attr.namespaceURI];
    if (prefix !== undefined) {
      props.set(`${prefix}:${attr.localName}`, attr.value);
    }
  }
  return props;
}

function childNamed(parent: Element, ns: string, local: string): Element | undefined {
  return Array.from(parent.children).find((child) => child.namespaceURI === ns && child.localName === local);
}

/** A region's (or a single-paragraph header's) paragraphs as code text. */
function paragraphsCode(element: Element, textStyles: OdsPageStyles["textStyles"]): string {
  return Array.from(element.children)
    .filter((child) => child.namespaceURI === TEXT_NS && (child.localName === "p" || child.localName === "h"))
    .map((paragraph) => inlineCode(paragraph, textStyles))
    .join("\n");
}

function inlineCode(node: Element, textStyles: OdsPageStyles["textStyles"]): string {
  let code = "";
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType === 3) {
      code += escapeHeaderFooterText(child.nodeValue ?? "");
      continue;
    }
    if (child.nodeType !== 1) {
      continue;
    }
    const element = child as Element;
    if (element.namespaceURI !== TEXT_NS) {
      code += escapeHeaderFooterText(element.textContent ?? "");
      continue;
    }
    const field = ODS_FIELD_CODES[element.localName];
    if (field !== undefined) {
      code += field;
    } else if (element.localName === "span") {
      const style = textStyles.get(element.getAttributeNS(TEXT_NS, "style-name") ?? "");
      const toggles = `${style?.bold ? "&B" : ""}${style?.italic ? "&I" : ""}`;
      code += toggles + inlineCode(element, textStyles) + toggles;
    } else if (element.localName === "s") {
      code += " ".repeat(Math.max(1, Number(element.getAttributeNS(TEXT_NS, "c") ?? "1") || 1));
    } else if (element.localName === "tab") {
      code += "\t";
    } else if (element.localName === "line-break") {
      code += "\n";
    } else {
      // An unmapped field (text:title, text:author-name …): the text it shows.
      code += escapeHeaderFooterText(element.textContent ?? "");
    }
  }
  return code;
}

function regionInfo(element: Element, textStyles: OdsPageStyles["textStyles"]): RegionInfo {
  const shown = element.getAttributeNS(STYLE_NS, "display") !== "false";
  const sections: HeaderFooterSections = {};
  const regions: [keyof HeaderFooterSections, string][] = [
    ["left", "region-left"],
    ["center", "region-center"],
    ["right", "region-right"],
  ];
  let regioned = false;
  for (const [key, local] of regions) {
    const region = childNamed(element, STYLE_NS, local);
    if (region !== undefined) {
      regioned = true;
      const code = paragraphsCode(region, textStyles);
      if (code !== "") {
        sections[key] = code;
      }
    }
  }
  if (!regioned) {
    const code = paragraphsCode(element, textStyles);
    if (code !== "") {
      sections.center = code;
    }
  }
  return { shown, sections };
}

/** styles.xml's page styles (see OdsPageStyles); empty when unreadable. */
export function readOdsPageStyles(stylesXml: string | undefined): OdsPageStyles {
  const styles: OdsPageStyles = { layouts: new Map(), masters: new Map(), textStyles: new Map() };
  if (stylesXml === undefined || !/page-layout|master-page/.test(stylesXml)) {
    return styles;
  }
  let doc: Document;
  try {
    doc = parseXml(stylesXml);
  } catch {
    return styles;
  }
  for (const layout of Array.from(doc.getElementsByTagNameNS(STYLE_NS, "page-layout"))) {
    const name = layout.getAttributeNS(STYLE_NS, "name");
    if (name === null) {
      continue;
    }
    const info: PageLayoutInfo = { props: propsOf(childNamed(layout, STYLE_NS, "page-layout-properties")) };
    const header = childNamed(layout, STYLE_NS, "header-style");
    const footer = childNamed(layout, STYLE_NS, "footer-style");
    if (header !== undefined) {
      info.header = propsOf(childNamed(header, STYLE_NS, "header-footer-properties"));
    }
    if (footer !== undefined) {
      info.footer = propsOf(childNamed(footer, STYLE_NS, "header-footer-properties"));
    }
    styles.layouts.set(name, info);
  }
  for (const style of Array.from(doc.getElementsByTagNameNS(STYLE_NS, "style"))) {
    if (style.getAttributeNS(STYLE_NS, "family") !== "text") {
      continue;
    }
    const props = propsOf(childNamed(style, STYLE_NS, "text-properties"));
    styles.textStyles.set(style.getAttributeNS(STYLE_NS, "name") ?? "", {
      bold: /^(?:bold|[6-9]00)$/.test(props.get("fo:font-weight") ?? ""),
      italic: /^(?:italic|oblique)$/.test(props.get("fo:font-style") ?? ""),
    });
  }
  for (const master of Array.from(doc.getElementsByTagNameNS(STYLE_NS, "master-page"))) {
    const name = master.getAttributeNS(STYLE_NS, "name");
    if (name === null) {
      continue;
    }
    const info: MasterPageInfo = { slots: {} };
    const layout = master.getAttributeNS(STYLE_NS, "page-layout-name");
    if (layout !== null) {
      info.layout = layout;
    }
    for (const child of Array.from(master.children)) {
      const slot = child.localName as Slot;
      if ((child.namespaceURI === STYLE_NS || child.namespaceURI === LOEXT_NS) && SLOTS.includes(slot) && info.slots[slot] === undefined) {
        info.slots[slot] = regionInfo(child, styles.textStyles);
      }
    }
    styles.masters.set(name, info);
  }
  return styles;
}

// --- a master page as a setup ------------------------------------------------------------------

function whole(text: string | undefined): number | undefined {
  return text !== undefined && /^\s*\d+\s*$/.test(text) ? Number(text) : undefined;
}

function hasText(region: RegionInfo | undefined): boolean {
  return region !== undefined && region.shown && Object.keys(region.sections).length > 0;
}

/** What a master page reads as (the layout fields of a setup; no breaks). */
export function odsMasterPageSetup(styles: OdsPageStyles, masterName: string): SheetPageSetup | undefined {
  const master = styles.masters.get(masterName);
  if (master === undefined) {
    return undefined;
  }
  const layout = master.layout === undefined ? undefined : styles.layouts.get(master.layout);
  const props = layout?.props ?? new Map<string, string>();
  const setup: SheetPageSetup = {};
  const width = odsLengthCm(props.get("fo:page-width"));
  const height = odsLengthCm(props.get("fo:page-height"));
  if (width !== undefined && height !== undefined && width > 0 && height > 0) {
    const portrait = { widthCm: round(Math.min(width, height), 3), heightCm: round(Math.max(width, height), 3) };
    const entry = paperEntryOf(portrait);
    setup.paper = entry === undefined ? portrait : { widthCm: entry.widthCm, heightCm: entry.heightCm, xlsxPaperSize: entry.code };
    const orientation = props.get("style:print-orientation");
    if (orientation === "landscape" || (orientation === undefined && width > height)) {
      setup.orientation = "landscape";
    }
  }
  const fitX = props.get("style:scale-to-X") ?? props.get("loext:scale-to-X");
  const fitY = props.get("style:scale-to-Y") ?? props.get("loext:scale-to-Y");
  const pages = whole(props.get("style:scale-to-pages"));
  const percent = /^\s*([0-9.]+)\s*%\s*$/.exec(props.get("style:scale-to") ?? "");
  if (fitX !== undefined || fitY !== undefined) {
    setup.fitTo = { width: whole(fitX) ?? 0, height: whole(fitY) ?? 0 };
  } else if (pages !== undefined && pages > 0) {
    setup.odsScaleToPages = pages;
  } else if (percent !== null) {
    setup.scale = Math.round(Number(percent[1]));
  }
  const print = (props.get("style:print") ?? "").split(/\s+/);
  if (print.includes("grid")) {
    setup.printGridLines = true;
  }
  if (print.includes("headers")) {
    setup.printHeadings = true;
  }
  if (props.get("style:print-page-order") === "ltr") {
    setup.pageOrder = "overThenDown";
  }
  const first = whole(props.get("style:first-page-number"));
  if (first !== undefined) {
    setup.firstPageNumber = first;
  }
  const centering = props.get("style:table-centering");
  if (centering === "horizontal" || centering === "both") {
    setup.centerHorizontally = true;
  }
  if (centering === "vertical" || centering === "both") {
    setup.centerVertically = true;
  }
  const margins = foldedMargins(props, layout, master);
  if (margins !== undefined) {
    setup.margins = margins;
  }
  const hf: SheetHeaderFooter = { odd: {} };
  for (const slot of SLOTS) {
    const region = master.slots[slot];
    if (!hasText(region)) {
      continue;
    }
    const [which, part] = SLOT_OF[slot];
    (hf[which] ??= {})[part] = { ...region!.sections };
  }
  if (master.slots["header-left"]?.shown || master.slots["footer-left"]?.shown) {
    hf.differentOddEven = true;
  }
  if (master.slots["header-first"]?.shown || master.slots["footer-first"]?.shown) {
    hf.differentFirst = true;
  }
  setup.headerFooter = hf;
  return normalizePageSetup(setup);
}

function foldedMargins(props: Map<string, string>, layout: PageLayoutInfo | undefined, master: MasterPageInfo): SheetMargins | undefined {
  const all = odsLengthCm(props.get("fo:margin"));
  const side = (name: string): number | undefined => odsLengthCm(props.get(`fo:margin-${name}`)) ?? all;
  const [top, bottom, left, right] = [side("top"), side("bottom"), side("left"), side("right")];
  if (top === undefined && bottom === undefined && left === undefined && right === undefined) {
    return undefined;
  }
  const headerShown = master.slots.header?.shown === true;
  const footerShown = master.slots.footer?.shown === true;
  const extent = (style: Map<string, string> | undefined, gap: "fo:margin-bottom" | "fo:margin-top"): number =>
    (odsLengthCm(style?.get("fo:min-height") ?? style?.get("svg:height")) ?? 0) + (odsLengthCm(style?.get(gap)) ?? 0);
  const t = top ?? 0;
  const b = bottom ?? 0;
  return {
    topCm: round(headerShown ? t + extent(layout?.header, "fo:margin-bottom") : t, 4),
    rightCm: round(right ?? 0, 4),
    bottomCm: round(footerShown ? b + extent(layout?.footer, "fo:margin-top") : b, 4),
    leftCm: round(left ?? 0, 4),
    headerCm: headerShown ? round(t, 4) : DEFAULT_SHEET_MARGINS.headerCm,
    footerCm: footerShown ? round(b, 4) : DEFAULT_SHEET_MARGINS.footerCm,
  };
}

// --- a table ----------------------------------------------------------------------------------

/** A table style's master page, from content.xml's automatic styles
 * ("Default" when it names none). */
export function odsMasterPageOf(contentPrefix: string, styleName: string | undefined): string {
  if (styleName === undefined) {
    return "Default";
  }
  for (const match of contentPrefix.matchAll(/<style:style\b[^>]*>/g)) {
    if (attributeValue(match[0], "style:name") === styleName && attributeValue(match[0], "style:family") === "table") {
      return attributeValue(match[0], "style:master-page-name") ?? "Default";
    }
  }
  return "Default";
}

/** The row and column styles of content.xml's automatic styles that break
 * the page before them. */
export function odsBreakingStyles(contentPrefix: string): { rows: Set<string>; cols: Set<string> } {
  const out = { rows: new Set<string>(), cols: new Set<string>() };
  if (!contentPrefix.includes("break-before")) {
    return out;
  }
  for (const match of contentPrefix.matchAll(/<style:style\b([^>]*?)(?:\/>|>([\s\S]*?)<\/style:style>)/g)) {
    const tag = `<style:style${match[1]}>`;
    const family = attributeValue(tag, "style:family");
    const name = attributeValue(tag, "style:name");
    if (name === undefined || !/fo:break-before\s*=\s*"page"/.test(match[2] ?? "")) {
      continue;
    }
    if (family === "table-row") {
      out.rows.add(name);
    } else if (family === "table-column") {
      out.cols.add(name);
    }
  }
  return out;
}

const TABLE_ITEM = /<(\/?)table:(table-row|table-header-rows|table-column|table-header-columns)\b([^>]*)>/g;

/** The header rows and columns and the broken rows and columns of a table's
 * text (rows and columns counted with their repeats). */
export function odsTableStructure(tableXml: string, breaking: { rows: ReadonlySet<string>; cols: ReadonlySet<string> }): {
  titles: PrintTitles;
  rowBreaks: number[];
  colBreaks: number[];
} {
  const out = { titles: {} as PrintTitles, rowBreaks: [] as number[], colBreaks: [] as number[] };
  if (!tableXml.includes("table-header-") && breaking.rows.size === 0 && breaking.cols.size === 0) {
    return out;
  }
  let row = 0;
  let col = 0;
  let headerRowsFrom = -1;
  let headerColsFrom = -1;
  for (const match of tableXml.matchAll(TABLE_ITEM)) {
    const closing = match[1] === "/";
    const local = match[2]!;
    const attrs = match[3]!;
    if (local === "table-header-rows") {
      if (!closing) {
        headerRowsFrom = row;
      } else if (headerRowsFrom >= 0 && row > headerRowsFrom && out.titles.rows === undefined) {
        out.titles.rows = [headerRowsFrom, row - 1];
      }
      continue;
    }
    if (local === "table-header-columns") {
      if (!closing) {
        headerColsFrom = col;
      } else if (headerColsFrom >= 0 && col > headerColsFrom && out.titles.cols === undefined) {
        out.titles.cols = [headerColsFrom, col - 1];
      }
      continue;
    }
    if (closing) {
      continue;
    }
    const tag = `<x${attrs}>`;
    const style = attributeValue(tag, "table:style-name");
    if (local === "table-row") {
      const repeat = Math.max(1, Number(attributeValue(tag, "table:number-rows-repeated") ?? "1") || 1);
      if (style !== undefined && breaking.rows.has(style)) {
        out.rowBreaks.push(row);
      }
      row += repeat;
    } else {
      const repeat = Math.max(1, Number(attributeValue(tag, "table:number-columns-repeated") ?? "1") || 1);
      if (style !== undefined && breaking.cols.has(style)) {
        out.colBreaks.push(col);
      }
      col += repeat;
    }
  }
  return out;
}

/** A table start tag's print ranges as ranges; null when absent or unreadable. */
function sourcePrintRanges(tableXml: string): ReturnType<typeof referenceRanges> | null {
  const start = /<table:table\b[^>]*>/.exec(tableXml)?.[0];
  const value = start === undefined ? undefined : attributeValue(start, "table:print-ranges");
  if (value === undefined) {
    return null;
  }
  const canonical = odsRangeAddressToCanonical(value);
  return canonical === null ? null : referenceRanges(canonical);
}

/** What one table's source says: its setup (breaks included) and its print
 * area and titles as reserved names. */
export function readOdsTablePageSetup(
  sheet: Pick<Sheet, "id" | "name" | "sourceXml">,
  contentPrefix: string,
  styles: OdsPageStyles,
  breaking = odsBreakingStyles(contentPrefix),
): { setup: SheetPageSetup | undefined; names: DefinedName[] } {
  const text = sheet.sourceXml ?? "";
  const master = odsMasterPageOf(contentPrefix, tableStyleNameOf(text));
  const structure = odsTableStructure(text, breaking);
  const setup = normalizePageSetup({ ...odsMasterPageSetup(styles, master), rowBreaks: structure.rowBreaks, colBreaks: structure.colBreaks });
  const names: DefinedName[] = [];
  const area = sourcePrintRanges(text);
  if (area !== null && area.length > 0) {
    names.push({ name: RESERVED_PREFIX + PRINT_AREA, scopeSheetId: sheet.id, formula: printAreaFormula(sheet.name, area) });
  }
  if (structure.titles.rows !== undefined || structure.titles.cols !== undefined) {
    names.push({ name: RESERVED_PREFIX + PRINT_TITLES, scopeSheetId: sheet.id, formula: printTitlesFormula(sheet.name, structure.titles) });
  }
  return { setup, names };
}

/** Every source table's setup installed on its sheet; the reserved names
 * their print ranges and header rows/columns mean, returned. */
export function readOdsPageSetups(sheets: Sheet[], contentPrefix: string, stylesXml: string | undefined): DefinedName[] {
  const styles = readOdsPageStyles(stylesXml);
  const breaking = odsBreakingStyles(contentPrefix);
  const names: DefinedName[] = [];
  for (const sheet of sheets) {
    if (sheet.kind !== "grid" || sheet.sourceXml === undefined) {
      continue;
    }
    const read = readOdsTablePageSetup(sheet, contentPrefix, styles, breaking);
    if (read.setup !== undefined) {
      sheet.pageSetup = read.setup;
    }
    names.push(...read.names);
  }
  return names;
}

// --- writing ----------------------------------------------------------------------------------

function layoutFields(setup: SheetPageSetup | undefined): SheetPageSetup | undefined {
  if (setup === undefined) {
    return undefined;
  }
  const { rowBreaks: _rows, colBreaks: _cols, ...rest } = setup;
  return normalizePageSetup(rest);
}

/** A whole element of `xml` whose start tag matches `open` (by its
 * `style:name`), with its offsets. */
function elementNamed(xml: string, local: string, name: string): { start: number; end: number; text: string } | undefined {
  const pattern = new RegExp(`<style:${local}\\b(?:[^>"']|"[^"]*"|'[^']*')*?(/?)>`, "g");
  for (const match of xml.matchAll(pattern)) {
    if (attributeValue(match[0], "style:name") !== name) {
      continue;
    }
    const start = match.index;
    if (match[1] === "/") {
      return { start, end: start + match[0].length, text: match[0] };
    }
    const close = xml.indexOf(`</style:${local}>`, start);
    if (close < 0) {
      return undefined;
    }
    const end = close + `</style:${local}>`.length;
    return { start, end, text: xml.slice(start, end) };
  }
  return undefined;
}

/** A child element (first, at any depth) of `xml` by its qualified name. */
function childTag(xml: string, qname: string): { start: number; end: number; tag: string } | undefined {
  const match = new RegExp(`<${qname}\\b(?:[^>"']|"[^"]*"|'[^']*')*?/?>`).exec(xml);
  return match === null ? undefined : { start: match.index, end: match.index + match[0].length, tag: match[0] };
}

const SCALE_ATTRIBUTES = ["style:scale-to", "style:scale-to-X", "style:scale-to-Y", "loext:scale-to-X", "loext:scale-to-Y", "style:scale-to-pages"];

function pagePropertiesTag(tag: string, setup: SheetPageSetup, margins: SheetMargins, headerShown: boolean, footerShown: boolean, paper: { widthCm: number; heightCm: number } | undefined): string {
  let next = tag;
  if (paper !== undefined) {
    const landscape = setup.orientation === "landscape";
    next = setAttribute(next, "fo:page-width", cmText(landscape ? paper.heightCm : paper.widthCm));
    next = setAttribute(next, "fo:page-height", cmText(landscape ? paper.widthCm : paper.heightCm));
  }
  next = setAttribute(next, "style:print-orientation", setup.orientation === "landscape" ? "landscape" : "portrait");
  next = setAttribute(next, "fo:margin", null);
  next = setAttribute(next, "fo:margin-top", cmText(headerShown ? margins.headerCm : margins.topCm));
  next = setAttribute(next, "fo:margin-bottom", cmText(footerShown ? margins.footerCm : margins.bottomCm));
  next = setAttribute(next, "fo:margin-left", cmText(margins.leftCm));
  next = setAttribute(next, "fo:margin-right", cmText(margins.rightCm));
  for (const name of SCALE_ATTRIBUTES) {
    next = setAttribute(next, name, null);
  }
  if (setup.fitTo !== undefined) {
    next = setAttribute(setAttribute(next, "style:scale-to-X", String(setup.fitTo.width)), "style:scale-to-Y", String(setup.fitTo.height));
  } else if (setup.odsScaleToPages !== undefined) {
    next = setAttribute(next, "style:scale-to-pages", String(setup.odsScaleToPages));
  } else if (setup.scale !== undefined) {
    next = setAttribute(next, "style:scale-to", `${setup.scale}%`);
  }
  const tokens = (attributeValue(next, "style:print") ?? "").split(/\s+/).filter((token) => token !== "" && token !== "grid" && token !== "headers");
  if (setup.printGridLines) {
    tokens.push("grid");
  }
  if (setup.printHeadings) {
    tokens.push("headers");
  }
  next = setAttribute(next, "style:print", tokens.length === 0 ? null : tokens.join(" "));
  next = setAttribute(next, "style:print-page-order", setup.pageOrder === "overThenDown" ? "ltr" : "ttb");
  next = setAttribute(next, "style:first-page-number", setup.firstPageNumber === undefined ? "continue" : String(setup.firstPageNumber));
  const centering =
    setup.centerHorizontally && setup.centerVertically ? "both" : setup.centerHorizontally ? "horizontal" : setup.centerVertically ? "vertical" : "none";
  return setAttribute(next, "style:table-centering", centering);
}

function bandStyle(sourceText: string | undefined, local: "header-style" | "footer-style", shown: boolean, extentCm: number): string {
  const qname = `style:${local}`;
  if (!shown) {
    return sourceText === undefined ? "" : sourceText;
  }
  const gap = local === "header-style" ? "fo:margin-bottom" : "fo:margin-top";
  const base = sourceText ?? `<${qname}><style:header-footer-properties/></${qname}>`;
  const props = childTag(base, "style:header-footer-properties");
  if (props === undefined) {
    return `<${qname}><style:header-footer-properties fo:min-height="${cmText(extentCm)}" ${gap}="${cmText(HEADER_GAP_CM)}"/></${qname}>`;
  }
  let tag = setAttribute(props.tag, "fo:min-height", cmText(extentCm));
  tag = setAttribute(setAttribute(tag, gap, cmText(HEADER_GAP_CM)), "svg:height", null);
  return base.slice(0, props.start) + tag + base.slice(props.end);
}

/** Code text as an ODS region's paragraphs (text styles for bold/italic
 * minted through `textStyle`). */
function regionParagraphs(code: string, textStyle: (bold: boolean, italic: boolean) => string): string {
  const paragraphs: string[] = [];
  let current = "";
  let bold = false;
  let italic = false;
  let open = false;
  const close = (): void => {
    if (open) {
      current += "</text:span>";
      open = false;
    }
  };
  const restyle = (): void => {
    close();
    if (bold || italic) {
      current += `<text:span text:style-name="${textStyle(bold, italic)}">`;
      open = true;
    }
  };
  const text = (value: string): void => {
    const lines = value.split("\n");
    lines.forEach((line, index) => {
      if (index > 0) {
        close();
        paragraphs.push(current);
        current = "";
        restyle();
      }
      current += escapeAttributeValue(line).replaceAll("&quot;", '"');
    });
  };
  for (const token of tokenizeHeaderFooter(code)) {
    switch (token.kind) {
      case "text":
        text(token.text);
        break;
      case "page":
        current += "<text:page-number>1</text:page-number>";
        break;
      case "pages":
        current += "<text:page-count>1</text:page-count>";
        break;
      case "date":
        current += "<text:date/>";
        break;
      case "time":
        current += "<text:time/>";
        break;
      case "sheet":
        current += "<text:sheet-name>???</text:sheet-name>";
        break;
      case "file":
        current += "<text:file-name>???</text:file-name>";
        break;
      case "toggle":
        if (token.toggle === "bold" || token.toggle === "italic") {
          if (token.toggle === "bold") {
            bold = !bold;
          } else {
            italic = !italic;
          }
          restyle();
        }
        break;
      default:
        break;
    }
  }
  close();
  paragraphs.push(current);
  return paragraphs.map((paragraph) => `<text:p>${paragraph}</text:p>`).join("");
}

function regionXml(slot: Slot, sections: HeaderFooterSections, textStyle: (bold: boolean, italic: boolean) => string): string {
  const regions: [keyof HeaderFooterSections, string][] = [
    ["left", "region-left"],
    ["center", "region-center"],
    ["right", "region-right"],
  ];
  const inner = regions
    .filter(([key]) => (sections[key] ?? "") !== "")
    .map(([key, local]) => `<style:${local}>${regionParagraphs(sections[key]!, textStyle)}</style:${local}>`)
    .join("");
  return `<style:${slot}>${inner}</style:${slot}>`;
}

/** The master page's top-level children, with their offsets in `text`. */
function masterChildren(text: string): { local: string; start: number; end: number }[] {
  const children: { local: string; start: number; end: number }[] = [];
  const openEnd = text.indexOf(">") + 1;
  let depth = 0;
  let current: { local: string; start: number; end: number } | null = null;
  for (const match of text.slice(openEnd).matchAll(/<(\/?)([\w.-]+:[\w.-]+)\b(?:[^>"']|"[^"]*"|'[^']*')*?(\/?)>/g)) {
    const at = openEnd + match.index;
    const closing = match[1] === "/";
    const selfClosing = match[3] === "/";
    if (closing) {
      depth -= 1;
      if (depth === 0 && current !== null) {
        current.end = at + match[0].length;
        children.push(current);
        current = null;
      }
      continue;
    }
    if (depth === 0) {
      const local = match[2]!.slice(match[2]!.indexOf(":") + 1);
      if (selfClosing) {
        children.push({ local, start: at, end: at + match[0].length });
        continue;
      }
      current = { local, start: at, end: at };
    }
    if (!selfClosing) {
      depth += 1;
    }
  }
  return children.filter((child) => child.local !== "master-page");
}

/** What a save writes for each table's page setup (see the module header). */
export class OdsPageSetupWriter {
  private readonly styles: OdsPageStyles;
  private readonly stylesXml: string;
  private layouts = "";
  private masters = "";
  private texts = "";
  private readonly textStyles = new Map<string, string>();

  constructor(
    private readonly contentPrefix: string,
    stylesXml: string | undefined,
    private readonly freshName: (prefix: string) => string,
  ) {
    this.stylesXml = stylesXml ?? "";
    this.styles = readOdsPageStyles(stylesXml);
  }

  /** The master page a table's style must name for `sheet`'s setup: its
   * source one (undefined) while that reads as the setup, else a minted one. */
  masterPageFor(sheet: Sheet, styleName: string | undefined): string | undefined {
    const sourceMaster = odsMasterPageOf(this.contentPrefix, styleName);
    const source = layoutFields(odsMasterPageSetup(this.styles, sourceMaster));
    const want = layoutFields(sheet.pageSetup);
    if (samePageSetup(source, want)) {
      return undefined;
    }
    return this.mint(sourceMaster, source ?? {}, want ?? {});
  }

  /** styles.xml with every minted page layout, master page and text style. */
  stylesXmlWith(xml: string): string {
    if (this.layouts === "" && this.masters === "" && this.texts === "") {
      return xml;
    }
    let next = xml;
    const automatic = this.texts + this.layouts;
    if (automatic !== "") {
      next = insertInto(next, "office:automatic-styles", automatic, /<office:master-styles\b|<\/[\w.-]+:document-styles>/);
    }
    if (this.masters !== "") {
      next = insertInto(next, "office:master-styles", this.masters, /<\/[\w.-]+:document-styles>/);
    }
    return next;
  }

  private textStyle(bold: boolean, italic: boolean): string {
    const key = `${bold}|${italic}`;
    let name = this.textStyles.get(key);
    if (name === undefined) {
      name = this.freshName("hf");
      this.textStyles.set(key, name);
      this.texts +=
        `<style:style style:name="${name}" style:family="text"><style:text-properties` +
        `${bold ? ` fo:font-weight="bold"` : ""}${italic ? ` fo:font-style="italic"` : ""}/></style:style>`;
    }
    return name;
  }

  private mint(sourceMaster: string, source: SheetPageSetup, want: SheetPageSetup): string {
    const layoutName = this.freshName("pm");
    const masterName = this.freshName("mp");
    const master = elementNamed(this.stylesXml, "master-page", sourceMaster);
    const sourceLayoutName = master === undefined ? undefined : attributeValue(master.text, "style:page-layout-name");
    const layout = sourceLayoutName === undefined ? undefined : elementNamed(this.stylesXml, "page-layout", sourceLayoutName);
    const margins = want.margins ?? DEFAULT_SHEET_MARGINS;
    const hf = want.headerFooter;
    const anyText = (part: keyof HeaderFooterPair): boolean =>
      [hf?.odd, hf?.even, hf?.first].some((pair) => formatHeaderFooter(pair?.[part] ?? {}) !== "");
    const headerShown = anyText("header") || Math.abs(margins.headerCm - DEFAULT_SHEET_MARGINS.headerCm) > 1e-9;
    const footerShown = anyText("footer") || Math.abs(margins.footerCm - DEFAULT_SHEET_MARGINS.footerCm) > 1e-9;

    // The page layout: a copy of the source's, its modelled attributes set.
    let layoutText = layout?.text ?? `<style:page-layout style:name="${layoutName}"/>`;
    if (layoutText.endsWith("/>") && !layoutText.includes("</style:page-layout>")) {
      layoutText = `${layoutText.slice(0, -2).trimEnd()}></style:page-layout>`;
    }
    layoutText = layoutText.replace(/^<style:page-layout\b[^>]*>/, (tag) => setAttribute(tag, "style:name", layoutName));
    const props = childTag(layoutText, "style:page-layout-properties");
    const propsTag = pagePropertiesTag(props?.tag ?? "<style:page-layout-properties/>", want, margins, headerShown, footerShown, want.paper);
    layoutText =
      props === undefined
        ? layoutText.replace(/^<style:page-layout\b[^>]*>/, (tag) => tag + propsTag)
        : layoutText.slice(0, props.start) + propsTag + layoutText.slice(props.end);
    for (const [local, shown, extent] of [
      ["header-style", headerShown, Math.max(0, margins.topCm - margins.headerCm - HEADER_GAP_CM)],
      ["footer-style", footerShown, Math.max(0, margins.bottomCm - margins.footerCm - HEADER_GAP_CM)],
    ] as const) {
      const existing = elementNamedAny(layoutText, `style:${local}`);
      const text = bandStyle(existing?.text, local, shown, extent);
      layoutText =
        existing === undefined
          ? layoutText.replace("</style:page-layout>", `${text}</style:page-layout>`)
          : layoutText.slice(0, existing.start) + text + layoutText.slice(existing.end);
    }
    this.layouts += layoutText;

    // The master page: a copy of the source's; a changed header or footer regenerated.
    let masterText = master?.text ?? `<style:master-page style:name="${masterName}"/>`;
    if (masterText.endsWith("/>") && !masterText.includes("</style:master-page>")) {
      masterText = `${masterText.slice(0, -2).trimEnd()}></style:master-page>`;
    }
    const openTag = masterText.slice(0, masterText.indexOf(">") + 1);
    const children = masterChildren(masterText);
    const bySlot = new Map(children.filter((child) => SLOTS.includes(child.local as Slot)).map((child) => [child.local as Slot, child]));
    const others = children.filter((child) => !SLOTS.includes(child.local as Slot));
    let body = "";
    for (const slot of SLOTS) {
      const [which, part] = SLOT_OF[slot];
      const sections = slotSections(want, which, part);
      const had = slotSections(source, which, part);
      const existing = bySlot.get(slot);
      if (existing !== undefined && formatHeaderFooter(sections ?? {}) === formatHeaderFooter(had ?? {}) && (sections !== undefined || slotShown(source, slot) === slotShown(want, slot))) {
        body += masterText.slice(existing.start, existing.end);
        continue;
      }
      const shown = slot === "header" ? headerShown : slot === "footer" ? footerShown : sections !== undefined;
      if (!shown) {
        continue;
      }
      body += regionXml(slot, sections ?? {}, (bold, italic) => this.textStyle(bold, italic));
    }
    body += others.map((child) => masterText.slice(child.start, child.end)).join("");
    let open = setAttribute(openTag, "style:name", masterName);
    open = setAttribute(open, "style:page-layout-name", layoutName);
    this.masters += `${open}${body}</style:master-page>`;
    return masterName;
  }
}

function slotSections(setup: SheetPageSetup, which: "odd" | "even" | "first", part: keyof HeaderFooterPair): HeaderFooterSections | undefined {
  const hf = setup.headerFooter;
  if (hf === undefined) {
    return undefined;
  }
  if (which === "even" && !hf.differentOddEven) {
    return undefined;
  }
  if (which === "first" && !hf.differentFirst) {
    return undefined;
  }
  const sections = hf[which]?.[part];
  return sections === undefined || formatHeaderFooter(sections) === "" ? (which === "odd" ? undefined : {}) : sections;
}

function slotShown(setup: SheetPageSetup, slot: Slot): boolean {
  const [which, part] = SLOT_OF[slot];
  return slotSections(setup, which, part) !== undefined;
}

function elementNamedAny(xml: string, qname: string): { start: number; end: number; text: string } | undefined {
  const match = new RegExp(`<${qname}\\b(?:[^>"']|"[^"]*"|'[^']*')*?(/?)>`).exec(xml);
  if (match === null) {
    return undefined;
  }
  if (match[1] === "/") {
    return { start: match.index, end: match.index + match[0].length, text: match[0] };
  }
  const close = xml.indexOf(`</${qname}>`, match.index);
  return close < 0 ? undefined : { start: match.index, end: close + qname.length + 3, text: xml.slice(match.index, close + qname.length + 3) };
}

/** `xml` with `content` appended inside its `qname` element (expanded when
 * self-closing; created before `before` when absent). */
function insertInto(xml: string, qname: string, content: string, before: RegExp): string {
  const close = xml.lastIndexOf(`</${qname}>`);
  if (close >= 0) {
    return xml.slice(0, close) + content + xml.slice(close);
  }
  const selfClosing = new RegExp(`<${qname}\\b[^>]*?/>`).exec(xml);
  if (selfClosing !== null) {
    return `${xml.slice(0, selfClosing.index)}<${qname}>${content}</${qname}>${xml.slice(selfClosing.index + selfClosing[0].length)}`;
  }
  const at = before.exec(xml)?.index ?? xml.length;
  return `${xml.slice(0, at)}<${qname}>${content}</${qname}>${xml.slice(at)}`;
}

/** A table's text with its `table:print-ranges` set from the model's
 * `_xlnm.Print_Area` — kept as written while it names the same ranges. */
export function withOdsPrintRanges(model: Pick<SpreadsheetModel, "names">, sheet: Pick<Sheet, "id">, tableXml: string): string {
  const want = printAreaOf(model, sheet);
  const had = sourcePrintRanges(tableXml);
  if (JSON.stringify(want ?? []) === JSON.stringify(had ?? [])) {
    return tableXml;
  }
  if (want === null) {
    return setTableAttribute(tableXml, "table:print-ranges", null);
  }
  const name = model.names.find((each) => each.scopeSheetId === sheet.id && each.name.toUpperCase() === (RESERVED_PREFIX + PRINT_AREA).toUpperCase());
  const address = name === undefined ? null : canonicalToOdsRangeAddress(name.formula);
  return address === null ? tableXml : setTableAttribute(tableXml, "table:print-ranges", address);
}
