// The .xlsx half of SPEC-136's page setup ("xlsx wire mapping"): a
// worksheet's `printOptions`, `pageMargins`, `pageSetup`, `headerFooter`,
// `rowBreaks`/`colBreaks` and `sheetPr/pageSetUpPr@fitToPage` read into
// Sheet.pageSetup, and spliced back.
//
// - Reading. Margins are inches (cm = in × 2.54, to 1e-6); `paperSize` maps
//   through pageSetup.ts's catalogue (an unknown code kept, Letter-sized),
//   `paperWidth`/`paperHeight` ("210mm", "8.5in") are a custom paper, and a
//   `<pageSetup>` without either is Letter, the schema default. Fit mode is
//   `pageSetUpPr@fitToPage="1"` with `fitToWidth`/`fitToHeight` (default 1,
//   0 = automatic); `firstPageNumber` counts only with
//   `useFirstPageNumber="1"`; a break is a `brk@man="1"`, its `id` the 0-based
//   row (column) the next page starts at. Defaults normalise away
//   (normalizePageSetup): a sheet whose part says nothing has no setup.
// - Splicing. spliceXlsxPageSetup changes only what differs from what the
//   part already reads as — attribute by attribute in place (setAttribute),
//   every unmodelled attribute (`r:id` to printerSettings, `horizontalDpi`,
//   `copies`, `blackAndWhite`, `draft`, `cellComments`, `errors`,
//   `gridLinesSet` …) keeping its bytes and position (the SPEC-104 `@show`
//   precedent); a missing element inserted at its CT_Worksheet place; a
//   header or footer whose text reads the same keeping its bytes. Margins are
//   written to 1/1000 inch (a sanctioned delta). An unchanged setup returns
//   the very string.
import { CM_PER_INCH } from "../../core/units";
import { formatHeaderFooter, parseHeaderFooter } from "../headerFooter";
import type { HeaderFooterPair, HeaderFooterSections, SheetHeaderFooter, SheetMargins, SheetPageSetup } from "../model";
import { DEFAULT_SHEET_MARGINS, LETTER_PAPER_CODE, normalizePageSetup, paperForXlsxCode, samePageSetup, xlsxCodeOfPaper } from "../pageSetup";
import { applySplices, decodeXmlText, escapeElementText, setAttribute, type Splice } from "../xmlSplice";
import { worksheetTail } from "./drawingXml";
import { scanAttr, type XmlScanAttribute } from "../../core/xmlScanner";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const MARGIN_ATTRIBUTES: readonly [keyof SheetMargins, string][] = [
  ["leftCm", "left"],
  ["rightCm", "right"],
  ["topCm", "top"],
  ["bottomCm", "bottom"],
  ["headerCm", "header"],
  ["footerCm", "footer"],
];
const HF_CHILDREN: readonly [keyof Pick<SheetHeaderFooter, "odd" | "even" | "first">, keyof HeaderFooterPair, string][] = [
  ["odd", "header", "oddHeader"],
  ["odd", "footer", "oddFooter"],
  ["even", "header", "evenHeader"],
  ["even", "footer", "evenFooter"],
  ["first", "header", "firstHeader"],
  ["first", "footer", "firstFooter"],
];
const PAGE_ELEMENTS = ["printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks"] as const;
type PageElement = (typeof PAGE_ELEMENTS)[number];

function truthy(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function inchesToCm(inches: number): number {
  return Math.round(inches * CM_PER_INCH * 1e6) / 1e6;
}

/** Centimetres as inches to 1/1000, trailing zeros dropped ("0.7"). */
export function formatInches(cm: number): string {
  return String(Math.round((cm / CM_PER_INCH) * 1000) / 1000);
}

/** A paperWidth/paperHeight measure ("210mm", "8.5in", "21cm") in cm. */
function measureCm(text: string | undefined): number | undefined {
  const match = text === undefined ? null : /^\s*([0-9]*\.?[0-9]+)\s*(mm|cm|in)\s*$/.exec(text);
  if (match === null) {
    return undefined;
  }
  const value = Number(match[1]);
  return match[2] === "mm" ? value / 10 : match[2] === "in" ? value * CM_PER_INCH : value;
}

function formatMm(cm: number): string {
  return `${Math.round(cm * 100) / 10}mm`;
}

function wholeNumber(text: string | undefined): number | undefined {
  if (text === undefined || !/^\s*-?\d+\s*$/.test(text)) {
    return undefined;
  }
  return Number(text);
}

// --- locating -----------------------------------------------------------------------------------

interface Located {
  local: string;
  start: number;
  end: number;
  attrs: XmlScanAttribute[];
}

interface PageLayout {
  prefix: string;
  rootClose: number;
  /** Every root child after `<sheetData>`, with its tail rank. */
  children: (Located & { rank: number; main: boolean })[];
  elements: Partial<Record<PageElement, Located>>;
  /** The `<sheetPr>` element (start..end) and its `<pageSetUpPr>` start tag. */
  sheetPr?: { start: number; end: number; tagEnd: number; selfClosing: boolean; name: string };
  pageSetUpPr?: { start: number; end: number };
}

const SHEET_PR = /<((?:[\w.-]+:)?)sheetPr\b(?:[^>"']|"[^"]*"|'[^']*')*?(\/?)>/;

function locate(xml: string): PageLayout {
  const tail = worksheetTail(xml);
  const layout: PageLayout = { prefix: tail.prefix, rootClose: tail.rootClose, children: tail.children, elements: {} };
  for (const child of tail.children) {
    if (child.main && (PAGE_ELEMENTS as readonly string[]).includes(child.local) && layout.elements[child.local as PageElement] === undefined) {
      layout.elements[child.local as PageElement] = child;
    }
  }
  const dataAt = xml.search(/<(?:[\w.-]+:)?sheetData\b/);
  const head = dataAt < 0 ? xml : xml.slice(0, dataAt);
  const match = SHEET_PR.exec(head);
  if (match !== null) {
    const start = match.index;
    const tagEnd = start + match[0].length;
    const selfClosing = match[2] === "/";
    const close = selfClosing ? -1 : head.indexOf(`</${match[1]}sheetPr>`, tagEnd);
    const end = selfClosing ? tagEnd : close < 0 ? tagEnd : close + `</${match[1]}sheetPr>`.length;
    layout.sheetPr = { start, end, tagEnd, selfClosing, name: `${match[1]}sheetPr` };
    if (!selfClosing && close >= 0) {
      const inner = /<(?:[\w.-]+:)?pageSetUpPr\b(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/.exec(head.slice(tagEnd, close));
      if (inner !== null) {
        layout.pageSetUpPr = { start: tagEnd + inner.index, end: tagEnd + inner.index + inner[0].length };
      }
    }
  }
  return layout;
}

// --- reading ----------------------------------------------------------------------------------

function attrsOf(located: Located | undefined): (name: string) => string | undefined {
  return (name) => (located === undefined ? undefined : scanAttr(located.attrs, name));
}

const HF_CHILD = /<((?:[\w.-]+:)?)(oddHeader|oddFooter|evenHeader|evenFooter|firstHeader|firstFooter)\b(?:[^>"']|"[^"]*"|'[^']*')*?(?:\/>|>([\s\S]*?)<\/\1\2\s*>)/g;

/** The text of each header/footer child of a `<headerFooter>` slice. */
function headerFooterTexts(slice: string): Map<string, { text: string; start: number; end: number }> {
  const texts = new Map<string, { text: string; start: number; end: number }>();
  for (const match of slice.matchAll(HF_CHILD)) {
    if (!texts.has(match[2]!)) {
      texts.set(match[2]!, { text: decodeXmlText(match[3] ?? ""), start: match.index, end: match.index + match[0].length });
    }
  }
  return texts;
}

function readHeaderFooter(xml: string, located: Located): SheetHeaderFooter {
  const attr = attrsOf(located);
  const texts = headerFooterTexts(xml.slice(located.start, located.end));
  const hf: SheetHeaderFooter = { odd: {} };
  for (const [which, part, local] of HF_CHILDREN) {
    const text = texts.get(local)?.text;
    if (text === undefined || text === "") {
      continue;
    }
    const pair = (hf[which] ??= {});
    pair[part] = parseHeaderFooter(text);
  }
  if (truthy(attr("differentOddEven"))) {
    hf.differentOddEven = true;
  }
  if (truthy(attr("differentFirst"))) {
    hf.differentFirst = true;
  }
  if (attr("scaleWithDoc") === "0" || attr("scaleWithDoc") === "false") {
    hf.scaleWithDoc = false;
  }
  if (attr("alignWithMargins") === "0" || attr("alignWithMargins") === "false") {
    hf.alignWithMargins = false;
  }
  return hf;
}

const BRK = /<(?:[\w.-]+:)?brk\b((?:[^>"']|"[^"]*"|'[^']*')*?)\/?>/g;

function readBreaks(slice: string): number[] {
  const out: number[] = [];
  for (const match of slice.matchAll(BRK)) {
    const attrs = match[1]!;
    const man = /\bman\s*=\s*["']([^"']*)["']/.exec(attrs)?.[1];
    const id = wholeNumber(/\bid\s*=\s*["']([^"']*)["']/.exec(attrs)?.[1]);
    if (truthy(man) && id !== undefined) {
      out.push(id);
    }
  }
  return out;
}

function readLayout(xml: string, layout: PageLayout): SheetPageSetup | undefined {
  const setup: SheetPageSetup = {};
  const options = attrsOf(layout.elements.printOptions);
  if (truthy(options("horizontalCentered"))) {
    setup.centerHorizontally = true;
  }
  if (truthy(options("verticalCentered"))) {
    setup.centerVertically = true;
  }
  if (truthy(options("gridLines"))) {
    setup.printGridLines = true;
  }
  if (truthy(options("headings"))) {
    setup.printHeadings = true;
  }
  if (layout.elements.pageMargins !== undefined) {
    const margin = attrsOf(layout.elements.pageMargins);
    const margins = { ...DEFAULT_SHEET_MARGINS };
    for (const [key, name] of MARGIN_ATTRIBUTES) {
      const value = Number(margin(name));
      if (margin(name) !== undefined && Number.isFinite(value)) {
        margins[key] = inchesToCm(value);
      }
    }
    setup.margins = margins;
  }
  const page = layout.elements.pageSetup;
  if (page !== undefined) {
    const attr = attrsOf(page);
    const code = wholeNumber(attr("paperSize"));
    const width = measureCm(attr("paperWidth"));
    const height = measureCm(attr("paperHeight"));
    if (code !== undefined) {
      setup.paper = paperForXlsxCode(code);
    } else if (width !== undefined && height !== undefined) {
      setup.paper = { widthCm: Math.min(width, height), heightCm: Math.max(width, height) };
    } else {
      setup.paper = paperForXlsxCode(LETTER_PAPER_CODE);
    }
    if (attr("orientation") === "landscape") {
      setup.orientation = "landscape";
    }
    const scale = wholeNumber(attr("scale"));
    if (scale !== undefined) {
      setup.scale = scale;
    }
    if (attr("pageOrder") === "overThenDown") {
      setup.pageOrder = "overThenDown";
    }
    const first = wholeNumber(attr("firstPageNumber"));
    if (first !== undefined && truthy(attr("useFirstPageNumber"))) {
      setup.firstPageNumber = first;
    }
  }
  if (layout.pageSetUpPr !== undefined && /\bfitToPage\s*=\s*["'](?:1|true)["']/.test(xml.slice(layout.pageSetUpPr.start, layout.pageSetUpPr.end))) {
    const attr = attrsOf(page);
    setup.fitTo = { width: wholeNumber(attr("fitToWidth")) ?? 1, height: wholeNumber(attr("fitToHeight")) ?? 1 };
  }
  if (layout.elements.headerFooter !== undefined) {
    setup.headerFooter = readHeaderFooter(xml, layout.elements.headerFooter);
  }
  for (const [element, key] of [
    ["rowBreaks", "rowBreaks"],
    ["colBreaks", "colBreaks"],
  ] as const) {
    const located = layout.elements[element];
    if (located !== undefined) {
      setup[key] = readBreaks(xml.slice(located.start, located.end));
    }
  }
  return normalizePageSetup(setup);
}

/** A worksheet part's page setup (see the module header); undefined when
 * it says nothing. */
export function readXlsxPageSetup(xml: string): SheetPageSetup | undefined {
  return readLayout(xml, locate(xml));
}

// --- splicing ---------------------------------------------------------------------------------

const tailRanks = new Map<string, number>();

function tailRank(local: string): number {
  let rank = tailRanks.get(local);
  if (rank === undefined) {
    rank = worksheetTail(`<worksheet xmlns="${MAIN_NS}"><sheetData/><${local}/></worksheet>`).children.find((child) => child.local === local)?.rank ?? -1;
    tailRanks.set(local, rank);
  }
  return rank;
}

function flag(tag: string, name: string, want: boolean, had: boolean): string {
  return want === had ? tag : setAttribute(tag, name, want ? "1" : null);
}

function emptyTag(tag: string): boolean {
  return !/\s[\w:.-]+\s*=/.test(tag);
}

function printOptionsTag(tag: string, source: SheetPageSetup, setup: SheetPageSetup): string {
  let next = tag;
  next = flag(next, "horizontalCentered", setup.centerHorizontally === true, source.centerHorizontally === true);
  next = flag(next, "verticalCentered", setup.centerVertically === true, source.centerVertically === true);
  next = flag(next, "gridLines", setup.printGridLines === true, source.printGridLines === true);
  next = flag(next, "headings", setup.printHeadings === true, source.printHeadings === true);
  return next;
}

function pageMarginsTag(tag: string, source: SheetMargins, margins: SheetMargins): string {
  let next = tag;
  for (const [key, name] of MARGIN_ATTRIBUTES) {
    if (Math.abs(source[key] - margins[key]) > 1e-9 || !new RegExp(`\\s${name}\\s*=`).test(next)) {
      next = setAttribute(next, name, formatInches(margins[key]));
    }
  }
  return next;
}

function pageSetupTag(tag: string, source: SheetPageSetup, setup: SheetPageSetup): string {
  let next = tag;
  if ((setup.orientation === "landscape") !== (source.orientation === "landscape")) {
    next = setAttribute(next, "orientation", setup.orientation === "landscape" ? "landscape" : null);
  }
  if (JSON.stringify(setup.paper ?? null) !== JSON.stringify(source.paper ?? null) && setup.paper !== undefined) {
    const code = xlsxCodeOfPaper(setup.paper);
    if (code !== undefined) {
      next = setAttribute(setAttribute(setAttribute(next, "paperSize", String(code)), "paperWidth", null), "paperHeight", null);
    } else {
      next = setAttribute(setAttribute(setAttribute(next, "paperSize", null), "paperWidth", formatMm(setup.paper.widthCm)), "paperHeight", formatMm(setup.paper.heightCm));
    }
  }
  if ((setup.scale ?? 100) !== (source.scale ?? 100)) {
    next = setAttribute(next, "scale", setup.scale === undefined || setup.scale === 100 ? null : String(setup.scale));
  }
  const fit = writtenFit(setup);
  if (fit !== undefined && JSON.stringify(fit) !== JSON.stringify(source.fitTo ?? null)) {
    next = setAttribute(next, "fitToWidth", fit.width === 1 ? null : String(fit.width));
    next = setAttribute(next, "fitToHeight", fit.height === 1 ? null : String(fit.height));
  }
  if ((setup.pageOrder === "overThenDown") !== (source.pageOrder === "overThenDown")) {
    next = setAttribute(next, "pageOrder", setup.pageOrder === "overThenDown" ? "overThenDown" : null);
  }
  if (setup.firstPageNumber !== source.firstPageNumber) {
    next =
      setup.firstPageNumber === undefined
        ? setAttribute(next, "useFirstPageNumber", null)
        : setAttribute(setAttribute(next, "firstPageNumber", String(setup.firstPageNumber)), "useFirstPageNumber", "1");
  }
  return next;
}

/** The fit an xlsx part can hold: SPEC-136's `odsScaleToPages` N writes as
 * fit 1 × N (a sanctioned delta). */
function writtenFit(setup: SheetPageSetup): { width: number; height: number } | undefined {
  if (setup.fitTo !== undefined) {
    return setup.fitTo;
  }
  return setup.odsScaleToPages !== undefined ? { width: 1, height: setup.odsScaleToPages } : undefined;
}

function sameSections(a: HeaderFooterSections | undefined, b: HeaderFooterSections | undefined): boolean {
  return formatHeaderFooter(a ?? {}) === formatHeaderFooter(b ?? {});
}

function headerFooterXml(xml: string, located: Located | undefined, source: SheetHeaderFooter | undefined, hf: SheetHeaderFooter, prefix: string): string {
  let tag = located === undefined ? `<${prefix}headerFooter>` : xml.slice(located.start, xml.indexOf(">", located.start) + 1);
  const selfClosing = tag.endsWith("/>");
  if (selfClosing) {
    tag = `${tag.slice(0, -2).trimEnd()}>`;
  }
  tag = flag(tag, "differentOddEven", hf.differentOddEven === true, source?.differentOddEven === true);
  tag = flag(tag, "differentFirst", hf.differentFirst === true, source?.differentFirst === true);
  if ((hf.scaleWithDoc === false) !== (source?.scaleWithDoc === false)) {
    tag = setAttribute(tag, "scaleWithDoc", hf.scaleWithDoc === false ? "0" : null);
  }
  if ((hf.alignWithMargins === false) !== (source?.alignWithMargins === false)) {
    tag = setAttribute(tag, "alignWithMargins", hf.alignWithMargins === false ? "0" : null);
  }
  const slice = located === undefined ? "" : xml.slice(located.start, located.end);
  const texts = headerFooterTexts(slice);
  let children = "";
  for (const [which, part, local] of HF_CHILDREN) {
    const want = hf[which]?.[part];
    const had = source?.[which]?.[part];
    const existing = texts.get(local);
    if (existing !== undefined && sameSections(want, had)) {
      children += slice.slice(existing.start, existing.end);
      continue;
    }
    const text = formatHeaderFooter(want ?? {});
    if (text !== "") {
      children += `<${prefix}${local}>${escapeElementText(text)}</${prefix}${local}>`;
    }
  }
  return `${tag}${children}</${prefix}headerFooter>`;
}

function breaksXml(local: "rowBreaks" | "colBreaks", breaks: readonly number[], prefix: string): string {
  if (breaks.length === 0) {
    return "";
  }
  const max = local === "rowBreaks" ? 16383 : 1048575;
  const items = breaks.map((id) => `<${prefix}brk id="${id}" max="${max}" man="1"/>`).join("");
  return `<${prefix}${local} count="${breaks.length}" manualBreakCount="${breaks.length}">${items}</${prefix}${local}>`;
}

/** The worksheet part with its page elements spliced to `setup` (see the
 * module header); the very string when it already reads as `setup`. */
export function spliceXlsxPageSetup(xml: string, setup: SheetPageSetup | undefined): string {
  const layout = locate(xml);
  const source = readLayout(xml, layout) ?? {};
  const want = normalizePageSetup(setup) ?? {};
  if (samePageSetup(source, want)) {
    return xml;
  }
  const p = layout.prefix;
  const splices: Splice[] = [];
  const replace = (located: Located | undefined, local: PageElement, text: string): void => {
    if (located !== undefined) {
      if (xml.slice(located.start, located.end) !== text) {
        splices.push({ start: located.start, end: located.end, text });
      }
      return;
    }
    if (text === "") {
      return;
    }
    const rank = tailRank(local);
    const at = layout.children.find((child) => child.rank > rank)?.start ?? layout.rootClose;
    splices.push({ start: at, end: at, text });
  };
  const startTag = (located: Located | undefined, local: string): { tag: string; rest: string } => {
    if (located === undefined) {
      return { tag: `<${p}${local}/>`, rest: "" };
    }
    const slice = xml.slice(located.start, located.end);
    const tagEnd = slice.indexOf(">") + 1;
    return { tag: slice.slice(0, tagEnd), rest: slice.slice(tagEnd) };
  };

  // printOptions
  const options = startTag(layout.elements.printOptions, "printOptions");
  const optionsTag = printOptionsTag(options.tag, source, want);
  replace(layout.elements.printOptions, "printOptions", layout.elements.printOptions === undefined && emptyTag(optionsTag) ? "" : optionsTag + options.rest);

  // pageMargins
  const sourceMargins = source.margins ?? DEFAULT_SHEET_MARGINS;
  const margins = want.margins ?? DEFAULT_SHEET_MARGINS;
  if (layout.elements.pageMargins !== undefined || want.margins !== undefined) {
    const located = startTag(layout.elements.pageMargins, "pageMargins");
    replace(layout.elements.pageMargins, "pageMargins", pageMarginsTag(located.tag, sourceMargins, margins) + located.rest);
  }

  // pageSetup
  const needsPageSetup =
    want.paper !== undefined ||
    want.orientation !== undefined ||
    want.scale !== undefined ||
    writtenFit(want) !== undefined ||
    want.pageOrder !== undefined ||
    want.firstPageNumber !== undefined;
  if (layout.elements.pageSetup !== undefined || needsPageSetup) {
    const located = startTag(layout.elements.pageSetup, "pageSetup");
    replace(layout.elements.pageSetup, "pageSetup", pageSetupTag(located.tag, source, want) + located.rest);
  }

  // headerFooter
  if (want.headerFooter !== undefined) {
    replace(layout.elements.headerFooter, "headerFooter", headerFooterXml(xml, layout.elements.headerFooter, source.headerFooter, want.headerFooter, p));
  } else if (layout.elements.headerFooter !== undefined && source.headerFooter !== undefined) {
    replace(layout.elements.headerFooter, "headerFooter", "");
  }

  // rowBreaks, colBreaks
  for (const local of ["rowBreaks", "colBreaks"] as const) {
    if (JSON.stringify(want[local] ?? []) !== JSON.stringify(source[local] ?? [])) {
      replace(layout.elements[local], local, breaksXml(local, want[local] ?? [], p));
    }
  }

  // sheetPr/pageSetUpPr@fitToPage
  const fit = writtenFit(want) !== undefined;
  if (fit !== (source.fitTo !== undefined)) {
    splices.push(...fitToPageSplices(xml, layout, fit, p));
  }
  return applySplices(xml, splices);
}

function fitToPageSplices(xml: string, layout: PageLayout, fit: boolean, prefix: string): Splice[] {
  if (layout.pageSetUpPr !== undefined) {
    const tag = xml.slice(layout.pageSetUpPr.start, layout.pageSetUpPr.end);
    return [{ start: layout.pageSetUpPr.start, end: layout.pageSetUpPr.end, text: setAttribute(tag, "fitToPage", fit ? "1" : null) }];
  }
  if (!fit) {
    return [];
  }
  const element = `<${prefix}pageSetUpPr fitToPage="1"/>`;
  const sheetPr = layout.sheetPr;
  if (sheetPr === undefined) {
    const root = /<(?:[\w.-]+:)?worksheet\b(?:[^>"']|"[^"]*"|'[^']*')*>/.exec(xml);
    const at = root === null ? 0 : root.index + root[0].length;
    return [{ start: at, end: at, text: `<${prefix}sheetPr>${element}</${prefix}sheetPr>` }];
  }
  if (sheetPr.selfClosing) {
    const tag = xml.slice(sheetPr.start, sheetPr.end);
    return [{ start: sheetPr.start, end: sheetPr.end, text: `${tag.slice(0, -2).trimEnd()}>${element}</${sheetPr.name}>` }];
  }
  // CT_SheetPr: tabColor, outlinePr, pageSetUpPr — the last of the three.
  const closeAt = xml.lastIndexOf(`</${sheetPr.name}`, sheetPr.end);
  return [{ start: closeAt, end: closeAt, text: element }];
}
