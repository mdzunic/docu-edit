// The .xlsx writer (SPEC-121 "From scratch", "The .xlsx substitution path").
// THIS FILE, with SPEC-121, is the normative source for how a workbook's
// package is rebuilt. Three paths, chosen by what the model holds (SPEC-080's
// for decks):
//
//  1. No retained package (`opaque` empty — a new workbook, or a
//     stripRetention'ed copy): the minimal package from the model, parts in
//     this order: [Content_Types].xml, _rels/.rels, docProps/core.xml,
//     docProps/app.xml (Application only — no TitlesOfParts to go stale),
//     xl/workbook.xml, its rels, xl/theme/theme1.xml (core's
//     OFFICE_THEME_XML), xl/styles.xml (one font, the two reserved fills,
//     one border, one xf, the Normal style) and one empty
//     xl/worksheets/sheetN.xml per sheet.
//  2. Retained package, nothing changed (no `structureDirty`, no dirty
//     sheet): every part verbatim, in source order.
//  3. Otherwise the substitution path: xl/workbook.xml gets a spliced
//     `<sheets>` (an original keeps its source sheetId and rId; copies and
//     new sheets get the next free ones), renumbered / removed / duplicated
//     `definedName@localSheetId`, every workbookView's `activeTab` and no
//     `firstSheet`; the workbook's rels and [Content_Types].xml are spliced,
//     never re-serialised; xl/calcChain.xml is dropped with its relationship
//     and Override; docProps/app.xml's sheet titles are rebuilt or removed;
//     a deleted sheet's part and everything reachable only through it are
//     skipped; each sheet part gets the retained sheet-reference pass
//     (sheetNamePass.ts, with a copy's own remap), its `<tabColor>` spliced
//     when the model's colour differs from the source's, and the first
//     sheetView's `tabSelected` made to agree with the active sheet; a new
//     sheet is written empty at the next free `xl/worksheets/sheet<N>.xml`.
//     Every other part is verbatim — the pass touches only the parts that
//     mention a remapped name.
//
// Interim until SPEC-122 (which replaces this paragraph): a dirty sheet with
// a source is written exactly as a clean one — its retained part plus the
// pass and the splices. Its formulas are right (the pass rewrites `<f>`),
// but cached values the edit pipeline recalculated are not refreshed here,
// so whenever a dirty sheet is written this way the workbook gets
// `calcPr@fullCalcOnLoad="1"` and the consumer recomputes on open. New sheets
// are always empty here; SPEC-122 serialises cells.
//
// `fullCalcOnLoad="1"` is also written when the caller asks
// (`options.fullCalcOnLoad`), when a recalculation left a stale cell
// (`model.forceFullCalcOnLoad`, SPEC-120) and, from scratch, when the
// model's own `calc.fullCalcOnLoad` says so.
import { now as clockNow } from "../../core/clock";
import { buildCorePropertiesXml } from "../../core/opcProperties";
import { OFFICE_THEME_XML } from "../../core/officeThemeXml";
import { decodedPartBytes } from "../../core/packageParts";
import { SPREADSHEETML_MAIN_CONTENT_TYPES } from "../../core/packageDetect";
import { sameModelColor } from "../../core/theme";
import { relsPathFor } from "../../core/xml";
import { scanAttr, scanXml } from "../../core/xmlScanner";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import { isEmptyRemap, type SheetRemap } from "../formula/sheetRefs";
import { pxToXlsxWidth } from "../geometry";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "../geometry";
import type { Sheet, SpreadsheetModel } from "../model";
import { copyRemap, sourceSheetRemap } from "../sourceNames";
import { applySplices, attributeValue, escapeAttributeValue, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { ownedClosure } from "./duplicate";
import {
  contentTypeIn,
  decodeUtf8,
  encodeUtf8,
  nextFreePartName,
  nextRelationshipId,
  packageEdges,
  reachableFrom,
  readContentTypes,
  readSourceWorkbook,
  relationshipKind,
  relationshipsOf,
  relationshipXml,
  relativeTarget,
  spliceContentTypes,
  spliceRels,
  targetPart,
  type SourceSheet,
  type SourceWorkbook,
} from "./package";
import { mentionsAnyName, passFor, remapXlsxFormula, remappedNames } from "./sheetNamePass";
import { modelColorToXlsxAttrs, xlsxColorFromScan } from "./themeColor";

export interface WriteXlsxOptions {
  /** The `modified` stamp of a from-scratch docProps/core.xml (defaults to
   * core `now()`) — injectable so goldens are exact. */
  modified?: Date;
  /** Ask the consumer to recalculate on open. */
  fullCalcOnLoad?: boolean;
}

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
/** Excel's calculation engine id, as a from-scratch `calcPr` names it. */
const CALC_ID = "191029";

/** Serialises a workbook to .xlsx bytes (see the module header). */
export async function writeXlsx(model: SpreadsheetModel, options: WriteXlsxOptions = {}): Promise<Uint8Array> {
  if (model.opaque.length === 0) {
    return zip(buildFromScratch(model, options));
  }
  const parts = decodedPartBytes(model.opaque);
  if (!model.structureDirty && !model.sheets.some((sheet) => sheet.dirty)) {
    return zip([...parts].map(([name, data]) => ({ name, data })));
  }
  return zip(buildSubstitution(model, parts, options));
}

// --- from scratch ------------------------------------------------------------------------

const encode = (text: string): Uint8Array => encodeUtf8(text);

function sheetStateAttr(sheet: Sheet): string {
  return sheet.visibility === "visible" ? "" : ` state="${sheet.visibility}"`;
}

/** An empty worksheet part (SPEC-121 "From scratch"): the tab colour, the
 * selected flag on the active sheet, the default sizes. */
export function emptyWorksheetXml(sheet: Sheet, active: boolean, model: SpreadsheetModel): string {
  const tabColor = sheet.tabColor === undefined ? "" : `<sheetPr><tabColor${modelColorToXlsxAttrs(sheet.tabColor, model.theme)}/></sheetPr>`;
  const colWidth =
    sheet.defaultColWidthPx === DEFAULT_COL_WIDTH_PX ? "" : ` defaultColWidth="${pxToXlsxWidth(sheet.defaultColWidthPx, 7)}"`;
  const rowHeight = Number.isFinite(sheet.defaultRowHeightPt) ? sheet.defaultRowHeightPt : DEFAULT_ROW_HEIGHT_PT;
  return (
    `${XML_DECLARATION}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${tabColor}<dimension ref="A1"/>` +
    `<sheetViews><sheetView workbookViewId="0"${active ? ` tabSelected="1"` : ""}/></sheetViews>` +
    `<sheetFormatPr${colWidth} defaultRowHeight="${rowHeight}"/><sheetData/>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
  );
}

/** The minimal styles part: one font (Calibri 11 through the minor theme
 * font, colour theme 1), the two reserved fills, one empty border, one
 * cell-style xf, one cell xf and the Normal style. */
export const FROM_SCRATCH_STYLES_XML =
  `${XML_DECLARATION}<styleSheet xmlns="${MAIN_NS}">` +
  `<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="0"/><tableStyles count="0"/></styleSheet>`;

/** A from-scratch app.xml: the application only. */
export const FROM_SCRATCH_APP_XML =
  `${XML_DECLARATION}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
  `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>docu-edit</Application></Properties>`;

function buildFromScratch(model: SpreadsheetModel, options: WriteXlsxOptions): ZipEntryInput[] {
  const modified = options.modified ?? clockNow();
  const fullCalc = options.fullCalcOnLoad === true || model.forceFullCalcOnLoad === true || model.calc?.fullCalcOnLoad === true;
  const mainType = SPREADSHEETML_MAIN_CONTENT_TYPES[model.packageVariant ?? "workbook"];
  const sheetParts = model.sheets.map((_, index) => `xl/worksheets/sheet${index + 1}.xml`);
  const n = model.sheets.length;
  const contentTypes =
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="${mainType}"/>` +
    sheetParts.map((part) => `<Override PartName="/${part}" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`).join("") +
    `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
  const rootRels =
    `${XML_DECLARATION}<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPE}/officeDocument" Target="xl/workbook.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${RELATIONSHIP_TYPE}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const workbook =
    `${XML_DECLARATION}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
    (model.dateSystem === "1904" ? `<workbookPr date1904="1"/>` : `<workbookPr/>`) +
    `<bookViews><workbookView activeTab="${model.activeSheetIndex}"/></bookViews><sheets>` +
    model.sheets
      .map((sheet, index) => `<sheet name="${escapeAttributeValue(sheet.name)}" sheetId="${index + 1}"${sheetStateAttr(sheet)} r:id="rId${index + 1}"/>`)
      .join("") +
    `</sheets><calcPr calcId="${CALC_ID}"${fullCalc ? ` fullCalcOnLoad="1"` : ""}/></workbook>`;
  const workbookRels =
    `${XML_DECLARATION}<Relationships xmlns="${PKG_REL_NS}">` +
    sheetParts.map((part, index) => relationshipXml(`rId${index + 1}`, `${RELATIONSHIP_TYPE}/worksheet`, part.slice("xl/".length))).join("") +
    relationshipXml(`rId${n + 1}`, `${RELATIONSHIP_TYPE}/theme`, "theme/theme1.xml") +
    relationshipXml(`rId${n + 2}`, `${RELATIONSHIP_TYPE}/styles`, "styles.xml") +
    `</Relationships>`;
  return [
    { name: "[Content_Types].xml", data: encode(contentTypes) },
    { name: "_rels/.rels", data: encode(rootRels) },
    { name: "docProps/core.xml", data: encode(buildCorePropertiesXml({ title: model.title }, modified)) },
    { name: "docProps/app.xml", data: encode(FROM_SCRATCH_APP_XML) },
    { name: "xl/workbook.xml", data: encode(workbook) },
    { name: "xl/_rels/workbook.xml.rels", data: encode(workbookRels) },
    { name: "xl/theme/theme1.xml", data: encode(OFFICE_THEME_XML) },
    { name: "xl/styles.xml", data: encode(FROM_SCRATCH_STYLES_XML) },
    ...model.sheets.map((sheet, index) => ({
      name: sheetParts[index]!,
      data: encode(emptyWorksheetXml(sheet, index === model.activeSheetIndex, model)),
    })),
  ];
}

// --- substitution -------------------------------------------------------------------------

/** Where each model sheet goes in the package. */
interface SheetPlan {
  sheet: Sheet;
  index: number;
  partName: string;
  sheetId: number;
  rId: string;
  /** The source `<sheet>` an original keeps. */
  source?: SourceSheet;
  /** A sheet with no retained part: written empty. */
  generated: boolean;
}

function planSheets(model: SpreadsheetModel, parts: Map<string, Uint8Array>, workbook: SourceWorkbook): SheetPlan[] {
  const byPart = new Map(workbook.sheets.map((sheet) => [sheet.partName, sheet] as const));
  const claimed = new Set<string>();
  const taken = new Set(parts.keys());
  const relIds = new Set(workbook.rels.map((rel) => rel.id));
  let maxSheetId = Math.max(0, ...workbook.sheets.map((sheet) => sheet.sheetId));
  return model.sheets.map((sheet, index) => {
    const source = sheet.copiedFrom === undefined && sheet.partName !== undefined ? byPart.get(sheet.partName) : undefined;
    if (source && !claimed.has(source.partName)) {
      claimed.add(source.partName);
      return { sheet, index, partName: source.partName, sheetId: source.sheetId, rId: source.rId, source, generated: false };
    }
    const retained = sheet.partName !== undefined && parts.has(sheet.partName);
    const partName = retained ? sheet.partName! : nextFreePartName(taken, "xl/worksheets/sheet1.xml");
    taken.add(partName);
    maxSheetId += 1;
    const rId = nextRelationshipId(relIds);
    relIds.add(rId);
    return { sheet, index, partName, sheetId: maxSheetId, rId, generated: !retained };
  });
}

function buildSubstitution(model: SpreadsheetModel, parts: Map<string, Uint8Array>, options: WriteXlsxOptions): ZipEntryInput[] {
  const workbook = readSourceWorkbook(parts);
  if (!workbook) {
    throw new Error("The workbook's retained package has no workbook part");
  }
  const remap = sourceSheetRemap(model, parts);
  const plans = planSheets(model, parts, workbook);
  const kept = new Set(plans.filter((plan) => plan.source).map((plan) => plan.source!.partName));
  const deletedSources = workbook.sheets.filter((sheet) => !kept.has(sheet.partName));
  const calcChainRels = workbook.rels.filter((rel) => relationshipKind(rel.type) === "calcChain");

  // What the package reaches before and after: the workbook's edges are its
  // surviving relationships plus every planned sheet part.
  const edges = packageEdges(parts);
  const sourceReach = reachableFrom([""], edges);
  const droppedRelIds = new Set([...deletedSources.map((sheet) => sheet.rId), ...calcChainRels.map((rel) => rel.id)]);
  const afterEdges = new Map(edges);
  afterEdges.set(workbook.path, [
    ...workbook.rels.filter((rel) => !droppedRelIds.has(rel.id) && !rel.external).flatMap((rel) => targetPart(workbook.path, rel) ?? []),
    ...plans.map((plan) => plan.partName),
  ]);
  const afterReach = reachableFrom([""], afterEdges);
  const skipped = new Set<string>();
  for (const part of sourceReach) {
    if (!afterReach.has(part) && part !== "") {
      skipped.add(part);
      skipped.add(relsPathFor(part));
    }
  }

  // Each copy's own closure takes the copy's remap.
  const ownerRemap = new Map<string, SheetRemap>();
  for (const plan of plans) {
    if (plan.sheet.copiedFrom !== undefined && !plan.generated) {
      const own = copyRemap(remap, plan.sheet);
      for (const part of ownedClosure(edges, plan.partName)) {
        ownerRemap.set(part, own);
      }
    }
  }

  const anyDirty = model.sheets.some((sheet) => sheet.dirty && !isGenerated(plans, sheet));
  const fullCalc = options.fullCalcOnLoad === true || model.forceFullCalcOnLoad === true || anyDirty;
  const activeIndex = model.activeSheetIndex;
  const out: ZipEntryInput[] = [];
  const contentTypesName = "[Content_Types].xml";
  const appPath = appPropertiesPath(parts);
  const lastWorksheet = [...parts.keys()].filter((name) => name.startsWith("xl/worksheets/") && !skipped.has(name)).at(-1);
  const generated = plans.filter((plan) => plan.generated);
  const planByPart = new Map(plans.map((plan) => [plan.partName, plan] as const));

  for (const [name, bytes] of parts) {
    if (skipped.has(name)) {
      continue;
    }
    if (calcChainRels.some((rel) => targetPart(workbook.path, rel) === name)) {
      continue;
    }
    let data = bytes;
    if (name === contentTypesName) {
      data = encode(contentTypesFor(decodeUtf8(bytes), parts, skipped, plans, workbook, calcChainRels));
    } else if (name === workbook.path) {
      data = encode(workbookXmlFor(workbook, plans, remap, activeIndex, fullCalc));
    } else if (name === workbook.relsPath) {
      data = encode(workbookRelsFor(workbook, plans, droppedRelIds));
    } else if (name === appPath) {
      data = encode(appXmlFor(decodeUtf8(bytes), workbook, plans, remap));
    } else {
      const plan = planByPart.get(name);
      if (plan) {
        data = sheetPartFor(bytes, plan, plan.index === activeIndex, ownerRemap.get(name) ?? remap, model);
      } else {
        data = passed(name, bytes, ownerRemap.get(name) ?? remap);
      }
    }
    out.push({ name, data });
    if (name === lastWorksheet) {
      out.push(...generated.map((plan) => ({ name: plan.partName, data: encode(emptyWorksheetXml(plan.sheet, plan.index === activeIndex, model)) })));
    }
  }
  if (lastWorksheet === undefined) {
    out.push(...generated.map((plan) => ({ name: plan.partName, data: encode(emptyWorksheetXml(plan.sheet, plan.index === activeIndex, model)) })));
  }
  return out;
}

function isGenerated(plans: readonly SheetPlan[], sheet: Sheet): boolean {
  return plans.some((plan) => plan.sheet === sheet && plan.generated);
}

/** A retained part through the pass, when it is a pass location that
 * mentions a remapped name; else its bytes untouched. */
function passed(name: string, bytes: Uint8Array, remap: SheetRemap): Uint8Array {
  const pass = passFor(name);
  if (!pass || isEmptyRemap(remap) || !mentionsAnyName(bytes, remappedNames(remap))) {
    return bytes;
  }
  const xml = decodeUtf8(bytes);
  const next = pass(xml, remap);
  return next === xml ? bytes : encode(next);
}

// --- sheet parts --------------------------------------------------------------------------

/** A retained sheet part with the pass, its tab colour and its selected
 * flag spliced — the same bytes when none of them changes anything. Without
 * the pass only the head before `<sheetData` is decoded: the cells (all but
 * a few hundred bytes of a large part) go back out as they came. */
function sheetPartFor(bytes: Uint8Array, plan: SheetPlan, active: boolean, remap: SheetRemap, model: SpreadsheetModel): Uint8Array {
  const needsPass = !isEmptyRemap(remap) && mentionsAnyName(bytes, remappedNames(remap));
  const cut = needsPass ? -1 : sheetDataOffset(bytes);
  if (cut < 0) {
    const xml = decodeUtf8(bytes);
    let next = needsPass ? (passFor(plan.partName) ?? ((text: string) => text))(xml, remap) : xml;
    next = spliceTabSelected(spliceTabColor(next, plan.sheet, model), active);
    return next === xml ? bytes : encode(next);
  }
  const head = decodeUtf8(bytes.subarray(0, cut));
  const marker = "<sheetData/>";
  const probe = head + marker;
  const next = spliceTabSelected(spliceTabColor(probe, plan.sheet, model), active);
  if (next === probe) {
    return bytes;
  }
  const headBytes = encode(next.slice(0, next.length - marker.length));
  const out = new Uint8Array(headBytes.length + bytes.length - cut);
  out.set(headBytes, 0);
  out.set(bytes.subarray(cut), headBytes.length);
  return out;
}

/** The byte offset of a worksheet's `<sheetData` (any prefix), or -1. */
function sheetDataOffset(bytes: Uint8Array): number {
  const needle = encode("sheetData");
  outer: for (let i = 1; i + needle.length <= bytes.length; i += 1) {
    for (let j = 0; j < needle.length; j += 1) {
      if (bytes[i + j] !== needle[j]) {
        continue outer;
      }
    }
    let lt = i - 1;
    while (lt >= 0 && bytes[lt] !== 0x3c && bytes[lt] !== 0x3e && bytes[lt] !== 0x20) {
      lt -= 1;
    }
    if (lt >= 0 && bytes[lt] === 0x3c && bytes[lt + 1] !== 0x2f) {
      return lt;
    }
  }
  return -1;
}

/** The first element named `local` (its start tag, and whether it closes
 * itself), and the root's start tag. */
function locate(xml: string, locals: readonly string[]): Map<string, { start: number; end: number; selfClosing: boolean; closeEnd: number }> {
  const found = new Map<string, { start: number; end: number; selfClosing: boolean; closeEnd: number }>();
  const open: string[] = [];
  scanXml(xml, (event) => {
    if (event.type === "open") {
      if (locals.includes(event.local) && !found.has(event.local)) {
        found.set(event.local, { start: event.start, end: event.end, selfClosing: event.selfClosing, closeEnd: event.selfClosing ? event.end : -1 });
        if (!event.selfClosing) {
          open.push(event.local);
        }
      } else if (!event.selfClosing) {
        open.push("");
      }
    } else if (event.type === "close") {
      const local = open.pop();
      const entry = local ? found.get(local) : undefined;
      if (entry && entry.closeEnd < 0) {
        entry.closeEnd = event.end;
      }
    }
    return found.has("sheetData") ? "stop" : undefined;
  });
  return found;
}

/** The root element's start tag and the prefix its children are written
 * with (`""` for the default namespace). */
function rootTag(xml: string): { start: number; tag: string; prefix: string } {
  let start = 0;
  for (;;) {
    start = xml.indexOf("<", start);
    if (start < 0 || !/^<[?!]/.test(xml.slice(start, start + 2))) {
      break;
    }
    start = xml.indexOf(">", start) + 1;
  }
  const tag = startTagAt(xml, start);
  const name = /^<([\w.:-]+)/.exec(tag)?.[1] ?? "";
  return { start, tag, prefix: name.includes(":") ? `${name.slice(0, name.indexOf(":"))}:` : "" };
}

/** `<tabColor>` inserted, replaced or removed when the model's tab colour
 * differs from the part's. */
function spliceTabColor(xml: string, sheet: Sheet, model: SpreadsheetModel): string {
  const found = locate(xml, ["sheetPr", "tabColor", "sheetData"]);
  const tab = found.get("tabColor");
  const sheetPr = found.get("sheetPr");
  let sourceColor = undefined;
  if (tab && sheetPr && tab.start > sheetPr.start && tab.start < sheetPr.closeEnd) {
    scanXml(xml.slice(tab.start, tab.end).replace(/\/?>$/, "/>"), (event) => {
      if (event.type === "open") {
        sourceColor = xlsxColorFromScan(event.attrs, model.indexedColors);
      }
    });
  }
  const tabInSheetPr = tab !== undefined && sheetPr !== undefined && tab.start > sheetPr.start && tab.start < sheetPr.closeEnd;
  if (sheet.tabColor === undefined ? sourceColor === undefined : sourceColor !== undefined && sameModelColor(sourceColor, sheet.tabColor)) {
    return xml;
  }
  const { prefix } = rootTag(xml);
  if (sheet.tabColor === undefined) {
    return tabInSheetPr ? xml.slice(0, tab.start) + xml.slice(tab.closeEnd) : xml;
  }
  const element = `<${prefix}tabColor${modelColorToXlsxAttrs(sheet.tabColor, model.theme)}/>`;
  if (tabInSheetPr) {
    return xml.slice(0, tab.start) + element + xml.slice(tab.closeEnd);
  }
  if (sheetPr) {
    if (sheetPr.selfClosing) {
      const tag = xml.slice(sheetPr.start, sheetPr.end);
      return `${xml.slice(0, sheetPr.start)}${tag.replace(/\s*\/>$/, ">")}${element}</${prefix}sheetPr>${xml.slice(sheetPr.end)}`;
    }
    return xml.slice(0, sheetPr.end) + element + xml.slice(sheetPr.end);
  }
  const root = rootTag(xml);
  const at = root.start + root.tag.length;
  return `${xml.slice(0, at)}<${prefix}sheetPr>${element}</${prefix}sheetPr>${xml.slice(at)}`;
}

/** The first sheetView's `tabSelected` agreeing with whether the sheet is
 * the active one (two selected tabs put Excel into group editing). */
function spliceTabSelected(xml: string, active: boolean): string {
  const view = locate(xml, ["sheetView", "sheetData"]).get("sheetView");
  if (!view) {
    return xml;
  }
  const tag = xml.slice(view.start, view.end);
  const selected = attributeValue(tag, "tabSelected");
  const isSelected = selected === "1" || selected === "true";
  if (isSelected === active) {
    return xml;
  }
  const next = setAttribute(tag, "tabSelected", active ? "1" : null);
  return xml.slice(0, view.start) + next + xml.slice(view.end);
}

// --- the workbook part ---------------------------------------------------------------------

/** CT_Workbook's children after `calcPr`: an absent `calcPr` is created
 * before the first of these. */
const AFTER_CALC_PR = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];

function workbookXmlFor(workbook: SourceWorkbook, plans: readonly SheetPlan[], remap: SheetRemap, activeIndex: number, fullCalc: boolean): string {
  const xml = workbook.xml;
  const root = rootTag(xml);
  const relPrefix = relationshipPrefix(root.tag);
  const splices: Splice[] = [];
  const sourceIndexToModel = new Map<number, number>();
  workbook.sheets.forEach((source, sourceIndex) => {
    const plan = plans.find((each) => each.source === source);
    if (plan) {
      sourceIndexToModel.set(sourceIndex, plan.index);
    }
  });
  const copiesBySourceIndex = new Map<number, SheetPlan[]>();
  for (const plan of plans) {
    if (plan.sheet.copiedFrom !== undefined) {
      const sourceIndex = workbook.sheets.findIndex((source) => source.name.toUpperCase() === plan.sheet.copiedFrom!.toUpperCase());
      if (sourceIndex >= 0) {
        copiesBySourceIndex.set(sourceIndex, [...(copiesBySourceIndex.get(sourceIndex) ?? []), plan]);
      }
    }
  }

  let depth = 0;
  let sheetsRange: { start: number; end: number } | null = null;
  let sheetsStart = -1;
  let calcPr: { start: number; end: number } | null = null;
  let insertCalcPrAt = -1;
  let definedName: { start: number; tagEnd: number; attrs: Parameters<typeof scanAttr>[0] } | null = null;
  let rootClose = -1;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      depth += 1;
      if (depth === 2) {
        if (event.local === "sheets") {
          sheetsStart = event.start;
          if (event.selfClosing) {
            sheetsRange = { start: event.start, end: event.end };
          }
        } else if (event.local === "calcPr") {
          calcPr = { start: event.start, end: event.end };
        } else if (AFTER_CALC_PR.includes(event.local) && insertCalcPrAt < 0) {
          insertCalcPrAt = event.start;
        }
      } else if (depth === 3 && event.local === "workbookView") {
        const tag = xml.slice(event.start, event.end);
        const next = setAttribute(setAttribute(tag, "activeTab", String(activeIndex)), "firstSheet", null);
        if (next !== tag) {
          splices.push({ start: event.start, end: event.end, text: next });
        }
      } else if (depth === 3 && event.local === "definedName") {
        if (event.selfClosing) {
          definedNameSplices(xml, { start: event.start, end: event.end, textStart: event.end, textEnd: event.end }, event.attrs, sourceIndexToModel, copiesBySourceIndex, remap, splices);
        } else {
          definedName = { start: event.start, tagEnd: event.end, attrs: event.attrs };
        }
      }
      if (event.selfClosing) {
        depth -= 1;
      }
    } else if (event.type === "close") {
      if (depth === 2 && event.local === "sheets" && sheetsStart >= 0) {
        sheetsRange = { start: sheetsStart, end: event.end };
      } else if (depth === 3 && event.local === "definedName" && definedName) {
        const current: { start: number; tagEnd: number; attrs: Parameters<typeof scanAttr>[0] } = definedName;
        definedNameSplices(
          xml,
          { start: current.start, end: event.end, textStart: current.tagEnd, textEnd: event.start },
          current.attrs,
          sourceIndexToModel,
          copiesBySourceIndex,
          remap,
          splices,
        );
        definedName = null;
      } else if (depth === 1) {
        rootClose = event.start;
      }
      depth -= 1;
    }
  });

  if (sheetsRange) {
    const range: { start: number; end: number } = sheetsRange;
    const sourceSheetsTag = startTagAt(xml, range.start);
    const elementName = /^<([\w.:-]+)/.exec(sourceSheetsTag)?.[1] ?? "sheets";
    const sheetName = elementName.replace(/sheets$/, "sheet");
    const children = plans
      .map((plan) => {
        if (plan.source) {
          let tag = xml.slice(plan.source.start, plan.source.end);
          tag = setAttribute(tag, "name", plan.sheet.name);
          tag = setAttribute(tag, "state", plan.sheet.visibility === "visible" ? null : plan.sheet.visibility);
          return tag;
        }
        return `<${sheetName} name="${escapeAttributeValue(plan.sheet.name)}" sheetId="${plan.sheetId}"${sheetStateAttr(plan.sheet)} ${relPrefix}id="${plan.rId}"/>`;
      })
      .join("");
    const open = sourceSheetsTag.endsWith("/>") ? sourceSheetsTag.replace(/\s*\/>$/, ">") : sourceSheetsTag;
    splices.push({ start: range.start, end: range.end, text: `${open}${children}</${elementName}>` });
  }
  if (fullCalc) {
    if (calcPr) {
      const range: { start: number; end: number } = calcPr;
      const tag = xml.slice(range.start, range.end);
      const next = setAttribute(tag, "fullCalcOnLoad", "1");
      if (next !== tag) {
        splices.push({ start: range.start, end: range.end, text: next });
      }
    } else {
      const at = insertCalcPrAt >= 0 ? insertCalcPrAt : rootClose;
      splices.push({ start: at, end: at, text: `<${root.prefix}calcPr fullCalcOnLoad="1"/>` });
    }
  }
  return applySplices(xml, splices);
}

/** The prefix bound to the relationships namespace on the root, colon
 * included (`r:` by default). */
function relationshipPrefix(rootStartTag: string): string {
  for (const match of rootStartTag.matchAll(/\sxmlns:([\w.-]+)\s*=\s*"([^"]*)"/g)) {
    if (match[2] === REL_NS || match[2] === "http://purl.oclc.org/ooxml/officeDocument/relationships") {
      return `${match[1]}:`;
    }
  }
  return "r:";
}

/** One definedName: its text through the pass; a sheet-scoped name
 * renumbered to its sheet's new index, removed with a deleted sheet, and
 * duplicated (after itself) for each copy of its sheet. */
function definedNameSplices(
  xml: string,
  element: { start: number; end: number; textStart: number; textEnd: number },
  attrs: Parameters<typeof scanAttr>[0],
  sourceIndexToModel: ReadonlyMap<number, number>,
  copiesBySourceIndex: ReadonlyMap<number, readonly SheetPlan[]>,
  remap: SheetRemap,
  splices: Splice[],
): void {
  const local = scanAttr(attrs, "localSheetId");
  const sourceIndex = local === undefined ? undefined : Number(local);
  const raw = xml.slice(element.textStart, element.textEnd);
  const text = (formula: SheetRemap): string => {
    const wire = decodeEntities(raw);
    const next = remapXlsxFormula(wire, formula);
    return next === wire ? raw : escapeText(next);
  };
  const tagText = xml.slice(element.start, element.textStart);
  const closeText = xml.slice(element.textEnd, element.end);
  const withIndex = (tag: string, index: number | undefined): string =>
    index === undefined ? tag : setAttribute(tag, "localSheetId", String(index));
  if (sourceIndex !== undefined && Number.isInteger(sourceIndex)) {
    const modelIndex = sourceIndexToModel.get(sourceIndex);
    if (modelIndex === undefined) {
      splices.push({ start: element.start, end: element.end, text: copiesFor(sourceIndex) });
      return;
    }
    const tag = withIndex(tagText, modelIndex);
    const body = text(remap);
    const own = tag === tagText && body === raw ? xml.slice(element.start, element.end) : `${tag}${body}${closeText}`;
    const copies = copiesFor(sourceIndex);
    if (own !== xml.slice(element.start, element.end) || copies !== "") {
      splices.push({ start: element.start, end: element.end, text: own + copies });
    }
    return;
  }
  const body = text(remap);
  if (body !== raw) {
    splices.push({ start: element.textStart, end: element.textEnd, text: body });
  }

  function copiesFor(index: number): string {
    return (copiesBySourceIndex.get(index) ?? [])
      .map((plan) => `${withIndex(tagText, plan.index)}${text(copyRemap(remap, plan.sheet))}${closeText}`)
      .join("");
  }
}

function decodeEntities(raw: string): string {
  return raw.replace(/&(amp|lt|gt|quot|apos);/g, (_whole, name: string) => ({ amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" })[name] ?? "");
}

function escapeText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// --- relationships and content types ----------------------------------------------------------

function workbookRelsFor(workbook: SourceWorkbook, plans: readonly SheetPlan[], dropped: ReadonlySet<string>): string {
  const sourceIds = new Set(workbook.rels.map((rel) => rel.id));
  const worksheetType = workbook.rels.find((rel) => relationshipKind(rel.type) === "worksheet")?.type ?? `${RELATIONSHIP_TYPE}/worksheet`;
  const added = plans
    .filter((plan) => !plan.source && !sourceIds.has(plan.rId))
    .map((plan) => relationshipXml(plan.rId, worksheetType, relativeTarget(workbook.path, plan.partName)));
  return spliceRels(workbook.relsXml, dropped, added);
}

/** The content type a part added by duplication takes: that of a part in
 * the same folder with the same stem (its origin's). */
function stemContentType(types: ReturnType<typeof readContentTypes>, partName: string): string | undefined {
  const stem = (name: string): string => name.replace(/(?:_\d+|\d+)(\.[^./]*)$/, "$1");
  const key = stem(partName);
  return types.overrides.find((entry) => stem(entry.partName) === key)?.contentType;
}

function contentTypesFor(
  xml: string,
  parts: Map<string, Uint8Array>,
  skipped: ReadonlySet<string>,
  plans: readonly SheetPlan[],
  workbook: SourceWorkbook,
  calcChainRels: readonly { target: string; external: boolean; id: string; type: string; start: number; end: number }[],
): string {
  const types = readContentTypes(xml);
  const remove = new Set<string>();
  for (const entry of types.overrides) {
    if (skipped.has(entry.partName)) {
      remove.add(entry.partName);
    }
  }
  for (const rel of calcChainRels) {
    const target = targetPart(workbook.path, rel);
    if (target !== undefined) {
      remove.add(target);
    }
  }
  const worksheetType =
    types.overrides.find((entry) => workbook.sheets.some((sheet) => sheet.partName === entry.partName && relationshipKind(sheet.relType) === "worksheet"))
      ?.contentType ?? WORKSHEET_CONTENT_TYPE;
  const add: { partName: string; contentType: string }[] = [];
  for (const plan of plans) {
    if (plan.generated) {
      add.push({ partName: plan.partName, contentType: worksheetType });
    }
  }
  // A part a duplicate added takes its origin's Override (the extension's
  // Default — application/xml for every .xml — would misdeclare it).
  const overridden = new Set(types.overrides.map((entry) => entry.partName));
  for (const name of parts.keys()) {
    if (skipped.has(name) || overridden.has(name) || name === "[Content_Types].xml" || name.endsWith(".rels")) {
      continue;
    }
    const type = stemContentType(types, name);
    if (type !== undefined && type !== contentTypeIn(types, name)) {
      add.push({ partName: name, contentType: type });
    }
  }
  return remove.size === 0 && add.length === 0 ? xml : spliceContentTypes(xml, remove, add);
}

// --- docProps/app.xml ----------------------------------------------------------------------

function appPropertiesPath(parts: Map<string, Uint8Array>): string | undefined {
  const rel = relationshipsOf(parts, "").find((entry) => relationshipKind(entry.type) === "extended-properties");
  return rel ? targetPart("", rel) : undefined;
}

/** app.xml with HeadingPairs/TitlesOfParts following the model (SPEC-121):
 * the worksheet group whose titles equal the source's grid sheet names is
 * rewritten with the model's, likewise chart sheets; sheet-qualified named
 * ranges follow renames and leave with their sheet; when a group cannot be
 * identified both elements are removed. Everything else verbatim. */
function appXmlFor(xml: string, workbook: SourceWorkbook, plans: readonly SheetPlan[], remap: SheetRemap): string {
  const heading = /<([\w.]+:)?HeadingPairs\b[^>]*>([\s\S]*?)<\/\1HeadingPairs>/.exec(xml);
  const titles = /<([\w.]+:)?TitlesOfParts\b[^>]*>([\s\S]*?)<\/\1TitlesOfParts>/.exec(xml);
  if (!heading || !titles) {
    return xml;
  }
  const vt = /<([\w.]+):variant\b/.exec(heading[2]!)?.[1] ?? "vt";
  const groups = [...heading[2]!.matchAll(new RegExp(`<${vt}:variant>\\s*<${vt}:lpstr>([^<]*)</${vt}:lpstr>\\s*</${vt}:variant>\\s*<${vt}:variant>\\s*<${vt}:i4>(\\d+)</${vt}:i4>\\s*</${vt}:variant>`, "g"))].map(
    (match) => ({ label: match[1]!, count: Number(match[2]) }),
  );
  const entries = [...titles[2]!.matchAll(new RegExp(`<${vt}:lpstr>([^<]*)</${vt}:lpstr>`, "g"))].map((match) => decodeEntities(match[1]!));
  const remove = (): string => applySplices(xml, [heading, titles].map((match) => ({ start: match.index, end: match.index + match[0].length, text: "" })));
  if (groups.reduce((sum, group) => sum + group.count, 0) !== entries.length) {
    return remove();
  }
  const sourceGrid = workbook.sheets.filter((sheet) => relationshipKind(sheet.relType) === "worksheet").map((sheet) => sheet.name);
  const sourceCharts = workbook.sheets.filter((sheet) => relationshipKind(sheet.relType) === "chartsheet").map((sheet) => sheet.name);
  const modelGrid = plans.filter((plan) => plan.sheet.kind === "grid").map((plan) => plan.sheet.name);
  const modelCharts = plans
    .filter((plan) => plan.sheet.kind === "foreign" && plan.source !== undefined && relationshipKind(plan.source.relType) === "chartsheet")
    .map((plan) => plan.sheet.name);
  const renamed = new Map(Object.entries(remap.renamed).map(([from, to]) => [from.toUpperCase(), to] as const));
  const deleted = new Set(remap.deleted.map((name) => name.toUpperCase()));
  let at = 0;
  let sawGrid = sourceGrid.length === 0;
  let sawCharts = sourceCharts.length === 0;
  const rebuilt: { label: string; items: string[] }[] = [];
  for (const group of groups) {
    const items = entries.slice(at, at + group.count);
    at += group.count;
    if (!sawGrid && sameNames(items, sourceGrid)) {
      sawGrid = true;
      rebuilt.push({ label: group.label, items: modelGrid });
    } else if (!sawCharts && sameNames(items, sourceCharts)) {
      sawCharts = true;
      rebuilt.push({ label: group.label, items: modelCharts });
    } else {
      rebuilt.push({ label: group.label, items: items.flatMap((item) => qualifiedTitle(item, renamed, deleted)) });
    }
  }
  if (!sawGrid || !sawCharts) {
    return remove();
  }
  const kept = rebuilt.filter((group) => group.items.length > 0);
  const escape = (text: string): string => escapeText(text);
  const headingBody =
    `<${vt}:vector size="${kept.length * 2}" baseType="variant">` +
    kept.map((group) => `<${vt}:variant><${vt}:lpstr>${escape(group.label)}</${vt}:lpstr></${vt}:variant><${vt}:variant><${vt}:i4>${group.items.length}</${vt}:i4></${vt}:variant>`).join("") +
    `</${vt}:vector>`;
  const allTitles = kept.flatMap((group) => group.items);
  const titlesBody = `<${vt}:vector size="${allTitles.length}" baseType="lpstr">${allTitles.map((title) => `<${vt}:lpstr>${escape(title)}</${vt}:lpstr>`).join("")}</${vt}:vector>`;
  const open = (match: RegExpExecArray): string => match[0].slice(0, match[0].indexOf(">") + 1);
  const close = (match: RegExpExecArray): string => match[0].slice(match[0].lastIndexOf("</"));
  return applySplices(xml, [
    { start: heading.index, end: heading.index + heading[0].length, text: `${open(heading)}${headingBody}${close(heading)}` },
    { start: titles.index, end: titles.index + titles[0].length, text: `${open(titles)}${titlesBody}${close(titles)}` },
  ]);
}

function sameNames(a: readonly string[], b: readonly string[]): boolean {
  return a.length === b.length && a.every((name, index) => name === b[index]);
}

/** A named-range title (`Data!Print_Area`) following its sheet: renamed,
 * or gone with a deleted sheet. */
function qualifiedTitle(item: string, renamed: ReadonlyMap<string, string>, deleted: ReadonlySet<string>): string[] {
  const bang = item.lastIndexOf("!");
  if (bang <= 0) {
    return [item];
  }
  const rawSheet = item.slice(0, bang);
  const sheet = rawSheet.startsWith("'") && rawSheet.endsWith("'") ? rawSheet.slice(1, -1).replaceAll("''", "'") : rawSheet;
  if (deleted.has(sheet.toUpperCase())) {
    return [];
  }
  const next = renamed.get(sheet.toUpperCase());
  if (next === undefined) {
    return [item];
  }
  const quoted = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(next) ? next : `'${next.replaceAll("'", "''")}'`;
  return [`${quoted}${item.slice(bang)}`];
}
