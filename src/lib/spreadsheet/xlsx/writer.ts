// The .xlsx writer (SPEC-121 "From scratch", "The .xlsx substitution path";
// SPEC-122 "The worksheet writer", "Shared strings", "Styles", "calcChain and
// calcPr"). THIS FILE, with SPEC-121 and SPEC-122, is the normative source
// for how a workbook's package is rebuilt. Three paths, chosen by what the
// model holds (SPEC-080's for decks):
//
//  1. No retained package (`opaque` empty — a new workbook, or a
//     stripRetention'ed copy): the minimal package from the model, parts in
//     this order: [Content_Types].xml, _rels/.rels, docProps/core.xml,
//     docProps/app.xml (Application only — no TitlesOfParts to go stale),
//     xl/workbook.xml (SPEC-128: with a `<definedNames>` after `<sheets>`
//     holding one hidden `_xlnm._FilterDatabase` per sheet with an
//     autofilter, when one has), its rels, xl/theme/theme1.xml (core's
//     OFFICE_THEME_XML), xl/styles.xml (xlsx/styles.ts's from-scratch base —
//     font 0 and xf 0 built from `model.styles[0]` — with every other style a
//     cell, row or column uses appended), xl/sharedStrings.xml when a cell
//     holds a string, and one xl/worksheets/sheetN.xml per sheet, written by
//     sheetXml.ts's serializeWorksheet (SPEC-121's empty-sheet template, byte
//     for byte, for an empty sheet).
//  2. Retained package, nothing changed (no `structureDirty`, no dirty
//     sheet): every part verbatim, in source order.
//  3. Otherwise the substitution path: xl/workbook.xml gets a spliced
//     `<sheets>` (an original keeps its source sheetId and rId; copies and
//     new sheets get the next free ones), renumbered / removed / duplicated
//     `definedName@localSheetId`, every workbookView's `activeTab` and no
//     `firstSheet`, and (SPEC-128, a save only — not substitutedXlsxParts)
//     the `_xlnm._FilterDatabase` names reconciled with the model by string
//     splice: for every grid sheet once `structureDirty`, else for the dirty
//     and new ones that have an autofilter or whose source part had one (an
//     Advanced Filter's name is an ordinary one), a sheet with an autofilter keeps its first source name
//     (localSheetId renumbered; the text as written while it names the
//     autofilter's range, else `'Sheet'!$A$1:$D$20` — xlsx/autoFilter.ts) or
//     gains one (first in `<definedNames>`, which is created before `calcPr`
//     and its later siblings when absent), and its other names — like a
//     deleted sheet's — go; a copy gets its own from its autofilter, never a
//     duplicate of its origin's; a clean sheet's stay verbatim; the workbook's rels and [Content_Types].xml are spliced,
//     never re-serialised; xl/calcChain.xml is dropped with its relationship
//     and Override; docProps/app.xml's sheet titles are rebuilt or removed;
//     a deleted sheet's part and everything reachable only through it are
//     skipped. A clean sheet part gets the retained sheet-reference pass
//     (sheetNamePass.ts, with a copy's own remap), its `<tabColor>` spliced
//     when the model's colour differs from the source's, and the first
//     sheetView's `tabSelected` made to agree with the active sheet — the
//     same bytes when none of them changes anything. A dirty sheet's
//     retained text gets the pass and the tab-colour splice FIRST; then
//     serializeWorksheet splices the model's rebuilt elements into it (the
//     model's formulas already name sheets as the model does, so they are
//     never passed twice). A new sheet is serialised from the template at the
//     next free `xl/worksheets/sheet<N>.xml`. One SharedStringsWriter and
//     one XlsxStyleWriter serve the whole save: both parts are only ever
//     appended to — clean parts reference their items by number — and are
//     rewritten only when something was appended; a workbook without one
//     gains it with its relationship and Override. Every other part is
//     verbatim — the pass touches only the parts that mention a remapped
//     name.
//
// SPEC-131 (xlsx/drawingXml.ts is the normative source for the wire forms):
// `drawingsDirty` joins the rebuild triggers. On the substitution path (a
// save, not substitutedXlsxParts) each `drawingsDirty` grid sheet's drawing
// part and its `.rels` are rebuilt from the model, a sheet gaining its first
// drawing gets `xl/drawings/drawing{n}.xml` (next free), its `.rels`, the
// worksheet's `<drawing r:id>` (in CT_Worksheet order, clean sheet or dirty),
// the sheet relationship and the drawing Override, and the last drawing gone
// takes the `<drawing>` and the relationship with it — the part, its `.rels`
// and its Override then fall to the reachability rule, as does a media or
// chart part no relationship targets any more (one still targeted from
// elsewhere — a header image, another sheet's drawing — survives). Media
// ("Media parts"): `model.media` is the source of truth for its paths — a
// part there is written from it ahead of an `opaque` twin, a new one at its
// `xl/media/` path (any other mediaId at the next free
// `xl/media/image{n}.{ext}`), with a `<Default>` for an extension
// `[Content_Types].xml` lacks. From scratch, every grid sheet's image and
// shape drawings are written in the canonical form with the same parts.
//
// SPEC-135 (xlsx/annotations.ts is the normative source for the wire forms):
// `annotationsDirty` joins the rebuild triggers. On the substitution path (a
// save) every `dirty` or `annotationsDirty` grid sheet's notes, threads and
// hyperlinks are written by writeXlsxSheetAnnotations over the sheet `.rels`
// the drawing and table writes left: the comments, VML and threaded parts
// rebuilt, spliced, minted or dropped (their relationships with them — the
// reachability rule then drops a part and its Override), the minted ones'
// Overrides and the `vml` Default added, and `<hyperlinks>`/`<legacyDrawing>`
// spliced into the worksheet part after the drawing and table splices — a
// sheet that is only `annotationsDirty` changes nothing else; the persons
// part gains the persons it lacks (or is minted with its workbook
// relationship). From scratch, every grid sheet's annotations are written
// with the same parts.
//
// `calcPr@fullCalcOnLoad="1"` is written (SPEC-122 "calcChain and calcPr")
// when the caller asks (`options.fullCalcOnLoad`) or a recalculation left a
// stale cell (`model.forceFullCalcOnLoad`, SPEC-120 — set, never cleared),
// and from scratch also when the model's own `calc.fullCalcOnLoad` says so;
// a source `fullCalcOnLoad="1"` stays. Dirty sheets are written from the
// model with the values the edit pipeline recalculated (a stale cell's last
// value as it is), so an edit alone never asks the consumer to recalculate.
import { now as clockNow } from "../../core/clock";
import { equalBytes } from "../../core/equality";
import { ensureMediaEncoded, extensionForContentType, pendingMediaByteLength, type MediaPart } from "../../core/media";
import { buildCorePropertiesXml } from "../../core/opcProperties";
import { OFFICE_THEME_XML } from "../../core/officeThemeXml";
import { base64ToBytes, decodedPartBytes } from "../../core/packageParts";
import { SPREADSHEETML_MAIN_CONTENT_TYPES } from "../../core/packageDetect";
import { DEFAULT_THEME, resolveFont, sameModelColor } from "../../core/theme";
import { relsPathFor } from "../../core/xml";
import { scanXml } from "../../core/xmlScanner";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import { isEmptyRemap, type SheetRemap } from "../formula/sheetRefs";
import { buildGridMetrics, maxDigitWidthPx } from "../geometry";
import { isModelledChart } from "../charts";
import type { DefinedName, Sheet, SheetDrawing, SpreadsheetModel } from "../model";
import { RESERVED_PREFIX, setReservedName } from "../names";
import { copyRemap, sourceSheetRemap } from "../sourceNames";
import { applySplices, attributeValue, escapeAttributeValue, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { FILTER_DATABASE_NAME, filterDatabaseNames, filterDatabaseText, sheetPartHasAutoFilter } from "./autoFilter";
import { sheetChartsChanged, sheetHasCreatedChart, XlsxChartWriter } from "./chartParts";
import {
  annotationsRebuilt,
  maxVmlIdmap,
  PERSON_CONTENT_TYPE,
  PERSON_RELATIONSHIP_TYPE,
  VML_CONTENT_TYPE,
  writeXlsxPersons,
  writeXlsxSheetAnnotations,
  type AnnotationWriteContext,
  type PersonsWrite,
} from "./annotations";
import { DRAWING_CONTENT_TYPE, DRAWING_RELATIONSHIP_TYPE, newRelsXml, spliceWorksheetDrawing, writeDrawingPart, writeSheetDrawings } from "./drawingXml";
import { definedNamesSplices, definedNamesXml, scanDefinedNames } from "./definedNames";
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
import { SharedStringsWriter } from "./sharedStrings";
import { mentionsAnyName, passFor, remappedNames } from "./sheetNamePass";
import { emptyWorksheetXml, serializeWorksheet, type XlsxSheetContext } from "./sheetXml";
import { DEFAULT_XLSX_FONT, XlsxStyleWriter } from "./styles";
import { planTableWrites, spliceWorksheetTableParts, TABLE_CONTENT_TYPE, type TableWritePlan } from "./tableXml";
import { modelColorToXlsxAttrs, xlsxColorFromScan } from "./themeColor";
import { spliceXlsxPageSetup } from "./pageSetup";

export { emptyWorksheetXml } from "./sheetXml";

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
  if (
    !model.structureDirty &&
    !model.namesDirty &&
    !model.sheets.some((sheet) => sheet.dirty || sheet.drawingsDirty || sheet.annotationsDirty || sheet.pageSetupDirty || sheetChartsChanged(sheet) || sheetTablesChanged(sheet))
  ) {
    return zip([...parts].map(([name, data]) => ({ name, data })));
  }
  return zip(buildSubstitution(model, parts, options, "write"));
}

/** SPEC-133: a sheet holding a table to regenerate or to write as a new
 * part (a deleted table marks its sheet dirty). */
function sheetTablesChanged(sheet: Sheet): boolean {
  return (sheet.tables ?? []).some((table) => table.dirty === true || table.partName === undefined);
}

/** The substitution path's parts for any model with a retained package —
 * what a reader upgrade re-reads an edited workbook's clean sheets from
 * (open.ts): every model sheet laid out in order with the retained
 * sheet-reference pass applied, a dirty sheet's retained part included (the
 * upgrade keeps the stored dirty sheets, so nothing is serialised from the
 * model), the shared strings, styles and calculation flag left as the
 * source has them. */
export function substitutedXlsxParts(model: SpreadsheetModel): Map<string, Uint8Array> {
  const entries = buildSubstitution({ ...model, structureDirty: true }, decodedPartBytes(model.opaque), {}, "layout");
  return new Map(entries.map((entry) => [entry.name, entry.data] as const));
}

/** The maximum digit width column widths convert with: the default style's
 * font, resolved as the reader resolves the Normal font (Calibri 11 when it
 * names no face or size). */
function workbookMdw(model: SpreadsheetModel): number {
  const font = model.styles[0]?.font;
  const family = font?.family !== undefined ? resolveFont(font.family, model.theme ?? DEFAULT_THEME) : DEFAULT_XLSX_FONT.family;
  return maxDigitWidthPx(family, font?.sizePt ?? DEFAULT_XLSX_FONT.sizePt);
}

const SHARED_STRINGS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

// --- from scratch ------------------------------------------------------------------------

const encode = (text: string): Uint8Array => encodeUtf8(text);

function sheetStateAttr(sheet: Sheet): string {
  return sheet.visibility === "visible" ? "" : ` state="${sheet.visibility}"`;
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
  // SPEC-132: a chart sheet whose chart is modelled is a chartsheet part.
  let chartSheets = 0;
  const sheetParts = model.sheets.map((sheet, index) => (writesChartSheet(sheet) ? `xl/chartsheets/sheet${(chartSheets += 1)}.xml` : `xl/worksheets/sheet${index + 1}.xml`));
  const n = model.sheets.length;
  // The sheets first: they fill the string and style tables.
  const strings = new SharedStringsWriter(undefined, model.theme);
  const styles = new XlsxStyleWriter(undefined, model);
  const mdw = workbookMdw(model);
  const worksheets = model.sheets.map((sheet, index) =>
    isChartSheetPart(sheetParts[index]!)
      ? chartsheetXml(sheet, index === model.activeSheetIndex, model)
      : spliceXlsxPageSetup(
          serializeWorksheet(sheet, undefined, { active: index === model.activeSheetIndex, theme: model.theme, mdw, strings, styles, dxfs: model.dxfs ?? [] }),
          sheet.pageSetup,
        ),
  );
  const sharedStrings = strings.xml();
  const layer = fromScratchDrawings(model, worksheets, sheetParts);
  const tables = fromScratchTables(model, sheetParts, layer);
  const notes = fromScratchAnnotations(model, sheetParts, layer, tables);
  const contentTypes =
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    (notes.vml ? `<Default Extension="vml" ContentType="${VML_CONTENT_TYPE}"/>` : "") +
    [...layer.defaults].map(([extension, type]) => `<Default Extension="${escapeAttributeValue(extension)}" ContentType="${escapeAttributeValue(type)}"/>`).join("") +
    `<Override PartName="/xl/workbook.xml" ContentType="${mainType}"/>` +
    sheetParts.map((part) => `<Override PartName="/${part}" ContentType="${isChartSheetPart(part) ? CHARTSHEET_CONTENT_TYPE : WORKSHEET_CONTENT_TYPE}"/>`).join("") +
    layer.drawingParts.map((part) => `<Override PartName="/${part}" ContentType="${DRAWING_CONTENT_TYPE}"/>`).join("") +
    layer.overrides.map((entry) => `<Override PartName="/${entry.partName}" ContentType="${entry.contentType}"/>`).join("") +
    tables.overrides.map((part) => `<Override PartName="/${part}" ContentType="${TABLE_CONTENT_TYPE}"/>`).join("") +
    notes.overrides.map((entry) => `<Override PartName="/${entry.partName}" ContentType="${entry.contentType}"/>`).join("") +
    `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/>` +
    (sharedStrings === undefined ? "" : `<Override PartName="/xl/sharedStrings.xml" ContentType="${SHARED_STRINGS_CONTENT_TYPE}"/>`) +
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
    `</sheets>${fromScratchDefinedNames(model)}<calcPr calcId="${CALC_ID}"${fullCalc ? ` fullCalcOnLoad="1"` : ""}/></workbook>`;
  const workbookRels =
    `${XML_DECLARATION}<Relationships xmlns="${PKG_REL_NS}">` +
    sheetParts.map((part, index) => relationshipXml(`rId${index + 1}`, `${RELATIONSHIP_TYPE}/${isChartSheetPart(part) ? "chartsheet" : "worksheet"}`, part.slice("xl/".length))).join("") +
    relationshipXml(`rId${n + 1}`, `${RELATIONSHIP_TYPE}/theme`, "theme/theme1.xml") +
    relationshipXml(`rId${n + 2}`, `${RELATIONSHIP_TYPE}/styles`, "styles.xml") +
    (sharedStrings === undefined ? "" : relationshipXml(`rId${n + 3}`, `${RELATIONSHIP_TYPE}/sharedStrings`, "sharedStrings.xml")) +
    (notes.persons === undefined ? "" : relationshipXml(`rId${n + 4}`, PERSON_RELATIONSHIP_TYPE, relativeTarget("xl/workbook.xml", notes.persons.partName))) +
    `</Relationships>`;
  return [
    { name: "[Content_Types].xml", data: encode(contentTypes) },
    { name: "_rels/.rels", data: encode(rootRels) },
    { name: "docProps/core.xml", data: encode(buildCorePropertiesXml({ title: model.title }, modified)) },
    { name: "docProps/app.xml", data: encode(FROM_SCRATCH_APP_XML) },
    { name: "xl/workbook.xml", data: encode(workbook) },
    { name: "xl/_rels/workbook.xml.rels", data: encode(workbookRels) },
    { name: "xl/theme/theme1.xml", data: encode(OFFICE_THEME_XML) },
    { name: "xl/styles.xml", data: encode(styles.xml()) },
    ...(sharedStrings === undefined ? [] : [{ name: "xl/sharedStrings.xml", data: encode(sharedStrings) }]),
    ...layer.worksheets.map((xml, index) => ({ name: sheetParts[index]!, data: encode(xml) })),
    ...layer.entries,
    ...tables.entries,
    ...notes.entries,
  ];
}

/** SPEC-135 from scratch: every grid sheet's notes, threads and hyperlinks
 * (xlsx/annotations.ts writeXlsxSheetAnnotations over an empty package) —
 * the minted parts with their Overrides, the sheet `.rels` (joining a
 * drawing's or table's when the sheet has one), the worksheet splices, and
 * the persons part the threads name. */
function fromScratchAnnotations(
  model: SpreadsheetModel,
  sheetParts: readonly string[],
  layer: { worksheets: string[]; entries: ZipEntryInput[] },
  tables: { entries: ZipEntryInput[] },
): { entries: ZipEntryInput[]; overrides: { partName: string; contentType: string }[]; vml: boolean; persons?: PersonsWrite } {
  const out: { entries: ZipEntryInput[]; overrides: { partName: string; contentType: string }[]; vml: boolean; persons?: PersonsWrite } = { entries: [], overrides: [], vml: false };
  const empty = new Map<string, Uint8Array>();
  const taken = new Set([...sheetParts, ...layer.entries.map((entry) => entry.name), ...tables.entries.map((entry) => entry.name)]);
  let idmap = 0;
  const ctx: AnnotationWriteContext = { parts: empty, taken, nextIdmap: () => (idmap += 1), persons: model.persons };
  model.sheets.forEach((sheet, index) => {
    const sheetPart = sheetParts[index]!;
    if (sheet.kind !== "grid" || isChartSheetPart(sheetPart)) {
      return;
    }
    const relsPath = relsPathFor(sheetPart);
    const holder = [layer.entries, tables.entries].find((entries) => entries.some((entry) => entry.name === relsPath));
    const relsEntry = holder?.find((entry) => entry.name === relsPath);
    const write = writeXlsxSheetAnnotations(sheet, sheetPart, relsEntry === undefined ? undefined : decodeUtf8(relsEntry.data), undefined, ctx);
    if (write === undefined) {
      return;
    }
    layer.worksheets[index] = write.worksheet(layer.worksheets[index]!);
    for (const [name, text] of write.texts) {
      if (holder !== undefined && name === relsPath) {
        holder[holder.findIndex((entry) => entry.name === relsPath)] = { name, data: encode(text) };
      } else {
        out.entries.push({ name, data: encode(text) });
      }
    }
    out.overrides.push(...write.added.flatMap((part) => (part.contentType === undefined ? [] : [{ partName: part.partName, contentType: part.contentType }])));
    out.vml ||= write.vml;
  });
  const persons = writeXlsxPersons(model, empty, taken);
  if (persons !== undefined) {
    out.persons = persons;
    out.entries.push({ name: persons.partName, data: encode(persons.xml) });
    out.overrides.push({ partName: persons.partName, contentType: PERSON_CONTENT_TYPE });
  }
  return out;
}

/** SPEC-132: a chart sheet written from scratch as a chartsheet — one whose
 * chart is modelled (an out-of-catalogue chart has no part to write; such a
 * sheet is written as an empty worksheet). */
function writesChartSheet(sheet: Sheet): boolean {
  return sheet.kind === "chart" && chartSheetDrawing(sheet) !== undefined;
}

/** Whether a from-scratch sheet part is a chartsheet. */
function isChartSheetPart(part: string): boolean {
  return part.startsWith("xl/chartsheets/");
}

/** A chart sheet's one chart drawing, when it is modelled. */
function chartSheetDrawing(sheet: Sheet): SheetDrawing | undefined {
  return (sheet.drawings ?? []).find((drawing) => drawing.content.kind === "chart" && isModelledChart(drawing.content));
}

const CHARTSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

/** A from-scratch chartsheet (ECMA-376 §18.3.1.12 CT_Chartsheet, as Excel
 * writes one): its tab colour, one sheet view (selected when active, zoomed
 * to fit), Excel's default margins and the `<drawing r:id="rId1"/>` its
 * `.rels` resolves to the drawing holding the chart. */
function chartsheetXml(sheet: Sheet, active: boolean, model: SpreadsheetModel): string {
  const sheetPr = sheet.tabColor === undefined ? "" : `<sheetPr><tabColor${modelColorToXlsxAttrs(sheet.tabColor, model.theme)}/></sheetPr>`;
  return (
    `${XML_DECLARATION}<chartsheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">${sheetPr}` +
    `<sheetViews><sheetView${active ? ` tabSelected="1"` : ""} zoomToFit="1" workbookViewId="0"/></sheetViews>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></chartsheet>`
  );
}

/** SPEC-131 from scratch: each grid sheet's image and shape drawings in a
 * canonical `xl/drawings/drawing{n}.xml` (n counting the sheets that have
 * one) with its `.rels`, the sheet's `.rels` (rId1, the from-scratch sheet
 * having no other) and `<drawing r:id="rId1"/>`; the media parts the
 * pictures embed and a `<Default>` per extension. SPEC-132: every modelled
 * chart drawing too, each relating a new `xl/charts/chart{n}.xml`
 * (xlsx/chartParts.ts) with its Override; a chart sheet's chart in its
 * chartsheet's own drawing part. */
function fromScratchDrawings(
  model: SpreadsheetModel,
  worksheets: readonly string[],
  sheetParts: readonly string[],
): { worksheets: string[]; entries: ZipEntryInput[]; drawingParts: string[]; overrides: { partName: string; contentType: string }[]; defaults: Map<string, string> } {
  const media = mediaWrites(model, new Set(sheetParts));
  const charts = new XlsxChartWriter(model, new Map(), new Set(sheetParts));
  const entries: ZipEntryInput[] = [];
  const mediaEntries = new Map<string, Uint8Array>();
  const drawingParts: string[] = [];
  const defaults = new Map<string, string>();
  const out = worksheets.map((xml, index) => {
    const sheet = model.sheets[index]!;
    const sheetPart = sheetParts[index]!;
    const chartSheet = sheetPart.startsWith("xl/chartsheets/");
    if ((sheet.kind !== "grid" && !chartSheet) || (sheet.drawings ?? []).length === 0) {
      return xml;
    }
    const partName = `xl/drawings/drawing${drawingParts.length + 1}.xml`;
    const drawings = chartSheet ? [chartSheetDrawing(sheet)!] : sheet.drawings!;
    const chartPart = (drawing: SheetDrawing): string | undefined => charts.partFor(sheet, drawing);
    const written = writeDrawingPart({ partName, drawings, metrics: buildGridMetrics(sheet), mediaPath: media.pathOf, chartPart });
    if (written.written === 0) {
      return xml;
    }
    drawingParts.push(partName);
    entries.push({ name: relsPathFor(sheetPart), data: encode(newRelsXml([relationshipXml("rId1", DRAWING_RELATIONSHIP_TYPE, relativeTarget(sheetPart, partName))])) });
    entries.push({ name: partName, data: encode(written.xml) });
    if (written.rels !== undefined) {
      entries.push({ name: relsPathFor(partName), data: encode(written.rels) });
    }
    for (const target of written.targets) {
      const part = media.partAt(target);
      if (part && !mediaEntries.has(target)) {
        mediaEntries.set(target, mediaBytes(part, undefined));
        defaults.set(target.slice(target.lastIndexOf(".") + 1).toLowerCase(), part.contentType);
      }
    }
    return chartSheet ? xml : spliceWorksheetDrawing(xml, "rId1");
  });
  entries.push(...[...mediaEntries].map(([name, data]) => ({ name, data })));
  const chartWrites = charts.writes();
  entries.push(...chartWrites.added);
  return { worksheets: out, entries, drawingParts, overrides: chartWrites.overrides, defaults };
}

/** A from-scratch workbook's `<definedNames>` (SPEC-133): every model name
 * in model order, each sheet's `_xlnm._FilterDatabase` first reconciled with
 * its autofilter (SPEC-128: one hidden name per filtered sheet); a name kept
 * unparsed in another format's dialect is not written. "" when none. */
function fromScratchDefinedNames(model: SpreadsheetModel): string {
  const scratch: SpreadsheetModel = { ...model, names: model.names ?? [] };
  for (const sheet of model.sheets) {
    if (sheet.kind === "grid") {
      reconcileFilterDatabase(scratch, sheet);
    }
  }
  const names = model.format === "xlsx" ? scratch.names : scratch.names.filter((name) => !name.formulaUnparsed);
  const index = new Map(model.sheets.map((sheet, at) => [sheet.id, at] as const));
  return definedNamesXml(names, (sheetId) => index.get(sheetId));
}

const FILTER_DATABASE_SHORT = FILTER_DATABASE_NAME.slice(RESERVED_PREFIX.length);

/** SPEC-128 through SPEC-133's names.ts, on a COPY of the caller's model
 * (`scratch.names` replaced, never mutated): a sheet with an autofilter
 * keeps its first `_xlnm._FilterDatabase` while it names the autofilter's
 * range (however written), else gets the range written as Excel does
 * (`'Sheet'!$A$1:$D$20`, xlsx/autoFilter.ts) — a new one first, hidden;
 * a sheet without one loses it; a sheet's further ones go. */
function reconcileFilterDatabase(scratch: SpreadsheetModel, sheet: Sheet): void {
  const key = FILTER_DATABASE_NAME.toUpperCase();
  const own = scratch.names.filter((name) => name.scopeSheetId === sheet.id && name.name.toUpperCase() === key);
  if (own.length > 1) {
    const extra = new Set(own.slice(1));
    scratch.names = scratch.names.filter((name) => !extra.has(name));
  }
  const filter = sheet.autoFilter;
  if (filter === undefined) {
    setReservedName(scratch, sheet.id, FILTER_DATABASE_SHORT, null);
    return;
  }
  const existing = own[0];
  if (existing === undefined || existing.formulaUnparsed || !filterDatabaseNames(existing.formula, sheet.name, filter.range)) {
    setReservedName(scratch, sheet.id, FILTER_DATABASE_SHORT, filterDatabaseText(sheet.name, filter.range));
  }
}

/** The names a save writes: the model's, with the `_xlnm._FilterDatabase`
 * of every sheet `reconciled` accepts brought in line (the same array when
 * nothing changed). */
function namesToWrite(model: SpreadsheetModel, plans: readonly SheetPlan[], reconciled: (plan: SheetPlan) => boolean): DefinedName[] {
  const scratch: SpreadsheetModel = { ...model, names: model.names };
  for (const plan of plans) {
    if (reconciled(plan)) {
      reconcileFilterDatabase(scratch, plan.sheet);
    }
  }
  return scratch.names;
}

/** SPEC-133 "Tables on the wire", from scratch: every grid sheet's tables as
 * new parts (xlsx/tableXml.ts planTableWrites over an empty package) — the
 * parts, the sheets' relationships (joining a drawing's `.rels` when the
 * sheet has one), `<tableParts>` and the Overrides. */
function fromScratchTables(model: SpreadsheetModel, sheetParts: readonly string[], layer: { worksheets: string[]; entries: ZipEntryInput[] }): { entries: ZipEntryInput[]; overrides: string[] } {
  if (!model.sheets.some((sheet) => (sheet.tables ?? []).length > 0)) {
    return { entries: [], overrides: [] };
  }
  const relsEntry = (sheetPart: string): number => layer.entries.findIndex((entry) => entry.name === relsPathFor(sheetPart));
  const plan = planTableWrites(
    new Map(),
    model.sheets.map((sheet, index) => ({ sheet, partName: sheetParts[index]!, retained: false })),
    (sheetPart) => {
      const at = relsEntry(sheetPart);
      return at < 0 ? undefined : decodeUtf8(layer.entries[at]!.data);
    },
    new Set([...sheetParts, ...layer.entries.map((entry) => entry.name)]),
  );
  layer.worksheets = layer.worksheets.map((xml, index) => {
    const ids = plan.tableParts.get(sheetParts[index]!);
    return ids === undefined ? xml : spliceWorksheetTableParts(xml, ids);
  });
  const entries: ZipEntryInput[] = [];
  for (const [relsPath, text] of plan.rels) {
    const at = layer.entries.findIndex((entry) => entry.name === relsPath);
    if (at >= 0) {
      layer.entries[at] = { name: relsPath, data: encode(text) };
    } else {
      entries.push({ name: relsPath, data: encode(text) });
    }
  }
  for (const [name, text] of plan.texts) {
    entries.push({ name, data: encode(text) });
  }
  return { entries, overrides: plan.added };
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

/** "write" serialises dirty and new sheets from the model through the
 * save's string and style writers; "layout" (substitutedXlsxParts) only lays
 * the retained parts out — dirty ones included — with new sheets empty. */
type SubstitutionMode = "write" | "layout";

/** A part the save adds: its bytes, relationship and Override. */
interface AddedPart {
  name: string;
  data: Uint8Array;
  rel: string;
  contentType: string;
}

function buildSubstitution(model: SpreadsheetModel, parts: Map<string, Uint8Array>, options: WriteXlsxOptions, mode: SubstitutionMode): ZipEntryInput[] {
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
  // A copy's own closure (SPEC-121) counts as reached: a part of it the save
  // no longer relates (SPEC-133: a table part the model now holds) goes.
  const sourceReach = reachableFrom(["", ...plans.filter((plan) => !plan.generated && parts.has(plan.partName)).map((plan) => plan.partName)], edges);
  const droppedRelIds = new Set([...deletedSources.map((sheet) => sheet.rId), ...calcChainRels.map((rel) => rel.id)]);
  const afterEdges = new Map(edges);
  afterEdges.set(workbook.path, [
    ...workbook.rels.filter((rel) => !droppedRelIds.has(rel.id) && !rel.external).flatMap((rel) => targetPart(workbook.path, rel) ?? []),
    ...plans.map((plan) => plan.partName),
  ]);
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

  // SPEC-131: the rebuilt drawing layers' relationships (a save only), with
  // SPEC-132's chart parts (a patched one takes its pass before the patch).
  const drawingPlan = mode === "write" ? planDrawingWrites(model, parts, plans, (part) => ownerRemap.get(part) ?? remap) : undefined;
  for (const [owner, targets] of drawingPlan?.edges ?? []) {
    afterEdges.set(owner, targets);
  }
  // SPEC-133: the tables' parts and relationships (a save only), over the
  // sheet `.rels` the drawing layer left.
  const tablePlan = mode === "write" ? planSubstitutionTables(parts, plans, drawingPlan) : undefined;
  for (const [owner, targets] of tablePlan?.edges ?? []) {
    afterEdges.set(owner, targets);
  }
  // SPEC-135: the annotation parts and relationships (a save only), over the
  // sheet `.rels` the drawing and table writes left.
  const annotationPlan = mode === "write" ? planAnnotationWrites(model, parts, plans, drawingPlan, tablePlan) : undefined;
  for (const [owner, targets] of annotationPlan?.edges ?? []) {
    afterEdges.set(owner, targets);
  }
  const afterReach = reachableFrom([""], afterEdges);
  const skipped = new Set<string>();
  for (const part of sourceReach) {
    if (!afterReach.has(part) && part !== "") {
      skipped.add(part);
      skipped.add(relsPathFor(part));
    }
  }

  const write = mode === "write";
  const fullCalc = write && (options.fullCalcOnLoad === true || model.forceFullCalcOnLoad === true);
  const activeIndex = model.activeSheetIndex;

  // The sheet parts first: a rebuilt sheet appends to the string and style
  // tables, which are written after them.
  const tablePart = (kind: string): string | undefined => {
    const rel = workbook.rels.find((entry) => relationshipKind(entry.type) === kind && !entry.external);
    const name = rel ? targetPart(workbook.path, rel) : undefined;
    return name !== undefined && parts.has(name) ? name : undefined;
  };
  const stringsPath = tablePart("sharedStrings");
  const stylesPath = tablePart("styles");
  const strings = new SharedStringsWriter(stringsPath === undefined ? undefined : decodeUtf8(parts.get(stringsPath)!), model.theme);
  const styles = new XlsxStyleWriter(stylesPath === undefined ? undefined : decodeUtf8(parts.get(stylesPath)!), model);
  const mdw = workbookMdw(model);
  // SPEC-129: a conditional format re-emitted from its source text takes the
  // sheet-name pass the retained part took.
  const context = (plan: SheetPlan, sheetRemap?: SheetRemap): XlsxSheetContext => {
    const pass = sheetRemap !== undefined && !isEmptyRemap(sheetRemap) ? passFor(plan.partName) : undefined;
    return {
      active: plan.index === activeIndex,
      theme: model.theme,
      mdw,
      strings,
      styles,
      dxfs: model.dxfs ?? [],
      ...(pass ? { passFragment: (xml: string) => pass(xml, sheetRemap!) } : {}),
    };
  };
  const sheetBytes = new Map<string, Uint8Array>();
  // SPEC-131: a sheet whose drawing part comes or goes gets its <drawing>.
  // SPEC-133: then the `<tableParts>` the table plan rewrote.
  const withDrawing = (partName: string, bytes: Uint8Array): Uint8Array => {
    const rId = drawingPlan?.worksheetDrawing.get(partName);
    const drawn = rId === undefined ? bytes : encode(spliceWorksheetDrawing(decodeUtf8(bytes), rId));
    const ids = tablePlan?.tableParts.get(partName);
    const tabled = ids === undefined ? drawn : encode(spliceWorksheetTableParts(decodeUtf8(drawn), ids));
    // SPEC-135: then `<hyperlinks>` and `<legacyDrawing>`.
    const annotate = annotationPlan?.worksheets.get(partName);
    if (annotate === undefined) {
      return tabled;
    }
    const xml = decodeUtf8(tabled);
    const next = annotate(xml);
    return next === xml ? tabled : encode(next);
  };
  // SPEC-136: then a dirty or page-setup-edited sheet's page elements.
  const withPageSetup = (sheet: Sheet, bytes: Uint8Array): Uint8Array => {
    if (!write || sheet.kind !== "grid" || !(sheet.dirty || sheet.pageSetupDirty)) {
      return bytes;
    }
    const xml = decodeUtf8(bytes);
    const next = spliceXlsxPageSetup(xml, sheet.pageSetup);
    return next === xml ? bytes : encode(next);
  };
  for (const plan of plans) {
    const active = plan.index === activeIndex;
    if (plan.generated) {
      const xml = write && plan.sheet.kind === "grid" ? serializeWorksheet(plan.sheet, undefined, context(plan)) : emptyWorksheetXml(plan.sheet, active, model);
      sheetBytes.set(plan.partName, withPageSetup(plan.sheet, withDrawing(plan.partName, encode(xml))));
      continue;
    }
    const bytes = parts.get(plan.partName);
    if (bytes === undefined || skipped.has(plan.partName)) {
      continue;
    }
    const sheetRemap = ownerRemap.get(plan.partName) ?? remap;
    sheetBytes.set(
      plan.partName,
      withPageSetup(
        plan.sheet,
        withDrawing(
          plan.partName,
          write && plan.sheet.dirty && plan.sheet.kind === "grid"
            ? encode(serializeWorksheet(plan.sheet, retainedBase(bytes, plan, sheetRemap, model), context(plan, sheetRemap)))
            : sheetPartFor(bytes, plan, active, sheetRemap, model),
        ),
      ),
    );
  }

  // The string and style tables, appended to (or added) when a sheet needed it.
  const replaced = new Map<string, Uint8Array>();
  const added: AddedPart[] = [];
  const relIds = new Set([...workbook.rels.map((rel) => rel.id), ...plans.map((plan) => plan.rId)]);
  const relType = (kind: string): string => {
    const worksheetType = workbook.rels.find((rel) => relationshipKind(rel.type) === "worksheet")?.type ?? `${RELATIONSHIP_TYPE}/worksheet`;
    return `${worksheetType.slice(0, worksheetType.lastIndexOf("/"))}/${kind}`;
  };
  const taken = new Set([...parts.keys(), ...plans.map((plan) => plan.partName)]);
  const table = (path: string | undefined, kind: string, like: string, contentType: string, xml: string): void => {
    if (path !== undefined) {
      replaced.set(path, encode(xml));
      return;
    }
    const name = taken.has(like) ? nextFreePartName(taken, like) : like;
    taken.add(name);
    const id = nextRelationshipId(relIds);
    relIds.add(id);
    added.push({ name, data: encode(xml), rel: relationshipXml(id, relType(kind), relativeTarget(workbook.path, name)), contentType });
  };
  if (write && strings.changed) {
    table(stringsPath, "sharedStrings", "xl/sharedStrings.xml", SHARED_STRINGS_CONTENT_TYPE, strings.xml()!);
  }
  if (write && styles.changed) {
    table(stylesPath, "styles", "xl/styles.xml", STYLES_CONTENT_TYPE, styles.xml());
  }
  // SPEC-135: a minted persons part joins the workbook (a retained one is
  // among the annotation plan's texts).
  const persons = annotationPlan?.persons;
  if (persons?.minted) {
    const id = nextRelationshipId(relIds);
    relIds.add(id);
    added.push({ name: persons.partName, data: encode(persons.xml), rel: relationshipXml(id, PERSON_RELATIONSHIP_TYPE, relativeTarget(workbook.path, persons.partName)), contentType: PERSON_CONTENT_TYPE });
  }

  const out: ZipEntryInput[] = [];
  const contentTypesName = "[Content_Types].xml";
  const appPath = appPropertiesPath(parts);
  const lastWorksheet = [...parts.keys()].filter((name) => name.startsWith("xl/worksheets/") && !skipped.has(name)).at(-1);
  const generated = plans.filter((plan) => plan.generated);
  // SPEC-133: regenerated table parts and rewritten sheet `.rels` replace
  // the retained ones (never passed: model text); new ones are appended.
  const tableReplaced = new Map<string, Uint8Array>();
  const tableAdded: ZipEntryInput[] = [];
  for (const [name, xml] of [...(tablePlan?.rels ?? []), ...(tablePlan?.texts ?? [])]) {
    const data = encode(xml);
    if (parts.has(name) || drawingPlan?.added.some((entry) => entry.name === name)) {
      tableReplaced.set(name, data);
    } else {
      tableAdded.push({ name, data });
    }
  }
  // SPEC-135: the annotation plan's texts win over the drawing and table
  // writes' (it read the sheet `.rels` they left); its new parts follow.
  const annotationText = (name: string): Uint8Array | undefined => annotationPlan?.texts.get(name);
  const knownNames = new Set([...parts.keys(), ...(drawingPlan?.added ?? []).map((entry) => entry.name), ...tableAdded.map((entry) => entry.name)]);
  const annotationAdded = [...(annotationPlan?.texts ?? [])].filter(([name]) => !knownNames.has(name)).map(([name, data]) => ({ name, data }));
  const appended = (): ZipEntryInput[] => [
    ...generated.map((plan) => ({ name: plan.partName, data: sheetBytes.get(plan.partName)! })),
    ...added.map((part) => ({ name: part.name, data: part.data })),
    ...(drawingPlan?.added ?? []).map((entry) => ({ name: entry.name, data: annotationText(entry.name) ?? tableReplaced.get(entry.name) ?? entry.data })),
    ...tableAdded.map((entry) => ({ name: entry.name, data: annotationText(entry.name) ?? entry.data })),
    ...annotationAdded,
  ];
  // SPEC-131 "Media parts": a model media part is written from the model.
  const modelMedia = (name: string, bytes: Uint8Array): Uint8Array | undefined => {
    const part = drawingPlan?.media.partAt(name);
    return part ? mediaBytes(part, bytes) : undefined;
  };

  for (const [name, bytes] of parts) {
    if (skipped.has(name)) {
      continue;
    }
    if (calcChainRels.some((rel) => targetPart(workbook.path, rel) === name)) {
      continue;
    }
    let data = bytes;
    if (name === contentTypesName) {
      data = encode(contentTypesFor(decodeUtf8(bytes), parts, skipped, plans, workbook, calcChainRels, added, drawingPlan, tablePlan?.added ?? [], annotationPlan));
    } else if (name === workbook.path) {
      data = encode(workbookXmlFor(workbook, plans, remap, activeIndex, fullCalc, workbookNames(model, plans, write ? filterDatabaseScope(model, parts) : undefined)));
    } else if (name === workbook.relsPath) {
      data = encode(workbookRelsFor(workbook, plans, droppedRelIds, added));
    } else if (name === appPath) {
      data = encode(appXmlFor(decodeUtf8(bytes), workbook, plans, remap));
    } else {
      data =
        sheetBytes.get(name) ??
        annotationText(name) ??
        tableReplaced.get(name) ??
        drawingPlan?.replaced.get(name) ??
        replaced.get(name) ??
        modelMedia(name, bytes) ??
        passed(name, bytes, ownerRemap.get(name) ?? remap);
    }
    out.push({ name, data });
    if (name === lastWorksheet) {
      out.push(...appended());
    }
  }
  if (lastWorksheet === undefined) {
    out.push(...appended());
  }
  return out;
}

/** SPEC-133 "Names on the wire": the names workbookXmlFor rebuilds
 * `<definedNames>` from — the model's, the `_xlnm._FilterDatabase`s of the
 * sheets `reconciled` accepts brought in line (a save); undefined (the part's
 * names left as written) when the model holds none to speak of (a record no
 * reader upgrade gave names) or nothing changed them. */
function workbookNames(model: SpreadsheetModel, plans: readonly SheetPlan[], reconciled: ((plan: SheetPlan) => boolean) | undefined): readonly DefinedName[] | undefined {
  if (model.names === undefined) {
    return undefined;
  }
  const names = reconciled === undefined ? model.names : namesToWrite(model, plans, reconciled);
  return model.namesDirty || model.structureDirty || names !== model.names ? names : undefined;
}

/** SPEC-133 "Tables on the wire": the save's table writes over the planned
 * sheets, each sheet's `.rels` as the drawing layer left it. */
function planSubstitutionTables(parts: Map<string, Uint8Array>, plans: readonly SheetPlan[], drawingPlan: DrawingPlan | undefined): TableWritePlan {
  const relsText = (sheetPart: string): string | undefined => {
    const relsPath = relsPathFor(sheetPart);
    const bytes = drawingPlan?.replaced.get(relsPath) ?? drawingPlan?.added.find((entry) => entry.name === relsPath)?.data ?? parts.get(relsPath);
    return bytes === undefined ? undefined : decodeUtf8(bytes);
  };
  const taken = new Set([...parts.keys(), ...plans.map((plan) => plan.partName), ...(drawingPlan?.added ?? []).map((entry) => entry.name)]);
  return planTableWrites(
    parts,
    plans.map((plan) => ({ sheet: plan.sheet, partName: plan.partName, retained: !plan.generated && parts.has(plan.partName) })),
    relsText,
    taken,
  );
}

/** A dirty sheet's retained text with the pass and its tab colour applied —
 * the base serializeWorksheet splices the model's elements into. */
function retainedBase(bytes: Uint8Array, plan: SheetPlan, remap: SheetRemap, model: SpreadsheetModel): string {
  const xml = decodeUtf8(bytes);
  const pass = !isEmptyRemap(remap) && mentionsAnyName(bytes, remappedNames(remap)) ? passFor(plan.partName) : undefined;
  return spliceTabColor(pass ? pass(xml, remap) : xml, plan.sheet, model);
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

/** SPEC-128: which sheets' `_xlnm._FilterDatabase` names a save reconciles
 * with the model — every grid sheet once the structure changed (sheets
 * added, copied, moved, renamed or deleted), else the dirty ones (a
 * structural edit moves an autofilter without touching the structure) —
 * among those the model owns the names of: a sheet with an autofilter, or
 * one whose source part held one (autoFilter.ts's sheetPartHasAutoFilter).
 * A clean sheet's names stay as the source wrote them; an Advanced Filter's
 * go through the ordinary defined-name splice. */
function filterDatabaseScope(model: SpreadsheetModel, parts: ReadonlyMap<string, Uint8Array>): (plan: SheetPlan) => boolean {
  const hadAutoFilter = new Map<string, boolean>();
  const sourceHad = (partName: string): boolean => {
    let had = hadAutoFilter.get(partName);
    if (had === undefined) {
      const bytes = parts.get(partName);
      had = bytes !== undefined && sheetPartHasAutoFilter(decodeUtf8(bytes));
      hadAutoFilter.set(partName, had);
    }
    return had;
  };
  const key = FILTER_DATABASE_NAME.toUpperCase();
  const named = (sheetId: string): boolean => (model.names ?? []).some((name) => name.scopeSheetId === sheetId && name.name.toUpperCase() === key);
  return (plan) =>
    plan.sheet.kind === "grid" &&
    (model.structureDirty === true || plan.sheet.dirty === true || plan.source === undefined) &&
    (plan.sheet.autoFilter !== undefined || (plan.source !== undefined ? sourceHad(plan.source.partName) : named(plan.sheet.id)));
}

/** CT_Workbook's children after `calcPr`: an absent `calcPr` is created
 * before the first of these. */
const AFTER_CALC_PR = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];

/** The rebuilt workbook part: `<sheets>`, every workbookView's `activeTab`
 * (no `firstSheet`), SPEC-133's `<definedNames>` from `names` when given
 * (xlsx/definedNames.ts definedNamesSplices: `localSheetId` recomputed from
 * the scope sheet's current index, a name matched to its source element
 * keeping its bytes) and `calcPr@fullCalcOnLoad`. */
function workbookXmlFor(
  workbook: SourceWorkbook,
  plans: readonly SheetPlan[],
  remap: SheetRemap,
  activeIndex: number,
  fullCalc: boolean,
  names: readonly DefinedName[] | undefined,
): string {
  const xml = workbook.xml;
  const root = rootTag(xml);
  const relPrefix = relationshipPrefix(root.tag);
  const splices: Splice[] = [];

  let depth = 0;
  let sheetsRange: { start: number; end: number } | null = null;
  let sheetsStart = -1;
  let calcPr: { start: number; end: number } | null = null;
  let insertCalcPrAt = -1;
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
      }
      if (event.selfClosing) {
        depth -= 1;
      }
    } else if (event.type === "close") {
      if (depth === 2 && event.local === "sheets" && sheetsStart >= 0) {
        sheetsRange = { start: sheetsStart, end: event.end };
      } else if (depth === 1) {
        rootClose = event.start;
      }
      depth -= 1;
    }
  });

  if (names !== undefined) {
    // Before the calcPr splices: a definedNames created where an absent
    // calcPr is also created goes first.
    const index = new Map(plans.map((plan) => [plan.sheet.id, plan.index] as const));
    const sourceIndex = new Map(plans.flatMap((plan) => (plan.source ? [[plan.sheet.id, workbook.sheets.indexOf(plan.source)] as const] : [])));
    splices.push(
      ...definedNamesSplices(xml, scanDefinedNames(xml), names, {
        indexOf: (sheetId) => index.get(sheetId),
        sourceIndexOf: (sheetId) => sourceIndex.get(sheetId),
        remap,
      }),
    );
  }

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
        // A root that binds no prefix to the relationships namespace (an
        // openpyxl workbook declares it per <sheet>) gets it declared here.
        const declare = relPrefix === null ? ` xmlns:r="${REL_NS}"` : "";
        return `<${sheetName} name="${escapeAttributeValue(plan.sheet.name)}" sheetId="${plan.sheetId}"${sheetStateAttr(plan.sheet)}${declare} ${relPrefix ?? "r:"}id="${plan.rId}"/>`;
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
 * included; null when the root binds none. */
function relationshipPrefix(rootStartTag: string): string | null {
  for (const match of rootStartTag.matchAll(/\sxmlns:([\w.-]+)\s*=\s*"([^"]*)"/g)) {
    if (match[2] === REL_NS || match[2] === "http://purl.oclc.org/ooxml/officeDocument/relationships") {
      return `${match[1]}:`;
    }
  }
  return null;
}

function decodeEntities(raw: string): string {
  return raw.replace(/&(amp|lt|gt|quot|apos);/g, (_whole, name: string) => ({ amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" })[name] ?? "");
}

function escapeText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// --- relationships and content types ----------------------------------------------------------

function workbookRelsFor(workbook: SourceWorkbook, plans: readonly SheetPlan[], dropped: ReadonlySet<string>, parts: readonly AddedPart[]): string {
  const sourceIds = new Set(workbook.rels.map((rel) => rel.id));
  const worksheetType = workbook.rels.find((rel) => relationshipKind(rel.type) === "worksheet")?.type ?? `${RELATIONSHIP_TYPE}/worksheet`;
  const added = plans
    .filter((plan) => !plan.source && !sourceIds.has(plan.rId))
    .map((plan) => relationshipXml(plan.rId, worksheetType, relativeTarget(workbook.path, plan.partName)));
  return spliceRels(workbook.relsXml, dropped, [...added, ...parts.map((part) => part.rel)]);
}

/** The content type a part added by duplication takes: that of a part in
 * the same folder with the same stem (its origin's). */
function stemContentType(types: ReturnType<typeof readContentTypes>, partName: string): string | undefined {
  const key = partStem(partName);
  return types.overrides.find((entry) => partStem(entry.partName) === key)?.contentType;
}

/** A part name without the number before its extension (`chart2.xml`,
 * `table1_2.xml` → `chart.xml`, `table1.xml`). */
function partStem(name: string): string {
  return name.replace(/(?:_\d+|\d+)(\.[^./]*)$/, "$1");
}

function contentTypesFor(
  xml: string,
  parts: Map<string, Uint8Array>,
  skipped: ReadonlySet<string>,
  plans: readonly SheetPlan[],
  workbook: SourceWorkbook,
  calcChainRels: readonly { target: string; external: boolean; id: string; type: string; start: number; end: number }[],
  addedParts: readonly AddedPart[],
  drawingPlan: DrawingPlan | undefined,
  tableParts: readonly string[],
  annotationPlan: AnnotationPlan | undefined,
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
  for (const part of addedParts) {
    add.push({ partName: part.name, contentType: part.contentType });
  }
  add.push(...(drawingPlan?.overrides ?? []));
  add.push(...tableParts.map((partName) => ({ partName, contentType: TABLE_CONTENT_TYPE })));
  add.push(...(annotationPlan?.overrides ?? []));
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
  const out = remove.size === 0 && add.length === 0 ? xml : spliceContentTypes(xml, remove, add);
  const defaults = new Map(drawingPlan?.defaults ?? []);
  if (annotationPlan?.vml === true && !types.defaults.has("vml")) {
    defaults.set("vml", VML_CONTENT_TYPE);
  }
  return spliceDefaults(out, defaults);
}

/** SPEC-135: what a save's annotation writes change in the package. */
interface AnnotationPlan {
  /** Parts written: comments, VML, threaded parts, sheet `.rels`, a retained persons part. */
  texts: Map<string, Uint8Array>;
  /** Each written sheet part's relationship targets after the save. */
  edges: Map<string, string[]>;
  /** Each written sheet's `<hyperlinks>`/`<legacyDrawing>` splice. */
  worksheets: Map<string, (xml: string) => string>;
  /** Minted parts' Overrides. */
  overrides: { partName: string; contentType: string }[];
  /** A VML part was minted (the `vml` Default may be needed). */
  vml: boolean;
  persons?: PersonsWrite;
}

/** Every `dirty` or `annotationsDirty` grid sheet's annotations written
 * (xlsx/annotations.ts writeXlsxSheetAnnotations) over the sheet `.rels`
 * the drawing and table writes left, then the persons. A part rewritten to
 * its source bytes is not one. */
function planAnnotationWrites(
  model: SpreadsheetModel,
  parts: Map<string, Uint8Array>,
  plans: readonly SheetPlan[],
  drawingPlan: DrawingPlan | undefined,
  tablePlan: TableWritePlan | undefined,
): AnnotationPlan {
  const plan: AnnotationPlan = { texts: new Map(), edges: new Map(), worksheets: new Map(), overrides: [], vml: false };
  const taken = new Set([
    ...parts.keys(),
    ...plans.map((each) => each.partName),
    ...(drawingPlan?.added ?? []).map((entry) => entry.name),
    ...(tablePlan?.texts.keys() ?? []),
    ...(tablePlan?.rels.keys() ?? []),
  ]);
  let idmap = maxVmlIdmap(parts);
  const ctx: AnnotationWriteContext = { parts, taken, nextIdmap: () => (idmap += 1), persons: model.persons };
  const relsText = (sheetPart: string): string | undefined => {
    const relsPath = relsPathFor(sheetPart);
    const tabled = tablePlan?.rels.get(relsPath);
    if (tabled !== undefined) {
      return tabled;
    }
    const bytes = drawingPlan?.replaced.get(relsPath) ?? drawingPlan?.added.find((entry) => entry.name === relsPath)?.data ?? parts.get(relsPath);
    return bytes === undefined ? undefined : decodeUtf8(bytes);
  };
  for (const sheetPlan of plans) {
    if (!annotationsRebuilt(sheetPlan.sheet)) {
      continue;
    }
    const retained = !sheetPlan.generated ? parts.get(sheetPlan.partName) : undefined;
    const write = writeXlsxSheetAnnotations(sheetPlan.sheet, sheetPlan.partName, relsText(sheetPlan.partName), retained === undefined ? undefined : decodeUtf8(retained), ctx);
    if (write === undefined) {
      continue;
    }
    for (const [name, text] of write.texts) {
      const data = encode(text);
      const source = parts.get(name);
      if (source === undefined || !equalBytes(source, data)) {
        plan.texts.set(name, data);
      }
    }
    plan.edges.set(sheetPlan.partName, write.edges);
    plan.worksheets.set(sheetPlan.partName, write.worksheet);
    plan.overrides.push(...write.added.flatMap((part) => (part.contentType === undefined ? [] : [{ partName: part.partName, contentType: part.contentType }])));
    plan.vml ||= write.vml;
  }
  const persons = writeXlsxPersons(model, parts, taken);
  if (persons !== undefined) {
    plan.persons = persons;
    if (!persons.minted) {
      plan.texts.set(persons.partName, encode(persons.xml));
    }
  }
  return plan;
}

/** `[Content_Types].xml` with a `<Default>` per extension in `defaults`,
 * after the last Default (else first inside the root). */
function spliceDefaults(xml: string, defaults: ReadonlyMap<string, string>): string {
  if (defaults.size === 0) {
    return xml;
  }
  let at = -1;
  let rootEnd = -1;
  let prefix = "";
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === "Types") {
      rootEnd = event.end;
      prefix = event.name.includes(":") ? event.name.slice(0, event.name.indexOf(":") + 1) : "";
    } else if ((event.type === "open" || event.type === "close") && event.local === "Default") {
      at = event.end;
    }
  });
  const elements = [...defaults]
    .map(([extension, type]) => `<${prefix}Default Extension="${escapeAttributeValue(extension)}" ContentType="${escapeAttributeValue(type)}"/>`)
    .join("");
  const insert = at >= 0 ? at : rootEnd;
  return insert < 0 ? xml : xml.slice(0, insert) + elements + xml.slice(insert);
}

// --- SPEC-131: drawing layers and media ----------------------------------------------------------

/** A media part path the xlsx writers keep as it is. */
const XLSX_MEDIA_PATH = /^xl\/media\/[^/]+$/;

/** Where a save writes each media part ("Media parts"). */
interface MediaWrites {
  /** The package path a mediaId is written at: its own when it is an
   * `xl/media/` path, else the next free `xl/media/image{n}.{ext}` (once per
   * id); undefined for an id neither the model nor the package holds. */
  pathOf(mediaId: string): string | undefined;
  /** The model media part written at a package path. */
  partAt(path: string): MediaPart | undefined;
}

function mediaWrites(model: Pick<SpreadsheetModel, "media">, taken: Set<string>): MediaWrites {
  const byId = new Map((model.media ?? []).map((part) => [part.mediaId, part] as const));
  const byPath = new Map([...byId].filter(([id]) => XLSX_MEDIA_PATH.test(id)));
  const assigned = new Map<string, string>();
  const nextPath = (contentType: string): string => {
    let max = 0;
    for (const name of [...taken, ...byPath.keys()]) {
      const match = /^xl\/media\/image(\d+)\.[^/]*$/.exec(name);
      if (match) {
        max = Math.max(max, Number(match[1]));
      }
    }
    return `xl/media/image${max + 1}.${extensionForContentType(contentType)}`;
  };
  return {
    pathOf(mediaId) {
      const part = byId.get(mediaId);
      if (!part) {
        return taken.has(mediaId) ? mediaId : undefined;
      }
      if (XLSX_MEDIA_PATH.test(mediaId)) {
        return mediaId;
      }
      let path = assigned.get(mediaId);
      if (path === undefined) {
        path = nextPath(part.contentType);
        assigned.set(mediaId, path);
        byPath.set(path, part);
        taken.add(path);
      }
      return path;
    },
    partAt: (path) => byPath.get(path),
  };
}

/** A media part's bytes: the retained twin's while the part is still the
 * bytes as read (its encode deferred), else its decoded `data`. */
function mediaBytes(part: MediaPart, source: Uint8Array | undefined): Uint8Array {
  if (source !== undefined && pendingMediaByteLength(part) > 0) {
    return source;
  }
  ensureMediaEncoded(part);
  return part.data === "" && source !== undefined ? source : base64ToBytes(part.data);
}

/** What a save's drawing layers change in the package. */
interface DrawingPlan {
  /** Retained parts rewritten (drawing parts, their `.rels`, sheet `.rels`). */
  replaced: Map<string, Uint8Array>;
  /** Parts added (new drawing parts and `.rels`, new media). */
  added: ZipEntryInput[];
  /** Relationship edges after the save, by owner. */
  edges: Map<string, string[]>;
  /** Worksheet parts whose `<drawing r:id>` is set (an id) or removed (null). */
  worksheetDrawing: Map<string, string | null>;
  overrides: { partName: string; contentType: string }[];
  /** Extensions `[Content_Types].xml` gains a `<Default>` for. */
  defaults: Map<string, string>;
  media: MediaWrites;
}

/** Every `drawingsDirty` grid sheet's layer written (xlsx/drawingXml.ts
 * writeSheetDrawings) — and SPEC-132, a grid sheet's holding a created chart
 * (no anchor XML of its own) — then the new media parts the written drawing
 * parts target. A rewrite that reproduces its source's bytes is not one.
 * SPEC-132: the chart parts (xlsx/chartParts.ts) join it — a created chart's
 * new part with its Override, and every retained chart part the model
 * regenerates or patches, always written as planned (model text: never
 * passed). A chart sheet's drawing part is never rewritten. */
function planDrawingWrites(model: SpreadsheetModel, parts: Map<string, Uint8Array>, plans: readonly SheetPlan[], remapOf: (partName: string) => SheetRemap): DrawingPlan {
  const taken = new Set([...parts.keys(), ...plans.map((plan) => plan.partName)]);
  const media = mediaWrites(model, taken);
  const charts = new XlsxChartWriter(model, parts, taken, remapOf);
  const plan: DrawingPlan = { replaced: new Map(), added: [], edges: new Map(), worksheetDrawing: new Map(), overrides: [], defaults: new Map(), media };
  for (const sheetPlan of plans) {
    const { sheet } = sheetPlan;
    if (sheet.kind !== "grid" || !(sheet.drawingsDirty || sheetHasCreatedChart(sheet))) {
      continue;
    }
    const retained = !sheetPlan.generated && parts.has(sheetPlan.partName);
    const chartPart = (drawing: SheetDrawing): string | undefined => charts.partFor(sheet, drawing);
    const write = writeSheetDrawings({ parts, sheet, sheetPart: sheetPlan.partName, retained, taken, mediaPath: media.pathOf, chartPart });
    for (const [name, text] of write?.texts ?? []) {
      const data = encode(text);
      const source = parts.get(name);
      if (source === undefined) {
        plan.added.push({ name, data });
      } else if (!equalBytes(source, data)) {
        plan.replaced.set(name, data);
      }
    }
    for (const [owner, targets] of write?.edges ?? []) {
      plan.edges.set(owner, targets);
    }
    if (write?.worksheetDrawing !== undefined) {
      plan.worksheetDrawing.set(sheetPlan.partName, write.worksheetDrawing);
    }
    if (write?.addedPart !== undefined) {
      plan.overrides.push({ partName: write.addedPart, contentType: DRAWING_CONTENT_TYPE });
    }
  }
  const chartWrites = charts.writes();
  for (const [name, data] of chartWrites.replaced) {
    plan.replaced.set(name, data);
  }
  plan.added.push(...chartWrites.added);
  plan.overrides.push(...chartWrites.overrides);
  const typesBytes = parts.get("[Content_Types].xml");
  const types = typesBytes === undefined ? undefined : readContentTypes(decodeUtf8(typesBytes));
  for (const path of new Set([...plan.edges.values()].flat())) {
    const part = parts.has(path) ? undefined : media.partAt(path);
    if (!part) {
      continue;
    }
    plan.added.push({ name: path, data: mediaBytes(part, undefined) });
    const extension = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    if (types?.defaults.has(extension) !== true && !plan.defaults.has(extension)) {
      plan.defaults.set(extension, part.contentType);
    }
  }
  return plan;
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
  // A chart sheet is "chart" (SPEC-132) or, unresolvable, "foreign".
  const modelCharts = plans
    .filter((plan) => plan.sheet.kind !== "grid" && plan.source !== undefined && relationshipKind(plan.source.relType) === "chartsheet")
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
  return applySplices(xml, [
    { start: heading.index, end: heading.index + heading[0].length, text: `${openTagOf(heading)}${headingBody}${closeTagOf(heading)}` },
    { start: titles.index, end: titles.index + titles[0].length, text: `${openTagOf(titles)}${titlesBody}${closeTagOf(titles)}` },
  ]);
}

/** The start tag of a matched `<x …>…</x>` element. */
function openTagOf(match: RegExpExecArray): string {
  return match[0].slice(0, match[0].indexOf(">") + 1);
}

/** The close tag of a matched `<x …>…</x>` element. */
function closeTagOf(match: RegExpExecArray): string {
  return match[0].slice(match[0].lastIndexOf("</"));
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
