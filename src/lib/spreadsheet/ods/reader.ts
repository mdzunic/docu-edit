// The .ods reader (SPEC-116, SPEC-117): an ODF package whose mimetype
// declares a spreadsheet (or its template) → the workbook model. content.xml
// is read ONCE, start to finish, with the forward-only scanner
// (core/xmlScanner.ts), whose exact offsets are also how each table's
// verbatim XML is sliced out: `sourceXml` is the whitespace before a
// `<table:table>` plus the element itself, and the model's
// `odsContentEnvelope` holds everything before the first and after the last
// — so prefix + every sourceXml + suffix is content.xml, exactly (the splice
// base SPEC-121/122's writers build on).
//
// What is read:
// - styles (ods/styles.ts): styles.xml's default, common and font-face
//   declarations, and content.xml's `office:font-face-decls` and
//   `office:automatic-styles` — sliced out of the scan by offset (they
//   precede the body) and parsed whole — flattened into `styles` (index 0
//   the "Default" style, then every referenced cell style name in
//   first-reference order); SPEC-123: styles.xml's common table-cell styles
//   → `namedStyles`, and each cell style's `namedStyle`;
// - `table:table` elements in order (each one a grid sheet); a table whose
//   style has `table:display="false"` is hidden, its
//   `tableooo:tab-color` is the tab colour;
// - each table's columns, rows and cells (ods/tableXml.ts): column widths,
//   visibility and default cell styles; row heights, custom heights,
//   visibility (hidden or filter-hidden) and default cell styles, the
//   defaults taken from the last column and row runs; cell values and style
//   ids; styled empty cells (stored only where they differ from the
//   effective default, a run to XFD becoming the row's style); covered cells
//   with a value or a style; spans → merges. `table:number-rows-repeated` / `table:number-columns-repeated`
//   are honoured without ever materialising a repeated run that holds
//   nothing to store (LibreOffice's trailing 1048553-row and 16380-column
//   runs, styled or not, cost nothing); everything materialised is counted
//   against MAX_WORKBOOK_CELLS before it is built;
// - rows nested in header rows, row groups and `table:table-rows`, in
//   document order; a table in the sheet's drawing layer is not the sheet's;
// - values by `office:value-type` (float/percentage/currency → number,
//   date/time → serial, boolean, string → `office:string-value` or the
//   `text:p` text with paragraphs joined by "\n" and `text:s`/`text:tab`/
//   `text:line-break` expanded); `calcext:value-type="error"` → the error
//   code from the text; a date, time or percentage value with no data style
//   gets its implied format (SPEC-118); a formula cell contributes its
//   cached value; a string whose paragraphs hold `text:span`s with a font of
//   their own gets `runs` (each span's font from its automatic text style);
// - SPEC-119: `table:formula` → `formula` through fromOdsFormula, its
//   dialect prefix resolved against the `xmlns:*` declarations on
//   `office:document-content` (a formula that does not parse keeps the whole
//   attribute value, prefix included, `formulaUnparsed`); a formula cell with
//   no `office:value-type` keeps `value: null`; the matrix spans →
//   `arrayRange` (tableXml.ts);
// - `table:null-date` → the date system (see serial.ts); settings.xml
//   (ods/settings.ts) → the active table and each table's frozen panes,
//   scroll position, gridlines, zero display, zoom and cursor; meta.xml →
//   properties;
// - SPEC-120: `table:calculation-settings` (case sensitivity, regular
//   expressions, wildcards, whole-cell matching, `table:iteration`) and
//   settings.xml's AutoCalculate → `calc`, read only.
// - SPEC-122 "Model additions": what the table writer re-emits for content
//   the model does not hold. Each cell's unmodelled start-tag attributes —
//   all but style-name, the repeat, span and matrix-span counts,
//   `table:formula`, `office:value-type`, the value attributes,
//   `office:currency` and `calcext:value-type` — verbatim with their leading
//   whitespace, in source order (`preservedAttrs`); every child element but
//   the text paragraphs (annotations, frames, unknown elements), verbatim in
//   source order (`preservedXml`); and the cell's `text:p` sequence verbatim
//   (`preservedText`) when a paragraph holds inline content the model does
//   not — an attribute on the paragraph, a heading, or any element other
//   than text:s, text:tab, text:line-break and a text:span whose only
//   attribute is a style holding nothing but a run font (a link, a field, a
//   note, a frame…). Whitespace between a cell's children is not content. A
//   cell holding only preserved content is stored, every repetition counted
//   against the budget; past the grid it shows nothing, like a style.
//   `table:protected="true"` → `Sheet.protection`;
//   `office:spreadsheet@table:structure-protected="true"` →
//   `structureProtected` (both read only — the XML stays in the source).
// - SPEC-128: a row's `table:visibility="filter"` → `filterHidden` and
//   `"collapse"` → `hidden`, exactly as the wire says (tableXml.ts; ODS rows
//   are never classified); the suffix's `table:database-range` that is a
//   table's autofilter (ods/databaseRanges.ts's claim) → `Sheet.autoFilter`
//   (its criteria, sort record, verbatim element and name; a filter that
//   does not decompose per column kept as `odsUnmappedFilter`).
// - SPEC-129: each table's `calcext:conditional-formats` (else the legacy
//   `style:map`s of its cells' automatic styles) → `conditionalFormats`,
//   the named styles they apply → `dxfs` (ods/conditionalFormats.ts, after
//   the scan, from each sheet's sourceXml).
// - SPEC-130: content.xml's `table:content-validations` (in the envelope
//   prefix) and the cells' `table:content-validation-name` → each sheet's
//   `validations` and the workbook's `odsOrphanValidations`
//   (ods/contentValidations.ts): the names are accumulated during the scan
//   as row-run × column-run rectangles, never materialised — the attribute
//   is modelled, so it is no longer one of a cell's `preservedAttrs`.
import { detectPackage, spreadsheetFormatOf } from "../../core/packageDetect";
import { parseOdfMeta } from "../../core/odfMeta";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE } from "../../core/odfXml";
import { XmlParseError } from "../../core/xml";
import { scanAttr, scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { unzip, ZipError } from "../../core/zip";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import type { GridRange } from "../address";
import { toCellError, type CalcSettings, type CellFont, type CellStyle, type CellValue, type Sheet, type TextRun } from "../model";
import {
  addressText,
  CellBudget,
  encodeOpaque,
  parseError,
  partText,
  resolveActiveSheet,
  type ParsedSpreadsheet,
} from "../parsed";
import { IMPLIED_PERCENTAGE, IMPLIED_TIME, impliedDateCode } from "../impliedFormats";
import { odsDateValueToSerial, odsTimeValueToSerial } from "../serial";
import { readOdsFormula } from "../formula/dialect";
import type { NamedCellStyle } from "../styles";
import type { Dxf } from "../conditionalFormats";
import { readOdsTableConditionalFormats, type OdsCfReadContext } from "./conditionalFormats";
import { readOdsContentValidations, ValidationRegions } from "./contentValidations";
import { readOdsSettings, type OdsSettings } from "./settings";
import {
  emptyOdsStyleSheet,
  namespaceDeclarations,
  OdsStyleBook,
  readOdsContentStyles,
  readOdsStyleSheet,
  type OdsStyleSheet,
} from "./styles";
import { readOdsAutoFilters } from "./databaseRanges";
import { OdsTableReader, repeatCount, type OdsCellFormula, type OdsCellPreserved } from "./tableXml";

/** LibreOffice's extension namespace (`calcext:value-type="error"`). */
const CALCEXT_NAMESPACE = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

/** The null date ODF assumes when a document declares none. */
const DEFAULT_NULL_DATE = "1899-12-30";

interface PendingTable {
  name: string;
  hidden: boolean;
  tabColor?: string;
  /** Where this sheet's sourceXml starts in content.xml. */
  start: number;
  level: number;
  reader: OdsTableReader;
  /** SPEC-122: `table:protected="true"`. */
  protected: boolean;
  /** SPEC-130: the cells naming each content validation. */
  validations: ValidationRegions;
}

/** A stretch of a cell's text sharing one span font. */
interface RunDraft {
  text: string;
  font: CellFont | undefined;
}

interface PendingCell {
  level: number;
  covered: boolean;
  attrs: readonly XmlScanAttribute[];
  paragraphs: string[];
  /** The text:p being collected, or null. */
  paragraph: string | null;
  paragraphLevel: number;
  /** >0 while inside a subtree whose text is not the cell's (an annotation,
   * a drawing). */
  skipDepth: number;
  /** The cell's text as runs (paragraphs separated by "\n"). */
  runs: RunDraft[];
  /** The open `text:span`s, innermost last, with the font each gives. */
  spans: { level: number; font: CellFont | undefined }[];
  /** Whether any span gave a font — only then are runs kept. */
  hasSpanFont: boolean;
  // --- SPEC-122 captures ---
  /** The unmodelled start-tag attributes, verbatim (empty: none). */
  preservedAttrs: string;
  /** The child elements other than paragraphs, verbatim (empty: none). */
  preservedXml: string;
  /** Each paragraph child's [start, end) offsets in content.xml. */
  paragraphRanges: number[];
  /** Where the open direct child element starts (-1: none open), and
   * whether it is a paragraph. */
  childStart: number;
  childIsParagraph: boolean;
  /** A paragraph holds inline content the model does not. */
  inline: boolean;
}

/** The cell attributes the model holds or the writer regenerates (SPEC-122
 * "Model additions"), by namespace and local name. */
const MODELLED_CELL_ATTRIBUTES: ReadonlyMap<string, ReadonlySet<string>> = new Map([
  [
    TABLE_NAMESPACE,
    new Set([
      "style-name",
      "number-columns-repeated",
      "number-columns-spanned",
      "number-rows-spanned",
      "formula",
      "number-matrix-columns-spanned",
      "number-matrix-rows-spanned",
      // SPEC-130: the model's Sheet.validations.
      "content-validation-name",
    ]),
  ],
  [OFFICE_NAMESPACE, new Set(["value-type", "value", "date-value", "time-value", "boolean-value", "string-value", "currency"])],
  [CALCEXT_NAMESPACE, new Set(["value-type"])],
]);

function isModelledCellAttribute(attr: XmlScanAttribute): boolean {
  return attr.ns !== undefined && (MODELLED_CELL_ATTRIBUTES.get(attr.ns)?.has(attr.local) ?? false);
}

/** A raw attribute of a start tag: its leading whitespace, name and quoted
 * value, verbatim. */
const RAW_ATTRIBUTE = /(\s+)([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g;

/** A cell's unmodelled start-tag attributes, verbatim with their leading
 * whitespace and in source order — "" when every attribute is modelled (the
 * overwhelmingly common case, decided without touching the raw text). */
function unmodelledAttributes(xml: string, start: number, end: number, attrs: readonly XmlScanAttribute[]): string {
  if (attrs.every(isModelledCellAttribute)) {
    return "";
  }
  const kept = new Set(attrs.filter((attr) => !isModelledCellAttribute(attr)).map((attr) => attr.name));
  const tag = xml.slice(start, end);
  let out = "";
  for (const match of tag.matchAll(RAW_ATTRIBUTE)) {
    if (kept.has(match[2]!)) {
      out += match[0];
    }
  }
  return out;
}

/** The captures a finished cell hands the table reader, or undefined. */
function cellPreserved(cell: PendingCell, xml: string): OdsCellPreserved | undefined {
  if (cell.preservedAttrs === "" && cell.preservedXml === "" && !cell.inline) {
    return undefined;
  }
  const preserved: OdsCellPreserved = {};
  if (cell.preservedAttrs !== "") {
    preserved.attrs = cell.preservedAttrs;
  }
  if (cell.preservedXml !== "") {
    preserved.xml = cell.preservedXml;
  }
  if (cell.inline && cell.paragraphRanges.length > 0) {
    let text = "";
    for (let i = 0; i < cell.paragraphRanges.length; i += 2) {
      text += xml.slice(cell.paragraphRanges[i], cell.paragraphRanges[i + 1]);
    }
    preserved.text = text;
  }
  return preserved.attrs === undefined && preserved.xml === undefined && preserved.text === undefined ? undefined : preserved;
}

/** Records a finished direct child of a cell: a paragraph's offsets, or any
 * other element verbatim. */
function captureChild(cell: PendingCell, xml: string, paragraph: boolean, start: number, end: number): void {
  if (paragraph) {
    cell.paragraphRanges.push(start, end);
  } else {
    cell.preservedXml += xml.slice(start, end);
  }
}

/** A cell's formula (SPEC-119), read at the position its first repetition
 * takes; undefined for a cell without one. */
function cellFormula(
  attrs: readonly XmlScanAttribute[],
  namespaces: ReadonlyMap<string, string>,
  host: { row: number; col: number },
): OdsCellFormula | undefined {
  const wire = scanAttr(attrs, "formula", TABLE_NAMESPACE);
  if (wire === undefined) {
    return undefined;
  }
  const formula: OdsCellFormula = readOdsFormula(wire, namespaces, host);
  const spans = ["number-matrix-columns-spanned", "number-matrix-rows-spanned"];
  if (spans.some((local) => scanAttr(attrs, local, TABLE_NAMESPACE) !== undefined)) {
    formula.matrix = { cols: repeatCount(attrs, spans[0]!), rows: repeatCount(attrs, spans[1]!) };
  }
  return formula;
}

/** The date system a declared null date selects (see serial.ts). */
function dateSystemOf(nullDate: string): { dateSystem: "1900" | "1904"; odsNullDate?: string } {
  const day = nullDate.trim().slice(0, 10);
  if (day === DEFAULT_NULL_DATE) {
    return { dateSystem: "1900" };
  }
  if (day === "1904-01-01") {
    return { dateSystem: "1904" };
  }
  return { dateSystem: "1900", odsNullDate: day };
}

/** A cell's value from its declared type; null for an empty cell. */
function odsCellValue(
  cell: PendingCell,
  dates: { dateSystem: "1900" | "1904"; odsNullDate?: string },
  where: string,
): CellValue {
  const officeAttr = (local: string): string | undefined => scanAttr(cell.attrs, local, OFFICE_NAMESPACE);
  const text = cell.paragraphs.join("\n");
  const valueType =
    scanAttr(cell.attrs, "value-type", CALCEXT_NAMESPACE) === "error" ? "error" : officeAttr("value-type");
  switch (valueType) {
    case "float":
    case "percentage":
    case "currency": {
      const raw = officeAttr("value");
      const n = raw === undefined ? Number.NaN : Number(raw);
      if (!Number.isFinite(n)) {
        throw parseError(where, `"${raw ?? ""}" is not a number`);
      }
      return n;
    }
    case "date": {
      const raw = officeAttr("date-value") ?? "";
      try {
        return odsDateValueToSerial(raw, dates.dateSystem, dates.odsNullDate);
      } catch {
        throw parseError(where, `"${raw}" is not a date`);
      }
    }
    case "time": {
      const raw = officeAttr("time-value") ?? "";
      try {
        return odsTimeValueToSerial(raw);
      } catch {
        throw parseError(where, `"${raw}" is not a time`);
      }
    }
    case "boolean": {
      const raw = officeAttr("boolean-value");
      if (raw === "true" || raw === "1") {
        return true;
      }
      if (raw === "false" || raw === "0") {
        return false;
      }
      throw parseError(where, `"${raw ?? ""}" is not a boolean`);
    }
    case "string":
      return officeAttr("string-value") ?? text;
    case "error":
      return toCellError(text);
    default:
      // No (or a void) value type: text a producer wrote anyway is shown as
      // a string; nothing at all is an empty cell.
      return text.length > 0 ? text : null;
  }
}

/** The format a cell's value type implies when its style gives none
 * (SPEC-118 "Implied formats"): dates `yyyy-mm-dd` (+ time), times
 * `[hh]:mm:ss`, percentages `0.00%`. */
function impliedCodeOf(cell: PendingCell, value: CellValue): string | undefined {
  if (typeof value !== "number") {
    return undefined;
  }
  switch (scanAttr(cell.attrs, "value-type", OFFICE_NAMESPACE)) {
    case "date":
      return impliedDateCode(value);
    case "time":
      return IMPLIED_TIME;
    case "percentage":
      return IMPLIED_PERCENTAGE;
    default:
      return undefined;
  }
}

/** Appends text to a cell's runs, in the innermost open span's font. */
function appendRunText(cell: PendingCell, text: string, font: CellFont | undefined): void {
  const last = cell.runs.at(-1);
  if (last && last.font === font) {
    last.text += text;
  } else {
    cell.runs.push({ text, font });
  }
}

/** Appends text to the paragraph being collected and to the runs. */
function appendText(cell: PendingCell, text: string): void {
  cell.paragraph += text;
  appendRunText(cell, text, cell.spans.at(-1)?.font);
}

/** Whether an element inside a cell's paragraph is inline content the model
 * holds (SPEC-122): text:s, text:tab, text:line-break, and a text:span whose
 * only attribute is a style holding nothing but a run font. */
function modelledInline(
  ns: string | undefined,
  local: string,
  attrs: readonly XmlScanAttribute[],
  styleBook: () => OdsStyleBook,
): boolean {
  if (ns !== TEXT_NAMESPACE) {
    return false;
  }
  if (local === "s" || local === "tab" || local === "line-break") {
    return true;
  }
  if (local !== "span") {
    return false;
  }
  if (attrs.some((attr) => !(attr.ns === TEXT_NAMESPACE && attr.local === "style-name"))) {
    return false;
  }
  const styleName = scanAttr(attrs, "style-name", TEXT_NAMESPACE);
  return styleName === undefined || styleBook().textStyleModelled(styleName);
}

/** A string cell's rich-text runs — only when a span gave a font and the
 * runs spell the value exactly. Fonts are copied: the model shares nothing
 * with the style book. */
function cellRuns(cell: PendingCell, value: CellValue): TextRun[] | undefined {
  if (!cell.hasSpanFont || typeof value !== "string") {
    return undefined;
  }
  const runs: TextRun[] = [];
  let lastFont: CellFont | undefined;
  for (const draft of cell.runs) {
    if (draft.text.length === 0) {
      continue;
    }
    const last = runs.at(-1);
    if (last && lastFont === draft.font) {
      last.text += draft.text;
      continue;
    }
    runs.push(draft.font ? { text: draft.text, font: { ...draft.font } } : { text: draft.text });
    lastFont = draft.font;
  }
  if (!runs.some((run) => run.font) || runs.map((run) => run.text).join("") !== value) {
    return undefined;
  }
  return runs;
}

interface OdsContent {
  sheets: Sheet[];
  styles: CellStyle[];
  /** SPEC-123: the common table-cell styles, in document order. */
  namedStyles: NamedCellStyle[];
  /** SPEC-129: the differential styles the tables' conditional formats
   * apply, one per named style. */
  dxfs: Dxf[];
  prefix: string;
  suffix: string;
  dates: { dateSystem: "1900" | "1904"; odsNullDate?: string };
  /** SPEC-120: what `table:calculation-settings` says (ODF's defaults when
   * it, or an attribute, is absent). */
  calc: CalcSettings;
  /** SPEC-122: `office:spreadsheet@table:structure-protected="true"`. */
  structureProtected: boolean;
  /** SPEC-130: the content validations no cell references, verbatim. */
  orphanValidations: string[];
}

/** SPEC-120 "Calculation settings" from `table:calculation-settings` and its
 * `table:iteration`, ODF's defaults filling what is absent: case-sensitive
 * true, regular expressions true, wildcards false, whole cell true. Wildcards
 * win over regular expressions (LibreOffice offers one or the other). */
function odsCalcSettings(attrs: readonly XmlScanAttribute[] | undefined, iterationStatus: string | undefined): CalcSettings {
  const attr = (local: string): string | undefined => (attrs ? scanAttr(attrs, local, TABLE_NAMESPACE) : undefined);
  const calc: CalcSettings = {};
  if (attr("case-sensitive") !== "false") {
    calc.caseSensitive = true;
  }
  if (attr("use-wildcards") !== "true") {
    calc.criteria = attr("use-regular-expressions") === "false" ? "plain" : "regex";
  }
  if (attr("search-criteria-must-apply-to-whole-cell") === "false") {
    calc.partialCellMatch = true;
  }
  if (iterationStatus === "enable") {
    calc.iterate = true;
  }
  return calc;
}

/** Reads content.xml, `sheetStyles` being styles.xml's declarations.
 * Returns null when the cell ceiling stopped the scan. */
export function readOdsContent(
  xml: string,
  budget: CellBudget,
  sheetStyles: OdsStyleSheet = emptyOdsStyleSheet(),
): OdsContent | null {
  const sheets: Sheet[] = [];
  /** SPEC-130: each finished table's validation names, as rectangles. */
  const validationRegions: Map<string, GridRange[]>[] = [];
  let dates: OdsContent["dates"] = { dateSystem: "1900" };
  let calcAttrs: readonly XmlScanAttribute[] | undefined;
  let iterationStatus: string | undefined;
  let depth = 0;
  let spreadsheetLevel = -1;
  let firstTableStart = -1;
  let previousTableEnd = -1;
  let table: PendingTable | null = null;
  /** Whether a `table:table-row` is open, and its level. */
  let inRow = false;
  let rowLevel = -1;
  let cell: PendingCell | null = null;
  /** >0 while inside a table's drawing layer (`table:shapes`, a draw:*
   * element): a text box there may hold a table of its own, whose rows are
   * not the sheet's. */
  let drawingDepth = 0;
  let tooLarge = false;
  let structureProtected = false;
  /** content.xml's style sections, sliced verbatim, and the namespace
   * declarations they need to parse on their own. */
  const styleSlices: string[] = [];
  let declarations = "";
  /** The root's prefix → namespace URI declarations, which a formula's
   * dialect prefix resolves against (SPEC-119). */
  let namespaces: ReadonlyMap<string, string> = new Map();
  let sliceLevel = -1;
  let sliceStart = -1;
  let book: OdsStyleBook | null = null;
  const styleBook = (): OdsStyleBook => (book ??= new OdsStyleBook(sheetStyles, readOdsContentStyles(styleSlices, declarations)));

  const finishCell = (done: PendingCell): boolean => {
    const reader = table!.reader;
    const where = `"${table!.name}"!${addressText(reader.rowIndex, reader.colIndex)}`;
    const formula = cellFormula(done.attrs, namespaces, { row: reader.rowIndex, col: reader.colIndex });
    const typed = scanAttr(done.attrs, "value-type", OFFICE_NAMESPACE) !== undefined || scanAttr(done.attrs, "value-type", CALCEXT_NAMESPACE) !== undefined;
    // A formula with no value type has no cached value (SPEC-119).
    const value = formula && !typed ? null : odsCellValue(done, dates, where);
    const validation = scanAttr(done.attrs, "content-validation-name", TABLE_NAMESPACE);
    if (validation !== undefined) {
      table!.validations.cell(validation, reader.colIndex, repeatCount(done.attrs, "number-columns-repeated"));
    }
    return reader.cell(done.attrs, done.covered, value, cellRuns(done, value), impliedCodeOf(done, value), formula, cellPreserved(done, xml));
  };

  const finishTable = (done: PendingTable, end: number): boolean => {
    const layout = done.reader.finish();
    if (layout === null) {
      return false;
    }
    const sheet: Sheet = {
      id: crypto.randomUUID(),
      name: done.name,
      visibility: done.hidden ? "hidden" : "visible",
      kind: "grid",
      rows: layout.rows,
      sourceXml: xml.slice(done.start, end),
      columns: layout.columns,
      defaultColWidthPx: layout.defaultColWidthPx,
      defaultRowHeightPt: layout.defaultRowHeightPt,
      merges: layout.merges,
      conditionalFormats: [],
      validations: [],
    };
    if (done.tabColor !== undefined) {
      sheet.tabColor = done.tabColor;
    }
    if (done.protected) {
      sheet.protection = { sheet: true };
    }
    sheets.push(sheet);
    validationRegions.push(done.validations.result());
    previousTableEnd = end;
    return true;
  };

  scanXml(xml, (event) => {
    if (cell !== null) {
      const inside = cell;
      if (event.type === "text") {
        if (inside.paragraph !== null && inside.skipDepth === 0) {
          appendText(inside, event.text);
        }
        return undefined;
      }
      if (event.type === "close") {
        if (depth === inside.level) {
          cell = null;
          depth -= 1;
          if (!finishCell(inside)) {
            tooLarge = true;
            return "stop";
          }
          return undefined;
        }
        if (depth === inside.level + 1 && inside.childStart >= 0) {
          captureChild(inside, xml, inside.childIsParagraph, inside.childStart, event.end);
          inside.childStart = -1;
        }
        if (inside.skipDepth > 0) {
          inside.skipDepth -= 1;
        } else if (inside.spans.length > 0 && depth === inside.spans.at(-1)!.level) {
          inside.spans.pop();
        } else if (inside.paragraph !== null && depth === inside.paragraphLevel) {
          inside.paragraphs.push(inside.paragraph);
          inside.paragraph = null;
        }
        depth -= 1;
        return undefined;
      }
      const level = depth + 1;
      if (level === inside.level + 1) {
        const paragraph = event.ns === TEXT_NAMESPACE && (event.local === "p" || event.local === "h");
        if (paragraph && (event.local === "h" || event.attrs.length > 0)) {
          inside.inline = true;
        }
        if (event.selfClosing) {
          captureChild(inside, xml, paragraph, event.start, event.end);
        } else {
          inside.childStart = event.start;
          inside.childIsParagraph = paragraph;
        }
      } else if (inside.childIsParagraph && inside.childStart >= 0 && !inside.inline && !modelledInline(event.ns, event.local, event.attrs, styleBook)) {
        inside.inline = true;
      }
      if (inside.skipDepth > 0) {
        if (!event.selfClosing) {
          inside.skipDepth += 1;
        }
      } else if ((event.ns === OFFICE_NAMESPACE && event.local === "annotation") || event.ns === DRAW_NAMESPACE) {
        if (!event.selfClosing) {
          inside.skipDepth = 1;
        }
      } else if (event.ns === TEXT_NAMESPACE && inside.paragraph === null && level === inside.level + 1) {
        if (event.local === "p" || event.local === "h") {
          if (inside.paragraphs.length > 0) {
            appendRunText(inside, "\n", undefined);
          }
          if (event.selfClosing) {
            inside.paragraphs.push("");
          } else {
            inside.paragraph = "";
            inside.paragraphLevel = level;
          }
        }
      } else if (event.ns === TEXT_NAMESPACE && inside.paragraph !== null) {
        if (event.local === "s") {
          const count = Number(scanAttr(event.attrs, "c", TEXT_NAMESPACE) ?? 1);
          appendText(inside, " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 10_000) : 1));
        } else if (event.local === "tab") {
          appendText(inside, "\t");
        } else if (event.local === "line-break") {
          appendText(inside, "\n");
        } else if (event.local === "span" && !event.selfClosing) {
          const outer = inside.spans.at(-1)?.font;
          const styleName = scanAttr(event.attrs, "style-name", TEXT_NAMESPACE);
          const own = styleName === undefined ? undefined : styleBook().textFont(styleName);
          inside.hasSpanFont ||= own !== undefined;
          inside.spans.push({ level, font: own && outer ? { ...outer, ...own } : (own ?? outer) });
        }
      }
      if (!event.selfClosing) {
        depth += 1;
      }
      return undefined;
    }

    if (event.type === "text") {
      return undefined;
    }

    if (sliceLevel > 0) {
      // Inside a style section being sliced out whole.
      if (event.type === "open") {
        depth += event.selfClosing ? 0 : 1;
      } else {
        if (depth === sliceLevel) {
          styleSlices.push(xml.slice(sliceStart, event.end));
          sliceLevel = -1;
        }
        depth -= 1;
      }
      return undefined;
    }

    if (drawingDepth > 0) {
      if (event.type === "open" && !event.selfClosing) {
        drawingDepth += 1;
        depth += 1;
      } else if (event.type === "close") {
        drawingDepth -= 1;
        depth -= 1;
      }
      return undefined;
    }
    if (
      event.type === "open" &&
      table !== null &&
      (event.ns === DRAW_NAMESPACE || (event.ns === TABLE_NAMESPACE && event.local === "shapes"))
    ) {
      if (!event.selfClosing) {
        drawingDepth = 1;
        depth += 1;
      }
      return undefined;
    }

    if (event.type === "open") {
      const level = depth + 1;
      if (level === 1) {
        declarations = namespaceDeclarations(event.attrs);
        namespaces = new Map(event.attrs.flatMap((attr) => (attr.name.startsWith("xmlns:") ? [[attr.name.slice(6), attr.value] as const] : [])));
      } else if (event.ns === OFFICE_NAMESPACE) {
        if (level === 2 && (event.local === "font-face-decls" || event.local === "automatic-styles")) {
          if (!event.selfClosing) {
            sliceLevel = level;
            sliceStart = event.start;
            depth += 1;
          }
          return undefined;
        }
        if (event.local === "spreadsheet") {
          spreadsheetLevel = level;
          structureProtected ||= scanAttr(event.attrs, "structure-protected", TABLE_NAMESPACE) === "true";
        }
      } else if (event.ns === TABLE_NAMESPACE) {
        if (event.local === "null-date" && spreadsheetLevel > 0 && table === null) {
          dates = dateSystemOf(scanAttr(event.attrs, "date-value", TABLE_NAMESPACE) ?? DEFAULT_NULL_DATE);
        } else if (event.local === "calculation-settings" && level === spreadsheetLevel + 1 && table === null) {
          calcAttrs = event.attrs;
        } else if (event.local === "iteration" && spreadsheetLevel > 0 && table === null) {
          iterationStatus = scanAttr(event.attrs, "status", TABLE_NAMESPACE);
        } else if (event.local === "table" && table === null && level === spreadsheetLevel + 1) {
          let start = event.start;
          if (previousTableEnd >= 0) {
            start = previousTableEnd;
          } else {
            while (start > 0 && /\s/.test(xml[start - 1]!)) {
              start -= 1;
            }
            firstTableStart = start;
          }
          const name = scanAttr(event.attrs, "name", TABLE_NAMESPACE) ?? `Sheet${sheets.length + 1}`;
          const tableStyle = styleBook().tableStyle(scanAttr(event.attrs, "style-name", TABLE_NAMESPACE));
          const pending: PendingTable = {
            name,
            hidden: tableStyle.hidden,
            start,
            level,
            reader: new OdsTableReader(name, styleBook(), budget),
            protected: scanAttr(event.attrs, "protected", TABLE_NAMESPACE) === "true",
            validations: new ValidationRegions(),
          };
          if (tableStyle.tabColor !== undefined) {
            pending.tabColor = tableStyle.tabColor;
          }
          if (event.selfClosing) {
            if (!finishTable(pending, event.end)) {
              tooLarge = true;
              return "stop";
            }
          } else {
            table = pending;
          }
        } else if (event.local === "table-column" && table !== null && !inRow) {
          table.reader.column(event.attrs);
        } else if (event.local === "table-row" && table !== null && !inRow) {
          table.reader.startRow(event.attrs);
          table.validations.startRow(table.reader.rowIndex, repeatCount(event.attrs, "number-rows-repeated"));
          if (event.selfClosing) {
            table.validations.endRow();
            if (!table.reader.endRow()) {
              tooLarge = true;
              return "stop";
            }
          } else {
            inRow = true;
            depth += 1;
            rowLevel = level;
            return undefined;
          }
        } else if ((event.local === "table-cell" || event.local === "covered-table-cell") && inRow) {
          const pending: PendingCell = {
            level,
            covered: event.local === "covered-table-cell",
            attrs: event.attrs,
            paragraphs: [],
            paragraph: null,
            paragraphLevel: -1,
            skipDepth: 0,
            runs: [],
            spans: [],
            hasSpanFont: false,
            preservedAttrs: unmodelledAttributes(xml, event.start, event.end, event.attrs),
            preservedXml: "",
            paragraphRanges: [],
            childStart: -1,
            childIsParagraph: false,
            inline: false,
          };
          if (event.selfClosing) {
            if (!finishCell(pending)) {
              tooLarge = true;
              return "stop";
            }
          } else {
            cell = pending;
          }
        }
      }
      if (!event.selfClosing) {
        depth += 1;
      }
      return undefined;
    }

    // A close tag.
    if (event.ns === TABLE_NAMESPACE) {
      if (event.local === "table-row" && inRow && depth === rowLevel) {
        inRow = false;
        table!.validations.endRow();
        if (!table!.reader.endRow()) {
          tooLarge = true;
          depth -= 1;
          return "stop";
        }
      } else if (event.local === "table" && table !== null && depth === table.level) {
        const done = table;
        table = null;
        if (!finishTable(done, event.end)) {
          tooLarge = true;
          depth -= 1;
          return "stop";
        }
      }
    } else if (event.ns === OFFICE_NAMESPACE && event.local === "spreadsheet") {
      spreadsheetLevel = -1;
    }
    depth -= 1;
    return undefined;
  });

  if (tooLarge) {
    return null;
  }
  if (sheets.length === 0 || firstTableStart < 0) {
    throw parseError("content.xml", "the spreadsheet has no tables");
  }
  const suffix = xml.slice(previousTableEnd);
  for (const [index, autoFilter] of readOdsAutoFilters(suffix, sheets.map((sheet) => sheet.name))) {
    sheets[index]!.autoFilter = autoFilter;
  }
  // SPEC-129: each table's conditional formats, the styles they apply
  // gathered into one dxf table.
  const conditional = { dxfs: [] as Dxf[] };
  const cfContext: OdsCfReadContext = { styles: styleBook(), namespaces };
  for (const sheet of sheets) {
    sheet.conditionalFormats = readOdsTableConditionalFormats(sheet.sourceXml!, cfContext, conditional);
  }
  // SPEC-130: the prefix's content validations over the names the cells carry.
  const prefix = xml.slice(0, firstTableStart);
  const validations = readOdsContentValidations(prefix, validationRegions);
  sheets.forEach((sheet, index) => {
    sheet.validations = validations.sheets[index]!;
  });
  return {
    sheets,
    styles: styleBook().styles,
    namedStyles: styleBook().namedStyles(),
    dxfs: conditional.dxfs,
    prefix,
    suffix,
    dates,
    calc: odsCalcSettings(calcAttrs, iterationStatus),
    structureProtected,
    orphanValidations: validations.orphans,
  };
}

/** settings.xml's `ActiveTable` (a table name), or undefined. */
export function readActiveTableName(xml: string): string | undefined {
  return readOdsSettings(xml).activeTable;
}

/** Wraps a malformed part's failure (from the scanner or the platform
 * parser) as the workbook's parse error. */
function scanned<T>(partName: string, read: () => T): T {
  try {
    return read();
  } catch (error) {
    if (error instanceof XmlScanError || error instanceof XmlParseError) {
      throw parseError(partName, `the part is not well-formed XML — ${error.message}`);
    }
    throw error;
  }
}

/** Gives each sheet the view fields settings.xml holds for its table (the
 * view-level items for a table it does not list). */
function applyViewSettings(sheets: readonly Sheet[], settings: OdsSettings): void {
  for (const sheet of sheets) {
    Object.assign(sheet, settings.tables.get(sheet.name) ?? settings.defaultView);
  }
}

/** Parses an `.ods` (or `.ots`) package. `preunzipped`, when the open funnel
 * already inflated the package, is used instead of `data`. */
export async function parseOds(
  data: Uint8Array | ArrayBuffer,
  preunzipped?: Map<string, Uint8Array>,
): Promise<ParsedSpreadsheet> {
  let parts = preunzipped;
  if (!parts) {
    try {
      parts = await unzip(data);
    } catch (error) {
      if (error instanceof ZipError) {
        throw new SpreadsheetOpenError("This file isn't an OpenDocument spreadsheet we can open");
      }
      throw error;
    }
  }
  const detected = spreadsheetFormatOf(detectPackage(parts));
  if (detected === null || detected === "xlsb" || detected.format !== "ods") {
    throw new SpreadsheetOpenError("This file isn't an OpenDocument spreadsheet we can open");
  }
  const contentBytes = parts.get("content.xml");
  if (!contentBytes) {
    throw parseError("content.xml", "the spreadsheet's content part is missing");
  }
  const stylesBytes = parts.get("styles.xml");
  const sheetStyles = stylesBytes
    ? scanned("styles.xml", () => readOdsStyleSheet(partText(stylesBytes)))
    : emptyOdsStyleSheet();
  const content = scanned("content.xml", () => readOdsContent(partText(contentBytes), new CellBudget(), sheetStyles));
  if (content === null) {
    throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
  }
  const settingsBytes = parts.get("settings.xml");
  const settings = settingsBytes ? scanned("settings.xml", () => readOdsSettings(partText(settingsBytes))) : undefined;
  if (settings) {
    applyViewSettings(content.sheets, settings);
  }
  const activeName = settings?.activeTable;
  const activeIndex = activeName === undefined ? 0 : content.sheets.findIndex((sheet) => sheet.name === activeName);
  const meta = parts.get("meta.xml");

  const parsed: ParsedSpreadsheet = {
    sheets: content.sheets,
    styles: content.styles,
    activeSheetIndex: resolveActiveSheet(Math.max(activeIndex, 0), content.sheets),
    dateSystem: content.dates.dateSystem,
    odsContentEnvelope: { prefix: content.prefix, suffix: content.suffix },
    opaque: await encodeOpaque(parts),
    properties: meta ? parseOdfMeta(partText(meta)) : {},
    dxfs: content.dxfs,
  };
  if (content.dates.odsNullDate !== undefined) {
    parsed.odsNullDate = content.dates.odsNullDate;
  }
  if (detected.packageVariant) {
    parsed.packageVariant = detected.packageVariant;
  }
  const calc: CalcSettings = { ...content.calc };
  if (settings?.autoCalculate === false) {
    calc.manual = true;
  }
  if (Object.keys(calc).length > 0) {
    parsed.calc = calc;
  }
  if (content.structureProtected) {
    parsed.structureProtected = true;
  }
  if (content.namedStyles.length > 0) {
    parsed.namedStyles = content.namedStyles;
  }
  if (content.orphanValidations.length > 0) {
    parsed.odsOrphanValidations = content.orphanValidations;
  }
  return parsed;
}
