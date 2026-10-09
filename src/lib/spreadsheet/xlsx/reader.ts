// The .xlsx reader (SPEC-116): an OPC package whose main part declares
// SpreadsheetML (plain, template, macro-enabled or macro-enabled template)
// → the workbook model. Reads the sheet list, the date system, the active
// tab and the core properties from the small parts with DOMParser
// (core/xml.ts); reads shared strings and every worksheet's cells with the
// forward-only scanner (core/xmlScanner.ts) — a million-cell sheet as a DOM
// costs an order of magnitude more memory than its text.
//
// What is read (everything else — styles, theme, drawings, comments, tables,
// calcChain, VBA, custom XML — stays in `opaque`, untouched):
// - sheets in `<sheets>` order: name, `state` → visibility, and the
//   relationship type → kind (worksheet → "grid"; chartsheet, dialogsheet,
//   macrosheet → "foreign", shown as a placeholder and carried verbatim);
// - shared strings: `<t>` or a run of `<r>`s (their `<rPr>` is SPEC-117's),
//   phonetic `rPh`/`phoneticPr` skipped, `_xHHHH_` escapes decoded;
// - cells: rows and cells without `r` placed sequentially; `t` absent/n, s,
//   str, inlineStr, b, e, d; a formula's cached `<v>` is the value (`<f>` is
//   SPEC-119's); a `<c>` with no value (a style alone) is not stored;
// - `workbookPr@date1904`, `bookViews/workbookView@activeTab`,
//   `docProps/core.xml`.
// The cell ceiling stops the scan the moment the stored count would pass
// MAX_WORKBOOK_CELLS; a cell or row past the grid is a parse error naming its
// address.
import { detectPackage, spreadsheetFormatOf } from "../../core/packageDetect";
import { parseCoreProperties } from "../../core/opcProperties";
import { parseRelationships, parseXml, relsPathFor, resolvePartPath, XmlParseError, type Relationship } from "../../core/xml";
import { scanAttr, scanXml, XmlScanError, type XmlScanEvent } from "../../core/xmlScanner";
import { unzip, ZipError } from "../../core/zip";
import type { DocumentProperties } from "../../core/documentProperties";
import { parseA1 } from "../address";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { MAX_COLS, MAX_ROWS, toCellError, type Cell, type CellValue, type Row, type Sheet, type TextRun } from "../model";
import {
  addressText,
  CellBudget,
  encodeOpaque,
  normaliseCells,
  normaliseRows,
  outOfGridError,
  parseError,
  partText,
  resolveActiveSheet,
  type ParsedSpreadsheet,
} from "../parsed";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "../geometry";
import { odsDateValueToSerial } from "../serial";
import { defaultStyles } from "../styles";

/** SpreadsheetML's main namespace, transitional and strict. */
const MAIN_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main",
]);

/** The `r:` relationship-id attribute namespace, transitional and strict. */
const RELATIONSHIP_ID_NAMESPACES = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

const CORE_PROPERTIES_RELATIONSHIP = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

/** A shared string: its text, plus its runs when it was rich. */
export interface SharedString {
  text: string;
  runs?: TextRun[];
}

/** One `<sheet>` of the workbook part, resolved through its relationship. */
interface WorkbookSheet {
  name: string;
  visibility: Sheet["visibility"];
  kind: Sheet["kind"];
  partName: string;
}

function isMain(ns: string | undefined): boolean {
  return ns !== undefined && MAIN_NAMESPACES.has(ns);
}

/** `_xHHHH_` escapes decoded to their code unit. `_x005F_` is a literal
 * underscore, so `_x005F_x0041_` reads as `_x0041_` (the match after an
 * escape resumes past it). */
export function decodeXmlEscapes(text: string): string {
  if (!text.includes("_x")) {
    return text;
  }
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}

function childElements(parent: Element, local: string): Element[] {
  return Array.from(parent.children).filter((child) => child.localName === local && isMain(child.namespaceURI ?? undefined));
}

function relationshipId(element: Element): string | null {
  for (const ns of RELATIONSHIP_ID_NAMESPACES) {
    const id = element.getAttributeNS(ns, "id");
    if (id) {
      return id;
    }
  }
  return null;
}

/** The kind a sheet's relationship type makes it: only a worksheet is a
 * grid this spec can show. */
function sheetKindOf(relationshipType: string): Sheet["kind"] {
  return relationshipType.endsWith("/worksheet") ? "grid" : "foreign";
}

function readRelationships(parts: Map<string, Uint8Array>, partPath: string): Map<string, Relationship> {
  const bytes = parts.get(relsPathFor(partPath));
  if (!bytes) {
    return new Map();
  }
  try {
    return parseRelationships(partText(bytes));
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw parseError(relsPathFor(partPath), "its relationships are not well-formed XML");
    }
    throw error;
  }
}

/** The workbook part: its sheets, date system, active tab and the shared
 * strings part it points at. */
function readWorkbook(
  parts: Map<string, Uint8Array>,
  workbookPath: string,
): { sheets: WorkbookSheet[]; date1904: boolean; activeTab: number; sharedStringsPath: string | undefined } {
  const bytes = parts.get(workbookPath);
  if (!bytes) {
    throw parseError(workbookPath, "the workbook part is missing");
  }
  let root: Element;
  try {
    root = parseXml(partText(bytes)).documentElement;
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw parseError(workbookPath, "the workbook part is not well-formed XML");
    }
    throw error;
  }
  const relationships = readRelationships(parts, workbookPath);
  const date1904Attr = childElements(root, "workbookPr")[0]?.getAttribute("date1904");
  const bookViews = childElements(root, "bookViews")[0];
  const firstView = bookViews ? childElements(bookViews, "workbookView")[0] : undefined;
  const activeTab = Number(firstView?.getAttribute("activeTab") ?? 0);

  const sheets: WorkbookSheet[] = [];
  const sheetsElement = childElements(root, "sheets")[0];
  for (const element of sheetsElement ? childElements(sheetsElement, "sheet") : []) {
    const name = element.getAttribute("name") ?? `Sheet${sheets.length + 1}`;
    const rId = relationshipId(element);
    const relationship = rId ? relationships.get(rId) : undefined;
    if (!relationship) {
      throw parseError(workbookPath, `sheet "${name}" has no part in the package`);
    }
    const state = element.getAttribute("state");
    sheets.push({
      name,
      visibility: state === "hidden" ? "hidden" : state === "veryHidden" ? "veryHidden" : "visible",
      kind: sheetKindOf(relationship.type),
      partName: resolvePartPath(workbookPath, relationship.target),
    });
  }
  if (sheets.length === 0) {
    throw parseError(workbookPath, "the workbook has no sheets");
  }
  const sharedStrings = [...relationships.values()].find((rel) => rel.type.endsWith("/sharedStrings"));
  return {
    sheets,
    date1904: date1904Attr === "1" || date1904Attr === "true",
    activeTab: Number.isFinite(activeTab) ? activeTab : 0,
    sharedStringsPath: sharedStrings ? resolvePartPath(workbookPath, sharedStrings.target) : undefined,
  };
}

/** Collects the text of a string item — `<si>` or `<is>` — from the events
 * inside it: plain `<t>` text, or `<r>` runs; phonetic runs are skipped. */
class StringItemReader {
  private plain = "";
  private runs: TextRun[] = [];
  private runText: string | null = null;
  private inText = false;
  private skipDepth = 0;

  /** Feeds one event from inside the item (not the item's own tags). */
  feed(event: XmlScanEvent): void {
    if (this.skipDepth > 0) {
      if (event.type === "open" && !event.selfClosing) {
        this.skipDepth += 1;
      } else if (event.type === "close") {
        this.skipDepth -= 1;
      }
      return;
    }
    if (event.type === "text") {
      if (this.inText) {
        if (this.runText !== null) {
          this.runText += event.text;
        } else {
          this.plain += event.text;
        }
      }
      return;
    }
    if (!isMain(event.ns)) {
      if (event.type === "open" && !event.selfClosing) {
        this.skipDepth = 1;
      }
      return;
    }
    if (event.type === "open") {
      if (event.local === "rPh" || event.local === "phoneticPr" || event.local === "rPr") {
        if (!event.selfClosing) {
          this.skipDepth = 1;
        }
      } else if (event.local === "r") {
        this.runText = event.selfClosing ? null : "";
      } else if (event.local === "t") {
        this.inText = !event.selfClosing;
      }
      return;
    }
    if (event.local === "t") {
      this.inText = false;
    } else if (event.local === "r" && this.runText !== null) {
      this.runs.push({ text: decodeXmlEscapes(this.runText) });
      this.runText = null;
    }
  }

  /** The item's string, with its runs when it had any. */
  finish(): SharedString {
    if (this.runs.length > 0) {
      return { text: this.runs.map((run) => run.text).join(""), runs: this.runs };
    }
    return { text: decodeXmlEscapes(this.plain) };
  }
}

/** Reads `xl/sharedStrings.xml` with the scanner. */
export function readSharedStrings(xml: string): SharedString[] {
  const strings: SharedString[] = [];
  let item: StringItemReader | null = null;
  let depth = 0;
  let itemDepth = -1;
  scanXml(xml, (event) => {
    if (event.type === "open") {
      if (item === null && isMain(event.ns) && event.local === "si") {
        if (event.selfClosing) {
          strings.push({ text: "" });
          return;
        }
        item = new StringItemReader();
        itemDepth = depth + 1;
      } else if (item !== null) {
        item.feed(event);
      }
      if (!event.selfClosing) {
        depth += 1;
      }
      return;
    }
    if (event.type === "close") {
      if (item !== null && depth === itemDepth) {
        strings.push(item.finish());
        item = null;
      } else if (item !== null) {
        item.feed(event);
      }
      depth -= 1;
      return;
    }
    item?.feed(event);
  });
  return strings;
}

/** One `<c>` being read. */
interface PendingCell {
  col: number;
  type: string | undefined;
  /** The `<v>` text — null when the cell had no `<v>`. */
  v: string | null;
  inlineString: StringItemReader | null;
  /** Nesting inside the cell: >0 while inside `<v>`, `<is>`, `<f>`, … */
  depth: number;
  inV: boolean;
}

/** A cell value from its type and cached text; null for "no value". */
function cellValue(
  pending: PendingCell,
  sharedStrings: readonly SharedString[],
  dateSystem: "1900" | "1904",
  where: string,
): { value: CellValue; runs?: TextRun[] } {
  const v = pending.v;
  switch (pending.type) {
    case undefined:
    case "n": {
      if (v === null || v.trim() === "") {
        return { value: null };
      }
      const n = Number(v);
      if (!Number.isFinite(n)) {
        throw parseError(where, `"${v}" is not a number`);
      }
      return { value: n };
    }
    case "s": {
      if (v === null || v.trim() === "") {
        return { value: null };
      }
      const index = Number(v);
      const shared = Number.isInteger(index) ? sharedStrings[index] : undefined;
      if (!shared) {
        throw parseError(where, `shared string ${v} does not exist`);
      }
      return shared.runs ? { value: shared.text, runs: shared.runs } : { value: shared.text };
    }
    case "str":
      return { value: v === null ? null : decodeXmlEscapes(v) };
    case "inlineStr": {
      if (pending.inlineString === null) {
        return { value: v === null ? null : decodeXmlEscapes(v) };
      }
      const inline = pending.inlineString.finish();
      return inline.runs ? { value: inline.text, runs: inline.runs } : { value: inline.text };
    }
    case "b": {
      const flag = v?.trim();
      if (flag === undefined || flag === "") {
        return { value: null };
      }
      if (flag === "1" || flag === "true") {
        return { value: true };
      }
      if (flag === "0" || flag === "false") {
        return { value: false };
      }
      throw parseError(where, `"${v}" is not a boolean`);
    }
    case "e":
      return { value: v === null || v.trim() === "" ? null : toCellError(v) };
    case "d": {
      if (v === null || v.trim() === "") {
        return { value: null };
      }
      try {
        return { value: odsDateValueToSerial(v, dateSystem) };
      } catch {
        throw parseError(where, `"${v}" is not an ISO 8601 date`);
      }
    }
    default:
      throw parseError(where, `cell type "${pending.type}" is not a SpreadsheetML type`);
  }
}

const RAW_REFERENCE = /^\$?([A-Za-z]+)\$?([0-9]+)$/;

/** The 0-based column of a `c@r`, failing loudly — out-of-grid addresses
 * name themselves, anything unreadable is malformed. */
function columnOfReference(reference: string, sheetName: string): number {
  const ref = parseA1(reference);
  if (ref) {
    return ref.col;
  }
  if (RAW_REFERENCE.test(reference)) {
    throw outOfGridError(sheetName, reference);
  }
  throw parseError(`"${sheetName}"`, `"${reference}" is not a cell reference`);
}

/** Reads one worksheet part's `sheetData` into sparse rows with the
 * scanner. Returns null when the cell ceiling stopped the scan. */
export function readWorksheetRows(
  xml: string,
  sheetName: string,
  sharedStrings: readonly SharedString[],
  dateSystem: "1900" | "1904",
  budget: CellBudget,
): Row[] | null {
  const rows: Row[] = [];
  let inSheetData = false;
  let lastRow = -1;
  let rowIndex = -1;
  let rowCells: Cell[] | null = null;
  let lastCol = -1;
  let cellsOrdered = true;
  let rowsOrdered = true;
  let pending: PendingCell | null = null;
  let tooLarge = false;

  const finishCell = (cell: PendingCell): "stop" | undefined => {
    const where = `"${sheetName}"!${addressText(rowIndex, cell.col)}`;
    const { value, runs } = cellValue(cell, sharedStrings, dateSystem, where);
    if (value === null) {
      return undefined;
    }
    if (!budget.take(1)) {
      tooLarge = true;
      return "stop";
    }
    const stored: Cell = runs ? { col: cell.col, value, runs } : { col: cell.col, value };
    if (rowCells!.length > 0 && rowCells![rowCells!.length - 1]!.col >= cell.col) {
      cellsOrdered = false;
    }
    rowCells!.push(stored);
    return undefined;
  };

  const finishRow = (): void => {
    if (rowCells && rowCells.length > 0) {
      if (!cellsOrdered) {
        normaliseCells(rowCells);
      }
      if (rows.length > 0 && rows[rows.length - 1]!.index >= rowIndex) {
        rowsOrdered = false;
      }
      rows.push({ index: rowIndex, cells: rowCells });
    }
    lastRow = rowIndex;
    rowCells = null;
  };

  scanXml(xml, (event) => {
    if (pending !== null) {
      // Inside a <c>: route everything to the pending cell.
      if (event.type === "close" && pending.depth === 0) {
        const cell = pending;
        pending = null;
        lastCol = cell.col;
        return finishCell(cell);
      }
      if (event.type === "open") {
        if (pending.depth === 0 && isMain(event.ns)) {
          if (event.local === "v") {
            pending.v = "";
            pending.inV = !event.selfClosing;
          } else if (event.local === "is") {
            pending.inlineString = new StringItemReader();
          }
        } else if (pending.inlineString !== null && pending.depth > 0) {
          pending.inlineString.feed(event);
        }
        if (!event.selfClosing) {
          pending.depth += 1;
        }
        return undefined;
      }
      if (event.type === "close") {
        pending.depth -= 1;
        if (pending.depth === 0) {
          pending.inV = false;
        } else if (pending.inlineString !== null) {
          pending.inlineString.feed(event);
        }
        return undefined;
      }
      if (pending.inV) {
        pending.v += event.text;
      } else if (pending.inlineString !== null && pending.depth > 1) {
        pending.inlineString.feed(event);
      }
      return undefined;
    }

    if (event.type === "text" || !isMain(event.ns)) {
      return undefined;
    }
    if (event.type === "open") {
      if (event.local === "sheetData") {
        inSheetData = !event.selfClosing;
      } else if (inSheetData && event.local === "row") {
        const r = scanAttr(event.attrs, "r");
        const number = r === undefined ? lastRow + 2 : Number(r);
        if (!Number.isInteger(number) || number < 1) {
          throw parseError(`"${sheetName}"`, `"${r}" is not a row number`);
        }
        if (number > MAX_ROWS) {
          throw outOfGridError(sheetName, `row ${number}`);
        }
        rowIndex = number - 1;
        rowCells = [];
        lastCol = -1;
        cellsOrdered = true;
        if (event.selfClosing) {
          finishRow();
        }
      } else if (rowCells !== null && event.local === "c") {
        const r = scanAttr(event.attrs, "r");
        const col = r === undefined ? lastCol + 1 : columnOfReference(r, sheetName);
        if (col >= MAX_COLS) {
          throw outOfGridError(sheetName, addressText(rowIndex, col));
        }
        const cell: PendingCell = { col, type: scanAttr(event.attrs, "t"), v: null, inlineString: null, depth: 0, inV: false };
        if (event.selfClosing) {
          lastCol = col;
          return finishCell(cell);
        }
        pending = cell;
      }
      return undefined;
    }
    if (event.local === "sheetData") {
      inSheetData = false;
    } else if (event.local === "row" && rowCells !== null) {
      finishRow();
    }
    return undefined;
  });

  if (tooLarge) {
    return null;
  }
  if (!rowsOrdered) {
    normaliseRows(rows);
  }
  return rows;
}

function coreProperties(parts: Map<string, Uint8Array>): DocumentProperties {
  const rootRelationships = readRelationships(parts, "");
  const core = [...rootRelationships.values()].find((rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP);
  const bytes = core ? parts.get(resolvePartPath("", core.target)) : undefined;
  return bytes ? parseCoreProperties(partText(bytes)) : {};
}

/** Wraps a scanner failure in a part as the workbook's parse error. */
function scanned<T>(partName: string, read: () => T): T {
  try {
    return read();
  } catch (error) {
    if (error instanceof XmlScanError) {
      throw parseError(partName, `the part is not well-formed XML — ${error.message}`);
    }
    throw error;
  }
}

/** Parses an `.xlsx` (or `.xlsm`/`.xltx`/`.xltm`) package. `preunzipped`,
 * when the open funnel already inflated the package, is used instead of
 * `data` (an upgrade passes the retained parts and no bytes at all). */
export async function parseXlsx(
  data: Uint8Array | ArrayBuffer,
  preunzipped?: Map<string, Uint8Array>,
): Promise<ParsedSpreadsheet> {
  let parts = preunzipped;
  if (!parts) {
    try {
      parts = await unzip(data);
    } catch (error) {
      if (error instanceof ZipError) {
        throw new SpreadsheetOpenError("This file isn't an Excel workbook we can open");
      }
      throw error;
    }
  }
  const info = detectPackage(parts);
  const detected = spreadsheetFormatOf(info);
  if (info.container !== "opc" || detected === null || detected === "xlsb" || detected.format !== "xlsx") {
    throw new SpreadsheetOpenError("This file isn't an Excel workbook we can open");
  }
  const workbook = readWorkbook(parts, info.mainPartPath);
  const dateSystem = workbook.date1904 ? "1904" : "1900";
  const sharedStringsBytes = workbook.sharedStringsPath ? parts.get(workbook.sharedStringsPath) : undefined;
  const sharedStrings = sharedStringsBytes
    ? scanned(workbook.sharedStringsPath!, () => readSharedStrings(partText(sharedStringsBytes)))
    : [];

  const budget = new CellBudget();
  const sheets: Sheet[] = [];
  for (const entry of workbook.sheets) {
    const sheet: Sheet = {
      id: crypto.randomUUID(),
      name: entry.name,
      visibility: entry.visibility,
      kind: entry.kind,
      rows: [],
      partName: entry.partName,
      columns: [],
      defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
      defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
      merges: [],
    };
    const bytes = parts.get(entry.partName);
    if (entry.kind === "grid" && bytes) {
      const rows = scanned(entry.partName, () =>
        readWorksheetRows(partText(bytes), entry.name, sharedStrings, dateSystem, budget),
      );
      if (rows === null) {
        throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
      }
      sheet.rows = rows;
    }
    sheets.push(sheet);
  }

  const parsed: ParsedSpreadsheet = {
    sheets,
    styles: defaultStyles(),
    activeSheetIndex: resolveActiveSheet(workbook.activeTab, sheets),
    dateSystem,
    opaque: await encodeOpaque(parts),
    properties: coreProperties(parts),
  };
  if (detected.packageVariant) {
    parsed.packageVariant = detected.packageVariant;
  }
  return parsed;
}
