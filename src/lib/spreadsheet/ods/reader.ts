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
//   first-reference order);
// - `table:table` elements in order (each one a grid sheet); a table whose
//   style has `table:display="false"` is hidden, its
//   `tableooo:tab-color` is the tab colour;
// - each table's columns, rows and cells (ods/tableXml.ts): column widths,
//   visibility and default cell styles; row heights, custom heights,
//   visibility and default cell styles, the defaults taken from the last
//   column and row runs; cell values and style ids; styled empty cells
//   (stored only where they differ from the effective default, a run to XFD
//   becoming the row's style); covered cells with a value or a style; spans
//   → merges. `table:number-rows-repeated` / `table:number-columns-repeated`
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
// - `table:null-date` → the date system (see serial.ts); settings.xml
//   (ods/settings.ts) → the active table and each table's frozen panes,
//   scroll position, gridlines, zero display, zoom and cursor; meta.xml →
//   properties.
import { detectPackage, spreadsheetFormatOf } from "../../core/packageDetect";
import { parseOdfMeta } from "../../core/odfMeta";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE } from "../../core/odfXml";
import { XmlParseError } from "../../core/xml";
import { scanAttr, scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { unzip, ZipError } from "../../core/zip";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { toCellError, type CellFont, type CellStyle, type CellValue, type Sheet, type TextRun } from "../model";
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
import { readOdsSettings, type OdsSettings } from "./settings";
import {
  emptyOdsStyleSheet,
  namespaceDeclarations,
  OdsStyleBook,
  readOdsContentStyles,
  readOdsStyleSheet,
  type OdsStyleSheet,
} from "./styles";
import { OdsTableReader } from "./tableXml";

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
  prefix: string;
  suffix: string;
  dates: { dateSystem: "1900" | "1904"; odsNullDate?: string };
}

/** Reads content.xml, `sheetStyles` being styles.xml's declarations.
 * Returns null when the cell ceiling stopped the scan. */
export function readOdsContent(
  xml: string,
  budget: CellBudget,
  sheetStyles: OdsStyleSheet = emptyOdsStyleSheet(),
): OdsContent | null {
  const sheets: Sheet[] = [];
  let dates: OdsContent["dates"] = { dateSystem: "1900" };
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
  /** content.xml's style sections, sliced verbatim, and the namespace
   * declarations they need to parse on their own. */
  const styleSlices: string[] = [];
  let declarations = "";
  let sliceLevel = -1;
  let sliceStart = -1;
  let book: OdsStyleBook | null = null;
  const styleBook = (): OdsStyleBook => (book ??= new OdsStyleBook(sheetStyles, readOdsContentStyles(styleSlices, declarations)));

  const finishCell = (done: PendingCell): boolean => {
    const reader = table!.reader;
    const where = `"${table!.name}"!${addressText(reader.rowIndex, reader.colIndex)}`;
    const value = odsCellValue(done, dates, where);
    return reader.cell(done.attrs, done.covered, value, cellRuns(done, value), impliedCodeOf(done, value));
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
    };
    if (done.tabColor !== undefined) {
      sheet.tabColor = done.tabColor;
    }
    sheets.push(sheet);
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
        }
      } else if (event.ns === TABLE_NAMESPACE) {
        if (event.local === "null-date" && spreadsheetLevel > 0 && table === null) {
          dates = dateSystemOf(scanAttr(event.attrs, "date-value", TABLE_NAMESPACE) ?? DEFAULT_NULL_DATE);
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
          if (event.selfClosing) {
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
  return {
    sheets,
    styles: styleBook().styles,
    prefix: xml.slice(0, firstTableStart),
    suffix: xml.slice(previousTableEnd),
    dates,
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
  };
  if (content.dates.odsNullDate !== undefined) {
    parsed.odsNullDate = content.dates.odsNullDate;
  }
  if (detected.packageVariant) {
    parsed.packageVariant = detected.packageVariant;
  }
  return parsed;
}
