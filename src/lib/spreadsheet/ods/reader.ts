// The .ods reader (SPEC-116): an ODF package whose mimetype declares a
// spreadsheet (or its template) → the workbook model. content.xml is read
// with the forward-only scanner (core/xmlScanner.ts), whose exact offsets
// are also how each table's verbatim XML is sliced out: `sourceXml` is the
// whitespace before a `<table:table>` plus the element itself, and the
// model's `odsContentEnvelope` holds everything before the first and after
// the last — so prefix + every sourceXml + suffix is content.xml, exactly
// (the splice base SPEC-121/122's writers build on).
//
// What is read:
// - `table:table` elements in order (each one a grid sheet); a table whose
//   automatic style has `table:display="false"` is hidden;
// - rows — including those nested in header rows, row groups and
//   `table:table-rows` — in document order; `table:number-rows-repeated` /
//   `table:number-columns-repeated` honoured without ever materialising a
//   repeated EMPTY row or cell (LibreOffice's trailing 1048553-row and
//   16380-column runs cost nothing); a repeated non-empty row or cell is
//   expanded, every copy counted against MAX_WORKBOOK_CELLS before it is
//   built; `table:covered-table-cell` occupies its column and stores nothing;
// - values by `office:value-type` (float/percentage/currency → number,
//   date/time → serial, boolean, string → `office:string-value` or the
//   `text:p` text with paragraphs joined by "\n" and `text:s`/`text:tab`/
//   `text:line-break` expanded); `calcext:value-type="error"` → the error code
//   from the text; a formula cell contributes its cached value;
// - `table:null-date` → the date system (see serial.ts); settings.xml's
//   `ActiveTable` → the active sheet; meta.xml → properties.
import { detectPackage, spreadsheetFormatOf } from "../../core/packageDetect";
import { parseOdfMeta } from "../../core/odfMeta";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE } from "../../core/odfXml";
import { scanAttr, scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { unzip, ZipError } from "../../core/zip";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { MAX_COLS, MAX_ROWS, toCellError, type Cell, type CellValue, type Row, type Sheet } from "../model";
import {
  addressText,
  CellBudget,
  encodeOpaque,
  outOfGridError,
  parseError,
  partText,
  resolveActiveSheet,
  type ParsedSpreadsheet,
} from "../parsed";
import { odsDateValueToSerial, odsTimeValueToSerial } from "../serial";

/** LibreOffice's extension namespace (`calcext:value-type="error"`). */
const CALCEXT_NAMESPACE = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const CONFIG_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

/** The null date ODF assumes when a document declares none. */
const DEFAULT_NULL_DATE = "1899-12-30";

/** A repeat count attribute (`table:number-rows-repeated`, …): 1 when absent
 * or unreadable. */
function repeatCount(attrs: readonly XmlScanAttribute[], local: string): number {
  const value = Number(scanAttr(attrs, local, TABLE_NAMESPACE));
  return Number.isInteger(value) && value > 1 ? value : 1;
}

interface PendingTable {
  name: string;
  hidden: boolean;
  /** Where this sheet's sourceXml starts in content.xml. */
  start: number;
  level: number;
  rows: Row[];
  /** The 0-based index the next row lands on. */
  rowIndex: number;
}

interface PendingRow {
  level: number;
  repeat: number;
  cells: Cell[];
  /** The 0-based column the next cell lands on. */
  colIndex: number;
}

interface PendingCell {
  level: number;
  covered: boolean;
  repeat: number;
  attrs: readonly XmlScanAttribute[];
  paragraphs: string[];
  /** The text:p being collected, or null. */
  paragraph: string | null;
  paragraphLevel: number;
  /** >0 while inside a subtree whose text is not the cell's (an annotation,
   * a drawing). */
  skipDepth: number;
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

interface OdsContent {
  sheets: Sheet[];
  prefix: string;
  suffix: string;
  dates: { dateSystem: "1900" | "1904"; odsNullDate?: string };
}

/** Reads content.xml. Returns null when the cell ceiling stopped the scan. */
export function readOdsContent(xml: string, budget: CellBudget): OdsContent | null {
  const sheets: Sheet[] = [];
  const hiddenStyles = new Set<string>();
  let dates: OdsContent["dates"] = { dateSystem: "1900" };
  let depth = 0;
  let automaticStylesLevel = -1;
  let tableStyle: string | null = null;
  let spreadsheetLevel = -1;
  let firstTableStart = -1;
  let previousTableEnd = -1;
  let table: PendingTable | null = null;
  let row: PendingRow | null = null;
  let cell: PendingCell | null = null;
  let tooLarge = false;

  const finishCell = (done: PendingCell): boolean => {
    const current = row!;
    if (done.covered) {
      current.colIndex += done.repeat;
      return true;
    }
    const where = `"${table!.name}"!${addressText(table!.rowIndex, current.colIndex)}`;
    const value = odsCellValue(done, dates, where);
    if (value === null) {
      current.colIndex += done.repeat;
      return true;
    }
    if (table!.rowIndex >= MAX_ROWS) {
      throw outOfGridError(table!.name, addressText(table!.rowIndex, current.colIndex));
    }
    if (current.colIndex + done.repeat > MAX_COLS) {
      throw outOfGridError(table!.name, addressText(table!.rowIndex, Math.max(current.colIndex, MAX_COLS)));
    }
    if (!budget.take(done.repeat)) {
      return false;
    }
    for (let k = 0; k < done.repeat; k += 1) {
      current.cells.push({ col: current.colIndex + k, value });
    }
    current.colIndex += done.repeat;
    return true;
  };

  const finishRow = (done: PendingRow): boolean => {
    const owner = table!;
    if (done.cells.length === 0) {
      owner.rowIndex += done.repeat;
      return true;
    }
    if (owner.rowIndex + done.repeat > MAX_ROWS) {
      throw outOfGridError(owner.name, addressText(MAX_ROWS, done.cells[0]!.col));
    }
    if (done.repeat > 1 && !budget.take((done.repeat - 1) * done.cells.length)) {
      return false;
    }
    owner.rows.push({ index: owner.rowIndex, cells: done.cells });
    for (let k = 1; k < done.repeat; k += 1) {
      owner.rows.push({ index: owner.rowIndex + k, cells: done.cells.map((stored) => ({ ...stored })) });
    }
    owner.rowIndex += done.repeat;
    return true;
  };

  const finishTable = (done: PendingTable, end: number): void => {
    sheets.push({
      id: crypto.randomUUID(),
      name: done.name,
      visibility: done.hidden ? "hidden" : "visible",
      kind: "grid",
      rows: done.rows,
      sourceXml: xml.slice(done.start, end),
    });
    previousTableEnd = end;
  };

  scanXml(xml, (event) => {
    if (cell !== null) {
      const inside = cell;
      if (event.type === "text") {
        if (inside.paragraph !== null && inside.skipDepth === 0) {
          inside.paragraph += event.text;
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
          inside.paragraph += " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 10_000) : 1);
        } else if (event.local === "tab") {
          inside.paragraph += "\t";
        } else if (event.local === "line-break") {
          inside.paragraph += "\n";
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

    if (event.type === "open") {
      const level = depth + 1;
      if (event.ns === OFFICE_NAMESPACE) {
        if (event.local === "automatic-styles") {
          automaticStylesLevel = level;
        } else if (event.local === "spreadsheet") {
          spreadsheetLevel = level;
        }
      } else if (event.ns === STYLE_NAMESPACE && automaticStylesLevel > 0) {
        if (event.local === "style" && scanAttr(event.attrs, "family", STYLE_NAMESPACE) === "table") {
          tableStyle = scanAttr(event.attrs, "name", STYLE_NAMESPACE) ?? null;
        } else if (event.local === "table-properties" && tableStyle !== null) {
          if (scanAttr(event.attrs, "display", TABLE_NAMESPACE) === "false") {
            hiddenStyles.add(tableStyle);
          }
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
          const styleName = scanAttr(event.attrs, "style-name", TABLE_NAMESPACE);
          const pending: PendingTable = {
            name: scanAttr(event.attrs, "name", TABLE_NAMESPACE) ?? `Sheet${sheets.length + 1}`,
            hidden: styleName !== undefined && hiddenStyles.has(styleName),
            start,
            level,
            rows: [],
            rowIndex: 0,
          };
          if (event.selfClosing) {
            finishTable(pending, event.end);
          } else {
            table = pending;
          }
        } else if (event.local === "table-row" && table !== null && row === null) {
          const pending: PendingRow = { level, repeat: repeatCount(event.attrs, "number-rows-repeated"), cells: [], colIndex: 0 };
          if (event.selfClosing) {
            finishRow(pending);
          } else {
            row = pending;
          }
        } else if ((event.local === "table-cell" || event.local === "covered-table-cell") && row !== null) {
          const pending: PendingCell = {
            level,
            covered: event.local === "covered-table-cell",
            repeat: repeatCount(event.attrs, "number-columns-repeated"),
            attrs: event.attrs,
            paragraphs: [],
            paragraph: null,
            paragraphLevel: -1,
            skipDepth: 0,
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
      if (event.local === "table-row" && row !== null && depth === row.level) {
        const done = row;
        row = null;
        if (!finishRow(done)) {
          tooLarge = true;
          depth -= 1;
          return "stop";
        }
      } else if (event.local === "table" && table !== null && depth === table.level) {
        finishTable(table, event.end);
        table = null;
      }
    } else if (event.ns === STYLE_NAMESPACE && event.local === "style") {
      tableStyle = null;
    } else if (event.ns === OFFICE_NAMESPACE) {
      if (event.local === "automatic-styles") {
        automaticStylesLevel = -1;
      } else if (event.local === "spreadsheet") {
        spreadsheetLevel = -1;
      }
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
    prefix: xml.slice(0, firstTableStart),
    suffix: xml.slice(previousTableEnd),
    dates,
  };
}

/** settings.xml's `ActiveTable` (a table name), or undefined. */
export function readActiveTableName(xml: string): string | undefined {
  let collecting = false;
  let text = "";
  let found: string | undefined;
  scanXml(xml, (event) => {
    if (event.type === "open" && event.ns === CONFIG_NAMESPACE && event.local === "config-item") {
      if (scanAttr(event.attrs, "name", CONFIG_NAMESPACE) === "ActiveTable") {
        if (event.selfClosing) {
          found = "";
          return "stop";
        }
        collecting = true;
        text = "";
      }
      return undefined;
    }
    if (event.type === "text" && collecting) {
      text += event.text;
      return undefined;
    }
    if (event.type === "close" && collecting) {
      found = text;
      return "stop";
    }
    return undefined;
  });
  return found;
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
  const content = scanned("content.xml", () => readOdsContent(partText(contentBytes), new CellBudget()));
  if (content === null) {
    throw new SpreadsheetOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
  }
  const settings = parts.get("settings.xml");
  const activeName = settings ? scanned("settings.xml", () => readActiveTableName(partText(settings))) : undefined;
  const activeIndex = activeName === undefined ? 0 : content.sheets.findIndex((sheet) => sheet.name === activeName);
  const meta = parts.get("meta.xml");

  const parsed: ParsedSpreadsheet = {
    sheets: content.sheets,
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
