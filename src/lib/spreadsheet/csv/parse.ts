// Reading CSV and TSV (SPEC-126 "Parsing", "The sheet") — THIS FILE is the
// normative source for the RFC 4180 parser and for how a delimited-text file
// becomes a one-sheet workbook. The sibling modules own the rest of the
// pipeline: sniff.ts decodes the bytes and sniffs the delimiter, `sep=` line,
// quoting and locale; infer.ts decides what each field becomes; write.ts
// writes a sheet back in the file's own dialect.
//
// The parser is a single pass over the decoded string, a state machine with
// `slice`-based field extraction (never a per-character concatenation):
// - A field is quoted only if its first character is `"`. Inside it `""` is
//   one `"`; the delimiter, CR and LF are data, and CRLF or a lone CR there
//   becomes LF (the grid's in-cell line break). The closing `"` ends quoting;
//   whatever follows it up to the next delimiter or record end is appended
//   as it is (`"ab"cd` → `abcd`, Excel's leniency). End of file inside quotes
//   ends the field there — no error.
// - In an unquoted field every character but the delimiter, CR and LF is
//   data, `"` included (`ab"cd` stays `ab"cd`); a space before an opening
//   quote makes the field unquoted (` "b"` stays ` "b"`), per RFC 4180.
// - Outside quotes CRLF, LF and a lone CR each end a record. A line ending at
//   the very end of the file ends the last record and starts none. A record
//   of exactly one empty field is a blank line.
//
// readCsvBytes is synchronous (SPEC-137 owns moving it off the main thread).
// The workbook it builds holds the source bytes as its single `opaque` part
// (CSV_SOURCE_PART, BOM included) — what byte preservation, re-import and
// the reader-version re-read all read — and records the dialect in
// `model.csv`. Record n (after any `sep=` line) is row n and field m column
// m; blank lines advance the row index and store nothing, as empty fields
// do. A typed field's number format becomes the cell's style through
// internStyle, so `styles` holds Normal plus one style per distinct inferred
// format. The bounds refuse rather than truncate: more than
// MAX_WORKBOOK_CELLS non-empty fields stops the parse as the count passes
// the ceiling, more than MAX_ROWS records or a record with more than
// MAX_COLS fields is refused naming the limit.
import { bytesToBase64, decodedPartBytes } from "../../core/packageParts";
import { WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "../geometry";
import {
  MAX_COLS,
  MAX_ROWS,
  MAX_WORKBOOK_CELLS,
  type Cell,
  type CellStyle,
  type CsvDelimiter,
  type CsvDialect,
  type CsvEncoding,
  type Row,
  type Sheet,
  type SpreadsheetModel,
} from "../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { SPREADSHEET_READER_VERSION } from "../readerVersion";
import { buildPreview } from "../preview";
import { cellStyle, defaultStyles, internStyle } from "../styles";
import { inferCsvField } from "./infer";
import { CSV_TOO_MANY_ROWS_MESSAGE, csvTooManyColumnsMessage } from "./messages";
import { decodeCsvBytes, inferLocaleTag, sepLineOf, sniffDelimiter, sniffSample } from "./sniff";

export { inferCsvField } from "./infer";

/** The name of the one `opaque` part a csv/tsv workbook keeps: the source
 * bytes as they came in. */
export const CSV_SOURCE_PART = "csv-source";

/** A delimited-text file this product will not open (binary data, past a
 * bound). */
export class CsvOpenError extends Error {}

export interface CsvReadOptions {
  /** Preferences.spreadsheetLocale, passed by the app layer. */
  userLocaleTag: SpreadsheetLocaleTag;
  /** Given → that facet is not sniffed. */
  delimiter?: CsvDelimiter;
  encoding?: CsvEncoding;
  localeTag?: SpreadsheetLocaleTag;
}

/** What one full parse learned about the file's layout. */
export interface CsvShape {
  /** The first record terminator outside quotes ("\r\n" if none). */
  lineEnding: CsvDialect["lineEnding"];
  /** The last record was terminated. */
  finalLineEnding: boolean;
  /** The shared field count, when every non-blank record had the same. */
  width?: number;
  records: number;
}

const QUOTE = 0x22;
const CR = 0x0d;
const LF = 0x0a;
const LINE_BREAK = /\r\n?|\n/g;

/** Whether a record is a blank line: exactly one empty field. */
export function isBlankRecord(fields: readonly string[]): boolean {
  return fields.length === 1 && fields[0] === "";
}

/** Parses `text` record by record (see the module header), handing each
 * record's fields, which of them were quoted and the 1-based line it starts
 * on to `onRecord`; "stop" ends the parse there. */
export function parseCsvText(
  text: string,
  delimiter: CsvDelimiter,
  onRecord: (fields: string[], quoted: boolean[], line: number) => void | "stop",
): CsvShape {
  const delim = delimiter.charCodeAt(0);
  const n = text.length;
  /** The end of an unquoted run from `from`: the next delimiter, CR or LF. */
  const runEnd = (from: number): number => {
    let at = from;
    while (at < n) {
      const code = text.charCodeAt(at);
      if (code === delim || code === CR || code === LF) {
        break;
      }
      at += 1;
    }
    return at;
  };
  let i = 0;
  let line = 1;
  let records = 0;
  let lineEnding: CsvDialect["lineEnding"] | undefined;
  let finalLineEnding = false;
  /** undefined: no non-blank record yet; null: the counts differ. */
  let width: number | null | undefined;
  while (i < n) {
    const fields: string[] = [];
    const quoted: boolean[] = [];
    const startLine = line;
    let terminated = false;
    for (;;) {
      if (text.charCodeAt(i) === QUOTE) {
        let value = "";
        let chunk = i + 1;
        let at = chunk;
        for (;;) {
          const close = text.indexOf('"', at);
          if (close < 0) {
            value += text.slice(chunk);
            at = n;
            break;
          }
          if (text.charCodeAt(close + 1) === QUOTE) {
            value += text.slice(chunk, close + 1);
            at = close + 2;
            chunk = at;
            continue;
          }
          value += text.slice(chunk, close);
          at = close + 1;
          break;
        }
        if (value.includes("\r") || value.includes("\n")) {
          line += value.match(LINE_BREAK)?.length ?? 0;
          value = value.replace(/\r\n?/g, "\n");
        }
        const tail = runEnd(at);
        fields.push(tail > at ? value + text.slice(at, tail) : value);
        quoted.push(true);
        i = tail;
      } else {
        const end = runEnd(i);
        fields.push(text.slice(i, end));
        quoted.push(false);
        i = end;
      }
      if (i >= n) {
        break;
      }
      const code = text.charCodeAt(i);
      if (code === delim) {
        i += 1;
        continue;
      }
      let terminator: CsvDialect["lineEnding"];
      if (code === CR && text.charCodeAt(i + 1) === LF) {
        terminator = "\r\n";
        i += 2;
      } else {
        terminator = code === CR ? "\r" : "\n";
        i += 1;
      }
      lineEnding ??= terminator;
      line += 1;
      terminated = true;
      break;
    }
    records += 1;
    if (!isBlankRecord(fields)) {
      if (width === undefined) {
        width = fields.length;
      } else if (width !== fields.length) {
        width = null;
      }
    }
    if (terminated && i >= n) {
      finalLineEnding = true;
    }
    if (onRecord(fields, quoted, startLine) === "stop") {
      break;
    }
  }
  const shape: CsvShape = { lineEnding: lineEnding ?? "\r\n", finalLineEnding, records };
  if (typeof width === "number") {
    shape.width = width;
  }
  return shape;
}

const SHEET_NAME_FORBIDDEN = /[[\]:*?/\\]/g;
const MAX_SHEET_NAME_LENGTH = 31;

/** The one sheet's name (SPEC-126 "The sheet"): the workbook title with each
 * of `[ ] : * ? / \` replaced by `_`, truncated to 31 characters, or
 * "Sheet1" if that leaves nothing — as Excel names an opened CSV's sheet
 * after its file. */
export function csvSheetName(title: string): string {
  const name = Array.from(title.replace(SHEET_NAME_FORBIDDEN, "_")).slice(0, MAX_SHEET_NAME_LENGTH).join("");
  return name.length > 0 ? name : "Sheet1";
}

/** The source bytes a csv/tsv workbook retains (CSV_SOURCE_PART). */
export function csvSourceBytes(model: Pick<SpreadsheetModel, "opaque">): Uint8Array {
  return decodedPartBytes(model.opaque).get(CSV_SOURCE_PART) ?? new Uint8Array(0);
}

/** Reads delimited-text bytes into a new one-sheet workbook (see the module
 * header). Throws CsvOpenError for a file it will not open. */
export function readCsvBytes(bytes: Uint8Array, format: "csv" | "tsv", title: string, options: CsvReadOptions): SpreadsheetModel {
  const decoded = decodeCsvBytes(bytes, options.encoding);
  const sep = sepLineOf(decoded.text);
  const body = sep ? decoded.text.slice(sep.length) : decoded.text;
  const delimiter = options.delimiter ?? sep?.delimiter ?? sniffDelimiter(body, format);
  const sample = sniffSample(body, delimiter);
  const localeTag = options.localeTag ?? inferLocaleTag(sample.fields, options.userLocaleTag);
  const locale = SPREADSHEET_LOCALES[localeTag];

  const model: SpreadsheetModel = {
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title,
    format,
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets: [],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [{ name: CSV_SOURCE_PART, data: bytesToBase64(bytes) }],
    styles: defaultStyles(),
    defaultStyleId: 0,
    names: [],
    dxfs: [],
  };
  /** Each distinct inferred format's style id, interned once. */
  const styleIds = new Map<string, number>();
  const styleIdOf = (numberFormat: string, builtinFormatId: number | undefined): number => {
    const key = `${builtinFormatId ?? ""}|${numberFormat}`;
    let id = styleIds.get(key);
    if (id === undefined) {
      const style: CellStyle = builtinFormatId === undefined ? { numberFormat } : { numberFormat, builtinFormatId };
      id = internStyle(model, style);
      styleIds.set(key, id);
    }
    return id;
  };

  const rows: Row[] = [];
  const lineOffset = sep ? 1 : 0;
  let rowIndex = 0;
  let stored = 0;
  const shape = parseCsvText(body, delimiter, (fields, _quoted, line) => {
    if (rowIndex >= MAX_ROWS) {
      throw new CsvOpenError(CSV_TOO_MANY_ROWS_MESSAGE);
    }
    if (fields.length > MAX_COLS) {
      throw new CsvOpenError(csvTooManyColumnsMessage(line + lineOffset));
    }
    const cells: Cell[] = [];
    fields.forEach((field, col) => {
      if (field === "") {
        return;
      }
      const inferred = inferCsvField(field, locale, "1900");
      if (inferred.value === null) {
        return;
      }
      stored += 1;
      if (stored > MAX_WORKBOOK_CELLS) {
        throw new CsvOpenError(WORKBOOK_TOO_LARGE_MESSAGE);
      }
      const cell: Cell = { col, value: inferred.value };
      if (inferred.numberFormat !== undefined) {
        cell.styleId = styleIdOf(inferred.numberFormat, inferred.builtinFormatId);
      }
      cells.push(cell);
    });
    if (cells.length > 0) {
      rows.push({ index: rowIndex, cells });
    }
    rowIndex += 1;
  });

  const sheet: Sheet = {
    id: crypto.randomUUID(),
    name: csvSheetName(title),
    visibility: "visible",
    kind: "grid",
    rows,
    columns: [],
    merges: [],
    conditionalFormats: [],
    validations: [],
    defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
    defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
  };
  const dialect: CsvDialect = {
    delimiter,
    encoding: decoded.encoding,
    bom: decoded.bom,
    lineEnding: sep?.lineEnding ?? shape.lineEnding,
    finalLineEnding: shape.finalLineEnding,
    quoteAll: sample.quoteAll,
    localeTag,
    sheetId: sheet.id,
  };
  if (shape.width !== undefined) {
    dialect.width = shape.width;
  }
  if (sep) {
    dialect.sepLine = true;
  }
  model.sheets = [sheet];
  model.csv = dialect;
  model.preview = buildPreview(model, SPREADSHEET_LOCALES[options.userLocaleTag]);
  return model;
}

/** Re-reads a csv/tsv workbook's retained source bytes with other options
 * (the import dialog's "Apply"): the sheet read from the file
 * (`csv.sheetId`, else the first) keeps its id, name and view settings and
 * takes the re-read rows, clean; the re-read formats are interned into the
 * workbook's style table (any other sheet keeps its ids) and the dialect is
 * the new read's. Returns a new model object — nothing given is mutated. */
export function reimportCsv(model: SpreadsheetModel, options: CsvReadOptions): SpreadsheetModel {
  const format = model.format === "tsv" ? "tsv" : "csv";
  const fresh = readCsvBytes(csvSourceBytes(model), format, model.title, options);
  const read = fresh.sheets[0]!;
  const at = Math.max(0, model.sheets.findIndex((sheet) => sheet.id === model.csv?.sheetId));
  const previous = model.sheets[at];
  const next: SpreadsheetModel = { ...model, styles: model.styles };
  const ids = new Map<number, number>();
  const rows = read.rows.map((row) => ({
    ...row,
    cells: row.cells.map((cell) => {
      if (cell.styleId === undefined) {
        return cell;
      }
      let id = ids.get(cell.styleId);
      if (id === undefined) {
        id = internStyle(next, cellStyle(fresh, cell.styleId));
        ids.set(cell.styleId, id);
      }
      return { ...cell, styleId: id };
    }),
  }));
  let sheet: Sheet;
  if (previous) {
    const { dirty: _dirty, ...kept } = previous;
    sheet = { ...kept, rows };
  } else {
    sheet = { ...read, rows };
  }
  next.sheets = previous ? model.sheets.map((each, index) => (index === at ? sheet : each)) : [sheet];
  next.csv = { ...fresh.csv!, sheetId: sheet.id };
  next.preview = buildPreview(next, SPREADSHEET_LOCALES[options.userLocaleTag]);
  return next;
}
