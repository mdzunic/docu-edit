// Writing CSV and TSV (SPEC-126 "Writing", "Copies") — THIS FILE is the
// normative source for the CSV writer, its field text and what a CSV cannot
// carry. A csv/tsv workbook's own download is its source bytes while the
// sheet read from the file is unedited (download.ts); once edited, writeCsv
// writes `sheets[0]` — only ever that sheet — in the source's dialect, so a
// one-field edit changes exactly that field's bytes.
//
// - Field text. null → ""; a string → itself; TRUE/FALSE; an error → its
//   code; a number under General → SPEC-118's plain notation (15 significant
//   digits, the locale's decimal, no grouping, scientific only beyond 1e15 /
//   below 1e−9) rather than the 11-character General, which is a
//   fit-to-a-column rule and a file has no column; any other format →
//   plainText(formatValue(…)) with no width. A formula cell contributes its
//   value.
// - Records. Rows 0 … usedRange.bottom from column A; a row with no
//   non-empty field is an empty line. With `width`, every non-empty record
//   has max(width, usedRange.right + 1) fields; without it, each ends at its
//   last stored cell. Merge-covered cells are empty; an empty sheet writes
//   no records.
// - Quoting. A field is quoted when it holds the delimiter, `"`, CR or LF,
//   or when `quoteAll` holds and it is non-empty; `"` is doubled; an in-cell
//   LF is written as LF inside the quotes.
// - Line endings. `lineEnding` between records and after the last when
//   `finalLineEnding`; `sepLine` writes `sep=` + delimiter + `lineEnding`
//   first.
// - Encoding (encodeCsvText). UTF-8 through TextEncoder, `EF BB BF` when
//   `bom`; UTF-16 by code unit in the dialect's byte order, its BOM when
//   `bom`; Windows-1252 through a reverse table built once by decoding bytes
//   00–FF — a character outside it makes the encoder return null, and the
//   sheet is then written as UTF-8 with a BOM (`encodingFellBack`). Nothing
//   here mutates the model: the next download decides again.
import { usedRange, type GridRange } from "../address";
import { isCellError, type CellStyle, type CellValue, type CsvDialect, type CsvEncoding, type Sheet, type SpreadsheetModel } from "../model";
import { formatCodeOf } from "../numberFormat/builtin";
import { plainNotation } from "../numberFormat/decimal";
import { SPREADSHEET_LOCALES, type SpreadsheetLocale } from "../numberFormat/locale";
import { formatValue, plainText, type FormatContext } from "../numberFormat/render";
import { cellStyle, effectiveStyleId, styleKey } from "../styles";
import { CSV_ENCODING_FALLBACK_NOTICE } from "./messages";

const GENERAL = /^general$/i;

/** A cell's text in a CSV file (see the module header). `dateQuirks` is the
 * calendar a workbook's dates render in (SPEC-118; "excel" unless an
 * ods-born workbook is copied). */
export function csvFieldText(
  value: CellValue,
  numberFormat: string | undefined,
  locale: SpreadsheetLocale,
  dateSystem: "1900" | "1904",
  dateQuirks: FormatContext["dateQuirks"] = "excel",
): string {
  if (value === null) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (isCellError(value)) {
    return value.error;
  }
  if (numberFormat === undefined || GENERAL.test(numberFormat.trim())) {
    return plainNotation(value, locale.decimal);
  }
  return plainText(formatValue(value, numberFormat, { locale, dateSystem, dateQuirks }));
}

// --- losses ---------------------------------------------------------------------------------

export type CsvLoss = "sheets" | "formulas" | "merges" | "formatting";

/** A style's look as CSV would lose it: font, fill, borders and alignment. */
function lookKey(style: CellStyle): string {
  return styleKey({ font: style.font, fill: style.fill, borders: style.borders, alignment: style.alignment });
}

/** Whether anything on the sheet — a cell, a row or a column — resolves to a
 * style whose font, fill, borders or alignment are set beyond the default
 * style's (what every cell of a bare file shows; number formats are not a
 * loss — their effect is written). */
function hasFormatting(model: SpreadsheetModel, sheet: Sheet): boolean {
  const plain = lookKey(cellStyle(model, 0));
  const verdicts = new Map<number, boolean>();
  const formatted = (id: number | undefined): boolean => {
    if (id === undefined) {
      return false;
    }
    let verdict = verdicts.get(id);
    if (verdict === undefined) {
      const style = cellStyle(model, id);
      const set = style.font !== undefined || style.fill !== undefined || style.borders !== undefined || style.alignment !== undefined;
      verdict = set && lookKey(style) !== plain;
      verdicts.set(id, verdict);
    }
    return verdict;
  };
  return (
    sheet.columns.some((span) => formatted(span.styleId)) ||
    sheet.rows.some((row) => formatted(row.styleId) || row.cells.some((cell) => formatted(cell.styleId)))
  );
}

/** What a CSV of the sheet at `sheetIndex` leaves out: "sheets" when any
 * other sheet stores a cell; "formulas" when the sheet has a formula;
 * "merges" when it has merges; "formatting" when anything on it resolves to
 * a font, fill, border or alignment. Column widths, row heights, hidden rows
 * and frozen panes are view settings and are not reported. */
export function csvLosses(model: SpreadsheetModel, sheetIndex: number): CsvLoss[] {
  const losses: CsvLoss[] = [];
  if (model.sheets.some((sheet, index) => index !== sheetIndex && sheet.rows.some((row) => row.cells.length > 0))) {
    losses.push("sheets");
  }
  const sheet = model.sheets[sheetIndex];
  if (!sheet) {
    return losses;
  }
  if (sheet.rows.some((row) => row.cells.some((cell) => cell.formula !== undefined))) {
    losses.push("formulas");
  }
  if (sheet.merges.length > 0) {
    losses.push("merges");
  }
  if (hasFormatting(model, sheet)) {
    losses.push("formatting");
  }
  return losses;
}

// --- encoding --------------------------------------------------------------------------------

let windows1252: Map<number, number> | null = null;

/** Each Windows-1252 character's byte, from decoding bytes 00–FF once. */
function windows1252Table(): Map<number, number> {
  if (!windows1252) {
    const bytes = Uint8Array.from({ length: 256 }, (_, byte) => byte);
    const text = new TextDecoder("windows-1252").decode(bytes);
    windows1252 = new Map();
    for (let byte = 0; byte < text.length; byte += 1) {
      const unit = text.charCodeAt(byte);
      if (unit !== 0xfffd && !windows1252.has(unit)) {
        windows1252.set(unit, byte);
      }
    }
  }
  return windows1252;
}

/** `text` in `encoding` (with its BOM when `bom`), or null when a character
 * has no Windows-1252 byte. Never a TextEncoder label it does not support:
 * TextEncoder writes UTF-8 only. */
export function encodeCsvText(text: string, encoding: CsvEncoding, bom: boolean): Uint8Array | null {
  if (encoding === "utf-8") {
    const body = new TextEncoder().encode(text);
    if (!bom) {
      return body;
    }
    const out = new Uint8Array(body.length + 3);
    out.set([0xef, 0xbb, 0xbf]);
    out.set(body, 3);
    return out;
  }
  if (encoding === "windows-1252") {
    const table = windows1252Table();
    const out = new Uint8Array(text.length);
    for (let at = 0; at < text.length; at += 1) {
      const byte = table.get(text.charCodeAt(at));
      if (byte === undefined) {
        return null;
      }
      out[at] = byte;
    }
    return out;
  }
  const little = encoding === "utf-16le";
  const offset = bom ? 2 : 0;
  const out = new Uint8Array(offset + text.length * 2);
  if (bom) {
    out.set(little ? [0xff, 0xfe] : [0xfe, 0xff]);
  }
  for (let at = 0; at < text.length; at += 1) {
    const unit = text.charCodeAt(at);
    const high = unit >> 8;
    const low = unit & 0xff;
    out[offset + at * 2] = little ? low : high;
    out[offset + at * 2 + 1] = little ? high : low;
  }
  return out;
}

// --- records ---------------------------------------------------------------------------------

function quoteField(text: string, delimiter: string, quoteAll: boolean): string {
  if (text === "") {
    return "";
  }
  if (quoteAll || text.includes(delimiter) || text.includes('"') || text.includes("\r") || text.includes("\n")) {
    return `"${text.replaceAll('"', '""')}"`;
  }
  return text;
}

/** Whether (row, col) is covered by a merge other than as its anchor. */
function mergeCovered(merges: readonly GridRange[], row: number, col: number): boolean {
  return merges.some(
    (merge) => row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right && (row !== merge.top || col !== merge.left),
  );
}

/** The sheet's records as text (see the module header). */
function sheetRecords(model: SpreadsheetModel, sheet: Sheet | undefined, dialect: Omit<CsvDialect, "sheetId">): string[] {
  const used = sheet && sheet.kind === "grid" ? usedRange(sheet) : null;
  if (!sheet || !used) {
    return [];
  }
  const locale = SPREADSHEET_LOCALES[dialect.localeTag];
  const dateQuirks = model.format === "ods" ? "proleptic" : "excel";
  const fieldCount = dialect.width === undefined ? undefined : Math.max(dialect.width, used.right + 1);
  const records: string[] = [];
  for (const row of sheet.rows) {
    if (row.index > used.bottom) {
      break;
    }
    while (records.length < row.index) {
      records.push("");
    }
    const texts: string[] = [];
    let last = -1;
    for (const cell of row.cells) {
      if (cell.value === null || mergeCovered(sheet.merges, row.index, cell.col)) {
        continue;
      }
      const code = formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row.index, cell.col)), locale);
      const text = csvFieldText(cell.value, code, locale, model.dateSystem, dateQuirks);
      if (text !== "") {
        texts[cell.col] = text;
        last = cell.col;
      }
    }
    if (last < 0) {
      records.push("");
      continue;
    }
    const fields: string[] = [];
    for (let col = 0; col < (fieldCount ?? last + 1); col += 1) {
      fields.push(quoteField(texts[col] ?? "", dialect.delimiter, dialect.quoteAll));
    }
    records.push(fields.join(dialect.delimiter));
  }
  return records;
}

/** The sheet at `sheetIndex` as a CSV in `dialect` (see the module header);
 * `encodingFellBack` when its characters forced UTF-8 with a BOM. */
export function sheetToCsv(
  model: SpreadsheetModel,
  sheetIndex: number,
  dialect: Omit<CsvDialect, "sheetId">,
): { bytes: Uint8Array; encodingFellBack: boolean } {
  const records = sheetRecords(model, model.sheets[sheetIndex], dialect);
  let text = records.join(dialect.lineEnding);
  if (dialect.finalLineEnding && records.length > 0) {
    text += dialect.lineEnding;
  }
  if (dialect.sepLine) {
    text = `sep=${dialect.delimiter}${dialect.lineEnding}${text}`;
  }
  const bytes = encodeCsvText(text, dialect.encoding, dialect.bom);
  if (bytes) {
    return { bytes, encodingFellBack: false };
  }
  return { bytes: encodeCsvText(text, "utf-8", true)!, encodingFellBack: true };
}

/** The dialect a csv/tsv workbook without a recorded one writes in. */
function fallbackDialect(model: SpreadsheetModel): Omit<CsvDialect, "sheetId"> {
  return {
    delimiter: model.format === "tsv" ? "\t" : ",",
    encoding: "utf-8",
    bom: false,
    lineEnding: "\r\n",
    finalLineEnding: true,
    quoteAll: false,
    localeTag: "en-US",
  };
}

/** An edited csv/tsv workbook's file: `sheets[0]` in the source's dialect,
 * with CSV_ENCODING_FALLBACK_NOTICE when Windows-1252 could not hold it. */
export function writeCsv(model: SpreadsheetModel): { bytes: Uint8Array; notices: string[] } {
  const { sheetId: _sheetId, ...dialect } = model.csv ?? { ...fallbackDialect(model), sheetId: "" };
  const { bytes, encodingFellBack } = sheetToCsv(model, 0, dialect);
  return { bytes, notices: encodingFellBack ? [CSV_ENCODING_FALLBACK_NOTICE] : [] };
}
