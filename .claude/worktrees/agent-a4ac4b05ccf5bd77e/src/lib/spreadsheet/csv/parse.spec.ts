// SPEC-126 "Parsing" and "The sheet": every rule of the RFC 4180 state
// machine — quoting, `""`, delimiters and line breaks inside quotes (CRLF and
// lone CR there becoming LF), record ends outside quotes, a quote mid-field
// kept, text after a closing quote appended, an unterminated quote running to
// end of file, blank lines, no record for the final line ending — and how
// readCsvBytes builds the one-sheet workbook and refuses past its bounds.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { base64ToBytes } from "../../core/packageParts";
import { WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { MAX_COLS, MAX_ROWS, MAX_WORKBOOK_CELLS, type CsvDelimiter } from "../model";
import { SPREADSHEET_READER_VERSION } from "../open";
import { buildSpec126ExcelUtf8CsvBytes } from "../fixtures/build";
import { CSV_TOO_MANY_COLUMNS_MESSAGE, CSV_TOO_MANY_ROWS_MESSAGE, csvTooManyColumnsMessage } from "./messages";
import { CSV_SOURCE_PART, CsvOpenError, csvSheetName, csvSourceBytes, isBlankRecord, parseCsvText, readCsvBytes, reimportCsv } from "./parse";

const encoder = new TextEncoder();

/** Every record of `text`, with its quoting and starting line. */
function parse(text: string, delimiter: CsvDelimiter = ",") {
  const records: { fields: string[]; quoted: boolean[]; line: number }[] = [];
  const shape = parseCsvText(text, delimiter, (fields, quoted, line) => {
    records.push({ fields, quoted, line });
  });
  return { records, fields: records.map((record) => record.fields), shape };
}

function read(text: string, format: "csv" | "tsv" = "csv") {
  return readCsvBytes(encoder.encode(text), format, "Book", { userLocaleTag: "en-US" });
}

/** The first sheet's cells as [row, col, value]. */
function cells(model: ReturnType<typeof read>) {
  return model.sheets[0]!.rows.flatMap((row) => row.cells.map((cell) => [row.index, cell.col, cell.value]));
}

describe("parseCsvText — quoted fields", () => {
  it('reads "" inside a quoted field as one quote and reports which fields were quoted', () => {
    const { records } = parse('"a ""b"" c",d,""""');
    expect(records[0]!.fields).toEqual(['a "b" c', "d", '"']);
    expect(records[0]!.quoted).toEqual([true, false, true]);
  });

  it("keeps the delimiter, CR and LF inside quotes as data", () => {
    expect(parse('"a,b","c\nd",e').fields).toEqual([["a,b", "c\nd", "e"]]);
    expect(parse('"a;b";c', ";").fields).toEqual([["a;b", "c"]]);
    expect(parse('"a\tb"\tc', "\t").fields).toEqual([["a\tb", "c"]]);
  });

  it("turns CRLF and a lone CR inside quotes into LF, the grid's in-cell line break", () => {
    expect(parse('"a\r\nb","c\rd","e\nf"\r\n').fields).toEqual([["a\nb", "c\nd", "e\nf"]]);
  });

  it("appends whatever follows a closing quote up to the next delimiter (Excel's leniency)", () => {
    expect(parse('"ab"cd,e').fields).toEqual([["abcd", "e"]]);
    expect(parse('"ab" c"d,e').fields).toEqual([['ab c"d', "e"]]);
  });

  it("runs an unterminated quote to the end of the file without an error", () => {
    const { fields, shape } = parse('a,"bc\nd,e\r\nf');
    expect(fields).toEqual([["a", "bc\nd,e\nf"]]);
    expect(shape.records).toBe(1);
    expect(shape.finalLineEnding).toBe(false);
  });
});

describe("parseCsvText — unquoted fields and records", () => {
  it("keeps a quote mid-field literally", () => {
    const { records } = parse('ab"cd,e"f"');
    expect(records[0]!.fields).toEqual(['ab"cd', 'e"f"']);
    expect(records[0]!.quoted).toEqual([false, false]);
  });

  it("makes a field with a space before its opening quote unquoted (RFC 4180)", () => {
    expect(parse(' "b",c').fields).toEqual([[' "b"', "c"]]);
  });

  it("ends a record at CRLF, LF or a lone CR outside quotes", () => {
    expect(parse("a\r\nb\nc\rd").fields).toEqual([["a"], ["b"], ["c"], ["d"]]);
  });

  it("records the first terminator found outside quotes as the line ending", () => {
    expect(parse('"x\r\ny"\na\r\nb').shape.lineEnding).toBe("\n");
    expect(parse("a\r\nb\nc").shape.lineEnding).toBe("\r\n");
    expect(parse("a\rb").shape.lineEnding).toBe("\r");
    expect(parse("a,b").shape.lineEnding).toBe("\r\n");
  });

  it("parses blank lines as empty rows", () => {
    const { fields } = parse("a\n\nb\n\n");
    expect(fields).toEqual([["a"], [""], ["b"], [""]]);
    expect(fields.map((record) => isBlankRecord(record))).toEqual([false, true, false, true]);
  });

  it("produces no record for the line ending that ends the file", () => {
    expect(parse("a\nb\n").shape).toMatchObject({ records: 2, finalLineEnding: true });
    expect(parse("a\nb").shape).toMatchObject({ records: 2, finalLineEnding: false });
    expect(parse("\n").fields).toEqual([[""]]);
    expect(parse("").shape).toEqual({ lineEnding: "\r\n", finalLineEnding: false, records: 0 });
  });

  it("gives a trailing delimiter its empty last field", () => {
    expect(parse("a,\nb,,").fields).toEqual([
      ["a", ""],
      ["b", "", ""],
    ]);
  });

  it("sets width only when every non-blank record has the same field count", () => {
    expect(parse("a,b\n\nc,d\n").shape.width).toBe(2);
    expect(parse("a,b\nc\n").shape.width).toBeUndefined();
  });

  it("numbers each record by the 1-based line it starts on, quoted line breaks included", () => {
    expect(parse('a\n"b\r\nc\rd",e\nf').records.map((record) => record.line)).toEqual([1, 2, 5]);
  });

  it('stops where onRecord says "stop"', () => {
    const seen: string[][] = [];
    const shape = parseCsvText("a\nb\nc\n", ",", (fields) => {
      seen.push(fields);
      return seen.length === 2 ? "stop" : undefined;
    });
    expect(seen).toEqual([["a"], ["b"]]);
    expect(shape.records).toBe(2);
  });

  it("is one pass of slice-based field extraction, never per-character concatenation", () => {
    const source = readFileSync(join(dirname(fileURLToPath(import.meta.url)), "parse.ts"), "utf8");
    const body = source.slice(source.indexOf("export function parseCsvText"), source.indexOf("const SHEET_NAME_FORBIDDEN"));
    expect(body).toContain("text.slice(");
    expect(body).not.toMatch(/\+=\s*(?:text\[|text\.charAt\(|String\.fromCharCode)/);
    // A long field comes back whole.
    const long = "x".repeat(200_000);
    expect(parse(`"${long}""",${long}`).fields).toEqual([[`${long}"`, long]]);
  });
});

describe("readCsvBytes — the workbook", () => {
  it("builds one visible grid sheet, the 1900 date system, the source as the one opaque part and the dialect", () => {
    const bytes = buildSpec126ExcelUtf8CsvBytes();
    const model = readCsvBytes(bytes, "csv", "Contacts", { userLocaleTag: "en-US" });
    expect(model).toMatchObject({
      kind: "spreadsheet",
      title: "Contacts",
      format: "csv",
      readerVersion: SPREADSHEET_READER_VERSION,
      activeSheetIndex: 0,
      dateSystem: "1900",
      defaultStyleId: 0,
    });
    expect(model.sheets).toHaveLength(1);
    expect(model.sheets[0]).toMatchObject({ name: "Contacts", visibility: "visible", kind: "grid", columns: [], merges: [] });
    expect(model.sheets[0]!.dirty).toBeUndefined();
    expect(model.structureDirty).toBeUndefined();
    expect(model.opaque).toHaveLength(1);
    expect(model.opaque[0]!.name).toBe(CSV_SOURCE_PART);
    expect(equalBytes(base64ToBytes(model.opaque[0]!.data), bytes)).toBe(true);
    expect(equalBytes(csvSourceBytes(model), bytes)).toBe(true);
    expect(model.csv).toEqual({
      delimiter: ",",
      encoding: "utf-8",
      bom: true,
      lineEnding: "\r\n",
      finalLineEnding: true,
      quoteAll: false,
      width: 5,
      localeTag: "en-US",
      sheetId: model.sheets[0]!.id,
    });
    // The model is plain JSON (SPEC-116's invariant).
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
  });

  it("puts record n in row n and field m in column m; blank lines advance the row and empty fields store nothing", () => {
    const model = read("a,,b\n\n,c\n");
    expect(cells(model)).toEqual([
      [0, 0, "a"],
      [0, 2, "b"],
      [2, 1, "c"],
    ]);
  });

  it("styles Normal plus one style per distinct inferred format, shared by the cells that use it", () => {
    // 2.20 and 3.70 need "0.00" to keep their trailing zero; 2.25 would not.
    const model = read("1.50,2.20,12%\n3.70,7%,x\n");
    expect(model.styles).toEqual([{}, { numberFormat: "0.00" }, { numberFormat: "0%" }]);
    const ids = model.sheets[0]!.rows.flatMap((row) => row.cells.map((cell) => cell.styleId));
    expect(ids).toEqual([1, 1, 2, 1, 2, undefined]);
  });

  it("builds the card preview in the user's locale", () => {
    const model = readCsvBytes(encoder.encode("1.5,x\n"), "csv", "Book", { userLocaleTag: "de-DE", localeTag: "en-US" });
    expect(model.preview).toEqual({ sheetName: "Book", cells: [["1,5", "x"]] });
  });

  it("names the sheet after the title: [ ] : * ? / \\ become _, at most 31 characters, Sheet1 when nothing is left", () => {
    expect(csvSheetName("Q3 [draft]: a*b?c/d\\e")).toBe("Q3 _draft__ a_b_c_d_e");
    expect(csvSheetName("x".repeat(40))).toBe("x".repeat(31));
    expect(csvSheetName("")).toBe("Sheet1");
    expect(read("a").sheets[0]!.name).toBe("Book");
  });
});

describe("readCsvBytes — bounds", () => {
  it("refuses more than MAX_ROWS records with CSV_TOO_MANY_ROWS_MESSAGE, and opens exactly MAX_ROWS", () => {
    expect(CSV_TOO_MANY_ROWS_MESSAGE).toBe("This file has more than 1,048,576 rows — too many to open here");
    expect(() => read("\n".repeat(MAX_ROWS))).not.toThrow();
    expect(() => read("\n".repeat(MAX_ROWS) + "x")).toThrow(CSV_TOO_MANY_ROWS_MESSAGE);
    expect(() => read("\n".repeat(MAX_ROWS) + "x")).toThrow(CsvOpenError);
  });

  it("refuses a record with more than MAX_COLS fields, naming its 1-based line", () => {
    expect(CSV_TOO_MANY_COLUMNS_MESSAGE).toBe("Line N has more than 16,384 fields — too many to open here");
    expect(csvTooManyColumnsMessage(3)).toBe("Line 3 has more than 16,384 fields — too many to open here");
    const wide = ",".repeat(MAX_COLS);
    expect(() => read(`a\n"two\nlines"\n${wide}\n`)).toThrow("Line 4 has more than 16,384 fields — too many to open here");
    // The sep= line counts as a line of the file.
    expect(() => read(`sep=,\n${wide}`)).toThrow("Line 2 has more than 16,384 fields");
    expect(() => read(",".repeat(MAX_COLS - 1))).not.toThrow();
  });

  it("stops the parse as the non-empty field count passes MAX_WORKBOOK_CELLS (WORKBOOK_TOO_LARGE_MESSAGE)", () => {
    const record = Array.from({ length: MAX_COLS }, () => "a").join(",");
    const full = Math.ceil((MAX_WORKBOOK_CELLS + 1) / MAX_COLS);
    // A record past the ceiling with too many fields: reaching it would be a
    // different error, so the cell ceiling must have stopped the parse first.
    const text = `${Array.from({ length: full }, () => record).join("\n")}\n${",".repeat(MAX_COLS + 1)}`;
    expect(() => read(text)).toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
    // Empty fields are not cells: the same count of them opens.
    expect(() => read(`${",".repeat(MAX_COLS - 1)}\n`.repeat(full))).not.toThrow();
  });
});

describe("reimportCsv", () => {
  it("re-reads the retained bytes with other options, keeping the workbook and sheet ids, clean", () => {
    const model = readCsvBytes(encoder.encode("a;b,c\n1;2,5\n"), "csv", "Book", { userLocaleTag: "en-US", delimiter: "," });
    expect(model.csv?.delimiter).toBe(",");
    const sheetId = model.sheets[0]!.id;
    const next = reimportCsv(model, { userLocaleTag: "en-US", delimiter: ";", encoding: "utf-8", localeTag: "de-DE" });
    expect(next).not.toBe(model);
    expect(next.id).toBe(model.id);
    expect(next.sheets[0]!.id).toBe(sheetId);
    expect(next.sheets[0]!.dirty).toBeUndefined();
    expect(next.csv).toMatchObject({ delimiter: ";", localeTag: "de-DE", sheetId });
    expect(cells(next)).toEqual([
      [0, 0, "a"],
      [0, 1, "b,c"],
      [1, 0, 1],
      [1, 1, 2.5],
    ]);
    // Nothing given was mutated.
    expect(model.csv?.delimiter).toBe(",");
    expect(cells(model)).toEqual([
      [0, 0, "a;b"],
      [0, 1, "c"],
      [1, 0, "1;2"],
      [1, 1, 5],
    ]);
    expect(equalBytes(csvSourceBytes(next), csvSourceBytes(model))).toBe(true);
  });
});
