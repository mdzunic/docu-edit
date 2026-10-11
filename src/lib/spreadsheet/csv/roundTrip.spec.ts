// SPEC-126 round trips through spreadsheetDownload: an untouched csv/tsv
// workbook downloads as its exact source bytes (equalBytes) for every
// fixture — also after a sheet was added, renamed or deleted elsewhere and
// after the import options changed; a one-field edit rewrites exactly that
// field's bytes; every written file re-opens to the same cells; and the
// notices — CSV_LOSS_NOTICE, the encoding fallback — appear exactly when
// something is lost. Also the copies: a csv/tsv workbook as .xlsx and .ods
// (stripRetention's from-scratch path), and any workbook's current sheet as
// CSV/TSV.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { entryFromInput, setCellEntry } from "../cells";
import { csvSourceIntact, spreadsheetDownload, spreadsheetSheetCopy } from "../download";
import {
  buildMiniBookXlsxBytes,
  buildSpec126ExcelUtf8CsvBytes,
  buildSpec126MacPipeCsvBytes,
  buildSpec126QuotedCsvBytes,
  buildSpec126Semicolon1252CsvBytes,
  buildSpec126SepCsvBytes,
  buildSpec126UnicodeTsvBytes,
  SPEC126_CSV,
} from "../fixtures/build";
import type { Row, SpreadsheetModel } from "../model";
import { parseInput } from "../numberFormat/input";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { readSpreadsheetBytes, SPREADSHEET_READER_VERSION, upgradeSpreadsheetModel } from "../open";
import { addSheet, deleteSheet, renameSheet } from "../sheets";
import { CSV_ENCODING_FALLBACK_NOTICE, CSV_LOSS_NOTICE } from "./messages";
import { readCsvBytes, reimportCsv } from "./parse";

const FIXTURES: { name: string; format: "csv" | "tsv"; build: () => Uint8Array }[] = [
  { name: SPEC126_CSV.excelUtf8.fileName, format: "csv", build: buildSpec126ExcelUtf8CsvBytes },
  { name: SPEC126_CSV.semicolon1252.fileName, format: "csv", build: buildSpec126Semicolon1252CsvBytes },
  { name: SPEC126_CSV.quoted.fileName, format: "csv", build: buildSpec126QuotedCsvBytes },
  { name: SPEC126_CSV.sep.fileName, format: "csv", build: buildSpec126SepCsvBytes },
  { name: SPEC126_CSV.unicode.fileName, format: "tsv", build: buildSpec126UnicodeTsvBytes },
  { name: SPEC126_CSV.macPipe.fileName, format: "csv", build: buildSpec126MacPipeCsvBytes },
];

function open(bytes: Uint8Array, format: "csv" | "tsv", title = "Book", userLocaleTag: SpreadsheetLocaleTag = "en-US"): SpreadsheetModel {
  return readCsvBytes(bytes, format, title, { userLocaleTag });
}

/** Commits `text` into a cell as the user typing it in `locale` would, and
 * marks the sheet edited as the undo pipeline does. */
function type(model: SpreadsheetModel, at: { row: number; col: number }, text: string, locale: SpreadsheetLocaleTag = "en-US"): void {
  setCellEntry(model, 0, at, entryFromInput(parseInput(text, SPREADSHEET_LOCALES[locale], model.dateSystem)));
  model.sheets = model.sheets.map((sheet, index) => (index === 0 ? { ...sheet, dirty: true } : sheet));
}

/** Forces the writer: the first sheet marked edited with nothing changed. */
function touched(model: SpreadsheetModel): SpreadsheetModel {
  return { ...model, sheets: model.sheets.map((sheet, index) => (index === 0 ? { ...sheet, dirty: true as const } : sheet)) };
}

/** The bytes of `source` with the field `from` replaced by `to`, encoded as
 * the fixture is. */
function replaced(text: string, from: string, to: string, encode: (text: string) => Uint8Array): Uint8Array {
  expect(text.split(from)).toHaveLength(2);
  return encode(text.replace(from, to));
}

function withBom(body: string): Uint8Array {
  return Uint8Array.from([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(body)]);
}

/** The first sheet's values, row by row, as [col, value]. */
function valuesOf(book: SpreadsheetModel): unknown {
  return book.sheets[0]!.rows.map((row) => row.cells.map((cell) => [cell.col, cell.value]));
}

/** The first sheet's cells' number formats, in order. */
function formatsOf(book: SpreadsheetModel): unknown {
  return book.sheets[0]!.rows.flatMap((row) => row.cells.map((cell) => book.styles[cell.styleId ?? 0]?.numberFormat));
}

/** A sheet's rows without view-only differences: cells as read. */
function cellsOf(rows: readonly Row[]): unknown {
  return rows.map((row) => ({ index: row.index, cells: row.cells }));
}

describe("byte preservation", () => {
  for (const fixture of FIXTURES) {
    it(`downloads an untouched ${fixture.name} as its exact source bytes`, async () => {
      const bytes = fixture.build();
      const model = open(bytes, fixture.format, "Book");
      const download = await spreadsheetDownload(model);
      expect(equalBytes(download.bytes, bytes)).toBe(true);
      expect(download.notices).toEqual([]);
      expect(download.fileName).toBe(`Book.${fixture.format}`);
      expect(download.mimeType).toBe(fixture.format === "csv" ? "text/csv" : "text/tab-separated-values");
    });
  }

  it("still downloads the source bytes after a sheet was added, renamed or deleted elsewhere", async () => {
    const bytes = buildSpec126Semicolon1252CsvBytes();
    const model = open(bytes, "csv");
    addSheet(model, 0);
    addSheet(model, 1);
    renameSheet(model, 1, "Notes");
    renameSheet(model, 0, "Daten");
    deleteSheet(model, 2);
    expect(model.structureDirty).toBe(true);
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Daten", "Notes"]);
    expect(csvSourceIntact(model)).toBe(true);
    const download = await spreadsheetDownload(model);
    expect(equalBytes(download.bytes, bytes)).toBe(true);
    // The added sheets are empty: nothing CSV cannot carry.
    expect(download.notices).toEqual([]);
  });

  it("still downloads the source bytes after the import options were changed", async () => {
    const bytes = buildSpec126Semicolon1252CsvBytes();
    let model = open(bytes, "csv");
    model = reimportCsv(model, { userLocaleTag: "en-US", delimiter: ",", encoding: "windows-1252", localeTag: "en-US" });
    expect(model.sheets[0]!.rows[0]!.cells[0]!.value).toBe("Name;Zip;Amount");
    expect(equalBytes((await spreadsheetDownload(model)).bytes, bytes)).toBe(true);
    model = reimportCsv(model, { userLocaleTag: "en-US", delimiter: ";", encoding: "utf-16le", localeTag: "de-DE" });
    expect(equalBytes((await spreadsheetDownload(model)).bytes, bytes)).toBe(true);
  });

  it("writes the file once sheets[0] is not the sheet read from the file", async () => {
    const bytes = buildSpec126SepCsvBytes();
    const model = open(bytes, "csv");
    addSheet(model, 0);
    model.sheets = [model.sheets[1]!, model.sheets[0]!];
    expect(csvSourceIntact(model)).toBe(false);
    const download = await spreadsheetDownload(model);
    // The empty new sheet is sheets[0]: just the sep= line; the source's
    // content now sits on another sheet.
    expect(new TextDecoder().decode(download.bytes)).toBe("sep=;\r\n");
    expect(download.notices).toEqual([CSV_LOSS_NOTICE]);
  });
});

describe("one-field edits", () => {
  it("rewrites exactly the edited field of spec126-excel-utf8.csv", async () => {
    const { text, edit } = SPEC126_CSV.excelUtf8;
    const model = open(buildSpec126ExcelUtf8CsvBytes(), "csv");
    type(model, edit, edit.to);
    const download = await spreadsheetDownload(model);
    expect(equalBytes(download.bytes, replaced(text, `\r\n${edit.from},`, `\r\n${edit.to},`, withBom))).toBe(true);
    expect(download.notices).toEqual([]);
  });

  it("rewrites exactly the edited field of spec126-semicolon-1252.csv, still Windows-1252", async () => {
    const { text, edit } = SPEC126_CSV.semicolon1252;
    const model = open(buildSpec126Semicolon1252CsvBytes(), "csv");
    type(model, edit, edit.to);
    const download = await spreadsheetDownload(model);
    const expected = new TextEncoder().encode(text.replace(edit.from, edit.to));
    expect(new TextDecoder("windows-1252").decode(download.bytes)).toBe(new TextDecoder().decode(expected));
    expect(download.bytes.length).toBe(buildSpec126Semicolon1252CsvBytes().length - edit.from.length + edit.to.length);
    expect(download.notices).toEqual([]);
  });

  it("rewrites exactly the edited field of spec126-unicode.tsv, still UTF-16 LE with its BOM", async () => {
    const { text, edit } = SPEC126_CSV.unicode;
    const model = open(buildSpec126UnicodeTsvBytes(), "tsv");
    type(model, edit, edit.to);
    const download = await spreadsheetDownload(model);
    expect([...download.bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
    expect(new TextDecoder("utf-16le").decode(download.bytes.subarray(2))).toBe(text.replace(edit.from, edit.to));
  });

  it("falls back to UTF-8 with a BOM and says so when a Windows-1252 file gains a character it cannot hold", async () => {
    const model = open(buildSpec126Semicolon1252CsvBytes(), "csv");
    type(model, { row: 1, col: 0 }, "Müller Ω");
    const download = await spreadsheetDownload(model);
    expect([...download.bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder().decode(download.bytes.subarray(3))).toContain("Müller Ω;01234;1234,5\r\n");
    expect(download.notices).toEqual([CSV_ENCODING_FALLBACK_NOTICE]);
    expect(model.csv?.encoding).toBe("windows-1252");
  });
});

describe("re-opening what was written", () => {
  for (const fixture of FIXTURES) {
    it(`re-reads a written ${fixture.name} to the same cells`, async () => {
      const model = open(fixture.build(), fixture.format);
      const written = (await spreadsheetDownload(touched(model))).bytes;
      const again = readCsvBytes(written, fixture.format, "Book", {
        userLocaleTag: "en-US",
        delimiter: model.csv!.delimiter,
        encoding: model.csv!.encoding,
        localeTag: model.csv!.localeTag,
      });
      expect(cellsOf(again.sheets[0]!.rows)).toEqual(cellsOf(model.sheets[0]!.rows));
      expect(again.styles).toEqual(model.styles);
    });
  }

  it("re-reads an edited file to the edited cells", async () => {
    const model = open(buildSpec126ExcelUtf8CsvBytes(), "csv");
    type(model, { row: 5, col: 1 }, "1.25");
    type(model, { row: 0, col: 6 }, "Extra, with comma");
    const written = (await spreadsheetDownload(model)).bytes;
    const again = open(written, "csv");
    expect(cellsOf(again.sheets[0]!.rows)).toEqual(cellsOf(model.sheets[0]!.rows));
  });
});

describe("losses", () => {
  it("still downloads but raises CSV_LOSS_NOTICE once for a formula, another sheet with content, a merge or formatting", async () => {
    expect(CSV_LOSS_NOTICE).toBe("Formatting, formulas and other sheets aren't saved in CSV — download a copy as .xlsx to keep them");
    const model = open(buildSpec126Semicolon1252CsvBytes(), "csv");
    type(model, { row: 1, col: 3 }, "=1+1");
    model.sheets = model.sheets.map((sheet, index) =>
      index === 0 ? { ...sheet, merges: [{ top: 6, left: 0, bottom: 6, right: 1 }] } : sheet,
    );
    addSheet(model, 0);
    model.sheets[1] = { ...model.sheets[1]!, rows: [{ index: 0, cells: [{ col: 0, value: "kept only in .xlsx" }] }] };
    const download = await spreadsheetDownload(model);
    expect(download.notices).toEqual([CSV_LOSS_NOTICE]);
    expect(download.bytes.length).toBeGreaterThan(0);
  });

  it("raises no toast for a download with nothing lost", async () => {
    const model = open(buildSpec126Semicolon1252CsvBytes(), "csv");
    type(model, { row: 1, col: 2 }, "99");
    expect((await spreadsheetDownload(model)).notices).toEqual([]);
  });
});

describe("copies", () => {
  for (const target of ["xlsx", "ods"] as const) {
    it(`copies a csv workbook as .${target} through stripRetention, re-opening with the same values and inferred formats`, async () => {
      const model = open(buildSpec126ExcelUtf8CsvBytes(), "csv", "Contacts");
      const copy = await spreadsheetDownload(model, target);
      expect(copy.fileName).toBe(`Contacts.${target}`);
      expect(copy.notices).toEqual([]);
      const reopened = await readSpreadsheetBytes(copy.bytes, target, "Contacts");
      expect(reopened.csv).toBeUndefined();
      expect(valuesOf(reopened)).toEqual(valuesOf(model));
      expect(formatsOf(reopened)).toEqual(formatsOf(model));
      // The open model is never mutated.
      expect(model.format).toBe("csv");
      expect(model.csv).toBeDefined();
    });
  }

  it("copies any workbook's active sheet as CSV UTF-8 with a BOM, CRLF and the user's list separator", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    const before = JSON.stringify(model);
    const copy = spreadsheetSheetCopy(model, "csv", SPREADSHEET_LOCALES["en-US"]);
    expect([...copy.bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    const text = new TextDecoder().decode(copy.bytes.subarray(3));
    expect(text.endsWith("\r\n")).toBe(true);
    expect(text.split("\r\n").some((record) => record.split(",").includes("Revenue"))).toBe(true);
    expect(copy.fileName).toBe(`mini-book - ${model.sheets[model.activeSheetIndex]!.name}.csv`);
    expect(copy.mimeType).toBe("text/csv");
    expect(JSON.stringify(model)).toBe(before);

    const german = spreadsheetSheetCopy(model, "csv", SPREADSHEET_LOCALES["de-DE"]);
    expect(new TextDecoder().decode(german.bytes).split("\r\n")[1]).toContain(";");
    const tsv = spreadsheetSheetCopy(model, "tsv", SPREADSHEET_LOCALES["de-DE"]);
    expect(tsv.fileName.endsWith(".tsv")).toBe(true);
    expect(new TextDecoder().decode(tsv.bytes).split("\r\n")[1]).toContain("\t");
  });

  it("names a one-sheet workbook's copy after its title, and raises the losses toast without other sheets", async () => {
    const model = open(buildSpec126SepCsvBytes(), "csv", "Fruit");
    const plain = await spreadsheetDownload(model, "tsv");
    expect(plain.fileName).toBe("Fruit.tsv");
    expect(new TextDecoder().decode(plain.bytes.subarray(3))).toBe("Item\tQty\tPrice\r\nApples\t3\t0.5\r\nPears\t12\t1.25\r\n");
    expect(plain.notices).toEqual([]);
    addSheet(model, 0);
    model.sheets[1] = { ...model.sheets[1]!, rows: [{ index: 0, cells: [{ col: 0, value: 1 }] }] };
    model.activeSheetIndex = 0;
    // Another sheet's content is not this copy's loss.
    expect(spreadsheetSheetCopy(model, "csv").notices).toEqual([]);
    type(model, { row: 0, col: 4 }, "=1+1");
    expect(spreadsheetSheetCopy(model, "csv").notices).toEqual([CSV_LOSS_NOTICE]);
  });
});

describe("reader versions", () => {
  it("re-reads an unedited csv workbook behind the version with its recorded dialect, keeping id, title and sheet id", async () => {
    const model = open(buildSpec126Semicolon1252CsvBytes(), "csv", "Kunden");
    const sheetId = model.sheets[0]!.id;
    const stale: SpreadsheetModel = { ...model, readerVersion: SPREADSHEET_READER_VERSION - 1, sheets: [{ ...model.sheets[0]!, rows: [] }] };
    const upgraded = await upgradeSpreadsheetModel(stale);
    expect(upgraded).not.toBe(stale);
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.id).toBe(model.id);
    expect(upgraded.title).toBe("Kunden");
    expect(upgraded.sheets[0]!.id).toBe(sheetId);
    expect(upgraded.csv).toEqual(model.csv);
    expect(cellsOf(upgraded.sheets[0]!.rows)).toEqual(cellsOf(model.sheets[0]!.rows));
  });

  it("only re-stamps an edited one, and never runs a ReaderUpgrade on it", async () => {
    const model = open(buildSpec126Semicolon1252CsvBytes(), "csv");
    type(model, { row: 1, col: 2 }, "99");
    const stale: SpreadsheetModel = { ...model, readerVersion: 5 };
    const upgraded = await upgradeSpreadsheetModel(stale);
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.sheets).toBe(stale.sheets);
    expect(upgraded.csv).toBe(stale.csv);
  });

  it("returns a current csv workbook as the same object", async () => {
    const model = open(buildSpec126SepCsvBytes(), "csv");
    expect(await upgradeSpreadsheetModel(model)).toBe(model);
  });
});
