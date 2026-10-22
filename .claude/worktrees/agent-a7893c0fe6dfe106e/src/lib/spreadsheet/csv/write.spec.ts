// SPEC-126 "Writing": csvFieldText per type (General as the 15-digit plain
// notation, every other format "as shown" at unlimited width), quoting,
// width padding, line endings and the sep= line, every encoding including
// the Windows-1252 fallback, and csvLosses.
import { describe, expect, it } from "vitest";
import { setCellEntry } from "../cells";
import { DEFAULT_COL_WIDTH_PX } from "../geometry";
import type { CsvDialect, Sheet, SpreadsheetModel } from "../model";
import { generalNumberText } from "../numberFormat/decimal";
import { SPREADSHEET_LOCALES } from "../numberFormat/locale";
import { emptySheet } from "../sheets";
import { internStyle } from "../styles";
import { CSV_ENCODING_FALLBACK_NOTICE } from "./messages";
import { readCsvBytes } from "./parse";
import { csvFieldText, csvLosses, encodeCsvText, sheetToCsv, writeCsv } from "./write";

const enUS = SPREADSHEET_LOCALES["en-US"];
const deDE = SPREADSHEET_LOCALES["de-DE"];
const encoder = new TextEncoder();
const decoder = new TextDecoder();

const DIALECT: Omit<CsvDialect, "sheetId"> = {
  delimiter: ",",
  encoding: "utf-8",
  bom: false,
  lineEnding: "\n",
  finalLineEnding: true,
  quoteAll: false,
  localeTag: "en-US",
};

function read(text: string): SpreadsheetModel {
  return readCsvBytes(encoder.encode(text), "csv", "Book", { userLocaleTag: "en-US", delimiter: ",", localeTag: "en-US" });
}

/** The sheet's CSV text in `dialect` (UTF-8 without a BOM). */
function csvText(model: SpreadsheetModel, dialect: Partial<Omit<CsvDialect, "sheetId">> = {}, sheetIndex = 0): string {
  return decoder.decode(sheetToCsv(model, sheetIndex, { ...DIALECT, ...dialect }).bytes);
}

/** Marks the first sheet edited, as the undo pipeline does. */
function dirty(model: SpreadsheetModel): void {
  model.sheets = model.sheets.map((sheet, index) => (index === 0 ? { ...sheet, dirty: true } : sheet));
}

describe("csvFieldText", () => {
  it("writes null as nothing, a string as itself, booleans as TRUE/FALSE and an error as its code", () => {
    expect(csvFieldText(null, undefined, enUS, "1900")).toBe("");
    expect(csvFieldText("=1+1", undefined, enUS, "1900")).toBe("=1+1");
    expect(csvFieldText(" a ", "0.00", enUS, "1900")).toBe(" a ");
    expect(csvFieldText(true, undefined, enUS, "1900")).toBe("TRUE");
    expect(csvFieldText(false, "0.00", enUS, "1900")).toBe("FALSE");
    expect(csvFieldText({ error: "#DIV/0!" }, undefined, enUS, "1900")).toBe("#DIV/0!");
  });

  it("writes a General number in plain notation — 15 significant digits, the locale's decimal, no grouping — not the 11-character General", () => {
    expect(csvFieldText(123456789012, undefined, enUS, "1900")).toBe("123456789012");
    expect(generalNumberText(123456789012)).toBe("1.23457E+11");
    expect(csvFieldText(1234567.891, undefined, deDE, "1900")).toBe("1234567,891");
    expect(csvFieldText(0.1 + 0.2, undefined, enUS, "1900")).toBe("0.3");
    expect(csvFieldText(2 / 3, "General", enUS, "1900")).toBe("0.666666666666667");
    expect(csvFieldText(1e15, undefined, enUS, "1900")).toBe("1E+15");
    expect(csvFieldText(123456789012345, undefined, enUS, "1900")).toBe("123456789012345");
    expect(csvFieldText(1e-10, undefined, enUS, "1900")).toBe("1E-10");
    expect(csvFieldText(-7.25, undefined, enUS, "1900")).toBe("-7.25");
  });

  it("writes any other format as shown at unlimited width, never ####", () => {
    expect(csvFieldText(1.5, "0.00", enUS, "1900")).toBe("1.50");
    expect(csvFieldText(1234.56, "#,##0.00", deDE, "1900")).toBe("1.234,56");
    expect(csvFieldText(0.12, "0%", enUS, "1900")).toBe("12%");
    expect(csvFieldText(46303, "yyyy-mm-dd", enUS, "1900")).toBe("2026-10-08");
    expect(csvFieldText(46303 - 1462, "yyyy-mm-dd", enUS, "1904")).toBe("2026-10-08");
    expect(csvFieldText(46303, "dd.mm.yyyy", deDE, "1900")).toBe("08.10.2026");
    expect(csvFieldText(123456789012345, "#,##0", enUS, "1900")).toBe("123,456,789,012,345");
    // A fill character is written once (plainText).
    expect(csvFieldText(5, "*-0", enUS, "1900")).toBe("-5");
  });
});

describe("records", () => {
  it("quotes a field holding the delimiter, a quote, CR or LF, doubling the quote; an in-cell LF stays LF", () => {
    const model = read('"a,b","say ""hi""","x\ny",plain\n');
    expect(csvText(model)).toBe('"a,b","say ""hi""","x\ny",plain\n');
    expect(csvText(model, { delimiter: ";" })).toBe('a,b;"say ""hi""";"x\ny";plain\n');
  });

  it("quotes every non-empty field when quoteAll holds, never an empty one", () => {
    expect(csvText(read("a,,1\n"), { quoteAll: true })).toBe('"a",,"1"\n');
  });

  it("pads every non-empty record to max(width, usedRange.right + 1) fields; a row with no cell is an empty line", () => {
    const model = read("a,b,c\nd\n\n,,f\n");
    expect(csvText(model, { width: 3 })).toBe("a,b,c\nd,,\n\n,,f\n");
    expect(csvText(model, { width: 2 })).toBe("a,b,c\nd,,\n\n,,f\n");
    expect(csvText(model, { width: 5 })).toBe("a,b,c,,\nd,,,,\n\n,,f,,\n");
  });

  it("ends each record at its last stored cell without width, from column A", () => {
    expect(csvText(read(",b\nc,,\n"))).toBe(",b\nc\n");
  });

  it("writes rows 0 … usedRange.bottom: leading empty rows as empty lines, nothing after the last stored row", () => {
    const model = read("\n\na\n\n\n");
    expect(csvText(model)).toBe("\n\na\n");
  });

  it("leaves merge-covered cells empty", () => {
    const model = read("a,b,c\nd,e,f\n");
    model.sheets = [{ ...model.sheets[0]!, merges: [{ top: 0, left: 0, bottom: 1, right: 1 }] }];
    expect(csvText(model, { width: 3 })).toBe("a,,c\n,,f\n");
  });

  it("writes an empty sheet as no records (a sep= line alone when the source had one)", () => {
    const model = read("");
    expect(csvText(model)).toBe("");
    expect(csvText(model, { sepLine: true, delimiter: ";", lineEnding: "\r\n" })).toBe("sep=;\r\n");
  });

  it("puts lineEnding between records, after the last only with finalLineEnding, and the sep= line first", () => {
    const model = read("a\nb\n");
    expect(csvText(model, { lineEnding: "\r\n" })).toBe("a\r\nb\r\n");
    expect(csvText(model, { lineEnding: "\r", finalLineEnding: false })).toBe("a\rb");
    expect(csvText(model, { sepLine: true, delimiter: "|" })).toBe("sep=|\na\nb\n");
  });

  it("writes each cell's text in the dialect's locale", () => {
    const model = read("1.5,2026-10-08,1.50\n");
    expect(csvText(model, { delimiter: ";", localeTag: "de-DE" })).toBe("1,5;2026-10-08;1,50\n");
  });
});

describe("encodeCsvText", () => {
  it("writes UTF-8 through TextEncoder, with EF BB BF when bom", () => {
    expect([...encodeCsvText("é", "utf-8", false)!]).toEqual([0xc3, 0xa9]);
    expect([...encodeCsvText("é", "utf-8", true)!]).toEqual([0xef, 0xbb, 0xbf, 0xc3, 0xa9]);
  });

  it("writes UTF-16 by code unit in either byte order, with its BOM when bom", () => {
    expect([...encodeCsvText("a😀", "utf-16le", true)!]).toEqual([0xff, 0xfe, 0x61, 0x00, 0x3d, 0xd8, 0x00, 0xde]);
    expect([...encodeCsvText("a😀", "utf-16be", true)!]).toEqual([0xfe, 0xff, 0x00, 0x61, 0xd8, 0x3d, 0xde, 0x00]);
    expect([...encodeCsvText("a", "utf-16le", false)!]).toEqual([0x61, 0x00]);
    expect(new TextDecoder("utf-16le").decode(encodeCsvText("Ελληνικά 😀", "utf-16le", false)!)).toBe("Ελληνικά 😀");
  });

  it("writes Windows-1252 through the reverse of TextDecoder's own table, every byte 00–FF", () => {
    const all = Uint8Array.from({ length: 256 }, (_, byte) => byte);
    const text = new TextDecoder("windows-1252").decode(all);
    expect(encodeCsvText(text, "windows-1252", false)).toEqual(all);
    expect([...encodeCsvText("Müller €", "windows-1252", false)!]).toEqual([0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x20, 0x80]);
  });

  it("returns null for a character Windows-1252 cannot hold", () => {
    expect(encodeCsvText("Ω", "windows-1252", false)).toBeNull();
    expect(encodeCsvText("ok 😀", "windows-1252", false)).toBeNull();
  });
});

/** A small Windows-1252 `;` file. */
function semicolon1252(): SpreadsheetModel {
  const text = "Name;Ort\nMüller;Köln\n";
  const bytes = Uint8Array.from(Array.from(text), (char) => (char === "ü" ? 0xfc : char === "ö" ? 0xf6 : char.charCodeAt(0)));
  return readCsvBytes(bytes, "csv", "Book", { userLocaleTag: "en-US" });
}

/** The first sheet with `patch` applied, copy-on-write. */
function sheetWith(model: SpreadsheetModel, patch: Partial<Sheet>): void {
  model.sheets = [{ ...model.sheets[0]!, ...patch }, ...model.sheets.slice(1)];
}

describe("writeCsv", () => {

  it("writes the first sheet in the source's dialect", () => {
    const model = semicolon1252();
    dirty(model);
    const { bytes, notices } = writeCsv(model);
    expect(notices).toEqual([]);
    expect(new TextDecoder("windows-1252").decode(bytes)).toBe("Name;Ort\nMüller;Köln\n");
  });

  it("writes only sheets[0], whatever the active sheet", () => {
    const model = read("a,b\n");
    const other = emptySheet("Other");
    model.sheets = [...model.sheets, { ...other, rows: [{ index: 0, cells: [{ col: 0, value: "never" }] }] }];
    model.activeSheetIndex = 1;
    dirty(model);
    expect(decoder.decode(writeCsv(model).bytes)).toBe("a,b\n");
  });

  it("falls back to UTF-8 with a BOM and CSV_ENCODING_FALLBACK_NOTICE when Windows-1252 cannot hold a character, mutating nothing", () => {
    expect(CSV_ENCODING_FALLBACK_NOTICE).toBe("Some characters can't be saved as Windows-1252, so this file was saved as UTF-8");
    const model = semicolon1252();
    setCellEntry(model, 0, { row: 1, col: 1 }, { value: "Ωmega" });
    dirty(model);
    const before = structuredClone(model.csv);
    const { bytes, notices } = writeCsv(model);
    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(decoder.decode(bytes.subarray(3))).toBe("Name;Ort\nMüller;Ωmega\n");
    expect(notices).toEqual([CSV_ENCODING_FALLBACK_NOTICE]);
    expect(model.csv).toEqual(before);
    // The next download decides again.
    setCellEntry(model, 0, { row: 1, col: 1 }, { value: "Köln" });
    expect(writeCsv(model).notices).toEqual([]);
  });
});

describe("csvLosses", () => {
  it("reports nothing for a plain sheet, number formats included (their effect is written)", () => {
    const model = read("1.50,12%,2026-10-08,x\n");
    expect(model.styles.length).toBeGreaterThan(1);
    expect(csvLosses(model, 0)).toEqual([]);
  });

  it('reports "sheets" when another sheet stores a cell — an empty one is no loss', () => {
    const model = read("a\n");
    model.sheets = [...model.sheets, emptySheet("Sheet2")];
    expect(csvLosses(model, 0)).toEqual([]);
    model.sheets = [model.sheets[0]!, { ...model.sheets[1]!, rows: [{ index: 3, cells: [{ col: 2, value: 1 }] }] }];
    expect(csvLosses(model, 0)).toEqual(["sheets"]);
    expect(csvLosses(model, 1)).toEqual(["sheets"]);
  });

  it('reports "formulas" for a formula on the written sheet', () => {
    const model = read("a\n");
    setCellEntry(model, 0, { row: 1, col: 3 }, { value: 2, formula: "1+1" });
    expect(csvLosses(model, 0)).toEqual(["formulas"]);
  });

  it('reports "merges" for a merge on the written sheet', () => {
    const model = read("a,b\n");
    sheetWith(model, { merges: [{ top: 0, left: 0, bottom: 0, right: 1 }] });
    expect(csvLosses(model, 0)).toEqual(["merges"]);
  });

  it('reports "formatting" for a cell, row or column resolving to a font, fill, border or alignment', () => {
    for (const style of [
      { font: { bold: true as const } },
      { fill: { kind: "solid" as const, color: "#ffff00" } },
      { borders: { top: { style: "thin" as const } } },
      { alignment: { horizontal: "center" as const } },
    ]) {
      const model = read("a\n");
      const id = internStyle(model, style);
      setCellEntry(model, 0, { row: 0, col: 0 }, { value: "a" });
      sheetWith(model, { rows: [{ index: 0, cells: [{ col: 0, value: "a", styleId: id }] }] });
      expect(csvLosses(model, 0), JSON.stringify(style)).toEqual(["formatting"]);
    }
    const rowStyled = read("a\n");
    sheetWith(rowStyled, { rows: [{ index: 0, cells: [{ col: 0, value: "a" }], styleId: internStyle(rowStyled, { font: { italic: true } }) }] });
    expect(csvLosses(rowStyled, 0)).toEqual(["formatting"]);
    const columnStyled = read("a\n");
    sheetWith(columnStyled, { columns: [{ min: 0, max: 0, styleId: internStyle(columnStyled, { alignment: { wrap: true } }) }] });
    expect(csvLosses(columnStyled, 0)).toEqual(["formatting"]);
  });

  it("does not report view settings: column widths, row heights, hidden rows, frozen panes", () => {
    const model = read("a\nb\n");
    sheetWith(model, {
      columns: [{ min: 0, max: 0, widthPx: DEFAULT_COL_WIDTH_PX * 2 }],
      rows: [
        { index: 0, cells: [{ col: 0, value: "a" }], heightPt: 30, customHeight: true },
        { index: 1, cells: [{ col: 0, value: "b" }], hidden: true },
      ],
      frozen: { rows: 1, cols: 0 },
    });
    expect(csvLosses(model, 0)).toEqual([]);
  });

  it("counts a workbook's own default look as no formatting (an .xlsx Normal style sets a font)", () => {
    const model = read("a\n");
    model.styles = [{ font: { sizePt: 11 } }, { font: { sizePt: 11 }, numberFormat: "0.00" }, { font: { sizePt: 14 } }];
    sheetWith(model, { rows: [{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }] });
    expect(csvLosses(model, 0)).toEqual([]);
    sheetWith(model, { rows: [{ index: 0, cells: [{ col: 0, value: 1, styleId: 2 }] }] });
    expect(csvLosses(model, 0)).toEqual(["formatting"]);
  });
});
