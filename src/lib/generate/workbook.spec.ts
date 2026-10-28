// SPEC-140 — the fluent WorkbookBuilder: every method's model output, each
// validation error with its exact path text, the current-sheet rules, cell
// input (dates in both date systems, serial 60 included), parseText (a
// field "=SUM(A1:A2)" stays text), the CellStyleSpec mapping, geometry and
// merges' value refusal, the materialisation order (a later style patch
// wins; the last cell write wins), reuse of one builder for two outputs, and
// the output methods SPEC-106's deck builder has.
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it, vi } from "vitest";
import { setClock } from "../core/clock";
import { equalBytes } from "../core/equality";
import { DEFAULT_THEME } from "../core/theme";
import { unzip } from "../core/zip";
import { getCell, parseA1 } from "../spreadsheet/address";
import { FORMAT_TOO_LARGE_MESSAGE } from "../spreadsheet/formatting";
import { DEFAULT_SPREADSHEET_TITLE, type Cell, type CellStyle, type SpreadsheetModel } from "../spreadsheet/model";
import { ODS_MIMETYPE, writeOds } from "../spreadsheet/ods/writer";
import { dateToSerial } from "../spreadsheet/serial";
import { createBlankWorkbook, LAST_VISIBLE_SHEET_MESSAGE, SHEET_NAME_MESSAGES } from "../spreadsheet/sheets";
import { writeXlsx } from "../spreadsheet/xlsx/writer";
import { DocumentGenerationError } from "./errors";
import { WorkbookBuilder } from "./workbook";
import { NO_SHEET_MESSAGE } from "./workbookPlan";

const MODIFIED = new Date(Date.UTC(2026, 9, 8, 9, 30));
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

afterEach(() => {
  setClock(null);
  vi.unstubAllGlobals();
});

function cellAt(model: SpreadsheetModel, ref: string, sheet = 0): Cell | undefined {
  const at = parseA1(ref)!;
  return getCell(model.sheets[sheet]!, at.row, at.col);
}

function styleAt(model: SpreadsheetModel, ref: string, sheet = 0): CellStyle {
  return model.styles[cellAt(model, ref, sheet)?.styleId ?? 0]!;
}

/** A model with its minted ids masked, for comparing two materialisations. */
function masked(model: SpreadsheetModel): unknown {
  return { ...model, id: "", sheets: model.sheets.map((sheet) => ({ ...sheet, id: "" })) };
}

const fails = (build: () => unknown): string => {
  try {
    build();
  } catch (error) {
    expect(error).toBeInstanceOf(DocumentGenerationError);
    return (error as Error).message;
  }
  throw new Error("expected a DocumentGenerationError");
};

/** The Reference's "Q3 sales" workbook. */
function q3Sales(): WorkbookBuilder {
  return new WorkbookBuilder({ title: "Q3 sales" })
    .defineStyle("header", { font: { bold: true }, fill: "#dce6f1", borders: { bottom: "thin" } })
    .addSheet("Sales")
    .setRow("A1", ["Region", "July", "August", "September", "Total"])
    .setRows("A2", [
      ["East", 1200, 1350, 1410],
      ["West", 980, 1010, 1130],
    ])
    .setFormula("E2", "SUM(B2:D2)")
    .setFormula("E3", "SUM(B3:D3)")
    .style("A1:E1", "header")
    .numberFormat("B2:E3", "$#,##0.00")
    .columnWidth("A", { chars: 14 })
    .freeze({ rows: 1 });
}

describe("the blank workbook and the constructor (AC-1)", () => {
  it("starts from createBlankWorkbook(): no sheets added is one empty Sheet1", () => {
    const model = new WorkbookBuilder().toWorkbook();
    const blank = createBlankWorkbook();
    expect(model.sheets.map((sheet) => [sheet.name, sheet.rows.length])).toEqual([["Sheet1", 0]]);
    expect(model.title).toBe(DEFAULT_SPREADSHEET_TITLE);
    expect(model.format).toBe("xlsx");
    expect(model.dateSystem).toBe("1900");
    expect(model.styles).toEqual(blank.styles);
    expect(model.theme).toEqual(DEFAULT_THEME);
    expect(model.opaque).toEqual([]);
  });

  it("takes title, format, dateSystem and locale; the builder's sheets replace Sheet1", () => {
    const model = new WorkbookBuilder({ title: "Budget", format: "ods", dateSystem: "1904", locale: "de-DE" }).addSheet("Income").addSheet("Expenses").toWorkbook();
    expect([model.title, model.format, model.dateSystem]).toEqual(["Budget", "ods", "1904"]);
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Income", "Expenses"]);
  });

  it("validates its options with the description's root paths", () => {
    expect(fails(() => new WorkbookBuilder({ format: "csv" } as never))).toBe('format: expected one of "xlsx", "ods", got "csv"');
    expect(fails(() => new WorkbookBuilder({ dateSystem: "1901" } as never))).toBe('dateSystem: expected one of "1900", "1904", got "1901"');
    expect(fails(() => new WorkbookBuilder({ locale: "nl-NL" } as never))).toMatch(/^locale: expected one of "en-US", "en-GB"/);
    expect(fails(() => new WorkbookBuilder({ title: 7 } as never))).toBe("title: expected a string");
    expect(fails(() => new WorkbookBuilder({ theme: "dark" } as never))).toBe("theme: unknown property");
    expect(fails(() => new WorkbookBuilder("Q3" as never))).toBe("options: expected an object");
  });
});

describe("every method returns this (AC-2)", () => {
  it("chains", () => {
    const builder = new WorkbookBuilder();
    const calls: ((b: WorkbookBuilder) => WorkbookBuilder)[] = [
      (b) => b.defineStyle("bold", { font: { bold: true } }),
      (b) => b.addSheet("One"),
      (b) => b.addSheet("Two"),
      (b) => b.sheet("One"),
      (b) => b.setCell("A1", 1),
      (b) => b.setRow("A2", [1, 2]),
      (b) => b.setRows("A3", [[1], [2]]),
      (b) => b.setFormula("B1", "A1*2"),
      (b) => b.style("A1", "bold"),
      (b) => b.numberFormat("A1:A4", "0.00"),
      (b) => b.columnWidth("A", { chars: 10 }),
      (b) => b.rowHeight(1, { pt: 20 }),
      (b) => b.merge("D1:E1"),
      (b) => b.freeze({ rows: 1 }),
    ];
    for (const call of calls) {
      expect(call(builder)).toBe(builder);
    }
  });

  it("a validation error names the path and leaves the builder as it was", () => {
    const builder = new WorkbookBuilder().addSheet("Sales").setCell("A1", 1);
    expect(fails(() => builder.setCell("B5", { formula: "SUMM(A1)" }))).toBe("sheets[0].cells.B5.formula: unknown function SUMM");
    expect(fails(() => builder.setRows("A1", [[2], [Number.NaN]]))).toBe("sheets[0].data[0].rows[1][0]: expected a finite number");
    const model = builder.toWorkbook();
    expect(cellAt(model, "A1")?.value).toBe(1);
    expect(cellAt(model, "B5")).toBeUndefined();
    // The failed setRows was not counted: the next one is still data[0].
    expect(fails(() => builder.setRow("A1", [Number.POSITIVE_INFINITY]))).toBe("sheets[0].data[0].rows[0][0]: expected a finite number");
  });
});

describe("sheets (AC-3, AC-4, AC-5)", () => {
  it("addSheet appends and makes current, with tabColor, hidden and showGridLines", () => {
    const model = new WorkbookBuilder()
      .addSheet("Data", { tabColor: "#4472C4", showGridLines: false })
      .setCell("A1", "on Data")
      .addSheet("Lookup", { hidden: true })
      .setCell("A1", "on Lookup")
      .toWorkbook();
    const [data, lookup] = model.sheets;
    expect(data).toMatchObject({ name: "Data", visibility: "visible", tabColor: "#4472c4", showGridLines: false });
    expect(lookup).toMatchObject({ name: "Lookup", visibility: "hidden" });
    expect(lookup!.tabColor).toBeUndefined();
    expect(lookup!.showGridLines).toBeUndefined();
    expect(cellAt(model, "A1", 0)?.value).toBe("on Data");
    expect(cellAt(model, "A1", 1)?.value).toBe("on Lookup");
    expect(model.activeSheetIndex).toBe(0);
  });

  it("validates the options at the sheet's path", () => {
    expect(fails(() => new WorkbookBuilder().addSheet("A", { tabColor: "blue" }))).toBe('sheets[0].tabColor: expected a "#rrggbb" colour, got "blue"');
    expect(fails(() => new WorkbookBuilder().addSheet("A").addSheet("B", { hidden: "yes" } as never))).toBe("sheets[1].hidden: expected a boolean");
    expect(fails(() => new WorkbookBuilder().addSheet("A", { zoom: 2 } as never))).toBe("sheets[0].zoom: unknown property");
  });

  it("sheet(name) makes an existing sheet current, in any case", () => {
    const model = new WorkbookBuilder().addSheet("Income").addSheet("Summary").sheet("income").setCell("B2", 5).toWorkbook();
    expect(cellAt(model, "B2", 0)?.value).toBe(5);
    expect(model.sheets[1]!.rows).toEqual([]);
    expect(fails(() => new WorkbookBuilder().addSheet("Income").sheet("Costs"))).toBe('no sheet called "Costs" — call addSheet first');
  });

  it('every cell, range, size, merge and freeze method throws "no sheet — call addSheet first" with no sheet', () => {
    const calls: ((b: WorkbookBuilder) => unknown)[] = [
      (b) => b.setCell("A1", 1),
      (b) => b.setRow("A1", [1]),
      (b) => b.setRows("A1", [[1]]),
      (b) => b.setFormula("A1", "1"),
      (b) => b.style("A1", { fill: "#ffffff" }),
      (b) => b.numberFormat("A1", "0"),
      (b) => b.columnWidth("A", { px: 80 }),
      (b) => b.rowHeight(1, { pt: 20 }),
      (b) => b.merge("A1:B1"),
      (b) => b.freeze({ rows: 1 }),
    ];
    for (const call of calls) {
      expect(fails(() => call(new WorkbookBuilder()))).toBe(NO_SHEET_MESSAGE);
    }
    expect(NO_SHEET_MESSAGE).toBe("no sheet — call addSheet first");
  });

  it("names pass SPEC-121's validateSheetName, its message at sheets[i].name", () => {
    const name = (value: string): string => fails(() => new WorkbookBuilder().addSheet("Sales").addSheet(value));
    expect(name(" ")).toBe(`sheets[1].name: ${SHEET_NAME_MESSAGES.blank}`);
    expect(name("x".repeat(32))).toBe(`sheets[1].name: ${SHEET_NAME_MESSAGES.tooLong}`);
    for (const bad of ["a:b", "a\\b", "a/b", "a?b", "a*b", "a[b", "a]b"]) {
      expect(name(bad)).toBe(`sheets[1].name: ${SHEET_NAME_MESSAGES.characters}`);
    }
    expect(name("'Q3")).toBe(`sheets[1].name: ${SHEET_NAME_MESSAGES.apostrophe}`);
    expect(name("Q3'")).toBe(`sheets[1].name: ${SHEET_NAME_MESSAGES.apostrophe}`);
    expect(name("history")).toBe(`sheets[1].name: ${SHEET_NAME_MESSAGES.reserved}`);
    expect(name("SALES")).toBe(`sheets[1].name: There's already a sheet called "SALES"`);
    expect(new WorkbookBuilder().addSheet("x".repeat(31)).addSheet("It's Q3").toWorkbook().sheets).toHaveLength(2);
  });

  it("refuses a workbook whose every sheet is hidden", () => {
    expect(fails(() => new WorkbookBuilder().addSheet("A", { hidden: true }).addSheet("B", { hidden: true }).toWorkbook())).toBe(`sheets[1].hidden: ${LAST_VISIBLE_SHEET_MESSAGE}`);
    expect(new WorkbookBuilder().addSheet("A", { hidden: true }).addSheet("B").toWorkbook().activeSheetIndex).toBe(1);
  });
});

/** A one-cell workbook. */
function one(input: unknown, dateSystem: "1900" | "1904" = "1900"): SpreadsheetModel {
  return new WorkbookBuilder({ dateSystem }).addSheet("S").setCell("A1", input as never).toWorkbook();
}

describe("cell input (AC-6, AC-7)", () => {

  it("strings up to 32,767 characters, a leading = staying text", () => {
    expect(cellAt(one("Region"), "A1")).toEqual({ col: 0, value: "Region" });
    expect(cellAt(one("=SUM(A1:A2)"), "A1")).toEqual({ col: 0, value: "=SUM(A1:A2)" });
    expect(cellAt(one({ value: "=1+1" }), "A1")).toEqual({ col: 0, value: "=1+1" });
    expect(cellAt(one("x".repeat(32_767)), "A1")?.value).toHaveLength(32_767);
    expect(fails(() => one("x".repeat(32_768)))).toBe("sheets[0].cells.A1: text longer than 32,767 characters");
    expect(fails(() => one({ value: "x".repeat(32_768) }))).toBe("sheets[0].cells.A1.value: text longer than 32,767 characters");
    expect(cellAt(one(""), "A1")).toBeUndefined();
  });

  it("finite numbers (-0 is 0), booleans, null and errors", () => {
    expect(cellAt(one(1234.5), "A1")?.value).toBe(1234.5);
    expect(Object.is(cellAt(one(-0), "A1")?.value, 0)).toBe(true);
    expect(cellAt(one(true), "A1")?.value).toBe(true);
    expect(cellAt(one({ value: false }), "A1")?.value).toBe(false);
    expect(cellAt(one(null), "A1")).toBeUndefined();
    expect(cellAt(one({ error: "#N/A" }), "A1")).toEqual({ col: 0, value: { error: "#N/A" } });
    expect(cellAt(one({ error: "#DIV/0!" }), "A1")?.value).toEqual({ error: "#DIV/0!" });
    expect(fails(() => one({ error: "#OOPS" }))).toMatch(/^sheets\[0\]\.cells\.A1\.error: expected one of "#NULL!", "#DIV\/0!"/);
    expect(fails(() => one(Number.NaN))).toBe("sheets[0].cells.A1: expected a finite number");
    expect(fails(() => one(Number.NEGATIVE_INFINITY))).toBe("sheets[0].cells.A1: expected a finite number");
    expect(fails(() => one([1, 2]))).toBe("sheets[0].cells.A1: expected a string, number, boolean, Date, null or cell object");
    expect(fails(() => one(undefined))).toBe("sheets[0].cells.A1: expected a string, number, boolean, Date, null or cell object");
  });

  it("null clears an earlier write, and a style may still store the cell", () => {
    const model = new WorkbookBuilder().addSheet("S").setCell("A1", 5).setCell("A1", null).setCell("B1", { style: { fill: "#ffff00" } }).toWorkbook();
    expect(cellAt(model, "A1")).toBeUndefined();
    expect(cellAt(model, "B1")?.value).toBeNull();
    expect(styleAt(model, "B1").fill).toEqual({ kind: "solid", color: "#ffff00" });
  });

  it("validates the cell object closed, with at most one content key", () => {
    expect(fails(() => new WorkbookBuilder().addSheet("S").setCell("C4", { fomula: "1" } as never))).toBe("sheets[0].cells.C4.fomula: unknown property");
    expect(fails(() => one({ value: 1, formula: "2" }))).toBe('sheets[0].cells.A1: specify at most one of "value", "date", "error" and "formula"');
    expect(fails(() => new WorkbookBuilder().addSheet("S").setCell("A0", 1))).toBe('sheets[0].cells: expected a cell reference like "B3", got "A0"');
    expect(fails(() => new WorkbookBuilder().addSheet("S").setCell("XFE1", 1))).toBe('sheets[0].cells: expected a cell reference like "B3", got "XFE1"');
  });
});

describe("dates (AC-8, AC-9)", () => {

  it("a JS Date is read in UTC, an ISO date by its calendar fields, both through serial.ts", () => {
    const day = dateToSerial(2026, 10, 8, "1900");
    expect(day).toBe(46303);
    for (const input of [new Date(Date.UTC(2026, 9, 8)), { date: "2026-10-08" }]) {
      const model = one(input);
      expect(cellAt(model, "A1")?.value).toBe(day);
      expect(styleAt(model, "A1").numberFormat).toBe("yyyy-mm-dd");
    }
    // Late in the UTC day, whatever the machine's zone: still the 8th.
    expect(cellAt(one(new Date("2026-10-08T23:30:00Z")), "A1")?.value).toBe(day + 23.5 / 24);
    for (const input of [new Date(Date.UTC(2026, 9, 8, 13, 45)), { date: "2026-10-08T13:45" }, { date: "2026-10-08T13:45:00" }, { date: "2026-10-08T13:45:00.000" }]) {
      const model = one(input);
      expect(cellAt(model, "A1")?.value).toBeCloseTo(day + (13 * 60 + 45) / 1440, 10);
      expect(styleAt(model, "A1").numberFormat).toBe("yyyy-mm-dd hh:mm");
    }
    expect(cellAt(one({ date: "2026-10-08T00:00:01.5" }), "A1")?.value).toBeCloseTo(day + 1.5 / 86_400, 12);
  });

  it("converts in the workbook's date system, serial 60 being the 1900 system's fictitious 1900-02-29", () => {
    expect(cellAt(one({ date: "2026-10-08" }, "1904"), "A1")?.value).toBe(46303 - 1462);
    expect(cellAt(one(new Date(Date.UTC(1904, 0, 1)), "1904"), "A1")?.value).toBe(0);
    expect(cellAt(one({ date: "1900-01-01" }), "A1")?.value).toBe(1);
    expect(cellAt(one({ date: "1900-02-28" }), "A1")?.value).toBe(59);
    expect(cellAt(one({ date: "1900-02-29" }), "A1")?.value).toBe(60);
    expect(cellAt(one({ date: "1900-03-01" }), "A1")?.value).toBe(61);
    expect(cellAt(one(new Date(Date.UTC(1900, 1, 28))), "A1")?.value).toBe(59);
    expect(cellAt(one(new Date(Date.UTC(1900, 2, 1))), "A1")?.value).toBe(61);
    expect(fails(() => one({ date: "1900-02-29" }, "1904"))).toBe('sheets[0].cells.A1.date: "1900-02-29" is not a calendar date and time');
  });

  it("a supplied number format replaces the default", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setCell("A1", { date: "2026-10-08", numberFormat: "d mmm yyyy" })
      .setCell("A2", { date: "2026-10-08", style: { numberFormat: "mm/dd/yy", font: { bold: true } } })
      .toWorkbook();
    expect(styleAt(model, "A1").numberFormat).toBe("d mmm yyyy");
    expect(styleAt(model, "A2")).toMatchObject({ numberFormat: "mm/dd/yy", font: { bold: true } });
  });

  it("an invalid Date, an offset-bearing or malformed ISO string is a validation error", () => {
    expect(fails(() => one(new Date(Number.NaN)))).toBe("sheets[0].cells.A1: invalid Date");
    for (const offset of ["2026-10-08T13:45:00Z", "2026-10-08T13:45:00+02:00", "2026-10-08T13:45-0500"]) {
      expect(fails(() => one({ date: offset }))).toBe(`sheets[0].cells.A1.date: "${offset}" has a UTC offset — give the calendar date and time to store, without one`);
    }
    expect(fails(() => one({ date: "08/10/2026" }))).toBe('sheets[0].cells.A1.date: expected "YYYY-MM-DD[THH:MM[:SS[.sss]]]", got "08/10/2026"');
    expect(fails(() => one({ date: "2026-02-30" }))).toBe('sheets[0].cells.A1.date: "2026-02-30" is not a calendar date and time');
    expect(fails(() => one({ date: "2026-10-08T24:00" }))).toBe('sheets[0].cells.A1.date: "2026-10-08T24:00" is not a calendar date and time');
    expect(fails(() => one({ date: 46303 }))).toBe("sheets[0].cells.A1.date: expected a string");
  });
});

describe("setCell, setRow and setRows (AC-10)", () => {
  it("setRow writes left to right from ref, setRows a 2-D block", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setRow("B2", ["a", 1, true])
      .setRows("C4", [
        [1, 2],
        [3, null, 5],
      ])
      .toWorkbook();
    expect(model.sheets[0]!.rows).toEqual([
      { index: 1, cells: [{ col: 1, value: "a" }, { col: 2, value: 1 }, { col: 3, value: true }] },
      { index: 3, cells: [{ col: 2, value: 1 }, { col: 3, value: 2 }] },
      { index: 4, cells: [{ col: 2, value: 3 }, { col: 4, value: 5 }] },
    ]);
  });

  it("counts setRow and setRows as data[j], each value at rows[r][c]", () => {
    const builder = new WorkbookBuilder().addSheet("S").setRow("A1", [1]).setRows("A2", [[2]]);
    expect(fails(() => builder.setRows("A3", [["ok", { error: "#BAD" as never }]]))).toMatch(/^sheets\[0\]\.data\[2\]\.rows\[0\]\[1\]\.error: expected one of/);
    expect(fails(() => builder.setRows("A3", [[1], 2] as never))).toBe("sheets[0].data[2].rows[1]: expected an array");
    expect(fails(() => builder.setRows("A3", "x" as never))).toBe("sheets[0].data[2].rows: expected an array");
    expect(fails(() => builder.setRow("A3", "x" as never))).toBe("sheets[0].data[2].rows[0]: expected an array");
    expect(fails(() => builder.setRows("3A", [[1]]))).toBe('sheets[0].data[2].at: expected a cell reference like "B3", got "3A"');
    expect(fails(() => builder.setRows("A3", [[1]], { parse: true } as never))).toBe("sheets[0].data[2].parse: unknown property");
    expect(fails(() => builder.setRow("XFD1", [1, 2]))).toBe("sheets[0].data[2].rows[0][1]: past the last cell of the sheet (XFD1048576)");
  });
});

describe("parseText (AC-11, AC-12)", () => {
  it("routes strings through inferCsvField in the builder's locale", () => {
    const typed = new WorkbookBuilder({ locale: "en-GB" })
      .addSheet("S")
      .setRows("A1", [["1,234.50", "45%", "10/08/2026", "TRUE", "#N/A", "007", "plain", ""]], { parseText: true })
      .toWorkbook();
    expect(cellAt(typed, "A1")?.value).toBe(1234.5);
    expect(styleAt(typed, "A1").numberFormat).toBe("#,##0.00");
    expect(cellAt(typed, "B1")?.value).toBe(0.45);
    expect(styleAt(typed, "B1").numberFormat).toBe("0%");
    expect(cellAt(typed, "C1")?.value).toBe(dateToSerial(2026, 8, 10, "1900"));
    expect(styleAt(typed, "C1")).toMatchObject({ numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(cellAt(typed, "D1")?.value).toBe(true);
    expect(cellAt(typed, "E1")?.value).toEqual({ error: "#N/A" });
    expect(cellAt(typed, "F1")?.value).toBe("007");
    expect(cellAt(typed, "G1")?.value).toBe("plain");
    expect(cellAt(typed, "H1")).toBeUndefined();
  });

  it("follows SPEC-126's lossless rule: en-US types 10/8/2026, keeps 10/08/2026 as text", () => {
    const model = new WorkbookBuilder().addSheet("S").setRows("A1", [["10/8/2026", "10/08/2026", "1.234,50"]], { parseText: true }).toWorkbook();
    expect(cellAt(model, "A1")?.value).toBe(46303);
    expect(cellAt(model, "B1")?.value).toBe("10/08/2026");
    expect(cellAt(model, "C1")?.value).toBe("1.234,50");
    const german = new WorkbookBuilder({ locale: "de-DE" }).addSheet("S").setRows("A1", [["1.234,50"]], { parseText: true }).toWorkbook();
    expect(cellAt(german, "A1")?.value).toBe(1234.5);
  });

  it("never makes a formula: =SUM(A1:A2) and =HYPERLINK(…) stay text", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setRows("A1", [["=SUM(A1:A2)", '=HYPERLINK("http://example.com","x")', "+1+2", "@SUM(1)"]], { parseText: true })
      .toWorkbook();
    for (const cell of model.sheets[0]!.rows[0]!.cells) {
      expect(cell.formula).toBeUndefined();
      expect(typeof cell.value).toBe("string");
    }
    expect(cellAt(model, "A1")?.value).toBe("=SUM(A1:A2)");
  });

  it("does not use SPEC-122's parseInput anywhere in the workbook modules", () => {
    const dir = dirname(fileURLToPath(import.meta.url));
    const modules = readdirSync(dir).filter((name) => /^(workbook\w*|readWorkbook)\.ts$/.test(name));
    expect(modules.length).toBeGreaterThanOrEqual(5);
    for (const name of modules) {
      expect(readFileSync(join(dir, name), "utf8"), name).not.toMatch(/parseInput|numberFormat\/input/);
    }
  });

  it("only applies with parseText: true", () => {
    const model = new WorkbookBuilder().addSheet("S").setRows("A1", [["1,234.50"]], { parseText: false }).setRows("A2", [["45%"]]).toWorkbook();
    expect(cellAt(model, "A1")?.value).toBe("1,234.50");
    expect(cellAt(model, "A2")?.value).toBe("45%");
    expect(fails(() => new WorkbookBuilder().addSheet("S").setRows("A1", [["1"]], { parseText: "yes" } as never))).toBe("sheets[0].data[0].parseText: expected a boolean");
  });
});

describe("formulas (AC-13, AC-14, AC-15, AC-16)", () => {
  it("takes canonical text with or without a leading =, through setFormula or { formula }", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setRow("A1", [1, 2])
      .setFormula("C1", "=SUM(A1:B1)")
      .setCell("D1", { formula: "A1*B1" })
      .setFormula("E1", "sum( a1 : b1 )")
      .toWorkbook();
    expect(cellAt(model, "C1")).toEqual({ col: 2, value: 3, formula: "SUM(A1:B1)" });
    expect(cellAt(model, "D1")).toEqual({ col: 3, value: 2, formula: "A1*B1" });
    expect(cellAt(model, "E1")?.formula).toBe("SUM(A1:B1)");
  });

  it("a FormulaSyntaxError becomes the error at …formula, with its offset", () => {
    const formula = (text: string): string => fails(() => new WorkbookBuilder().addSheet("S").setFormula("A9", text));
    expect(formula("SUM(A1))")).toBe('sheets[0].cells.A9.formula: unexpected ")" at 7');
    expect(formula("=SUM(A1))")).toBe('sheets[0].cells.A9.formula: unexpected ")" at 7');
    expect(formula("1+")).toBe("sheets[0].cells.A9.formula: unexpected end of formula at 2");
    expect(formula("=")).toBe("sheets[0].cells.A9.formula: expected a formula");
    expect(fails(() => new WorkbookBuilder().addSheet("S").setRows("A1", [[1, { formula: "A1+" }]]))).toBe("sheets[0].data[0].rows[0][1].formula: unexpected end of formula at 3");
  });

  it("rejects a function neither SPEC-120 nor SPEC-127 registers, by name", () => {
    expect(fails(() => new WorkbookBuilder().addSheet("S").setFormula("B5", "SUMM(B1:B4)"))).toBe("sheets[0].cells.B5.formula: unknown function SUMM");
    expect(fails(() => new WorkbookBuilder().addSheet("S").setFormula("B5", "ROUND(FROBNICATE(1),2)"))).toBe("sheets[0].cells.B5.formula: unknown function FROBNICATE");
    const library = new WorkbookBuilder().addSheet("S").setRow("A1", [3, 1, 2]).setFormula("A2", 'TEXTJOIN("-",TRUE,A1:C1)').setFormula("B2", "XLOOKUP(1,A1:C1,A1:C1)").toWorkbook();
    expect(cellAt(library, "A2")?.value).toBe("3-1-2");
    expect(cellAt(library, "B2")?.value).toBe(1);
  });

  it("rejects what this API cannot build: defined names, table and external references", () => {
    expect(fails(() => new WorkbookBuilder().addSheet("S").setFormula("A1", "Rate*2"))).toBe("sheets[0].cells.A1.formula: unknown name Rate (defined names are not built by this API)");
    expect(fails(() => new WorkbookBuilder().addSheet("S").setFormula("A1", "SUM(Sales[Total])"))).toMatch(/^sheets\[0\]\.cells\.A1\.formula: Sales\[Total\] is a table reference/);
    expect(fails(() => new WorkbookBuilder().addSheet("S").setFormula("A1", "[1]Data!A1"))).toBe("sheets[0].cells.A1.formula: [1]Data!A1 is an external reference");
  });

  it("at output, rejects a sheet the workbook does not have and any circular reference", () => {
    const missing = new WorkbookBuilder().addSheet("Sales").setFormula("B2", "Costs!A1*2");
    expect(fails(() => missing.toWorkbook())).toBe('sheets[0].cells.B2.formula: refers to a sheet the workbook does not have ("Costs")');
    expect(cellAt(missing.addSheet("Costs").toWorkbook(), "B2")?.value).toBe(0);
    expect(fails(() => new WorkbookBuilder().addSheet("Sales").setFormula("B2", "B2+1").toWorkbook())).toBe("sheets[0].cells.B2.formula: circular reference (Sales!B2)");
    expect(fails(() => new WorkbookBuilder().addSheet("Q3 Sales").setFormula("B2", "C2").setFormula("C2", "B2").toWorkbook())).toBe("sheets[0].cells.B2.formula: circular reference ('Q3 Sales'!B2)");
    const threeD = new WorkbookBuilder().addSheet("Jan").addSheet("Feb").addSheet("Total").setFormula("A1", "SUM(Jan:Mar!A1)");
    expect(fails(() => threeD.toWorkbook())).toBe('sheets[2].cells.A1.formula: refers to a sheet the workbook does not have ("Mar")');
  });

  it("computes every formula's cached value with SPEC-120's engine", () => {
    setClock(new Date(Date.UTC(2026, 9, 8, 12)));
    const model = new WorkbookBuilder()
      .addSheet("Data")
      .setRows("A1", [
        [10, 20],
        [30, 40],
      ])
      .addSheet("Summary")
      .setFormula("A1", "SUM(Data!A1:B2)")
      .setFormula("A2", "A1/4")
      .setFormula("A3", 'IF(A2>20,"high","low")&"!"')
      .setFormula("A4", "1/0")
      .setFormula("A5", "AVERAGE(Data!A1:A2)")
      .setFormula("A6", "TODAY()")
      .setFormula("A7", "A8*2")
      .setCell("A8", 21)
      .toWorkbook();
    const values = model.sheets[1]!.rows.map((row) => row.cells[0]!.value);
    expect(values).toEqual([100, 25, "high!", { error: "#DIV/0!" }, 20, dateToSerial(2026, 10, 8, "1900"), 42, 21]);
    expect(model.forceFullCalcOnLoad).toBeUndefined();
  });
});

describe("styles (AC-17, AC-18, AC-19)", () => {
  it("maps CellStyleSpec onto CellStyle field for field", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setCell("A1", "x")
      .style("A1", {
        font: { family: "Arial", sizePt: 14, bold: true, italic: true, underline: "double", strike: true, color: "#1F4E79" },
        fill: "#DCE6F1",
        borders: { top: "thin", bottom: { style: "double", color: "#FF0000" } },
        alignment: { horizontal: "center", vertical: "top", wrap: true, indent: 2, rotation: 45 },
        numberFormat: "0.00",
      })
      .toWorkbook();
    const { namedStyle, ...style } = styleAt(model, "A1");
    expect(namedStyle).toBe("Normal");
    expect(style).toEqual({
      font: { family: "Arial", sizePt: 14, bold: true, italic: true, underline: "double", strike: true, color: "#1f4e79" },
      fill: { kind: "solid", color: "#dce6f1" },
      borders: { top: { style: "thin" }, bottom: { style: "double", color: "#ff0000" } },
      alignment: { horizontal: "center", vertical: "top", wrap: true, indent: 2, textRotation: 45 },
      numberFormat: "0.00",
    });
  });

  it("false removes, as SPEC-123's null does", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setRow("A1", [1, 2])
      .style("A1:B1", { font: { bold: true, italic: true }, alignment: { wrap: true, indent: 3 } })
      .style("B1", { font: { bold: false }, alignment: { wrap: false, indent: 0 } })
      .toWorkbook();
    expect(styleAt(model, "A1")).toMatchObject({ font: { bold: true, italic: true }, alignment: { wrap: true, indent: 3 } });
    expect(styleAt(model, "B1").font).toEqual({ family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" }, italic: true });
    expect(styleAt(model, "B1").alignment).toBeUndefined();
  });

  it("all sets every edge of every cell, outline the outer edges only, a named edge being each cell's", () => {
    const all = new WorkbookBuilder().addSheet("S").style("B2:C3", { borders: { all: "thin" } }).toWorkbook();
    for (const ref of ["B2", "C2", "B3", "C3"]) {
      expect(styleAt(all, ref).borders).toEqual({ top: { style: "thin" }, right: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" } });
    }
    const outline = new WorkbookBuilder().addSheet("S").style("B2:D4", { borders: { outline: { style: "medium", color: "#000080" } } }).toWorkbook();
    const medium = { style: "medium", color: "#000080" };
    expect(styleAt(outline, "B2").borders).toEqual({ top: medium, left: medium });
    expect(styleAt(outline, "C2").borders).toEqual({ top: medium });
    expect(styleAt(outline, "D2").borders).toEqual({ top: medium, right: medium });
    expect(styleAt(outline, "B3").borders).toEqual({ left: medium });
    expect(styleAt(outline, "C3").borders).toBeUndefined();
    expect(styleAt(outline, "C4").borders).toEqual({ bottom: medium });
    expect(styleAt(outline, "D4").borders).toEqual({ right: medium, bottom: medium });
    // A named edge is CellStyle's, every cell's own; it applies after all and
    // outline, so it wins where they meet.
    const named = new WorkbookBuilder().addSheet("S").style("B2:C3", { borders: { all: "thin", outline: "thick", top: "double" } }).toWorkbook();
    expect(styleAt(named, "B2").borders).toEqual({ top: { style: "double" }, right: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thick" } });
    expect(styleAt(named, "C3").borders).toEqual({ top: { style: "double" }, right: { style: "thick" }, bottom: { style: "thick" }, left: { style: "thin" } });
  });

  it("applies style, numberFormat and a cell's own style in call order", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setCell("A1", { value: 1, numberFormat: "0.0" })
      .numberFormat("A1:A3", "0.00")
      .setCell("A2", { value: 2, numberFormat: "0%" })
      .setRow("A3", [3])
      .style("A1:A3", { fill: "#ff0000" })
      .style("A3", { fill: "#0000ff" })
      .toWorkbook();
    expect(styleAt(model, "A1")).toMatchObject({ numberFormat: "0.00", fill: { kind: "solid", color: "#ff0000" } });
    expect(styleAt(model, "A2")).toMatchObject({ numberFormat: "0%", fill: { kind: "solid", color: "#ff0000" } });
    expect(styleAt(model, "A3")).toMatchObject({ numberFormat: "0.00", fill: { kind: "solid", color: "#0000ff" } });
  });

  it("styles whole columns and rows by range class", () => {
    const model = new WorkbookBuilder().addSheet("S").setCell("C2", 1).style("B:C", { font: { italic: true } }).numberFormat("2:2", "0.0").toWorkbook();
    const sheet = model.sheets[0]!;
    expect(sheet.columns).toEqual([{ min: 1, max: 2, styleId: expect.any(Number) }]);
    expect(model.styles[sheet.columns[0]!.styleId!]!.font?.italic).toBe(true);
    expect(styleAt(model, "C2")).toMatchObject({ font: { italic: true }, numberFormat: "0.0" });
  });

  it("defineStyle registers a preset, used by name", () => {
    const model = q3Sales().toWorkbook();
    expect(styleAt(model, "A1")).toMatchObject({ font: { bold: true }, fill: { kind: "solid", color: "#dce6f1" }, borders: { bottom: { style: "thin" } } });
    expect(styleAt(model, "E2").numberFormat).toBe("$#,##0.00");
    expect(styleAt(model, "A2").numberFormat).toBeUndefined();
    expect(fails(() => new WorkbookBuilder().addSheet("S").style("A1", "nope"))).toBe('sheets[0].styles[0].style: no style called "nope" — call defineStyle first');
    expect(fails(() => new WorkbookBuilder().addSheet("S").setCell("A1", { value: 1, style: "nope" }))).toBe('sheets[0].cells.A1.style: no style called "nope" — call defineStyle first');
  });

  it("validates every spec field at its path, colours as #rrggbb only", () => {
    const define = (spec: unknown): string => fails(() => new WorkbookBuilder().defineStyle("header", spec as never));
    expect(define({ font: { sizePt: 0 } })).toBe("styles.header.font.sizePt: expected a size between 1 and 409 pt");
    expect(define({ font: { colour: "#000000" } })).toBe("styles.header.font.colour: unknown property");
    expect(define({ font: { underline: "wavy" } })).toBe('styles.header.font.underline: expected one of "single", "double", got "wavy"');
    expect(define({ font: { color: { theme: "accent1" } } })).toBe("styles.header.font.color: expected a string");
    expect(define({ fill: "red" })).toBe('styles.header.fill: expected a "#rrggbb" colour, got "red"');
    expect(define({ borders: { top: "wavy" } })).toMatch(/^styles\.header\.borders\.top: expected one of "thin", "medium"/);
    expect(define({ borders: { diagonal: "thin" } })).toBe("styles.header.borders.diagonal: unknown property");
    expect(define({ borders: { all: { style: "thin", color: "#00f" } } })).toBe('styles.header.borders.all.color: expected a "#rrggbb" colour, got "#00f"');
    expect(define({ alignment: { horizontal: "fill" } })).toBe('styles.header.alignment.horizontal: expected one of "left", "center", "right", "justify", got "fill"');
    expect(define({ alignment: { rotation: 200 } })).toBe("styles.header.alignment.rotation: expected an integer between 0 and 180");
    expect(define({ alignment: { indent: 1.5 } })).toBe("styles.header.alignment.indent: expected an integer between 0 and 250");
    expect(define({ protection: {} })).toBe("styles.header.protection: unknown property");
    expect(define({ numberFormat: '"abc' })).toBe('styles.header.numberFormat: invalid number format ""abc": Unterminated quote (at offset 0)');
    expect(fails(() => new WorkbookBuilder().addSheet("S").style("A1", { fill: "red" }))).toBe('sheets[0].styles[0].style.fill: expected a "#rrggbb" colour, got "red"');
    expect(fails(() => new WorkbookBuilder().addSheet("S").numberFormat("A1", "0;0;0;@;0"))).toBe(
      'sheets[0].styles[0].numberFormat: invalid number format "0;0;0;@;0": A format code has at most four sections (at offset 7)',
    );
    expect(fails(() => new WorkbookBuilder().addSheet("S").setCell("A1", { value: 1, numberFormat: "[Red" }))).toBe(
      'sheets[0].cells.A1.numberFormat: invalid number format "[Red": Unterminated bracket (at offset 0)',
    );
    expect(fails(() => new WorkbookBuilder().addSheet("S").numberFormat("A1:", "0"))).toBe('sheets[0].styles[0].range: expected a range like "A1:D4", "B:D" or "3:5", got "A1:"');
  });

  it("counts style and numberFormat calls as styles[k]", () => {
    const builder = new WorkbookBuilder().addSheet("S").style("A1", { fill: "#ffffff" }).numberFormat("A1", "0");
    expect(fails(() => builder.style("A1", { font: { sizePt: 500 } }))).toBe("sheets[0].styles[2].style.font.sizePt: expected a size between 1 and 409 pt");
  });

  it("an applyStylePatch refusal becomes the error at that call's path", () => {
    const builder = new WorkbookBuilder().addSheet("S").numberFormat("A1", "0").style("A1:B1000000", { fill: "#ffff00" });
    expect(fails(() => builder.toWorkbook())).toBe(`sheets[0].styles[1].style: ${FORMAT_TOO_LARGE_MESSAGE}`);
  });
});

describe("sizes, merges and panes (AC-20, AC-21, AC-22, AC-23)", () => {
  it("columnWidth stores widthPx by the MDW formula and keeps sourceWidthChars", () => {
    const model = new WorkbookBuilder().addSheet("S").columnWidth("B:D", { chars: 14 }).columnWidth("F", { px: 120 }).columnWidth("g", { chars: 255 }).toWorkbook();
    // trunc(((256·14 + trunc(128/7)) / 256) · 7) = 98, MDW 7 for Calibri 11.
    expect(model.sheets[0]!.columns).toEqual([
      { min: 1, max: 3, widthPx: 98, sourceWidthChars: 14 },
      { min: 5, max: 5, widthPx: 120 },
      { min: 6, max: 6, widthPx: 1785, sourceWidthChars: 255 },
    ]);
  });

  it("adjacent columns with equal settings share one span", () => {
    const model = new WorkbookBuilder().addSheet("S").columnWidth("A", { chars: 10 }).columnWidth("B", { chars: 10 }).columnWidth("C", { chars: 12 }).toWorkbook();
    expect(model.sheets[0]!.columns).toEqual([
      { min: 0, max: 1, widthPx: 70, sourceWidthChars: 10 },
      { min: 2, max: 2, widthPx: 84, sourceWidthChars: 12 },
    ]);
  });

  it("validates widths: 0–255 characters, one unit, a column range", () => {
    const width = (columns: string, spec: unknown): string => fails(() => new WorkbookBuilder().addSheet("S").columnWidth(columns, spec as never));
    expect(width("A", { chars: 256 })).toBe("sheets[0].columns[0].width.chars: expected at most 255 characters");
    expect(width("A", { chars: -1 })).toBe("sheets[0].columns[0].width.chars: expected a non-negative number");
    expect(width("A", { px: 2000 })).toBe("sheets[0].columns[0].width.px: expected at most 1785 px (255 characters)");
    expect(width("A", { chars: 1, px: 1 })).toBe('sheets[0].columns[0].width: specify either "chars" or "px"');
    expect(width("A", {})).toBe('sheets[0].columns[0].width: specify either "chars" or "px"');
    expect(width("A", { pt: 1 })).toBe("sheets[0].columns[0].width.pt: unknown property");
    expect(width("A1", { chars: 1 })).toBe('sheets[0].columns[0].columns: expected a column or columns like "B" or "B:D", got "A1"');
  });

  it("rowHeight sets heightPt and customHeight, rows 1-based", () => {
    const model = new WorkbookBuilder().addSheet("S").setCell("A3", "x").rowHeight(3, { pt: 24 }).rowHeight("5:6", { px: 32 }).toWorkbook();
    expect(model.sheets[0]!.rows).toEqual([
      { index: 2, cells: [{ col: 0, value: "x" }], heightPt: 24, customHeight: true },
      { index: 4, cells: [], heightPt: 24, customHeight: true },
      { index: 5, cells: [], heightPt: 24, customHeight: true },
    ]);
    const height = (rows: unknown, spec: unknown): string => fails(() => new WorkbookBuilder().addSheet("S").rowHeight(rows as never, spec as never));
    expect(height(1, { pt: 410 })).toBe("sheets[0].rows[0].height.pt: expected at most 409 pt");
    expect(height(1, { px: 546 })).toBe("sheets[0].rows[0].height.px: expected at most 409 pt");
    expect(height(0, { pt: 10 })).toBe('sheets[0].rows[0].rows: expected a row or rows like 3, "3" or "3:5"');
    expect(height("A", { pt: 10 })).toBe('sheets[0].rows[0].rows: expected a row or rows like 3, "3" or "3:5"');
    expect(height(1, { chars: 10 })).toBe("sheets[0].rows[0].height.chars: unknown property");
  });

  it("merge goes through mergeCells", () => {
    const model = new WorkbookBuilder().addSheet("S").setCell("A1", "Title").style("C1", { fill: "#eeeeee" }).merge("A1:D1").merge("B3:C5").toWorkbook();
    expect(model.sheets[0]!.merges).toEqual([
      { top: 0, left: 0, bottom: 0, right: 3 },
      { top: 2, left: 1, bottom: 4, right: 2 },
    ]);
    expect(cellAt(model, "A1")?.value).toBe("Title");
    expect(cellAt(model, "C1")?.value).toBeNull();
  });

  it("refuses a merge that would discard a value, or that overlaps another", () => {
    const discard = new WorkbookBuilder().addSheet("S").merge("F1:G1").merge("F3:G3").merge("A1:D1").setRow("A1", ["Title", null, "lost"]);
    expect(fails(() => discard.toWorkbook())).toBe("sheets[0].merges[2]: A1:D1 would discard the value in C1");
    const formula = new WorkbookBuilder().addSheet("S").setFormula("B2", "1+1").merge("A1:B2");
    expect(fails(() => formula.toWorkbook())).toBe("sheets[0].merges[0]: A1:B2 would discard the value in B2");
    expect(fails(() => new WorkbookBuilder().addSheet("S").merge("A1:B2").merge("B2:C3"))).toBe("sheets[0].merges[1]: B2:C3 overlaps the merge A1:B2");
    expect(fails(() => new WorkbookBuilder().addSheet("S").merge("A1"))).toBe('sheets[0].merges[0]: expected a range of two or more cells like "A1:D1", got "A1"');
  });

  it("freeze sets Sheet.frozen; the last call wins and 0 × 0 clears", () => {
    expect(new WorkbookBuilder().addSheet("S").freeze({ rows: 1 }).toWorkbook().sheets[0]!.frozen).toEqual({ rows: 1, cols: 0 });
    expect(new WorkbookBuilder().addSheet("S").freeze({ rows: 2, cols: 1 }).toWorkbook().sheets[0]!.frozen).toEqual({ rows: 2, cols: 1 });
    expect(new WorkbookBuilder().addSheet("S").freeze({ rows: 2 }).freeze({ rows: 0, cols: 0 }).toWorkbook().sheets[0]!.frozen).toBeUndefined();
    expect(fails(() => new WorkbookBuilder().addSheet("S").freeze({ rows: -1 }))).toBe("sheets[0].freeze.rows: expected an integer between 0 and 1048575");
    expect(fails(() => new WorkbookBuilder().addSheet("S").freeze({ top: 1 } as never))).toBe("sheets[0].freeze.top: unknown property");
  });
});

describe("materialisation (AC-24, AC-25)", () => {
  it("the last write to a cell wins, and a later style patch wins", () => {
    const model = new WorkbookBuilder()
      .addSheet("S")
      .setCell("A1", 1)
      .setRows("A1", [[2, 3]])
      .setFormula("B1", "A1*10")
      .style("A1", { fill: "#ff0000" })
      .style("A1", { fill: "#00ff00" })
      .toWorkbook();
    expect(cellAt(model, "A1")?.value).toBe(2);
    expect(cellAt(model, "B1")).toMatchObject({ value: 20, formula: "A1*10" });
    expect(styleAt(model, "A1").fill).toEqual({ kind: "solid", color: "#00ff00" });
  });

  it("satisfies SPEC-116's invariants: sorted, sparse, plain JSON, no retained source", () => {
    const model = q3Sales().addSheet("Notes").setCell("C9", "z").setCell("A9", "a").setCell("B2", null).setCell("A1", { date: "2026-10-08" }).merge("A3:B3").toWorkbook();
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
    expect(model.opaque).toEqual([]);
    const text = JSON.stringify(model);
    expect(text).not.toContain("partName");
    expect(text).not.toContain("sourceXml");
    for (const sheet of model.sheets) {
      expect(sheet.rows.map((row) => row.index)).toEqual(sheet.rows.map((row) => row.index).toSorted((a, b) => a - b));
      for (const row of sheet.rows) {
        expect(row.cells.map((cell) => cell.col)).toEqual(row.cells.map((cell) => cell.col).toSorted((a, b) => a - b));
        for (const cell of row.cells) {
          expect(cell.value !== null || cell.formula !== undefined || cell.styleId !== undefined).toBe(true);
        }
      }
    }
    expect(model.sheets[1]!.rows.map((row) => row.cells.map((cell) => cell.col))).toEqual([[0], [0, 2]]);
  });

  it("builds a fresh model on every output, so one builder serves two", async () => {
    const builder = q3Sales();
    const first = builder.toWorkbook();
    const second = builder.toWorkbook();
    expect(first).not.toBe(second);
    expect(masked(first)).toEqual(masked(second));
    first.sheets[0]!.rows = [];
    expect(cellAt(builder.toWorkbook(), "E2")?.value).toBe(3960);
    const xlsx = await builder.toBytes("xlsx", { modified: MODIFIED });
    builder.setCell("A4", "North");
    expect(cellAt(builder.toWorkbook(), "A4")?.value).toBe("North");
    expect(equalBytes(xlsx, await q3Sales().toBytes("xlsx", { modified: MODIFIED }))).toBe(true);
  });
});

describe("output (AC-27, AC-28, AC-29)", () => {
  it("toWorkbook holds the computed values, toBytes writes the builder's format by default", async () => {
    const model = q3Sales().toWorkbook();
    expect(cellAt(model, "E2")).toMatchObject({ value: 3960, formula: "SUM(B2:D2)" });
    expect(cellAt(model, "E3")).toMatchObject({ value: 3120, formula: "SUM(B3:D3)" });
    const xlsx = await unzip(await q3Sales().toBytes());
    expect(xlsx.has("xl/workbook.xml")).toBe(true);
    const ods = await unzip(await new WorkbookBuilder({ format: "ods" }).toBytes());
    expect(new TextDecoder().decode(ods.get("mimetype"))).toBe(ODS_MIMETYPE);
    expect((await unzip(await q3Sales().toBytes("ods"))).has("content.xml")).toBe(true);
  });

  it("the bytes are exactly the writers' from-scratch output for toWorkbook()", async () => {
    for (const format of ["xlsx", "ods"] as const) {
      const ours = await unzip(await q3Sales().toBytes(format, { modified: MODIFIED }));
      const model = { ...q3Sales().toWorkbook(), format };
      const theirs = await unzip(format === "xlsx" ? await writeXlsx(model, { modified: MODIFIED }) : await writeOds(model, { modified: MODIFIED }));
      expect([...ours.keys()]).toEqual([...theirs.keys()]);
      for (const [name, data] of theirs) {
        expect(equalBytes(ours.get(name)!, data), name).toBe(true);
      }
    }
  });

  it("the same input with the same modified gives equalBytes parts; modified reaches the writers", async () => {
    for (const format of ["xlsx", "ods"] as const) {
      const first = await unzip(await q3Sales().toBytes(format, { modified: MODIFIED }));
      const second = await unzip(await q3Sales().toBytes(format, { modified: MODIFIED }));
      for (const [name, data] of first) {
        expect(equalBytes(second.get(name)!, data), `${format} ${name}`).toBe(true);
      }
    }
    const xlsx = await unzip(await q3Sales().toBytes("xlsx", { modified: MODIFIED }));
    expect(new TextDecoder().decode(xlsx.get("docProps/core.xml"))).toContain("2026-10-08T09:30:00Z");
    const ods = await unzip(await q3Sales().toBytes("ods", { modified: MODIFIED }));
    expect(new TextDecoder().decode(ods.get("meta.xml"))).toContain("2026-10-08T09:30:00");
  });

  it("validates format and options before writing", async () => {
    await expect(q3Sales().toBytes("csv" as never)).rejects.toThrow('format: expected one of "xlsx", "ods", got "csv"');
    await expect(q3Sales().toBytes("xlsx", { fullCalcOnLoad: true } as never)).rejects.toThrow("options.fullCalcOnLoad: unknown property");
    await expect(q3Sales().toBytes("xlsx", { modified: "today" } as never)).rejects.toThrow("options.modified: expected a valid Date");
    await expect(new WorkbookBuilder().addSheet("S").setFormula("A1", "A1").toBytes()).rejects.toThrow("sheets[0].cells.A1.formula: circular reference (S!A1)");
  });

  it("toBlob carries the format's MIME type; toBuffer needs a global Buffer", async () => {
    expect((await q3Sales().toBlob()).type).toBe(XLSX_MIME);
    expect((await q3Sales().toBlob("ods")).type).toBe(ODS_MIMETYPE);
    const buffer = await q3Sales().toBuffer("ods", { modified: MODIFIED });
    expect(Buffer.isBuffer(buffer)).toBe(true);
    vi.stubGlobal("Buffer", undefined);
    await expect(q3Sales().toBuffer()).rejects.toThrow(DocumentGenerationError);
    await expect(q3Sales().toBuffer()).rejects.toThrow(/Buffer is not available in this environment/);
  });

  it("fileName follows spreadsheetDownload's naming rule", () => {
    expect(q3Sales().fileName()).toBe("Q3 sales.xlsx");
    expect(q3Sales().fileName("ods")).toBe("Q3 sales.ods");
    expect(new WorkbookBuilder({ title: "Budget", format: "ods" }).fileName()).toBe("Budget.ods");
    expect(new WorkbookBuilder({ title: "   " }).fileName()).toBe("Untitled spreadsheet.xlsx");
    expect(new WorkbookBuilder().fileName("ods")).toBe("Untitled spreadsheet.ods");
    expect(fails(() => new WorkbookBuilder().fileName("pdf" as never))).toBe('format: expected one of "xlsx", "ods", got "pdf"');
  });
});
