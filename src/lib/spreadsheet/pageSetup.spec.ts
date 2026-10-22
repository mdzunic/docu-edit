// SPEC-136 "The page-setup model": defaults, the paper catalogue, the
// commands (replaced, never mutated; marked per format), the print area and
// titles accessors over SPEC-133's reserved names, and the "pageBreaks"
// reference adjuster (no MoveAdjuster; names move print areas).
import { describe, expect, it } from "vitest";
import type { DefinedName, Sheet, SpreadsheetModel } from "./model";
import { MOVE_ADJUSTERS } from "./move";
import {
  adjustBreaks,
  DEFAULT_SHEET_MARGINS,
  effectivePaper,
  insertPageBreak,
  normalizePageSetup,
  paperForXlsxCode,
  paperNameOf,
  printAreaOf,
  printTitlesOf,
  removePageBreak,
  resetPageBreaks,
  setPageSetup,
  setPrintArea,
  setPrintTitles,
  XLSX_PAPER_SIZES,
  xlsxCodeOfPaper,
} from "./pageSetup";
import { applyStructuralEdit, REFERENCE_ADJUSTERS, type StructuralEdit } from "./structure";

function sheet(extra: Partial<Sheet> = {}): Sheet {
  return { id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows: [{ index: 0, cells: [{ col: 0, value: 1 }] }], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], conditionalFormats: [], validations: [], partName: "xl/worksheets/sheet1.xml", ...extra };
}

function book(format: "xlsx" | "ods", extra: Partial<Sheet> = {}, names: DefinedName[] = []): SpreadsheetModel {
  return { kind: "spreadsheet", id: "b", title: "Book", format, readerVersion: 16, sheets: [sheet(extra)], activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0, dxfs: [], names } as SpreadsheetModel;
}

const edit = (kind: StructuralEdit["kind"], at: number, count: number): StructuralEdit => ({ kind, sheetId: "s1", at, count });

describe("the model's defaults and the paper catalogue", () => {
  it("DEFAULT_SHEET_MARGINS is Excel's Normal; normalizePageSetup drops every default and an empty setup", () => {
    expect(DEFAULT_SHEET_MARGINS).toEqual({ topCm: 1.905, rightCm: 1.778, bottomCm: 1.905, leftCm: 1.778, headerCm: 0.762, footerCm: 0.762 });
    expect(normalizePageSetup({ scale: 100, margins: { ...DEFAULT_SHEET_MARGINS }, rowBreaks: [], headerFooter: { odd: { header: { left: "" } } } })).toBeUndefined();
    expect(normalizePageSetup({ scale: 75, rowBreaks: [5, 0, 5, 2] })).toEqual({ scale: 75, rowBreaks: [2, 5] });
  });

  it("XLSX_PAPER_SIZES maps codes 1, 3, 5, 8, 9, 11 onto PAPER_SIZES and 7, 12, 13, 14, 66, 70 by dimension; an unknown code is kept as Letter", () => {
    expect(XLSX_PAPER_SIZES.map((entry) => `${entry.code}:${entry.name}`)).toEqual(["1:Letter", "3:Tabloid", "5:Legal", "7:Executive", "8:A3", "9:A4", "11:A5", "12:B4", "13:B5", "14:Folio", "66:A2", "70:A6"]);
    expect(paperForXlsxCode(9)).toEqual({ widthCm: 21, heightCm: 29.7, xlsxPaperSize: 9 });
    expect(paperForXlsxCode(42)).toEqual({ widthCm: 21.59, heightCm: 27.94, xlsxPaperSize: 42 });
    expect(xlsxCodeOfPaper({ widthCm: 42, heightCm: 59.4 })).toBe(66);
    expect(xlsxCodeOfPaper({ widthCm: 20, heightCm: 20 })).toBeUndefined();
    expect(paperNameOf({ widthCm: 20, heightCm: 20 }, { widthCm: 21, heightCm: 29.7 })).toBe("Custom");
    expect(paperNameOf(undefined, { widthCm: 21, heightCm: 29.7 })).toBe("A4");
  });

  it("effectivePaper: the setup's paper, else the user's default page size", () => {
    expect(effectivePaper(sheet(), { defaultPageSizeId: "Letter" })).toEqual({ widthCm: 21.59, heightCm: 27.94 });
    expect(effectivePaper(sheet({ pageSetup: { paper: { widthCm: 29.7, heightCm: 42, xlsxPaperSize: 8 } } }), { defaultPageSizeId: "Letter" })).toEqual({ widthCm: 29.7, heightCm: 42 });
  });
});

describe("the commands", () => {
  it("setPageSetup replaces the setup (the old object untouched) and marks an xlsx sheet pageSetupDirty, never dirty", () => {
    const model = book("xlsx", { pageSetup: { scale: 80 } });
    const before = model.sheets[0]!.pageSetup!;
    expect(setPageSetup(model, model.sheets[0]!, { orientation: "landscape", scale: undefined })).toBe(true);
    expect(before).toEqual({ scale: 80 });
    expect(model.sheets[0]!.pageSetup).toEqual({ orientation: "landscape" });
    expect(model.sheets[0]!.pageSetupDirty).toBe(true);
    expect(model.sheets[0]!.dirty).toBeUndefined();
    expect(setPageSetup(model, "s1", { orientation: "landscape" })).toBe(false);
  });

  it("ods: a layout edit is pageSetupDirty, a break edit dirty", () => {
    const layout = book("ods");
    setPageSetup(layout, "s1", { scale: 50 });
    expect([layout.sheets[0]!.pageSetupDirty, layout.sheets[0]!.dirty]).toEqual([true, undefined]);
    const breaks = book("ods");
    setPageSetup(breaks, "s1", { rowBreaks: [3] });
    expect([breaks.sheets[0]!.pageSetupDirty, breaks.sheets[0]!.dirty]).toEqual([undefined, true]);
  });

  it("insertPageBreak: Excel's rule by active cell — nothing at A1, a column break in row 1, a row break in column A, both elsewhere", () => {
    const model = book("xlsx");
    expect(insertPageBreak(model, "s1", { row: 0, col: 0 })).toBe(false);
    insertPageBreak(model, "s1", { row: 0, col: 3 });
    insertPageBreak(model, "s1", { row: 7, col: 0 });
    insertPageBreak(model, "s1", { row: 9, col: 5 });
    expect(model.sheets[0]!.pageSetup).toEqual({ rowBreaks: [7, 9], colBreaks: [3, 5] });
    removePageBreak(model, "s1", { row: 9, col: 5 });
    expect(model.sheets[0]!.pageSetup).toEqual({ rowBreaks: [7], colBreaks: [3] });
    resetPageBreaks(model, "s1");
    expect(model.sheets[0]!.pageSetup).toBeUndefined();
  });

  it("insertPageBreak on a fit-to-page sheet switches it to the percent the fit printed at, as Excel does", () => {
    const model = book("xlsx", { pageSetup: { fitTo: { width: 1, height: 0 } } });
    insertPageBreak(model, "s1", { row: 29, col: 0 }, 72);
    expect(model.sheets[0]!.pageSetup).toEqual({ scale: 72, rowBreaks: [29] });
  });

  it("print area and titles are the sheet-scoped reserved names, read back by printAreaOf / printTitlesOf", () => {
    const model = book("xlsx");
    setPrintArea(model, "s1", [
      { top: 0, left: 0, bottom: 19, right: 3 },
      { top: 0, left: 5, bottom: 0, right: 5 },
    ]);
    expect(model.names).toEqual([{ name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Sheet1!$A$1:$D$20,Sheet1!$F$1" }]);
    expect(model.namesDirty).toBe(true);
    expect(printAreaOf(model, model.sheets[0]!)).toEqual([
      { top: 0, left: 0, bottom: 19, right: 3 },
      { top: 0, left: 5, bottom: 0, right: 5 },
    ]);
    setPrintTitles(model, "s1", { rows: [0, 1], cols: [0, 0] });
    expect(model.names[1]).toEqual({ name: "_xlnm.Print_Titles", scopeSheetId: "s1", formula: "Sheet1!$1:$2,Sheet1!$A:$A" });
    expect(printTitlesOf(model, model.sheets[0]!)).toEqual({ rows: [0, 1], cols: [0, 0] });
    setPrintArea(model, "s1", null);
    expect(printAreaOf(model, model.sheets[0]!)).toBeNull();
  });
});

describe('the "pageBreaks" adjuster (SPEC-124)', () => {
  it("is registered after SPEC-135's, with no MoveAdjuster: a cut-paste leaves breaks where they are", () => {
    expect(REFERENCE_ADJUSTERS.at(-1)!.id).toBe("pageBreaks");
    expect(MOVE_ADJUSTERS.map((adjuster) => adjuster.id)).not.toContain("pageBreaks");
  });

  it("a break at or after an inserted band shifts by count; one inside a deleted band moves to `at`, deduplicated", () => {
    expect(adjustBreaks([3, 5, 9], edit("insertRows", 5, 2))).toEqual([3, 7, 11]);
    expect(adjustBreaks([3, 5, 6, 9], edit("deleteRows", 4, 3))).toEqual([3, 4, 6]);
    expect(adjustBreaks([2, 4], edit("insertCols", 0, 1))).toEqual([3, 5]);
  });

  it("a row inserted above a break moves it (end to end), and the print area moves once, by the names adjuster", () => {
    const model = book("xlsx", { pageSetup: { rowBreaks: [10], colBreaks: [2] } }, [{ name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Sheet1!$A$5:$C$20" }]);
    expect(applyStructuralEdit(model, edit("insertRows", 3, 2))).toEqual({ ok: true });
    expect(model.sheets[0]!.pageSetup).toEqual({ rowBreaks: [12], colBreaks: [2] });
    expect(model.names[0]!.formula).toBe("Sheet1!$A$7:$C$22");
  });
});
