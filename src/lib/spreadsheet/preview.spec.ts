// SPEC-116: the library card's preview (buildPreview — the active sheet's
// top-left 12 × 6, trimmed to what is used) and the foreign-sheet
// placeholder's text. SPEC-118: values through their number formats at full
// width, in the locale the caller hands in.
import { describe, expect, it } from "vitest";
import { foreignSheetMessage, isChartSheet } from "./foreign";
import type { CellStyle, Row, Sheet, SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { buildPreview, PREVIEW_COLS, PREVIEW_ROWS } from "./preview";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const DE_DE = SPREADSHEET_LOCALES["de-DE"];

function sheet(name: string, rows: Row[], kind: Sheet["kind"] = "grid", partName?: string): Sheet {
  return { conditionalFormats: [], validations: [], id: name, name, visibility: "visible", kind, rows, partName, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

function model(sheets: Sheet[], activeSheetIndex = 0, opaque: SpreadsheetModel["opaque"] = []): SpreadsheetModel {
  return {
    dxfs: [],
    kind: "spreadsheet",
    id: "w",
    title: "W",
    format: "xlsx",
    readerVersion: 1,
    sheets,
    activeSheetIndex,
    dateSystem: "1900",
    opaque,
    styles: [{}],
    defaultStyleId: 0,
  };
}

describe("buildPreview", () => {
  it("shows the active sheet's corner as General text, \"\" for empty cells", () => {
    const rows: Row[] = [
      { index: 0, cells: [{ col: 0, value: "Name" }, { col: 2, value: 123456789012 }] },
      { index: 2, cells: [{ col: 1, value: true }, { col: 2, value: { error: "#N/A" } }] },
    ];
    expect(buildPreview(model([sheet("Other", []), sheet("Main", rows)], 1), EN_US)).toEqual({
      sheetName: "Main",
      cells: [
        ["Name", "", "1.23457E+11"],
        ["", "", ""],
        ["", "TRUE", "#N/A"],
      ],
    });
  });

  it("stops at PREVIEW_ROWS × PREVIEW_COLS", () => {
    const rows: Row[] = Array.from({ length: 30 }, (_row, index) => ({
      index,
      cells: Array.from({ length: 10 }, (_cell, col) => ({ col, value: index * 10 + col })),
    }));
    const preview = buildPreview(model([sheet("Big", rows)]), EN_US);
    expect(PREVIEW_ROWS).toBe(12);
    expect(PREVIEW_COLS).toBe(6);
    expect(preview.cells).toHaveLength(12);
    expect(preview.cells.every((row) => row.length === 6)).toBe(true);
    expect(preview.cells[11]![5]).toBe("115");
  });

  it("has no cells for an empty or foreign active sheet", () => {
    expect(buildPreview(model([sheet("Empty", [])]), EN_US)).toEqual({ sheetName: "Empty", cells: [] });
    expect(buildPreview(model([sheet("Chart", [], "foreign")]), EN_US)).toEqual({ sheetName: "Chart", cells: [] });
  });

  it("formats each value with its cell's number format at full width, in the locale it is given (SPEC-118)", () => {
    const styles: CellStyle[] = [{}, { numberFormat: "0.00%" }, { builtinFormatId: 14, numberFormat: "m/d/yyyy" }, { numberFormat: '_("$"* #,##0.00_)' }];
    const rows: Row[] = [
      {
        index: 0,
        cells: [
          { col: 0, value: 0.125, styleId: 1 },
          { col: 1, value: 46303.5729166667, styleId: 2 },
          { col: 2, value: 123456789012 }, // General never fits a card's cell: full width
          { col: 3, value: 5, styleId: 3 }, // the fill as one character
          { col: 4, value: true, styleId: 1 },
        ],
      },
    ];
    const workbook = { ...model([sheet("Main", rows)]), styles };
    expect(buildPreview(workbook, EN_US).cells).toEqual([["12.50%", "10/8/2026", "1.23457E+11", " $ 5.00 ", "TRUE"]]);
    expect(buildPreview(workbook, DE_DE).cells).toEqual([["12,50%", "08.10.2026", "1,23457E+11", " $ 5,00 ", "TRUE"]]);
  });

  it("renders an ods-born workbook's dates before 1900 as the calendar's (proleptic Gregorian)", () => {
    const rows: Row[] = [{ index: 0, cells: [{ col: 0, value: -6, styleId: 1 }] }];
    const ods = { ...model([sheet("Main", rows)]), format: "ods" as const, styles: [{}, { numberFormat: "yyyy-mm-dd" }] };
    expect(buildPreview(ods, EN_US).cells).toEqual([["1899-12-25"]]);
    expect(buildPreview({ ...ods, format: "xlsx" }, EN_US).cells).toEqual([["########"]]);
  });
});

/** A retained `[Content_Types].xml` with the given overrides. */
function contentTypes(overrides: string): SpreadsheetModel["opaque"][number] {
  return {
    name: "[Content_Types].xml",
    data: btoa(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${overrides}</Types>`),
  };
}

describe("the foreign-sheet placeholder", () => {
  it("names a chart sheet by its declared content type", () => {
    const chart = sheet("Chart1", [], "foreign", "xl/sheets/odd-name.xml");
    const declared = model(
      [chart],
      0,
      [contentTypes(`<Override PartName="/xl/sheets/odd-name.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>`)],
    );
    expect(isChartSheet(declared, chart)).toBe(true);
    expect(foreignSheetMessage(declared, chart)).toBe("This chart sheet can't be shown yet");
  });

  it("falls back to Excel's chartsheets/ folder when nothing is declared, else the generic text", () => {
    const chart = sheet("Chart1", [], "foreign", "xl/chartsheets/sheet1.xml");
    const dialog = sheet("Dialog1", [], "foreign", "xl/dialogsheets/sheet1.xml");
    expect(foreignSheetMessage(model([chart]), chart)).toBe("This chart sheet can't be shown yet");
    expect(foreignSheetMessage(model([dialog]), dialog)).toBe("This sheet type can't be shown yet");
    const macro = sheet("Macro1", [], "foreign", "xl/macrosheets/sheet1.xml");
    const declared = model(
      [macro],
      0,
      [contentTypes(`<Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>`)],
    );
    expect(foreignSheetMessage(declared, macro)).toBe("This sheet type can't be shown yet");
  });
});
