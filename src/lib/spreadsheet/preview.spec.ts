// SPEC-116: the library card's preview (buildPreview — the active sheet's
// top-left 12 × 6 through the General display, trimmed to what is used) and
// the foreign-sheet placeholder's text.
import { describe, expect, it } from "vitest";
import { foreignSheetMessage, isChartSheet } from "./foreign";
import type { Row, Sheet, SpreadsheetModel } from "./model";
import { buildPreview, PREVIEW_COLS, PREVIEW_ROWS } from "./preview";

function sheet(name: string, rows: Row[], kind: Sheet["kind"] = "grid", partName?: string): Sheet {
  return { id: name, name, visibility: "visible", kind, rows, partName };
}

function model(sheets: Sheet[], activeSheetIndex = 0, opaque: SpreadsheetModel["opaque"] = []): SpreadsheetModel {
  return { kind: "spreadsheet", id: "w", title: "W", format: "xlsx", readerVersion: 1, sheets, activeSheetIndex, dateSystem: "1900", opaque };
}

describe("buildPreview", () => {
  it("shows the active sheet's corner as General text, \"\" for empty cells", () => {
    const rows: Row[] = [
      { index: 0, cells: [{ col: 0, value: "Name" }, { col: 2, value: 123456789012 }] },
      { index: 2, cells: [{ col: 1, value: true }, { col: 2, value: { error: "#N/A" } }] },
    ];
    expect(buildPreview(model([sheet("Other", []), sheet("Main", rows)], 1))).toEqual({
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
    const preview = buildPreview(model([sheet("Big", rows)]));
    expect(PREVIEW_ROWS).toBe(12);
    expect(PREVIEW_COLS).toBe(6);
    expect(preview.cells).toHaveLength(12);
    expect(preview.cells.every((row) => row.length === 6)).toBe(true);
    expect(preview.cells[11]![5]).toBe("115");
  });

  it("has no cells for an empty or foreign active sheet", () => {
    expect(buildPreview(model([sheet("Empty", [])]))).toEqual({ sheetName: "Empty", cells: [] });
    expect(buildPreview(model([sheet("Chart", [], "foreign")]))).toEqual({ sheetName: "Chart", cells: [] });
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
