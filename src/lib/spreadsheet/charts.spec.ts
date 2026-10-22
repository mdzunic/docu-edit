// SPEC-132 charts.ts: references resolved against the grid, the refresh
// rule, insert from selection, the commands, and charts following
// structural edits, moves and sheet renames/deletes.
import { describe, expect, it } from "vitest";
import { defaultChartOptions, type ChartVisual } from "../core/charts";
import { parseA1 } from "./address";
import {
  buildChartFromRange,
  chartsAffectedBy,
  insertSheetChart,
  NO_CHART_DATA_MESSAGE,
  refreshedChartContent,
  refreshSheetChart,
  resolveChartRange,
} from "./charts";
import { applyDrawingResult } from "./drawings";
import { cellKey } from "./formula/dependencies";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "./geometry";
import type { Cell, CellValue, Row, Sheet, SheetChartContent, SheetChartRefs, SheetDrawing, SpreadsheetModel } from "./model";
import { moveRange } from "./move";
import { SPREADSHEET_READER_VERSION } from "./open";
import { deleteSheet, renameSheet } from "./sheets";
import { applyStructuralEdit } from "./structure";
import { defaultStyles } from "./styles";

/** A grid sheet holding `values` by A1 address. */
function sheetOf(name: string, values: Record<string, CellValue> = {}, extra: Partial<Sheet> = {}): Sheet {
  const byRow = new Map<number, Cell[]>();
  for (const [address, value] of Object.entries(values)) {
    const ref = parseA1(address)!;
    byRow.set(ref.row, [...(byRow.get(ref.row) ?? []), { col: ref.col, value }]);
  }
  const rows: Row[] = [...byRow.entries()].toSorted((a, b) => a[0] - b[0]).map(([index, cells]) => ({ index, cells: cells.toSorted((a, b) => a.col - b.col) }));
  return {
    id: name,
    name,
    visibility: "visible",
    kind: "grid",
    rows,
    columns: [],
    defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
    defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
    merges: [],
    conditionalFormats: [],
    validations: [],
    ...extra,
  };
}

function bookOf(sheets: Sheet[]): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "book",
    title: "Charts",
    format: "xlsx",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: defaultStyles(),
    defaultStyleId: 0,
    names: [],
    dxfs: [],
  };
}

const DATA = { A1: "Quarter", B1: "North", A2: "Q1", A3: "Q2", A4: "Q3", B2: 10, B3: 20, B4: 30 };
const REFS: SheetChartRefs = { series: [{ name: "Data!$B$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" }], plotVisibleOnly: true };
const VISUAL: ChartVisual = { chartType: "columnClustered", categories: ["Q1", "Q2", "Q3"], series: [{ name: "North", values: [10, 20, 30], color: "#4e79a7" }], options: defaultChartOptions() };

const corner = (col: number, row: number) => ({ col, colOffEmu: 0, row, rowOffEmu: 0 });

function chartDrawing(refs: SheetChartRefs = REFS, visual: ChartVisual = VISUAL): SheetDrawing {
  return { id: "dr1", name: "Chart 1", anchor: { kind: "twoCell", from: corner(5, 20), to: corner(10, 35) }, content: { kind: "chart", visual, refs, sourcePart: "xl/charts/chart1.xml" } };
}

/** Data (the cells) and Summary (the chart) — the chart plots another sheet's data. */
function workbook(values: Record<string, CellValue> = DATA, refs: SheetChartRefs = REFS): SpreadsheetModel {
  return bookOf([sheetOf("Data", values), sheetOf("Summary", {}, { drawings: [chartDrawing(refs)] })]);
}

const chartOf = (model: SpreadsheetModel, sheet = "Summary"): SheetChartContent => {
  const content = model.sheets.find((each) => each.name === sheet)!.drawings![0]!.content;
  if (content.kind !== "chart") {
    throw new Error("not a chart");
  }
  return content;
};

describe("resolveChartRange", () => {
  it("resolves a sheet-qualified cell, a range and a parenthesised union; anything else is null", () => {
    const model = workbook();
    expect(resolveChartRange(model, "Data!$B$1")).toEqual([{ sheetId: "Data", range: { top: 0, left: 1, bottom: 0, right: 1 } }]);
    expect(resolveChartRange(model, "(Data!$A$2:$A$3,Data!$B$4)")).toEqual([
      { sheetId: "Data", range: { top: 1, left: 0, bottom: 2, right: 0 } },
      { sheetId: "Data", range: { top: 3, left: 1, bottom: 3, right: 1 } },
    ]);
    expect(resolveChartRange(model, "SUM(Data!$B$2:$B$4)")).toBeNull();
    expect(resolveChartRange(model, "[1]Data!$B$2")).toBeNull();
    expect(resolveChartRange(model, "#REF!")).toBeNull();
    expect(resolveChartRange(model, "Sales")).toBeNull();
    expect(resolveChartRange(model, "Sales", () => "Data!$B$2:$B$4")).toEqual([{ sheetId: "Data", range: { top: 1, left: 1, bottom: 3, right: 1 } }]);
  });
});

describe("refresh", () => {
  it("blank and #N/A cells are gaps, text plots as 0, hidden rows are skipped with plotVisibleOnly — and an unchanged chart keeps its identity", () => {
    const model = workbook({ ...DATA, B2: null, B3: { error: "#N/A" }, B4: "n/a" });
    const content = chartOf(model);
    expect(refreshedChartContent(model, "Summary", content).visual.series[0]!.values).toEqual([null, null, 0]);
    const hidden = workbook();
    hidden.sheets[0] = { ...hidden.sheets[0]!, rows: hidden.sheets[0]!.rows.map((row) => (row.index === 2 ? { ...row, hidden: true as const } : row)) };
    expect(refreshedChartContent(hidden, "Summary", chartOf(hidden)).visual.series[0]!.values).toEqual([10, 30]);
    const all = workbook(DATA, { ...REFS, plotVisibleOnly: false });
    all.sheets[0] = hidden.sheets[0]!;
    expect(refreshedChartContent(all, "Summary", chartOf(all)).visual.series[0]!.values).toEqual([10, 20, 30]);
    const same = workbook();
    expect(refreshedChartContent(same, "Summary", chartOf(same))).toBe(chartOf(same));
  });

  it("refreshSheetChart sets cachesDirty on a changed chart (not for display) and chartsAffectedBy finds only charts whose references hold a changed cell", () => {
    const model = workbook({ ...DATA, B3: 99 });
    expect(chartsAffectedBy(model, [cellKey("Data", 2, 1)])).toEqual([{ sheetId: "Summary", drawingId: "dr1" }]);
    expect(chartsAffectedBy(model, [cellKey("Data", 40, 7)])).toEqual([]);
    const display = workbook({ ...DATA, B3: 99 });
    expect(refreshSheetChart(display, "Summary", "dr1", { forDisplay: true })).toBe(true);
    expect(chartOf(display).cachesDirty).toBeUndefined();
    expect(refreshSheetChart(model, "Summary", "dr1")).toBe(true);
    expect(chartOf(model).visual.series[0]!.values).toEqual([10, 99, 30]);
    expect(chartOf(model).cachesDirty).toBe(true);
    expect(refreshSheetChart(model, "Summary", "dr1")).toBe(false);
  });

  it("refreshes a 12-series × 100-category chart after a one-cell edit within SPEC-052's 100 ms budget", () => {
    const values: Record<string, CellValue> = {};
    for (let row = 1; row <= 100; row++) {
      for (let col = 1; col <= 12; col++) {
        values[`${String.fromCharCode(65 + col)}${row}`] = row * col;
      }
    }
    const refs: SheetChartRefs = { series: Array.from({ length: 12 }, (_, i) => ({ values: `Data!$${String.fromCharCode(66 + i)}$1:$${String.fromCharCode(66 + i)}$100` })), plotVisibleOnly: true };
    const model = workbook(values, refs);
    refreshSheetChart(model, "Summary", "dr1");
    model.sheets[0] = { ...model.sheets[0]!, rows: model.sheets[0]!.rows.map((row) => (row.index === 5 ? { ...row, cells: row.cells.map((cell) => (cell.col === 1 ? { ...cell, value: 1e6 } : cell)) } : row)) };
    const start = performance.now();
    expect(chartsAffectedBy(model, [cellKey("Data", 5, 1)])).toHaveLength(1);
    expect(refreshSheetChart(model, "Summary", "dr1")).toBe(true);
    expect(performance.now() - start).toBeLessThan(100);
  });
});

describe("the charts adjusters", () => {
  it("a row insert above the chart's data moves its references, and the refreshed chart plots the same values", () => {
    const model = workbook();
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: "Data", at: 0, count: 1 })).toEqual({ ok: true });
    const content = chartOf(model);
    expect(content.refs!.series[0]).toEqual({ name: "Data!$B$2", categories: "Data!$A$3:$A$5", values: "Data!$B$3:$B$5" });
    expect(content.refsDirty).toBe(true);
    expect(content.visual.series[0]!.values).toEqual([10, 20, 30]);
    expect(content.visual.categories).toEqual(["Q1", "Q2", "Q3"]);
  });

  it("a row insert inside the data range extends the references, and the caches gain the new blank point", () => {
    const model = workbook();
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: "Data", at: 2, count: 1 })).toEqual({ ok: true });
    const content = chartOf(model);
    expect(content.refs!.series[0]).toEqual({ name: "Data!$B$1", categories: "Data!$A$2:$A$5", values: "Data!$B$2:$B$5" });
    expect(content.visual.series[0]!.values).toEqual([10, null, 20, 30]);
    expect(content.cachesDirty).toBe(true);
  });

  it("moveRange of the whole data block re-points the references; the chart drawing stays put", () => {
    const model = workbook();
    const anchor = model.sheets[1]!.drawings![0]!.anchor;
    expect(moveRange(model, { fromSheetId: "Data", from: { top: 0, left: 0, bottom: 3, right: 1 }, toSheetId: "Data", toTop: 10, toLeft: 3 })).toEqual({ ok: true });
    const content = chartOf(model);
    expect(content.refs!.series[0]).toEqual({ name: "Data!$E$11", categories: "Data!$D$12:$D$14", values: "Data!$E$12:$E$14" });
    expect(content.visual.series[0]!.values).toEqual([10, 20, 30]);
    expect(model.sheets[1]!.drawings![0]!.anchor).toBe(anchor);
  });

  it("a sheet rename renames the references; deleting the data's sheet makes them #REF! and the series keeps its caches", () => {
    const model = workbook();
    renameSheet(model, 0, "Facts");
    expect(chartOf(model).refs!.series[0]!.values).toBe("Facts!$B$2:$B$4");
    expect(chartOf(model).refsDirty).toBe(true);
    deleteSheet(model, 0);
    const content = chartOf(model);
    expect(content.refs!.series[0]!.values).toContain("#REF!");
    expect(content.visual.series[0]!.values).toEqual([10, 20, 30]);
  });
});

const select = (top: number, left: number, bottom: number, right: number) => ({ ranges: [{ top, left, bottom, right }], active: { row: top, col: left } });

describe("insert from selection", () => {
  const TABLE = { A1: "Item", B1: "Plan", C1: "Actual", A2: "Alpha", A3: "Beta", A4: "Gamma", B2: 4, B3: 6, B4: 8, C2: 5, C3: 5, C4: 9 };

  it("headers become names and categories, series run down the columns, and the chart lands two columns right of the data on its top row", () => {
    const model = bookOf([sheetOf("Summary", TABLE)]);
    const built = buildChartFromRange(model, "Summary", select(0, 0, 3, 2), "line");
    expect(built.ok && built.refs.series).toEqual([
      { name: "Summary!$B$1", categories: "Summary!$A$2:$A$4", values: "Summary!$B$2:$B$4" },
      { name: "Summary!$C$1", categories: "Summary!$A$2:$A$4", values: "Summary!$C$2:$C$4" },
    ]);
    const result = insertSheetChart(model, "Summary", select(0, 0, 3, 2), "line");
    expect("sheet" in result).toBe(true);
    if (!("sheet" in result)) {
      return;
    }
    applyDrawingResult(model, result);
    const drawing = model.sheets[0]!.drawings!.find((each) => each.id === result.drawingId)!;
    expect(drawing.anchor).toMatchObject({ kind: "twoCell", editAs: "twoCell", from: { col: 4, row: 0, colOffEmu: 0, rowOffEmu: 0 } });
    expect(drawing.content.kind === "chart" && drawing.content.visual.chartType).toBe("line");
    expect(drawing.content.kind === "chart" && drawing.content.visual.series.map((series) => series.values)).toEqual([
      [4, 6, 8],
      [5, 5, 9],
    ]);
  });

  it("a single cell expands to its current region; an empty selection with an empty region inserts nothing", () => {
    const model = bookOf([sheetOf("Summary", TABLE)]);
    const built = buildChartFromRange(model, "Summary", select(2, 1, 2, 1), "columnClustered");
    expect(built.ok && built.source.ranges).toEqual([{ top: 0, left: 0, bottom: 3, right: 2 }]);
    expect(insertSheetChart(model, "Summary", select(20, 20, 20, 20), "columnClustered")).toEqual({ message: NO_CHART_DATA_MESSAGE });
  });

  it("names the chart \"Chart {k}\" past every name the sheet's drawings already take", () => {
    // Imported "Chart 3" (cNvPr id 2) and "Chart 2" (id 3): the next drawing number is 3, a name "Chart 3" already takes.
    const imported: SheetDrawing = { ...chartDrawing(), id: "dr1", name: "Chart 3", sourceXml: `<xdr:twoCellAnchor><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 3"/></xdr:nvGraphicFramePr></xdr:graphicFrame></xdr:twoCellAnchor>` };
    const second: SheetDrawing = { ...imported, id: "dr2", name: "Chart 2", sourceXml: imported.sourceXml!.replace('id="2" name="Chart 3"', 'id="3" name="Chart 2"') };
    const model = bookOf([sheetOf("Summary", TABLE, { drawings: [imported, second] })]);
    const names: string[] = [];
    for (let i = 0; i < 2; i += 1) {
      const result = insertSheetChart(model, "Summary", select(0, 0, 3, 2), "line");
      if (!("sheet" in result)) {
        throw new Error(result.message);
      }
      applyDrawingResult(model, result);
      names.push(model.sheets[0]!.drawings!.find((each) => each.id === result.drawingId)!.name);
    }
    expect(names).toEqual(["Chart 4", "Chart 5"]);
  });

  it("with fewer data rows than columns, series run across the rows", () => {
    const model = bookOf([sheetOf("Wide", { A1: "", B1: "Jan", C1: "Feb", D1: "Mar", A2: "Sales", B2: 1, C2: 2, D2: 3 })]);
    const built = buildChartFromRange(model, "Wide", select(0, 0, 1, 3), "columnClustered");
    expect(built.ok && built.refs.series).toEqual([{ name: "Wide!$A$2", categories: "Wide!$B$1:$D$1", values: "Wide!$B$2:$D$2" }]);
    expect(built.ok && built.visual.options.title).toBe("Sales");
  });
});
