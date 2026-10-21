// SPEC-132 round trips on spec132-charts.xlsx (and an .ods made from it): an
// untouched chart's part is written byte-identical, a chart whose cells
// changed is patched and reopens with the refreshed caches, a chart created
// here reopens as it was made (equalStructural: the same visual and
// references), and a second write of the reopened workbook is byte-stable.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { chartContentOf, chartsAffectedBy, insertSheetChart, refreshSheetChart } from "./charts";
import { applyDrawingResult } from "./drawings";
import { cellKey } from "./formula/dependencies";
import { buildSpec132ChartsXlsxBytes } from "./fixtures/build";
import type { SheetChartContent, SpreadsheetModel } from "./model";
import { writeOds } from "./ods/writer";
import { readSpreadsheetBytes } from "./open";
import { stripRetention } from "./sheets";
import { writeXlsx } from "./xlsx/writer";

async function opened(): Promise<{ bytes: Uint8Array; model: SpreadsheetModel }> {
  const bytes = await buildSpec132ChartsXlsxBytes();
  return { bytes, model: await readSpreadsheetBytes(bytes, "xlsx", "Charts") };
}

/** The chart drawing named `name` on sheet `sheet`. */
function chartNamed(model: SpreadsheetModel, sheet: string, name: string): SheetChartContent {
  const drawing = model.sheets.find((each) => each.name === sheet)!.drawings!.find((each) => each.name === name)!;
  return chartContentOf(drawing)!;
}

/** Data!D3 set to `value`, as the edit pipeline does: the cell replaced copy-on-write, the sheet dirty, the charts it feeds refreshed. */
function editD3(model: SpreadsheetModel, value: number): void {
  const data = model.sheets.find((sheet) => sheet.name === "Data")!;
  const rows = data.rows.map((row) => (row.index === 2 ? { ...row, cells: row.cells.map((cell) => (cell.col === 3 ? { ...cell, value } : cell)) } : row));
  model.sheets = model.sheets.map((sheet) => (sheet === data ? { ...sheet, rows, dirty: true as const } : sheet));
  for (const target of chartsAffectedBy(model, [cellKey(data.id, 2, 3)])) {
    refreshSheetChart(model, target.sheetId, target.drawingId);
  }
}

const structural = (content: SheetChartContent) => ({ visual: content.visual, refs: content.refs });

describe("SPEC-132 chart round trips — spec132-charts.xlsx", () => {
  it("opening and writing an untouched workbook writes every chart part (and its rels, style and colors parts) byte-identical", async () => {
    const { bytes, model } = await opened();
    const source = await unzip(bytes);
    const written = await unzip(await writeXlsx(model));
    const chartParts = [...source.keys()].filter((name) => name.startsWith("xl/charts/"));
    expect(chartParts.length).toBeGreaterThan(6);
    for (const name of chartParts) {
      expect(equalBytes(written.get(name)!, source.get(name)!), name).toBe(true);
    }
  });

  it("a changed cell patches the chart part it feeds, which reopens with the refreshed cache; untouched chart parts stay byte-identical; a second write is byte-stable", async () => {
    const { bytes, model } = await opened();
    editD3(model, 90);
    expect(chartNamed(model, "Data", "Chart 1").cachesDirty).toBe(true);
    const first = await writeXlsx(model);
    const source = await unzip(bytes);
    const written = await unzip(first);
    expect(new TextDecoder().decode(written.get("xl/charts/chart1.xml")!)).toContain("<c:v>90</c:v>");
    // The pie and the scatter do not plot D3: their parts pass through.
    for (const name of ["xl/charts/chart3.xml", "xl/charts/chart4.xml", "xl/charts/chartEx1.xml", "xl/charts/style1.xml", "xl/charts/colors1.xml"]) {
      expect(equalBytes(written.get(name)!, source.get(name)!), name).toBe(true);
    }
    const reopened = await readSpreadsheetBytes(first, "xlsx", "Charts");
    expect(chartNamed(reopened, "Data", "Chart 1").visual.series[2]!.values).toEqual([12, 90, 18, 24]);
    const second = await unzip(await writeXlsx(reopened));
    for (const name of [...written.keys()].filter((each) => each.startsWith("xl/charts/"))) {
      expect(equalBytes(second.get(name)!, written.get(name)!), name).toBe(true);
    }
  });

  it("a chart created here reopens with the visual and references it was made with, in .xlsx and in an .ods copy", async () => {
    const { model } = await opened();
    const summary = model.sheets.find((sheet) => sheet.name === "Summary")!;
    const result = insertSheetChart(model, summary.id, { ranges: [{ top: 0, left: 0, bottom: 3, right: 2 }], active: { row: 0, col: 0 } }, "line");
    if (!("sheet" in result)) {
      throw new Error(result.message);
    }
    applyDrawingResult(model, result);
    const made = chartContentOf(result.sheet.drawings!.find((drawing) => drawing.id === result.drawingId)!)!;

    const xlsx = await readSpreadsheetBytes(await writeXlsx(model), "xlsx", "Charts");
    const xlsxCharts = xlsx.sheets.find((sheet) => sheet.name === "Summary")!.drawings!.map((drawing) => chartContentOf(drawing)!);
    expect(xlsxCharts.map(structural)).toContainEqual(structural(made));

    const ods = await readSpreadsheetBytes(await writeOds(stripRetention({ ...model, format: "ods" })), "ods", "Charts");
    const odsCharts = ods.sheets.find((sheet) => sheet.name === "Summary")!.drawings!.flatMap((drawing) => {
      const content = chartContentOf(drawing);
      return content ? [content] : [];
    });
    // The series references cross; a generated ODF object writes no chart style, so
    // plotVisibleOnly reads back as ODF's default (hidden cells included).
    expect(odsCharts.map((content) => content.refs?.series)).toContainEqual(made.refs!.series);
    expect(odsCharts.map((content) => content.visual.series.map((series) => series.values))).toContainEqual(made.visual.series.map((series) => series.values));
  });
});
