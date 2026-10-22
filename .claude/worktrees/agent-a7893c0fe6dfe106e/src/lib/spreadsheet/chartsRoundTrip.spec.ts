// SPEC-132 round trips on spec132-charts.xlsx and spec132-charts.ods: an
// untouched chart's part is written byte-identical, a chart whose cells
// changed is patched in place and reopens with the refreshed caches, moved
// references reach the patched part (and, for an unmodelled chart, the
// retained rewrite) and nothing else, a chart created here reopens as it was
// made (equalStructural: the same visual — its colours resolved, ODF having
// no theme — and references) in both formats, and a second write of the
// reopened workbook is byte-stable.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { DEFAULT_THEME, resolveColor } from "../core/theme";
import { unzip } from "../core/zip";
import { chartContentOf, chartsAffectedBy, insertSheetChart, refreshSheetChart } from "./charts";
import { applyDrawingResult } from "./drawings";
import { cellKey } from "./formula/dependencies";
import { buildSpec132ChartsOdsBytes, buildSpec132ChartsXlsxBytes, SPEC132_CHARTS_ODS } from "./fixtures/build";
import type { SheetChartContent, SpreadsheetModel } from "./model";
import { writeOds } from "./ods/writer";
import { readSpreadsheetBytes } from "./open";
import { stripRetention } from "./sheets";
import { applyStructuralEdit } from "./structure";
import { writeXlsx } from "./xlsx/writer";

async function opened(): Promise<{ bytes: Uint8Array; model: SpreadsheetModel }> {
  const bytes = await buildSpec132ChartsXlsxBytes();
  return { bytes, model: await readSpreadsheetBytes(bytes, "xlsx", "Charts") };
}

async function openedOds(): Promise<{ bytes: Uint8Array; model: SpreadsheetModel }> {
  const bytes = await buildSpec132ChartsOdsBytes();
  return { bytes, model: await readSpreadsheetBytes(bytes, "ods", "Charts") };
}

const decoder = new TextDecoder();
const text = (parts: Map<string, Uint8Array>, name: string): string => decoder.decode(parts.get(name)!);

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

/** A chart's visual and references, its series colours resolved through
 * the workbook's theme (ODF stores no theme references, only the colour
 * they resolved to). */
const structural = (content: SheetChartContent, model: SpreadsheetModel) => ({
  visual: { ...content.visual, series: content.visual.series.map((series) => ({ ...series, color: resolveColor(series.color, model.theme ?? DEFAULT_THEME) })) },
  refs: content.refs,
});

/** Its references as .xlsx can state them: no plot-area range (an ODS-only reference). */
const withoutPlotArea = (content: SheetChartContent): SheetChartContent => {
  const { plotArea: _plotArea, ...refs } = content.refs!;
  return { ...content, refs };
};

/** Every chart content on a sheet. */
function chartsOn(model: SpreadsheetModel, sheet: string): SheetChartContent[] {
  return model.sheets
    .find((each) => each.name === sheet)!
    .drawings!.flatMap((drawing) => {
      const content = chartContentOf(drawing);
      return content ? [content] : [];
    });
}

/** Insert ▸ Chart ▸ Line over Summary!A1:C4; the content made. */
function insertLineOnSummary(model: SpreadsheetModel): SheetChartContent {
  const summary = model.sheets.find((sheet) => sheet.name === "Summary")!;
  const result = insertSheetChart(model, summary.id, { ranges: [{ top: 0, left: 0, bottom: 3, right: 2 }], active: { row: 0, col: 0 } }, "line");
  if (!("sheet" in result)) {
    throw new Error(result.message);
  }
  applyDrawingResult(model, result);
  return chartContentOf(result.sheet.drawings!.find((drawing) => drawing.id === result.drawingId)!)!;
}

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
    const made = insertLineOnSummary(model);
    expect(made.refs!.plotVisibleOnly).toBe(true);

    const xlsx = await readSpreadsheetBytes(await writeXlsx(model), "xlsx", "Charts");
    expect(chartsOn(xlsx, "Summary").map((content) => structural(content, xlsx))).toContainEqual(structural(made, model));

    // Across formats too — plotVisibleOnly included (the object's visible-cells-only style).
    const ods = await readSpreadsheetBytes(await writeOds(stripRetention({ ...model, format: "ods" })), "ods", "Charts");
    expect(chartsOn(ods, "Summary").map((content) => structural(content, ods))).toContainEqual(structural(made, model));
  });
});

/** An ODS chart object's address attribute values, in document order. */
const addresses = (xml: string): string[] =>
  [...xml.matchAll(/(?:table:cell-range-address|chart:values-cell-range-address|chart:label-cell-address)="([^"]*)"/g)].map((match) => match[1]!);
/** The object with its address attributes blanked: what stays when only its references move. */
const withoutAddresses = (xml: string): string => xml.replace(/(table:cell-range-address|chart:values-cell-range-address|chart:label-cell-address)="[^"]*"/g, '$1=""');
/** The object up to its `local-table` cache. */
const beforeLocalTable = (xml: string): string => xml.slice(0, xml.indexOf(`<table:table table:name="local-table">`));

describe("SPEC-132 chart round trips — spec132-charts.ods", () => {
  const { bar, surface } = SPEC132_CHARTS_ODS.charts;

  it("reads the bar chart's caches and references, the surface chart as unsupported; an untouched workbook is written byte-identical, every object part included", async () => {
    const { bytes, model } = await openedOds();
    const barChart = chartNamed(model, "Data", bar.name);
    expect(barChart.sourcePart).toBe(bar.part);
    expect(barChart.refs).toEqual(bar.refs);
    expect(barChart.visual.chartType).toBe(bar.chartType);
    expect(barChart.visual.series.map((series) => series.values)).toEqual([
      [10, 20, 30, 40],
      [15, 25, 35, 45],
      [12, 30, 18, 24],
    ]);
    const surfaceChart = chartNamed(model, "Data", surface.name);
    expect(surfaceChart.visual.unsupportedChartXml).toBe(text(await unzip(bytes), surface.part));
    expect(surfaceChart.refs).toBeUndefined();

    const source = await unzip(bytes);
    const written = await unzip(await writeOds(model));
    expect([...written.keys()].toSorted()).toEqual([...source.keys()].toSorted());
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), name).toBe(true);
    }
  });

  it("a changed cell patches the object it feeds — its local-table refreshed, every other byte kept — which reopens with the refreshed cache; the surface chart's object passes through; a second write is byte-stable", async () => {
    const { bytes, model } = await openedOds();
    editD3(model, 90);
    expect(chartNamed(model, "Data", bar.name).cachesDirty).toBe(true);
    const first = await writeOds(model);
    const source = await unzip(bytes);
    const written = await unzip(first);
    const object = text(written, bar.part);
    expect(object).toContain(`<table:table-cell office:value-type="float" office:value="90"><text:p>90</text:p></table:table-cell>`);
    expect(beforeLocalTable(object)).toBe(beforeLocalTable(text(source, bar.part)));
    for (const name of [surface.part, `${bar.object}/styles.xml`, `${bar.object}/meta.xml`, bar.replacement, surface.replacement, "META-INF/manifest.xml"]) {
      expect(equalBytes(written.get(name)!, source.get(name)!), name).toBe(true);
    }
    const reopened = await readSpreadsheetBytes(first, "ods", "Charts");
    expect(chartNamed(reopened, "Data", bar.name).visual.series[2]!.values).toEqual([12, 90, 18, 24]);
    const second = await unzip(await writeOds(reopened));
    for (const name of [...written.keys()].filter((each) => each.startsWith("Object"))) {
      expect(equalBytes(second.get(name)!, written.get(name)!), name).toBe(true);
    }
  });

  it("a row inserted above the data moves the bar chart's addresses and its frame's notify list from the model, and the surface chart's through the retained rewrite — nothing else in either object changes", async () => {
    const { bytes, model } = await openedOds();
    const data = model.sheets.find((sheet) => sheet.name === "Data")!;
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: data.id, at: 0, count: 1 })).toEqual({ ok: true });
    const barChart = chartNamed(model, "Data", bar.name);
    expect(barChart.refsDirty).toBe(true);
    expect(barChart.cachesDirty).toBeUndefined();
    const source = await unzip(bytes);
    const written = await unzip(await writeOds(model));
    const object = text(written, bar.part);
    expect(addresses(object)).toEqual(["$Data.A2:.D6", "$Data.A3:.A6", "$Data.B3:.B6", "$Data.B2", "$Data.C3:.C6", "$Data.C2", "$Data.D3:.D6", "$Data.D2"]);
    expect(withoutAddresses(object)).toBe(withoutAddresses(text(source, bar.part)));
    const surfaceObject = text(written, surface.part);
    expect(addresses(surfaceObject)).toEqual(["Data.A2:Data.D6", "Data.A3:Data.A6", "Data.B3:Data.B6", "Data.B2", "Data.C3:Data.C6", "Data.C2", "Data.D3:Data.D6", "Data.D2"]);
    expect(withoutAddresses(surfaceObject)).toBe(withoutAddresses(text(source, surface.part)));
    const content = text(written, "content.xml");
    expect(content).toContain(`<draw:object draw:notify-on-update-of-ranges="$Data.A2:.D6" xlink:href="./${bar.object}"`);
    expect(content).toContain(`<draw:object draw:notify-on-update-of-ranges="Data.A2:Data.D6" xlink:href="./${surface.object}"`);

    const reopened = await readSpreadsheetBytes(await writeOds(model), "ods", "Charts");
    expect(chartNamed(reopened, "Data", bar.name).refs!.series[0]).toEqual({ name: "Data!B2", categories: "Data!A3:A6", values: "Data!B3:B6" });
    expect(chartNamed(reopened, "Data", bar.name).visual.series.map((series) => series.values)).toEqual(barChart.visual.series.map((series) => series.values));
  });

  it("a chart created here reopens with the visual and references it was made with, in .ods (its own new object) and in an .xlsx copy", async () => {
    const { bytes, model } = await openedOds();
    const made = insertLineOnSummary(model);
    const written = await writeOds(model);
    const parts = await unzip(written);
    const source = await unzip(bytes);
    expect([...parts.keys()].filter((name) => !source.has(name))).toEqual(["Object 3/content.xml"]);
    expect(made.refs!.plotArea).toBe("Summary!$A$1:$C$4");
    const ods = await readSpreadsheetBytes(written, "ods", "Charts");
    expect(chartsOn(ods, "Summary").map((content) => structural(content, ods))).toContainEqual(structural(made, model));
    // A second write of the reopened workbook is byte-stable.
    const second = await unzip(await writeOds(ods));
    for (const [name, data] of parts) {
      expect(equalBytes(second.get(name)!, data), name).toBe(true);
    }

    const xlsx = await readSpreadsheetBytes(await writeXlsx(stripRetention({ ...model, format: "xlsx" })), "xlsx", "Charts");
    expect(chartsOn(xlsx, "Summary").map((content) => structural(content, xlsx))).toContainEqual(structural(withoutPlotArea(made), model));
    // The bar chart crosses too; the surface chart (out of the catalogue) does not.
    expect(chartsOn(xlsx, "Data").map((content) => content.visual.chartType)).toEqual(["columnClustered"]);
  });
});
