// SPEC-132 xlsx chart write planner (chartParts.ts) on spec132-charts.xlsx:
// the path each chart drawing takes — verbatim, patched, regenerated, a new
// part — and what each writes.
import { describe, expect, it } from "vitest";
import { CHART_CONTENT_TYPE } from "../../core/chartPartOoxml";
import { unzip } from "../../core/zip";
import { chartContentOf, chartsAffectedBy, insertSheetChart, refreshSheetChart, setSheetChartOptions } from "../charts";
import { applyDrawingResult } from "../drawings";
import { cellKey } from "../formula/dependencies";
import { buildSpec132ChartsXlsxBytes } from "../fixtures/build";
import type { SheetChartContent, SheetDrawing, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { sheetChartsChanged, sheetHasCreatedChart, XlsxChartWriter } from "./chartParts";
import { xlsxChartPartXml } from "./chartXml";

const decoder = new TextDecoder();

async function opened(): Promise<{ model: SpreadsheetModel; parts: Map<string, Uint8Array> }> {
  const bytes = await buildSpec132ChartsXlsxBytes();
  return { model: await readSpreadsheetBytes(bytes, "xlsx", "Charts"), parts: await unzip(bytes) };
}

const writerFor = (model: SpreadsheetModel, parts: Map<string, Uint8Array>): XlsxChartWriter => new XlsxChartWriter(model, parts, new Set(parts.keys()));
const sheetNamed = (model: SpreadsheetModel, name: string) => model.sheets.find((sheet) => sheet.name === name)!;
const drawingNamed = (model: SpreadsheetModel, sheet: string, name: string): SheetDrawing => sheetNamed(model, sheet).drawings!.find((drawing) => drawing.name === name)!;

/** A reference one row down ("Data!$B$1" → "Data!$B$2"). */
const moved = (text: string): string => text.replace(/\$(\d+)/g, (_, row: string) => `$${Number(row) + 1}`);
/** A chart part with its c:f texts blanked. */
const withoutFormulas = (xml: string): string => xml.replace(/<c:f>[^<]*<\/c:f>/g, "<c:f/>");

/** A chart drawing's content replaced copy-on-write. */
function setChart(model: SpreadsheetModel, sheet: string, name: string, update: (content: SheetChartContent) => SheetChartContent): void {
  model.sheets = model.sheets.map((each) =>
    each.name !== sheet ? each : { ...each, drawings: each.drawings!.map((drawing) => (drawing.name === name ? { ...drawing, content: update(chartContentOf(drawing)!) } : drawing)) },
  );
}

describe("XlsxChartWriter", () => {
  it("writes nothing for an untouched workbook; each chart drawing names the part its anchor relates", async () => {
    const { model, parts } = await opened();
    expect(model.sheets.map(sheetChartsChanged)).toEqual([false, false, false, false]);
    const writer = writerFor(model, parts);
    expect(writer.partFor(sheetNamed(model, "Data"), drawingNamed(model, "Data", "Chart 1"))).toBe("xl/charts/chart1.xml");
    expect(writer.partFor(sheetNamed(model, "Sales Chart"), drawingNamed(model, "Sales Chart", "Chart 6"))).toBe("xl/charts/chart6.xml");
    // The waterfall is SPEC-131's foreign drawing, not a chart.
    expect(writer.partFor(sheetNamed(model, "Data"), drawingNamed(model, "Data", "Chart 7"))).toBeUndefined();
    expect(writer.writes()).toEqual({ replaced: new Map(), added: [], overrides: [] });
  });

  it("patches a refreshed chart's part — its caches from the cells, every other byte kept — and only that part", async () => {
    const { model, parts } = await opened();
    const data = sheetNamed(model, "Data");
    model.sheets = model.sheets.map((sheet) =>
      sheet === data ? { ...sheet, rows: sheet.rows.map((row) => (row.index === 2 ? { ...row, cells: row.cells.map((cell) => (cell.col === 3 ? { ...cell, value: 90 } : cell)) } : row)) } : sheet,
    );
    for (const target of chartsAffectedBy(model, [cellKey(data.id, 2, 3)])) {
      refreshSheetChart(model, target.sheetId, target.drawingId);
    }
    expect(sheetChartsChanged(sheetNamed(model, "Data"))).toBe(true);
    const writes = writerFor(model, parts).writes();
    expect([...writes.replaced.keys()]).toEqual(["xl/charts/chart1.xml"]);
    const source = decoder.decode(parts.get("xl/charts/chart1.xml"));
    const patched = decoder.decode(writes.replaced.get("xl/charts/chart1.xml"));
    expect(patched).toContain(`<c:pt idx="1"><c:v>90</c:v></c:pt>`);
    expect(patched.replace(/<c:v>90<\/c:v>/, "<c:v>30</c:v>")).toBe(source);
  });

  it("regenerates an edited chart's part from its visual and references", async () => {
    const { model, parts } = await opened();
    const data = sheetNamed(model, "Data");
    const chart = drawingNamed(model, "Data", "Chart 1");
    const content = chartContentOf(chart)!;
    applyDrawingResult(model, setSheetChartOptions(model, data.id, chart.id, { ...content.visual.options, legend: "bottom" })!);
    const edited = chartContentOf(drawingNamed(model, "Data", "Chart 1"))!;
    expect(edited.edited).toBe(true);
    const writes = writerFor(model, parts).writes();
    expect(decoder.decode(writes.replaced.get("xl/charts/chart1.xml"))).toBe(xlsxChartPartXml(edited.visual, edited.refs));
    expect(writes.added).toEqual([]);
  });

  it("mints one new part (with its Override) for a created chart, named past every chart part the package holds", async () => {
    const { model, parts } = await opened();
    const summary = sheetNamed(model, "Summary");
    const result = insertSheetChart(model, summary.id, { ranges: [{ top: 0, left: 0, bottom: 3, right: 2 }], active: { row: 0, col: 0 } }, "line");
    if (!("sheet" in result)) {
      throw new Error(result.message);
    }
    applyDrawingResult(model, result);
    const created = sheetNamed(model, "Summary").drawings!.find((drawing) => drawing.id === result.drawingId)!;
    expect(sheetHasCreatedChart(sheetNamed(model, "Summary"))).toBe(true);
    const writer = writerFor(model, parts);
    expect(writer.partFor(sheetNamed(model, "Summary"), created)).toBe("xl/charts/chart7.xml");
    expect(writer.partFor(sheetNamed(model, "Summary"), created)).toBe("xl/charts/chart7.xml");
    const content = chartContentOf(created)!;
    const writes = writer.writes();
    expect(writes.replaced.size).toBe(0);
    expect(writes.added.map((part) => [part.name, decoder.decode(part.data)])).toEqual([["xl/charts/chart7.xml", xlsxChartPartXml(content.visual, content.refs)]]);
    expect(writes.overrides).toEqual([{ partName: "xl/charts/chart7.xml", contentType: CHART_CONTENT_TYPE }]);
  });

  it("never writes an out-of-catalogue chart's part, and patches only a read-only chart's c:f texts", async () => {
    const { model, parts } = await opened();
    setChart(model, "Data", "Chart 5", (content) => ({ ...content, refsDirty: true }));
    setChart(model, "Data", "Chart 3", (content) => ({
      ...content,
      readOnlyReason: "Pivot charts can't be edited yet",
      refsDirty: true,
      refs: { ...content.refs!, series: content.refs!.series.map((each) => ({ name: moved(each.name!), categories: moved(each.categories!), values: moved(each.values!) })) },
    }));
    const writes = writerFor(model, parts).writes();
    expect([...writes.replaced.keys()]).toEqual(["xl/charts/chart3.xml"]);
    const source = decoder.decode(parts.get("xl/charts/chart3.xml"));
    const patched = decoder.decode(writes.replaced.get("xl/charts/chart3.xml"));
    expect([...patched.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((match) => match[1])).toEqual(["Data!$B$2", "Data!$A$3:$A$6", "Data!$B$3:$B$6"]);
    expect(withoutFormulas(patched)).toBe(withoutFormulas(source));
  });
});
