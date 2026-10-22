// SPEC-132 ods chart write planner (chartParts.ts) on spec132-charts.ods:
// the path each chart drawing's embedded object takes — passed through,
// patched, regenerated, a new object — and what each writes.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { chartContentOf, chartsAffectedBy, insertSheetChart, refreshSheetChart, setSheetChartOptions } from "../charts";
import { applyDrawingResult } from "../drawings";
import { cellKey } from "../formula/dependencies";
import { buildSpec132ChartsOdsBytes, SPEC132_CHARTS_ODS } from "../fixtures/build";
import type { SheetChartContent, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { hasOdsChartWork, planOdsCharts } from "./chartParts";
import { odsChartManifestEntries } from "./chartXml";

const decoder = new TextDecoder();
const { bar, surface } = SPEC132_CHARTS_ODS.charts;

async function opened(): Promise<{ model: SpreadsheetModel; parts: Map<string, Uint8Array> }> {
  const bytes = await buildSpec132ChartsOdsBytes();
  return { model: await readSpreadsheetBytes(bytes, "ods", "Charts"), parts: await unzip(bytes) };
}

const sheetNamed = (model: SpreadsheetModel, name: string) => model.sheets.find((sheet) => sheet.name === name)!;
const chartNamed = (model: SpreadsheetModel, sheet: string, name: string): SheetChartContent => chartContentOf(sheetNamed(model, sheet).drawings!.find((drawing) => drawing.name === name)!)!;

/** An object up to its `local-table` cache. */
const head = (xml: string): string => xml.slice(0, xml.indexOf(`<table:table table:name="local-table">`));

/** A chart drawing's content replaced copy-on-write. */
function setChart(model: SpreadsheetModel, sheet: string, name: string, update: (content: SheetChartContent) => SheetChartContent): void {
  model.sheets = model.sheets.map((each) =>
    each.name !== sheet ? each : { ...each, drawings: each.drawings!.map((drawing) => (drawing.name === name ? { ...drawing, content: update(chartContentOf(drawing)!) } : drawing)) },
  );
}

describe("planOdsCharts", () => {
  it("has no work for an untouched workbook: every object passes through", async () => {
    const { model, parts } = await opened();
    expect(hasOdsChartWork(model)).toBe(false);
    const plan = planOdsCharts(model, parts.keys());
    expect(plan.minted.size).toBe(0);
    expect(plan.newParts).toEqual([]);
    expect(plan.manifestEntries).toEqual([]);
    expect(plan.rewrite(bar.part, decoder.decode(parts.get(bar.part)))).toBeUndefined();
    expect(plan.rewrite(surface.part, decoder.decode(parts.get(surface.part)))).toBeUndefined();
  });

  it("patches a refreshed chart's object over the text it is given — its local-table refreshed, its addresses kept", async () => {
    const { model, parts } = await opened();
    const data = sheetNamed(model, "Data");
    model.sheets = model.sheets.map((sheet) =>
      sheet === data ? { ...sheet, rows: sheet.rows.map((row) => (row.index === 2 ? { ...row, cells: row.cells.map((cell) => (cell.col === 3 ? { ...cell, value: 90 } : cell)) } : row)) } : sheet,
    );
    for (const target of chartsAffectedBy(model, [cellKey(data.id, 2, 3)])) {
      refreshSheetChart(model, target.sheetId, target.drawingId);
    }
    expect(hasOdsChartWork(model)).toBe(true);
    const plan = planOdsCharts(model, parts.keys());
    const source = decoder.decode(parts.get(bar.part));
    const patched = plan.rewrite(bar.part, source)!;
    expect(patched).toContain(`office:value="90"`);
    expect(head(patched)).toBe(head(source));
    expect(plan.rewrite(surface.part, decoder.decode(parts.get(surface.part)))).toBeUndefined();
  });

  it("regenerates an edited chart's object from the model, whatever text it is given", async () => {
    const { model, parts } = await opened();
    const data = sheetNamed(model, "Data");
    const drawing = data.drawings!.find((each) => each.name === bar.name)!;
    applyDrawingResult(model, setSheetChartOptions(model, data.id, drawing.id, { ...chartNamed(model, "Data", bar.name).visual.options, legend: "bottom" })!);
    const plan = planOdsCharts(model, parts.keys());
    const regenerated = plan.rewrite(bar.part, decoder.decode(parts.get(bar.part)))!;
    expect(regenerated).toContain(`<chart:legend chart:legend-position="bottom"/>`);
    expect(regenerated).toContain(`<chart:plot-area chart:style-name="ch1" table:cell-range-address="$Data.A1:.D5">`);
    expect(plan.rewrite(bar.part, "anything")).toBe(regenerated);
  });

  it("mints a new object for a created chart, n past every object the package holds, with its parts and manifest entries", async () => {
    const { model, parts } = await opened();
    const summary = sheetNamed(model, "Summary");
    const result = insertSheetChart(model, summary.id, { ranges: [{ top: 0, left: 0, bottom: 3, right: 2 }], active: { row: 0, col: 0 } }, "line");
    if (!("sheet" in result)) {
      throw new Error(result.message);
    }
    applyDrawingResult(model, result);
    expect(hasOdsChartWork(model)).toBe(true);
    const created = sheetNamed(model, "Summary").drawings!.find((drawing) => drawing.id === result.drawingId)!;
    const plan = planOdsCharts(model, parts.keys());
    expect(plan.minted.get(created)).toBe("Object 3");
    expect(plan.newParts.map((part) => part.name)).toEqual(["Object 3/content.xml"]);
    expect(plan.newParts[0]!.text).toContain(`chart:class="chart:line"`);
    expect(plan.manifestEntries).toEqual([odsChartManifestEntries("Object 3")]);
  });

  it("never rewrites an out-of-catalogue chart's object", async () => {
    const { model, parts } = await opened();
    setChart(model, "Data", surface.name, (content) => ({ ...content, refsDirty: true, edited: true }));
    expect(hasOdsChartWork(model)).toBe(false);
    expect(planOdsCharts(model, parts.keys()).rewrite(surface.part, decoder.decode(parts.get(surface.part)))).toBeUndefined();
  });
});
