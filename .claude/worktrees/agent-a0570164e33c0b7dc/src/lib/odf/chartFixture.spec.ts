// SPEC-052: opening the fixtures — one chart of each catalogue type, one
// with negative values, one with a gap, one unsupported type as the
// placeholder. lo-chart-types.odt is hand-built by
// scripts/build-spec052-fixtures.mjs, NOT round-tripped through this
// project's own writer, so this is a genuine test of the reader
// (src/lib/odf/chart.ts) against real-shaped chart XML.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseOdt } from "./reader";
import { asChart, isChart } from "../documentModel";
import { type ChartType } from "../core/charts";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

describe("lo-chart-types.odt (AC-1, AC-40..43)", () => {
  it("renders one chart of each catalogue type, in document order", async () => {
    const fixture = loadFixture("lo-chart-types.odt");
    const parsed = await parseOdt(fixture);
    const charts = parsed.sections[0].blocks.filter(isChart).map(asChart);
    expect(charts).toHaveLength(12);

    const expectedTypes: ChartType[] = [
      "columnClustered",
      "columnStacked",
      "barClustered",
      "line",
      "lineMarkers",
      "area",
      "pie",
      "doughnut",
      "scatter",
      "columnClustered",
      "columnClustered",
    ];
    expectedTypes.forEach((type, i) => {
      expect(charts[i].chartType, `chart ${i}`).toBe(type);
      expect(charts[i].unsupportedChartXml, `chart ${i} should be supported`).toBeUndefined();
    });
  });

  it("renders a chart with negative values", async () => {
    const parsed = await parseOdt(loadFixture("lo-chart-types.odt"));
    const charts = parsed.sections[0].blocks.filter(isChart).map(asChart);
    expect(charts[9].series[0].values).toEqual([5, -3, 2]);
  });

  it("renders a chart with a gap (empty cell), distinct from zero", async () => {
    const parsed = await parseOdt(loadFixture("lo-chart-types.odt"));
    const charts = parsed.sections[0].blocks.filter(isChart).map(asChart);
    expect(charts[10].series[0].values).toEqual([10, null, 15]);
  });

  it("renders the unsupported (chart:stock) chart as a placeholder, not a catalogue type", async () => {
    const parsed = await parseOdt(loadFixture("lo-chart-types.odt"));
    const charts = parsed.sections[0].blocks.filter(isChart).map(asChart);
    const unsupported = charts[11];
    expect(unsupported.unsupportedChartXml).toBeDefined();
    expect(unsupported.unsupportedChartXml).toContain("chart:stock");
    expect(unsupported.sourcePart).toBe("Object 12/content.xml");
  });

  it("pie and doughnut carry only one series (the catalogue's own single-series rule)", async () => {
    const parsed = await parseOdt(loadFixture("lo-chart-types.odt"));
    const charts = parsed.sections[0].blocks.filter(isChart).map(asChart);
    expect(charts[6].series).toHaveLength(1); // pie
    expect(charts[7].series).toHaveLength(1); // doughnut
  });
});
