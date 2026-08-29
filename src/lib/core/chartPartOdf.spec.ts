// SPEC-092: the hoisted ODF chart-object codec, exercised on its new
// `ChartVisual` surface with explicit geometry. The text-product re-export
// specs (`odf/chart.spec.ts`) pin the identical behaviour through
// `odf/chart.ts`; this pins the core module directly.
import { describe, it, expect } from "vitest";
import {
  CHART_NAMESPACE,
  CHART_OBJECT_CONTENT_TYPE,
  odfChartContentXml,
  parseOdfChartXml,
} from "./chartPartOdf";
import { CHART_CATALOGUE, SINGLE_SERIES_CHART_TYPES, type ChartVisual } from "./charts";

function baseVisual(overrides: Partial<ChartVisual> = {}): ChartVisual {
  return {
    chartType: "columnClustered",
    categories: ["Cat 1", "Cat 2", "Cat 3"],
    series: [
      { name: "Series A", values: [1, 2, 3], color: "#4e79a7" },
      { name: "Series B", values: [4, null, 6], color: "#f28e2b" },
    ],
    options: { legend: "right", dataLabels: false, gridlines: true },
    ...overrides,
  };
}

describe("odfChartContentXml / parseOdfChartXml — ChartVisual round trip", () => {
  for (const { chartType } of CHART_CATALOGUE) {
    it(`round-trips ${chartType}`, () => {
      const chart = baseVisual({ chartType });
      const parsed = parseOdfChartXml(odfChartContentXml(chart, 12, 7.5));
      expect(parsed).not.toBeNull();
      expect(parsed!.chartType).toBe(chartType);
      const expected = SINGLE_SERIES_CHART_TYPES.has(chartType) ? chart.series.slice(0, 1) : chart.series;
      expect(parsed!.series.map((s) => s.name)).toEqual(expected.map((s) => s.name));
      expect(parsed!.series.map((s) => s.values)).toEqual(expected.map((s) => s.values));
    });
  }

  it("writes the passed geometry onto chart:chart svg:width/height", () => {
    const xml = odfChartContentXml(baseVisual(), 8, 4.5);
    expect(xml).toContain('svg:width="8cm"');
    expect(xml).toContain('svg:height="4.5cm"');
  });

  it("preserves a gap (null) distinct from zero", () => {
    const parsed = parseOdfChartXml(odfChartContentXml(baseVisual({ series: [{ name: "S", values: [0, null, 5], color: "#111111" }] }), 12, 7.5));
    expect(parsed!.series[0].values).toEqual([0, null, 5]);
  });

  it("returns null for a chart:class outside the catalogue", () => {
    const bad = odfChartContentXml(baseVisual(), 12, 7.5).replace('chart:class="chart:bar"', 'chart:class="chart:surface"');
    expect(parseOdfChartXml(bad)).toBeNull();
  });

  it("returns null for malformed XML", () => {
    expect(parseOdfChartXml("<not-xml")).toBeNull();
  });
});

describe("constants", () => {
  it("expose the ODF chart namespace and object content type", () => {
    expect(CHART_NAMESPACE).toBe("urn:oasis:names:tc:opendocument:xmlns:chart:1.0");
    expect(CHART_OBJECT_CONTENT_TYPE).toBe("application/vnd.oasis.opendocument.chart");
  });
});
