// SPEC-092: the hoisted OOXML chart-part codec, exercised on its new
// `ChartVisual` surface (no `Chart`/`DrawingPlacement` geometry). The
// text-product re-export specs (`docx/chart.spec.ts`) pin the identical
// behaviour through `docx/chart.ts`; this pins the core module directly so
// the presentation product can rely on it.
import { describe, it, expect } from "vitest";
import {
  CHART_CONTENT_TYPE,
  CHART_RELATIONSHIP,
  C_NAMESPACE,
  chartPartXml,
  parseChartXml,
} from "./chartPartOoxml";
import { CHART_CATALOGUE, SINGLE_SERIES_CHART_TYPES, type ChartVisual } from "./charts";

function baseVisual(overrides: Partial<ChartVisual> = {}): ChartVisual {
  return {
    chartType: "columnClustered",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "Revenue", values: [10, 20, 30], color: "#4e79a7" },
      { name: "Cost", values: [4, null, 6], color: "#f28e2b" },
    ],
    options: { legend: "right", dataLabels: false, gridlines: true },
    ...overrides,
  };
}

describe("chartPartXml / parseChartXml — ChartVisual round trip", () => {
  for (const { chartType } of CHART_CATALOGUE) {
    it(`round-trips ${chartType}`, () => {
      const chart = baseVisual({ chartType });
      const parsed = parseChartXml(chartPartXml(chart));
      expect(parsed).not.toBeNull();
      expect(parsed!.chartType).toBe(chartType);
      const expected = SINGLE_SERIES_CHART_TYPES.has(chartType) ? chart.series.slice(0, 1) : chart.series;
      expect(parsed!.series.map((s) => s.name)).toEqual(expected.map((s) => s.name));
      expect(parsed!.series.map((s) => s.values)).toEqual(expected.map((s) => s.values));
      expect(parsed!.series.map((s) => s.color)).toEqual(expected.map((s) => s.color));
    });
  }

  it("preserves a gap (null) distinct from zero", () => {
    const parsed = parseChartXml(chartPartXml(baseVisual({ series: [{ name: "S", values: [0, null, 5], color: "#111111" }] })));
    expect(parsed!.series[0].values).toEqual([0, null, 5]);
  });

  it("round-trips options: title, axis titles, legend, labels, gridlines, min/max", () => {
    const options: ChartVisual["options"] = {
      title: "Quarterly Revenue",
      categoryAxisTitle: "Quarter",
      valueAxisTitle: "USD",
      legend: "bottom",
      dataLabels: true,
      gridlines: true,
      valueAxisMin: -5,
      valueAxisMax: 40,
    };
    const parsed = parseChartXml(chartPartXml(baseVisual({ options })));
    expect(parsed!.options).toEqual(options);
  });

  it("returns null for a combo chart (two type elements)", () => {
    const combo = chartPartXml(baseVisual()).replace("</c:barChart>", "</c:barChart><c:lineChart/>");
    expect(parseChartXml(combo)).toBeNull();
  });

  it("returns null for malformed XML", () => {
    expect(parseChartXml("<not-xml")).toBeNull();
  });
});

describe("constants", () => {
  it("expose the OOXML chart namespace, relationship and content type", () => {
    expect(C_NAMESPACE).toBe("http://schemas.openxmlformats.org/drawingml/2006/chart");
    expect(CHART_RELATIONSHIP).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
    expect(CHART_CONTENT_TYPE).toBe("application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
  });
});
