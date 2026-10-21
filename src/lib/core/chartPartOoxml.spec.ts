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
  type ChartPartRefs,
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

// SPEC-132 AC-18: the optional `refs` parameter — a sheet chart's own `c:f`
// texts in place of SPEC-052's placeholder formulas. Without refs the output
// is unchanged (the pptx golden and both products' round trips pin it).
describe("chartPartXml(chart, refs) — the host's references (SPEC-132)", () => {
  const refs: ChartPartRefs = {
    series: [
      { name: "Data!$B$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" },
      { name: "'R&D'!$C$1", categories: "Data!$A$2:$A$4", values: "'R&D'!$C$2:$C$4" },
    ],
  };

  it("writes each series' wire c:f texts, escaped, with the visual's caches — never the placeholders", () => {
    const xml = chartPartXml(baseVisual(), refs);
    expect(xml).toContain(
      '<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>',
    );
    expect(xml).toContain("<c:tx><c:strRef><c:f>'R&amp;D'!$C$1</c:f>");
    expect(xml).toContain(
      '<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt>',
    );
    expect(xml).toContain(
      "<c:val><c:numRef><c:f>'R&amp;D'!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode>" +
        '<c:ptCount val="3"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:val>',
    );
    expect(xml).not.toContain("Sheet1!");
    expect(xml).toContain('<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>');
  });

  it("writes a literal c:tx/c:v for a series with no name reference", () => {
    const xml = chartPartXml(baseVisual(), { series: [{ values: "Data!$B$2:$B$4" }, { name: "Data!$C$1", values: "Data!$C$2:$C$4" }] });
    expect(xml).toContain('<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:spPr>');
    expect(xml).toContain("<c:tx><c:strRef><c:f>Data!$C$1</c:f>");
    expect(parseChartXml(xml)!.series.map((s) => s.name)).toEqual(["Revenue", "Cost"]);
  });

  it("escapes a literal name", () => {
    const visual = baseVisual({ series: [{ name: "P&L <net>", values: [1, 2, 3], color: "#4e79a7" }] });
    expect(chartPartXml(visual, { series: [{ values: "Data!$B$2:$B$4" }] })).toContain("<c:tx><c:v>P&amp;L &lt;net&gt;</c:v></c:tx>");
  });

  it("writes no c:cat or c:xVal for a series with no categories reference", () => {
    for (const chartType of ["columnClustered", "line", "area", "pie", "scatter"] as const) {
      const xml = chartPartXml(baseVisual({ chartType }), { series: [{ name: "Data!$B$1", values: "Data!$B$2:$B$4" }, { values: "Data!$C$2:$C$4" }] });
      expect(xml).not.toContain("<c:cat>");
      expect(xml).not.toContain("<c:xVal>");
      expect(xml).not.toContain("<c:strLit>");
    }
  });

  it("writes a c:numLit holding the values for a series with no values reference", () => {
    const xml = chartPartXml(baseVisual(), { series: [{ name: "Data!$B$1" }, {}] });
    expect(xml).toContain(
      '<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/>' +
        '<c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numLit></c:val>',
    );
    expect(xml).toContain('<c:tx><c:v>Cost</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="F28E2B"/></a:solidFill></c:spPr><c:val><c:numLit>');
    expect(xml).not.toContain("<c:numRef>");
  });

  it("treats a series the refs do not reach as having no references", () => {
    const xml = chartPartXml(baseVisual(), { series: [refs.series[0]] });
    expect(xml).toContain('<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Cost</c:v></c:tx>');
  });

  it("writes a scatter's references as c:xVal and c:yVal number references, in CT_ScatterSer order", () => {
    const visual = baseVisual({ chartType: "scatter", categories: ["1", "2", "4"] });
    const xml = chartPartXml(visual, refs);
    expect(xml).toContain(
      '<c:xVal><c:numRef><c:f>Data!$A$2:$A$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>' +
        '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:xVal>' +
        "<c:yVal><c:numRef><c:f>Data!$B$2:$B$4</c:f>",
    );
    expect(xml).toContain('</c:numRef></c:yVal><c:smooth val="0"/></c:ser>');
    const parsed = parseChartXml(xml)!;
    expect(parsed.categories).toEqual(["1", "2", "4"]);
    expect(parsed.series.map((s) => s.values)).toEqual([
      [10, 20, 30],
      [4, null, 6],
    ]);
  });

  it("writes plotVisOnly 0 when the host plots hidden cells too", () => {
    expect(chartPartXml(baseVisual(), { ...refs, plotVisibleOnly: false })).toContain('<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/>');
    expect(chartPartXml(baseVisual(), { ...refs, plotVisibleOnly: true })).toContain('<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>');
  });

  it("writes only the first series' references for pie and doughnut", () => {
    for (const chartType of ["pie", "doughnut"] as const) {
      const xml = chartPartXml(baseVisual({ chartType }), refs);
      expect(xml.match(/<c:ser>/g)).toHaveLength(1);
      expect(xml).toContain("<c:f>Data!$B$2:$B$4</c:f>");
      expect(xml).not.toContain("$C$");
    }
  });

  it("keeps the part without refs exactly as before: placeholder formulas, no dispBlanksAs", () => {
    const xml = chartPartXml(baseVisual());
    expect(xml).toContain("<c:f>Sheet1!$B$1</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$A$2:$A$4</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$C$2:$C$4</c:f>");
    expect(xml).toContain('<c:plotVisOnly val="1"/></c:chart>');
    expect(xml).not.toContain("dispBlanksAs");
  });

  for (const { chartType } of CHART_CATALOGUE) {
    it(`round-trips ${chartType} written with refs through parseChartXml`, () => {
      const chart = baseVisual({ chartType, categories: chartType === "scatter" ? ["1", "2", "3"] : ["Q1", "Q2", "Q3"] });
      const parsed = parseChartXml(chartPartXml(chart, refs));
      expect(parsed).not.toBeNull();
      const expected = SINGLE_SERIES_CHART_TYPES.has(chartType) ? chart.series.slice(0, 1) : chart.series;
      expect(parsed!.chartType).toBe(chartType);
      expect(parsed!.categories).toEqual(chart.categories);
      expect(parsed!.series).toEqual(expected);
      // A pie or doughnut has no value axis, so no gridlines to read back.
      expect(parsed!.options).toEqual(SINGLE_SERIES_CHART_TYPES.has(chartType) ? { ...chart.options, gridlines: false } : chart.options);
    });
  }
});
