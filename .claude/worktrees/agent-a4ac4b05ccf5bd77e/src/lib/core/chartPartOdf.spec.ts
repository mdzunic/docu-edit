// SPEC-092: the hoisted ODF chart-object codec, exercised on its new
// `ChartVisual` surface with explicit geometry. The text-product re-export
// specs (`odf/chart.spec.ts`) pin the identical behaviour through
// `odf/chart.ts`; this pins the core module directly.
import { describe, it, expect } from "vitest";
import {
  CHART_NAMESPACE,
  CHART_OBJECT_CONTENT_TYPE,
  odfChartContentXml,
  odfChartTableXml,
  parseOdfChartXml,
  type OdfChartRefs,
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

describe("odfChartContentXml — a host's references (SPEC-132)", () => {
  const refs: OdfChartRefs = {
    plotArea: "$Data.$A$1:.$C$4",
    categories: "$Data.$A$2:.$A$4",
    series: [
      { values: "$Data.$B$2:.$B$4", label: "$Data.$B$1" },
      { values: "$Data.$C$2:.$C$4", label: "$Data.$C$1" },
    ],
  };

  it("writes each series' values and label address in place of the local-table ones", () => {
    const xml = odfChartContentXml(baseVisual(), 12, 7.5, undefined, refs);
    expect(xml).toContain(
      '<chart:series chart:class="chart:bar" chart:color="#4e79a7" chart:values-cell-range-address="$Data.$B$2:.$B$4" chart:label-cell-address="$Data.$B$1"/>',
    );
    expect(xml).toContain(
      '<chart:series chart:class="chart:bar" chart:color="#f28e2b" chart:values-cell-range-address="$Data.$C$2:.$C$4" chart:label-cell-address="$Data.$C$1"/>',
    );
    expect(xml).not.toContain("local-table.$");
  });

  it("an absent entry keeps the local-table address the object writes without refs", () => {
    const xml = odfChartContentXml(baseVisual(), 12, 7.5, undefined, { series: [{ values: "$Data.$B$2:.$B$4" }] });
    expect(xml).toContain('chart:values-cell-range-address="$Data.$B$2:.$B$4" chart:label-cell-address="local-table.$B$1"/>');
    expect(xml).toContain('chart:values-cell-range-address="local-table.$C$2:.$C$4" chart:label-cell-address="local-table.$C$1"/>');
    expect(xml).not.toContain("chart:categories");
    expect(xml).toContain("<chart:plot-area>");
  });

  it("puts the plot-area range on chart:plot-area and the categories on the x axis — after its title, before the y axis", () => {
    const xml = odfChartContentXml(baseVisual({ options: { legend: "none", dataLabels: true, gridlines: true, categoryAxisTitle: "Month" } }), 12, 7.5, undefined, refs);
    expect(xml).toContain('<chart:plot-area table:cell-range-address="$Data.$A$1:.$C$4" chart:data-labels="true">');
    expect(xml).toContain(
      '<chart:axis chart:dimension="x" chart:name="primary-x"><chart:title><text:p>Month</text:p></chart:title>' +
        '<chart:categories table:cell-range-address="$Data.$A$2:.$A$4"/></chart:axis><chart:axis chart:dimension="y"',
    );
  });

  it("gives a pie (and a doughnut) an x axis carrying just the categories, and only when there are some", () => {
    for (const chartType of ["pie", "doughnut"] as const) {
      const xml = odfChartContentXml(baseVisual({ chartType }), 12, 7.5, undefined, refs);
      expect(xml).toContain('<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="$Data.$A$2:.$A$4"/></chart:axis><chart:series');
      expect(xml).not.toContain('chart:dimension="y"');
      expect(xml.match(/<chart:series /g)).toHaveLength(1);
      expect(odfChartContentXml(baseVisual({ chartType }), 12, 7.5, undefined, { series: [] })).not.toContain("chart:axis");
    }
  });

  it("writes a series' domain as its chart:domain child (scatter x)", () => {
    const xml = odfChartContentXml(baseVisual({ chartType: "scatter" }), 12, 7.5, undefined, {
      series: [{ values: "$S.$B$2:.$B$4", label: "$S.$B$1", domain: "$S.$A$2:.$A$4" }, { values: "$S.$C$2:.$C$4" }],
    });
    expect(xml).toContain(
      'chart:values-cell-range-address="$S.$B$2:.$B$4" chart:label-cell-address="$S.$B$1"><chart:domain table:cell-range-address="$S.$A$2:.$A$4"/></chart:series>',
    );
    expect(xml).toContain('chart:values-cell-range-address="$S.$C$2:.$C$4" chart:label-cell-address="local-table.$C$1"/>');
  });

  it("escapes the wire text it is given and writes it otherwise verbatim (quoted names, the empty address)", () => {
    const xml = odfChartContentXml(baseVisual(), 12, 7.5, undefined, { plotArea: `$'R&D "x"'.A1:.B2`, series: [{ values: "$'Q1 Data'.$B$2:.$B$13", label: "" }] });
    expect(xml).toContain(`<chart:plot-area table:cell-range-address="$'R&amp;D &quot;x&quot;'.A1:.B2">`);
    expect(xml).toContain(`chart:values-cell-range-address="$'Q1 Data'.$B$2:.$B$13" chart:label-cell-address=""/>`);
  });

  it("still writes the local-table cache, and the chart reads back the same through parseOdfChartXml", () => {
    for (const { chartType } of CHART_CATALOGUE) {
      const chart = baseVisual({ chartType, options: { legend: "bottom", dataLabels: false, gridlines: true, title: "T", categoryAxisTitle: "C", valueAxisTitle: "V" } });
      const withRefs = odfChartContentXml(chart, 12, 7.5, undefined, { ...refs, series: refs.series.map((each) => ({ ...each, domain: "$Data.$A$2:.$A$4" })) });
      expect(withRefs).toContain(odfChartTableXml(chart));
      expect(parseOdfChartXml(withRefs)).toEqual(parseOdfChartXml(odfChartContentXml(chart, 12, 7.5)));
    }
  });

  it("odfChartTableXml is the local-table the object writes, holding the series pie and doughnut persist only", () => {
    for (const { chartType } of CHART_CATALOGUE) {
      const chart = baseVisual({ chartType });
      const table = odfChartTableXml(chart);
      expect(odfChartContentXml(chart, 12, 7.5)).toContain(`</chart:plot-area>${table}</chart:chart>`);
      expect(table.match(/<table:table-column\/>/g)).toHaveLength(1 + (SINGLE_SERIES_CHART_TYPES.has(chartType) ? 1 : 2));
    }
  });
});
