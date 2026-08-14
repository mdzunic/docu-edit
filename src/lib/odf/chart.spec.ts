// SPEC-052: the ODF chart reader/writer — a draw:frame > draw:object
// embedded chart:chart object, mirroring odf/shape.spec.ts's own fixture
// conventions.
import { describe, it, expect } from "vitest";
import { chartFrameXml, odfChartContentXml, parseChartFrame, parseOdfChartXml } from "./chart";
import { parseXml } from "../docx/xml";
import { equalStructural } from "../equality";
import { CHART_CATALOGUE, DEFAULT_CHART_COLORS, SINGLE_SERIES_CHART_TYPES, type Chart } from "../documentModel";
import { DRAW_NAMESPACE, SVG_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE } from "./xml";

// A real draw:frame lives inside content.xml, whose root already declares
// every namespace it uses — a standalone parse needs the same ambient
// declarations, mirroring odf/drawing.spec.ts's own wrap() helper.
function wrap(xml: string): Element {
  const wrapped =
    `<root xmlns:text="${TEXT_NAMESPACE}" xmlns:draw="${DRAW_NAMESPACE}" ` +
    `xmlns:svg="${SVG_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

const OBJECT_DIR = "Object 1";
const CONTENT_PART_PATH = `${OBJECT_DIR}/content.xml`;

function baseChart(overrides: Partial<Chart> = {}): Chart {
  return {
    type: "chart",
    chartType: "columnClustered",
    categories: ["Cat 1", "Cat 2", "Cat 3"],
    series: [
      { name: "Series A", values: [1, 2, 3], color: "#4e79a7" },
      { name: "Series B", values: [4, null, 6], color: "#f28e2b" },
    ],
    options: { legend: "right", dataLabels: false, gridlines: true },
    widthCm: 12,
    heightCm: 7.5,
    wrap: "square",
    ...overrides,
  };
}

function roundTripPart(chart: Chart): ReturnType<typeof parseOdfChartXml> {
  const xml = odfChartContentXml(chart);
  return parseOdfChartXml(xml);
}

function roundTripFrame(chart: Chart) {
  const partXml = odfChartContentXml(chart);
  const parts = new Map([[CONTENT_PART_PATH, new TextEncoder().encode(partXml)]]);
  const frameXml = chartFrameXml(chart, `./${OBJECT_DIR}`);
  return parseChartFrame(wrap(frameXml), parts);
}

describe("odfChartContentXml / parseOdfChartXml — catalogue round trip", () => {
  for (const { chartType } of CHART_CATALOGUE) {
    it(`round-trips ${chartType}`, () => {
      const chart = baseChart({ chartType });
      const parsed = roundTripPart(chart);
      expect(parsed).not.toBeNull();
      expect(parsed!.chartType).toBe(chartType);
      expect(parsed!.categories).toEqual(chart.categories);
      const expectedSeries = SINGLE_SERIES_CHART_TYPES.has(chartType) ? chart.series.slice(0, 1) : chart.series;
      expect(parsed!.series.map((s) => s.name)).toEqual(expectedSeries.map((s) => s.name));
      expect(parsed!.series.map((s) => s.color)).toEqual(expectedSeries.map((s) => s.color));
      expect(parsed!.series.map((s) => s.values)).toEqual(expectedSeries.map((s) => s.values));
    });
  }

  it("preserves a gap (null) distinct from zero", () => {
    const chart = baseChart({ series: [{ name: "S", values: [0, null, 5], color: "#111111" }] });
    const parsed = roundTripPart(chart);
    expect(parsed!.series[0].values).toEqual([0, null, 5]);
  });

  it("round-trips negative values", () => {
    const chart = baseChart({ series: [{ name: "S", values: [-4, 2, -1.5], color: "#111111" }] });
    const parsed = roundTripPart(chart);
    expect(parsed!.series[0].values).toEqual([-4, 2, -1.5]);
  });

  it("round-trips title, axis titles, legend, data labels, gridlines and axis min/max", () => {
    const chart = baseChart({
      options: {
        title: "Quarterly Revenue",
        categoryAxisTitle: "Quarter",
        valueAxisTitle: "USD (millions)",
        legend: "bottom",
        dataLabels: true,
        gridlines: true,
        valueAxisMin: -10,
        valueAxisMax: 50,
      },
    });
    const parsed = roundTripPart(chart);
    expect(parsed!.options).toEqual(chart.options);
  });

  it("legend 'none' round-trips with no chart:legend element", () => {
    const chart = baseChart({ options: { legend: "none", dataLabels: false, gridlines: false } });
    const xml = odfChartContentXml(chart);
    expect(xml).not.toContain("<chart:legend");
    const parsed = roundTripPart(chart);
    expect(parsed!.options.legend).toBe("none");
  });

  it("uses chart:chart, chart:plot-area and table:table data (AC-37)", () => {
    const xml = odfChartContentXml(baseChart());
    expect(xml).toContain("<chart:chart");
    expect(xml).toContain("<chart:plot-area");
    expect(xml).toContain("<table:table");
    expect(xml).toContain('chart:class="chart:bar"');
  });

  it("distinguishes columnClustered/columnStacked/barClustered via vertical+stacked", () => {
    expect(odfChartContentXml(baseChart({ chartType: "columnClustered" }))).toContain('chart:vertical="true" chart:stacked="false"');
    expect(odfChartContentXml(baseChart({ chartType: "columnStacked" }))).toContain('chart:vertical="true" chart:stacked="true"');
    expect(odfChartContentXml(baseChart({ chartType: "barClustered" }))).toContain('chart:vertical="false" chart:stacked="false"');
  });

  it("returns null for malformed XML", () => {
    expect(parseOdfChartXml("<not-xml")).toBeNull();
  });

  it("returns null for a chart:class outside the catalogue", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">' +
      '<office:body><office:chart><chart:chart chart:class="chart:stock"></chart:chart></office:chart></office:body>' +
      "</office:document-content>";
    expect(parseOdfChartXml(xml)).toBeNull();
  });
});

describe("parseChartFrame — the draw:frame/draw:object wrapper", () => {
  it("round-trips size, wrap and alt text", () => {
    const chart = baseChart({ wrap: "square", altText: "Quarterly revenue by product line" });
    const parsed = roundTripFrame(chart);
    expect(parsed).not.toBeNull();
    expect(parsed!.widthCm).toBeCloseTo(chart.widthCm, 1);
    expect(parsed!.heightCm).toBeCloseTo(chart.heightCm, 1);
    expect(parsed!.wrap).toBe("square");
    expect(parsed!.altText).toBe("Quarterly revenue by product line");
    expect(parsed!.sourcePart).toBe(CONTENT_PART_PATH);
    expect(parsed!.chartType).toBe("columnClustered");
  });

  it("round-trips an inline chart", () => {
    const chart = baseChart({ wrap: "inline" });
    const parsed = roundTripFrame(chart);
    expect(parsed!.wrap).toBe("inline");
  });

  it("captures the raw XML as unsupportedChartXml for an out-of-catalogue chart", () => {
    const chart = baseChart();
    const badXml = odfChartContentXml(chart).replace('chart:class="chart:bar"', 'chart:class="chart:stock"');
    const parts = new Map([[CONTENT_PART_PATH, new TextEncoder().encode(badXml)]]);
    const frameXml = chartFrameXml(chart, `./${OBJECT_DIR}`);
    const parsed = parseChartFrame(wrap(frameXml), parts);
    expect(parsed).not.toBeNull();
    expect(parsed!.unsupportedChartXml).toBeDefined();
    expect(parsed!.unsupportedChartXml).toContain("chart:stock");
    expect(parsed!.sourcePart).toBe(CONTENT_PART_PATH);
  });

  it("returns null when the embedded object part is missing", () => {
    const chart = baseChart();
    const frameXml = chartFrameXml(chart, `./${OBJECT_DIR}`);
    expect(parseChartFrame(wrap(frameXml), new Map())).toBeNull();
  });

  it("returns null when the frame holds no draw:object (e.g. a picture)", () => {
    const notAChart =
      '<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" text:anchor-type="as-char" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" svg:width="5cm" svg:height="3cm">' +
      '<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>' +
      "</draw:frame>";
    const el = new DOMParser().parseFromString(notAChart, "application/xml").documentElement;
    expect(parseChartFrame(el, new Map())).toBeNull();
  });
});

describe("equalStructural of a full round trip", () => {
  it("a chart written, read back and re-written produces the same model", () => {
    const chart = baseChart({
      options: { title: "T", legend: "top", dataLabels: true, gridlines: true, valueAxisMin: 0, valueAxisMax: 10 },
    });
    const first = roundTripPart(chart)!;
    const rewritten = odfChartContentXml({ ...chart, chartType: first.chartType, categories: first.categories, series: first.series, options: first.options });
    const second = parseOdfChartXml(rewritten)!;
    expect(equalStructural(first, second)).toBe(true);
  });
});

describe("DEFAULT_CHART_COLORS fallback", () => {
  it("assigns a palette colour to a series with no chart:color", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">' +
      '<office:body><office:chart><chart:chart chart:class="chart:bar" chart:vertical="true" chart:stacked="false">' +
      '<chart:plot-area><chart:series chart:class="chart:bar" chart:values-cell-range-address="local-table.$B$2:.$B$2" chart:label-cell-address="local-table.$B$1"/></chart:plot-area>' +
      '<table:table table:name="local-table"><table:table-header-rows><table:table-row><table:table-cell/><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell></table:table-row></table:table-header-rows>' +
      '<table:table-rows><table:table-row><table:table-cell office:value-type="string"><text:p>Cat</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table-rows></table:table>' +
      "</chart:chart></office:chart></office:body>" +
      "</office:document-content>";
    const parsed = parseOdfChartXml(xml);
    expect(parsed!.series[0].color).toBe(DEFAULT_CHART_COLORS[0]);
  });
});
