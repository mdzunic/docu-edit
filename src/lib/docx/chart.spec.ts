// SPEC-052: the DrawingML chart reader/writer. Mirrors drawing.spec.ts's
// own fixture conventions (a drawingElement wrapper, a synthetic mainRels/
// parts pair standing in for a real package's relationships/other parts).
import { describe, it, expect } from "vitest";
import {
  CHART_RELATIONSHIP,
  chartDrawingXml,
  chartPartXml,
  parseChartDrawing,
  parseChartXml,
} from "./chart";
import { parseXml } from "../core/xml";
import { equalStructural } from "../equality";
import { type Chart } from "../documentModel";
import {
  CHART_CATALOGUE,
  DEFAULT_CHART_COLORS,
  SINGLE_SERIES_CHART_TYPES,
  type ChartType,
} from "../core/charts";

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function drawingElement(xml: string): Element {
  const wrapped = `<root xmlns:w="${WORD_NS}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

const MAIN_PART_PATH = "word/document.xml";
const CHART_PART_PATH = "word/charts/chart1.xml";

function relsMap(): Map<string, { id: string; type: string; target: string }> {
  return new Map([["rId9", { id: "rId9", type: CHART_RELATIONSHIP, target: "charts/chart1.xml" }]]);
}

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

/** Writes a Chart's part XML + wrapper drawing, then reads it straight back
 * — the round trip this module's whole contract rests on. */
function roundTripPart(chart: Chart): ReturnType<typeof parseChartXml> {
  const xml = chartPartXml(chart);
  return parseChartXml(xml);
}

function roundTripDrawing(chart: Chart) {
  const partXml = chartPartXml(chart);
  const parts = new Map([[CHART_PART_PATH, new TextEncoder().encode(partXml)]]);
  const drawingXmlText = chartDrawingXml(chart, "rId9", 3);
  return parseChartDrawing(drawingElement(drawingXmlText), relsMap(), MAIN_PART_PATH, parts);
}

describe("chartPartXml / parseChartXml — catalogue round trip", () => {
  for (const { chartType } of CHART_CATALOGUE) {
    it(`round-trips ${chartType}`, () => {
      // Scatter treats categories as numeric X values (c:xVal, not c:cat) —
      // a non-numeric category falls back to its 1-based index, so this
      // catalogue-wide test uses numeric-looking labels throughout; scatter's
      // own X-value semantics get a dedicated test below.
      const chart = baseChart({ chartType, categories: ["1", "2", "3"] });
      const parsed = roundTripPart(chart);
      expect(parsed).not.toBeNull();
      expect(parsed!.chartType).toBe(chartType);
      expect(parsed!.categories).toEqual(chart.categories);
      // Pie/doughnut render (and therefore round-trip) only the first
      // series (SPEC-052 "Chart catalogue").
      const expectedSeries = SINGLE_SERIES_CHART_TYPES.has(chartType) ? chart.series.slice(0, 1) : chart.series;
      expect(parsed!.series.map((s) => s.name)).toEqual(expectedSeries.map((s) => s.name));
      expect(parsed!.series.map((s) => s.color)).toEqual(expectedSeries.map((s) => s.color));
      expect(parsed!.series.map((s) => s.values)).toEqual(expectedSeries.map((s) => s.values));
    });
  }

  it("scatter's non-numeric category text falls back to its 1-based index as the X value (known simplification)", () => {
    const chart = baseChart({ chartType: "scatter", categories: ["Cat 1", "Cat 2", "Cat 3"] });
    const parsed = roundTripPart(chart);
    expect(parsed!.categories).toEqual(["1", "2", "3"]);
  });

  it("preserves a gap (null) distinct from zero", () => {
    const chart = baseChart({
      series: [{ name: "S", values: [0, null, 5], color: "#111111" }],
    });
    const parsed = roundTripPart(chart);
    expect(parsed!.series[0].values).toEqual([0, null, 5]);
  });

  it("round-trips negative values", () => {
    const chart = baseChart({
      series: [{ name: "S", values: [-4, 2, -1.5], color: "#111111" }],
    });
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

  it("legend 'none' round-trips with no c:legend element", () => {
    const chart = baseChart({ options: { legend: "none", dataLabels: false, gridlines: false } });
    const xml = chartPartXml(chart);
    expect(xml).not.toContain("<c:legend>");
    const parsed = roundTripPart(chart);
    expect(parsed!.options.legend).toBe("none");
  });

  it("distinguishes line from lineMarkers via c:marker", () => {
    const line = chartPartXml(baseChart({ chartType: "line" }));
    const markers = chartPartXml(baseChart({ chartType: "lineMarkers" }));
    expect(line).toContain('<c:marker val="0"/>');
    expect(markers).toContain('<c:marker val="1"/>');
  });

  it("distinguishes columnClustered/columnStacked/barClustered via barDir+grouping", () => {
    expect(chartPartXml(baseChart({ chartType: "columnClustered" }))).toContain('<c:barDir val="col"/><c:grouping val="clustered"/>');
    expect(chartPartXml(baseChart({ chartType: "columnStacked" }))).toContain('<c:barDir val="col"/><c:grouping val="stacked"/>');
    expect(chartPartXml(baseChart({ chartType: "barClustered" }))).toContain('<c:barDir val="bar"/><c:grouping val="clustered"/>');
  });

  it("uses word/charts/chartN.xml's c:barChart/c:lineChart/c:pieChart with cached c:numCache/c:strCache (AC-36)", () => {
    const xml = chartPartXml(baseChart({ chartType: "columnClustered" }));
    expect(xml).toContain("<c:barChart>");
    expect(xml).toContain("c:numCache");
    expect(xml).toContain("c:strCache");
    const lineXml = chartPartXml(baseChart({ chartType: "line" }));
    expect(lineXml).toContain("<c:lineChart>");
    const pieXml = chartPartXml(baseChart({ chartType: "pie" }));
    expect(pieXml).toContain("<c:pieChart>");
  });

  it("returns null for a combo chart (two chart-type elements)", () => {
    const barXml = chartPartXml(baseChart({ chartType: "columnClustered" }));
    const comboXml = barXml.replace(
      "</c:barChart>",
      '</c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:marker val="0"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>',
    );
    expect(parseChartXml(comboXml)).toBeNull();
  });

  it("returns null for an unrecognised chart type (e.g. a 3-D chart)", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<c:chart><c:plotArea><c:layout/><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/></c:bar3DChart></c:plotArea></c:chart>' +
      "</c:chartSpace>";
    expect(parseChartXml(xml)).toBeNull();
  });

  it("returns null for malformed XML", () => {
    expect(parseChartXml("<not-xml")).toBeNull();
  });
});

describe("parseChartDrawing — the w:drawing wrapper", () => {
  it("round-trips size, wrap, offset, zOrder and alt text for an anchored chart", () => {
    const chart = baseChart({
      wrap: "behind",
      offsetXCm: 2.5,
      offsetYCm: 1.25,
      zOrder: 3,
      altText: "Quarterly revenue by product line",
      relativeToH: "column",
      relativeToV: "paragraph",
    });
    const parsed = roundTripDrawing(chart);
    expect(parsed).not.toBeNull();
    expect(parsed!.widthCm).toBeCloseTo(chart.widthCm, 2);
    expect(parsed!.heightCm).toBeCloseTo(chart.heightCm, 2);
    expect(parsed!.wrap).toBe("behind");
    expect(parsed!.offsetXCm).toBeCloseTo(2.5, 2);
    expect(parsed!.offsetYCm).toBeCloseTo(1.25, 2);
    expect(parsed!.zOrder).toBe(3);
    expect(parsed!.altText).toBe("Quarterly revenue by product line");
    expect(parsed!.sourcePart).toBe(CHART_PART_PATH);
    expect(parsed!.chartType).toBe("columnClustered");
  });

  it("round-trips an inline chart with no offsets", () => {
    const chart = baseChart({ wrap: "inline", offsetXCm: undefined, offsetYCm: undefined, relativeToH: undefined, relativeToV: undefined });
    const parsed = roundTripDrawing(chart);
    expect(parsed!.wrap).toBe("inline");
    expect(parsed!.offsetXCm).toBeUndefined();
    expect(parsed!.offsetYCm).toBeUndefined();
  });

  it("captures the raw XML as unsupportedChartXml for an out-of-catalogue chart, still resolving placement", () => {
    const chart = baseChart({ chartType: "columnClustered" as ChartType });
    // Corrupt the part so it no longer parses as a catalogue type (mirrors
    // parseChartXml's own "combo chart" null case above).
    const partXml = chartPartXml(chart).replace(
      "</c:barChart>",
      '</c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:marker val="0"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>',
    );
    const parts = new Map([[CHART_PART_PATH, new TextEncoder().encode(partXml)]]);
    const drawingXmlText = chartDrawingXml(chart, "rId9", 3);
    const parsed = parseChartDrawing(drawingElement(drawingXmlText), relsMap(), MAIN_PART_PATH, parts);
    expect(parsed).not.toBeNull();
    expect(parsed!.unsupportedChartXml).toBeDefined();
    expect(parsed!.unsupportedChartXml).toContain("c:barChart");
    expect(parsed!.sourcePart).toBe(CHART_PART_PATH);
    // Placement (size/wrap) is still recovered even though content isn't.
    expect(parsed!.widthCm).toBeCloseTo(chart.widthCm, 2);
  });

  it("returns null when the drawing holds no c:chart reference (an image instead)", () => {
    const notAChart =
      '<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="100000" cy="100000"/>' +
      '<wp:docPr id="1" name="x"/><wp:cNvGraphicFramePr/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/>' +
      "</a:graphic></wp:inline></w:drawing>";
    expect(parseChartDrawing(drawingElement(notAChart), relsMap(), MAIN_PART_PATH, new Map())).toBeNull();
  });

  it("returns null when the referenced part is missing from the package", () => {
    const chart = baseChart();
    const drawingXmlText = chartDrawingXml(chart, "rId9", 3);
    expect(parseChartDrawing(drawingElement(drawingXmlText), relsMap(), MAIN_PART_PATH, new Map())).toBeNull();
  });
});

describe("equalStructural of a full round trip", () => {
  it("a chart written, read back and re-written produces the same model", () => {
    const chart = baseChart({
      options: { title: "T", legend: "top", dataLabels: true, gridlines: true, valueAxisMin: 0, valueAxisMax: 10 },
    });
    const first = roundTripPart(chart)!;
    const rewritten = chartPartXml({ ...chart, chartType: first.chartType, categories: first.categories, series: first.series, options: first.options });
    const second = parseChartXml(rewritten)!;
    expect(equalStructural(first, second)).toBe(true);
  });
});

describe("DEFAULT_CHART_COLORS fallback", () => {
  it("assigns a palette colour to a series with no c:spPr/a:solidFill", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      "<c:chart><c:plotArea><c:layout/>" +
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      '<c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
      '<c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>' +
      "</c:ser>" +
      '<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
      "</c:plotArea></c:chart></c:chartSpace>";
    const parsed = parseChartXml(xml);
    expect(parsed!.series[0].color).toBe(DEFAULT_CHART_COLORS[0]);
  });
});
