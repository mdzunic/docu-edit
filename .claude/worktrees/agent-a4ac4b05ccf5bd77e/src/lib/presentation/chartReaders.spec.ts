// SPEC-092 AC-7..AC-10: the slide-side chart-part resolution in both readers.
// Uses the readers' test seams so a full package isn't needed — a chart
// graphicFrame/drawFrame plus a map of part bytes is enough.
import { describe, it, expect } from "vitest";
import { parseChartGraphicFrameForTest } from "./pptx/reader";
import { parseChartDrawFrameForTest } from "./odp/reader";
import { chartPartXml } from "../core/chartPartOoxml";
import { odfChartContentXml } from "../core/chartPartOdf";
import type { ChartFrame } from "./model";
import type { ChartVisual } from "../core/charts";

const VISUAL: ChartVisual = {
  chartType: "columnClustered",
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "Revenue", values: [10, 20, 30], color: "#4e79a7" },
    { name: "Cost", values: [4, null, 6], color: "#f28e2b" },
  ],
  options: { legend: "right", dataLabels: false, gridlines: true },
};

const bytes = (s: string) => new TextEncoder().encode(s);

// --- pptx -------------------------------------------------------------------

const PPTX_GF =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 5"/>` +
  `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="914400" y="457200"/><a:ext cx="4114800" cy="2743200"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart r:id="rId9"/></a:graphicData></a:graphic></p:graphicFrame>`;

describe("pptx reader — chart graphicFrame (AC-7, AC-8, AC-8a)", () => {
  it("resolves an in-catalogue chart part into a painted ChartFrame", () => {
    const frame = parseChartGraphicFrameForTest(
      PPTX_GF,
      { rId9: "../charts/chart1.xml" },
      { "ppt/charts/chart1.xml": bytes(chartPartXml(VISUAL)) },
    ) as ChartFrame;
    expect(frame.kind).toBe("chart");
    expect(frame.chartType).toBe("columnClustered");
    expect(frame.series.map((s) => s.name)).toEqual(["Revenue", "Cost"]);
    expect(frame.sourcePart).toBe("ppt/charts/chart1.xml");
    expect(frame.xml).toContain("<p:graphicFrame");
    expect(frame.unsupportedChartXml).toBeUndefined();
    // Geometry from p:xfrm (EMU → cm): 914400 EMU = 2.54cm.
    expect(frame.xCm).toBeCloseTo(2.54, 2);
    expect(frame.widthCm).toBeCloseTo(11.43, 2);
  });

  it("keeps an out-of-catalogue but valid chartSpace as an unsupported chart (AC-8)", () => {
    const threeD =
      '<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      "<c:chart><c:plotArea><c:bar3DChart/></c:plotArea></c:chart></c:chartSpace>";
    const frame = parseChartGraphicFrameForTest(
      PPTX_GF,
      { rId9: "../charts/chart1.xml" },
      { "ppt/charts/chart1.xml": bytes(threeD) },
    ) as ChartFrame;
    expect(frame.kind).toBe("chart");
    expect(frame.unsupportedChartXml).toContain("bar3DChart");
    expect(frame.sourcePart).toBe("ppt/charts/chart1.xml");
  });

  it("keeps a ForeignFrame 'Chart' when the relationship cannot resolve (AC-8a)", () => {
    const frame = parseChartGraphicFrameForTest(PPTX_GF, {}, {});
    expect(frame.kind).toBe("foreign");
    expect(frame.kind === "foreign" && frame.label).toBe("Chart");
  });

  it("keeps a ForeignFrame 'Chart' when the resolved part is not a chartSpace (AC-8a)", () => {
    const frame = parseChartGraphicFrameForTest(
      PPTX_GF,
      { rId9: "../charts/chart1.xml" },
      { "ppt/charts/chart1.xml": bytes("<not-a-chart/>") },
    );
    expect(frame.kind).toBe("foreign");
  });
});

// --- odp --------------------------------------------------------------------

const ODP_FRAME =
  `<draw:frame draw:id="f1" svg:x="2cm" svg:y="1cm" svg:width="12cm" svg:height="7.5cm" presentation:class="chart">` +
  `<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `<draw:image xlink:href="./ObjectReplacements/Object 1"/>` +
  `</draw:frame>`;

describe("odp reader — chart draw:object (AC-9, AC-9a, AC-9b, AC-10)", () => {
  it("resolves an in-catalogue embedded chart into a painted ChartFrame", () => {
    const frame = parseChartDrawFrameForTest(ODP_FRAME, {
      "Object 1/content.xml": bytes(odfChartContentXml(VISUAL, 12, 7.5)),
    }) as ChartFrame;
    expect(frame.kind).toBe("chart");
    expect(frame.chartType).toBe("columnClustered");
    expect(frame.sourcePart).toBe("Object 1/content.xml");
    expect(frame.widthCm).toBeCloseTo(12, 2);
    expect(frame.heightCm).toBeCloseTo(7.5, 2);
  });

  it("ignores the ObjectReplacements preview image, not treating it as an image frame (AC-9a)", () => {
    const frame = parseChartDrawFrameForTest(ODP_FRAME, {
      "Object 1/content.xml": bytes(odfChartContentXml(VISUAL, 12, 7.5)),
    });
    expect(frame.kind).toBe("chart");
  });

  it("keeps an out-of-catalogue chart:class as an unsupported chart (AC-9b)", () => {
    const surface = odfChartContentXml(VISUAL, 12, 7.5).replace('chart:class="chart:bar"', 'chart:class="chart:surface"');
    const frame = parseChartDrawFrameForTest(ODP_FRAME, { "Object 1/content.xml": bytes(surface) }) as ChartFrame;
    expect(frame.kind).toBe("chart");
    expect(frame.unsupportedChartXml).toContain("chart:surface");
  });

  it("keeps a ForeignFrame 'Object' when the href does not resolve (AC-10)", () => {
    const frame = parseChartDrawFrameForTest(ODP_FRAME, {});
    expect(frame.kind).toBe("foreign");
    expect(frame.kind === "foreign" && frame.label).toBe("Object");
  });

  it("keeps a ForeignFrame 'Object' when the part is not an office:chart (AC-10)", () => {
    const frame = parseChartDrawFrameForTest(ODP_FRAME, { "Object 1/content.xml": bytes("<not-a-chart/>") });
    expect(frame.kind).toBe("foreign");
  });
});
