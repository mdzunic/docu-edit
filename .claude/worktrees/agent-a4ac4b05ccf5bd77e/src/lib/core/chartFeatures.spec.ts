// SPEC-092: the hoisted chart-feature markers and `lostChartFeatures`. The
// text-product coverage (`opaque.spec.ts`) pins the same marker tables through
// `collectEditedChartDropRecords`; this pins the core surface the presentation
// rewrite-notice reads.
import { describe, it, expect } from "vitest";
import { DOCX_CHART_FEATURE_MARKERS, ODT_CHART_FEATURE_MARKERS, lostChartFeatures } from "./chartFeatures";

describe("lostChartFeatures — ooxml", () => {
  it("names a trendline, error bars, per-point formatting and a custom number format", () => {
    const xml =
      "<c:chartSpace><c:trendline/><c:errBars/><c:dPt><c:idx/></c:dPt>" +
      '<c:numFmt formatCode="0.00%"/></c:chartSpace>';
    const lost = lostChartFeatures(xml, "ooxml");
    expect(lost).toContain("a trendline");
    expect(lost).toContain("error bars");
    expect(lost).toContain("per-point formatting (individual bar/slice colours or markers)");
    expect(lost).toContain("a custom number format");
  });

  it("names a secondary axis only when a second c:valAx is present", () => {
    expect(lostChartFeatures("<c:valAx><c:axId/></c:valAx>", "ooxml")).not.toContain("a secondary axis");
    expect(lostChartFeatures("<c:valAx><c:axId/></c:valAx><c:valAx><c:axId/></c:valAx>", "ooxml")).toContain("a secondary axis");
  });

  it("returns nothing for a plain part", () => {
    expect(lostChartFeatures("<c:chartSpace><c:barChart/></c:chartSpace>", "ooxml")).toEqual([]);
  });
});

describe("lostChartFeatures — odf", () => {
  it("names a trendline, error bars and per-point formatting", () => {
    const xml = "<chart:chart><chart:regression-curve/><chart:error-indicator/><chart:data-point chart:repeated='1'/></chart:chart>";
    const lost = lostChartFeatures(xml, "odf");
    expect(lost).toContain("a trendline");
    expect(lost).toContain("error bars");
    expect(lost).toContain("per-point formatting (individual bar/slice colours or markers)");
  });

  it("returns nothing for a plain part", () => {
    expect(lostChartFeatures('<chart:chart chart:class="chart:bar"/>', "odf")).toEqual([]);
  });
});

describe("marker tables", () => {
  it("expose the OOXML and ODF marker sets", () => {
    expect(DOCX_CHART_FEATURE_MARKERS.length).toBe(6);
    expect(ODT_CHART_FEATURE_MARKERS.length).toBe(4);
  });
});
