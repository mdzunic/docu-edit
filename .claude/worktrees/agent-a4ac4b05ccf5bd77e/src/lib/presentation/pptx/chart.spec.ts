// SPEC-092 AC-29, AC-29a: the pptx chart serialiser equals the hand-authored
// goldens, and the reader accepts them (plus attribute-reordered variants).
import { describe, it, expect } from "vitest";
import { serializeChartFrame } from "./serialize";
import { parseChartGraphicFrameForTest } from "./reader";
import { chartPartXml } from "../../core/chartPartOoxml";
import { CHART_GRAPHIC_FRAME_GOLDEN, CHART_PART_GOLDEN, GOLDEN_CHART_FRAME } from "./chartGolden";
import type { ChartFrame } from "../model";

const bytes = (s: string) => new TextEncoder().encode(s);

describe("serializeChartFrame — golden equality (AC-29)", () => {
  it("emits the canonical p:graphicFrame", () => {
    expect(serializeChartFrame(GOLDEN_CHART_FRAME, 2, "rId2")).toBe(CHART_GRAPHIC_FRAME_GOLDEN);
  });

  it("chartPartXml emits the canonical c:chartSpace", () => {
    expect(chartPartXml(GOLDEN_CHART_FRAME)).toBe(CHART_PART_GOLDEN);
  });
});

describe("reader accepts the goldens (AC-29a)", () => {
  it("round-trips the golden graphicFrame + part back to the model", () => {
    const frame = parseChartGraphicFrameForTest(
      CHART_GRAPHIC_FRAME_GOLDEN,
      { rId2: "../charts/chart1.xml" },
      { "ppt/charts/chart1.xml": bytes(CHART_PART_GOLDEN) },
    ) as ChartFrame;
    expect(frame.kind).toBe("chart");
    expect(frame.chartType).toBe(GOLDEN_CHART_FRAME.chartType);
    expect(frame.categories).toEqual(GOLDEN_CHART_FRAME.categories);
    expect(frame.series.map((s) => s.name)).toEqual(["S"]);
    expect(frame.series.map((s) => s.values)).toEqual([[1, 2]]);
    expect(frame.series[0].color).toBe("#4e79a7");
    expect(frame.sourcePart).toBe("ppt/charts/chart1.xml");
    expect(frame.widthCm).toBeCloseTo(12, 2);
    expect(frame.heightCm).toBeCloseTo(7.5, 2);
  });

  it("accepts a PowerPoint-flavoured variant with reordered attributes", () => {
    const reordered = CHART_GRAPHIC_FRAME_GOLDEN.replace(
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>',
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>',
    );
    const frame = parseChartGraphicFrameForTest(
      reordered,
      { rId2: "../charts/chart1.xml" },
      { "ppt/charts/chart1.xml": bytes(CHART_PART_GOLDEN) },
    );
    expect(frame.kind).toBe("chart");
  });
});
