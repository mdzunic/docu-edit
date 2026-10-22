// SPEC-092 AC-29, AC-29a: the odp chart serialiser equals the hand-authored
// golden, and the reader accepts it (plus attribute-reordered variants).
import { describe, it, expect } from "vitest";
import { serializeOdpChartFrame } from "./serialize";
import { parseChartDrawFrameForTest } from "./reader";
import { odfChartContentXml } from "../../core/chartPartOdf";
import { CHART_DRAW_FRAME_GOLDEN, GOLDEN_CHART_FRAME } from "./chartGolden";
import type { ChartFrame } from "../model";

const bytes = (s: string) => new TextEncoder().encode(s);

describe("serializeOdpChartFrame — golden equality (AC-29)", () => {
  it("emits the canonical draw:frame + draw:object", () => {
    expect(serializeOdpChartFrame(GOLDEN_CHART_FRAME, "./Object 1")).toBe(CHART_DRAW_FRAME_GOLDEN);
  });
});

describe("reader accepts the golden (AC-29a)", () => {
  it("round-trips the golden draw:frame + content.xml back to the model", () => {
    const content = odfChartContentXml(GOLDEN_CHART_FRAME, GOLDEN_CHART_FRAME.widthCm, GOLDEN_CHART_FRAME.heightCm);
    const frame = parseChartDrawFrameForTest(CHART_DRAW_FRAME_GOLDEN, {
      "Object 1/content.xml": bytes(content),
    }) as ChartFrame;
    expect(frame.kind).toBe("chart");
    expect(frame.chartType).toBe("columnClustered");
    expect(frame.categories).toEqual(["A", "B"]);
    expect(frame.series.map((s) => s.values)).toEqual([[1, 2]]);
    expect(frame.sourcePart).toBe("Object 1/content.xml");
    expect(frame.widthCm).toBeCloseTo(12, 2);
  });

  it("accepts an Impress-flavoured variant with reordered attributes", () => {
    const reordered = CHART_DRAW_FRAME_GOLDEN.replace(
      '<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>',
      '<draw:object xlink:actuate="onLoad" xlink:show="embed" xlink:type="simple" xlink:href="./Object 1"/>',
    );
    const content = odfChartContentXml(GOLDEN_CHART_FRAME, 12, 7.5);
    const frame = parseChartDrawFrameForTest(reordered, { "Object 1/content.xml": bytes(content) });
    expect(frame.kind).toBe("chart");
  });
});
