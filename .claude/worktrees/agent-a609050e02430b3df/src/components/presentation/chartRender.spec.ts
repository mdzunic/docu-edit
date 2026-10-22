// SPEC-092 AC-11a: chartVisualOf narrows a chart frame to the ChartVisual the
// canvas painter consumes, and returns null for any other frame kind.
import { describe, it, expect } from "vitest";
import { chartVisualOf } from "./slideRender";
import type { ChartFrame, ShapeFrame } from "../../lib/presentation/model";

const baseGeom = { id: "f1", xCm: 1, yCm: 1, widthCm: 12, heightCm: 7.5 };

describe("chartVisualOf", () => {
  it("returns the ChartVisual for a chart frame", () => {
    const frame: ChartFrame = {
      ...baseGeom,
      kind: "chart",
      chartType: "pie",
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2], color: "#4e79a7" }],
      options: { legend: "right", dataLabels: false, gridlines: false },
      sourcePart: "ppt/charts/chart1.xml",
    };
    const visual = chartVisualOf(frame);
    expect(visual).not.toBeNull();
    expect(visual!.chartType).toBe("pie");
    expect(visual!.categories).toEqual(["A", "B"]);
    expect(visual!.series[0].name).toBe("S");
    expect(visual!.unsupportedChartXml).toBeUndefined();
  });

  it("carries unsupportedChartXml through so paintChart draws the placeholder", () => {
    const frame: ChartFrame = {
      ...baseGeom,
      kind: "chart",
      chartType: "columnClustered",
      categories: [],
      series: [],
      options: { legend: "none", dataLabels: false, gridlines: false },
      unsupportedChartXml: "<c:chartSpace/>",
    };
    expect(chartVisualOf(frame)!.unsupportedChartXml).toBe("<c:chartSpace/>");
  });

  it("returns null for a non-chart frame", () => {
    const shape: ShapeFrame = { ...baseGeom, kind: "shape", geometry: "rect" };
    expect(chartVisualOf(shape)).toBeNull();
  });
});
