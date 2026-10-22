// SPEC-092 AC-16, AC-16a, AC-23: the chart commands and the rewrite-notice
// helper.
import { describe, it, expect } from "vitest";
import {
  insertChartFrame,
  lostChartFeaturesOnRewrite,
  setSlideChartData,
  setSlideChartOptions,
  setSlideChartSeriesColor,
} from "./charts";
import { moveFrame, resizeFrame, reorderFrame } from "./frames";
import { chartPartXml } from "../core/chartPartOoxml";
import { bytesToBase64 } from "../core/packageParts";
import type { ChartFrame, PresentationModel } from "./model";
import type { ChartVisual } from "../core/charts";

function deck(format: "pptx" | "odp" = "pptx"): PresentationModel {
  return {
    kind: "presentation",
    id: "d1",
    title: "Deck",
    format,
    slideWidthCm: 25.4,
    slideHeightCm: 19.05,
    slides: [{ id: "s1", frames: [] }],
  };
}

const RECT = { xCm: 6.7, yCm: 5.8, widthCm: 12, heightCm: 7.5 };

function chartOf(model: PresentationModel): ChartFrame {
  return model.slides[0].frames[0] as ChartFrame;
}

describe("insertChartFrame (AC-13a, AC-16)", () => {
  it("inserts a seeded, selected chart frame at the rect", () => {
    const model = deck();
    const id = insertChartFrame(model, 0, "pie", RECT);
    expect(id).not.toBeNull();
    const frame = chartOf(model);
    expect(frame.kind).toBe("chart");
    expect(frame.chartType).toBe("pie");
    expect(frame.categories.length).toBeGreaterThan(0);
    expect(frame.series.length).toBeGreaterThan(0);
    expect(frame.widthCm).toBe(12);
    expect(frame.heightCm).toBe(7.5);
    expect(frame.edited).toBe(true);
    expect(frame.sourcePart).toBeUndefined();
    expect(model.slides[0].dirty).toBe(true);
  });
});

describe("data / options commands (AC-16)", () => {
  it("setSlideChartData sets edited, drops xml, replaces data", () => {
    const model = deck();
    const id = insertChartFrame(model, 0, "columnClustered", RECT)!;
    const frame = chartOf(model);
    frame.xml = "<p:graphicFrame/>"; // pretend it came from a file
    frame.edited = undefined as unknown as true;
    setSlideChartData(model, 0, id, ["A", "B"], [{ name: "S", values: [1, 2], color: "#123456" }]);
    expect(frame.categories).toEqual(["A", "B"]);
    expect(frame.series[0].name).toBe("S");
    expect(frame.edited).toBe(true);
    expect(frame.xml).toBeUndefined();
  });

  it("setSlideChartOptions replaces options and marks edited", () => {
    const model = deck();
    const id = insertChartFrame(model, 0, "columnClustered", RECT)!;
    setSlideChartOptions(model, 0, id, { legend: "bottom", dataLabels: true, gridlines: false });
    const frame = chartOf(model);
    expect(frame.options.legend).toBe("bottom");
    expect(frame.options.dataLabels).toBe(true);
    expect(frame.edited).toBe(true);
  });

  it("setSlideChartSeriesColor recolours one series", () => {
    const model = deck();
    const id = insertChartFrame(model, 0, "columnClustered", RECT)!;
    setSlideChartSeriesColor(model, 0, id, 0, "#abcdef");
    expect(chartOf(model).series[0].color).toBe("#abcdef");
  });

  it("is a no-op on an unsupported chart", () => {
    const model = deck();
    insertChartFrame(model, 0, "columnClustered", RECT);
    const frame = chartOf(model);
    frame.unsupportedChartXml = "<c:chartSpace/>";
    expect(setSlideChartData(model, 0, frame.id, ["A"], [])).toBeNull();
    expect(setSlideChartOptions(model, 0, frame.id, frame.options)).toBeNull();
  });
});

describe("geometry commands never set edited (AC-16a)", () => {
  it("moveFrame/resizeFrame/reorderFrame drop the verbatim xml but leave edited untouched", () => {
    for (const op of [
      (m: PresentationModel, id: string) => moveFrame(m, 0, id, 1, 1),
      (m: PresentationModel, id: string) => resizeFrame(m, 0, id, { xCm: 1, yCm: 1, widthCm: 8, heightCm: 5 }),
      // Give reorder a second frame so "back" actually moves the chart.
      (m: PresentationModel, id: string) => reorderFrame(m, 0, id, "back"),
    ]) {
      const model = deck();
      insertChartFrame(model, 0, "line", { xCm: 0, yCm: 0, widthCm: 4, heightCm: 4 }); // z-order neighbour
      const id = insertChartFrame(model, 0, "columnClustered", RECT)!;
      const frame = model.slides[0].frames.find((f) => f.id === id) as ChartFrame;
      // Simulate an untouched imported chart: has verbatim xml, not edited.
      frame.xml = "<p:graphicFrame/>";
      frame.edited = undefined as unknown as true;
      op(model, id);
      expect(frame.xml).toBeUndefined(); // regenerate container at new geometry
      expect(frame.edited).toBeFalsy(); // but the chart PART is untouched
    }
  });
});

describe("lostChartFeaturesOnRewrite (AC-23)", () => {
  const VISUAL: ChartVisual = {
    chartType: "columnClustered",
    categories: ["Q1"],
    series: [{ name: "S", values: [1], color: "#4e79a7" }],
    options: { legend: "right", dataLabels: false, gridlines: true },
  };

  it("names lost features and the workbook line for a pptx chart with externalData", () => {
    const sourceXml =
      chartPartXml(VISUAL).replace("</c:chart>", "</c:chart><c:externalData r:id=\"rId1\"><c:autoUpdate val=\"0\"/></c:externalData>") +
      "";
    const withTrend = sourceXml.replace("<c:ser>", "<c:ser><c:trendline/>");
    const model = deck("pptx");
    model.opaque = [{ name: "ppt/charts/chart1.xml", data: bytesToBase64(new TextEncoder().encode(withTrend)) }];
    const frame: ChartFrame = {
      id: "c1",
      xCm: 0,
      yCm: 0,
      widthCm: 12,
      heightCm: 7.5,
      kind: "chart",
      ...VISUAL,
      sourcePart: "ppt/charts/chart1.xml",
    };
    const lost = lostChartFeaturesOnRewrite(model, frame);
    expect(lost).toContain("a trendline");
    expect(lost.some((l) => l.includes("workbook"))).toBe(true);
  });

  it("returns nothing for a chart created here (no sourcePart)", () => {
    const model = deck();
    const id = insertChartFrame(model, 0, "columnClustered", RECT)!;
    expect(lostChartFeaturesOnRewrite(model, chartOf(model))).toEqual([]);
    expect(id).toBeTruthy();
  });
});
