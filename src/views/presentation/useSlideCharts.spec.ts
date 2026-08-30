// View spec: the chart controller's command + dialog wiring (SPEC-092
// AC-13a/14/15/16/17/23).
import { computed, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { insertChartFrame } from "../../lib/presentation/charts";
import { defaultChartOptions } from "../../lib/core/charts";
import type { ChartFrame, PresentationModel, SlideFrame } from "../../lib/presentation/model";
import { useSlideCharts } from "./useSlideCharts";

function setup(seed?: (model: PresentationModel) => string) {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const seeded = seed ? seed(model.value!) : null;
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const currentIndex = shallowRef(0);
  const selectedId = shallowRef<string | null>(seeded);
  const selectedFrame = computed<SlideFrame | null>(
    () => currentSlide.value?.frames.find((f) => f.id === selectedId.value) ?? null,
  );
  const selectFrame = vi.fn((next: string | null) => {
    selectedId.value = next;
  });
  const record = vi.fn();
  // The real render trio (useSlideRender) re-identifies the slide + its frames
  // and retriggers the shallowRef on commit; mirror that so the controller's
  // computeds (which read the selected frame) recompute like they do in the app.
  const commit = vi.fn(() => {
    const m = model.value;
    if (!m) return;
    m.slides = m.slides.map((s) => ({ ...s, frames: s.frames.map((f) => ({ ...f })) }));
    model.value = { ...m };
  });
  const controller = useSlideCharts({
    model,
    currentSlide,
    currentIndex,
    selectedFrame,
    selectFrame,
    record,
    commit,
  });
  return { model, controller, selectedId, record, commit };
}

function chartFrame(model: PresentationModel): ChartFrame {
  const frame = model.slides[0].frames.find((f) => f.kind === "chart");
  if (!frame || frame.kind !== "chart") {
    throw new Error("no chart frame");
  }
  return frame;
}

describe("useSlideCharts", () => {
  it("inserts a 12×7.5cm chart centred on the slide, selected (AC-13a)", () => {
    const { model, controller, record, commit } = setup();
    controller.pickChartType("columnClustered");
    const frame = chartFrame(model.value!);
    expect(frame.widthCm).toBe(12);
    expect(frame.heightCm).toBe(7.5);
    expect(frame.xCm).toBeCloseTo((model.value!.slideWidthCm - 12) / 2, 3);
    expect(frame.chartType).toBe("columnClustered");
    expect(record).toHaveBeenCalledWith("Insert chart");
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("edits data / options / series colour through the command layer (AC-16/17)", () => {
    const { model, controller, record } = setup((m) => insertChartFrame(m, 0, "columnClustered", { xCm: 1, yCm: 1, widthCm: 12, heightCm: 7.5 })!);
    controller.applyData(["A", "B"], [{ name: "S1", color: "#123456", values: [1, 2] }]);
    expect(chartFrame(model.value!).categories).toEqual(["A", "B"]);
    expect(record).toHaveBeenCalledWith("Edit chart data");

    controller.applyOptions({ ...defaultChartOptions(), legend: "bottom", dataLabels: true });
    expect(chartFrame(model.value!).options.legend).toBe("bottom");
    expect(record).toHaveBeenCalledWith("Chart options");

    controller.applySeriesColor(0, "#abcdef");
    expect(chartFrame(model.value!).series[0].color).toBe("#abcdef");
    expect(record).toHaveBeenCalledWith("Chart series colour");
  });

  it("opens the data sheet on a supported chart but not an unsupported one (AC-14/15)", () => {
    const supported = setup((m) => insertChartFrame(m, 0, "columnClustered", { xCm: 1, yCm: 1, widthCm: 12, heightCm: 7.5 })!);
    supported.controller.openDataSheet();
    expect(supported.controller.editingData.value).toBe(true);

    const unsupported = setup((m) => {
      const id = insertChartFrame(m, 0, "columnClustered", { xCm: 1, yCm: 1, widthCm: 12, heightCm: 7.5 })!;
      chartFrame(m).unsupportedChartXml = "<c:bar3DChart/>";
      return id;
    });
    expect(unsupported.controller.isUnsupported.value).toBe(true);
    unsupported.controller.openDataSheet();
    expect(unsupported.controller.editingData.value).toBe(false);
    unsupported.controller.openOptions();
    expect(unsupported.controller.editingOptions.value).toBe(false);
  });

  it("deletes and duplicates the selected chart (toolbar)", () => {
    const { model, controller } = setup((m) => insertChartFrame(m, 0, "line", { xCm: 1, yCm: 1, widthCm: 12, heightCm: 7.5 })!);
    controller.duplicateChart();
    expect(model.value!.slides[0].frames.filter((f) => f.kind === "chart")).toHaveLength(2);
    controller.deleteChart();
    expect(model.value!.slides[0].frames.filter((f) => f.kind === "chart")).toHaveLength(1);
  });

  it("shows the rewrite notice only after the first edit of an imported chart (AC-23)", () => {
    const { controller } = setup((m) => {
      const id = insertChartFrame(m, 0, "columnClustered", { xCm: 1, yCm: 1, widthCm: 12, heightCm: 7.5 })!;
      const frame = chartFrame(m);
      frame.sourcePart = "ppt/charts/chart1.xml";
      delete frame.edited;
      m.opaque = [
        {
          name: "ppt/charts/chart1.xml",
          data: btoa("<c:chartSpace><c:chart><c:plotArea><c:trendline/></c:plotArea></c:chart></c:chartSpace>"),
        },
      ];
      return id;
    });
    // Untouched: no notice yet.
    expect(controller.rewriteNotice.value).toBe("");
    controller.applyOptions({ ...defaultChartOptions(), legend: "bottom" });
    expect(controller.rewriteNotice.value).toContain("trendline");
  });
});
