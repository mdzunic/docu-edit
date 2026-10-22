// The chart editing controller (SPEC-092): the presentation view's wiring for
// the "Insert → Chart" gallery, the floating chart toolbar (SlideChartLayer),
// and SPEC-052's data-sheet / options dialogs — reused verbatim. Mirrors
// useTables' shape: every gesture records one SPEC-080 undo entry, mutates the
// selected ChartFrame in place through the framework-free command layer
// (lib/presentation/charts.ts), and commits write-through. The dialogs' prefill
// and the rewrite notice are fresh-array/object computeds off the selected
// frame so the shallowRef re-identification the render trio performs on commit
// re-drives them (the useSlideRender reactivity rule, AC-17a).
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFrame, ChartFrame } from "../../lib/presentation/model";
import {
  asChartType,
  NEW_CHART_HEIGHT_CM,
  NEW_CHART_WIDTH_CM,
  SINGLE_SERIES_CHART_TYPES,
  type ChartOptions,
  type ChartSeries,
} from "../../lib/core/charts";
import { deleteFrame, duplicateFrame, type FrameRect } from "../../lib/presentation/frames";
import {
  insertChartFrame,
  lostChartFeaturesOnRewrite,
  setSlideChartData,
  setSlideChartOptions,
  setSlideChartSeriesColor,
} from "../../lib/presentation/charts";

export interface SlideChartsDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  /** The selected frame, or null (stageEdit.selection.selectedFrame). */
  selectedFrame: ComputedRef<SlideFrame | null>;
  /** Re-select a frame id after a structural change (or null when deleted). */
  selectFrame: (id: string | null) => void;
  record: (label?: string) => void;
  commit: () => void;
}

export function useSlideCharts(deps: SlideChartsDeps) {
  /** The selected chart frame, or null. */
  const selectedChart = computed<ChartFrame | null>(() => {
    const frame = deps.selectedFrame.value;
    return frame && frame.kind === "chart" ? frame : null;
  });
  /** A chart the editor cannot model (3-D, combo, radar…): its edit controls
   * are disabled and double-click is a no-op (AC-15). */
  const isUnsupported = computed(() => !!selectedChart.value && selectedChart.value.unsupportedChartXml !== undefined);

  // Which dialog (if any) is open for the selected chart.
  const editingData = ref(false);
  const editingOptions = ref(false);
  function closeDialogs(): void {
    editingData.value = false;
    editingOptions.value = false;
  }
  // Selecting a different chart (or deselecting) dismisses any open dialog.
  watch(
    () => selectedChart.value?.id ?? null,
    () => closeDialogs(),
  );

  /** "Insert → Chart" (AC-13a): a 12×7.5 cm chart with sample data and default
   * options centred on the slide, selected. */
  function pickChartType(chartType = "columnClustered"): void {
    const model = deps.model.value;
    const slide = deps.currentSlide.value;
    if (!model || !slide) {
      return;
    }
    const rect: FrameRect = {
      xCm: Math.max(0, (model.slideWidthCm - NEW_CHART_WIDTH_CM) / 2),
      yCm: Math.max(0, (model.slideHeightCm - NEW_CHART_HEIGHT_CM) / 2),
      widthCm: NEW_CHART_WIDTH_CM,
      heightCm: NEW_CHART_HEIGHT_CM,
    };
    deps.record("Insert chart");
    const id = insertChartFrame(model, deps.currentIndex.value, asChartType(chartType), rect);
    deps.selectFrame(id);
    deps.commit();
  }

  /** Double-click a supported chart, or the toolbar's "Edit data" button
   * (AC-14): open the data sheet. A no-op for an unsupported chart (AC-15). */
  function openDataSheet(id?: string): void {
    if (id) {
      deps.selectFrame(id);
    }
    if (!selectedChart.value || isUnsupported.value) {
      return;
    }
    editingOptions.value = false;
    editingData.value = true;
  }
  /** The toolbar's "Chart options" button (AC-14a). */
  function openOptions(): void {
    if (!selectedChart.value || isUnsupported.value) {
      return;
    }
    editingData.value = false;
    editingOptions.value = true;
  }

  /** The data sheet's "Done" commit (AC-16). */
  function applyData(categories: string[], series: ChartSeries[]): void {
    const model = deps.model.value;
    const frame = selectedChart.value;
    if (!model || !frame) {
      return;
    }
    deps.record("Edit chart data");
    setSlideChartData(model, deps.currentIndex.value, frame.id, categories, series);
    deps.commit();
    editingData.value = false;
  }
  /** The options panel's "Apply" commit (AC-16). */
  function applyOptions(options: ChartOptions): void {
    const model = deps.model.value;
    const frame = selectedChart.value;
    if (!model || !frame) {
      return;
    }
    deps.record("Chart options");
    setSlideChartOptions(model, deps.currentIndex.value, frame.id, options);
    deps.commit();
    editingOptions.value = false;
  }
  /** The options panel's live per-series colour input (AC-16). */
  function applySeriesColor(seriesIndex: number, color: string): void {
    const model = deps.model.value;
    const frame = selectedChart.value;
    if (!model || !frame) {
      return;
    }
    deps.record("Chart series colour");
    setSlideChartSeriesColor(model, deps.currentIndex.value, frame.id, seriesIndex, color);
    deps.commit();
  }

  /** The toolbar's Delete (SPEC-082 voice: "Delete chart"). */
  function deleteChart(): void {
    const model = deps.model.value;
    const frame = selectedChart.value;
    if (!model || !frame) {
      return;
    }
    deps.record("Delete chart");
    const next = deleteFrame(model, deps.currentIndex.value, frame.id);
    deps.selectFrame(next);
    deps.commit();
  }
  /** The toolbar's Duplicate ("Duplicate chart"). */
  function duplicateChart(): void {
    const model = deps.model.value;
    const frame = selectedChart.value;
    if (!model || !frame) {
      return;
    }
    deps.record("Duplicate chart");
    const id = duplicateFrame(model, deps.currentIndex.value, frame.id);
    deps.selectFrame(id);
    deps.commit();
  }

  /** The data-sheet dialog's prefill (fresh arrays per AC-17a). */
  const dataSheetInitial = computed(() => {
    const frame = selectedChart.value;
    if (!frame) {
      return { categories: [], series: [], singleSeriesOnly: false };
    }
    return {
      categories: frame.categories.slice(),
      series: frame.series.map((s) => ({ name: s.name, color: s.color, values: s.values.slice() })),
      singleSeriesOnly: SINGLE_SERIES_CHART_TYPES.has(frame.chartType),
    };
  });

  /** The options dialog's prefill (fresh arrays per AC-17a). */
  const optionsInitial = computed(() => {
    const frame = selectedChart.value;
    if (!frame) {
      return { legend: "none", dataLabels: false, gridlines: false, series: [], hasAxes: false };
    }
    return {
      title: frame.options.title,
      categoryAxisTitle: frame.options.categoryAxisTitle,
      valueAxisTitle: frame.options.valueAxisTitle,
      legend: frame.options.legend,
      dataLabels: frame.options.dataLabels,
      gridlines: frame.options.gridlines,
      valueAxisMin: frame.options.valueAxisMin,
      valueAxisMax: frame.options.valueAxisMax,
      series: frame.series.map((s) => ({ name: s.name, color: s.color })),
      hasAxes: frame.chartType !== "pie" && frame.chartType !== "doughnut",
    };
  });

  /** The one-line rewrite notice for the first edit of an imported chart
   * (AC-23): what the regenerate-from-model rewrite could not keep. Empty for a
   * chart inserted here, an untouched or unsupported chart, or one with nothing
   * exotic to lose. */
  const rewriteNotice = computed(() => {
    const model = deps.model.value;
    const frame = selectedChart.value;
    if (!model || !frame || !frame.edited || !frame.sourcePart) {
      return "";
    }
    const lost = lostChartFeaturesOnRewrite(model, frame);
    if (lost.length === 0) {
      return "";
    }
    return `Chart rebuilt from your edits — no longer includes ${lost.join("; ")}.`;
  });

  return {
    selectedChart,
    isUnsupported,
    editingData,
    editingOptions,
    closeDialogs,
    pickChartType,
    openDataSheet,
    openOptions,
    applyData,
    applyOptions,
    applySeriesColor,
    deleteChart,
    duplicateChart,
    dataSheetInitial,
    optionsInitial,
    rewriteNotice,
  };
}

export type SlideChartsController = ReturnType<typeof useSlideCharts>;
