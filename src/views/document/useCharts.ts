// Charts (SPEC-052): the selected chart's state (selection, measured
// handles for the floating toolbar's position), insertion at the caret
// (fixed 12cm x 7.5cm, no drag-to-size draw mode — see charts.ts's
// insertChart), the chart toolbar's delete/duplicate actions, and the
// data-sheet/options-panel commit paths. Mirrors useShapes.ts's own pattern,
// trimmed: no resize/rotate handles, no drag-to-move, no multi-select
// grouping UI (Chart still participates in shapes.ts's generic z-order/
// grouping machinery at the model level, SPEC-052 AC-24, just not through a
// dedicated Chart-side gesture here).
import { computed, nextTick, ref } from "vue";
import { caretSelection, documentStart, type CommandResult, type ModelSelection } from "../../lib/commands";
import { measureImageGeometry, noImageHandleGeometry } from "../../lib/caretGeometry";
import { CM_PER_PX } from "../../lib/core/units";
import {
  chartAt,
  deleteChart,
  insertChart,
  insertChartFromTable,
  moveChart,
  setChartAltText,
  setChartData,
  setChartOptions,
  setChartSeriesColor,
  setChartWrap,
  type ChartLocation,
} from "../../lib/charts";
import { chartDataFromTableCells, SINGLE_SERIES_CHART_TYPES, type ChartFromTableResult } from "../../lib/core/charts";
import { isChart, type DocumentModel, type ImageWrapMode } from "../../lib/documentModel";
import { type ChartOptions, type ChartSeries, type ChartType } from "../../lib/core/charts";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";

function sameChartLocation(a: ChartLocation, b: ChartLocation): boolean {
  return a.section === b.section && a.block === b.block;
}

export type ChartsDeps = EditorContext;

export function useCharts(deps: ChartsDeps) {
  const { doc, contentRoot, pageRef, editable, currentSelection, applyEdit, clearCellSelection, getCaret } = deps;

  const chartSelection = ref<ChartLocation | null>(null);
  const chartHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());
  /** Opens the data-sheet dialog for chartSelection's chart (double-click,
   * or the toolbar's "Edit data" button). */
  const editingChartData = ref(false);
  /** Opens the options panel dialog for chartSelection's chart. */
  const editingChartOptions = ref(false);

  function clearChartSelection() {
    chartSelection.value = null;
    chartHandles.value = null;
    editingChartData.value = false;
    editingChartOptions.value = false;
  }

  function remeasureChartHandles() {
    const location = chartSelection.value;
    const pageEl = pageRef.value;
    const root = contentRoot.value;
    if (!location || !pageEl || !root) {
      return;
    }
    const el = root.querySelector(
      `.doc-chart-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-chart`,
    );
    chartHandles.value = el ? measureImageGeometry(pageEl, el) : null;
  }

  /** A plain click selects just this chart (no shift-multi-select gesture,
   * unlike shapes.ts's handleShapeSelect — see this module's header). */
  function handleChartSelect(section = 0, block = 0) {
    clearCellSelection();
    const location: ChartLocation = { section, block };
    if (!chartSelection.value || !sameChartLocation(chartSelection.value, location)) {
      editingChartData.value = false;
      editingChartOptions.value = false;
    }
    chartSelection.value = location;
    void nextTick().then(remeasureChartHandles);
  }

  /** Double-clicking a chart, or "Edit data" on its toolbar (AC-4, AC-5):
   * opens the data sheet — never for a chart outside the catalogue (AC-30:
   * editing is disabled for an unsupported chart type; the toolbar already
   * hides its own Edit data/Chart options buttons behind the same check,
   * this guards the dblclick path). */
  function handleChartEditData(section = 0, block = 0) {
    handleChartSelect(section, block);
    const model = doc.value;
    const chart = model ? chartAt(model, { section, block }) : null;
    if (chart?.unsupportedChartXml !== undefined) {
      return;
    }
    editingChartData.value = true;
  }

  function openChartOptions() {
    const model = doc.value;
    const location = chartSelection.value;
    const chart = model && location ? chartAt(model, location) : null;
    if (chart?.unsupportedChartXml !== undefined) {
      return;
    }
    if (!chartSelection.value) {
      return;
    }
    editingChartOptions.value = true;
  }

  /** The chart insertChart/insertChartFromTable just landed, found the same
   * way shapeInsertedLocation does — a fresh chart lands adjacent to the
   * anchor block index. */
  function chartInsertedLocation(model: DocumentModel, anchor: CaretPosition): ChartLocation | null {
    const blocks = model.sections[anchor.section]?.blocks ?? [];
    for (let i = 0; i < blocks.length; i++) {
      if (isChart(blocks[i]) && Math.abs(i - anchor.block) <= 2) {
        return { section: anchor.section, block: i };
      }
    }
    return null;
  }

  /** "Insert → Chart" (AC-2, AC-3): a fixed 12cm x 7.5cm chart with sample
   * data at the caret, immediately selected — mirrors useShapes.ts's
   * insertTextBoxAction (a fixed-size insert, not a drag-to-size gesture). */
  function insertChartAction(chartType: ChartType) {
    const root = contentRoot.value as HTMLElement | null;
    const anchor = getCaret() ?? documentStart();
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(anchor);
    const result = insertChart(model, selection, chartType);
    applyEdit(result, selection, "insert chart");
    const inserted = chartInsertedLocation(result.model, anchor);
    if (inserted) {
      chartSelection.value = inserted;
      void nextTick().then(remeasureChartHandles);
    }
  }

  /** "Insert → Chart from table" (AC-21, AC-22). Returns the unparsable
   * cell texts (if any) so the caller can report them; the caller reads the
   * selected table's cell text into `cellText` (this module holds no table
   * selection state of its own). */
  function insertChartFromTableAction(
    tableLocation: { section: number; block: number },
    chartType: ChartType,
    cellText: string[][],
  ): ChartFromTableResult {
    const data = chartDataFromTableCells(cellText);
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model) {
      return data;
    }
    const selection = currentSelection(root) ?? caretSelection(documentStart());
    const result = insertChartFromTable(model, selection, tableLocation, chartType, data);
    applyEdit(result, selection, "insert chart from table");
    const inserted = chartInsertedLocation(result.model, { ...documentStart(), section: tableLocation.section, block: tableLocation.block });
    if (inserted) {
      chartSelection.value = inserted;
      void nextTick().then(remeasureChartHandles);
    }
    return data;
  }

  function applyChartMutate(mutate: (model: DocumentModel, selection: ModelSelection, location: ChartLocation) => CommandResult, label: string) {
    const location = chartSelection.value;
    const model = doc.value;
    if (!location || !model) {
      return;
    }
    const root = contentRoot.value as HTMLElement | null;
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const result = mutate(model, selection, location);
    applyEdit(result, selection, label);
  }

  /** The data sheet's commit (AC-7..AC-10): cell edits, row/column insert/
   * delete and pasted TSV/CSV all funnel through one final categories/series
   * replacement. */
  function applyChartData(categories: string[], series: ChartSeries[]) {
    applyChartMutate((model, selection, location) => setChartData(model, selection, location, categories, series), "edit chart data");
  }

  /** The options panel's commit (AC-14..AC-20). */
  function applyChartOptions(options: ChartOptions) {
    applyChartMutate((model, selection, location) => setChartOptions(model, selection, location, options), "edit chart options");
  }

  function applyChartSeriesColor(seriesIndex: number, color: string) {
    applyChartMutate(
      (model, selection, location) => setChartSeriesColor(model, selection, location, seriesIndex, color),
      "set chart series colour",
    );
  }

  function handleChartWrap(wrap: string) {
    applyChartMutate((model, selection, location) => setChartWrap(model, selection, location, wrap as ImageWrapMode), "set chart wrap");
    // A wrap change moves the chart between text flow and an absolute
    // offset; the selection handles must re-find it (same as useShapes).
    void nextTick().then(remeasureChartHandles);
  }

  function handleChartAltText(altText: string) {
    applyChartMutate((model, selection, location) => setChartAltText(model, selection, location, altText), "set chart alt text");
  }

  function handleChartDelete() {
    const location = chartSelection.value;
    if (!location) {
      return;
    }
    applyChartMutate((model, selection, loc) => deleteChart(model, selection, loc), "delete chart");
    clearChartSelection();
  }

  function handleChartDuplicate() {
    const location = chartSelection.value;
    const model = doc.value;
    if (!location || !model) {
      return;
    }
    const chart = chartAt(model, location);
    if (!chart) {
      return;
    }
    const root = contentRoot.value as HTMLElement | null;
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const next = {
      ...model,
      sections: model.sections.map((section, sectionIndex) =>
        sectionIndex === location.section
          ? { ...section, blocks: [...section.blocks] }
          : section,
      ),
    };
    const copy = { ...chart, offsetXCm: (chart.offsetXCm ?? 0) + 0.5, offsetYCm: (chart.offsetYCm ?? 0) + 0.5 };
    delete copy.groupOrdinal;
    // A catalogue-type chart's `sourcePart` names a package part the write
    // plan keys by name (chartWritePlan.ts/odf's own) — reusing it here
    // would give two live blocks the same part, and editing just the copy
    // would silently overwrite the original's data on save. Stripping it
    // makes the copy behave exactly like a freshly inserted chart (see
    // insertChart in charts.ts): always regenerated fresh, with its own
    // part. An unsupported chart's `sourcePart` bytes are never regenerated
    // (AC-29..31), so two blocks safely sharing it is not a hazard — leave
    // it, or duplicating an unsupported chart would call chartPartXml on
    // its placeholder stand-in fields and write bogus chart content.
    if (copy.unsupportedChartXml === undefined) {
      delete copy.sourcePart;
      delete copy.edited;
    }
    next.sections[location.section].blocks.splice(location.block + 1, 0, copy);
    applyEdit({ model: next, selection }, selection, "duplicate chart");
    chartSelection.value = { section: location.section, block: location.block + 1 };
    void nextTick().then(remeasureChartHandles);
  }

  function moveChartBy(dxCm: number, dyCm: number) {
    applyChartMutate((model, selection, location) => moveChart(model, selection, location, dxCm, dyCm), "move chart");
  }

  function chartToolbarStyle(): { left: number; top: number } {
    const geometry = chartHandles.value;
    const pageEl = pageRef.value;
    if (!geometry || !pageEl) {
      return { left: 0, top: 0 };
    }
    const pageRect = pageEl.getBoundingClientRect();
    const viewportTop = pageRect.top + geometry.top;
    const above = viewportTop - 44;
    return {
      left: pageRect.left + geometry.left,
      top: above >= 0 ? above : pageRect.top + geometry.top + geometry.height + 8,
    };
  }

  /** "Selection and hit testing" recovery, mirroring shapeAtClientPoint: a
   * chart is always opaque over its full box (no unfilled/textless
   * pointer-events:none interior), so a simple bounding-rect containment
   * check is enough — no outline-tolerance test needed. */
  function chartAtClientPoint(clientX: number, clientY: number): ChartLocation | null {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return null;
    }
    const els = root.querySelectorAll(".doc-chart-anchor");
    for (const el of Array.from(els).reverse()) {
      const target = el.querySelector(".doc-chart");
      if (!(target instanceof HTMLElement)) {
        continue;
      }
      const rect = target.getBoundingClientRect();
      if (clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom) {
        const section = Number((el as HTMLElement).dataset.section);
        const block = Number((el as HTMLElement).dataset.block);
        if (Number.isInteger(section) && Number.isInteger(block)) {
          return { section, block };
        }
      }
    }
    return null;
  }

  void editable;
  void CM_PER_PX;

  /** The data sheet dialog's prefill (SPEC-052 AC-4..AC-13): the selected
   * chart's current categories/series, and whether its type renders only the
   * first series (pie/doughnut). */
  const chartDataSheetInitial = computed(() => {
    const model = doc.value;
    const location = chartSelection.value;
    const chart = model && location ? chartAt(model, location) : null;
    if (!chart) {
      return { categories: [], series: [], singleSeriesOnly: false };
    }
    return {
      categories: chart.categories,
      series: chart.series,
      singleSeriesOnly: SINGLE_SERIES_CHART_TYPES.has(chart.chartType),
    };
  });

  /** The options panel's prefill (SPEC-052 AC-14..AC-20). */
  const chartOptionsInitial = computed(() => {
    const model = doc.value;
    const location = chartSelection.value;
    const chart = model && location ? chartAt(model, location) : null;
    if (!chart) {
      return { legend: "none", dataLabels: false, gridlines: false, series: [], hasAxes: false };
    }
    return {
      title: chart.options.title,
      categoryAxisTitle: chart.options.categoryAxisTitle,
      valueAxisTitle: chart.options.valueAxisTitle,
      legend: chart.options.legend,
      dataLabels: chart.options.dataLabels,
      gridlines: chart.options.gridlines,
      valueAxisMin: chart.options.valueAxisMin,
      valueAxisMax: chart.options.valueAxisMax,
      series: chart.series,
      hasAxes: chart.chartType !== "pie" && chart.chartType !== "doughnut",
    };
  });

  return {
    chartSelection,
    chartDataSheetInitial,
    chartOptionsInitial,
    chartHandles,
    editingChartData,
    editingChartOptions,
    clearChartSelection,
    remeasureChartHandles,
    handleChartSelect,
    handleChartEditData,
    openChartOptions,
    insertChartAction,
    insertChartFromTableAction,
    applyChartData,
    applyChartOptions,
    applyChartSeriesColor,
    handleChartWrap,
    handleChartAltText,
    handleChartDelete,
    handleChartDuplicate,
    moveChartBy,
    chartToolbarStyle,
    chartAtClientPoint,
  };
}
