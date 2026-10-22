// SPEC-132 useSheetCharts: Insert ▸ Chart as one labelled step (or a toast
// and nothing inserted), and the refresh hook — charts follow every edit,
// undo, redo and manual recalculation pass, never the open pass, and a
// refresh records no step of its own.
import { computed, effectScope, ref, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { defaultChartOptions } from "../../lib/core/charts";
import { CHART_UNDO_LABELS, NO_CHART_DATA_MESSAGE } from "../../lib/spreadsheet/charts";
import { cellKey } from "../../lib/spreadsheet/formula/dependencies";
import { buildGridMetrics } from "../../lib/spreadsheet/geometry";
import type { Sheet, SheetChartContent, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useSheetCharts, type SheetChartsDeps } from "./useSheetCharts";

const VALUES: [number, number, number | string][] = [
  [0, 0, "Item"],
  [0, 1, "Plan"],
  [1, 0, "Alpha"],
  [1, 1, 4],
  [2, 0, "Beta"],
  [2, 1, 6],
];

function workbook(): SpreadsheetModel {
  const model = createBlankWorkbook();
  const rows = [0, 1, 2].map((index) => ({ index, cells: VALUES.filter(([row]) => row === index).map(([, col, value]) => ({ col, value })) }));
  model.sheets = [{ ...model.sheets[0]!, rows }];
  return model;
}

function harness(model: SpreadsheetModel, selected: { top: number; left: number; bottom: number; right: number }) {
  const workbookRef = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const record = vi.fn();
  const commit = vi.fn();
  const select = vi.fn();
  const toast = vi.fn();
  const sheetsChanged = vi.fn();
  let listener: ((changed: string[], reason?: "open" | "edit" | "manual") => void) | undefined;
  const deps = {
    book: { workbook: workbookRef, activeIndex, activeSheet: computed<Sheet | null>(() => workbookRef.value?.sheets[activeIndex.value] ?? null), sheetsChanged },
    undo: { record, commit },
    recalc: { onResult: (each: typeof listener) => (listener = each) },
    selection: { selection: ref({ active: { row: selected.top, col: selected.left }, ranges: [selected], activeRangeIndex: 0 }) },
    sheetMetrics: { metrics: computed(() => buildGridMetrics(workbookRef.value!.sheets[0]!)) },
    drawings: { selected: computed(() => null), select, remove: vi.fn(), canInsert: computed(() => true) },
    toast,
  } as unknown as SheetChartsDeps;
  const scope = effectScope();
  const charts = scope.run(() => useSheetCharts(deps))!;
  return { charts, record, commit, select, toast, sheetsChanged, fire: (changed: string[], reason?: "open" | "edit" | "manual") => listener!(changed, reason), scope };
}

const chartsOn = (model: SpreadsheetModel): SheetChartContent[] =>
  (model.sheets[0]!.drawings ?? []).flatMap((drawing) => (drawing.content.kind === "chart" ? [drawing.content] : []));

describe("useSheetCharts", () => {
  it("Insert ▸ Chart builds the chart from the selection as one \"Insert chart\" step and selects it", () => {
    const model = workbook();
    const { charts, record, commit, select, toast, scope } = harness(model, { top: 0, left: 0, bottom: 2, right: 1 });
    charts.pickChartType("columnClustered");
    expect(toast).not.toHaveBeenCalled();
    expect(record).toHaveBeenCalledTimes(1);
    expect(record.mock.calls[0]![0]).toBe(CHART_UNDO_LABELS.insert);
    expect(commit).toHaveBeenCalledTimes(1);
    const inserted = model.sheets[0]!.drawings!.at(-1)!;
    expect(inserted.content.kind === "chart" && inserted.content.visual.series[0]!.values).toEqual([4, 6]);
    expect(select).toHaveBeenCalledWith(inserted.id);
    scope.stop();
  });

  it("an empty selection with an empty current region toasts and inserts nothing (the model is untouched, so the undo stack keeps no step)", () => {
    const model = workbook();
    const sheets = model.sheets;
    const { charts, toast, scope } = harness(model, { top: 30, left: 30, bottom: 30, right: 30 });
    charts.pickChartType("line");
    expect(toast).toHaveBeenCalledWith(NO_CHART_DATA_MESSAGE);
    expect(model.sheets).toBe(sheets);
    expect(model.sheets[0]!.drawings ?? []).toEqual([]);
    scope.stop();
  });

  it("refreshes a chart whose cells a pass changed — never for the open pass — recording no step of its own", () => {
    const model = workbook();
    const sheetId = model.sheets[0]!.id;
    const content: SheetChartContent = {
      kind: "chart",
      visual: { chartType: "columnClustered", categories: ["Alpha", "Beta"], series: [{ name: "Plan", values: [4, 6], color: "#4e79a7" }], options: defaultChartOptions() },
      refs: { series: [{ name: `${model.sheets[0]!.name}!$B$1`, categories: `${model.sheets[0]!.name}!$A$2:$A$3`, values: `${model.sheets[0]!.name}!$B$2:$B$3` }], plotVisibleOnly: true },
    };
    model.sheets = [{ ...model.sheets[0]!, drawings: [{ id: "dr1", name: "Chart 1", anchor: { kind: "absolute", posEmu: { x: 0, y: 0 }, extEmu: { cx: 1, cy: 1 } }, content }] }];
    const { record, sheetsChanged, fire, scope } = harness(model, { top: 0, left: 0, bottom: 0, right: 0 });
    // The cell changes (as a committed edit leaves it).
    model.sheets = [{ ...model.sheets[0]!, rows: model.sheets[0]!.rows.map((row) => (row.index === 2 ? { ...row, cells: row.cells.map((cell) => (cell.col === 1 ? { ...cell, value: 60 } : cell)) } : row)) }];
    fire([cellKey(sheetId, 2, 1)], "open");
    expect(chartsOn(model)[0]!.visual.series[0]!.values).toEqual([4, 6]);
    fire([cellKey(sheetId, 2, 1)], "edit");
    expect(chartsOn(model)[0]!.visual.series[0]!.values).toEqual([4, 60]);
    expect(chartsOn(model)[0]!.cachesDirty).toBe(true);
    expect(sheetsChanged).toHaveBeenCalled();
    expect(record).not.toHaveBeenCalled();
    scope.stop();
  });
});
