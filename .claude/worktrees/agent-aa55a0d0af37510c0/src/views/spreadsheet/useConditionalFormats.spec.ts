import { computed, effectScope, nextTick, ref, shallowRef } from "vue";
import { describe, expect, it } from "vitest";
import { parseRange } from "../../lib/spreadsheet/address";
import { addConditionalFormat, CONDITIONAL_FORMATTING_LABEL, type ConditionalFormat } from "../../lib/spreadsheet/conditionalFormats";
import { cellKey, type CellKey } from "../../lib/spreadsheet/formula/dependencies";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { useConditionalFormats } from "./useConditionalFormats";

const SCALE: ConditionalFormat = {
  id: "f",
  ranges: [parseRange("A1:A3")!],
  rules: [{ id: "r", priority: 1, kind: "colorScale", stops: [{ at: { type: "min" }, color: "#000000" }, { at: { type: "max" }, color: "#ffffff" }] }],
};

function sheet(id: string, values: number[], formats: ConditionalFormat[]): Sheet {
  return {
    id,
    name: id,
    visibility: "visible",
    kind: "grid",
    rows: values.map((value, index) => ({ index, cells: [{ col: 0, value }] })),
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    conditionalFormats: formats,
    validations: [],
  };
}

function setup() {
  const model: SpreadsheetModel = {
    kind: "spreadsheet",
    id: "w",
    title: "W",
    format: "xlsx",
    readerVersion: 10,
    sheets: [sheet("s1", [1, 2, 3], [SCALE]), sheet("s2", [5, 50], [])],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
    names: [],
    dxfs: [{ fill: { kind: "solid", color: "#C6EFCE" } }],
  };
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => workbook.value?.sheets[activeIndex.value] ?? null);
  const results: ((changed: CellKey[]) => void)[] = [];
  const steps: ((kind: "commit" | "undo" | "redo", label: string) => void)[] = [];
  const scope = effectScope();
  const view = scope.run(() =>
    useConditionalFormats({
      book: { workbook, activeSheet },
      recalc: { onResult: (listener) => results.push(listener) },
      undo: { onStep: (listener) => steps.push(listener) },
    }),
  )!;
  /** Sets column A's value at `row` on the active sheet, copy-on-write. */
  const setValue = (row: number, value: number) => {
    const at = activeIndex.value;
    const current = model.sheets[at]!;
    model.sheets = model.sheets.map((each, i) => (i === at ? { ...current, rows: current.rows.map((r) => (r.index === row ? { index: row, cells: [{ col: 0, value }] } : r)) } : each));
  };
  return { model, view, activeIndex, setValue, recalcDone: (keys: CellKey[]) => results.forEach((listener) => listener(keys)), step: (kind: "commit" | "undo" | "redo", label: string) => steps.forEach((listener) => listener(kind, label)) };
}

describe("useConditionalFormats", () => {
  it("answers the sheet on show, cell by cell", () => {
    const { view } = setup();
    expect(view.resultAt(0, 0)?.scaleColor).toBe("#000000");
    expect(view.resultAt(2, 0)?.scaleColor).toBe("#ffffff");
    expect(view.resultAt(0, 5)).toBeUndefined();
  });

  it("invalidates from each recalculation result's changed keys", () => {
    const { view, setValue, recalcDone } = setup();
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#808080");
    setValue(2, 101);
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#808080"); // still the cached pass's
    const before = view.revision.value;
    recalcDone([cellKey("s1", 2, 0)]);
    expect(view.revision.value).toBe(before + 1);
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#030303");
  });

  it("invalidates everything on undo and redo, and on a rule command's step", () => {
    const { model, view, setValue, step } = setup();
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#808080");
    setValue(2, 101);
    step("undo", "Typing");
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#030303");
    setValue(2, 3);
    step("commit", "Typing"); // an edit's own step waits for its pass's keys
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#030303");
    step("commit", CONDITIONAL_FORMATTING_LABEL);
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#808080");
    addConditionalFormat(model, "s1", [parseRange("A1:A3")!], { kind: "cellIs", operator: "greaterThan", formulas: ["1"], dxfId: 0 });
    expect(view.resultAt(1, 0)).toEqual({ dxf: { fill: { kind: "solid", color: "#C6EFCE" } } }); // the new rule ranks first and takes the fill
  });

  it("keeps one evaluator per open sheet and invalidates on a sheet switch", async () => {
    const { model, view, activeIndex } = setup();
    expect(view.resultAt(0, 0)?.scaleColor).toBe("#000000");
    model.sheets = [model.sheets[0]!, { ...model.sheets[1]!, conditionalFormats: [{ ...SCALE, ranges: [parseRange("A1:A2")!] }] }];
    const before = view.revision.value;
    activeIndex.value = 1;
    await nextTick();
    expect(view.revision.value).toBeGreaterThan(before);
    expect(view.resultAt(0, 0)?.scaleColor).toBe("#000000");
    expect(view.resultAt(1, 0)?.scaleColor).toBe("#ffffff");
    activeIndex.value = 0;
    await nextTick();
    expect(view.resultAt(2, 0)?.scaleColor).toBe("#ffffff");
  });
});
