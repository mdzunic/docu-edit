// SPEC-124 useMerges: Merge & Centre with the warning when another cell holds
// a value (Merge clears it, Cancel leaves everything), merging at once when
// nothing would be lost, the button pressed on a merged active cell with
// Unmerge its primary action, each command one labelled undo step, and the
// painter's confirmDiscard.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import { buildMergeIndex } from "../../lib/spreadsheet/merges";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { cellStyle } from "../../lib/spreadsheet/styles";
import { useMerges } from "./useMerges";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const scopes: EffectScope[] = [];
afterEach(() => scopes.splice(0).forEach((scope) => scope.stop()));

function harness(rows: Row[], range: GridRange, merges: GridRange[] = []) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows, merges }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>({ active: { row: range.top, col: range.left }, ranges: [range], activeRangeIndex: 0 });
  const toasts: string[] = [];
  const readOnly = ref(false);
  const scope = effectScope();
  scopes.push(scope);
  const { undo, merge } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const merging = useMerges({
      book: { workbook, activeSheet },
      viewState: { syncInto: vi.fn() },
      selection: { selection: computed(() => selection.value) },
      sheetMetrics: { merges: computed(() => buildMergeIndex(activeSheet.value?.merges ?? [])) },
      editing: { undo: stack, cell: { editing: computed(() => false) }, readOnly: computed(() => readOnly.value) },
      toast: (message) => toasts.push(message),
    });
    return { undo: stack, merge: merging };
  })!;
  const sheet = () => workbook.value!.sheets[0]!;
  return { merge, undo, sheet, workbook, selection, toasts, readOnly };
}

const A1_C1: GridRange = { top: 0, left: 0, bottom: 0, right: 2 };

describe("SPEC-124 useMerges", () => {
  it("warns before Merge & Centre clears B1's value; Merge merges, clears it and centres A1, as one 'Merge cells' step", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: "Title" }, { col: 1, value: "Qty" }] }], A1_C1);
    h.merge.merge("centre");
    expect(h.merge.warning.value).not.toBeNull();
    expect(h.sheet().merges).toEqual([]);
    h.merge.confirm();
    expect(h.merge.warning.value).toBeNull();
    expect(h.sheet().merges).toEqual([A1_C1]);
    const [anchor, ...rest] = h.sheet().rows[0]!.cells;
    expect(rest).toEqual([]);
    expect(cellStyle(h.workbook.value!, anchor!.styleId ?? 0).alignment).toEqual({ horizontal: "center" });
    expect(h.undo.undoLabel.value).toBe("Merge cells");
  });

  it("Cancel leaves everything as it was", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: "Title" }, { col: 1, value: "Qty" }] }], A1_C1);
    const sheet = h.sheet();
    h.merge.merge("cells");
    h.merge.cancel();
    expect(h.sheet()).toBe(sheet);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("merges at once when only the upper-left cell holds a value", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: "Title" }] }], A1_C1);
    h.merge.merge("across");
    expect(h.merge.warning.value).toBeNull();
    expect(h.sheet().merges).toEqual([A1_C1]);
  });

  it("shows pressed on a merged active cell, its primary action then Unmerge ('Unmerge cells')", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: "Title" }] }], A1_C1, [A1_C1]);
    expect(h.merge.merged.value).toBe(true);
    h.merge.primary();
    expect(h.sheet().merges).toEqual([]);
    expect(h.undo.undoLabel.value).toBe("Unmerge cells");
    expect(h.merge.merged.value).toBe(false);
    h.merge.primary();
    expect(h.sheet().merges).toEqual([A1_C1]);
  });

  it("confirmDiscard runs at once without a discard, after Merge with one", () => {
    const h = harness([], A1_C1);
    const proceed = vi.fn();
    h.merge.confirmDiscard(false, proceed);
    expect(proceed).toHaveBeenCalledTimes(1);
    h.merge.confirmDiscard(true, proceed);
    expect(proceed).toHaveBeenCalledTimes(1);
    h.merge.confirm();
    expect(proceed).toHaveBeenCalledTimes(2);
  });

  it("SPEC-138: disabled while the workbook is read-only — no merge, unmerge or pending warning applies", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: "Title" }, { col: 1, value: "Qty" }] }], A1_C1);
    expect(h.merge.disabled.value).toBe(false);
    h.merge.merge("centre");
    expect(h.merge.warning.value).not.toBeNull();
    h.readOnly.value = true;
    expect(h.merge.disabled.value).toBe(true);
    const sheet = h.sheet();
    h.merge.confirm();
    h.merge.merge("across");
    h.merge.primary();
    expect(h.merge.warning.value).toBeNull();
    expect(h.sheet()).toBe(sheet);
    expect(h.undo.canUndo.value).toBe(false);

    const merged = harness([{ index: 0, cells: [{ col: 0, value: "Title" }] }], A1_C1, [A1_C1]);
    merged.readOnly.value = true;
    merged.merge.unmerge();
    merged.merge.primary();
    expect(merged.sheet().merges).toEqual([A1_C1]);
    expect(merged.undo.canUndo.value).toBe(false);
  });

  it("toasts a refusal: a merge partly over an array", () => {
    const h = harness([{ index: 0, cells: [{ col: 1, value: 2, formula: "A1:A2*2", arrayRange: { top: 0, left: 1, bottom: 1, right: 1 } }] }], { top: 0, left: 0, bottom: 0, right: 1 });
    h.merge.merge("cells");
    h.merge.confirm();
    expect(h.toasts).toEqual(["You can't change part of an array"]);
    expect(h.undo.canUndo.value).toBe(false);
  });
});
