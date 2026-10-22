// SPEC-125 useRangeMove: a drag starts only on a single range while Ready;
// the preview is the snapped destination (kept inside the grid) with its
// address; a drop moves the cells as ONE "Move cells" step — formulas
// pointing into the block re-pointed — or, as a copy, pastes them as ONE
// "Paste" step with relative references shifted; a drop onto data outside
// the source asks first (Replace / Cancel); a refused drop toasts and
// records nothing; Escape, a drop outside the grid and a drop where it
// started do nothing.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import { getCell } from "../../lib/spreadsheet/address";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import { PASTE_MERGE_MESSAGE } from "../../lib/spreadsheet/clipboard";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useRangeMove } from "./useRangeMove";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
});

const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

/** D1 = =B2+1; B2 = 1 (bold), C2 = 2, B3 = =B2*10; E10 = "occupied";
 * H2:I2 merged. */
function rowsOf(): Row[] {
  return [
    { index: 0, cells: [{ col: 3, value: 2, formula: "B2+1" }] },
    { index: 1, cells: [{ col: 1, value: 1, styleId: 1 }, { col: 2, value: 2 }, { col: 7, value: "merged" }] },
    { index: 2, cells: [{ col: 1, value: 10, formula: "B2*10" }] },
    { index: 9, cells: [{ col: 4, value: "occupied" }] },
  ];
}

function harness() {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.styles = [{}, { font: { bold: true } }];
  model.sheets = [{ ...model.sheets[0]!, rows: rowsOf(), merges: [range(1, 7, 1, 8)] }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>({ active: { row: 1, col: 1 }, ranges: [range(1, 1, 2, 2)], activeRangeIndex: 0 });
  const editing = ref(false);
  const toasts: string[] = [];
  const scope = effectScope();
  scopes.push(scope);
  const { undo, move } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const controller = useRangeMove({
      book: { workbook, activeSheet },
      viewState: { syncInto: vi.fn() },
      selection: {
        selection: computed(() => selection.value),
        set: (next: SheetSelection) => {
          selection.value = next;
        },
        formatContext: computed(() => ({ locale: EN_US, dateSystem: "1900" as const })),
      },
      editing: { undo: stack, cell: { editing: computed(() => editing.value) } },
      toast: (message) => toasts.push(message),
      measurer: fixedMeasurer(7, 20),
    });
    return { undo: stack, move: controller };
  })!;
  const cell = (r: number, c: number) => getCell(workbook.value!.sheets[0]!, r, c);
  /** Presses the outline at B2 and drags so B2 lands on `to`. */
  const dragTo = (to: { row: number; col: number }): void => {
    expect(move.begin({ row: 1, col: 1 })).toBe(true);
    move.dragTo(to);
  };
  return { move, undo, selection, editing, toasts, cell, dragTo };
}

describe("SPEC-125 useRangeMove — the drag", () => {
  it("starts only on a single range while Ready; hover is the grid's to set", () => {
    const h = harness();
    h.move.setHover(true);
    expect(h.move.hover.value).toBe(true);
    h.editing.value = true;
    expect(h.move.begin({ row: 1, col: 1 })).toBe(false);
    h.editing.value = false;
    h.selection.value = { active: { row: 0, col: 0 }, ranges: [range(0, 0, 0, 0), range(4, 4, 4, 4)], activeRangeIndex: 0 };
    expect(h.move.begin({ row: 0, col: 0 })).toBe(false);
    expect(h.move.drag.value).toBeNull();
  });

  it("previews the snapped destination with its address, kept inside the grid", () => {
    const h = harness();
    expect(h.move.begin({ row: 2, col: 2 })).toBe(true); // pressed on C3's edge
    expect(h.move.preview.value).toBeNull();
    h.move.dragTo({ row: 5, col: 5 });
    expect(h.move.preview.value).toEqual({ range: range(4, 4, 5, 5), address: "E5:F6" });
    h.move.dragTo({ row: 0, col: 0 });
    expect(h.move.preview.value).toEqual({ range: range(0, 0, 1, 1), address: "A1:B2" });
  });

  it("Escape, a drop outside the grid, or a drop where it started do nothing", () => {
    const h = harness();
    h.dragTo({ row: 4, col: 4 });
    document.body.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(h.move.drag.value).toBeNull();
    h.move.drop({ copy: false, outside: false });
    h.dragTo({ row: 4, col: 4 });
    h.move.drop({ copy: false, outside: true });
    h.dragTo({ row: 1, col: 1 });
    h.move.drop({ copy: false, outside: false });
    expect(h.undo.canUndo.value).toBe(false);
    expect(h.cell(1, 1)?.value).toBe(1);
  });
});

describe("SPEC-125 useRangeMove — the drop", () => {
  it("moves the cells as one 'Move cells' step, re-pointing the formulas that refer to them", () => {
    const h = harness();
    h.dragTo({ row: 4, col: 4 });
    h.move.drop({ copy: false, outside: false });
    expect(h.undo.undoLabel.value).toBe("Move cells");
    expect([h.cell(1, 1), h.cell(1, 2), h.cell(2, 1)]).toEqual([undefined, undefined, undefined]);
    expect(h.cell(4, 4)).toEqual({ col: 4, value: 1, styleId: 1 });
    expect(h.cell(5, 4)?.formula).toBe("E5*10");
    expect(h.cell(0, 3)?.formula).toBe("E5+1");
    expect(h.selection.value.ranges).toEqual([range(4, 4, 5, 5)]);
    h.undo.undo();
    expect(h.cell(1, 1)).toEqual({ col: 1, value: 1, styleId: 1 });
    expect(h.cell(0, 3)?.formula).toBe("B2+1");
  });

  it("a copy drop pastes as one 'Paste' step: the source stays, relative references shift, nothing is re-pointed", () => {
    const h = harness();
    h.dragTo({ row: 4, col: 4 });
    h.move.drop({ copy: true, outside: false });
    expect(h.undo.undoLabel.value).toBe("Paste");
    expect(h.cell(1, 1)).toEqual({ col: 1, value: 1, styleId: 1 });
    expect(h.cell(4, 4)).toEqual({ col: 4, value: 1, styleId: 1 });
    expect(h.cell(5, 4)?.formula).toBe("E5*10");
    expect(h.cell(0, 3)?.formula).toBe("B2+1");
    expect(h.selection.value.ranges).toEqual([range(4, 4, 5, 5)]);
  });

  it("onto data outside the source it asks first — Cancel keeps everything, Replace moves", () => {
    const h = harness();
    h.dragTo({ row: 8, col: 4 }); // E9:F10 holds E10
    h.move.drop({ copy: false, outside: false });
    expect(h.move.overwrite.value).not.toBeNull();
    expect(h.undo.canUndo.value).toBe(false);
    h.move.cancelOverwrite();
    expect(h.move.overwrite.value).toBeNull();
    expect(h.cell(9, 4)?.value).toBe("occupied");
    h.dragTo({ row: 8, col: 4 });
    h.move.drop({ copy: false, outside: false });
    h.move.confirmOverwrite();
    expect(h.move.overwrite.value).toBeNull();
    expect(h.undo.undoLabel.value).toBe("Move cells");
    expect(h.cell(8, 4)?.value).toBe(1);
    expect(h.cell(9, 4)?.formula).toBe("E9*10");
  });

  it("overlapping only its own source it does not ask", () => {
    const h = harness();
    h.dragTo({ row: 2, col: 1 }); // B3:C4 overlaps B3, a source cell
    h.move.drop({ copy: false, outside: false });
    expect(h.move.overwrite.value).toBeNull();
    expect([h.cell(1, 1), h.cell(2, 1)?.value, h.cell(2, 2)?.value]).toEqual([undefined, 1, 2]);
    expect(h.cell(3, 1)?.formula).toBe("B3*10");
  });

  it("a drop partly over a merge is refused with a toast, as a move and as a copy", () => {
    const h = harness();
    for (const copy of [false, true]) {
      h.dragTo({ row: 1, col: 8 }); // I2:J3 crosses H2:I2
      h.move.drop({ copy, outside: false });
    }
    expect(h.toasts).toEqual([PASTE_MERGE_MESSAGE, PASTE_MERGE_MESSAGE]);
    expect(h.undo.canUndo.value).toBe(false);
    expect(h.move.overwrite.value).toBeNull();
  });
});
