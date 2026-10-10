// SPEC-125 useFillHandle: the drag's target (extend along the axis moved
// further, shrink back inside); the handle only on a single range while
// Ready; release fills as ONE "Fill" step — the series a plain drag fills,
// Ctrl (Windows/Linux) or Alt toggling copy ↔ series — and shows the chip
// with the applied option pressed; a chip option re-applies the fill as the
// SAME undo step; Escape cancels a drag and dismisses the chip, as do any
// key, a press elsewhere and another sheet; a shrink clears the cells left
// behind; double-click fills to the guide column's extent; Ctrl+D / Ctrl+R
// in the grid are "Fill down" / "Fill right"; a refusal toasts and records
// nothing.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import { getCell } from "../../lib/spreadsheet/address";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import { PASTE_MERGE_MESSAGE } from "../../lib/spreadsheet/clipboard";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { fillDragTarget, useFillHandle } from "./useFillHandle";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  document.body.innerHTML = "";
});

const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

/** A keydown dispatched on `on`, as the browser would. */
function press(on: Element, init: KeyboardEventInit): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  on.dispatchEvent(event);
  return event;
}

/** A1:A5 = 1..5, B1 = =A1*2, D2 = "Mon", E2 = 1 (bold), F2 = "x". */
function rowsOf(): Row[] {
  return [
    { index: 0, cells: [{ col: 0, value: 1 }, { col: 1, value: 2, formula: "A1*2" }] },
    { index: 1, cells: [{ col: 0, value: 2 }, { col: 3, value: "Mon" }, { col: 4, value: 1, styleId: 1 }, { col: 5, value: "x" }] },
    { index: 2, cells: [{ col: 0, value: 3 }] },
    { index: 3, cells: [{ col: 0, value: 4 }] },
    { index: 4, cells: [{ col: 0, value: 5 }] },
  ];
}

function harness(merges: GridRange[] = []) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.styles = [{}, { font: { bold: true } }];
  model.sheets = [{ ...model.sheets[0]!, rows: rowsOf(), merges }, { ...model.sheets[0]!, id: "other", name: "Sheet2", rows: [] }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => workbook.value?.sheets[activeIndex.value] ?? null);
  const selection = ref<SheetSelection>({ active: { row: 1, col: 3 }, ranges: [range(1, 3, 1, 3)], activeRangeIndex: 0 });
  const editing = ref(false);
  const toasts: string[] = [];
  const grid = document.createElement("div");
  grid.dataset.testid = "spreadsheet-grid";
  const target = document.createElement("div");
  grid.append(target);
  document.body.append(grid);
  const scope = effectScope();
  scopes.push(scope);
  const { undo, fill } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const handle = useFillHandle({
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
      locale: () => EN_US,
      toast: (message) => toasts.push(message),
      measurer: fixedMeasurer(7, 20),
    });
    return { undo: stack, fill: handle };
  })!;
  const select = (r: GridRange): void => {
    selection.value = { active: { row: r.top, col: r.left }, ranges: [r], activeRangeIndex: 0 };
  };
  const value = (r: number, c: number) => getCell(workbook.value!.sheets[0]!, r, c)?.value;
  const column = (c: number, from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => value(from + i, c));
  /** A drag of the handle from the selection to `to`, released with `mods`. */
  const dragFill = (to: { row: number; col: number }, mods = { ctrl: false, alt: false }): void => {
    expect(fill.begin()).toBe(true);
    fill.dragTo(to);
    fill.release(mods);
  };
  return { fill, undo, workbook, selection, select, editing, toasts, target, value, column, dragFill, activeIndex };
}

describe("SPEC-125 fillDragTarget", () => {
  const seed = range(2, 2, 3, 3);
  it("extends along the axis the pointer moved further on, either way", () => {
    expect(fillDragTarget(seed, { row: 9, col: 4 })).toEqual(range(2, 2, 9, 3));
    expect(fillDragTarget(seed, { row: 4, col: 9 })).toEqual(range(2, 2, 3, 9));
    expect(fillDragTarget(seed, { row: 0, col: 2 })).toEqual(range(0, 2, 3, 3));
    expect(fillDragTarget(seed, { row: 3, col: 0 })).toEqual(range(2, 0, 3, 3));
  });

  it("shrinks back inside the seed, never past its first row or column", () => {
    expect(fillDragTarget(seed, { row: 2, col: 3 })).toEqual(range(2, 2, 2, 3));
    expect(fillDragTarget(seed, { row: 3, col: 2 })).toEqual(range(2, 2, 3, 2));
    expect(fillDragTarget(seed, { row: 3, col: 3 })).toEqual(seed);
  });
});

describe("SPEC-125 useFillHandle — the drag and the chip", () => {
  it("sits on a single range while Ready, and only then", () => {
    const h = harness();
    expect(h.fill.seed.value).toEqual(range(1, 3, 1, 3));
    h.editing.value = true;
    expect(h.fill.seed.value).toBeNull();
    expect(h.fill.begin()).toBe(false);
    h.editing.value = false;
    h.selection.value = { active: { row: 0, col: 0 }, ranges: [range(0, 0, 0, 0), range(3, 3, 3, 3)], activeRangeIndex: 0 };
    expect(h.fill.seed.value).toBeNull();
  });

  it("previews the target while dragging and fills the series on release: one 'Fill' step, the chip on Fill series", () => {
    const h = harness();
    expect(h.fill.begin()).toBe(true);
    expect(h.fill.preview.value).toBeNull();
    h.fill.dragTo({ row: 4, col: 3 });
    expect(h.fill.preview.value).toEqual(range(1, 3, 4, 3));
    h.fill.release({ ctrl: false, alt: false });
    expect(h.fill.preview.value).toBeNull();
    expect(h.column(3, 1, 4)).toEqual(["Mon", "Tue", "Wed", "Thu"]);
    expect(h.undo.undoLabel.value).toBe("Fill");
    expect(h.selection.value.ranges).toEqual([range(1, 3, 4, 3)]);
    expect(h.fill.chip.value).toMatchObject({ seed: range(1, 3, 1, 3), target: range(1, 3, 4, 3), applied: "series" });
    h.undo.undo();
    expect(h.column(3, 2, 4)).toEqual([undefined, undefined, undefined]);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("a lone number copies; Ctrl (Windows/Linux) or Alt held at release toggles to the series", () => {
    const h = harness();
    h.select(range(1, 4, 1, 4));
    h.dragFill({ row: 3, col: 4 });
    expect(h.column(4, 1, 3)).toEqual([1, 1, 1]);
    expect(h.fill.chip.value?.applied).toBe("copy");
    h.undo.undo();
    h.select(range(1, 4, 1, 4));
    h.dragFill({ row: 3, col: 4 }, { ctrl: true, alt: false });
    expect(h.column(4, 1, 3)).toEqual([1, 2, 3]);
    h.undo.undo();
    h.select(range(1, 3, 1, 3));
    h.dragFill({ row: 3, col: 3 }, { ctrl: false, alt: true });
    expect(h.column(3, 1, 3)).toEqual(["Mon", "Mon", "Mon"]);
  });

  it("a chip option re-applies the fill as the same undo step, and is pressed after", () => {
    const h = harness();
    h.select(range(1, 4, 1, 4));
    h.dragFill({ row: 3, col: 4 });
    h.fill.choose("series");
    expect(h.column(4, 1, 3)).toEqual([1, 2, 3]);
    expect(h.fill.chip.value?.applied).toBe("series");
    h.fill.choose("formats");
    expect(h.column(4, 2, 3)).toEqual([null, null]); // styled blanks
    expect(getCell(h.workbook.value!.sheets[0]!, 3, 4)?.styleId).toBe(1);
    h.fill.choose("no-formats");
    expect(h.column(4, 2, 3)).toEqual([1, 1]);
    expect(getCell(h.workbook.value!.sheets[0]!, 3, 4)?.styleId).toBeUndefined();
    expect(h.undo.undoLabel.value).toBe("Fill");
    h.undo.undo();
    expect(h.undo.canUndo.value).toBe(false);
    expect(h.column(4, 2, 3)).toEqual([undefined, undefined]);
  });

  it("Escape cancels a drag; Escape, any other key, a press elsewhere or another sheet dismisses the chip", async () => {
    const h = harness();
    expect(h.fill.begin()).toBe(true);
    h.fill.dragTo({ row: 5, col: 3 });
    press(h.target, { key: "Escape" });
    expect(h.fill.drag.value).toBeNull();
    h.fill.release({ ctrl: false, alt: false });
    expect(h.undo.canUndo.value).toBe(false);

    const chipped = (): void => {
      h.select(range(1, 3, 1, 3));
      h.dragFill({ row: 2, col: 3 });
      expect(h.fill.chip.value).not.toBeNull();
    };
    chipped();
    press(h.target, { key: "Escape" });
    expect(h.fill.chip.value).toBeNull();
    chipped();
    press(h.target, { key: "Shift" });
    expect(h.fill.chip.value).not.toBeNull();
    press(h.target, { key: "a" });
    expect(h.fill.chip.value).toBeNull();
    chipped();
    const chip = document.createElement("div");
    chip.dataset.testid = "sheet-fill-options";
    const option = document.createElement("button");
    chip.append(option);
    document.body.append(chip);
    option.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(h.fill.chip.value).not.toBeNull();
    h.target.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(h.fill.chip.value).toBeNull();
    chipped();
    h.activeIndex.value = 1;
    await nextTick();
    expect(h.fill.chip.value).toBeNull();
  });

  it("a shrink clears the contents of the cells left behind, without a chip", () => {
    const h = harness();
    h.select(range(0, 0, 4, 0));
    h.dragFill({ row: 2, col: 0 });
    expect(h.column(0, 0, 4)).toEqual([1, 2, 3, undefined, undefined]);
    expect(h.fill.chip.value).toBeNull();
    expect(h.selection.value.ranges).toEqual([range(0, 0, 2, 0)]);
    expect(h.undo.undoLabel.value).toBe("Fill");
  });

  it("refuses a fill whose band crosses a merge, with a toast and no step", () => {
    const h = harness([range(3, 3, 3, 4)]);
    h.dragFill({ row: 4, col: 3 });
    expect(h.toasts).toEqual([PASTE_MERGE_MESSAGE]);
    expect(h.undo.canUndo.value).toBe(false);
    expect(h.fill.chip.value).toBeNull();
  });
});

describe("SPEC-125 useFillHandle — double-click, fill down and right", () => {
  it("double-click fills down beside the guide column, formulas shifted; without adjacent data nothing", () => {
    const h = harness();
    h.select(range(0, 1, 0, 1));
    h.fill.doubleClick();
    // The harness's recalc is a stub: the formulas are what to read.
    const formulas = Array.from({ length: 6 }, (_, row) => getCell(h.workbook.value!.sheets[0]!, row, 1)?.formula);
    expect(formulas).toEqual(["A1*2", "A2*2", "A3*2", "A4*2", "A5*2", undefined]);
    expect(h.undo.undoLabel.value).toBe("Fill");
    expect(h.fill.chip.value?.target).toEqual(range(0, 1, 4, 1));
    h.undo.undo();
    h.select(range(10, 8, 10, 8));
    h.fill.doubleClick();
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("Ctrl+D and Ctrl+R in the grid fill down and right, each one step; outside the grid or while editing they do nothing", () => {
    const h = harness();
    h.select(range(1, 5, 3, 5));
    const down = press(h.target, { key: "d", code: "KeyD", ctrlKey: true });
    expect(down.defaultPrevented).toBe(true);
    expect(h.column(5, 1, 3)).toEqual(["x", "x", "x"]);
    expect(h.undo.undoLabel.value).toBe("Fill down");
    h.select(range(1, 5, 1, 7));
    const right = press(h.target, { key: "r", code: "KeyR", ctrlKey: true });
    expect(right.defaultPrevented).toBe(true);
    expect([h.value(1, 6), h.value(1, 7)]).toEqual(["x", "x"]);
    expect(h.undo.undoLabel.value).toBe("Fill right");
    h.undo.undo();
    h.undo.undo();
    const outside = document.createElement("input");
    document.body.append(outside);
    h.select(range(1, 5, 3, 5));
    expect(press(outside, { key: "d", code: "KeyD", ctrlKey: true }).defaultPrevented).toBe(false);
    h.editing.value = true;
    expect(press(h.target, { key: "d", code: "KeyD", ctrlKey: true }).defaultPrevented).toBe(false);
    expect(h.undo.canUndo.value).toBe(false);
  });
});
