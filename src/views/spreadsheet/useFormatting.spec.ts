// SPEC-123 useFormatting: every command is ONE undo step with its label;
// toggles and value controls read the ACTIVE cell's effective style (a Bold
// on a selection whose active cell is bold un-bolds every cell); decimal
// stepping and indent are computed from the active cell; the sheet shortcuts
// act only in the grid and never while a cell is edited; Clear contents is
// SPEC-122's Delete; a refusal records nothing and shows a toast.
import { describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import { MAX_ROWS, type CellStyle, type Row, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";
import { useFormatting } from "./useFormatting";

const EN_US = SPREADSHEET_LOCALES["en-US"];

function harness(rows: Row[], styles: CellStyle[], ranges: GridRange[] = [{ top: 0, left: 0, bottom: 1, right: 1 }]) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows }];
  model.styles = styles;
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>({ active: { row: ranges[0]!.top, col: ranges[0]!.left }, ranges, activeRangeIndex: 0 });
  const labels: string[] = [];
  const toasts: string[] = [];
  const editing = ref(false);
  const clearSelection = vi.fn();
  const scope = effectScope();
  const formatting = scope.run(() =>
    useFormatting({
      book: { workbook, activeSheet },
      viewState: { syncInto: vi.fn() },
      selection: { selection: computed(() => selection.value), formatContext: computed(() => ({ locale: EN_US, dateSystem: "1900" as const })) },
      editing: { undo: { record: (label) => labels.push(label), commit: () => workbook.value && (workbook.value = { ...workbook.value }) }, cell: { editing: computed(() => editing.value), clearSelection } },
      locale: () => EN_US,
      toast: (message) => toasts.push(message),
    }),
  )!;
  const shown = (row: number, col: number): CellStyle => cellStyle(workbook.value!, effectiveStyleId(workbook.value!.sheets[0]!, row, col));
  return { formatting, model: workbook, labels, toasts, editing, clearSelection, shown, scope };
}

describe("SPEC-123 useFormatting", () => {
  it("reflects the active cell and toggles from it: Bold on a bold active cell un-bolds every cell", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }, { col: 1, value: 2 }] }, { index: 1, cells: [{ col: 0, value: 3, styleId: 1 }] }], [{}, { font: { bold: true } }]);
    expect(h.formatting.toolbar.value.bold).toBe(true);
    h.formatting.run({ kind: "toggle", format: "bold" });
    expect(h.labels).toEqual(["Bold"]);
    for (const [row, col] of [[0, 0], [0, 1], [1, 0], [1, 1]] as const) {
      expect(h.shown(row, col).font?.bold).toBeUndefined();
    }
    h.formatting.run({ kind: "toggle", format: "bold" });
    expect(h.shown(1, 1).font?.bold).toBe(true);
    h.scope.stop();
  });

  it("labels each command and computes indent and decimals from the active cell", () => {
    const h = harness([{ index: 0, cells: [{ col: 0, value: 1.25 }] }], [{}]);
    h.formatting.run({ kind: "fill", color: "#ff0000" });
    h.formatting.run({ kind: "border", preset: "all" });
    h.formatting.run({ kind: "number-format", preset: "currency" });
    h.formatting.run({ kind: "decimals", delta: 1 });
    h.formatting.run({ kind: "indent", delta: 1 });
    h.formatting.run({ kind: "rotation", degrees: 45 });
    h.formatting.run({ kind: "rotation", degrees: 45 });
    h.formatting.run({ kind: "cell-style", name: "Good" });
    h.formatting.run({ kind: "clear", what: "formats" });
    expect(h.labels).toEqual(["Fill colour", "Borders", "Number format", "Number format", "Indent", "Text rotation", "Text rotation", "Cell style", "Clear formats"]);
    h.scope.stop();
    const g = harness([{ index: 0, cells: [{ col: 0, value: 1.25 }] }], [{}]);
    g.formatting.run({ kind: "number-format", preset: "currency" });
    g.formatting.run({ kind: "decimals", delta: 1 });
    expect(g.shown(0, 0).numberFormat).toBe('"$"#,##0.000');
    g.formatting.run({ kind: "indent", delta: 1 });
    expect(g.shown(1, 1).alignment).toEqual({ indent: 1, horizontal: "left" });
    g.formatting.run({ kind: "fill", color: null });
    expect(g.shown(0, 0).fill).toBeUndefined();
    expect(g.formatting.toolbar.value.numberFormatLabel).toBe("Currency");
    g.scope.stop();
  });

  it("Clear contents is SPEC-122's Delete, and a refusal shows a toast", () => {
    const h = harness([], [{}], [{ top: 0, left: 0, bottom: 999_999, right: 3 }]);
    h.formatting.run({ kind: "clear", what: "contents" });
    expect(h.clearSelection).toHaveBeenCalledOnce();
    h.formatting.run({ kind: "toggle", format: "bold" });
    expect(h.toasts).toEqual(["That's too many cells to format one by one — select whole columns or rows instead"]);
    h.scope.stop();
  });

  it("acts on the sheet shortcuts only in the grid and never while a cell is edited", () => {
    const h = harness([], [{}], [{ top: 0, left: 0, bottom: MAX_ROWS - 1, right: 0 }]);
    const grid = document.createElement("div");
    grid.dataset.testid = "spreadsheet-grid";
    const editor = grid.appendChild(document.createElement("div"));
    document.body.append(grid);
    const press = (target: Element, init: KeyboardEventInit): KeyboardEvent => {
      const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
      target.dispatchEvent(event);
      return event;
    };
    expect(press(editor, { key: "b", ctrlKey: true }).defaultPrevented).toBe(true);
    expect([h.labels, h.toasts, h.model.value!.sheets[0]!.columns]).toEqual([["Bold"], [], [{ min: 0, max: 0, styleId: 1 }]]);
    expect(press(document.body, { key: "i", ctrlKey: true }).defaultPrevented).toBe(false);
    h.editing.value = true;
    expect(press(editor, { key: "i", ctrlKey: true }).defaultPrevented).toBe(false);
    h.editing.value = false;
    press(editor, { key: "%", code: "Digit5", ctrlKey: true, shiftKey: true });
    expect(h.labels).toEqual(["Bold", "Number format"]);
    grid.remove();
    h.scope.stop();
  });
});
