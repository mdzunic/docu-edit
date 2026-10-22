// SPEC-124 useStructure: the header menu's entries and labels, selecting a
// header outside the selection first, every command as ONE undo step with
// its label (bands deleted together), the insert/delete shortcuts and dialog,
// Shift+F10 on whole rows, the size dialogs, autofit, freeze, refusals as
// toasts, and an undo restoring rows, columns, merges, frozen, opaque and
// sourceXml by identity.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import type { GridRange } from "../../lib/spreadsheet/address";
import { TOO_MANY_ROWS_MESSAGE } from "../../lib/spreadsheet/bands";
import { buildMergeIndex } from "../../lib/spreadsheet/merges";
import { MAX_COLS, MAX_ROWS, type Row, type Sheet, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { NavContext, SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import type { SheetViewFields } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";
import { useStructure } from "./useStructure";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const LAST_COL = MAX_COLS - 1;
const LAST_ROW = MAX_ROWS - 1;

const wholeRows = (top: number, bottom: number): GridRange => ({ top, bottom, left: 0, right: LAST_COL });
const wholeCols = (left: number, right: number): GridRange => ({ left, right, top: 0, bottom: LAST_ROW });
const sel = (ranges: GridRange[], active = { row: ranges[0]!.top, col: ranges[0]!.left }): SheetSelection => ({ active, ranges, activeRangeIndex: 0 });

function harness(sheetFields: Partial<Sheet>, initial: SheetSelection) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, ...sheetFields }];
  model.opaque = [{ name: "xl/workbook.xml", data: "" }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>(initial);
  const toasts: string[] = [];
  const editing = ref(false);
  const scope = effectScope();
  const nav = (): NavContext => ({
    sheet: workbook.value!.sheets[0]!,
    merges: buildMergeIndex(workbook.value!.sheets[0]!.merges),
    isRowHidden: () => false,
    isColHidden: () => false,
    pageRows: 20,
    pageCols: 10,
  });
  scopes.push(scope);
  const { undo, structure } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const commands = useStructure({
      book: { workbook, activeSheet },
      viewState: {
        syncInto: vi.fn(),
        update: (_id: string, patch: SheetViewFields) => {
          if (patch.selection) {
            selection.value = patch.selection;
          }
        },
      },
      selection: {
        selection: computed(() => selection.value),
        set: (next: SheetSelection) => {
          selection.value = next;
        },
        formatContext: computed(() => ({ locale: EN_US, dateSystem: "1900" as const })),
      },
      sheetMetrics: { nav },
      editing: { undo: stack, cell: { editing: computed(() => editing.value) } },
      measurer: fixedMeasurer(7, 20),
      toast: (message) => toasts.push(message),
    });
    return { undo: stack, structure: commands };
  })!;
  const sheet = (): Sheet => workbook.value!.sheets[0]!;
  return { structure, undo, selection, toasts, editing, sheet, workbook, scope };
}

const rows = (...indices: number[]): Row[] => indices.map((index) => ({ index, cells: [{ col: 0, value: index }] }));

/** Every harness's scope — stopped after each test, so no window listener
 * outlives its test. */
const scopes: EffectScope[] = [];

afterEach(() => {
  document.body.innerHTML = "";
  scopes.splice(0).forEach((scope) => scope.stop());
});

describe("SPEC-124 useStructure — the header menu", () => {
  it("labels its entries from the band holding the clicked header", () => {
    const h = harness({}, sel([wholeRows(2, 3)]));
    h.structure.openHeaderMenu("row", 2, 10, 20);
    expect(h.structure.headerItems.value.map((item) => [item.testid, item.label, item.disabled])).toEqual([
      ["sheet-insert-before", "Insert 2 rows above", false],
      ["sheet-insert-after", "Insert 2 rows below", false],
      ["sheet-delete-band", "Delete rows 3–4", false],
      ["sheet-hide-band", "Hide rows 3–4", false],
      ["sheet-unhide-band", "Unhide rows", true],
      ["sheet-size-band", "Row height…", false],
      ["sheet-autofit-band", "Autofit row height", false],
    ]);
    h.structure.openHeaderMenu("col", 1, 0, 0);
    expect(h.structure.headerItems.value.slice(0, 4).map((item) => item.label)).toEqual(["Insert 1 column left", "Insert 1 column right", "Delete column", "Hide column"]);
  });

  it("selects a header outside the selection first", () => {
    const h = harness({}, sel([{ top: 0, left: 0, bottom: 0, right: 0 }]));
    h.structure.openHeaderMenu("col", 3, 0, 0);
    expect(h.selection.value.ranges).toEqual([wholeCols(3, 3)]);
    expect(h.structure.headerMenu.value).toEqual({ axis: "col", index: 3, x: 0, y: 0 });
  });

  it("enables Unhide when the selection spans a hidden row — select-all reaches a hidden first row", () => {
    const h = harness({ rows: [{ index: 0, cells: [], hidden: true }] }, sel([wholeRows(0, LAST_ROW)]));
    h.structure.openHeaderMenu("row", 4, 0, 0);
    expect(h.structure.headerItems.value.find((item) => item.command === "unhide")!.disabled).toBe(false);
  });

  it("inserts above as one 'Insert rows' step and moves the live selection with the rows", () => {
    const h = harness({ rows: rows(1, 2, 5) }, sel([wholeRows(2, 3)]));
    h.structure.openHeaderMenu("row", 2, 0, 0);
    h.structure.runHeaderMenu("insert-before");
    expect(h.structure.headerMenu.value).toBeNull();
    expect(h.sheet().rows.map((row) => row.index)).toEqual([1, 4, 7]);
    expect(h.undo.undoLabel.value).toBe("Insert rows");
    expect(h.selection.value.ranges).toEqual([wholeRows(4, 5)]);
    h.structure.openHeaderMenu("row", 4, 0, 0);
    h.structure.runHeaderMenu("insert-after");
    expect(h.sheet().rows.map((row) => row.index)).toEqual([1, 4, 9]);
  });

  it("deletes every band of the selection as ONE step", () => {
    const h = harness({ rows: rows(0, 2, 4, 6) }, { active: { row: 2, col: 0 }, ranges: [wholeRows(2, 2), wholeRows(4, 5)], activeRangeIndex: 0 });
    h.structure.openHeaderMenu("row", 4, 0, 0);
    expect(h.structure.headerItems.value.find((item) => item.command === "delete")!.label).toBe("Delete rows");
    h.structure.runHeaderMenu("delete");
    expect(h.sheet().rows.map((row) => [row.index, row.cells[0]!.value])).toEqual([
      [0, 0],
      [3, 6],
    ]);
    expect(h.undo.undoLabel.value).toBe("Delete rows");
    h.undo.undo();
    expect(h.sheet().rows.map((row) => row.index)).toEqual([0, 2, 4, 6]);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("hides and unhides every band as one step each", () => {
    const h = harness({}, sel([wholeCols(1, 2)]));
    h.structure.openHeaderMenu("col", 1, 0, 0);
    h.structure.runHeaderMenu("hide");
    expect(h.sheet().columns).toEqual([{ min: 1, max: 2, hidden: true }]);
    expect(h.undo.undoLabel.value).toBe("Hide columns");
    h.structure.openHeaderMenu("col", 1, 0, 0);
    h.structure.runHeaderMenu("unhide");
    expect(h.sheet().columns).toEqual([]);
    expect(h.undo.undoLabel.value).toBe("Unhide columns");
  });

  it("refuses hiding more rows that store nothing than MAX_FORMAT_MATERIALISED, with a toast and no step", () => {
    const h = harness({}, sel([wholeRows(0, LAST_ROW)]));
    h.structure.openHeaderMenu("row", 0, 0, 0);
    h.structure.runHeaderMenu("hide");
    expect(h.toasts).toEqual([TOO_MANY_ROWS_MESSAGE]);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("undo restores rows, columns, merges, frozen, opaque and sourceXml by identity", () => {
    const h = harness(
      { rows: rows(3), columns: [{ min: 0, max: 0, widthPx: 90 }], merges: [{ top: 3, left: 0, bottom: 4, right: 1 }], frozen: { rows: 4, cols: 1 }, sourceXml: "<table:table/>" },
      sel([wholeRows(1, 1)]),
    );
    const before = h.sheet();
    const opaque = h.workbook.value!.opaque;
    h.structure.openHeaderMenu("row", 1, 0, 0);
    h.structure.runHeaderMenu("insert-before");
    expect(h.sheet().rows).not.toBe(before.rows);
    expect(h.sheet().frozen).toEqual({ rows: 5, cols: 1 });
    h.undo.undo();
    const after = h.sheet();
    expect(after.rows).toBe(before.rows);
    expect(after.columns).toBe(before.columns);
    expect(after.merges).toBe(before.merges);
    expect(after.frozen).toBe(before.frozen);
    expect(after.sourceXml).toBe(before.sourceXml);
    expect(h.workbook.value!.opaque).toBe(opaque);
  });
});

/** A focusable element inside a grid, for the keys. */
function grid(): HTMLElement {
  document.body.innerHTML = `<div data-testid="spreadsheet-grid"><div id="cell" tabindex="0"></div></div>`;
  return document.getElementById("cell")!;
}

function press(target: Element, init: KeyboardEventInit): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
}

describe("SPEC-124 useStructure — the shortcuts and the insert/delete dialog", () => {
  it("Ctrl+Alt+= on whole rows inserts as many above; Ctrl+Alt+- deletes them", () => {
    const h = harness({ rows: rows(4) }, sel([wholeRows(1, 2)]));
    const target = grid();
    expect(press(target, { key: "=", code: "Equal", ctrlKey: true, altKey: true }).defaultPrevented).toBe(true);
    expect(h.sheet().rows.map((row) => row.index)).toEqual([6]);
    expect(h.undo.undoLabel.value).toBe("Insert rows");
    press(target, { key: "-", code: "Minus", ctrlKey: true, altKey: true });
    expect(h.undo.undoLabel.value).toBe("Delete rows");
    h.scope.stop();
  });

  it("on any other selection opens the dialog, whose Entire row / Entire column act on the active range", () => {
    const h = harness({ rows: [{ index: 0, cells: [{ col: 0, value: "a" }, { col: 2, value: "c" }] }] }, sel([{ top: 0, left: 1, bottom: 3, right: 2 }]));
    const target = grid();
    press(target, { key: "=", code: "Equal", ctrlKey: true, altKey: true });
    expect(h.structure.insertDelete.value).toBe("insert");
    h.structure.chooseInsertDelete("col");
    expect(h.structure.insertDelete.value).toBeNull();
    expect(h.sheet().rows[0]!.cells.map((cell) => cell.col)).toEqual([0, 4]);
    expect(h.undo.undoLabel.value).toBe("Insert columns");
    press(target, { key: "-", code: "Minus", ctrlKey: true, altKey: true });
    expect(h.structure.insertDelete.value).toBe("delete");
    h.structure.chooseInsertDelete("row");
    expect(h.sheet().rows).toEqual([]);
    h.scope.stop();
  });

  it("never acts while a cell is being edited, nor outside the grid", () => {
    const h = harness({ rows: rows(4) }, sel([wholeRows(1, 2)]));
    const target = grid();
    h.editing.value = true;
    press(target, { key: "=", code: "Equal", ctrlKey: true, altKey: true });
    h.editing.value = false;
    press(document.body, { key: "=", code: "Equal", ctrlKey: true, altKey: true });
    expect(h.undo.canUndo.value).toBe(false);
    h.scope.stop();
  });

  it("Shift+F10 on whole rows opens the header menu; on cells it is the cell menu's", () => {
    const h = harness({}, sel([{ top: 0, left: 0, bottom: 0, right: 0 }]));
    const target = grid();
    expect(press(target, { key: "F10", shiftKey: true }).defaultPrevented).toBe(false);
    expect(h.structure.headerMenu.value).toBeNull();
    h.selection.value = sel([wholeRows(3, 3)]);
    press(target, { key: "F10", shiftKey: true });
    expect(h.structure.headerMenu.value?.axis).toBe("row");
    expect(h.structure.headerMenu.value?.index).toBe(3);
    h.scope.stop();
  });
});

describe("SPEC-124 useStructure — sizes, autofit and freeze", () => {
  it("opens the size dialog with the common size, blank when the rows differ; OK resizes in one step", () => {
    const h = harness({ rows: [{ index: 1, cells: [], heightPt: 20, customHeight: true }, { index: 2, cells: [], heightPt: 20, customHeight: true }] }, sel([wholeRows(1, 2)]));
    h.structure.openSize("row", [{ start: 1, end: 2 }]);
    expect(h.structure.sizeDialog.value?.initial).toBe("20");
    h.structure.openSize("row", [{ start: 1, end: 3 }]);
    expect(h.structure.sizeDialog.value?.initial).toBe("");
    h.structure.applySize(30.1);
    expect(h.sheet().rows.map((row) => row.heightPt)).toEqual([30, 30, 30]);
    expect(h.undo.undoLabel.value).toBe("Resize rows");
  });

  it("column width in characters: 8.43 keeps its xlsx width and is 64 px; 0 hides", () => {
    const h = harness({}, sel([wholeCols(2, 2)]));
    h.structure.openSize("col", [{ start: 2, end: 2 }]);
    expect(h.structure.sizeDialog.value?.initial).toBe("8.43");
    h.structure.applySize(10.5);
    expect(h.sheet().columns).toEqual([{ min: 2, max: 2, widthPx: 78, sourceWidthChars: 11.2109375 }]);
    h.structure.openSize("col", [{ start: 2, end: 2 }]);
    expect(h.structure.sizeDialog.value?.initial).toBe("10.5");
    h.structure.applySize(0);
    expect(h.sheet().columns).toEqual([{ min: 2, max: 2, widthPx: 78, sourceWidthChars: 11.2109375, hidden: true }]);
    expect(h.undo.undoLabel.value).toBe("Resize columns");
  });

  it("autofits through the injected measurer as one 'Autofit' step", () => {
    const h = harness({ rows: [{ index: 0, cells: [{ col: 0, value: "twelve chars" }] }] }, sel([wholeCols(0, 0)]));
    h.structure.formatBands("col", "autofit");
    expect(h.sheet().columns).toEqual([{ min: 0, max: 0, widthPx: 12 * 7 + 5 }]);
    expect(h.undo.undoLabel.value).toBe("Autofit");
  });

  it("freezes the top row, the first column or at the active cell (disabled at A1), and unfreezes", () => {
    const h = harness({}, sel([{ top: 0, left: 0, bottom: 0, right: 0 }]));
    expect(h.structure.freezeState.value).toEqual({ frozen: false, atA1: true, hasGrid: true });
    h.structure.freeze("at-cell");
    expect(h.undo.canUndo.value).toBe(false);
    h.structure.freeze("top-row");
    expect(h.sheet().frozen).toEqual({ rows: 1, cols: 0 });
    expect(h.sheet().dirty).toBe(true);
    expect(h.undo.undoLabel.value).toBe("Freeze panes");
    h.selection.value = sel([{ top: 4, left: 2, bottom: 4, right: 2 }]);
    h.structure.freeze("at-cell");
    expect(h.sheet().frozen).toEqual({ rows: 4, cols: 2 });
    expect(h.structure.freezeState.value.frozen).toBe(true);
    h.structure.freeze("unfreeze");
    expect(h.sheet().frozen).toBeUndefined();
    h.structure.freeze("first-column");
    expect(h.sheet().frozen).toEqual({ rows: 0, cols: 1 });
  });
});
