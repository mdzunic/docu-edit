// SPEC-128 useSortFilter: the Data menu's sorts — the range each selection
// kind covers, each sort one "Sort" undo step, a refusal opening sheet-alert
// and changing nothing — the Sort dialog, Data ▸ Filter / Clear filters /
// Reapply filter, the popover's criteria as single undo steps handing the
// visibility-sensitive cells to recalc, the header buttons, the status text,
// and the keys (ignored from inputs and while a cell is being edited).
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { findRow, getCell, type GridRange } from "../../lib/spreadsheet/address";
import { EMPTY_FILTER_REGION_MESSAGE } from "../../lib/spreadsheet/filter";
import { cellKey } from "../../lib/spreadsheet/formula/dependencies";
import { buildMergeIndex } from "../../lib/spreadsheet/merges";
import { rowIsHidden, type CellValue, type Row, type Sheet, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { NavContext, SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { SORT_REFUSAL_MESSAGES } from "../../lib/spreadsheet/sort";
import { useSortFilter } from "./useSortFilter";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const scopes: EffectScope[] = [];
const nodes: Element[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  nodes.splice(0).forEach((node) => node.remove());
});

/** Rows from a table of values, row 0 first (null: no cell). */
function rowsOf(table: CellValue[][]): Row[] {
  return table.map((values, index) => ({
    index,
    cells: values.flatMap((value, col) => (value === null ? [] : [{ col, value }])),
  }));
}

const SALES: CellValue[][] = [
  ["Region", "Amount"],
  ["East", 30],
  ["West", 10],
  ["North", 20],
  ["West", 40],
];

const single = (row: number, col: number): SheetSelection => ({
  active: { row, col },
  ranges: [{ top: row, left: col, bottom: row, right: col }],
  activeRangeIndex: 0,
});

function harness(rows: Row[], sheetFields: Partial<Sheet> = {}) {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows, ...sheetFields }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref<SheetSelection>(single(0, 0));
  const editing = ref(false);
  const recalcCalls: string[][] = [];
  const scope = effectScope();
  scopes.push(scope);
  const merges = computed(() => buildMergeIndex(activeSheet.value?.merges ?? []));
  const nav = (): NavContext => ({
    sheet: activeSheet.value!,
    merges: merges.value,
    isRowHidden: (r) => rowIsHidden(findRow(activeSheet.value!, r)),
    isColHidden: () => false,
    pageRows: 20,
    pageCols: 10,
  });
  const locale = SPREADSHEET_LOCALES["en-US"];
  const { undo, sf } = scope.run(() => {
    const stack = useSpreadsheetUndo({
      model: workbook,
      recalc: (keys, done) => {
        recalcCalls.push(keys);
        done([]);
      },
      onApplied: () => triggerRef(workbook),
      activateSheet: () => {},
    });
    const sortFilter = useSortFilter({
      book: { workbook, activeSheet },
      viewState: { syncInto: vi.fn() },
      selection: {
        selection: computed(() => selection.value),
        set: (next) => (selection.value = next),
        formatContext: computed(() => ({ locale, dateSystem: "1900" as const })),
      },
      sheetMetrics: { nav, merges },
      editing: { undo: stack, cell: { editing: computed(() => editing.value) } },
      locale: () => locale,
    });
    return { undo: stack, sf: sortFilter };
  })!;
  const sheet = () => workbook.value!.sheets[0]!;
  /** A column's values over rows `from`..`to`. */
  const column = (col: number, from = 1, to = 4) =>
    Array.from({ length: to - from + 1 }, (_, i) => getCell(sheet(), from + i, col)?.value ?? null);
  const hiddenRows = () => sheet().rows.filter((row) => row.filterHidden).map((row) => row.index);
  return { sf, undo, sheet, column, hiddenRows, selection, editing, recalcCalls, workbook };
}

/** A grid element in the document, for the keys (inGrid reads the DOM). */
function gridTarget(): HTMLElement {
  const grid = document.createElement("div");
  grid.dataset.testid = "spreadsheet-grid";
  const inner = document.createElement("div");
  grid.append(inner);
  document.body.append(grid);
  nodes.push(grid);
  return inner;
}

function press(target: EventTarget, init: KeyboardEventInit): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
}

const A1_B5: GridRange = { top: 0, left: 0, bottom: 4, right: 1 };

describe("SPEC-128 useSortFilter — sorting", () => {
  it("Sort A to Z on one cell sorts its current region below the guessed header by that column, as one 'Sort' step that undo restores", () => {
    const h = harness(rowsOf(SALES));
    const before = h.sheet().rows;
    h.selection.value = single(2, 1);
    h.sf.sortAscending();
    expect(h.column(1)).toEqual([10, 20, 30, 40]);
    expect(h.column(0)).toEqual(["West", "North", "East", "West"]);
    expect(getCell(h.sheet(), 0, 0)?.value).toBe("Region");
    expect(h.undo.undoLabel.value).toBe("Sort");
    expect(h.sheet().sortState).toMatchObject({ ref: { top: 1, left: 0, bottom: 4, right: 1 }, keys: [{ col: 1, descending: false }] });
    h.undo.undo();
    expect(h.sheet().rows).toBe(before);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("Sort Z to A on a multi-cell selection sorts exactly that selection", () => {
    const h = harness(rowsOf(SALES));
    h.selection.value = { active: { row: 1, col: 0 }, ranges: [{ top: 1, left: 0, bottom: 4, right: 0 }], activeRangeIndex: 0 };
    h.sf.sortDescending();
    expect(h.column(0)).toEqual(["West", "West", "North", "East"]);
    expect(h.column(1)).toEqual([30, 10, 20, 40]);
  });

  it("a selection inside the autofilter's range sorts that range below its header and records autoFilter.sort", () => {
    const h = harness(rowsOf(SALES), { autoFilter: { range: A1_B5, columns: [] } });
    h.selection.value = { active: { row: 2, col: 1 }, ranges: [{ top: 2, left: 1, bottom: 3, right: 1 }], activeRangeIndex: 0 };
    h.sf.sortDescending();
    expect(h.column(1)).toEqual([40, 30, 20, 10]);
    expect(h.column(0)).toEqual(["West", "East", "North", "West"]);
    expect(h.sheet().autoFilter?.sort?.keys).toEqual([{ col: 1, descending: true }]);
  });

  it("Sort range… over the autofilter's range locks its header row as a header: an OK without headers still sorts below it", () => {
    const h = harness(rowsOf(SALES), { autoFilter: { range: A1_B5, columns: [] } });
    h.selection.value = single(2, 1);
    h.sf.openSortDialog();
    expect(h.sf.sortDialog.value).toMatchObject({ range: A1_B5, hasHeader: true, headerLocked: true });
    h.sf.confirmSort({ keys: [{ col: 1, descending: false }], hasHeader: false, caseSensitive: false });
    expect(h.column(0, 0)).toEqual(["Region", "West", "North", "East", "West"]);
    expect(h.column(1, 0)).toEqual(["Amount", 10, 20, 30, 40]);
    // Off the autofilter the box stays the user's.
    const plain = harness(rowsOf(SALES));
    plain.selection.value = single(2, 1);
    plain.sf.openSortDialog();
    expect(plain.sf.sortDialog.value?.headerLocked).toBe(false);
  });

  it("a multi-area selection refuses: sheet-alert states why, nothing changes and no undo step is left", () => {
    const h = harness(rowsOf(SALES));
    const before = h.sheet();
    h.selection.value = {
      active: { row: 1, col: 0 },
      ranges: [
        { top: 1, left: 0, bottom: 2, right: 0 },
        { top: 1, left: 1, bottom: 2, right: 1 },
      ],
      activeRangeIndex: 0,
    };
    h.sf.sortAscending();
    expect(h.sf.alert.value).toBe(SORT_REFUSAL_MESSAGES.multiArea);
    expect(h.sheet()).toBe(before);
    expect(h.undo.canUndo.value).toBe(false);
    h.sf.closeAlert();
    expect(h.sf.alert.value).toBeNull();
    h.sf.openSortDialog();
    expect(h.sf.sortDialog.value).toBeNull();
    expect(h.sf.alert.value).toBe(SORT_REFUSAL_MESSAGES.multiArea);
  });

  it("merged cells spanning rows refuse with the pinned message", () => {
    const h = harness(rowsOf(SALES), { merges: [{ top: 1, left: 0, bottom: 2, right: 0 }] });
    h.selection.value = single(3, 1);
    h.sf.sortAscending();
    expect(h.sf.alert.value).toBe(SORT_REFUSAL_MESSAGES.merges);
    expect(h.column(1)).toEqual([30, 10, 20, 40]);
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("Sort range… opens the dialog on the region with the header guess and both sets of column choices; OK sorts, a refusal stays inline", () => {
    const h = harness(rowsOf(SALES));
    h.selection.value = single(2, 1);
    h.sf.openSortDialog();
    const dialog = h.sf.sortDialog.value!;
    expect(dialog).toMatchObject({ range: A1_B5, hasHeader: true, keyCol: 1, error: "" });
    expect(dialog.withHeader).toEqual([
      { col: 0, label: "Region", numeric: false },
      { col: 1, label: "Amount", numeric: true },
    ]);
    expect(dialog.withoutHeader.map((choice) => choice.label)).toEqual(["Column A", "Column B"]);
    h.sf.confirmSort({ keys: [{ col: 0, descending: false }, { col: 0, descending: true }], hasHeader: true, caseSensitive: false });
    expect(h.sf.sortDialog.value?.error).toBe("Column A is used by more than one level — remove the duplicate.");
    expect(h.undo.canUndo.value).toBe(false);
    h.sf.confirmSort({ keys: [{ col: 0, descending: true }, { col: 1, descending: false }], hasHeader: true, caseSensitive: false });
    expect(h.sf.sortDialog.value).toBeNull();
    expect(h.column(0)).toEqual(["West", "West", "North", "East"]);
    expect(h.column(1)).toEqual([10, 40, 20, 30]);
    expect(h.undo.undoLabel.value).toBe("Sort");
  });

  it("does nothing while a cell is being edited", () => {
    const h = harness(rowsOf(SALES));
    const before = h.sheet();
    h.editing.value = true;
    h.selection.value = single(2, 1);
    h.sf.sortAscending();
    h.sf.toggleFilter();
    expect(h.sheet()).toBe(before);
  });
});

describe("SPEC-128 useSortFilter — filtering", () => {
  it("Filter turns the autofilter on over the current region and off again, each one 'Filter' step; an empty region refuses in sheet-alert", () => {
    const h = harness(rowsOf(SALES));
    h.selection.value = single(0, 0);
    h.sf.toggleFilter();
    expect(h.sheet().autoFilter?.range).toEqual(A1_B5);
    expect(h.sf.filterOn.value).toBe(true);
    expect(h.undo.undoLabel.value).toBe("Filter");
    h.sf.toggleFilter();
    expect(h.sheet().autoFilter).toBeUndefined();
    h.selection.value = single(10, 6);
    h.sf.toggleFilter();
    expect(h.sf.alert.value).toBe(EMPTY_FILTER_REGION_MESSAGE);
    expect(h.sheet().autoFilter).toBeUndefined();
  });

  it("the popover's OK writes a values criterion as one 'Filter column' step, hides the rejected rows and hands the SUBTOTAL cells to recalc", () => {
    const rows = rowsOf(SALES);
    rows[0]!.cells.push({ col: 3, value: 100, formula: "SUBTOTAL(9,B2:B5)" });
    const h = harness(rows, { autoFilter: { range: A1_B5, columns: [] } });
    h.sf.popover.open(0, { left: 10, top: 20 });
    const view = h.sf.popover.view.value!;
    expect(view.title).toBe("Region");
    expect(view.conditionsLabel).toBe("Text Filters");
    expect(view.sortAscLabel).toBe("Sort A to Z");
    expect(view.entries.map((entry) => entry.text)).toEqual(["East", "North", "West"]);
    expect(h.sf.popover.openColId.value).toBe(0);
    h.sf.popover.apply(view.entries.filter((entry) => entry.text !== "West").map((entry) => entry.key));
    expect(h.sf.popover.view.value).toBeNull();
    expect(h.hiddenRows()).toEqual([2, 4]);
    expect(h.sheet().autoFilter?.columns).toEqual([{ colId: 0, criteria: { kind: "values", values: ["East", "North"], blanks: false, dateGroups: [] } }]);
    expect(h.undo.undoLabel.value).toBe("Filter column");
    expect(h.recalcCalls.at(-1)).toContain(cellKey(h.sheet().id, 0, 3));
    expect(h.sf.statusText.value).toBe("2 of 4 records found");
    expect(h.sf.isRowFiltered(1)).toBe(true);
    expect(h.sf.isRowFiltered(0)).toBe(false);
    expect(h.sf.filterButtonAt(0, 0)).toMatchObject({ label: 'Filter "Region"', state: "filtered", icon: "filter", expanded: false });
    h.undo.undo();
    expect(h.hiddenRows()).toEqual([]);
    expect(h.sf.statusText.value).toBe("");
  });

  it("OK with nothing toggled writes nothing; with every entry checked again it removes the criteria", () => {
    const h = harness(rowsOf(SALES), {
      autoFilter: { range: A1_B5, columns: [{ colId: 0, criteria: { kind: "values", values: ["East"], blanks: false, dateGroups: [] } }] },
    });
    h.sf.reapplyFilter();
    expect(h.hiddenRows()).toEqual([2, 3, 4]);
    h.sf.popover.open(0, { left: 0, top: 0 });
    const entries = h.sf.popover.view.value!.entries;
    h.sf.popover.apply(entries.filter((entry) => entry.checked).map((entry) => entry.key));
    expect(h.undo.undoLabel.value).toBe("Reapply filter");
    h.sf.popover.open(0, { left: 0, top: 0 });
    h.sf.popover.apply(entries.map((entry) => entry.key));
    expect(h.sheet().autoFilter?.columns).toEqual([]);
    expect(h.hiddenRows()).toEqual([]);
  });

  it("Clear filters removes every criterion and keeps the filter, never touching a user-hidden row; Reapply re-evaluates; both do nothing without a filter", () => {
    const rows = rowsOf(SALES);
    rows[1] = { ...rows[1]!, hidden: true };
    const h = harness(rows, {
      autoFilter: { range: A1_B5, columns: [{ colId: 1, criteria: { kind: "custom", join: "and", conditions: [{ operator: "greaterThan", value: "15" }] } }] },
    });
    h.sf.reapplyFilter();
    expect(h.hiddenRows()).toEqual([2]);
    h.sf.clearFilters();
    expect(h.undo.undoLabel.value).toBe("Clear filters");
    expect(h.sheet().autoFilter).toMatchObject({ range: A1_B5, columns: [] });
    expect(h.hiddenRows()).toEqual([]);
    expect(findRow(h.sheet(), 1)?.hidden).toBe(true);
    h.sf.toggleFilter();
    const steps = h.undo.undoLabel.value;
    h.sf.clearFilters();
    h.sf.reapplyFilter();
    expect(h.undo.undoLabel.value).toBe(steps);
    expect(h.sf.alert.value).toBeNull();
  });

  it("the popover's conditions: Begins With… opens the custom dialog pre-set, Top 10… its dialog, Above Average applies at once", () => {
    const h = harness(rowsOf(SALES), { autoFilter: { range: A1_B5, columns: [] } });
    h.sf.popover.open(0, { left: 0, top: 0 });
    h.sf.popover.condition("beginsWith");
    expect(h.sf.popover.view.value).toBeNull();
    expect(h.sf.popover.custom.value).toMatchObject({ colId: 0, kind: "text", title: "Region", draft: { op1: "beginsWith", value1: "", op2: "" } });
    h.sf.popover.confirmCustom({ op1: "beginsWith", value1: "We", join: "and", op2: "", value2: "" });
    expect(h.sf.popover.custom.value).toBeNull();
    expect(h.sheet().autoFilter?.columns[0]?.criteria).toEqual({ kind: "custom", join: "and", conditions: [{ operator: "equal", value: "We*" }] });
    expect(h.hiddenRows()).toEqual([1, 3]);

    h.sf.popover.open(1, { left: 0, top: 0 });
    expect(h.sf.popover.view.value?.conditionsLabel).toBe("Number Filters");
    expect(h.sf.popover.view.value?.sortAscLabel).toBe("Sort Smallest to Largest");
    h.sf.popover.condition("top10");
    expect(h.sf.popover.top10.value).toMatchObject({ colId: 1, title: "Amount" });
    h.sf.popover.confirmTop10({ top: true, percent: false, value: "501" });
    expect(h.sf.popover.top10.value).not.toBeNull();
    h.sf.popover.confirmTop10({ top: true, percent: false, value: "1" });
    expect(h.sheet().autoFilter?.columns[1]?.criteria).toEqual({ kind: "top10", top: true, percent: false, value: 1 });

    h.sf.popover.open(1, { left: 0, top: 0 });
    h.sf.popover.condition("aboveAverage");
    expect(h.sheet().autoFilter?.columns[1]?.criteria).toEqual({ kind: "dynamic", type: "aboveAverage" });
    expect(h.undo.undoLabel.value).toBe("Filter column");
  });

  it("a preserved criterion offers only Clear filter; the popover's sorts sort the filter's data rows", () => {
    const h = harness(rowsOf(SALES), {
      autoFilter: { range: A1_B5, columns: [{ colId: 0, criteria: { kind: "preserved", xml: "<colorFilter dxfId=\"0\"/>", label: "colour" } }] },
    });
    h.sf.popover.open(0, { left: 0, top: 0 });
    expect(h.sf.popover.view.value).toMatchObject({ preserved: true, hasCriteria: true, entries: [] });
    h.sf.popover.clear();
    expect(h.sheet().autoFilter?.columns).toEqual([]);
    expect(h.undo.undoLabel.value).toBe("Clear filter");
    h.sf.popover.open(1, { left: 0, top: 0 });
    h.sf.popover.sort(true);
    expect(h.column(1)).toEqual([40, 30, 20, 10]);
    expect(h.sf.filterButtonAt(0, 1)).toMatchObject({ state: "sorted-desc", icon: "sort-descending" });
    expect(h.undo.undoLabel.value).toBe("Sort");
  });

  it("buttons: one per header cell of the range, named by the header (a blank one by its column), none where hideButton, aria-expanded with the popover", () => {
    const rows = rowsOf([
      ["Region", null, "Rep"],
      ["East", 1, "Ann"],
      ["West", 2, "Bo"],
    ]);
    const h = harness(rows, { autoFilter: { range: { top: 0, left: 0, bottom: 2, right: 2 }, columns: [{ colId: 2, hideButton: true }] } });
    expect(h.sf.filterButtonAt(0, 0)).toEqual({ label: 'Filter "Region"', state: "none", icon: "chevron-down", expanded: false });
    expect(h.sf.filterButtonAt(0, 1)?.label).toBe('Filter "Column B"');
    expect(h.sf.filterButtonAt(0, 2)).toBeNull();
    expect(h.sf.filterButtonAt(1, 0)).toBeNull();
    expect(h.sf.filterButtonAt(0, 3)).toBeNull();
    h.sf.pressFilterButton(1, { left: 0, top: 0 });
    expect(h.sf.filterButtonAt(0, 1)?.expanded).toBe(true);
    h.sf.pressFilterButton(1, { left: 0, top: 0 });
    expect(h.sf.popover.view.value).toBeNull();
  });
});

describe("SPEC-128 useSortFilter — keys", () => {
  it("Ctrl+Shift+L in the grid toggles the filter and Ctrl+Alt+L reapplies it; from an input outside the grid or while editing they do nothing", () => {
    const h = harness(rowsOf(SALES));
    const target = gridTarget();
    const outside = document.createElement("input");
    document.body.append(outside);
    nodes.push(outside);
    press(outside, { key: "L", code: "KeyL", ctrlKey: true, shiftKey: true });
    expect(h.sheet().autoFilter).toBeUndefined();
    h.editing.value = true;
    press(target, { key: "L", code: "KeyL", ctrlKey: true, shiftKey: true });
    expect(h.sheet().autoFilter).toBeUndefined();
    h.editing.value = false;
    const event = press(target, { key: "L", code: "KeyL", ctrlKey: true, shiftKey: true });
    expect(event.defaultPrevented).toBe(true);
    expect(h.sheet().autoFilter?.range).toEqual(A1_B5);
    press(target, { key: "l", code: "KeyL", ctrlKey: true, altKey: true });
    expect(h.undo.undoLabel.value).toBe("Reapply filter");
  });

  it("Alt+Down on a header cell opens its popover; Escape returns to the header cell", () => {
    const h = harness(rowsOf(SALES), { autoFilter: { range: A1_B5, columns: [] } });
    const target = gridTarget();
    h.selection.value = single(2, 1);
    expect(press(target, { key: "ArrowDown", altKey: true }).defaultPrevented).toBe(false);
    expect(h.sf.popover.view.value).toBeNull();
    h.selection.value = single(0, 1);
    expect(press(target, { key: "ArrowDown", altKey: true }).defaultPrevented).toBe(true);
    expect(h.sf.popover.view.value?.title).toBe("Amount");
    h.selection.value = single(3, 0);
    h.sf.popover.close("escape");
    expect(h.sf.popover.view.value).toBeNull();
    expect(h.selection.value.active).toEqual({ row: 0, col: 1 });
  });
});
