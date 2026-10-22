// SPEC-133's useNames: the Name Box (references navigate, names and tables
// select, new names are defined as one "Define name" step, invalid entries
// toast, Escape restores, the dropdown's contents), the Name Manager's rows
// (visible names only) and its editor (live validation, "Define name",
// "Edit name" rewriting every formula that used the old name, "Delete name").
import { describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import type { DefinedName, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { NAME_PROBLEM_MESSAGES } from "../../lib/spreadsheet/names";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { singleCell } from "../../lib/spreadsheet/selection";
import { ARRAY_VALUE_TEXT, FORMULA_PROBLEM_MESSAGE, HIDDEN_SHEET_MESSAGE, NOT_A_RANGE_MESSAGE } from "./nameBox";
import { NAME_UNDO_LABELS, useNames } from "./useNames";
import type { SheetViewFields } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const NAMES: DefinedName[] = [
  { name: "Margin", formula: "Data!$B$2:$B$5" },
  { name: "TaxRate", formula: "0.2" },
  { name: "Local", scopeSheetId: "s1", formula: "Sheet1!$C$1" },
  { name: "Elsewhere", scopeSheetId: "data", formula: "Data!$A$1" },
  { name: "Secret", formula: "Data!$A$1", hidden: true },
  { name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Sheet1!$A$1:$D$10" },
  { name: "Pair", formula: "{1,2}" },
  { name: "Ghost", formula: "Hidden!$A$1" },
];

function workbook(): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  const sheet1: Sheet = { ...base, id: "s1", name: "Sheet1", rows: [{ index: 0, cells: [{ col: 0, value: 50, formula: "SUM(Margin)" }] }] };
  const data: Sheet = {
    ...base,
    id: "data",
    name: "Data",
    rows: [1, 2, 3, 4].map((index) => ({ index, cells: [{ col: 1, value: index * 10 }] })),
    tables: [
      {
        id: "t1",
        name: "Sales",
        ref: { top: 6, left: 0, bottom: 9, right: 1 },
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Item" }, { name: "Amount" }],
        style: { showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
      },
    ],
  };
  const hidden: Sheet = { ...base, id: "hidden", name: "Hidden", visibility: "hidden", rows: [] };
  model.sheets = [sheet1, data, hidden];
  model.names = NAMES.map((name) => ({ ...name }));
  return model;
}

function setup() {
  const scope = effectScope();
  const model = workbook();
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  const viewState = {
    update: (sheetId: string, patch: SheetViewFields) => void (fields.value = { ...fields.value, [sheetId]: { ...fields.value[sheetId], ...patch } }),
    syncInto: () => {},
  };
  const selection = computed(() => fields.value[activeSheet.value?.id ?? ""]?.selection ?? singleCell({ row: 0, col: 0 }));
  const select = (top: number, left: number, bottom = top, right = left) =>
    viewState.update(activeSheet.value!.id, { selection: { active: { row: top, col: left }, ranges: [{ top, left, bottom, right }], activeRangeIndex: 0 } });
  const toast = vi.fn();
  const focusCell = vi.fn();
  const revealTick = ref(0);
  const result = scope.run(() => {
    const undo = useSpreadsheetUndo({ model: wb, recalc: (_changed, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
    const names = useNames({
      book: { workbook: wb, activeIndex, activeSheet, selectSheet: (index) => void (activeIndex.value = index) },
      viewState,
      selection: { selection },
      editing: { undo, cell: { editing: computed(() => false), focusCell, revealTick } },
      toast,
    });
    return { undo, names };
  })!;
  return { ...result, model, activeIndex, selection, select, toast, focusCell, revealTick, stop: () => scope.stop() };
}

const flush = async () => {
  await nextTick();
  await nextTick();
  await nextTick();
};

describe("useNames — the Name Box", () => {
  it("navigates to a typed reference, switching sheet, revealing it and handing focus back", async () => {
    const t = setup();
    t.names.input("data!b3:c4");
    expect(t.names.draft.value).toBe("data!b3:c4");
    await t.names.submit();
    await flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value).toEqual({ active: { row: 2, col: 1 }, ranges: [{ top: 2, left: 1, bottom: 3, right: 2 }], activeRangeIndex: 0 });
    expect(t.revealTick.value).toBe(1);
    expect(t.focusCell).toHaveBeenCalled();
    expect(t.names.draft.value).toBeNull();
    expect(t.toast).not.toHaveBeenCalled();
    t.stop();
  });

  it("selects a visible name's range (any case) and a table's data body", async () => {
    const t = setup();
    t.names.input("margin");
    await t.names.submit();
    await flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value.ranges).toEqual([{ top: 1, left: 1, bottom: 4, right: 1 }]);
    t.activeIndex.value = 0;
    t.names.input("Sales");
    await t.names.submit();
    await flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value.ranges).toEqual([{ top: 7, left: 0, bottom: 9, right: 1 }]);
    t.stop();
  });

  it("toasts for a name that is no range or is on a hidden sheet, changing nothing", async () => {
    const t = setup();
    t.names.input("TaxRate");
    await t.names.submit();
    t.names.input("Ghost");
    await t.names.submit();
    expect(t.toast.mock.calls).toEqual([[NOT_A_RANGE_MESSAGE], [HIDDEN_SHEET_MESSAGE]]);
    expect(t.activeIndex.value).toBe(0);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("defines a valid new name for the selection as one \"Define name\" step (a union for several areas)", async () => {
    const t = setup();
    t.select(1, 1, 4, 1);
    t.names.input("Profit");
    await t.names.submit();
    expect(t.model.names.at(-1)).toEqual({ name: "Profit", formula: "Sheet1!$B$2:$B$5" });
    expect(t.undo.undoLabel.value).toBe(NAME_UNDO_LABELS.define);
    t.undo.undo();
    expect(t.model.names.some((name) => name.name === "Profit")).toBe(false);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("toasts an invalid entry's NAME_PROBLEM_MESSAGES text and changes nothing", async () => {
    const t = setup();
    for (const [text, problem] of [
      ["Nowhere!", "invalidCharacters"],
      ["R2C3", "looksLikeReference"],
      ["_xlnm.Foo", "reserved"],
      ["Sales", "tableConflict"],
    ] as const) {
      t.names.input(text);
      await t.names.submit();
      if (text !== "Sales") {
        expect(t.toast).toHaveBeenLastCalledWith(NAME_PROBLEM_MESSAGES[problem]);
      }
    }
    expect(t.toast).toHaveBeenCalledTimes(3); // Sales is a table: it selects
    expect(t.model.names).toHaveLength(NAMES.length);
    t.stop();
  });

  it("Escape restores the box and returns to the grid", () => {
    const t = setup();
    t.names.input("Typed");
    t.names.cancel(true);
    expect(t.names.draft.value).toBeNull();
    expect(t.focusCell).toHaveBeenCalledTimes(1);
    expect(t.model.names).toHaveLength(NAMES.length);
    t.stop();
  });

  it("the dropdown lists the visible workbook and active-sheet names and every table, alphabetically; a pick selects", async () => {
    const t = setup();
    expect(t.names.options.value.map((option) => [option.name, option.kind])).toEqual([
      ["Ghost", "name"],
      ["Local", "name"],
      ["Margin", "name"],
      ["Pair", "name"],
      ["Sales", "table"],
      ["TaxRate", "name"],
    ]);
    await t.names.pick(t.names.options.value.find((option) => option.name === "Local")!.key);
    await flush();
    expect(t.selection.value.ranges).toEqual([{ top: 0, left: 2, bottom: 0, right: 2 }]);
    t.activeIndex.value = 1;
    expect(t.names.options.value.map((option) => option.name)).toEqual(["Elsewhere", "Ghost", "Margin", "Pair", "Sales", "TaxRate"]);
    t.stop();
  });
});

describe("useNames — the Name Manager", () => {
  it("lists every visible name with its value, refers-to and scope — not hidden or reserved ones", () => {
    const t = setup();
    t.names.openManager();
    const rows = t.names.rows.value;
    expect(rows.map((row) => row.name)).toEqual(["Elsewhere", "Ghost", "Local", "Margin", "Pair", "TaxRate"]);
    expect(rows.find((row) => row.name === "Margin")).toMatchObject({ value: ARRAY_VALUE_TEXT, refersTo: "=Data!$B$2:$B$5", scope: "Workbook" });
    expect(rows.find((row) => row.name === "TaxRate")).toMatchObject({ value: "0.2", refersTo: "=0.2" });
    expect(rows.find((row) => row.name === "Elsewhere")).toMatchObject({ value: "", scope: "Data" });
    expect(rows.find((row) => row.name === "Pair")!.value).toBe(ARRAY_VALUE_TEXT);
    expect(rows.some((row) => row.name.includes("Print_Area") || row.name === "Secret")).toBe(false);
    expect(t.names.selectedKey.value).toBe(rows[0]!.key);
    t.stop();
  });

  it("New validates live and defines one \"Define name\" step", () => {
    const t = setup();
    t.names.openManager();
    t.names.openNew();
    expect(t.names.editor.value?.fields.refersTo).toBe("=Sheet1!$A$1");
    expect(t.names.editorError.value).toBe(NAME_PROBLEM_MESSAGES.empty);
    t.names.setField("name", "margin");
    expect(t.names.editorError.value).toBe(NAME_PROBLEM_MESSAGES.duplicate);
    t.names.setField("scope", "s1");
    expect(t.names.editorError.value).toBe("");
    t.names.setField("refersTo", "=SUM(");
    expect(t.names.editorError.value).toBe(FORMULA_PROBLEM_MESSAGE);
    expect(t.names.confirmEditor()).toBe(false);
    t.names.setField("refersTo", "=sum(Data!$B$2:$B$3)*2");
    t.names.setField("comment", "Doubled");
    expect(t.names.confirmEditor()).toBe(true);
    expect(t.model.names.at(-1)).toEqual({ name: "margin", scopeSheetId: "s1", formula: "SUM(Data!$B$2:$B$3)*2", comment: "Doubled" });
    expect(t.undo.undoLabel.value).toBe(NAME_UNDO_LABELS.define);
    expect(t.names.editor.value).toBeNull();
    t.stop();
  });

  it("Edit renames as one \"Edit name\" step, rewriting every formula that used the old name; the scope stays", () => {
    const t = setup();
    t.names.openManager();
    t.names.select(t.names.rows.value.find((row) => row.name === "Margin")!.key);
    t.names.openEdit();
    t.names.setField("scope", "s1");
    expect(t.names.editor.value?.fields.scope).toBe("");
    t.names.setField("name", "Gain");
    expect(t.names.confirmEditor()).toBe(true);
    expect(t.model.sheets[0]!.rows[0]!.cells[0]!.formula).toBe("SUM(Gain)");
    expect(t.model.names.find((name) => name.name === "Gain")?.formula).toBe("Data!$B$2:$B$5");
    expect(t.undo.undoLabel.value).toBe(NAME_UNDO_LABELS.edit);
    t.undo.undo();
    expect(t.model.sheets[0]!.rows[0]!.cells[0]!.formula).toBe("SUM(Margin)");
    t.stop();
  });

  it("Delete removes the selected name as one \"Delete name\" step and selects the next row", () => {
    const t = setup();
    t.names.openManager();
    const keys = t.names.rows.value.map((row) => row.key);
    t.names.select(keys[3]!);
    t.names.deleteSelected();
    expect(t.model.names.some((name) => name.name === "Margin")).toBe(false);
    expect(t.undo.undoLabel.value).toBe(NAME_UNDO_LABELS.delete);
    expect(t.names.selectedKey.value).toBe(keys[4]);
    t.stop();
  });

  it("Refers to points: a pick replaces a lone reference, inserts after an operator, and follows each selection", async () => {
    const t = setup();
    t.names.openManager();
    t.names.openNew();
    t.names.startPick();
    t.select(2, 1, 3, 2);
    await nextTick();
    expect(t.names.editor.value?.fields.refersTo).toBe("=Sheet1!$B$3:$C$4");
    t.names.endPick();
    t.names.setField("refersTo", "=Sheet1!$A$1+");
    t.names.startPick();
    t.activeIndex.value = 1;
    await nextTick();
    expect(t.names.editor.value?.fields.refersTo).toBe("=Sheet1!$A$1+Data!$A$1");
    t.names.endPick();
    expect(t.names.picking.value).toBe(false);
    t.stop();
  });
});
