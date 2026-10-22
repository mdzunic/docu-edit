// SPEC-134's useSheetGoTo over a workbook built here, with an in-memory view
// state: Control+G and F5 (preventDefault; another dialog keeps them; the
// cell editor consumes them), the list (the Name Box's names and tables,
// then the last ten origins), a jump (the target selected, the sheet
// switched, the origin remembered — so F5 then Enter returns), errors that
// leave the selection alone, Go To Special's defaults, filters and refusal,
// and the Name Box resolving through the same grammar and remembering its
// jumps (useNames' setJump).
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef } from "vue";
import { GO_TO_HIDDEN_SHEET_MESSAGE, GO_TO_INVALID_MESSAGE, GO_TO_NOT_A_RANGE_MESSAGE, GO_TO_SPECIAL_NONE_MESSAGE } from "../../lib/spreadsheet/find";
import type { Cell, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { NAME_PROBLEM_MESSAGES } from "../../lib/spreadsheet/names";
import { singleCell, type SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { GO_TO_RECENT_LIMIT, useSheetGoTo } from "./useSheetGoTo";
import { useNames } from "./useNames";
import type { SheetViewFields } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

function rowsOf(cells: Record<string, string | number | Partial<Cell>>): Row[] {
  const rows = new Map<number, Cell[]>();
  for (const [address, spec] of Object.entries(cells)) {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!;
    const col = letters!.charCodeAt(0) - 65;
    rows.set(Number(digits) - 1, [...(rows.get(Number(digits) - 1) ?? []), typeof spec === "object" ? { col, value: null, ...spec } : { col, value: spec }]);
  }
  return [...rows].toSorted(([a], [b]) => a - b).map(([index, list]) => ({ index, cells: list.toSorted((x, y) => x.col - y.col) }));
}

function workbook(): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  const sheet = (id: string, name: string, rows: Row[], extra: Partial<Sheet> = {}): Sheet => ({ ...base, id, name, rows, ...extra });
  model.sheets = [
    sheet("sales", "Sales", rowsOf({ A1: "Region", A2: "North", B2: 1200, C2: { formula: "B2*2", value: 2400 }, B10: { formula: "SUM(B2:B6)", value: 1200 } })),
    sheet("other", "Other", rowsOf({ A1: 1 }), {
      tables: [
        {
          id: "t1",
          name: "Orders",
          ref: { top: 4, left: 0, bottom: 7, right: 1 },
          headerRow: true,
          totalsRow: false,
          columns: [{ name: "Item" }, { name: "Qty" }],
          style: { showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
        },
      ],
    }),
    sheet("archive", "Archive", [], { visibility: "hidden" }),
  ];
  model.names = [
    { name: "Totals", formula: "Sales!$B$10" },
    { name: "TaxRate", formula: "0.2" },
    { name: "Local", scopeSheetId: "other", formula: "Other!$C$1" },
    { name: "Secret", formula: "Sales!$A$1", hidden: true },
    { name: "_xlnm.Print_Area", scopeSheetId: "sales", formula: "Sales!$A$1:$C$10" },
    { name: "Ghost", formula: "Archive!$A$1" },
  ];
  return model;
}

/** The jump's ticks. */
async function flush(): Promise<void> {
  for (let i = 0; i < 3; i += 1) {
    await nextTick();
  }
}

/** A keydown at the window, as the browser dispatches it. */
function key(init: KeyboardEventInit): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  window.dispatchEvent(event);
  return event;
}

let stops: (() => void)[] = [];
afterEach(() => {
  stops.forEach((stop) => stop());
  stops = [];
  document.body.innerHTML = "";
});

function setup(options: { isMac?: boolean } = {}) {
  const model = workbook();
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  const viewState = {
    update: (sheetId: string, patch: SheetViewFields) => void (fields.value = { ...fields.value, [sheetId]: { ...fields.value[sheetId], ...patch } }),
    syncInto: () => {},
  };
  const selection = computed((): SheetSelection => fields.value[activeSheet.value?.id ?? ""]?.selection ?? singleCell({ row: 0, col: 0 }));
  const editing = ref(false);
  const focusCell = vi.fn();
  const revealTick = ref(0);
  const toast = vi.fn();
  const book = { workbook: wb, activeIndex, activeSheet, selectSheet: (index: number) => void (activeIndex.value = index) };
  const cell = { editing: computed(() => editing.value), focusCell, revealTick };
  const scope = effectScope();
  const parts = scope.run(() => {
    const undo = useSpreadsheetUndo({ model: wb, recalc: (_changed, done) => done([]), onApplied: () => {}, activateSheet: () => {} });
    const names = useNames({ book, viewState, selection: { selection }, editing: { undo, cell }, toast });
    const goTo = useSheetGoTo({ book, viewState, selection: { selection }, editing: { cell }, names, isMac: options.isMac ?? false });
    return { names, goTo };
  })!;
  stops.push(() => scope.stop());
  const select = (top: number, left: number, bottom = top, right = left) =>
    viewState.update(activeSheet.value!.id, { selection: { active: { row: top, col: left }, ranges: [{ top, left, bottom, right }], activeRangeIndex: 0 } });
  return { ...parts, model, activeIndex, selection, editing, focusCell, revealTick, toast, flush, key, select };
}

describe("useSheetGoTo — the keys", () => {
  it("opens on Control+G and F5 with preventDefault, on every platform; Cmd+G stays macOS's find-next", () => {
    for (const isMac of [false, true]) {
      const t = setup({ isMac });
      const control = t.key({ key: "g", code: "KeyG", ctrlKey: true });
      expect(control.defaultPrevented).toBe(true);
      expect(t.goTo.dialogOpen.value).toBe(true);
      t.goTo.cancel();
      expect(t.key({ key: "F5", code: "F5" }).defaultPrevented).toBe(true);
      expect(t.goTo.dialogOpen.value).toBe(true);
      t.goTo.cancel();
      stops.pop()!();
    }
    const mac = setup({ isMac: true });
    expect(mac.key({ key: "g", code: "KeyG", metaKey: true }).defaultPrevented).toBe(false);
    expect(mac.goTo.dialogOpen.value).toBe(false);
  });

  it("leaves F5 alone while another dialog is open, and consumes it while the cell editor is", () => {
    const t = setup();
    const dialog = document.createElement("div");
    dialog.setAttribute("role", "dialog");
    document.body.append(dialog);
    expect(t.key({ key: "F5" }).defaultPrevented).toBe(false);
    dialog.remove();
    t.editing.value = true;
    expect(t.key({ key: "F5" }).defaultPrevented).toBe(true);
    expect(t.key({ key: "g", ctrlKey: true }).defaultPrevented).toBe(true);
    expect(t.goTo.dialogOpen.value).toBe(false);
  });
});

describe("useSheetGoTo — the dialog", () => {
  it("lists the Name Box's names and tables — never hidden or reserved ones — then the recent origins", async () => {
    const t = setup();
    t.goTo.openDialog();
    expect(t.goTo.items.value.map((item) => [item.text, item.kind])).toEqual([
      ["Ghost", "name"],
      ["Orders", "table"],
      ["TaxRate", "name"],
      ["Totals", "name"],
    ]);
    t.goTo.setReference("B3");
    await t.goTo.confirm();
    t.goTo.openDialog();
    expect(t.goTo.items.value.at(-1)).toMatchObject({ kind: "recent", text: "Sales!A1" });
  });

  it("selects a reference's ranges, the active cell at the first's top-left, switching sheet and revealing it", async () => {
    const t = setup();
    t.goTo.openDialog();
    // Every part on one sheet: an unqualified part is the active sheet's.
    t.goTo.setReference("other!b3:c4,E9");
    expect(await t.goTo.confirm()).toBe(false);
    t.goTo.setReference("other!b3:c4,Other!E9");
    expect(await t.goTo.confirm()).toBe(true);
    await t.flush();
    expect(t.goTo.dialogOpen.value).toBe(false);
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value).toEqual({ active: { row: 2, col: 1 }, ranges: [{ top: 2, left: 1, bottom: 3, right: 2 }, { top: 8, left: 4, bottom: 8, right: 4 }], activeRangeIndex: 0 });
    expect(t.revealTick.value).toBe(1);
    expect(t.focusCell).toHaveBeenCalled();
  });

  it("shows an error and leaves the selection unchanged", async () => {
    const t = setup();
    t.select(4, 4);
    t.goTo.openDialog();
    for (const [text, message] of [
      ["Nowhere!A1", GO_TO_INVALID_MESSAGE],
      ["R1C1", GO_TO_INVALID_MESSAGE],
      ["", GO_TO_INVALID_MESSAGE],
      ["TaxRate", GO_TO_NOT_A_RANGE_MESSAGE],
      ["Archive!A1", GO_TO_HIDDEN_SHEET_MESSAGE],
    ] as const) {
      t.goTo.setReference(text);
      expect(await t.goTo.confirm()).toBe(false);
      expect(t.goTo.error.value).toBe(message);
    }
    expect(t.goTo.dialogOpen.value).toBe(true);
    expect(t.selection.value.active).toEqual({ row: 4, col: 4 });
    t.goTo.setReference("A1");
    expect(t.goTo.error.value).toBe("");
  });

  it("remembers the origin of every jump, most recent first, ten at most — so F5 then Enter returns", async () => {
    const t = setup();
    t.select(1, 1, 2, 2);
    t.goTo.openDialog();
    t.goTo.setReference("Totals");
    await t.goTo.confirm();
    await t.flush();
    expect(t.selection.value.active).toEqual({ row: 9, col: 1 });
    expect(t.goTo.recent.value).toEqual(["Sales!B2:C3"]);
    // F5 then Enter: the field holds the last origin.
    t.goTo.openDialog();
    expect(t.goTo.reference.value).toBe("Sales!B2:C3");
    await t.goTo.confirm();
    await t.flush();
    expect(t.selection.value.ranges).toEqual([{ top: 1, left: 1, bottom: 2, right: 2 }]);
    expect(t.goTo.recent.value).toEqual(["Sales!B10", "Sales!B2:C3"]);
    for (let row = 0; row < 12; row += 1) {
      t.goTo.setReference(`A${row + 20}`);
      await t.goTo.confirm();
      await t.flush();
    }
    expect(t.goTo.recent.value).toHaveLength(GO_TO_RECENT_LIMIT);
    expect(t.goTo.recent.value[0]).toBe("Sales!A30");
  });

  it("fills the field from a chosen item and goes on an activated one", async () => {
    const t = setup();
    t.goTo.openDialog();
    const orders = t.goTo.items.value.find((item) => item.text === "Orders")!;
    t.goTo.choose(orders.key);
    expect([t.goTo.reference.value, t.goTo.chosen.value]).toEqual(["Orders", orders.key]);
    expect(t.goTo.dialogOpen.value).toBe(true);
    expect(await t.goTo.activate(orders.key)).toBe(true);
    await t.flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value.ranges).toEqual([{ top: 5, left: 0, bottom: 7, right: 1 }]);
  });
});

describe("useSheetGoTo — Go To Special", () => {
  it("opens from Special… with every type checked, the filters enabled for Constants and Formulas only", () => {
    const t = setup();
    t.goTo.openDialog();
    t.goTo.openSpecial();
    expect([t.goTo.dialogOpen.value, t.goTo.specialOpen.value]).toEqual([false, true]);
    expect(t.goTo.specialTypes.value).toEqual({ numbers: true, text: true, logicals: true, errors: true });
    const enabled = (["blanks", "constants", "formulas", "errors", "lastCell", "visible"] as const).map((kind) => (t.goTo.setKind(kind), t.goTo.typesEnabled.value));
    expect(enabled).toEqual([false, true, true, false, false, false]);
  });

  it("selects the matching cells as rectangles, the active cell at the first's top-left", async () => {
    const t = setup();
    t.goTo.openSpecial();
    t.goTo.setKind("formulas");
    expect(await t.goTo.confirmSpecial()).toBe(true);
    await t.flush();
    expect(t.goTo.specialOpen.value).toBe(false);
    expect(t.selection.value).toEqual({ active: { row: 1, col: 2 }, ranges: [{ top: 1, left: 2, bottom: 1, right: 2 }, { top: 9, left: 1, bottom: 9, right: 1 }], activeRangeIndex: 0 });
  });

  it("says \"No cells were found.\" and changes nothing when nothing matches", async () => {
    const t = setup();
    t.select(1, 1);
    t.goTo.openSpecial();
    t.goTo.setKind("constants");
    for (const type of ["numbers", "text", "logicals", "errors"] as const) {
      t.goTo.setType(type, false);
    }
    expect(await t.goTo.confirmSpecial()).toBe(false);
    expect(t.goTo.specialMessage.value).toBe(GO_TO_SPECIAL_NONE_MESSAGE);
    expect(t.goTo.specialOpen.value).toBe(true);
    expect(t.selection.value.active).toEqual({ row: 1, col: 1 });
  });
});

describe("useSheetGoTo — the Name Box shares the grammar and the history", () => {
  it("accepts exactly Go To's references and names, remembering where each jump came from", async () => {
    const t = setup();
    t.names.input("A1,C3:D4");
    await t.names.submit();
    await t.flush();
    expect(t.selection.value.ranges).toEqual([{ top: 0, left: 0, bottom: 0, right: 0 }, { top: 2, left: 2, bottom: 3, right: 3 }]);
    t.names.input("totals");
    await t.names.submit();
    await t.flush();
    expect(t.selection.value.active).toEqual({ row: 9, col: 1 });
    expect(t.goTo.recent.value).toEqual(["Sales!A1,Sales!C3:D4", "Sales!A1"]);
  });

  it("keeps SPEC-133's order: a name that is no range or on a hidden sheet toasts; an invalid new name toasts its problem", async () => {
    const t = setup();
    for (const text of ["TaxRate", "Ghost", "Nowhere!", "Secret"]) {
      t.names.input(text);
      await t.names.submit();
    }
    expect(t.toast.mock.calls).toEqual([[GO_TO_NOT_A_RANGE_MESSAGE], [GO_TO_HIDDEN_SHEET_MESSAGE], [NAME_PROBLEM_MESSAGES.invalidCharacters], [NAME_PROBLEM_MESSAGES.duplicate]]);
    expect(t.goTo.recent.value).toEqual([]);
    // A valid new name is defined, as before.
    t.names.input("Profit");
    await t.names.submit();
    expect(t.model.names.at(-1)).toEqual({ name: "Profit", formula: "Sales!$A$1" });
  });
});
