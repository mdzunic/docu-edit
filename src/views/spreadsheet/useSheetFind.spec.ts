// SPEC-134's useSheetFind over a workbook built here, with the real undo
// stack (its recalculation a synchronous lib pass), an in-memory view state,
// a manual slot scheduler standing in for scheduleEagerWork and fake timers
// for the debounce: the keys (open, replace row, Escape, the editor and
// dialogs, macOS Cmd+G), the debounced chunked walk and its "Searching…"
// count, as-you-type, next / previous with the wrap hint, scope (a
// selection's ranges kept, workbook scope switching sheets), the hidden-match
// hint, an invalid expression, the marks, Find all, Replace's refusal and
// move-on, and Replace all as ONE undo step.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import { setCellEntry } from "../../lib/spreadsheet/cells";
import { FIND_ALL_LIMIT, FIND_DEBOUNCE_MS, FIND_HIDDEN_MATCH_HINT, FIND_SEARCHING_HINT, FIND_SLICE_BUDGET_MS, findAllCapMessage } from "../../lib/spreadsheet/find";
import { recalculate } from "../../lib/spreadsheet/formula/recalc";
import type { Cell, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { singleCell, type SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { findKeyAction } from "./sheetFindKeys";
import { findResultRows } from "./sheetFindRows";
import { findMarkKey } from "./sheetLayerProps";
import { FIND_UNDO_LABELS, useSheetFind } from "./useSheetFind";
import type { SheetViewFields } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const EN = SPREADSHEET_LOCALES["en-US"];

function a1(address: string): { row: number; col: number } {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(address)!;
  return { row: Number(digits) - 1, col: letters!.charCodeAt(0) - 65 };
}

function rowsOf(cells: Record<string, string | number | Partial<Cell>>): Row[] {
  const rows = new Map<number, Cell[]>();
  for (const [address, spec] of Object.entries(cells)) {
    const { row, col } = a1(address);
    rows.set(row, [...(rows.get(row) ?? []), typeof spec === "object" ? { col, value: null, ...spec } : { col, value: spec }]);
  }
  return [...rows].toSorted(([a], [b]) => a - b).map(([index, list]) => ({ index, cells: list.toSorted((x, y) => x.col - y.col) }));
}

const ifNorth = (n: number, value: number): Partial<Cell> => ({ formula: `IF(A${n}="North",1,0)`, value });

/** The fixture's shape: Sales (active), the protected Locked, the hidden Archive. */
function workbook(): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  const sheet = (id: string, name: string, rows: Row[], extra: Partial<Sheet> = {}): Sheet => ({ ...base, id, name, rows, ...extra });
  model.sheets = [
    sheet("sales", "Sales", rowsOf({ A1: "Region", A2: "North", C2: ifNorth(2, 1), A3: "South", C3: ifNorth(3, 0), A4: "Northwest", C4: ifNorth(4, 0), B5: 300, C5: { formula: "B5/0", value: { error: "#DIV/0!" } }, A6: "north", C6: ifNorth(6, 0) })),
    sheet("locked", "Locked", rowsOf({ A1: "North office" }), { protection: { sheet: true } }),
    sheet("archive", "Archive", rowsOf({ A1: "North" }), { visibility: "hidden" }),
  ];
  for (const each of model.sheets) {
    delete each.dirty;
  }
  return model;
}

/** A keydown dispatched at `target` (default the window), as the browser would. */
function dispatchKey(init: KeyboardEventInit, target: EventTarget = window): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
}

const press = (init: KeyboardEventInit): KeyboardEvent => new KeyboardEvent("keydown", init);

let scopes: (() => void)[] = [];
beforeEach(() => {
  vi.useFakeTimers();
});
afterEach(() => {
  scopes.forEach((stop) => stop());
  scopes = [];
  vi.useRealTimers();
  document.body.innerHTML = "";
});

function setup(model = workbook(), options: { isMac?: boolean; clockStep?: number } = {}) {
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  const viewState = {
    update: (sheetId: string, patch: SheetViewFields) => void (fields.value = { ...fields.value, [sheetId]: { ...fields.value[sheetId], ...patch } }),
    syncInto: () => {},
  };
  const selection = computed((): SheetSelection => fields.value[activeSheet.value?.id ?? ""]?.selection ?? singleCell({ row: 0, col: 0 }));
  const select = (at: string, ...ranges: GridRange[]) => viewState.update(activeSheet.value!.id, { selection: { active: a1(at), ranges: ranges.length > 0 ? ranges : [{ top: a1(at).row, left: a1(at).col, bottom: a1(at).row, right: a1(at).col }], activeRangeIndex: 0 } });
  const editing = ref(false);
  const focusCell = vi.fn();
  const recalcs: number[] = [];
  const slots: (() => void)[] = [];
  const budgets: number[] = [];
  let clock = 0;
  const scope = effectScope();
  const result = scope.run(() => {
    const undo = useSpreadsheetUndo({
      model: wb,
      recalc: (changed, done) => {
        recalcs.push(changed.length);
        done(recalculate(model, changed).changedSheetIds);
      },
      onApplied: () => triggerRef(wb),
      activateSheet: (id) => void (activeIndex.value = model.sheets.findIndex((sheet) => sheet.id === id)),
    });
    const spied = { record: vi.fn(undo.record), commit: vi.fn(undo.commit) };
    const find = useSheetFind({
      book: { workbook: wb, activeIndex, activeSheet, selectSheet: (index) => void (activeIndex.value = index) },
      viewState,
      selection: { selection, formatContext: computed(() => workbookFormatContext(model, EN)) },
      editing: { undo: spied, cell: { editing: computed(() => editing.value), focusCell, revealTick: ref(0) } },
      locale: () => EN,
      schedule: (callback, budget) => {
        budgets.push(budget);
        let live = true;
        slots.push(() => live && callback(budget));
        return { cancel: () => void (live = false) };
      },
      now: () => (clock += options.clockStep ?? 0),
      isMac: options.isMac ?? false,
      measurer: { widthPx: (text) => text.length * 7, lineHeightPx: () => 20 },
    });
    return { undo, spied, find };
  })!;
  scopes.push(() => scope.stop());
  /** Runs every queued slot (and the ones they queue). */
  const drain = () => {
    while (slots.length > 0) {
      slots.shift()!();
    }
  };
  /** The debounce, every slot, and the jump's ticks. */
  const settle = async () => {
    await nextTick(); // the watchers arm the debounce
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS);
    drain();
    for (let i = 0; i < 4; i += 1) {
      await nextTick();
    }
  };
  const active = () => {
    const { row, col } = selection.value.active;
    return `${activeSheet.value?.name}!${String.fromCharCode(65 + col)}${row + 1}`;
  };
  return { ...result, model, wb, activeIndex, selection, select, editing, focusCell, recalcs, slots, budgets, drain, settle, key: dispatchKey, active };
}

async function searchFor(t: ReturnType<typeof setup>, text: string): Promise<void> {
  t.find.openFind();
  t.find.barEvents["update:query"](text);
  await t.settle();
}

describe("useSheetFind — the keys", () => {
  it("opens on Cmd/Ctrl+F and with the replace row on Cmd/Ctrl+H, preventing the browser's find", () => {
    const t = setup();
    const event = t.key({ key: "f", ctrlKey: true });
    expect(event.defaultPrevented).toBe(true);
    expect(t.find.bar.value.visible).toBe(true);
    expect(t.find.bar.value.expanded).toBe(false);
    expect(t.key({ key: "h", ctrlKey: true }).defaultPrevented).toBe(true);
    expect(t.find.bar.value.expanded).toBe(true);
  });

  it("closes on Escape from the bar's buttons, clearing the marks and returning focus to the grid", async () => {
    const t = setup();
    await searchFor(t, "north");
    expect(t.find.marks.value?.matches.size).toBe(3);
    const bar = document.createElement("div");
    bar.dataset.testid = "find-bar";
    const button = document.createElement("button");
    bar.append(button);
    document.body.append(bar);
    expect(t.key({ key: "Escape" }, button).defaultPrevented).toBe(true);
    expect(t.find.bar.value.visible).toBe(false);
    expect(t.find.marks.value).toBeNull();
    expect(t.focusCell).toHaveBeenCalled();
  });

  it("consumes the keys and does nothing while the cell editor is open; leaves them to a dialog", () => {
    const t = setup();
    t.editing.value = true;
    expect(t.key({ key: "f", ctrlKey: true }).defaultPrevented).toBe(true);
    expect(t.find.bar.value.visible).toBe(false);
    t.editing.value = false;
    const dialog = document.createElement("div");
    dialog.setAttribute("role", "dialog");
    const input = document.createElement("input");
    dialog.append(input);
    document.body.append(dialog);
    expect(t.key({ key: "f", ctrlKey: true }, input).defaultPrevented).toBe(false);
    expect(t.find.bar.value.visible).toBe(false);
  });

  it("steps with Cmd+G / Cmd+Shift+G on macOS while the bar is open", async () => {
    const t = setup(workbook(), { isMac: true });
    expect(t.key({ key: "g", metaKey: true }).defaultPrevented).toBe(false);
    expect(t.key({ key: "f", metaKey: true }).defaultPrevented).toBe(true);
    t.find.barEvents["update:query"]("north");
    await t.settle();
    expect(t.active()).toBe("Sales!A2");
    expect(t.key({ key: "g", metaKey: true }).defaultPrevented).toBe(true);
    await t.settle();
    expect(t.active()).toBe("Sales!A4");
    t.key({ key: "G", metaKey: true, shiftKey: true });
    await t.settle();
    expect(t.active()).toBe("Sales!A2");
  });

  it("classifies the keys through findKeyAction", () => {
    expect(findKeyAction(press({ key: "f", ctrlKey: true }), false, false)).toBe("open");
    expect(findKeyAction(press({ key: "h", metaKey: true }), false, true)).toBe("openReplace");
    expect(findKeyAction(press({ key: "g", metaKey: true }), true, true)).toBe("next");
    expect(findKeyAction(press({ key: "g", metaKey: true, shiftKey: true }), true, true)).toBe("prev");
    expect(findKeyAction(press({ key: "g", ctrlKey: true }), true, false)).toBeNull();
    expect(findKeyAction(press({ key: "Escape" }), true, false)).toBeNull();
  });
});

describe("useSheetFind — the walk and navigation", () => {
  it("debounces the query, then walks in scheduleEagerWork slots of FIND_SLICE_BUDGET_MS and selects the first hit at or after the active cell", async () => {
    const t = setup();
    t.find.openFind();
    t.drain();
    t.find.barEvents["update:query"]("north");
    await nextTick();
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS - 1);
    expect(t.slots).toHaveLength(0);
    vi.advanceTimersByTime(1);
    expect(t.slots).toHaveLength(1);
    t.drain();
    await t.settle();
    expect(t.budgets.every((budget) => budget === FIND_SLICE_BUDGET_MS)).toBe(true);
    expect(t.find.bar.value.matchCount).toBe(3);
    expect(t.find.bar.value.currentIndex).toBe(0);
    expect(t.active()).toBe("Sales!A2");
  });

  it("shows the running total with \"Searching…\" until the walk is done", async () => {
    const rows: Row[] = Array.from({ length: 3000 }, (_, index) => ({ index, cells: [{ col: 0, value: index % 3 === 0 ? "north" : "south" }] }));
    const model = workbook();
    model.sheets[0] = { ...model.sheets[0]!, rows };
    const t = setup(model, { clockStep: 1 });
    t.find.openFind();
    t.drain();
    t.find.barEvents["update:query"]("north");
    await nextTick();
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS);
    t.slots.shift()!();
    expect(t.find.bar.value.hint).toBe(FIND_SEARCHING_HINT);
    expect(t.find.bar.value.matchCount).toBeGreaterThan(0);
    expect(t.find.bar.value.matchCount).toBeLessThan(1000);
    await t.settle();
    expect(t.find.bar.value.matchCount).toBe(1000);
    expect(t.find.bar.value.hint).toBe("");
  });

  it("steps next and previous from the active cell, wrapping with \"Wrapped to top\"", async () => {
    const t = setup();
    await searchFor(t, "north");
    t.find.barEvents.next();
    await t.settle();
    expect([t.active(), t.find.bar.value.currentIndex]).toEqual(["Sales!A4", 1]);
    // Two presses in quick succession: each steps on from the last.
    t.find.barEvents.next();
    t.find.barEvents.next();
    await t.settle();
    expect(t.active()).toBe("Sales!A2");
    expect(t.find.bar.value.hint).toBe("Wrapped to top");
    t.find.barEvents.prev();
    await t.settle();
    expect(t.active()).toBe("Sales!A6");
    expect(t.find.bar.value.hint).toBe("Wrapped to top");
    // From wherever the user has since clicked.
    t.select("A3");
    t.find.barEvents.next();
    await t.settle();
    expect(t.active()).toBe("Sales!A4");
  });

  it("searches only a multi-cell selection, keeping its ranges while the active cell moves", async () => {
    const t = setup();
    const ranges = [{ top: 3, left: 0, bottom: 5, right: 0 }];
    t.select("A4", ...ranges);
    await searchFor(t, "north");
    expect(t.find.bar.value.matchCount).toBe(2);
    t.find.barEvents.next();
    await t.settle();
    expect(t.selection.value).toEqual({ active: a1("A6"), ranges, activeRangeIndex: 0 });
    expect(t.find.bar.value.matchCount).toBe(2);
  });

  it("switches the active sheet for a workbook-scope hit, and never searches the hidden sheet", async () => {
    const t = setup();
    t.find.within.value = "workbook";
    expect(t.find.bar.value.placeholder).toBe("Find in workbook");
    await searchFor(t, "office");
    expect(t.find.bar.value.matchCount).toBe(1);
    expect(t.activeIndex.value).toBe(1);
    expect(t.active()).toBe("Locked!A1");
    t.find.within.value = "sheet";
    expect(t.find.bar.value.placeholder).toBe("Find in sheet");
  });

  it("hints a current match in a hidden row or column", async () => {
    const model = workbook();
    model.sheets[0] = { ...model.sheets[0]!, rows: model.sheets[0]!.rows.map((row) => (row.index === 3 ? { ...row, hidden: true as const } : row)) };
    const t = setup(model);
    await searchFor(t, "northwest");
    expect(t.active()).toBe("Sales!A4");
    expect(t.find.bar.value.hint).toBe(FIND_HIDDEN_MATCH_HINT);
  });

  it("shows an invalid expression's error and runs nothing", async () => {
    const t = setup();
    t.find.regex.value = true;
    await searchFor(t, "(north");
    expect(t.find.bar.value.error).toBe("Invalid regular expression");
    expect(t.find.bar.value.matchCount).toBe(0);
    expect(t.slots).toHaveLength(0);
  });

  it("keeps every option for the view session", async () => {
    const t = setup();
    t.find.openFind();
    t.find.barEvents["update:matchCase"](true);
    t.find.entireCell.value = true;
    t.find.lookIn.value = "formulas";
    t.find.within.value = "workbook";
    t.find.barEvents["update:useRegex"](true);
    t.find.close();
    t.find.openFind();
    expect([t.find.bar.value.matchCase, t.find.bar.value.useRegex, t.find.entireCell.value, t.find.lookIn.value, t.find.within.value]).toEqual([true, true, true, "formulas", "workbook"]);
    expect(t.find.bar.value.showWholeWords).toBe(false);
    expect([t.find.bar.value.scopeLabel, t.find.bar.value.scopeTitle]).toEqual([
      "Cells only",
      "Search covers cell values and formulas. Notes, comments, shapes, charts and hidden sheets are not searched.",
    ]);
  });

  it("restarts the walk after any model change", async () => {
    const t = setup();
    await searchFor(t, "north");
    expect(t.find.bar.value.matchCount).toBe(3);
    setCellEntry(t.model, 0, a1("B8"), { value: "North pole" });
    triggerRef(t.wb);
    await t.settle();
    expect(t.find.bar.value.matchCount).toBe(4);
  });

  it("marks the active sheet's matches and the current one, and nothing once closed", async () => {
    const t = setup();
    await searchFor(t, "north");
    expect(t.find.marks.value).toEqual({ sheetId: "sales", matches: new Set([findMarkKey(1, 0), findMarkKey(3, 0), findMarkKey(5, 0)]), current: findMarkKey(1, 0) });
    t.find.close();
    expect(t.find.marks.value).toBeNull();
  });
});

describe("useSheetFind — Find all", () => {
  it("lists every match in search order; activating a row selects it as find-next does", async () => {
    const t = setup();
    t.find.lookIn.value = "formulas";
    await searchFor(t, "north");
    t.find.findAll();
    const rows = t.find.results.value!.rows;
    expect(rows.map((row) => row.ref)).toEqual(["A2", "C2", "C3", "A4", "C4", "A6", "C6"]);
    expect(rows[1]).toMatchObject({ sheetName: "Sales", value: "1", formula: '=IF(A2="North",1,0)' });
    t.find.activate(3);
    await t.settle();
    expect([t.active(), t.find.bar.value.currentIndex]).toEqual(["Sales!A4", 3]);
    expect(t.find.results.value!.cap).toBe("");
  });

  it("caps the list at FIND_ALL_LIMIT with the count of every match", () => {
    const model = workbook();
    const hits = Array.from({ length: FIND_ALL_LIMIT + 7 }, () => ({ sheetId: "sales", row: 1, col: 0 }));
    const results = findResultRows(model, hits, workbookFormatContext(model, EN));
    expect(results.rows).toHaveLength(FIND_ALL_LIMIT);
    expect(results.cap).toBe(findAllCapMessage(FIND_ALL_LIMIT + 7));
    expect(results.cap).toBe("Showing the first 5,000 of 5,007 matches");
  });
});

describe("useSheetFind — replace", () => {
  it("Replace: rewrites the current cell as ONE \"Replace\" step and moves to the next match", async () => {
    const t = setup();
    await searchFor(t, "north");
    t.find.barEvents["update:replaceText"]("N.");
    t.find.barEvents.replace();
    expect(t.model.sheets[0]!.rows[1]!.cells[0]!.value).toBe("N.");
    expect(t.undo.undoLabel.value).toBe(FIND_UNDO_LABELS.replace);
    await t.settle();
    expect(t.active()).toBe("Sales!A4");
  });

  it("refuses a replacement that breaks a formula: the cell unchanged, the address named, the match kept", async () => {
    const t = setup();
    t.find.lookIn.value = "formulas";
    await searchFor(t, 'IF(A4="North"');
    expect(t.active()).toBe("Sales!C4");
    t.find.barEvents["update:replaceText"]("IF((");
    t.find.barEvents.replace();
    expect(t.find.bar.value.hint).toBe("That replacement would break the formula in C4");
    expect(t.model.sheets[0]!.rows[3]!.cells[1]!.formula).toBe('IF(A4="North",1,0)');
    expect([t.active(), t.find.bar.value.currentIndex]).toEqual(["Sales!C4", 0]);
    expect(t.undo.canUndo.value).toBe(false);
  });

  it("Replace all: ONE step — one record before, one commit after — dirtying only touched sheets, one recalculation, reported", async () => {
    const t = setup();
    t.find.lookIn.value = "formulas";
    t.find.within.value = "workbook";
    await searchFor(t, "north");
    const lockedRows = t.model.sheets[1]!.rows;
    t.find.barEvents["update:replaceText"]("N.");
    t.spied.record.mockClear();
    t.spied.commit.mockClear();
    t.recalcs.length = 0;
    t.find.barEvents["replace-all"]();
    expect(t.spied.record.mock.calls).toEqual([[FIND_UNDO_LABELS.replaceAll, "workbook"]]);
    expect(t.spied.commit).toHaveBeenCalledTimes(1);
    expect(t.recalcs).toHaveLength(1);
    expect(t.find.bar.value.hint).toBe("Replaced 7 occurrences in 7 cells · skipped 1 protected");
    expect(t.model.sheets[0]!.dirty).toBe(true);
    expect(t.model.sheets[1]!.dirty).toBeUndefined();
    expect(t.model.sheets[1]!.rows).toBe(lockedRows);
    expect(t.model.sheets[0]!.rows[1]!.cells.map((cell) => [cell.value, cell.formula])).toEqual([
      ["N.", undefined],
      [1, 'IF(A2="N.",1,0)'],
    ]);
    // One undo restores every value and formula.
    t.undo.undo();
    expect(t.model.sheets[0]!.rows[1]!.cells.map((cell) => [cell.value, cell.formula])).toEqual([
      ["North", undefined],
      [1, 'IF(A2="North",1,0)'],
    ]);
    expect(t.undo.canUndo.value).toBe(false);
  });

  it("scopes Replace all to the active sheet for Within = Sheet, and in Values mode enters text, never a formula", async () => {
    const t = setup();
    await searchFor(t, "north");
    t.find.barEvents["update:replaceText"]("=1+1");
    t.spied.record.mockClear();
    t.find.barEvents["replace-all"]();
    expect(t.spied.record.mock.calls).toEqual([[FIND_UNDO_LABELS.replaceAll, { sheetId: "sales" }]]);
    expect(t.find.bar.value.hint).toBe("Replaced 3 occurrences in 3 cells");
    expect(t.model.sheets[0]!.rows[1]!.cells[0]).toMatchObject({ value: "=1+1" });
    expect(t.model.sheets[0]!.rows[1]!.cells[0]!.formula).toBeUndefined();
    expect(t.model.sheets[0]!.rows[3]!.cells[0]!.value).toBe("=1+1west");
  });
});
