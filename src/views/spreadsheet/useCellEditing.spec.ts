// SPEC-122 "Editor keys", "Input parsing", "Cell commands", "Protection":
// useCellEditing's mode machine (Ready → Enter/Edit, the status bar's mode
// label), the key table — commits and moves, the Tab-run return, cycling
// inside a multi-cell selection, Enter-mode arrows, Escape, Alt+Enter,
// Ctrl/Cmd+Enter's Fill, Delete's Clear contents, IME composition — the
// commit (parsing in the user's locale, formula validation and its messages,
// the length limits, one labelled undo step through SPEC-121's pipeline that
// recalculates dependents on every sheet and marks the sheet dirty, nothing
// for an unchanged entry, wrap on a line break and the row refit), the
// draft's own undo while editing, and the refusals for protection and
// arrays. Driven the way the grid drives it: the editor's keydown first, and
// a key it passes on is SPEC-117's navigation (useSheetGrid's fallback); a
// typed character lands natively and is read back through `input`.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import { getCell, parseA1 } from "../../lib/spreadsheet/address";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import { cellKey, type CellKey } from "../../lib/spreadsheet/formula/dependencies";
import { recalculate } from "../../lib/spreadsheet/formula/recalc";
import type { Cell, CellStyle, CellValue, Row, Sheet, SpreadsheetModel, TextRun } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { addSheet, createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { cellStyle } from "../../lib/spreadsheet/styles";
import type { SheetViewFields } from "./useSheetViewState";

const platform = { mac: false };
vi.mock("../../lib/core/platform", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/core/platform")>()),
  isMacPlatform: () => platform.mac,
}));

const { useCellEditing } = await import("./useCellEditing");
const { useSheetMetrics } = await import("./useSheetMetrics");
const { useSheetSelection } = await import("./useSheetSelection");
const { useSpreadsheetUndo } = await import("./useSpreadsheetUndo");
const { ARRAY_PART_MESSAGE, FORMULA_ERROR_MESSAGE, FORMULA_TOO_LONG_MESSAGE, PROTECTED_CELL_MESSAGE } = await import("./cellEntry");

const scopes: Array<() => void> = [];
afterEach(() => {
  for (const stop of scopes.splice(0)) {
    stop();
  }
  platform.mac = false;
});

type CellSpec = CellValue | Omit<Cell, "col">;

/** Rows from A1 addresses: a value, or a cell without its column. */
function rowsOf(cells: Record<string, CellSpec>): Row[] {
  const byRow = new Map<number, Cell[]>();
  for (const [address, spec] of Object.entries(cells)) {
    const at = parseA1(address)!;
    const cell: Cell = spec !== null && typeof spec === "object" && "value" in spec ? { col: at.col, ...spec } : { col: at.col, value: spec };
    byRow.set(at.row, [...(byRow.get(at.row) ?? []), cell]);
  }
  return [...byRow]
    .toSorted((a, b) => a[0] - b[0])
    .map(([index, rowCells]) => ({ index, cells: rowCells.toSorted((a, b) => a.col - b.col) }));
}

/** A clean blank workbook (no sheet dirty) whose Sheet1 holds `cells`;
 * `extraStyles` are appended after the Normal style (ids from 1). */
function book(cells: Record<string, CellSpec> = {}, options: { extraStyles?: CellStyle[]; sheet?: Partial<Sheet> } = {}): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows: rowsOf(cells), ...options.sheet }];
  delete model.sheets[0]!.dirty;
  model.styles = [...model.styles, ...(options.extraStyles ?? [])];
  return model;
}

/** A second, clean sheet named `name` holding `cells`. */
function withSheet(model: SpreadsheetModel, name: string, cells: Record<string, CellSpec> = {}): SpreadsheetModel {
  addSheet(model, model.sheets.length - 1);
  const index = model.sheets.length - 1;
  model.sheets = model.sheets.map((sheet, i) => (i === index ? { ...sheet, name, rows: rowsOf(cells) } : sheet));
  delete model.sheets[index]!.dirty;
  delete model.structureDirty;
  model.activeSheetIndex = 0;
  return model;
}

const PAGE = { rows: 20, cols: 8 };

/** useCellEditing over `model` as useWorkbookEditing composes it: the real
 * selection, metrics and undo stack (its recalculation a synchronous lib
 * pass), a per-sheet in-memory view state, a toast spy and a fixed measurer
 * (7 px per character, 20 px lines: one line is the default 15 pt row). */
function setup(model: SpreadsheetModel) {
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(model.activeSheetIndex);
  const activeSheet = computed(() => workbook.value?.sheets[activeIndex.value] ?? null);
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  const viewState = {
    active: computed(() => fields.value[activeSheet.value?.id ?? ""] ?? {}),
    updateActive(patch: SheetViewFields) {
      const id = activeSheet.value?.id ?? "";
      fields.value = { ...fields.value, [id]: { ...fields.value[id], ...patch } };
    },
  };
  const toasts: string[] = [];
  const recalls: CellKey[][] = [];
  const scope = effectScope();
  const parts = scope.run(() => {
    const sheetMetrics = useSheetMetrics(() => activeSheet.value);
    const selection = useSheetSelection({
      sheet: () => activeSheet.value,
      sheetMetrics,
      viewState,
      model: () => workbook.value,
      locale: () => SPREADSHEET_LOCALES["en-US"],
    });
    const undo = useSpreadsheetUndo({
      model: workbook,
      recalc: (changed, done) => {
        recalls.push(changed);
        done(recalculate(model, changed).changedSheetIds);
      },
      onApplied: () => triggerRef(workbook),
      activateSheet: (id) => {
        activeIndex.value = model.sheets.findIndex((sheet) => sheet.id === id);
      },
    });
    const editing = useCellEditing({
      book: {
        workbook,
        activeIndex,
        activeSheet,
        selectSheet: (index) => {
          activeIndex.value = index;
        },
      },
      selection,
      sheetMetrics,
      undo,
      locale: () => SPREADSHEET_LOCALES["en-US"],
      toast: (message) => toasts.push(message),
      measurer: fixedMeasurer(7, 20),
    });
    return { selection, undo, editing };
  })!;
  scopes.push(() => scope.stop());
  const { selection, undo, editing } = parts;

  /** A key as the grid delivers it: the editor's table first, else
   * SPEC-117's navigation. */
  function press(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
    const event = new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true, ...init });
    if (!editing.onKeydown(event) && selection.handleKey(event, PAGE)) {
      event.preventDefault();
    }
    return event;
  }

  /** Types `text` a character at a time: each keydown passes, then the
   * browser inserts the character at the caret (into the empty editor in
   * Ready) and the editor reads its host back. */
  function type(text: string): void {
    for (const char of text) {
      if (press(char).defaultPrevented) {
        continue;
      }
      if (editing.mode.value === "ready") {
        editing.input(char, char.length);
      } else {
        const at = editing.caret.value;
        const current = editing.text.value;
        editing.input(current.slice(0, at) + char + current.slice(at), at + char.length);
      }
    }
  }

  /** Selects a cell ("B3") or a range ("B2:C3", active at its first cell) as a click and Shift+click do. */
  function select(target: string): void {
    const [from, to] = target.split(":");
    const start = parseA1(from!)!;
    // The grid's pointer asks the editor first (an edit commits; a click forgets a Tab-run).
    if (editing.pressCell({ row: start.row, col: start.col }, false, false) !== "select") {
      return;
    }
    selection.press({ row: start.row, col: start.col }, { shift: false, mod: false });
    if (to) {
      const end = parseA1(to)!;
      selection.press({ row: end.row, col: end.col }, { shift: true, mod: false });
    }
    selection.release();
  }

  /** The live cell at an address on sheet `index` (default: the first). */
  function cell(address: string, index = 0): Cell | undefined {
    const at = parseA1(address)!;
    return getCell(model.sheets[index]!, at.row, at.col);
  }

  const active = (): string => selection.nameBoxText.value;
  const ranges = () => selection.selection.value.ranges;

  return { model, workbook, activeIndex, selection, undo, editing, toasts, recalls, press, type, select, cell, active, ranges };
}

describe("modes", () => {
  it("Ready: typing opens Enter mode with the typed text replacing the content, caret after it", () => {
    const { editing, type, select, cell } = setup(book({ B3: "Revenue" }));
    select("B3");
    expect(editing.mode.value).toBe("ready");
    expect(editing.modeLabel.value).toBe("Ready");
    type("R");
    expect(editing.mode.value).toBe("enter");
    expect(editing.modeLabel.value).toBe("Enter");
    expect(editing.text.value).toBe("R");
    expect(editing.caret.value).toBe(1);
    expect(editing.target.value?.at).toEqual({ row: 2, col: 1 });
    type("ent");
    expect(editing.text.value).toBe("Rent");
    // Nothing is written until the commit.
    expect(cell("B3")?.value).toBe("Revenue");
  });

  it("an IME composition starting in Ready opens Enter mode on an empty draft; its text lands when it ends", () => {
    const { editing, select, press, cell } = setup(book({ A1: "old" }));
    select("A1");
    expect(editing.compositionStart()).toBe(true);
    expect(editing.mode.value).toBe("enter");
    expect(editing.text.value).toBe("");
    editing.compositionEnd();
    editing.input("日本", 2);
    press("Enter");
    expect(cell("A1")?.value).toBe("日本");
  });

  it("F2 opens Edit mode on the cell's edit text with the caret at the end", () => {
    const styles: CellStyle[] = [{ numberFormat: "#,##0.00" }, { builtinFormatId: 14, numberFormat: "m/d/yyyy" }, { numberFormat: "0.00%" }, { quotePrefix: true }];
    const { editing, press, select } = setup(
      book(
        {
          A1: { value: 2400, formula: "A2*2" },
          A2: { value: 1234.5, styleId: 1 },
          A3: { value: 46303, styleId: 2 },
          A4: { value: 0.125, styleId: 3 },
          A5: { value: "007", styleId: 4 },
          A6: { value: 3, formula: "of:=FOO(", formulaUnparsed: true },
          A7: "Revenue",
        },
        { extraStyles: styles },
      ),
    );
    const expected: Array<[string, string]> = [
      ["A1", "=A2*2"], // "=" + the formula
      ["A2", "1234.5"], // SPEC-118's editText: no grouping
      ["A3", "10/8/2026"], // a date in the locale's short date
      ["A4", "12.5%"],
      ["A5", "'007"], // a quote-prefixed string keeps its apostrophe
      ["A6", "=of:=FOO("], // an unparsed formula as stored
      ["A7", "Revenue"],
      ["A8", ""],
    ];
    for (const [address, text] of expected) {
      select(address);
      const event = press("F2");
      expect(event.defaultPrevented).toBe(true);
      expect([address, editing.mode.value, editing.modeLabel.value, editing.text.value, editing.caret.value]).toEqual([address, "edit", "Edit", text, text.length]);
      press("Escape");
      expect(editing.mode.value).toBe("ready");
    }
  });

  it("opens Edit mode at a given character (a click in the formula bar)", () => {
    const { editing, select } = setup(book({ A1: "Revenue" }));
    select("A1");
    expect(editing.openEdit({ at: 3 })).toBe(true);
    expect([editing.mode.value, editing.text.value, editing.caret.value]).toEqual(["edit", "Revenue", 3]);
  });

  it("F2 inside the editor toggles Enter ↔ Edit and keeps the draft", () => {
    const { editing, type, press } = setup(book());
    type("abc");
    press("F2");
    expect([editing.mode.value, editing.modeLabel.value, editing.text.value]).toEqual(["edit", "Edit", "abc"]);
    press("F2");
    expect([editing.mode.value, editing.modeLabel.value, editing.text.value]).toEqual(["enter", "Enter", "abc"]);
  });

  it("the mode label reads Point while a reference is being pointed at, and Enter again once typing goes on", () => {
    const { editing, type, press, select } = setup(book());
    select("B3");
    type("=");
    expect(editing.modeLabel.value).toBe("Enter");
    press("ArrowUp");
    expect(editing.text.value).toBe("=B2");
    expect(editing.modeLabel.value).toBe("Point");
    type("+");
    expect(editing.modeLabel.value).toBe("Enter");
    press("F2");
    expect(editing.modeLabel.value).toBe("Edit");
  });

  it("Escape cancels: the cell is restored, nothing is written or recorded, and the active cell stays", () => {
    const { editing, undo, type, press, select, cell, active, model } = setup(book({ A1: "abc" }));
    const sheets = model.sheets;
    select("A1");
    type("xyz");
    expect(press("Escape").defaultPrevented).toBe(true);
    expect([editing.mode.value, editing.text.value]).toEqual(["ready", ""]);
    press("F2");
    type("!");
    press("Escape");
    expect(cell("A1")?.value).toBe("abc");
    expect(model.sheets).toBe(sheets);
    expect(undo.canUndo.value).toBe(false);
    expect(active()).toBe("A1");
  });
});

describe("Enter, Tab and the arrows", () => {
  it("Enter / Shift+Enter commit and move down / up; Tab / Shift+Tab right / left", () => {
    const { type, press, select, cell, active } = setup(book());
    select("B2");
    type("down");
    expect(press("Enter").defaultPrevented).toBe(true);
    expect([cell("B2")?.value, active()]).toEqual(["down", "B3"]);
    type("up");
    press("Enter", { shiftKey: true });
    expect([cell("B3")?.value, active()]).toEqual(["up", "B2"]);
    type("right");
    expect(press("Tab").defaultPrevented).toBe(true);
    expect([cell("B2")?.value, active()]).toEqual(["right", "C2"]);
    type("left");
    press("Tab", { shiftKey: true });
    expect([cell("C2")?.value, active()]).toEqual(["left", "B2"]);
  });

  it("Enter and Tab in Ready move the same way", () => {
    const { editing, press, select, active } = setup(book());
    select("B2");
    press("Enter");
    expect(active()).toBe("B3");
    press("Tab");
    expect(active()).toBe("C3");
    press("Tab", { shiftKey: true });
    press("Enter", { shiftKey: true });
    expect(active()).toBe("B2");
    expect(editing.mode.value).toBe("ready");
  });

  it("Enter after a run of Tabs returns to the run's starting column one row down", () => {
    const { type, press, select, cell, active } = setup(book());
    select("B2");
    type("a");
    press("Tab");
    // Typing the next entry keeps the run.
    type("b");
    press("Tab");
    type("c");
    press("Enter");
    expect([cell("B2")?.value, cell("C2")?.value, cell("D2")?.value]).toEqual(["a", "b", "c"]);
    expect(active()).toBe("B3");
    // The return is used once: the next Enter goes straight down.
    type("d");
    press("Enter");
    expect(active()).toBe("B4");
  });

  it("a navigation key or a click forgets the Tab-run's column", () => {
    const { editing, type, press, select, active } = setup(book());
    select("B2");
    type("a");
    press("Tab");
    press("ArrowDown");
    expect(active()).toBe("C3");
    type("b");
    press("Enter");
    expect(active()).toBe("C4");

    select("B6");
    type("a");
    press("Tab");
    expect(editing.pressCell({ row: 9, col: 4 }, false, false)).toBe("select");
    select("E10");
    type("b");
    press("Enter");
    expect(active()).toBe("E11");
  });

  it("inside a multi-cell selection Enter cycles down the column and on to the next, without collapsing it", () => {
    const { type, press, select, cell, active, ranges } = setup(book());
    select("B2:C3");
    const range = { top: 1, left: 1, bottom: 2, right: 2 };
    const visits: string[] = [];
    for (const text of ["1", "2", "3", "4"]) {
      visits.push(active());
      type(text);
      press("Enter");
      expect(ranges()).toEqual([range]);
    }
    expect(visits).toEqual(["B2", "B3", "C2", "C3"]);
    expect(active()).toBe("B2");
    expect([cell("B2")?.value, cell("B3")?.value, cell("C2")?.value, cell("C3")?.value]).toEqual([1, 2, 3, 4]);
    // Shift reverses, wrapping back to the last cell.
    type("5");
    press("Enter", { shiftKey: true });
    expect([cell("B2")?.value, active()]).toEqual([5, "C3"]);
    expect(ranges()).toEqual([range]);
  });

  it("inside a multi-cell selection Tab cycles along the row and on to the next; Shift+Tab reverses", () => {
    const { type, press, select, active, ranges } = setup(book());
    select("B2:C3");
    const visits: string[] = [];
    for (const text of ["1", "2", "3", "4"]) {
      visits.push(active());
      type(text);
      press("Tab");
    }
    expect(visits).toEqual(["B2", "C2", "B3", "C3"]);
    expect(active()).toBe("B2");
    type("5");
    press("Tab", { shiftKey: true });
    expect(active()).toBe("C3");
    expect(ranges()).toEqual([{ top: 1, left: 1, bottom: 2, right: 2 }]);
    // No Tab-run inside a selection: Enter cycles down the column.
    type("6");
    press("Enter");
    expect(active()).toBe("B2");
  });

  it("Enter-mode arrows commit and move; Edit-mode arrows are left to the caret", () => {
    const { editing, type, press, select, cell, active } = setup(book());
    select("B2");
    type("5");
    const arrow = press("ArrowRight");
    expect(arrow.defaultPrevented).toBe(true);
    expect([cell("B2")?.value, active(), editing.mode.value]).toEqual([5, "C2", "ready"]);

    press("F2");
    type("x");
    for (const key of ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"]) {
      const event = press(key);
      // The editor's (no grid navigation) but not prevented: the browser moves the caret.
      expect(event.defaultPrevented).toBe(false);
    }
    expect([editing.mode.value, editing.text.value, active()]).toEqual(["edit", "x", "C2"]);
    expect(cell("C2")).toBeUndefined();
  });

  it("Alt/Option+Enter inserts a line break at the caret", () => {
    const { editing, type, press } = setup(book());
    type("one");
    expect(press("Enter", { altKey: true }).defaultPrevented).toBe(true);
    expect([editing.mode.value, editing.text.value, editing.caret.value]).toEqual(["enter", "one\n", 4]);
    type("two");
    expect(editing.text.value).toBe("one\ntwo");
  });

  it("a click on another cell commits (a click inside the editor turns Enter mode into Edit)", () => {
    const { editing, type, select, cell } = setup(book());
    select("A1");
    type("abc");
    expect(editing.pressCell({ row: 0, col: 0 }, false, true)).toBe("native");
    expect(editing.mode.value).toBe("edit");
    expect(editing.pressCell({ row: 4, col: 2 }, false, false)).toBe("select");
    expect([editing.mode.value, cell("A1")?.value]).toEqual(["ready", "abc"]);
  });

  it("a header press commits first; a failing commit keeps the editor open", () => {
    const { editing, type, select, cell } = setup(book());
    select("A1");
    type("x");
    expect(editing.pressOutside()).toBe(true);
    expect(cell("A1")?.value).toBe("x");
    // With the caret after ")" no reference can be pointed: the click commits, which fails.
    type("=1+)");
    expect(editing.pressOutside()).toBe(false);
    expect([editing.mode.value, editing.error.value, editing.caret.value]).toEqual(["enter", FORMULA_ERROR_MESSAGE, 3]);
    editing.caret.value = 4;
    expect(editing.pressCell({ row: 4, col: 2 }, false, false)).toBe("handled");
    expect([editing.mode.value, editing.error.value, cell("A1")?.value]).toEqual(["enter", FORMULA_ERROR_MESSAGE, "x"]);
  });
});

describe("Ctrl/Cmd+Enter and Delete", () => {
  it("Ctrl+Enter writes the entry into every cell of the selection as one step \"Fill\"", () => {
    const { editing, undo, type, press, select, cell, ranges } = setup(book());
    select("D1:D3");
    type("x");
    expect(press("Enter", { ctrlKey: true }).defaultPrevented).toBe(true);
    expect([cell("D1")?.value, cell("D2")?.value, cell("D3")?.value]).toEqual(["x", "x", "x"]);
    expect(editing.mode.value).toBe("ready");
    expect(ranges()).toEqual([{ top: 0, left: 3, bottom: 2, right: 3 }]);
    expect(undo.undoLabel.value).toBe("Fill");
    undo.undo();
    expect([cell("D1"), cell("D2"), cell("D3")]).toEqual([undefined, undefined, undefined]);
    expect(undo.canUndo.value).toBe(false);
  });

  it("Cmd+Enter fills too, each formula shifted relatively from the edited cell", () => {
    platform.mac = true;
    const { type, press, select, cell } = setup(book({ A1: 1, A2: 2, A3: 3 }));
    select("B1:B3");
    type("=A1*2");
    press("Enter", { metaKey: true });
    expect([cell("B1")?.formula, cell("B2")?.formula, cell("B3")?.formula]).toEqual(["A1*2", "A2*2", "A3*2"]);
    expect([cell("B1")?.value, cell("B2")?.value, cell("B3")?.value]).toEqual([2, 4, 6]);
  });

  it("Delete or Backspace in Ready clears the selection's values and formulas as one step \"Clear contents\", keeping styles", () => {
    const { editing, undo, press, select, cell } = setup(book({ A1: { value: "styled", styleId: 1 }, A2: { value: 2, formula: "1+1" }, B1: true, C5: "outside" }, { extraStyles: [{ font: { bold: true } }] }));
    select("A1:B2");
    expect(press("Delete").defaultPrevented).toBe(true);
    expect(editing.mode.value).toBe("ready");
    expect(cell("A1")).toEqual({ col: 0, value: null, styleId: 1 });
    expect([cell("A2"), cell("B1"), cell("C5")?.value]).toEqual([undefined, undefined, "outside"]);
    expect(undo.undoLabel.value).toBe("Clear contents");
    undo.undo();
    expect([cell("A1")?.value, cell("A2")?.formula, cell("B1")?.value]).toEqual(["styled", "1+1", true]);

    select("C5");
    press("Backspace");
    expect(cell("C5")).toBeUndefined();
    expect(undo.undoLabel.value).toBe("Clear contents");
  });

  it("Delete over a whole column refits only the rows it cleared; a row it did not touch keeps the height its file gave it", () => {
    // 16.5 pt from the file; the measurer finds one 20 px line (the default 15 pt).
    const rows: Row[] = [
      { index: 0, cells: [{ col: 0, value: "a" }, { col: 1, value: "b" }], heightPt: 16.5 },
      { index: 1, cells: [{ col: 1, value: "c" }], heightPt: 16.5 },
    ];
    const { undo, press, select, model } = setup(book({}, { sheet: { rows } }));
    select("A1:A1048576");
    press("Delete");
    expect(model.sheets[0]!.rows[0]).toEqual({ index: 0, cells: [{ col: 1, value: "b" }] });
    expect(model.sheets[0]!.rows[1]).toBe(rows[1]);
    expect(undo.undoLabel.value).toBe("Clear contents");
    undo.undo();
    expect(model.sheets[0]!.rows).toEqual(rows);
  });

  it("SPEC-128: Delete over filter-hidden rows clears the visible rows only, as one step", () => {
    const rows: Row[] = [
      { index: 0, cells: [{ col: 0, value: "West" }, { col: 1, value: 1 }] },
      { index: 1, cells: [{ col: 0, value: "East" }, { col: 1, value: 2 }], filterHidden: true },
      { index: 2, cells: [{ col: 0, value: "North" }, { col: 1, value: 3 }], hidden: true },
      { index: 3, cells: [{ col: 0, value: "West" }, { col: 1, value: 4 }] },
    ];
    const { undo, press, select, model } = setup(book({}, { sheet: { rows } }));
    select("A1:B4");
    press("Delete");
    // The user-hidden row is no filter's: it is cleared, as without a filter.
    expect(model.sheets[0]!.rows).toEqual([rows[1], { index: 2, cells: [], hidden: true }]);
    expect(model.sheets[0]!.rows[0]).toBe(rows[1]);
    expect(undo.undoLabel.value).toBe("Clear contents");
    undo.undo();
    expect(model.sheets[0]!.rows).toEqual(rows);
  });

  it("Delete and Backspace inside the editor are text editing, not Clear contents", () => {
    const { editing, undo, type, press, select, cell } = setup(book({ A1: "keep" }));
    select("A1");
    press("F2");
    type("!");
    for (const key of ["Delete", "Backspace"]) {
      expect(press(key).defaultPrevented).toBe(false);
    }
    expect(editing.mode.value).toBe("edit");
    expect(cell("A1")?.value).toBe("keep");
    expect(undo.canUndo.value).toBe(false);
  });
});

describe("IME composition", () => {
  it("ignores Enter, Tab, the arrows and Escape while a composition is active (isComposing, keyCode 229)", () => {
    const { editing, type, press, select, cell, active } = setup(book());
    select("A1");
    type("か");
    for (const init of [{ isComposing: true }, { keyCode: 229 }]) {
      for (const key of ["Enter", "Tab", "ArrowDown", "Escape"]) {
        const event = press(key, init);
        // The editor's (nothing reaches the grid), untouched by it.
        expect(event.defaultPrevented, `${key} ${JSON.stringify(init)}`).toBe(false);
      }
    }
    expect([editing.mode.value, editing.text.value, active()]).toEqual(["enter", "か", "A1"]);
    expect(cell("A1")).toBeUndefined();
  });

  it("ignores keys between compositionstart and compositionend, whatever the event says", () => {
    const { editing, type, press, select, cell, active } = setup(book());
    select("A1");
    type("a");
    editing.compositionStart();
    press("Enter");
    press("Escape");
    expect([editing.mode.value, active()]).toEqual(["enter", "A1"]);
    editing.compositionEnd();
    editing.input("a漢", 2);
    press("Enter");
    expect([cell("A1")?.value, active()]).toEqual(["a漢", "A2"]);
  });
});

describe("commit", () => {
  it("parses per Input parsing: grouped numbers, percent, ISO dates; the suggested format lands on a General cell", () => {
    const { type, press, select, cell, model } = setup(book());
    const format = (address: string) => {
      const style = cellStyle(model, cell(address)?.styleId ?? 0);
      return [style.numberFormat, style.builtinFormatId];
    };
    select("A2");
    type("1,200");
    press("Tab");
    expect(cell("A2")?.value).toBe(1200);
    expect(format("A2")).toEqual(["#,##0", undefined]);
    select("C1");
    type("10%");
    press("Enter");
    expect(cell("C1")?.value).toBe(0.1);
    expect(format("C1")).toEqual(["0%", undefined]);
    type("2026-10-08");
    press("Enter");
    expect(cell("C2")?.value).toBe(46303);
    expect(format("C2")[1]).toBe(14);
    type("'007");
    press("Enter");
    expect(cell("C3")?.value).toBe("007");
    expect(cellStyle(model, cell("C3")?.styleId ?? 0).quotePrefix).toBe(true);
    type("true");
    press("Enter");
    expect(cell("C4")?.value).toBe(true);
  });

  it("a percent-formatted cell turns \"10\" into 0.1 and a Text cell keeps \"=A1\" as text", () => {
    const { type, press, select, cell } = setup(book({ A1: { value: null, styleId: 1 }, A2: { value: null, styleId: 2 } }, { extraStyles: [{ numberFormat: "0%" }, { numberFormat: "@" }] }));
    select("A1");
    type("10");
    press("Enter");
    type("=A1");
    press("Enter");
    expect([cell("A1")?.value, cell("A2")?.value, cell("A2")?.formula]).toEqual([0.1, "=A1", undefined]);
  });

  it("stores a formula canonically and closes unbalanced trailing \"(\"", () => {
    const { type, press, select, cell } = setup(book({ A1: 2 }));
    select("B1");
    type("=sum(a1, 3)");
    press("Enter");
    expect(cell("B1")?.formula).toBe("SUM(A1,3)");
    expect(cell("B1")?.value).toBe(5);
    type("=SUM(1,2");
    press("Enter");
    expect([cell("B2")?.formula, cell("B2")?.value]).toEqual(["SUM(1,2)", 3]);
  });

  it("a syntax error keeps the editor open with the message and the caret at the error", () => {
    const { editing, undo, type, press, select, cell, active } = setup(book());
    select("E1");
    type("=1+)+2");
    press("Enter");
    expect(editing.mode.value).toBe("enter");
    expect(editing.error.value).toBe(FORMULA_ERROR_MESSAGE);
    // The error is at ")" — after "=1+".
    expect(editing.caret.value).toBe(3);
    expect([cell("E1"), active(), undo.canUndo.value]).toEqual([undefined, "E1", false]);
    // Typing goes on from the error; the message goes with it.
    editing.input("=1+2", 4);
    expect(editing.error.value).toBe("");
    press("Escape");
    expect(cell("E1")).toBeUndefined();
  });

  it("a formula over 8,192 characters shows its message; one of 8,192 is stored", () => {
    const { editing, press, select, cell } = setup(book());
    select("A1");
    const atLimit = `"${"a".repeat(8190)}"`;
    editing.input(`=${atLimit}&1`, atLimit.length + 3);
    press("Enter");
    expect([editing.mode.value, editing.error.value]).toEqual(["enter", FORMULA_TOO_LONG_MESSAGE]);
    expect(cell("A1")).toBeUndefined();
    editing.input(`=${atLimit}`, atLimit.length + 1);
    press("Enter");
    expect(cell("A1")?.formula).toHaveLength(8192);
  });

  it("text over 32,767 characters cannot be typed", () => {
    const { editing, type, press, select, cell } = setup(book());
    select("A1");
    type("a");
    expect(editing.input("a".repeat(32_768), 32_768)).toBe(false);
    expect(editing.text.value).toBe("a");
    expect(editing.input("a".repeat(32_767), 32_767)).toBe(true);
    // At the limit, a line break does not fit either.
    press("Enter", { altKey: true });
    expect(editing.text.value).toHaveLength(32_767);
    press("Enter");
    expect(cell("A1")?.value).toHaveLength(32_767);
  });

  it("is one undo step labelled with the first 20 characters and the address, scoped to the sheet", () => {
    const { undo, type, press, select, cell } = setup(withSheet(book(), "Other"));
    select("B3");
    type("abcdefghijklmnopqrstuvwxyz");
    press("Enter");
    expect(undo.undoLabel.value).toBe('Typing "abcdefghijklmnopqrst" in B3');
    type("=A2*2");
    press("Enter");
    expect(undo.undoLabel.value).toBe('Typing "=A2*2" in B4');
    undo.undo();
    expect(cell("B4")).toBeUndefined();
    undo.undo();
    expect(cell("B3")).toBeUndefined();
    expect(undo.canUndo.value).toBe(false);
  });

  it("recalculates the edited cell's dependents on every sheet and marks the changed sheets dirty", () => {
    const model = withSheet(book({ A1: 5, B1: { value: 10, formula: "A1*2" } }), "Data", { A1: { value: 6, formula: "Sheet1!A1+1" } });
    const { recalls, type, press, select, cell } = setup(model);
    expect(model.sheets.map((sheet) => sheet.dirty)).toEqual([undefined, undefined]);
    select("A1");
    type("7");
    press("Enter");
    expect(recalls).toHaveLength(1);
    expect(recalls[0]).toContain(cellKey(model.sheets[0]!.id, 0, 0));
    expect([cell("A1")?.value, cell("B1")?.value, cell("A1", 1)?.value]).toEqual([7, 14, 8]);
    expect(model.sheets.map((sheet) => sheet.dirty)).toEqual([true, true]);
  });

  it("an unchanged commit records nothing — the same text, or text that parses to the same entry", () => {
    const { undo, recalls, type, press, select, cell, model, active } = setup(book({ A1: "abc", A2: 5 }));
    const sheets = model.sheets;
    select("A1");
    press("F2");
    press("Enter");
    expect(active()).toBe("A2");
    type("5.0");
    press("Enter");
    expect(model.sheets).toBe(sheets);
    expect([undo.canUndo.value, recalls.length, model.sheets[0]!.dirty]).toEqual([false, 0, undefined]);
    expect([cell("A1")?.value, cell("A2")?.value]).toEqual(["abc", 5]);
  });

  it("an unchanged commit records nothing, even on a row whose own height is not the measured one", () => {
    // 16.5 pt from the file; the measurer finds one 20 px line (the default 15 pt).
    const { undo, recalls, type, press, select, model } = setup(book({}, { sheet: { rows: [{ index: 1, cells: [{ col: 0, value: 5 }], heightPt: 16.5 }] } }));
    const sheets = model.sheets;
    select("A2");
    type("5.0");
    press("Enter");
    expect(model.sheets).toBe(sheets);
    expect([undo.canUndo.value, recalls.length, model.sheets[0]!.rows[0]!.heightPt]).toEqual([false, 0, 16.5]);
  });

  it("a changed rich-text cell stores plain text; an unchanged one keeps its runs", () => {
    const runs: TextRun[] = [{ text: "Hello ", font: { bold: true } }, { text: "world" }];
    const { type, press, select, cell } = setup(book({ A1: { value: "Hello world", runs } }));
    select("A1");
    press("F2");
    press("Enter");
    expect(cell("A1")?.runs).toEqual(runs);
    select("A1");
    press("F2");
    type("!");
    press("Enter");
    expect(cell("A1")).toEqual({ col: 0, value: "Hello world!" });
  });

  it("a line break turns wrap on through the cell's style and refits the row inside the same step", () => {
    const { undo, type, press, select, cell, model } = setup(book());
    select("C3");
    type("one");
    press("Enter", { altKey: true });
    type("two");
    press("Enter");
    expect(cell("C3")?.value).toBe("one\ntwo");
    expect(cellStyle(model, cell("C3")?.styleId ?? 0).alignment?.wrap).toBe(true);
    // Two 20 px lines: 40 px = 30 pt, never customHeight.
    const row = model.sheets[0]!.rows.find((r) => r.index === 2);
    expect([row?.heightPt, row?.customHeight]).toEqual([30, undefined]);
    undo.undo();
    expect(model.sheets[0]!.rows).toEqual([]);
  });

  it("a cell edited on its own keeps the rest of the sheet's rows and cells untouched (copy-on-write)", () => {
    const { type, press, select, model } = setup(book({ A1: "a", A5: "e" }));
    const [row1, row5] = model.sheets[0]!.rows;
    select("C3");
    type("c");
    press("Enter");
    const rows = model.sheets[0]!.rows;
    expect(rows.map((row) => row.index)).toEqual([0, 2, 4]);
    expect(rows[0]).toBe(row1);
    expect(rows[2]).toBe(row5);
  });
});

describe("the draft's own undo while editing", () => {
  it("Ctrl+Z / Ctrl+Y step the draft — prevented, so the workbook's undo listener ignores them — and the workbook undo applies after the commit", () => {
    const { editing, undo, type, press, select, cell } = setup(book());
    select("A1");
    type("a");
    press("Enter");
    expect(undo.undoLabel.value).toBe('Typing "a" in A1');
    type("xy");
    const back = press("z", { ctrlKey: true });
    expect(back.defaultPrevented).toBe(true);
    expect(editing.text.value).toBe("x");
    // Reaching the window, the prevented key undoes nothing in the workbook.
    window.dispatchEvent(back);
    expect([cell("A1")?.value, undo.canUndo.value]).toEqual(["a", true]);
    const again = press("y", { ctrlKey: true });
    window.dispatchEvent(again);
    expect([editing.text.value, editing.mode.value]).toEqual(["xy", "enter"]);
    expect(cell("A1")?.value).toBe("a");
    press("Escape");
    // In Ready the key is the workbook's.
    const ready = press("z", { ctrlKey: true });
    expect(ready.defaultPrevented).toBe(false);
    window.dispatchEvent(ready);
    expect(cell("A1")).toBeUndefined();
  });
});

describe("protection", () => {
  const protectedBook = () =>
    book({ A1: "locked", B1: { value: "free", styleId: 1 }, B2: { value: null, styleId: 1 } }, { extraStyles: [{ protection: { locked: false } }], sheet: { protection: { sheet: true } } });

  it("a locked cell on a protected sheet refuses the editor — F2, typing, a composition — with the toast", () => {
    const { editing, toasts, type, press, select } = setup(protectedBook());
    select("A1");
    press("F2");
    expect(editing.mode.value).toBe("ready");
    expect(editing.input("x", 1)).toBe(false);
    type("y");
    expect(editing.compositionStart()).toBe(false);
    editing.compositionEnd();
    expect(editing.mode.value).toBe("ready");
    expect(toasts).toEqual([PROTECTED_CELL_MESSAGE, PROTECTED_CELL_MESSAGE, PROTECTED_CELL_MESSAGE, PROTECTED_CELL_MESSAGE]);
    // An empty position takes the sheet's default style: locked too.
    select("D4");
    press("F2");
    expect(toasts).toHaveLength(5);
  });

  it("refuses Delete over a selection holding a locked cell, clearing nothing", () => {
    const { undo, toasts, press, select, cell } = setup(protectedBook());
    select("A1:B1");
    press("Delete");
    expect([cell("A1")?.value, cell("B1")?.value]).toEqual(["locked", "free"]);
    expect(toasts).toEqual([PROTECTED_CELL_MESSAGE]);
    expect(undo.canUndo.value).toBe(false);
  });

  it("refuses a Fill whose selection holds a locked cell, keeping the editor open", () => {
    const harness = setup(protectedBook());
    const { editing, undo, toasts, type, press, cell } = harness;
    // B1:B2 is unlocked; A1 is not. B2:A1 with B2 active.
    harness.selection.press({ row: 1, col: 1 }, { shift: false, mod: false });
    harness.selection.press({ row: 0, col: 0 }, { shift: true, mod: false });
    harness.selection.release();
    expect(harness.active()).toBe("B2");
    type("z");
    press("Enter", { ctrlKey: true });
    expect(toasts).toEqual([PROTECTED_CELL_MESSAGE]);
    expect(editing.mode.value).toBe("enter");
    expect([cell("A1")?.value, cell("B1")?.value, cell("B2")?.value]).toEqual(["locked", "free", null]);
    expect(undo.canUndo.value).toBe(false);
  });

  it("unlocked cells edit normally", () => {
    const { editing, toasts, type, press, select, cell } = setup(protectedBook());
    select("B1");
    press("F2");
    expect(editing.mode.value).toBe("edit");
    type("dom");
    press("Enter");
    expect(cell("B1")?.value).toBe("freedom");
    type("x");
    press("Tab");
    expect(cell("B2")?.value).toBe("x");
    select("B1:B2");
    press("Delete");
    expect([cell("B1")?.value, cell("B2")?.value]).toEqual([null, null]);
    expect(toasts).toEqual([]);
  });
});

describe("arrays", () => {
  const arrayBook = () =>
    book({
      A1: 1,
      A2: 2,
      A3: 3,
      B1: { value: 2, formula: "A1:A3*2", arrayRange: { top: 0, left: 1, bottom: 2, right: 1 } },
      B2: 4,
      B3: 6,
    });

  it("refuses editing or clearing part of an array with the toast", () => {
    const { editing, toasts, type, press, select, cell } = setup(arrayBook());
    select("B2");
    press("F2");
    type("9");
    press("Delete");
    expect(editing.mode.value).toBe("ready");
    expect(toasts).toEqual([ARRAY_PART_MESSAGE, ARRAY_PART_MESSAGE, ARRAY_PART_MESSAGE]);
    expect(cell("B2")?.value).toBe(4);
  });

  it("the anchor edits as its formula, without braces, and keeps its arrayRange", () => {
    const { editing, press, select, cell } = setup(arrayBook());
    select("B1");
    press("F2");
    expect(editing.text.value).toBe("=A1:A3*2");
    editing.input("=A1:A3*3", 8);
    press("Enter");
    expect(cell("B1")?.formula).toBe("A1:A3*3");
    expect(cell("B1")?.arrayRange).toEqual({ top: 0, left: 1, bottom: 2, right: 1 });
  });
});

describe("sheets", () => {
  it("choosing another tab while not pointing commits first; a failing commit keeps the sheet", () => {
    const { editing, activeIndex, type, select, cell } = setup(withSheet(book(), "Other"));
    select("A1");
    type("x");
    editing.selectSheet(1);
    expect([cell("A1")?.value, activeIndex.value, editing.mode.value]).toEqual(["x", 1, "ready"]);
    editing.selectSheet(0);
    type("=1+)");
    editing.selectSheet(1);
    expect([activeIndex.value, editing.mode.value, editing.error.value]).toEqual([0, "enter", FORMULA_ERROR_MESSAGE]);
  });

  it("Ctrl+Enter while pointing on another sheet fills the edited sheet's selection, not the one on show", () => {
    const { editing, activeIndex, selection, type, press, select, cell } = setup(withSheet(book(), "Other"));
    select("D1:D2");
    type("=");
    editing.selectSheet(1);
    expect([activeIndex.value, editing.mode.value]).toEqual([1, "enter"]);
    // The sheet on show has a selection of its own.
    selection.set({ active: { row: 4, col: 5 }, ranges: [{ top: 4, left: 5, bottom: 8, right: 5 }], activeRangeIndex: 0 });
    type("7");
    press("Enter", { ctrlKey: true });
    expect([activeIndex.value, editing.mode.value]).toEqual([0, "ready"]);
    expect([cell("D1")?.formula, cell("D2")?.formula]).toEqual(["7", "7"]);
    expect([cell("F5"), cell("F5", 1)]).toEqual([undefined, undefined]);
  });

  it("is a grid sheet's only: nothing opens without a workbook", async () => {
    const { editing, workbook, select } = setup(book());
    select("A1");
    workbook.value = null;
    await nextTick();
    expect(editing.openEdit()).toBe(false);
    expect(editing.input("x", 1)).toBe(false);
  });
});
