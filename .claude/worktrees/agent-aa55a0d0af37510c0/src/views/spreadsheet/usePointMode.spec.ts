// SPEC-122 "Point mode": usePointMode as useCellEditing drives it — where a
// reference may be pointed (after "=", an operator, "(" or ","), arrows that
// insert and then replace the reference at the caret, Shift+arrows that
// extend it to a range, Ctrl/Cmd+arrows that jump to the data edge, clicks
// and drags, sheet-qualified references when another sheet is shown (quoted
// as SPEC-119 prints them) with commit and cancel returning to the edited
// sheet, typing that ends the point, F4's `$` cycle, and the highlights the
// grid outlines: colour by first appearance modulo 8, only references on
// the displayed sheet, the pointed one dashed.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef } from "vue";
import { getCell, parseA1 } from "../../lib/spreadsheet/address";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import { recalculate } from "../../lib/spreadsheet/formula/recalc";
import type { Cell, CellStyle, CellValue, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { addSheet, createBlankWorkbook } from "../../lib/spreadsheet/sheets";
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

  return { model, workbook, activeIndex, selection, undo, editing, toasts, press, type, select, cell, active, ranges };
}

describe("where a reference may be pointed", () => {
  it("after \"=\", an operator, \"(\" or \",\" — and never in a text or after a value", () => {
    const { editing, type, select } = setup(book());
    select("C3");
    for (const [text, allowed] of [
      ["=", true],
      ["=1+", true],
      ["=SUM(", true],
      ["=SUM(A1,", true],
      ["=A1<>", true],
      ["=-", true],
      ["=1", false],
      ["=A1", false],
      ['="x', false],
    ] as const) {
      editing.input(text, text.length);
      expect([text, editing.point.allowed()]).toEqual([text, allowed]);
    }
    editing.cancel();
    type("abc");
    expect(editing.point.allowed()).toBe(false);
  });
});

describe("arrows", () => {
  it("insert a reference at the caret, then replace it; Shift extends it to a range (the e2e's =SUM(B1:B2))", () => {
    const { editing, type, press, select, cell } = setup(book({ B1: 1000, B2: 1400 }));
    select("B3");
    type("=SUM(");
    expect(press("ArrowUp").defaultPrevented).toBe(true);
    expect(editing.text.value).toBe("=SUM(B2");
    press("ArrowUp");
    expect(editing.text.value).toBe("=SUM(B1");
    press("ArrowDown", { shiftKey: true });
    expect(editing.text.value).toBe("=SUM(B1:B2");
    type(")");
    expect(editing.text.value).toBe("=SUM(B1:B2)");
    expect(editing.point.pointing.value).toBe(false);
    press("Enter");
    expect([cell("B3")?.formula, cell("B3")?.value]).toEqual(["SUM(B1:B2)", 2400]);
  });

  it("typing ends the point: the next arrow after an operator inserts a new reference", () => {
    const { editing, type, press, select } = setup(book());
    select("C3");
    type("=");
    press("ArrowLeft");
    expect(editing.text.value).toBe("=B3");
    type("+");
    press("ArrowRight");
    expect(editing.text.value).toBe("=B3+D3");
    // After a reference that is no longer pointed, an arrow commits and moves.
    type("*2");
    press("ArrowDown");
    expect(editing.mode.value).toBe("ready");
  });

  it("Ctrl+arrows (Cmd on macOS) jump to the data edge", () => {
    const data = { A1: 1, A2: 2, A3: 3, A4: 4 };
    const { editing, type, press, select } = setup(book(data));
    select("A1");
    press("ArrowRight");
    type("=");
    press("ArrowLeft");
    press("ArrowDown", { ctrlKey: true });
    expect(editing.text.value).toBe("=A4");
    editing.cancel();
    platform.mac = true;
    select("B1");
    type("=");
    press("ArrowLeft");
    press("ArrowDown", { metaKey: true, shiftKey: true });
    expect(editing.text.value).toBe("=A1:A4");
  });
});

describe("clicks and drags", () => {
  it("a click inserts the cell's reference; a drag extends it; Shift+click extends too", () => {
    const { editing, type, select } = setup(book());
    select("D5");
    type("=");
    expect(editing.pressCell({ row: 0, col: 0 }, false, false)).toBe("handled");
    expect(editing.text.value).toBe("=A1");
    editing.pointDrag({ row: 2, col: 1 });
    expect(editing.text.value).toBe("=A1:B3");
    editing.pointRelease();
    editing.pointDrag({ row: 5, col: 5 });
    expect(editing.text.value).toBe("=A1:B3");
    editing.pressCell({ row: 3, col: 0 }, true, false);
    expect(editing.text.value).toBe("=A1:A4");
  });
});

describe("another sheet", () => {
  it("a tab chosen while pointing shows that sheet; references there are sheet-qualified as SPEC-119 quotes them; commit returns to the edited sheet", () => {
    const model = withSheet(book(), "My Sheet", { B2: 5 });
    const { editing, activeIndex, type, press, select, cell } = setup(model);
    select("A1");
    type("=");
    editing.selectSheet(1);
    expect([activeIndex.value, editing.mode.value]).toEqual([1, "enter"]);
    expect(editing.onShownSheet.value).toBe(false);
    editing.pressCell({ row: 1, col: 1 }, false, false);
    expect(editing.text.value).toBe("='My Sheet'!B2");
    press("ArrowDown");
    expect(editing.text.value).toBe("='My Sheet'!B3");
    press("Enter");
    expect(activeIndex.value).toBe(0);
    expect(cell("A1")?.formula).toBe("'My Sheet'!B3");
    expect(cell("B2", 1)?.value).toBe(5);
  });

  it("cancel returns to the edited sheet too", () => {
    const { editing, activeIndex, type, select, cell } = setup(withSheet(book(), "Data"));
    select("A1");
    type("=");
    editing.selectSheet(1);
    editing.pressCell({ row: 0, col: 0 }, false, false);
    expect(editing.text.value).toBe("=Data!A1");
    editing.cancel();
    expect([activeIndex.value, editing.mode.value, cell("A1")]).toEqual([0, "ready", undefined]);
  });
});

describe("F4", () => {
  it("cycles the reference at the caret A1 → $A$1 → A$1 → $A1 → A1", () => {
    const { editing, press, select } = setup(book());
    select("C1");
    editing.input("=A1+B2", 3);
    const seen: string[] = [];
    for (let i = 0; i < 4; i += 1) {
      expect(press("F4").defaultPrevented).toBe(true);
      seen.push(editing.text.value);
    }
    expect(seen).toEqual(["=$A$1+B2", "=A$1+B2", "=$A1+B2", "=A1+B2"]);
  });
});

describe("highlights", () => {
  it("colour references by first appearance modulo 8 and outline those on the displayed sheet, the pointed one dashed", () => {
    const model = withSheet(book(), "Data");
    const { editing, press, select } = setup(model);
    select("J1");
    const text = "=A1+B1+C1+D1+E1+F1+G1+H1+I1+a1+Data!A1+";
    editing.input(text, text.length);
    // I1 is the ninth distinct reference (back to the first colour); "a1" is A1's; Data!A1 is the tenth.
    expect(editing.point.spans.value.map((span) => span.colorIndex)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 0, 0, 1]);
    press("ArrowDown");
    expect(editing.text.value).toBe(`${text}J2`);
    const shown = editing.point.highlights.value;
    // Data!A1 is on another sheet: not outlined here.
    expect(shown).toHaveLength(11);
    expect(shown.at(-1)).toEqual({ range: { top: 1, left: 9, bottom: 1, right: 9 }, colorIndex: 2, dashed: true });
    expect(shown.filter((mark) => mark.dashed)).toHaveLength(1);
    expect(shown[0]).toEqual({ range: { top: 0, left: 0, bottom: 0, right: 0 }, colorIndex: 0, dashed: false });
    // On Data only Data!A1 is outlined.
    editing.selectSheet(1);
    expect(editing.point.highlights.value).toEqual([{ range: { top: 0, left: 0, bottom: 0, right: 0 }, colorIndex: 1, dashed: false }]);
  });

  it("a range reference is outlined over its whole range; nothing is outlined outside a formula", () => {
    const { editing, select } = setup(book());
    select("D1");
    editing.input("=SUM(A1:B3)", 11);
    expect(editing.point.highlights.value).toEqual([{ range: { top: 0, left: 0, bottom: 2, right: 1 }, colorIndex: 0, dashed: false }]);
    editing.input("A1:B3", 5);
    expect([editing.point.spans.value, editing.point.highlights.value]).toEqual([[], []]);
  });
});
