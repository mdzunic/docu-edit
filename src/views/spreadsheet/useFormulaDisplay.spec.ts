// SPEC-119 "The formula bar", "Show formulas": useFormulaDisplay's formula-bar
// text — formulaBarText for the active cell, muted for a dynamic array's
// member — and the Show formulas toggle: the sheet's `showFormulas` view
// state, flipped by Ctrl+` matched on event.code with Ctrl and no other
// modifier on every platform (macOS included), ignored from inputs and
// dialogs and on a foreign sheet, never anything but view state.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, shallowRef, type EffectScope } from "vue";
import type { Cell, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import type { SheetViewFields } from "./useSheetViewState";

const platform = { mac: false };
vi.mock("../../lib/core/platform", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/core/platform")>()),
  isMacPlatform: () => platform.mac,
}));

const { useSheetMetrics } = await import("./useSheetMetrics");
const { useSheetSelection } = await import("./useSheetSelection");
const { isShowFormulasShortcut, useFormulaDisplay } = await import("./useFormulaDisplay");

/** Row 0: A1 =SUM(B2:B4); B1:B3 a legacy array; C1:C3 a dynamic array;
 * D1:D2 a data table; E1 an unparsed formula; F1 a date; G1 text. */
function formulaSheet(kind: Sheet["kind"] = "grid"): Sheet {
  const rows: Row[] = [
    {
      index: 0,
      cells: [
        { col: 0, value: 60, formula: "SUM(B2:B4)" },
        { col: 1, value: 6, formula: "A2:A4*2", arrayRange: { top: 0, left: 1, bottom: 2, right: 1 } },
        { col: 2, value: 1, formula: "SORT(A2:A6)", arrayRange: { top: 0, left: 2, bottom: 2, right: 2 }, dynamicArray: true, cm: 1 },
        { col: 3, value: 10, dataTableXml: `<f t="dataTable" ref="D1:D2" dt2D="0" dtr="1" r1="H1"/>` },
        { col: 4, value: 3, formula: "of:=FOO(", formulaUnparsed: true },
        { col: 5, value: 46303, styleId: 1 },
        { col: 6, value: "text" },
      ] satisfies Cell[],
    },
    { index: 1, cells: [{ col: 1, value: 2 }, { col: 2, value: 2 }, { col: 3, value: 20 }] },
  ];
  return { conditionalFormats: [], validations: [], id: "s1", name: "Sheet1", visibility: "visible", kind, rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

function bookOf(sheet: Sheet, format: SpreadsheetModel["format"] = "xlsx"): SpreadsheetModel {
  return {
    dxfs: [],
    kind: "spreadsheet",
    id: "book",
    title: "Book",
    format,
    readerVersion: 4,
    sheets: [sheet],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}, { numberFormat: "m/d/yyyy", builtinFormatId: 14 }],
    defaultStyleId: 0,
  };
}

let scope: EffectScope;
beforeEach(() => {
  scope = effectScope();
});
afterEach(() => {
  scope.stop();
  platform.mac = false;
  document.body.innerHTML = "";
});

/** The composable over a workbook, with an in-memory view state. */
function setup(sheet: Sheet = formulaSheet(), format: SpreadsheetModel["format"] = "xlsx") {
  const model = bookOf(sheet, format);
  const fields = shallowRef<SheetViewFields>({});
  const updates: SheetViewFields[] = [];
  const viewState = {
    active: computed(() => fields.value),
    updateActive(patch: SheetViewFields) {
      updates.push(patch);
      fields.value = { ...fields.value, ...patch };
    },
  };
  const result = scope.run(() => {
    const sheetMetrics = useSheetMetrics(() => sheet);
    const selection = useSheetSelection({ sheet: () => sheet, sheetMetrics, viewState, model: () => model, locale: () => SPREADSHEET_LOCALES["en-US"] });
    return { selection, formulas: useFormulaDisplay({ sheet: () => sheet, model: () => model, viewState, selection }) };
  })!;
  const select = (row: number, col: number): void => {
    result.selection.press({ row, col }, { shift: false, mod: false });
    result.selection.release();
  };
  return { ...result, fields, updates, select };
}

/** Ctrl+` (or a variation) dispatched at `target` (default: the window). */
function press(init: KeyboardEventInit, target: EventTarget = window): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { code: "Backquote", key: "`", bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  return event;
}

describe("the formula bar's text", () => {
  it("is formulaBarText for the active cell: '=' + formula, arrays, data tables, unparsed as stored, editText otherwise", () => {
    const { formulas, select } = setup();
    expect(formulas.barText.value).toBe("=SUM(B2:B4)");
    select(1, 1);
    expect(formulas.barText.value).toBe("{=A2:A4*2}");
    select(1, 3);
    expect(formulas.barText.value).toBe("{=TABLE(H1,)}");
    select(0, 4);
    expect(formulas.barText.value).toBe("=of:=FOO(");
    select(0, 5);
    expect(formulas.barText.value).toBe("10/8/2026");
    select(0, 6);
    expect(formulas.barText.value).toBe("text");
    select(40, 40);
    expect(formulas.barText.value).toBe("");
  });

  it("SPEC-122: is empty for a hidden formula on a protected sheet only", () => {
    const sheet = formulaSheet();
    const hidden = { ...sheet, rows: sheet.rows.map((row) => (row.index === 0 ? { ...row, cells: row.cells.map((cell) => (cell.col === 0 ? { ...cell, styleId: 2 } : cell)) } : row)) };
    const run = (protectedSheet: boolean) => {
      const model = bookOf(protectedSheet ? { ...hidden, protection: { sheet: true } } : hidden);
      model.styles = [...model.styles, { protection: { hidden: true } }];
      const target = model.sheets[0]!;
      const fields = shallowRef<SheetViewFields>({});
      const viewState = { active: computed(() => fields.value), updateActive: () => {} };
      return scope.run(() => {
        const sheetMetrics = useSheetMetrics(() => target);
        const selection = useSheetSelection({ sheet: () => target, sheetMetrics, viewState, model: () => model, locale: () => SPREADSHEET_LOCALES["en-US"] });
        return useFormulaDisplay({ sheet: () => target, model: () => model, viewState, selection }).barText.value;
      });
    };
    expect(run(false)).toBe("=SUM(B2:B4)");
    expect(run(true)).toBe("");
  });

  it("shows an ods workbook's unparsed formula from its first '='", () => {
    const { formulas, select } = setup(formulaSheet(), "ods");
    select(0, 4);
    expect(formulas.barText.value).toBe("=FOO(");
  });

  it("marks a dynamic array's member — not its anchor — as a spill member", () => {
    const { formulas, select } = setup();
    select(0, 2);
    expect([formulas.barText.value, formulas.spillMember.value]).toEqual(["=SORT(A2:A6)", false]);
    select(1, 2);
    expect([formulas.barText.value, formulas.spillMember.value]).toEqual(["=SORT(A2:A6)", true]);
    select(1, 1);
    expect(formulas.spillMember.value).toBe(false);
  });
});

describe("Show formulas", () => {
  it("toggles the sheet's showFormulas view state — on, then off — and nothing else", () => {
    const { formulas, fields, updates } = setup();
    expect(formulas.showFormulas.value).toBe(false);
    formulas.toggle();
    expect(formulas.showFormulas.value).toBe(true);
    expect(fields.value.showFormulas).toBe(true);
    formulas.toggle();
    expect(formulas.showFormulas.value).toBe(false);
    expect(fields.value.showFormulas).toBeUndefined();
    expect(updates).toEqual([{ showFormulas: true }, { showFormulas: undefined }]);
  });

  it("toggles on Ctrl+` — matched on event.code, so a dead-key layout works — and consumes the key", async () => {
    const { formulas } = setup();
    const event = press({ ctrlKey: true });
    expect(event.defaultPrevented).toBe(true);
    expect(formulas.showFormulas.value).toBe(true);
    press({ ctrlKey: true, key: "Dead" });
    await nextTick();
    expect(formulas.showFormulas.value).toBe(false);
  });

  it("is Control on macOS too — Cmd+` is the system's and never toggles", () => {
    platform.mac = true;
    const { formulas } = setup();
    press({ metaKey: true });
    expect(formulas.showFormulas.value).toBe(false);
    press({ ctrlKey: true });
    expect(formulas.showFormulas.value).toBe(true);
  });

  it("ignores Ctrl+Shift+`, Ctrl+Alt+`, Ctrl+Cmd+`, a bare `, and Ctrl with another key", () => {
    const { formulas } = setup();
    for (const init of [
      { ctrlKey: true, shiftKey: true },
      { ctrlKey: true, altKey: true },
      { ctrlKey: true, metaKey: true },
      {},
      { ctrlKey: true, code: "KeyA", key: "a" },
      { ctrlKey: true, code: "Quote", key: "`" },
    ]) {
      const event = press(init);
      expect(event.defaultPrevented, JSON.stringify(init)).toBe(false);
    }
    expect(formulas.showFormulas.value).toBe(false);
    expect(isShowFormulasShortcut({ code: "Backquote", ctrlKey: true, shiftKey: false, altKey: false, metaKey: false })).toBe(true);
  });

  it("is ignored from an input, a textarea, a select and inside a dialog (isOutsideDocumentInput)", () => {
    const { formulas } = setup();
    const dialog = document.createElement("div");
    dialog.setAttribute("role", "dialog");
    const button = document.createElement("button");
    dialog.append(button);
    for (const target of [document.createElement("input"), document.createElement("textarea"), document.createElement("select"), button]) {
      document.body.append(target.closest('[role="dialog"]') ?? target);
      press({ ctrlKey: true }, target);
    }
    expect(formulas.showFormulas.value).toBe(false);
    press({ ctrlKey: true }, document.body);
    expect(formulas.showFormulas.value).toBe(true);
  });

  it("does nothing on a foreign sheet", () => {
    const { formulas, updates } = setup(formulaSheet("foreign"));
    press({ ctrlKey: true });
    formulas.toggle();
    expect(updates).toEqual([]);
    expect(formulas.barText.value).toBe("");
  });

  it("stops listening when its scope ends", () => {
    const { formulas } = setup();
    scope.stop();
    press({ ctrlKey: true });
    expect(formulas.showFormulas.value).toBe(false);
  });
});
