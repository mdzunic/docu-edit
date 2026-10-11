// SPEC-117's selection composable: the state the grid, the Name Box, the
// formula bar and the status bar read, mapped from keys and pointer presses
// onto lib/spreadsheet/selection.ts and stored as view state — the file's
// selection restored, a press snapped to its merge, a drag shown as
// "<rows>R × <cols>C" until release, the keyboard map behind the platform's
// primary modifier (inputs ignored), and the aggregates over stored cells,
// deferred to idle time for a selection over 100,000 stored cells. (The
// formula bar's text is useFormulaDisplay's since SPEC-119; the bar's
// assertions here read it there, over this selection.)
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, shallowRef, type EffectScope } from "vue";
import type { Cell, Row, Sheet } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import type { SheetViewFields } from "./useSheetViewState";

const idle: Array<() => void> = [];
const platform = { mac: false };

vi.mock("../../lib/core/idleScheduler", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/core/idleScheduler")>()),
  runWhenIdle: (callback: () => void) => {
    idle.push(callback);
  },
}));
vi.mock("../../lib/core/platform", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../lib/core/platform")>()),
  isMacPlatform: () => platform.mac,
}));

const { useSheetMetrics } = await import("./useSheetMetrics");
const { useSheetSelection, AGGREGATE_DEFER_THRESHOLD } = await import("./useSheetSelection");
const { useFormulaDisplay } = await import("./useFormulaDisplay");

/** A sheet from [row, col, value] triples (0-based), plus layout extras. */
function sheetOf(cells: ReadonlyArray<readonly [number, number, Cell["value"]]>, extra: Partial<Sheet> = {}): Sheet {
  const rows: Row[] = [];
  for (const [index, col, value] of cells.toSorted((a, b) => a[0] - b[0] || a[1] - b[1])) {
    const last = rows.at(-1);
    if (last?.index === index) {
      last.cells.push({ col, value });
    } else {
      rows.push({ index, cells: [{ col, value }] });
    }
  }
  return { conditionalFormats: [], id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...extra };
}

/** B3:B12 = 10, 20, … 100; A1 a two-line string; C5:E5 merged. */
function styledSheet(extra: Partial<Sheet> = {}): Sheet {
  const cells: Array<readonly [number, number, Cell["value"]]> = [[0, 0, "Line one\nLine two"]];
  for (let i = 0; i < 10; i += 1) {
    cells.push([2 + i, 1, (i + 1) * 10]);
  }
  cells.push([4, 2, "Merged title"]);
  return sheetOf(cells, { merges: [{ top: 4, left: 2, bottom: 4, right: 4 }], ...extra });
}

let scope: EffectScope;
afterEach(() => {
  scope?.stop();
  idle.length = 0;
  platform.mac = false;
});
beforeEach(() => {
  scope = effectScope();
});

/** The composable over `sheet`, with an in-memory view state. */
function setup(sheet: Sheet, initial: SheetViewFields = {}) {
  const fields = shallowRef<SheetViewFields>(initial);
  const updates: SheetViewFields[] = [];
  const viewState = {
    active: computed(() => fields.value),
    updateActive(patch: SheetViewFields) {
      updates.push(patch);
      fields.value = { ...fields.value, ...patch };
    },
  };
  const sheetRef = shallowRef(sheet);
  const result = scope.run(() => {
    const sheetMetrics = useSheetMetrics(() => sheetRef.value);
    const selection = useSheetSelection({ sheet: () => sheetRef.value, sheetMetrics, viewState });
    return { selection, formulas: useFormulaDisplay({ sheet: () => sheetRef.value, model: () => null, viewState, selection }) };
  })!;
  return { selection: result.selection, formulas: result.formulas, updates, fields, sheetRef };
}

function key(init: KeyboardEventInit & { key: string }, target?: EventTarget): KeyboardEvent {
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  if (target) {
    Object.defineProperty(event, "target", { value: target });
  }
  return event;
}

const PAGE = { rows: 20, cols: 8 };
const range = (sel: SheetSelection) => sel.ranges[sel.activeRangeIndex];

describe("useSheetSelection", () => {
  it("starts at A1 when the sheet has no selection, and restores the file's selection when it has one", () => {
    expect(setup(styledSheet()).selection.selection.value.active).toEqual({ row: 0, col: 0 });
    const fileSelection: SheetSelection = {
      active: { row: 2, col: 1 },
      ranges: [
        { top: 2, left: 1, bottom: 4, right: 1 },
        { top: 6, left: 3, bottom: 7, right: 4 },
      ],
      activeRangeIndex: 0,
    };
    const { selection } = setup(styledSheet(), { selection: fileSelection });
    expect(selection.selection.value).toEqual(fileSelection);
    expect(selection.nameBoxText.value).toBe("B3");
  });

  it("selects the pressed cell, shows the drag size from press until release, and stores it as view state", () => {
    const { selection, updates } = setup(styledSheet());
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    expect(selection.selection.value.active).toEqual({ row: 2, col: 1 });
    expect(selection.nameBoxText.value).toBe("1R × 1C");
    selection.dragTo({ row: 4, col: 2 });
    // The drag reached the merged C5:E5, so the range snaps to it.
    expect(range(selection.selection.value)).toEqual({ top: 2, left: 1, bottom: 4, right: 4 });
    expect(selection.nameBoxText.value).toBe("3R × 4C");
    selection.release();
    expect(selection.nameBoxText.value).toBe("B3");
    expect(updates.at(-1)?.selection?.active).toEqual({ row: 2, col: 1 });
  });

  it("names a drag of 3 rows by 2 columns '3R × 2C'", () => {
    const { selection } = setup(styledSheet());
    selection.press({ row: 6, col: 5 }, { shift: false, mod: false });
    selection.dragTo({ row: 8, col: 6 });
    expect(selection.nameBoxText.value).toBe("3R × 2C");
  });

  it("extends with Shift+press and adds a range with Mod+press", () => {
    const { selection } = setup(styledSheet());
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    selection.release();
    selection.press({ row: 4, col: 1 }, { shift: true, mod: false });
    selection.release();
    expect(selection.selection.value.ranges).toEqual([{ top: 2, left: 1, bottom: 4, right: 1 }]);
    expect(selection.selection.value.active).toEqual({ row: 2, col: 1 });
    selection.press({ row: 9, col: 6 }, { shift: false, mod: true });
    selection.release();
    expect(selection.selection.value.ranges).toHaveLength(2);
    expect(selection.selection.value.active).toEqual({ row: 9, col: 6 });
  });

  it("makes a merge's top-left the active cell when any member is pressed", () => {
    const { selection, formulas } = setup(styledSheet());
    selection.press({ row: 4, col: 3 }, { shift: false, mod: false });
    selection.release();
    expect(selection.selection.value.active).toEqual({ row: 4, col: 2 });
    expect(selection.nameBoxText.value).toBe("C5");
    expect(formulas.barText.value).toBe("Merged title");
  });

  it("shows the active cell's General text in the formula bar, line breaks preserved", () => {
    const { selection, formulas } = setup(styledSheet());
    expect(formulas.barText.value).toBe("Line one\nLine two");
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    expect(formulas.barText.value).toBe("10");
    selection.press({ row: 50, col: 50 }, { shift: false, mod: false });
    expect(formulas.barText.value).toBe("");
  });

  it("selects whole rows and columns from the headers, and everything from the corner", () => {
    const { selection } = setup(styledSheet());
    selection.pressHeader("column", 6, { shift: false, mod: false });
    expect(range(selection.selection.value)).toEqual({ top: 0, left: 6, bottom: 1_048_575, right: 6 });
    selection.dragTo({ row: 0, col: 7 });
    expect(range(selection.selection.value)).toEqual({ top: 0, left: 6, bottom: 1_048_575, right: 7 });
    selection.release();
    // A whole column is merge-snapped like any range: C crosses C5:E5.
    selection.pressHeader("column", 2, { shift: false, mod: false });
    expect(range(selection.selection.value)).toEqual({ top: 0, left: 2, bottom: 1_048_575, right: 4 });
    selection.release();
    selection.pressHeader("row", 5, { shift: false, mod: false });
    expect(range(selection.selection.value)).toEqual({ top: 5, left: 0, bottom: 5, right: 16_383 });
    selection.selectEverything();
    expect(selection.selection.value.ranges).toEqual([{ top: 0, left: 0, bottom: 1_048_575, right: 16_383 }]);
  });

  it("extends to whole rows or columns with Shift+header press, keeping the active cell and the other ranges", () => {
    const { selection } = setup(styledSheet());
    selection.press({ row: 9, col: 6 }, { shift: false, mod: false });
    selection.release();
    selection.press({ row: 2, col: 1 }, { shift: false, mod: true });
    selection.release();
    // B3 active (Mod-added beside G10); Shift+press row 6's header.
    selection.pressHeader("row", 5, { shift: true, mod: false });
    expect(selection.selection.value.active).toEqual({ row: 2, col: 1 });
    expect(selection.selection.value.ranges).toEqual([
      { top: 9, left: 6, bottom: 9, right: 6 },
      { top: 2, left: 0, bottom: 5, right: 16_383 },
    ]);
    expect(selection.nameBoxText.value).toBe("4R × 16384C");
    // Dragging on extends from the active cell, still keeping G10.
    selection.dragTo({ row: 0, col: 3 });
    expect(range(selection.selection.value)).toEqual({ top: 0, left: 0, bottom: 2, right: 16_383 });
    selection.release();
    expect(selection.nameBoxText.value).toBe("B3");
    expect(selection.selection.value.ranges).toHaveLength(2);
    // Columns the same way; from B3 to column C crosses C5:E5, so it snaps to E.
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    selection.release();
    selection.pressHeader("column", 2, { shift: true, mod: false });
    expect(selection.selection.value.active).toEqual({ row: 2, col: 1 });
    expect(range(selection.selection.value)).toEqual({ top: 0, left: 1, bottom: 1_048_575, right: 4 });
  });

  it("navigates by the keyboard map with Ctrl as the primary modifier off macOS", () => {
    const { selection } = setup(styledSheet());
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    selection.release();
    expect(selection.handleKey(key({ key: "ArrowDown" }), PAGE)).toBe(true);
    expect(selection.nameBoxText.value).toBe("B4");
    expect(selection.handleKey(key({ key: "ArrowDown", ctrlKey: true }), PAGE)).toBe(true);
    expect(selection.nameBoxText.value).toBe("B12");
    // Cmd is not the primary modifier here: a plain move.
    expect(selection.handleKey(key({ key: "ArrowUp", metaKey: true }), PAGE)).toBe(true);
    expect(selection.nameBoxText.value).toBe("B11");
    expect(selection.handleKey(key({ key: "x" }), PAGE)).toBe(false);
  });

  it("uses Cmd as the primary modifier on macOS", () => {
    platform.mac = true;
    const { selection } = setup(styledSheet());
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    selection.release();
    expect(selection.handleKey(key({ key: "ArrowDown", metaKey: true }), PAGE)).toBe(true);
    expect(selection.nameBoxText.value).toBe("B12");
  });

  it("ignores keys whose target is an input", () => {
    const { selection } = setup(styledSheet());
    const input = document.createElement("input");
    expect(selection.handleKey(key({ key: "ArrowDown" }, input), PAGE)).toBe(false);
    expect(selection.nameBoxText.value).toBe("A1");
  });

  it("selects the current region, then the whole sheet, on Ctrl+A twice", () => {
    const { selection } = setup(styledSheet());
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    selection.release();
    selection.handleKey(key({ key: "a", ctrlKey: true }), PAGE);
    const region = range(selection.selection.value)!;
    expect(region.top).toBeLessThanOrEqual(2);
    expect(region.bottom).toBe(11);
    selection.handleKey(key({ key: "a", ctrlKey: true }), PAGE);
    expect(selection.selection.value.ranges).toEqual([{ top: 0, left: 0, bottom: 1_048_575, right: 16_383 }]);
  });

  it("aggregates a multi-cell selection: Average 20, Count 3, Sum 60 for B3:B5", async () => {
    const { selection } = setup(styledSheet());
    selection.press({ row: 2, col: 1 }, { shift: false, mod: false });
    selection.release();
    await nextTick();
    expect(selection.statusTexts.value).toEqual({});
    selection.press({ row: 4, col: 1 }, { shift: true, mod: false });
    selection.release();
    await nextTick();
    expect(selection.statusTexts.value).toEqual({ average: "Average: 20", count: "Count: 3", sum: "Sum: 60" });
  });

  it(`defers the aggregates of a selection over ${AGGREGATE_DEFER_THRESHOLD.toLocaleString("en-US")} stored cells to idle time`, async () => {
    const rows: Row[] = [];
    for (let index = 0; index < 50_001; index += 1) {
      rows.push({ index, cells: [{ col: 0, value: 1 }, { col: 1, value: 2 }, { col: 2, value: "x" }] });
    }
    const { selection } = setup(sheetOf([], { rows }));
    selection.selectEverything();
    await nextTick();
    expect(selection.aggregates.value).toBeNull();
    expect(selection.statusTexts.value).toEqual({});
    expect(idle).toHaveLength(1);
    idle.shift()!();
    expect(selection.aggregates.value).toMatchObject({ count: 150_003, numericCount: 100_002, sum: 150_003 });
    expect(selection.statusTexts.value.count).toBe("Count: 150003");
  });
});
