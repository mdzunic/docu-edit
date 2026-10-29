// SPEC-117 "View state" (AC-46, AC-47): a change to a sheet's view fields is
// written back through storeSpreadsheetViewState at most once per
// VIEW_STATE_PERSIST_DELAY_MS, on a new model whose changed sheet is a new
// object keeping its `rows` — and flushed on close and when the page goes
// away. The page-leave flush starts at `beforeunload` (a write begun only at
// `pagehide` does not land before a reload), with `pagehide` as the fallback.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, ref, shallowRef } from "vue";
import { mount } from "@vue/test-utils";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";

const store = vi.hoisted(() => ({ writes: [] as SpreadsheetModel[] }));
vi.mock("../../lib/documentStore", () => ({
  storeSpreadsheetViewState: (model: SpreadsheetModel) => {
    store.writes.push(model);
    return Promise.resolve();
  },
}));

const { useSheetViewState, withSheetView, VIEW_STATE_PERSIST_DELAY_MS } = await import("./useSheetViewState");

function sheetOf(id: string): Sheet {
  return {
    conditionalFormats: [],
    validations: [],
    id,
    name: id,
    visibility: "visible",
    kind: "grid",
    rows: [{ index: 0, cells: [{ col: 0, value: 1 }] }],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
  };
}

function bookOf(): SpreadsheetModel {
  return {
    names: [],
    dxfs: [],
    kind: "spreadsheet",
    id: "book",
    title: "Book",
    format: "xlsx",
    readerVersion: 2,
    sheets: [sheetOf("s1"), sheetOf("s2")],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
  };
}

/** The composable over a two-sheet book, mounted so it can unmount. */
function setup(model = bookOf()) {
  const activeIndex = ref(model.activeSheetIndex);
  let api!: ReturnType<typeof useSheetViewState>;
  const wrapper = mount(
    defineComponent({
      setup() {
        api = useSheetViewState({ workbook: shallowRef(model), activeIndex });
        return () => null;
      },
    }),
  );
  return { api, model, activeIndex, wrapper };
}

let unmount: (() => void) | null = null;
beforeEach(() => {
  vi.useFakeTimers();
  store.writes.length = 0;
});
afterEach(() => {
  unmount?.();
  unmount = null;
  vi.useRealTimers();
});

describe("useSheetViewState", () => {
  it("writes a change once the delay has passed, on a new sheet object that keeps its rows", () => {
    const { api, model, wrapper } = setup();
    unmount = () => wrapper.unmount();
    api.updateActive({ showGridLines: false });
    vi.advanceTimersByTime(VIEW_STATE_PERSIST_DELAY_MS - 1);
    expect(store.writes).toEqual([]);
    vi.advanceTimersByTime(1);
    expect(store.writes).toHaveLength(1);
    const [written] = store.writes;
    expect(written!.sheets[0]!.showGridLines).toBe(false);
    expect(written!.sheets[0]).not.toBe(model.sheets[0]);
    expect(written!.sheets[0]!.rows).toBe(model.sheets[0]!.rows);
    expect(written!.sheets[1]).toBe(model.sheets[1]);
    expect(model.sheets[0]!.showGridLines).toBeUndefined();
  });

  it("writes at most once per delay however many changes land inside it", () => {
    const { api, wrapper } = setup();
    unmount = () => wrapper.unmount();
    api.updateActive({ showGridLines: false });
    api.updateActive({ zoomPercent: 150 });
    api.updateActive({ topLeftCell: { row: 40, col: 2 } });
    vi.advanceTimersByTime(VIEW_STATE_PERSIST_DELAY_MS);
    expect(store.writes).toHaveLength(1);
    expect(store.writes[0]!.sheets[0]).toMatchObject({ showGridLines: false, zoomPercent: 150, topLeftCell: { row: 40, col: 2 } });
  });

  it("flushes a pending change at beforeunload, so a reload inside the delay keeps it", () => {
    const { api, wrapper } = setup();
    unmount = () => wrapper.unmount();
    api.updateActive({ showGridLines: false });
    window.dispatchEvent(new Event("beforeunload"));
    expect(store.writes).toHaveLength(1);
    expect(store.writes[0]!.sheets[0]!.showGridLines).toBe(false);
    // Nothing is left for pagehide or the timer to write.
    window.dispatchEvent(new Event("pagehide"));
    vi.advanceTimersByTime(VIEW_STATE_PERSIST_DELAY_MS);
    expect(store.writes).toHaveLength(1);
  });

  it("flushes a pending change at pagehide where beforeunload never fired", () => {
    const { api, wrapper } = setup();
    unmount = () => wrapper.unmount();
    api.updateActive({ zoomPercent: 85 });
    window.dispatchEvent(new Event("pagehide"));
    expect(store.writes.map((model) => model.sheets[0]!.zoomPercent)).toEqual([85]);
  });

  it("flushes on close, and a closed view no longer flushes on a page leave", () => {
    const { api, wrapper } = setup();
    api.updateActive({ showGridLines: false });
    wrapper.unmount();
    expect(store.writes).toHaveLength(1);
    // Were its listeners still attached, this would be written now.
    api.updateActive({ zoomPercent: 150 });
    window.dispatchEvent(new Event("beforeunload"));
    window.dispatchEvent(new Event("pagehide"));
    expect(store.writes).toHaveLength(1);
  });

  it("writes nothing on a page leave when nothing changed", () => {
    const { wrapper } = setup();
    unmount = () => wrapper.unmount();
    window.dispatchEvent(new Event("beforeunload"));
    window.dispatchEvent(new Event("pagehide"));
    expect(store.writes).toEqual([]);
  });
});

// SPEC-138: a recorded view (a commit's or a draft's sheetView: the tab, the
// active cell, the top-left cell) changes only what disagrees with the view
// state the model carries — so a plain reopen, whose model carries what its
// commit recorded, changes nothing and re-mounts nothing.
describe("a recorded view (SPEC-138's sheetView)", () => {
  const WHOLE_SHEET: SheetSelection = { active: { row: 2, col: 1 }, ranges: [{ top: 0, left: 0, bottom: 1_048_575, right: 16_383 }], activeRangeIndex: 0 };

  function bookWithSelection(): SpreadsheetModel {
    const model = bookOf();
    model.sheets[0] = { ...model.sheets[0]!, selection: WHOLE_SHEET, topLeftCell: { row: 4, col: 0 } };
    return model;
  }

  it("a model that already carries it changes nothing: a whole-sheet selection on its active cell survives, where the user moved since stays", () => {
    const { api, model, activeIndex, wrapper } = setup(bookWithSelection());
    unmount = () => wrapper.unmount();
    const sheets = model.sheets;
    api.updateActive({ topLeftCell: { row: 40, col: 0 } });
    expect(api.applySheetView({ sheetId: "s1", active: { row: 2, col: 1 }, topRow: 4, leftCol: 0 })).toBe(false);
    expect(model.sheets).toBe(sheets);
    expect(model.sheets[0]!.selection).toBe(WHOLE_SHEET);
    expect(activeIndex.value).toBe(0);
    expect(api.active.value.selection).toBe(WHOLE_SHEET);
    expect(api.active.value.topLeftCell).toEqual({ row: 40, col: 0 });
    // An absent selection and top-left are A1's.
    const plain = setup();
    expect(plain.api.applySheetView({ sheetId: "s1", active: { row: 0, col: 0 }, topRow: 0, leftCol: 0 })).toBe(false);
    expect(plain.model.sheets[0]!.selection).toBeUndefined();
    plain.wrapper.unmount();
  });

  it("keeps a selection whose active cell is the view's while the top-left moves to the view's", () => {
    const model = bookWithSelection();
    expect(withSheetView(model, { sheetId: "s1", active: { row: 2, col: 1 }, topRow: 9, leftCol: 3 })).toBe(true);
    expect(model.sheets[0]!.selection).toBe(WHOLE_SHEET);
    expect(model.sheets[0]!.topLeftCell).toEqual({ row: 9, col: 3 });
  });

  it("a view that disagrees is applied: its sheet shown, the selection collapsed to its active cell, its top-left cell", () => {
    const { api, model, activeIndex, wrapper } = setup(bookWithSelection());
    unmount = () => wrapper.unmount();
    const rows = model.sheets[1]!.rows;
    expect(api.applySheetView({ sheetId: "s2", active: { row: 3, col: 2 }, topRow: 2, leftCol: 1 })).toBe(true);
    expect(model.activeSheetIndex).toBe(1);
    expect(activeIndex.value).toBe(1);
    const collapsed = { active: { row: 3, col: 2 }, ranges: [{ top: 3, left: 2, bottom: 3, right: 2 }], activeRangeIndex: 0 };
    expect(model.sheets[1]!.selection).toEqual(collapsed);
    expect(model.sheets[1]!.topLeftCell).toEqual({ row: 2, col: 1 });
    expect(model.sheets[1]!.rows).toBe(rows);
    expect(api.active.value).toEqual({ selection: collapsed, topLeftCell: { row: 2, col: 1 } });
    // On the same sheet, another active cell collapses the whole-sheet selection.
    const other = bookWithSelection();
    expect(withSheetView(other, { sheetId: "s1", active: { row: 0, col: 0 }, topRow: 4, leftCol: 0 })).toBe(true);
    expect(other.sheets[0]!.selection).toEqual({ active: { row: 0, col: 0 }, ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], activeRangeIndex: 0 });
    expect(other.sheets[0]!.topLeftCell).toEqual({ row: 4, col: 0 });
  });
});
