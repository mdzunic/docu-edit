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

const store = vi.hoisted(() => ({ writes: [] as SpreadsheetModel[] }));
vi.mock("../../lib/documentStore", () => ({
  storeSpreadsheetViewState: (model: SpreadsheetModel) => {
    store.writes.push(model);
    return Promise.resolve();
  },
}));

const { useSheetViewState, VIEW_STATE_PERSIST_DELAY_MS } = await import("./useSheetViewState");

function sheetOf(id: string): Sheet {
  return {
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
function setup() {
  const model = bookOf();
  let api!: ReturnType<typeof useSheetViewState>;
  const wrapper = mount(
    defineComponent({
      setup() {
        api = useSheetViewState({ workbook: shallowRef(model), activeIndex: ref(0) });
        return () => null;
      },
    }),
  );
  return { api, model, wrapper };
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
