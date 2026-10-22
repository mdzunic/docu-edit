// SPEC-137 in the workbook view: a stored workbook that is not resident is
// read back on open — the view shows its loading state meanwhile, never the
// not-found one — and released when the view goes; a load that came back
// incomplete shows WORKBOOK_STORAGE_INCOMPLETE_MESSAGE and the workbook is
// read-only: the formula bar is not editable, and the edit pipeline undoes
// any command at its commit, writing nothing through.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { effectScope, shallowRef } from "vue";
import { createMemoryHistory, createRouter } from "vue-router";
import type { WriteBatchOp } from "../lib/core/idb";
import type { Row, SpreadsheetModel } from "../lib/spreadsheet/model";

const data = new Map<string, unknown>();
const getGate: { wait: Promise<void> | null } = { wait: null };
const writeBatch = vi.fn(async (ops: WriteBatchOp[]) => {
  for (const op of ops) {
    if ("put" in op) {
      data.set((op.put as { id: string }).id, structuredClone(op.put));
    } else if ("delete" in op) {
      data.delete(op.delete);
    } else {
      for (const key of [...data.keys()]) {
        if (key >= op.deleteRange[0] && key <= op.deleteRange[1]) {
          data.delete(key);
        }
      }
    }
  }
});
const cleanups: Array<() => void> = [];

beforeEach(() => {
  vi.resetModules();
  data.clear();
  getGate.wait = null;
  writeBatch.mockClear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    getRecord: vi.fn(async (key: string) => {
      await getGate.wait;
      return data.has(key) ? structuredClone(data.get(key)) : undefined;
    }),
    writeBatch,
  }));
  vi.spyOn(Element.prototype, "clientWidth", "get").mockReturnValue(1280);
  vi.spyOn(Element.prototype, "clientHeight", "get").mockReturnValue(720);
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});

afterEach(() => {
  for (const cleanup of cleanups.splice(0)) {
    cleanup();
  }
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
});

function rowsOf(count: number): Row[] {
  return Array.from({ length: count }, (_, index) => ({ index, cells: [{ col: 0, value: `v${index}` }] }));
}

/** A 1,100-row workbook stored and released: in storage, not in memory. */
async function storedWorkbook() {
  const store = await import("../lib/documentStore");
  const { createBlankWorkbook } = await import("../lib/spreadsheet/sheets");
  const model = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows: rowsOf(1_100) }];
  store.storeSpreadsheet(model);
  await store.releaseSpreadsheet(model.id);
  expect(store.getSpreadsheet(model.id)).toBeUndefined();
  return { store, model };
}

async function mountSheets(id: string) {
  const { routes } = await import("../router");
  const { default: SpreadsheetView } = await import("./SpreadsheetView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/sheets/${id}`);
  const wrapper = mount(SpreadsheetView, { global: { plugins: [router] }, attachTo: document.body });
  cleanups.push(() => wrapper.unmount());
  return wrapper;
}

describe("the workbook view over lazy residency", () => {
  it("shows the loading state (not not-found) while the workbook loads, then the grid; releases it on unmount", async () => {
    const { store, model } = await storedWorkbook();
    let open = (): void => {};
    getGate.wait = new Promise((resolve) => {
      open = resolve;
    });
    const wrapper = await mountSheets(model.id);
    await flushPromises();
    expect(wrapper.find('[data-testid="spreadsheet-loading"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="spreadsheet-not-found"]').exists()).toBe(false);

    open();
    await vi.waitFor(() => expect(wrapper.find('[data-testid="spreadsheet-grid"]').exists()).toBe(true), { timeout: 5_000 });
    expect(wrapper.get('[data-ref="A1"]').text()).toBe("v0");
    expect(store.getSpreadsheet(model.id)?.sheets[0]!.rows).toHaveLength(1_100);

    wrapper.unmount();
    cleanups.length = 0;
    await vi.waitFor(() => expect(store.getSpreadsheet(model.id)).toBeUndefined());
    expect(store.getSpreadsheetShell(model.id)?.sheets[0]!.rows).toEqual([]);
  });

  it("incomplete storage: the alert shows, the formula bar is not editable, and nothing is written back", async () => {
    const { model } = await storedWorkbook();
    data.delete(`sheetrows:${model.id}:${model.sheets[0]!.id}:0001`);
    writeBatch.mockClear();
    const { WORKBOOK_STORAGE_INCOMPLETE_MESSAGE } = await import("../lib/spreadsheet/errors");
    const wrapper = await mountSheets(model.id);
    await vi.waitFor(() => expect(wrapper.find('[data-testid="workbook-storage-incomplete"]').exists()).toBe(true), { timeout: 5_000 });
    const alert = wrapper.get('[data-testid="workbook-storage-incomplete"]');
    expect(alert.attributes("role")).toBe("alert");
    expect(alert.text()).toBe(WORKBOOK_STORAGE_INCOMPLETE_MESSAGE);
    expect(wrapper.find(".formula-bar-content.is-editable").exists()).toBe(false);
    await flushPromises();
    expect(writeBatch).not.toHaveBeenCalled();
  });

  it("the edit pipeline undoes a command at its commit while read-only: nothing stacked, nothing written through", async () => {
    const { useSpreadsheetUndo } = await import("./spreadsheet/useSpreadsheetUndo");
    const { createBlankWorkbook } = await import("../lib/spreadsheet/sheets");
    const live = createBlankWorkbook();
    const before = live.sheets[0]!.rows;
    const onApplied = vi.fn();
    const scope = effectScope();
    scope.run(() => {
      const model = shallowRef<SpreadsheetModel | null>(live);
      const undo = useSpreadsheetUndo({ model, recalc: (_keys, done) => done([]), onApplied, activateSheet: () => {}, readOnly: () => true });
      undo.record("Typing", { sheetId: live.sheets[0]!.id });
      live.sheets = [{ ...live.sheets[0]!, rows: rowsOf(1) }];
      undo.commit();
      expect(live.sheets[0]!.rows).toBe(before);
      expect(undo.canUndo.value).toBe(false);
    });
    scope.stop();
    expect(onApplied).not.toHaveBeenCalled();
  });
});
