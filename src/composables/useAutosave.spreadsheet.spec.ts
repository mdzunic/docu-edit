// SPEC-116: autosave is a text-editor session, and a workbook id must never
// start one — beginSession bails before loading a draft or handing the
// editor a model, so nothing about a misrouted editor can be written over a
// stored workbook. The idb seam is mocked, as in useAutosave.media.spec.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h } from "vue";
import { mount } from "@vue/test-utils";
import { buildMiniBookXlsxBytes } from "../lib/spreadsheet/fixtures/build";

const putRecord = vi.fn();
const getRecord = vi.fn();
const deleteRecord = vi.fn();

beforeEach(() => {
  vi.resetModules();
  putRecord.mockReset().mockResolvedValue(undefined);
  getRecord.mockReset().mockResolvedValue(undefined);
  deleteRecord.mockReset().mockResolvedValue(undefined);
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord,
    getRecord,
    deleteRecord,
  }));
});

afterEach(() => {
  vi.doUnmock("../lib/idb");
});

describe("useAutosave on a workbook id (SPEC-116)", () => {
  it("beginSession bails: no draft is read, nothing is restored into the editor", async () => {
    const store = await import("../lib/documentStore");
    const { readSpreadsheetBytes } = await import("../lib/spreadsheet/open");
    const { useAutosave } = await import("./useAutosave");
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    store.storeSpreadsheet(model);

    const applyRestored = vi.fn();
    let api: ReturnType<typeof useAutosave> | null = null;
    mount(
      defineComponent({
        setup() {
          api = useAutosave({
            getCaret: () => null,
            getScroll: () => 0,
            isActive: () => true,
            applyRestored,
          });
          return () => h("div");
        },
      }),
    );
    await api!.beginSession(model.id);

    expect(applyRestored).not.toHaveBeenCalled();
    expect(getRecord.mock.calls.map(([key]) => key)).not.toContain(`draft:${model.id}`);
    expect(api!.recovery.value).toBeNull();
    // The stored model is still the workbook, every sheet's rows intact.
    expect(store.getSpreadsheet(model.id)?.sheets.map((sheet) => sheet.rows.length)).toEqual([11, 500, 2]);
    expect(store.getDocument(model.id)).toBeUndefined();
  });
});
