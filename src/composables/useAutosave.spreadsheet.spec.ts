// SPEC-116 made beginSession bail on a workbook id, because the draft
// machinery could not yet tell a workbook from a document; SPEC-138 lifted
// that bail — a workbook session reads its draft and hands the view the whole
// workbook, every sheet's rows intact, and a text model can still never be
// swapped in over it. The idb seam is mocked, as in useAutosave.media.spec.ts.
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

describe("useAutosave on a workbook id (SPEC-116, SPEC-138)", () => {
  it("beginSession runs a workbook session: the draft is read and the whole workbook is restored", async () => {
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

    expect(getRecord.mock.calls.map(([key]) => key)).toContain(`draft:${model.id}`);
    expect(applyRestored).toHaveBeenCalledTimes(1);
    expect(applyRestored.mock.calls[0]![0]).toMatchObject({ model, recovered: false });
    expect(api!.recovery.value).toBeNull();
    // The stored model is still the workbook, every sheet's rows intact.
    expect(store.getSpreadsheet(model.id)?.sheets.map((sheet) => sheet.rows.length)).toEqual([11, 500, 2]);
    expect(store.getDocument(model.id)).toBeUndefined();
    expect(() => store.replaceDocumentModel(model.id, { id: model.id } as never)).toThrow(/spreadsheet/);
  });
});
