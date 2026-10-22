// SPEC-107: autosave is a text-editor session, and a media id must never
// start one — beginSession bails before loading a draft or handing the
// editor a model, so nothing about a misrouted editor can be written over
// the stored bytes. The idb seam is mocked, as in useAutosave.flush.spec.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h } from "vue";
import { mount } from "@vue/test-utils";
import { buildPngFixtureBytes } from "../lib/media/fixtures/build";

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

describe("useAutosave on a media id (SPEC-107)", () => {
  it("beginSession bails: no draft is read, nothing is restored into the editor", async () => {
    const store = await import("../lib/documentStore");
    const { assembleMediaModel } = await import("../lib/media/open");
    const { useAutosave } = await import("./useAutosave");
    const model = await assembleMediaModel("Shot", "png", await buildPngFixtureBytes());
    store.storeMedia(model);

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
    // The draft lookup (draft:<id>) never happens; only the committed
    // record's refresh does.
    expect(getRecord.mock.calls.map(([key]) => key)).not.toContain(`draft:${model.id}`);
    expect(api!.recovery.value).toBeNull();
    // The stored model is still the media document, payload intact.
    expect(store.getMedia(model.id)?.media[0].data).toBe(model.media[0].data);
    expect(store.getDocument(model.id)).toBeUndefined();
  });
});
