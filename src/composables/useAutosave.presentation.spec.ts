// SPEC-103: useAutosave serves decks too. A deck's session records its
// selected slide beside every draft and commit, a newer draft reopens as
// recovered slides through replacePresentationModel, and Keep/Discard behave
// exactly as the text side's — the same state machine, unchanged. The idb
// seam is mocked, as in useAutosave.flush.spec.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h } from "vue";
import { mount } from "@vue/test-utils";
import type { PresentationModel } from "../lib/presentation/model";
import type { RestoredState } from "./useAutosave";

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

function deck(id: string, title: string, slides = 3): PresentationModel {
  return {
    kind: "presentation",
    id,
    title,
    format: "pptx",
    slides: Array.from({ length: slides }, (_, i) => ({ id: `s${i}`, frames: [] })),
    slideWidthCm: 25.4,
    slideHeightCm: 19.05,
  };
}

async function setup(id: string, options: { draft?: PresentationModel; draftSlide?: number } = {}) {
  const store = await import("../lib/documentStore");
  const { useAutosave } = await import("./useAutosave");
  const committed = store.replacePresentationModel(id, deck(id, "Committed"));
  store.getStoredRecord(id)!.slideIndex = 2;
  if (options.draft) {
    getRecord.mockImplementation(async (key: string) =>
      key === `draft:${id}`
        ? { id: key, docId: id, savedAt: Date.now() + 60_000, model: options.draft, caret: null, scroll: 0, slideIndex: options.draftSlide }
        : undefined,
    );
  }
  let slideIndex = 1;
  const restored: RestoredState[] = [];
  let api: ReturnType<typeof useAutosave> | null = null;
  mount(
    defineComponent({
      setup() {
        api = useAutosave({
          getCaret: () => null,
          getScroll: () => 0,
          getSlideIndex: () => slideIndex,
          isActive: () => true,
          applyRestored: (state) => restored.push(state),
        });
        return () => h("div");
      },
    }),
  );
  await api!.beginSession(id);
  return { api: api!, store, committed, restored, setSlide: (index: number) => (slideIndex = index) };
}

function writes(key: string) {
  return putRecord.mock.calls.map(([record]) => record).filter((record) => record.id === key);
}

describe("useAutosave on a deck (SPEC-103)", () => {
  it("opens a deck with no newer draft as itself, handing back the committed slide", async () => {
    const { api, restored, committed } = await setup("deck-open");
    expect(api.recovery.value).toBeNull();
    expect(restored).toHaveLength(1);
    expect(restored[0].model).toBe(committed);
    expect(restored[0].slideIndex).toBe(2);
    expect(restored[0].recovered).toBe(false);
  });

  it("records the deck's selected slide beside every draft", async () => {
    vi.useFakeTimers();
    try {
      const { api, setSlide } = await setup("deck-draft");
      setSlide(2);
      api.markDirty();
      expect(api.saveState.value).toBe("saving");
      await vi.advanceTimersByTimeAsync(2_500);
      await vi.waitFor(() => expect(writes("draft:deck-draft")).toHaveLength(1));
      expect(writes("draft:deck-draft")[0]).toMatchObject({ caret: null, scroll: 0, slideIndex: 2 });
    } finally {
      vi.useRealTimers();
    }
  });

  it("reopens over a newer draft as recovered slides, routed through replacePresentationModel", async () => {
    const draft = deck("deck-rec", "Drafted", 2);
    const { api, store, restored } = await setup("deck-rec", { draft, draftSlide: 1 });
    expect(api.recovery.value).not.toBeNull();
    expect(api.saveStateText.value).toBe("Not saved yet");
    expect(restored.at(-1)).toMatchObject({ recovered: true, slideIndex: 1 });
    expect(restored.at(-1)!.model).toBe(draft);
    expect(store.getPresentation("deck-rec")).toBe(draft);
  });

  it("Keep commits the recovered slides with the current slide and clears the bar", async () => {
    const draft = deck("deck-keep", "Drafted", 2);
    const { api, store, setSlide } = await setup("deck-keep", { draft, draftSlide: 1 });
    setSlide(1);
    putRecord.mockClear();
    await api.keepRecovered();
    expect(api.recovery.value).toBeNull();
    expect(api.saveState.value).toBe("saved");
    const committedWrite = writes("deck-keep").at(-1);
    expect(committedWrite.model.title).toBe("Drafted");
    expect(committedWrite.slideIndex).toBe(1);
    expect(store.getStoredRecord("deck-keep")?.slideIndex).toBe(1);
    expect(deleteRecord).toHaveBeenCalledWith("draft:deck-keep");
  });

  it("Discard deletes the draft and restores the committed deck and its slide", async () => {
    const draft = deck("deck-discard", "Drafted", 2);
    const { api, store, restored, committed } = await setup("deck-discard", { draft, draftSlide: 1 });
    api.discardRecovered();
    expect(api.recovery.value).toBeNull();
    expect(restored.at(-1)).toMatchObject({ recovered: false, slideIndex: 2 });
    expect(restored.at(-1)!.model).toBe(committed);
    expect(store.getPresentation("deck-discard")).toBe(committed);
    await vi.waitFor(() => expect(deleteRecord).toHaveBeenCalledWith("draft:deck-discard"));
  });
});
