// SPEC-103 — the deck's saving: useAutosave's state machine reused for decks
// (the header readout, drafts, recovery's Keep/Discard) composed with the
// version history controller (cadence, commit points, restore, preview, the
// over-budget degrade). The idb seam is an in-memory double of both object
// stores, so what gets written — and what never does — is observable.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h, ref, shallowRef } from "vue";
import { mount } from "@vue/test-utils";
import type { PresentationModel } from "../../lib/presentation/model";

function memoryIdb() {
  const documents = new Map<string, { id: string }>();
  const versions = new Map<string, { id: string }>();
  return {
    documents,
    versions,
    loadAllRecords: vi.fn(async () => [...documents.values()].map((r) => structuredClone(r))),
    getRecord: vi.fn(async (id: string) => structuredClone(documents.get(id))),
    putRecord: vi.fn(async (r: { id: string }) => void documents.set(r.id, structuredClone(r))),
    deleteRecord: vi.fn(async (id: string) => void documents.delete(id)),
    getVersionsRecord: vi.fn(async (id: string) => structuredClone(versions.get(id))),
    putVersionsRecord: vi.fn(async (r: { id: string }) => void versions.set(r.id, structuredClone(r))),
    deleteVersionsRecord: vi.fn(async (id: string) => void versions.delete(id)),
    loadAllTemplateRecords: vi.fn(async () => []),
    putTemplateRecord: vi.fn(async () => undefined),
    deleteTemplateRecord: vi.fn(async () => undefined),
  };
}

let idb: ReturnType<typeof memoryIdb>;

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  idb = memoryIdb();
  vi.doMock("../../lib/idb", () => idb);
});

afterEach(() => {
  vi.doUnmock("../../lib/idb");
  vi.doUnmock("../../lib/core/versionChain");
  vi.useRealTimers();
});

/** One real macrotask turn — the store yields through a MessageChannel,
 * which fake timers do not drive. */
function realTick(): Promise<void> {
  return new Promise((resolve) => {
    const channel = new MessageChannel();
    channel.port1.addEventListener(
      "message",
      () => {
        channel.port1.close();
        resolve();
      },
      { once: true },
    );
    channel.port1.start();
    channel.port2.postMessage(0);
  });
}

async function settle(turns = 20): Promise<void> {
  for (let i = 0; i < turns; i++) {
    await realTick();
  }
}

function deck(title = "Quarterly"): PresentationModel {
  return {
    kind: "presentation",
    id: "deck-1",
    title,
    format: "pptx",
    slides: Array.from({ length: 4 }, (_, i) => ({ id: `s${i}`, frames: [], partName: `ppt/slides/slide${i + 1}.xml` })),
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    opaque: [{ name: "[Content_Types].xml", data: "PFR5cGVzLz4=" }],
  };
}

/** Mounts useDeckSaving over a deck the store holds; `prepare` runs first
 * (seeding storage). */
async function harness(prepare?: (store: typeof import("../../lib/documentStore")) => void | Promise<void>) {
  const store = await import("../../lib/documentStore");
  const { useDeckSaving } = await import("./useDeckSaving");
  const { usePresentationUndo } = await import("./usePresentationUndo");
  await prepare?.(store);
  if (!store.getPresentation("deck-1")) {
    store.replacePresentationModel("deck-1", deck());
  }
  const model = shallowRef<PresentationModel | null>(store.getPresentation("deck-1")!);
  const currentIndex = ref(0);
  const order: string[] = [];
  const errors: string[] = [];
  let saving!: ReturnType<typeof useDeckSaving>;
  let undo!: ReturnType<typeof usePresentationUndo>;
  const wrapper = mount(
    defineComponent({
      setup() {
        const commit = () => {
          order.push(`commit ${model.value?.title}`);
          saving.noteEdit();
        };
        undo = usePresentationUndo({ model, onApply: commit });
        saving = useDeckSaving({
          model,
          presentationId: ref("deck-1"),
          currentIndex,
          record: (label) => {
            order.push(`record ${label}`);
            undo.record(label);
          },
          commit,
          reportError: (message) => errors.push(message),
        });
        return () => h("div");
      },
    }),
  );
  await settle();
  /** An edit, as a command makes one: record, mutate, commit. */
  function edit(mutate: (m: PresentationModel) => void) {
    undo.record("Edit");
    mutate(model.value!);
    model.value!.slides = model.value!.slides.slice();
    order.push(`commit ${model.value?.title}`);
    saving.noteEdit();
  }
  return { store, saving, model, currentIndex, order, errors, undo, wrapper, edit };
}

const versionKinds = () =>
  [...idb.versions.values()].filter((r) => r.id.startsWith("versions:")).flatMap((r) => (r as unknown as { versions: { trigger?: string }[] }).versions);

describe("the header's save state (AC-19)", () => {
  it("rests on All changes saved; an edit reads Saving… until its draft lands", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] });
    const { saving, edit } = await harness();
    expect(saving.saveState.value).toBe("saved");
    expect(saving.saveStateText.value).toBe("All changes saved");
    edit((m) => (m.title = "Edited"));
    expect(saving.saveState.value).toBe("saving");
    expect(saving.saveStateText.value).toBe("Saving…");
    await vi.advanceTimersByTimeAsync(2_500);
    await settle();
    expect(idb.documents.has("draft:deck-1")).toBe(true);
    expect(saving.saveState.value).toBe("saved");
  });

  it("a failed draft write reads the error state until a write succeeds", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] });
    const { saving, edit, store } = await harness();
    idb.putRecord.mockRejectedValue(new DOMException("quota", "QuotaExceededError"));
    edit((m) => (m.title = "Edited"));
    await vi.advanceTimersByTimeAsync(2_500);
    await settle();
    expect(saving.saveState.value).toBe("error");
    expect(saving.saveStateText.value).toBe("Not saved — browser storage is full");
    // A draft failure is the autosave's to report, not the storage banner's.
    expect(store.isStorageAvailable()).toBe(true);
    idb.putRecord.mockImplementation(async (r: { id: string }) => void idb.documents.set(r.id, structuredClone(r)));
    edit((m) => (m.title = "Again"));
    await vi.advanceTimersByTimeAsync(10_000);
    await settle();
    expect(saving.saveState.value).toBe("saved");
  });
});

describe("recovery (AC-22..AC-24)", () => {
  /** A committed deck in storage and a newer draft whose slides differ. */
  async function crashed(store: typeof import("../../lib/documentStore")) {
    const committed = deck("Committed");
    idb.documents.set("deck-1", { id: "deck-1", format: "pptx", lastModified: Date.now() - 60_000, model: { ...committed, opaque: [] }, caret: null, scroll: 0, slideIndex: 3 } as never);
    idb.documents.set("opaque:deck-1", { id: "opaque:deck-1", docId: "deck-1", opaque: committed.opaque } as never);
    const drafted = { ...deck("Drafted"), opaque: [], slides: deck().slides.slice(0, 2) };
    idb.documents.set("draft:deck-1", { id: "draft:deck-1", docId: "deck-1", savedAt: Date.now(), model: drafted, caret: null, scroll: 0, slideIndex: 1 } as never);
    await store.initLibrary();
  }

  it("reopens over a newer draft with the bar up and the drafted slides selected", async () => {
    const { saving, model, currentIndex } = await harness(crashed);
    expect(saving.recovery.value).not.toBeNull();
    expect(saving.recoveryLabel.value).toMatch(/^Recovered unsaved changes from /);
    expect(saving.saveState.value).toBe("recovered");
    expect(model.value?.title).toBe("Drafted");
    expect(model.value?.slides).toHaveLength(2);
    // The draft's package rejoined from the committed record.
    expect(model.value?.opaque).toEqual(deck().opaque);
    expect(currentIndex.value).toBe(1);
  });

  it("Keep commits the recovered deck and clears the bar", async () => {
    const { saving } = await harness(crashed);
    await saving.keepRecovered();
    await settle();
    expect(saving.recovery.value).toBeNull();
    expect((idb.documents.get("deck-1") as unknown as { model: PresentationModel }).model.title).toBe("Drafted");
    expect(idb.documents.has("draft:deck-1")).toBe(false);
  });

  it("Discard deletes the draft, restores the committed deck and re-selects its slide", async () => {
    const { saving, model, currentIndex } = await harness(crashed);
    saving.discardRecovered();
    await settle();
    expect(saving.recovery.value).toBeNull();
    expect(model.value?.title).toBe("Committed");
    expect(model.value?.slides).toHaveLength(4);
    expect(currentIndex.value).toBe(3);
    expect(idb.documents.has("draft:deck-1")).toBe(false);
  });
});

describe("snapshots (AC-34..AC-37)", () => {
  it("the cadence snapshots after five minutes of editing — and only when the deck changed", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout", "Date"] });
    const { edit } = await harness();
    await vi.advanceTimersByTimeAsync(6 * 60_000);
    await settle();
    expect(versionKinds()).toHaveLength(0);
    edit((m) => (m.title = "Edited"));
    await vi.advanceTimersByTimeAsync(5 * 60_000);
    await settle(60);
    expect(versionKinds().map((v) => v.trigger)).toEqual(["cadence"]);
  });

  it("a download is a commit point: it commits, then snapshots the changed deck", async () => {
    const { saving, edit } = await harness();
    saving.commitPoint();
    await settle(60);
    expect(versionKinds()).toHaveLength(0); // nothing changed yet
    edit((m) => (m.title = "Edited"));
    saving.commitPoint();
    await settle(60);
    expect((idb.documents.get("deck-1") as unknown as { model: PresentationModel }).model.title).toBe("Edited");
    expect(versionKinds().map((v) => v.trigger)).toEqual(["commit"]);
  });

  it("Keep version history off stops cadence and commit snapshots, never a restore's pair", async () => {
    const { saving, edit } = await harness();
    const { usePreferences } = await import("../../composables/usePreferences");
    usePreferences().setPreference("keepVersionHistory", false);
    edit((m) => (m.title = "Edited"));
    saving.commitPoint();
    await settle(60);
    expect(versionKinds()).toHaveLength(0);
    // Seed one version to restore, written directly through the library.
    const { recordDeckVersion } = await import("../../lib/presentationVersions");
    const seeded = await recordDeckVersion({ docId: "deck-1", model: deck("Seeded"), author: "You", trigger: "cadence", cachedPrev: null });
    await saving.history.restore(seeded.newEntry!.id);
    await settle(60);
    expect(versionKinds().map((v) => v.trigger)).toEqual(["cadence", "pre-restore", "post-restore"]);
  });
});

describe("restore (AC-42, AC-43)", () => {
  it("pre-restore → record('Restore version') → swap → commit() → post-restore; Ctrl+Z reverts the slides", async () => {
    const { saving, model, order, undo, edit } = await harness();
    const { recordDeckVersion } = await import("../../lib/presentationVersions");
    const old = { ...deck("Old title"), slides: deck().slides.slice(0, 2) };
    const target = await recordDeckVersion({ docId: "deck-1", model: old, author: "You", trigger: "commit", cachedPrev: null });
    edit((m) => (m.title = "Live title"));
    const liveSlides = model.value!.slides.map((s) => s.id);
    order.length = 0;

    await saving.history.restore(target.newEntry!.id);
    await settle(60);
    expect(order).toEqual(["record Restore version", "commit Old title"]);
    expect(model.value?.title).toBe("Old title");
    expect(model.value?.slides).toHaveLength(2);
    const triggers = versionKinds().map((v) => v.trigger);
    expect(triggers.slice(-2)).toEqual(["pre-restore", "post-restore"]);

    // Undo puts back what its snapshot holds — the slides — and nothing more.
    undo.undo();
    expect(model.value?.slides.map((s) => s.id)).toEqual(liveSlides);
    expect(model.value?.title).toBe("Old title");
  });
});

describe("preview (AC-40, AC-41)", () => {
  it("reconstructs the version on its own and leaves the live deck untouched", async () => {
    const { saving, model, currentIndex } = await harness();
    const { recordDeckVersion } = await import("../../lib/presentationVersions");
    const version = await recordDeckVersion({ docId: "deck-1", model: { ...deck("Then"), slides: deck().slides.slice(0, 3) }, author: "You", trigger: "commit", cachedPrev: null });
    const live = model.value;
    const liveSlides = live!.slides;
    currentIndex.value = 3;
    await saving.history.preview(version.newEntry!.id);
    expect(saving.history.previewModel.value?.title).toBe("Then");
    expect(saving.history.previewIndex.value).toBe(2); // clamped into the version
    expect(model.value).toBe(live);
    saving.history.exitPreview();
    expect(saving.history.previewModel.value).toBeNull();
    expect(model.value).toBe(live);
    expect(model.value!.slides).toBe(liveSlides);
    expect(model.value!.title).toBe("Quarterly");
  });
});

describe("the over-budget degrade (AC-46..AC-48)", () => {
  it("pauses with DECK_HISTORY_TOO_LARGE_MESSAGE and the edit still saves", async () => {
    vi.doMock("../../lib/core/versionChain", async (importOriginal) => ({
      ...(await importOriginal<typeof import("../../lib/core/versionChain")>()),
      HISTORY_BUDGET_BYTES: 10,
    }));
    const { saving, edit, store } = await harness();
    const { DECK_HISTORY_TOO_LARGE_MESSAGE } = await import("../../lib/presentationVersions");
    edit((m) => (m.title = "Edited"));
    saving.commitPoint();
    await settle(60);
    // The refusal follows the package's digests, which crypto.subtle runs off
    // the event loop: under the full suite's load they can outlast any fixed
    // count of turns, so wait on the latch itself.
    await vi.waitFor(() => expect(saving.history.pausedMessage.value).toBe(DECK_HISTORY_TOO_LARGE_MESSAGE));
    expect(saving.history.visibleVersions.value).toHaveLength(0);
    // Nothing but the index latch reached the versions store …
    expect([...idb.versions.keys()]).toEqual(["versions:deck-1"]);
    // … and the deck itself committed, with no storage banner.
    expect((idb.documents.get("deck-1") as unknown as { model: PresentationModel }).model.title).toBe("Edited");
    expect(store.isStorageAvailable()).toBe(true);
  });

  it("a history write failure never latches the banner nor touches the deck's records", async () => {
    idb.putVersionsRecord.mockRejectedValue(new DOMException("quota", "QuotaExceededError"));
    const { saving, edit, store } = await harness();
    edit((m) => (m.title = "Edited"));
    const before = idb.putRecord.mock.calls.length;
    saving.commitPoint();
    await settle(60);
    expect(store.isStorageAvailable()).toBe(true);
    expect(saving.history.visibleVersions.value).toHaveLength(0);
    // The commit's own writes only: the deck's shell (and its package, the
    // first time) — never a history record in the documents store.
    const written = idb.putRecord.mock.calls.slice(before).map(([r]) => (r as { id: string }).id);
    expect(written.every((id) => id === "deck-1" || id === "opaque:deck-1")).toBe(true);
    expect(saving.saveState.value).toBe("saved");
  });
});

describe("the panel's open state", () => {
  it("is the presentationHistoryOpen preference", async () => {
    const { saving } = await harness();
    const { loadPreferences } = await import("../../lib/preferences");
    expect(saving.history.open.value).toBe(false);
    saving.history.toggle();
    expect(saving.history.open.value).toBe(true);
    expect(loadPreferences().presentationHistoryOpen).toBe(true);
  });
});
