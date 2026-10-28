import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import type { DraftRecord, OpaqueRecord, PersistedRecord, StoredDocumentRecord } from "./documentStore";
import type { PresentationModel } from "./presentation/model";

// SPEC-010 — the committed/draft split. Fresh modules per test so the store's
// in-memory map and latched storage status start clean; the idb seam is
// mocked to observe writes and steer failures.
async function loadStore(overrides: Record<string, unknown> = {}) {
  const putRecord = vi.fn().mockResolvedValue(undefined);
  const getRecord = vi.fn().mockResolvedValue(undefined);
  const deleteRecord = vi.fn().mockResolvedValue(undefined);
  const loadAllRecords = vi.fn().mockResolvedValue([]);
  vi.doMock("./idb", () => ({
    putRecord,
    getRecord,
    deleteRecord,
    loadAllRecords,
    ...overrides,
  }));
  const store = await import("./documentStore");
  return { store, putRecord, getRecord, deleteRecord, loadAllRecords };
}

// documentStore's write-through paths now yield a real macrotask (via
// idleScheduler's MessageChannel-based yieldToMainThread, SPEC-047 AC-15/
// AC-16) before their putRecord call — a getOrCreateDocument()'s
// fire-and-forget persist() is still in flight when the synchronous test
// body continues, so a test that reconfigures the mocked putRecord right
// after must wait for that write to LAND first. Otherwise the deferred call
// picks up the reconfigured mock instead of the one active when it was
// issued, and because persist() is the SPEC-003 write-through it catches
// into markStorageUnavailable() — a rejection staged for the autosave path
// latches the banner the test is asserting stays clear.
//
// It waits on the write itself, not on one setTimeout(0) tick: the pending
// write resumes from a MessageChannel message task while the tick is a timer
// task, and the two task sources have no ordering guarantee between them (an
// already-expired timer runs in Node's timers phase, before the poll phase
// that drains the MessagePort). A single tick therefore raced the write and
// lost about once per 200 runs of this file. The closing expect keeps a
// barrier that stops working loud instead of silent.
//
// One putRecord call is the whole of a fresh document's persist(): a
// createDocument() model has no `media`, so persistMedia() skips its write
// and only the shell record goes out.
async function flushSetupWrite(putRecord: {
  mock: { calls: unknown[] };
}): Promise<void> {
  for (let tick = 0; tick < 100 && putRecord.mock.calls.length === 0; tick += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  expect(
    putRecord.mock.calls.length,
    "setup persist() never reached putRecord — the barrier below is not holding",
  ).toBeGreaterThan(0);
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.doUnmock("./idb");
});

describe("saveDraft", () => {
  it("writes a draft record keyed under the draft prefix", async () => {
    const { store, putRecord } = await loadStore();
    store.getOrCreateDocument("doc-1");
    await store.saveDraft("doc-1", { caret: { section: 0, block: 0, offset: 3 }, scroll: 120 });
    const draft = putRecord.mock.calls.at(-1)![0];
    expect(draft.id).toBe("draft:doc-1");
    expect(draft.docId).toBe("doc-1");
    expect(draft.model.id).toBe("doc-1");
    expect(draft.caret).toEqual({ section: 0, block: 0, offset: 3 });
    expect(draft.scroll).toBe(120);
    expect(typeof draft.savedAt).toBe("number");
  });

  it("writes changed media under the DRAFT media key, never the committed one (SPEC-047 AC-15)", async () => {
    const { store, putRecord } = await loadStore();
    const doc = store.getOrCreateDocument("doc-m");
    doc.media = [{ mediaId: "m1", contentType: "image/png", data: "aGk=" }];
    // Commit so the committed media latch reflects m1's array.
    await store.commitDocument("doc-m", { caret: null, scroll: 0 });
    putRecord.mockClear();

    // An edit replaces the media array (removing the image), then autosaves.
    doc.media = [];
    await store.saveDraft("doc-m", { caret: null, scroll: 0 });

    const written = putRecord.mock.calls.map((call) => call[0] as PersistedRecord);
    expect(written.some((record) => record.id === "draftmedia:doc-m")).toBe(true);
    // The committed media record must be untouched: "Discard" depends on it.
    expect(written.some((record) => record.id === "media:doc-m")).toBe(false);
  });

  it("skips the draft media write entirely while media matches the committed state", async () => {
    const { store, putRecord } = await loadStore();
    const doc = store.getOrCreateDocument("doc-s");
    doc.media = [{ mediaId: "m1", contentType: "image/png", data: "aGk=" }];
    await store.commitDocument("doc-s", { caret: null, scroll: 0 });
    putRecord.mockClear();

    // No media change between commit and autosave: no media write at all.
    await store.saveDraft("doc-s", { caret: null, scroll: 0 });
    const written = putRecord.mock.calls.map((call) => call[0] as PersistedRecord);
    expect(written.every((record) => !record.id.includes("media"))).toBe(true);
  });

  it("retries the committed media write after a failed one instead of latching it as persisted", async () => {
    const { store, putRecord } = await loadStore();
    const doc = store.getOrCreateDocument("doc-r");
    await flushSetupWrite(putRecord);
    doc.media = [{ mediaId: "m1", contentType: "image/png", data: "aGk=" }];
    putRecord.mockClear();
    putRecord.mockRejectedValueOnce(new DOMException("quota", "QuotaExceededError"));

    // First commit: the media write fails (the shell write may still land).
    await expect(store.commitDocument("doc-r", { caret: null, scroll: 0 })).rejects.toBeTruthy();
    putRecord.mockResolvedValue(undefined);

    // Second commit with the SAME media reference must retry the media
    // write — a failure latched as "already persisted" would strand the
    // images outside storage forever.
    await store.commitDocument("doc-r", { caret: null, scroll: 0 });
    const written = putRecord.mock.calls.map((call) => call[0] as PersistedRecord);
    expect(written.some((record) => record.id === "media:doc-r")).toBe(true);
  });

  it("rejects on write failure WITHOUT latching the SPEC-003 banner", async () => {
    const { store, putRecord } = await loadStore();
    store.getOrCreateDocument("doc-2");
    await flushSetupWrite(putRecord);
    putRecord.mockRejectedValue(
      new DOMException("quota", "QuotaExceededError"),
    );
    await expect(store.saveDraft("doc-2", { caret: null, scroll: 0 })).rejects.toBeTruthy();
    expect(store.isStorageAvailable()).toBe(true);
  });

  it("rejects for an unknown document id", async () => {
    const { store } = await loadStore();
    await expect(store.saveDraft("nope", { caret: null, scroll: 0 })).rejects.toBeTruthy();
  });
});

describe("loadDraft", () => {
  it("returns the stored draft", async () => {
    const { store, getRecord } = await loadStore();
    const model = { id: "doc-3" };
    getRecord.mockResolvedValue({
      id: "draft:doc-3",
      docId: "doc-3",
      savedAt: 42,
      model,
      caret: null,
      scroll: 0,
    });
    const draft = await store.loadDraft("doc-3");
    expect(getRecord).toHaveBeenCalledWith("draft:doc-3");
    expect(draft?.savedAt).toBe(42);
  });

  it("resolves undefined when absent or unreadable", async () => {
    const { store, getRecord } = await loadStore();
    expect(await store.loadDraft("doc-4")).toBeUndefined();
    getRecord.mockRejectedValue(new Error("no idb"));
    expect(await store.loadDraft("doc-4")).toBeUndefined();
  });
});

describe("commitDocument", () => {
  it("bumps lastModified, records caret and scroll, and clears the draft", async () => {
    const { store, putRecord, deleteRecord } = await loadStore();
    store.getOrCreateDocument("doc-5");
    const before = store.getStoredRecord("doc-5")!.lastModified;
    await new Promise((resolve) => setTimeout(resolve, 2));
    await store.commitDocument("doc-5", { caret: { section: 0, block: 1, offset: 7 }, scroll: 33 });
    const record = store.getStoredRecord("doc-5")!;
    expect(record.lastModified).toBeGreaterThan(before);
    expect(record.caret).toEqual({ section: 0, block: 1, offset: 7 });
    expect(record.scroll).toBe(33);
    // Not .toBe(record): commitDocument now writes a media-stripped shell
    // (SPEC-047 AC-15 — putRecord()'s structured clone must not carry the
    // in-memory model's `media` array), so the object identity changes even
    // though every field but `model.media` matches the in-memory record.
    expect(putRecord.mock.calls.at(-1)![0]).toEqual({ ...record, model: { ...record.model, media: [] } });
    expect(deleteRecord).toHaveBeenCalledWith("draft:doc-5");
  });

  it("rejects when the committed write fails, without latching the banner", async () => {
    const { store, putRecord } = await loadStore();
    store.getOrCreateDocument("doc-6");
    await flushSetupWrite(putRecord);
    putRecord.mockRejectedValue(new Error("full"));
    await expect(store.commitDocument("doc-6", { caret: null, scroll: 0 })).rejects.toBeTruthy();
    expect(store.isStorageAvailable()).toBe(true);
  });
});

describe("initLibrary", () => {
  it("hydrates committed records but never drafts", async () => {
    const committed: PersistedRecord = {
      id: "doc-7",
      format: "new",
      lastModified: 1,
      model: { id: "doc-7" } as never,
    };
    const draft: PersistedRecord = {
      id: "draft:doc-7",
      docId: "doc-7",
      savedAt: 2,
      model: { id: "doc-7" } as never,
      caret: null,
      scroll: 0,
    };
    const { store } = await loadStore({
      loadAllRecords: vi.fn().mockResolvedValue([committed, draft]),
    });
    await store.initLibrary();
    expect(store.getStoredRecord("doc-7")).toBe(committed);
    expect(store.getStoredRecord("draft:doc-7")).toBeUndefined();
    expect(store.listDocuments()).toHaveLength(1);
  });
});

describe("deleteDocument", () => {
  it("removes the draft along with the committed record", async () => {
    const { store, deleteRecord } = await loadStore();
    store.getOrCreateDocument("doc-8");
    store.deleteDocument("doc-8");
    expect(deleteRecord).toHaveBeenCalledWith("doc-8");
    expect(deleteRecord).toHaveBeenCalledWith("draft:doc-8");
  });
});

describe("refreshDocumentFromStorage", () => {
  it("adopts a stored record newer than the in-memory one", async () => {
    const { store, getRecord } = await loadStore();
    store.getOrCreateDocument("doc-9");
    const newer = {
      id: "doc-9",
      format: "new",
      lastModified: Date.now() + 60_000,
      model: { id: "doc-9" },
    };
    getRecord.mockResolvedValue(newer);
    const record = await store.refreshDocumentFromStorage("doc-9");
    expect(record).toBe(newer);
    expect(store.getStoredRecord("doc-9")).toBe(newer);
  });

  it("keeps the in-memory record when storage is stale or unreadable", async () => {
    const { store, getRecord } = await loadStore();
    store.getOrCreateDocument("doc-10");
    const memory = store.getStoredRecord("doc-10");
    getRecord.mockResolvedValue({
      id: "doc-10",
      format: "new",
      lastModified: 0,
      model: { id: "doc-10" },
    });
    expect(await store.refreshDocumentFromStorage("doc-10")).toBe(memory);
    getRecord.mockRejectedValue(new Error("no idb"));
    expect(await store.refreshDocumentFromStorage("doc-10")).toBe(memory);
  });
});

describe("replaceDocumentModel", () => {
  it("swaps the model without bumping lastModified or persisting", async () => {
    const { store, putRecord } = await loadStore();
    store.getOrCreateDocument("doc-11");
    const record = store.getStoredRecord("doc-11")!;
    const stamp = record.lastModified;
    const writes = putRecord.mock.calls.length;
    const replacement = { id: "other" } as never;
    const applied = store.replaceDocumentModel("doc-11", replacement);
    expect(applied).toBe(replacement);
    expect(record.model).toBe(replacement);
    expect(record.model.id).toBe("doc-11");
    expect(record.lastModified).toBe(stamp);
    expect(putRecord.mock.calls.length).toBe(writes);
  });
});

// --- SPEC-103: decks join the draft machinery ---------------------------------

/** A small deck carrying a retained package and an image, the two arrays the
 * store splits out of the shell. */
function deck(id: string): PresentationModel {
  return {
    kind: "presentation",
    id,
    title: "Deck",
    format: "pptx",
    slides: [
      { id: "s1", frames: [] },
      { id: "s2", frames: [] },
    ],
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    opaque: [
      { name: "[Content_Types].xml", data: "PFR5cGVzLz4=" },
      { name: "ppt/media/image1.png", data: "aGk=" },
    ],
    media: [{ mediaId: "ppt/media/image1.png", contentType: "image/png", data: "aGk=" }],
  };
}

function writtenKeys(putRecord: { mock: { calls: unknown[][] } }): string[] {
  return putRecord.mock.calls.map((call) => (call[0] as PersistedRecord).id);
}

function writtenRecord(putRecord: { mock: { calls: unknown[][] } }, key: string): PersistedRecord | undefined {
  return putRecord.mock.calls.map((call) => call[0] as PersistedRecord).find((record) => record.id === key);
}

describe("deck drafts (SPEC-103)", () => {
  it("saveDraft writes a deck's draft — the presentation guard is lifted (AC-1)", async () => {
    const { store, putRecord } = await loadStore();
    store.replacePresentationModel("deck-1", deck("deck-1"));
    await store.saveDraft("deck-1", { caret: null, scroll: 0, slideIndex: 1 });
    const draft = writtenRecord(putRecord, "draft:deck-1") as DraftRecord;
    expect(draft.docId).toBe("deck-1");
    expect(draft.slideIndex).toBe(1);
    expect(draft.caret).toBeNull();
    expect(draft.scroll).toBe(0);
    // The shell carries neither array: both went to their own records.
    expect((draft.model as PresentationModel).opaque).toEqual([]);
    expect(draft.model.media).toEqual([]);
  });

  it("getOrCreateDocument still fails loudly over a deck's id (AC-2)", async () => {
    const { store } = await loadStore();
    store.replacePresentationModel("deck-2", deck("deck-2"));
    expect(() => store.getOrCreateDocument("deck-2")).toThrow(/presentation/);
  });

  it("tick 1 writes shell + package + media; ticks 2..n write the shell alone (AC-14)", async () => {
    const { store, putRecord } = await loadStore();
    const model = store.replacePresentationModel("deck-3", deck("deck-3"));
    await store.saveDraft("deck-3", { caret: null, scroll: 0, slideIndex: 0 });
    expect(writtenKeys(putRecord).toSorted()).toEqual(["draft:deck-3", "draftmedia:deck-3", "draftopaque:deck-3"]);
    for (let tick = 2; tick <= 5; tick++) {
      putRecord.mockClear();
      // An edit as useSlideRender.commit() makes one: a fresh slides array and
      // a re-identified current slide — `media` and `opaque` untouched.
      model.slides = model.slides.slice();
      model.slides[0] = { ...model.slides[0], frames: [] };
      await store.saveDraft("deck-3", { caret: null, scroll: 0, slideIndex: tick % 2 });
      expect(writtenKeys(putRecord)).toEqual(["draft:deck-3"]);
    }
  });

  it("after a commit, a draft whose package and media match the committed ones writes the shell alone", async () => {
    const { store, putRecord } = await loadStore();
    store.replacePresentationModel("deck-4", deck("deck-4"));
    await store.commitDocument("deck-4", { caret: null, scroll: 0, slideIndex: 0 });
    putRecord.mockClear();
    await store.saveDraft("deck-4", { caret: null, scroll: 0, slideIndex: 0 });
    expect(writtenKeys(putRecord)).toEqual(["draft:deck-4"]);
  });

  it("a package that diverged from the committed one goes to the DRAFT key, never the committed one", async () => {
    const { store, putRecord } = await loadStore();
    const model = store.replacePresentationModel("deck-5", deck("deck-5"));
    await store.commitDocument("deck-5", { caret: null, scroll: 0 });
    putRecord.mockClear();
    model.opaque = [...model.opaque!]; // e.g. a restored version's package
    await store.saveDraft("deck-5", { caret: null, scroll: 0 });
    expect(writtenKeys(putRecord)).toContain("draftopaque:deck-5");
    expect(writtenKeys(putRecord)).not.toContain("opaque:deck-5");
  });

  it("commitDocument records the deck's slideIndex and splits the package out of the committed shell", async () => {
    const { store, putRecord } = await loadStore();
    store.replacePresentationModel("deck-6", deck("deck-6"));
    await store.commitDocument("deck-6", { caret: null, scroll: 0, slideIndex: 1 });
    expect(store.getStoredRecord("deck-6")?.slideIndex).toBe(1);
    expect((writtenRecord(putRecord, "opaque:deck-6") as OpaqueRecord).opaque).toEqual(deck("deck-6").opaque);
    const shell = writtenRecord(putRecord, "deck-6") as StoredDocumentRecord;
    expect((shell.model as PresentationModel).opaque).toEqual([]);
    expect(shell.slideIndex).toBe(1);
    // The in-memory model keeps its package: only the stored shell is slim.
    expect(store.getPresentation("deck-6")?.opaque).toEqual(deck("deck-6").opaque);
  });

  it("a failed package write clears its latch and keeps the package inline in the committed shell", async () => {
    const { store, putRecord } = await loadStore();
    store.replacePresentationModel("deck-7", deck("deck-7"));
    putRecord.mockImplementation(async (record: PersistedRecord) => {
      if (record.id === "opaque:deck-7") {
        throw new DOMException("quota", "QuotaExceededError");
      }
    });
    await expect(store.commitDocument("deck-7", { caret: null, scroll: 0 })).rejects.toBeTruthy();
    // Never a stripped shell over a package that did not land: the shell
    // write waits for the package write's outcome, then goes out inline.
    await vi.waitFor(() => expect(writtenRecord(putRecord, "deck-7")).toBeDefined());
    const inline = writtenRecord(putRecord, "deck-7") as StoredDocumentRecord;
    expect((inline.model as PresentationModel).opaque).toEqual(deck("deck-7").opaque);

    putRecord.mockReset().mockResolvedValue(undefined);
    await store.commitDocument("deck-7", { caret: null, scroll: 0 });
    // Same reference, retried — not latched as already persisted.
    expect(writtenKeys(putRecord)).toContain("opaque:deck-7");
    expect(((writtenRecord(putRecord, "deck-7") as StoredDocumentRecord).model as PresentationModel).opaque).toEqual([]);
  });

  it("loadDraft rejoins the draft's own package copy, else the committed one", async () => {
    const { store, getRecord } = await loadStore();
    const stored = new Map<string, unknown>();
    getRecord.mockImplementation(async (key: string) => stored.get(key));
    const shell = { ...deck("deck-8"), opaque: [], media: [] };
    stored.set("draft:deck-8", { id: "draft:deck-8", docId: "deck-8", savedAt: 5, model: shell, caret: null, scroll: 0, slideIndex: 1 });
    stored.set("opaque:deck-8", { id: "opaque:deck-8", docId: "deck-8", opaque: deck("deck-8").opaque });
    stored.set("media:deck-8", { id: "media:deck-8", docId: "deck-8", media: deck("deck-8").media });
    const committed = await store.loadDraft("deck-8");
    expect((committed!.model as PresentationModel).opaque).toEqual(deck("deck-8").opaque);
    expect(committed?.model.media).toEqual(deck("deck-8").media);
    expect(committed?.slideIndex).toBe(1);

    const own = [{ name: "[Content_Types].xml", data: "b3du" }];
    stored.set("draft:deck-8", { id: "draft:deck-8", docId: "deck-8", savedAt: 5, model: { ...shell, opaque: [] }, caret: null, scroll: 0 });
    stored.set("draftopaque:deck-8", { id: "draftopaque:deck-8", docId: "deck-8", opaque: own });
    expect(((await store.loadDraft("deck-8"))!.model as PresentationModel).opaque).toEqual(own);
  });
});

describe("the deck package record on load (SPEC-103)", () => {
  it("rejoins a stored package, seeds its latch, and never lists the record", async () => {
    const shell: StoredDocumentRecord = { id: "deck-9", format: "pptx", lastModified: 1, model: { ...deck("deck-9"), opaque: [], media: [] } };
    const opaque: OpaqueRecord = { id: "opaque:deck-9", docId: "deck-9", opaque: deck("deck-9").opaque! };
    const media: PersistedRecord = { id: "media:deck-9", docId: "deck-9", media: deck("deck-9").media! };
    const draftOpaque: OpaqueRecord = { id: "draftopaque:deck-9", docId: "deck-9", opaque: [] };
    const { store, putRecord } = await loadStore({
      loadAllRecords: vi.fn().mockResolvedValue([shell, opaque, media, draftOpaque]),
    });
    await store.initLibrary();
    expect(store.getPresentation("deck-9")?.opaque).toEqual(deck("deck-9").opaque);
    expect(store.listDocuments().map((entry) => entry.id)).toEqual(["deck-9"]);
    await store.touchDocument("deck-9");
    // Both latches seeded from storage: the persist writes the shell alone.
    expect(writtenKeys(putRecord)).toEqual(["deck-9"]);
  });

  it("a deck stored before the split keeps its inline package and the next persist migrates it (legacy-inline)", async () => {
    const legacy: StoredDocumentRecord = { id: "deck-10", format: "pptx", lastModified: 1, model: deck("deck-10") };
    const { store, putRecord } = await loadStore({ loadAllRecords: vi.fn().mockResolvedValue([legacy]) });
    await store.initLibrary();
    expect(store.getPresentation("deck-10")?.opaque).toEqual(deck("deck-10").opaque);
    await store.touchDocument("deck-10");
    // The latch was left unseeded, so the package is written out — and only
    // then is the shell stripped of it.
    expect((writtenRecord(putRecord, "opaque:deck-10") as OpaqueRecord).opaque).toEqual(deck("deck-10").opaque);
    expect(((writtenRecord(putRecord, "deck-10") as StoredDocumentRecord).model as PresentationModel).opaque).toEqual([]);
  });

  it("refreshDocumentFromStorage rejoins a newer stored deck's package", async () => {
    const { store, getRecord } = await loadStore();
    store.replacePresentationModel("deck-11", deck("deck-11"));
    const stored = new Map<string, unknown>([
      ["deck-11", { id: "deck-11", format: "pptx", lastModified: Date.now() + 60_000, model: { ...deck("deck-11"), title: "Newer", opaque: [] } }],
      ["opaque:deck-11", { id: "opaque:deck-11", docId: "deck-11", opaque: deck("deck-11").opaque }],
    ]);
    getRecord.mockImplementation(async (key: string) => stored.get(key));
    const record = await store.refreshDocumentFromStorage("deck-11");
    expect(record?.model.title).toBe("Newer");
    expect((record!.model as PresentationModel).opaque).toEqual(deck("deck-11").opaque);
  });
});

describe("deleteDocument (SPEC-103)", () => {
  // SPEC-114 adds the seventh: the document's search index record, a copy
  // of its text, goes with it.
  it("removes all seven keys: the record, its draft, both media and both package records, and its index record", async () => {
    const { store, deleteRecord } = await loadStore();
    store.replacePresentationModel("deck-12", deck("deck-12"));
    store.deleteDocument("deck-12");
    expect(deleteRecord).toHaveBeenCalledTimes(7);
    expect(deleteRecord.mock.calls.map((call) => call[0]).toSorted()).toEqual([
      "deck-12",
      "draft:deck-12",
      "draftmedia:deck-12",
      "draftopaque:deck-12",
      "index:deck-12",
      "media:deck-12",
      "opaque:deck-12",
    ]);
  });
});

describe("replacePresentationModel (SPEC-103)", () => {
  it("swaps a deck's model without bumping lastModified or persisting", async () => {
    const { store, putRecord } = await loadStore();
    store.replacePresentationModel("deck-13", deck("deck-13"));
    const record = store.getStoredRecord("deck-13")!;
    const stamp = record.lastModified;
    const writes = putRecord.mock.calls.length;
    const replacement = { ...deck("other"), title: "Recovered" };
    expect(store.replacePresentationModel("deck-13", replacement)).toBe(replacement);
    expect(record.model).toBe(replacement);
    expect(record.model.id).toBe("deck-13");
    expect(record.lastModified).toBe(stamp);
    expect(putRecord.mock.calls.length).toBe(writes);
  });

  it("never crosses kinds — a deck and a text document cannot overwrite each other", async () => {
    const { store } = await loadStore();
    store.replacePresentationModel("deck-14", deck("deck-14"));
    expect(() => store.replaceDocumentModel("deck-14", { id: "x" } as never)).toThrow(/presentation/);
    store.getOrCreateDocument("doc-14");
    expect(() => store.replacePresentationModel("doc-14", deck("doc-14"))).toThrow(/text document/);
  });
});
