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

  it("lets a draft write in flight land before it clears the drafts, so no draft outlives the commit (SPEC-138)", async () => {
    const { store, putRecord, deleteRecord } = await loadStore();
    store.getOrCreateDocument("doc-7");
    await flushSetupWrite(putRecord);
    const order: string[] = [];
    let land = (): void => {};
    putRecord.mockImplementation((record: { id: string }) =>
      record.id === "draft:doc-7"
        ? new Promise<void>((resolve) => {
            land = () => {
              order.push("put draft:doc-7");
              resolve();
            };
          })
        : Promise.resolve(),
    );
    deleteRecord.mockImplementation((key: string) => {
      order.push(`delete ${key}`);
      return Promise.resolve();
    });
    const ticks = async (count: number) => {
      for (let tick = 0; tick < count; tick += 1) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    };
    const draft = store.saveDraft("doc-7", { caret: null, scroll: 0 });
    for (let tick = 0; tick < 100 && !putRecord.mock.calls.some(([record]) => record.id === "draft:doc-7"); tick += 1) {
      await ticks(1);
    }
    const commit = store.commitDocument("doc-7", { caret: null, scroll: 0 });
    await ticks(20);
    expect(order).not.toContain("delete draft:doc-7");
    land();
    await Promise.all([draft, commit]);
    expect(order.indexOf("put draft:doc-7")).toBeGreaterThanOrEqual(0);
    expect(order.indexOf("put draft:doc-7")).toBeLessThan(order.indexOf("delete draft:doc-7"));
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

// --- SPEC-138: workbooks join the draft machinery -----------------------------
//
// A workbook draft never touches a committed record: its chunks shadow the
// committed ones under `draftsheetrows:` and its sources go to
// `draftsheetsource:` only once one was replaced, the shell carries the
// `rowKeys` manifest loadDraft trusts, and a commit writes the committed
// records in one batch, then clears the whole draft range. Costs are counted
// at the idb seam: a stateful in-memory double recording every operation.

const inRange = (key: string, [lower, upper]: [string, string]) => key >= lower && key <= upper;

/** An idb seam over a Map, every put, batch, delete and range recorded. */
function workbookIdb() {
  const data = new Map<string, unknown>();
  const puts: string[] = [];
  const batches: Array<Array<{ put?: { id: string }; delete?: string; deleteRange?: [string, string] }>> = [];
  const deletes: string[] = [];
  const ranges: Array<[string, string]> = [];
  const seam = {
    loadAllRecords: vi.fn(async () => [...data.values()].map((value) => structuredClone(value))),
    getRecord: vi.fn(async (key: string) => (data.has(key) ? structuredClone(data.get(key)) : undefined)),
    putRecord: vi.fn(async (record: { id: string }) => {
      puts.push(record.id);
      data.set(record.id, structuredClone(record));
    }),
    deleteRecord: vi.fn(async (key: string) => {
      deletes.push(key);
      data.delete(key);
    }),
    deleteRecordRange: vi.fn(async (lower: string, upper: string) => {
      ranges.push([lower, upper]);
      for (const key of data.keys()) {
        if (inRange(key, [lower, upper])) {
          data.delete(key);
        }
      }
    }),
    writeBatch: vi.fn(async (ops: Array<{ put?: { id: string }; delete?: string; deleteRange?: [string, string] }>) => {
      batches.push(ops);
      for (const op of ops) {
        if (op.put) {
          data.set(op.put.id, structuredClone(op.put));
        } else if (op.delete !== undefined) {
          data.delete(op.delete);
        } else if (op.deleteRange) {
          for (const key of data.keys()) {
            if (inRange(key, op.deleteRange)) {
              data.delete(key);
            }
          }
        }
      }
    }),
  };
  return { seam, data, puts, batches, deletes, ranges, clear: () => void [puts, batches, deletes, ranges].forEach((log) => (log.length = 0)) };
}

async function loadWorkbookStore() {
  const idb = workbookIdb();
  vi.doMock("./idb", () => idb.seam);
  const store = await import("./documentStore");
  const { createBlankWorkbook } = await import("./spreadsheet/sheets");
  return { store, idb, createBlankWorkbook };
}

/** Real ticks: every store write yields a macrotask first. */
async function settleWrites(ticks = 20) {
  for (let i = 0; i < ticks; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

type Row = import("./spreadsheet/model").Row;
type SpreadsheetModel = import("./spreadsheet/model").SpreadsheetModel;

function sheetRows(count: number, tag: string): Row[] {
  return Array.from({ length: count }, (_, index) => ({ index, cells: [{ col: 0, value: `${tag}${index}` }, { col: 1, value: index }] }));
}

/** Three sheets: "one" (1,100 rows — three chunks), "two" (600 rows — two
 * chunks) and "three" (3 rows), ods-style sources on each and a package. */
function threeSheets(createBlankWorkbook: () => SpreadsheetModel): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  model.sheets = [
    { ...base, id: "one", name: "One", rows: sheetRows(1_100, "a"), sourceXml: "<table:table table:name=\"One\"/>" },
    { ...base, id: "two", name: "Two", rows: sheetRows(600, "b"), sourceXml: "<table:table table:name=\"Two\"/>" },
    { ...base, id: "three", name: "Three", rows: sheetRows(3, "c"), sourceXml: "<table:table table:name=\"Three\"/>" },
  ];
  model.opaque = [{ name: "content.xml", data: "PGNvbnRlbnQvPg==" }];
  model.odsContentEnvelope = { prefix: "<doc>", suffix: "</doc>" };
  return model;
}

/** A copy-on-write cell edit, as the edit pipeline makes one. */
function edit(model: SpreadsheetModel, sheetIndex: number, rowIndex: number, value: string): void {
  const target = model.sheets[sheetIndex]!;
  const rows = target.rows.map((row) => (row.index === rowIndex ? { ...row, cells: [{ col: 0, value }, ...row.cells.slice(1)] } : row));
  const sheets = model.sheets.slice();
  sheets[sheetIndex] = { ...target, rows, dirty: true };
  model.sheets = sheets;
}

/** Stores, commits, and returns the workbook with the logs cleared. */
async function committedWorkbook() {
  const loaded = await loadWorkbookStore();
  const model = threeSheets(loaded.createBlankWorkbook);
  loaded.store.storeSpreadsheet(model);
  await loaded.store.commitDocument(model.id, { caret: null, scroll: 0 });
  await settleWrites();
  loaded.idb.clear();
  return { ...loaded, model, id: model.id };
}

const VIEW = { sheetId: "two", active: { row: 4, col: 1 }, topRow: 3, leftCol: 0 };

describe("workbook drafts (SPEC-138)", () => {
  it("the spreadsheet guard is lifted from saveDraft; getOrCreateDocument still refuses a workbook's id (AC-1, AC-2)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    edit(model, 1, 5, "edited");
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0, sheetView: VIEW });
    expect(idb.puts).toContain(`draft:${id}`);
    expect(() => store.getOrCreateDocument(id)).toThrow(/spreadsheet/);
    expect(store.DRAFT_SHEET_ROWS_KEY_PREFIX).toBe("draftsheetrows:");
    expect(store.DRAFT_SHEET_SOURCE_KEY_PREFIX).toBe("draftsheetsource:");
    expect(store.draftRowsKey(`sheetrows:${id}:two:0000`)).toBe(`draftsheetrows:${id}:two:0000`);
  });

  it("a draft shell carries no bulk, the manifest and the view; the sources stay committed while reference-equal (AC-6, AC-12, AC-15, AC-17, AC-18)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    edit(model, 1, 5, "edited");
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0, sheetView: VIEW });
    const draft = idb.data.get(`draft:${id}`) as import("./documentStore").DraftRecord;
    const shell = draft.model as SpreadsheetModel;
    expect(shell.sheets.every((sheet) => sheet.rows.length === 0 && sheet.sourceXml === undefined)).toBe(true);
    expect(shell.opaque).toEqual([]);
    expect(shell.odsContentEnvelope).toBeUndefined();
    expect(draft.sheetView).toEqual(VIEW);
    expect(draft.caret).toBeNull();
    expect(draft.scroll).toBe(0);
    expect(draft.sourceKey).toBe(`sheetsource:${id}`);
    expect(draft.rowKeys).toEqual({
      one: [`sheetrows:${id}:one:0000`, `sheetrows:${id}:one:0001`, `sheetrows:${id}:one:0002`],
      two: [`draftsheetrows:${id}:two:0000`, `sheetrows:${id}:two:0001`],
      three: [`sheetrows:${id}:three:0000`],
    });
    // Nothing a draft writes is a committed key.
    expect(idb.puts.every((key) => key.startsWith("draft"))).toBe(true);
    expect(idb.batches).toEqual([]);
  });

  it("counts operations per tick on a three-sheet workbook (AC-20, AC-21)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    // First tick after editing one cell of sheet 2: the shell and that chunk.
    edit(model, 1, 5, "first");
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draftsheetrows:${id}:two:0000`, `draft:${id}`]);
    // A second edit in the same chunk: the shell and that chunk again.
    idb.clear();
    edit(model, 1, 7, "second");
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draftsheetrows:${id}:two:0000`, `draft:${id}`]);
    // An edit to sheet 3 after it: the shell and sheet 3's chunk only.
    idb.clear();
    edit(model, 2, 1, "third");
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draftsheetrows:${id}:three:0000`, `draft:${id}`]);
    // A tick with nothing new: the shell alone.
    idb.clear();
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draft:${id}`]);
    expect(idb.batches).toEqual([]);
  });

  it("applies the draft-chunk latch in all three states, clearing it when a write fails (AC-13, AC-14)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    const committedRows = model.sheets[1]!.rows;
    edit(model, 1, 5, "edited");
    store.replaceSpreadsheetModel(id, model);
    // Otherwise: written, under the draft key.
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toContain(`draftsheetrows:${id}:two:0000`);
    // Equal to the last draft write: skipped.
    idb.clear();
    model.sheets = model.sheets.map((sheet) => ({ ...sheet }));
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draft:${id}`]);
    // Reference-equal to the committed slice again (an undo): no draft
    // record, and the stale draft copy is deleted.
    idb.clear();
    model.sheets = model.sheets.map((sheet, index) => (index === 1 ? { ...sheet, rows: committedRows } : sheet));
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draft:${id}`]);
    expect(idb.deletes).toContain(`draftsheetrows:${id}:two:0000`);
    expect((idb.data.get(`draft:${id}`) as { rowKeys: Record<string, string[]> }).rowKeys.two).toEqual([`sheetrows:${id}:two:0000`, `sheetrows:${id}:two:0001`]);
    // A failed write clears the latch: the next tick writes the chunk again.
    edit(model, 1, 9, "will fail");
    idb.seam.putRecord.mockRejectedValueOnce(new DOMException("quota", "QuotaExceededError"));
    await expect(store.saveDraft(id, { caret: null, scroll: 0 })).rejects.toBeTruthy();
    idb.clear();
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draftsheetrows:${id}:two:0000`, `draft:${id}`]);
    // The committed chunk was never written by any of it.
    expect((idb.data.get(`sheetrows:${id}:two:0000`) as { rows: Row[] }).rows[5]!.cells[0]!.value).toBe("b5");
  });

  it("writes draftsheetsource:<id> once a retained source was replaced, and only then (AC-19, AC-22)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    // A structural edit that rewrites sheet 3's retained source.
    model.sheets = model.sheets.map((sheet, index) => (index === 2 ? { ...sheet, sourceXml: "<table:table table:name=\"Three\" moved=\"1\"/>" } : sheet));
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draftsheetsource:${id}`, `draft:${id}`]);
    const draft = idb.data.get(`draft:${id}`) as { sourceKey: string };
    expect(draft.sourceKey).toBe(`draftsheetsource:${id}`);
    expect((idb.data.get(`draftsheetsource:${id}`) as { sheetSources: Record<string, string> }).sheetSources.three).toContain('moved="1"');
    expect((idb.data.get(`sheetsource:${id}`) as { sheetSources: Record<string, string> }).sheetSources.three).not.toContain("moved");
    // Unchanged on the next tick: not rewritten.
    idb.clear();
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.puts).toEqual([`draft:${id}`]);
  });

  it("loadDraft rejoins every sheet from its manifest — committed chunks for untouched ones — and refuses a manifest naming a missing record (AC-16)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    edit(model, 1, 5, "drafted");
    model.sheets = model.sheets.map((sheet, index) => (index === 2 ? { ...sheet, sourceXml: "<table:table moved=\"1\"/>" } : sheet));
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0, sheetView: VIEW });
    const draft = await store.loadDraft(id);
    const recovered = draft!.model as SpreadsheetModel;
    expect(recovered.sheets.map((sheet) => sheet.rows.length)).toEqual([1_100, 600, 3]);
    expect(recovered.sheets[1]!.rows[5]!.cells[0]!.value).toBe("drafted");
    expect(recovered.sheets[0]!.rows).toEqual(model.sheets[0]!.rows);
    expect(recovered.sheets[2]!.sourceXml).toBe("<table:table moved=\"1\"/>");
    expect(recovered.opaque).toEqual(model.opaque);
    expect(recovered.odsContentEnvelope).toEqual(model.odsContentEnvelope);
    expect(draft!.sheetView).toEqual(VIEW);

    idb.data.delete(`sheetrows:${id}:one:0001`);
    expect(await store.loadDraft(id)).toBeUndefined();
  });

  it("loadDraft runs the reader-version upgrade on a recovered workbook (AC-26)", async () => {
    const upgrade = vi.fn(async (model: SpreadsheetModel) => ({ ...model, readerVersion: 999, title: "Upgraded" }));
    vi.doMock("./spreadsheet/open", async (importOriginal) => ({ ...(await importOriginal<typeof import("./spreadsheet/open")>()), upgradeSpreadsheetModel: upgrade }));
    try {
      const { store, model, id } = await committedWorkbook();
      model.readerVersion = 1;
      edit(model, 0, 1, "old reader");
      store.replaceSpreadsheetModel(id, model);
      await store.saveDraft(id, { caret: null, scroll: 0 });
      const draft = await store.loadDraft(id);
      expect(upgrade).toHaveBeenCalledTimes(1);
      expect((draft!.model as SpreadsheetModel).title).toBe("Upgraded");
    } finally {
      vi.doUnmock("./spreadsheet/open");
    }
  });

  it("a commit writes the changed committed chunks and the shell in one batch, then clears the whole draft range (AC-23, AC-24)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    edit(model, 1, 5, "kept");
    // A sheet added and deleted again mid-session leaves its draft chunk behind…
    const extra = { ...model.sheets[2]!, id: "extra", name: "Extra", rows: sheetRows(2, "x") };
    model.sheets = [...model.sheets, extra];
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0 });
    expect(idb.data.has(`draftsheetrows:${id}:extra:0000`)).toBe(true);
    model.sheets = model.sheets.slice(0, 3);
    await store.saveDraft(id, { caret: null, scroll: 0 });
    idb.clear();

    await store.commitDocument(id, { caret: null, scroll: 0, sheetView: VIEW });
    expect(idb.batches).toHaveLength(1);
    const written = idb.batches[0]!.flatMap((op) => (op.put ? [op.put.id] : []));
    expect(written).toEqual([`sheetrows:${id}:two:0000`, id]);
    expect(idb.ranges).toEqual([[`draftsheetrows:${id}:`, `draftsheetrows:${id}:￿`]]);
    expect(idb.deletes).toEqual(expect.arrayContaining([`draftsheetsource:${id}`, `draft:${id}`]));
    // …which the range cleared with every other draft record.
    expect([...idb.data.keys()].filter((key) => key.startsWith("draft"))).toEqual([]);
    expect(store.getStoredRecord(id)!.sheetView).toEqual(VIEW);
  });

  it("deleteDraft and deleteDocument clear the draft range in one operation (AC-24)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    edit(model, 0, 600, "drafted");
    store.replaceSpreadsheetModel(id, model);
    await store.saveDraft(id, { caret: null, scroll: 0 });
    idb.clear();
    await store.deleteDraft(id);
    expect(idb.ranges).toEqual([[`draftsheetrows:${id}:`, `draftsheetrows:${id}:￿`]]);
    expect(idb.deletes).toEqual(expect.arrayContaining([`draftsheetsource:${id}`, `draft:${id}`]));
    expect([...idb.data.keys()].filter((key) => key.startsWith("draft"))).toEqual([]);

    await store.saveDraft(id, { caret: null, scroll: 0 });
    idb.clear();
    store.deleteDocument(id);
    await settleWrites();
    expect(idb.ranges).toEqual([[`draftsheetrows:${id}:`, `draftsheetrows:${id}:￿`]]);
    expect([...idb.data.keys()].filter((key) => key.includes(id))).toEqual([]);
  });

  it("the start-up read excludes both draft prefixes, and initLibrary lists no draft record (AC-25)", async () => {
    const idbModule = await vi.importActual<typeof import("./idb")>("./idb");
    const { store } = await loadWorkbookStore();
    expect(idbModule.DRAFT_RECORD_KEY_PREFIXES).toEqual([store.DRAFT_SHEET_ROWS_KEY_PREFIX, store.DRAFT_SHEET_SOURCE_KEY_PREFIX]);
    const ranges = idbModule.rangesExcluding([...idbModule.LAZY_RECORD_KEY_PREFIXES, ...idbModule.DRAFT_RECORD_KEY_PREFIXES]);
    const inSomeRange = (key: string) => ranges.some(({ lower, upper }) => (lower === undefined || key > lower) && (upper === undefined || key < upper));
    expect(inSomeRange("draftsheetrows:doc:s:0000")).toBe(false);
    expect(inSomeRange("draftsheetsource:doc")).toBe(false);
    expect(inSomeRange("draft:doc")).toBe(true);
    expect(inSomeRange("doc")).toBe(true);
  });

  it("initLibrary skips a draft chunk or source record if storage hands one over", async () => {
    const { store, idb } = await loadWorkbookStore();
    idb.data.set("draftsheetrows:doc:s:0000", { id: "draftsheetrows:doc:s:0000", docId: "doc", sheetId: "s", chunk: 0, rows: [] });
    idb.data.set("draftsheetsource:doc", { id: "draftsheetsource:doc", docId: "doc", opaque: [], sheetSources: {} });
    await store.initLibrary();
    expect(store.listDocuments()).toEqual([]);
  });
});

describe("replaceSpreadsheetModel and replaceStoredModel (SPEC-138)", () => {
  it("swap a workbook's model without bumping lastModified or persisting, and dispatch on kind (AC-3, AC-4)", async () => {
    const { store, idb, model, id, createBlankWorkbook } = await committedWorkbook();
    const record = store.getStoredRecord(id)!;
    const stamp = record.lastModified;
    const replacement = createBlankWorkbook();
    expect(store.replaceSpreadsheetModel(id, replacement)).toBe(replacement);
    expect(record.model).toBe(replacement);
    expect(replacement.id).toBe(id);
    expect(record.lastModified).toBe(stamp);
    expect(store.replaceStoredModel(id, model)).toBe(model);
    expect(store.getSpreadsheet(id)).toBe(model);
    await settleWrites();
    expect(idb.puts).toEqual([]);
    expect(idb.batches).toEqual([]);
  });

  it("never crosses kinds", async () => {
    const { store, id } = await committedWorkbook();
    expect(() => store.replaceDocumentModel(id, { id: "x" } as never)).toThrow(/spreadsheet/);
    expect(() => store.replaceStoredModel(id, deck(id))).toThrow(/spreadsheet/);
    store.getOrCreateDocument("text-1");
    const { createBlankWorkbook } = await import("./spreadsheet/sheets");
    expect(() => store.replaceSpreadsheetModel("text-1", createBlankWorkbook())).toThrow(/text document/);
  });

  it("never writes a workbook another tab holds: no persist, and its release drops back to the committed shell without writing (AC-33)", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    store.setSpreadsheetHeldElsewhere(id, true);
    // The open-time recalculation of a volatile cell replaces a row.
    model.sheets = model.sheets.map((sheet, index) =>
      index === 0 ? { ...sheet, rows: sheet.rows.map((row) => (row.index === 0 ? { ...row, cells: [{ col: 0, value: "recalculated" }] } : row)) } : sheet,
    );
    await store.touchDocument(id);
    await store.releaseSpreadsheet(id);
    await settleWrites();
    expect(idb.batches).toEqual([]);
    expect(idb.puts).toEqual([]);
    expect(store.getSpreadsheet(id)).toBeUndefined();
    expect((await store.loadSpreadsheet(id))!.sheets[0]!.rows[0]!.cells[0]!.value).toBe("a0");
    // Released, the workbook is this tab's to write again.
    await store.touchDocument(id);
    await settleWrites();
    expect(idb.batches).toHaveLength(1);
  });

  it("while edits are uncommitted a view-state write stays in memory, and a release drops back to the committed shell", async () => {
    const { store, idb, model, id } = await committedWorkbook();
    edit(model, 1, 5, "uncommitted");
    store.replaceSpreadsheetModel(id, model);
    await store.storeSpreadsheetViewState({ ...model, activeSheetIndex: 2 });
    await settleWrites();
    expect(idb.batches).toEqual([]);
    await store.releaseSpreadsheet(id);
    expect(store.getSpreadsheet(id)).toBeUndefined();
    const reloaded = await store.loadSpreadsheet(id);
    expect(reloaded!.sheets[1]!.rows[5]!.cells[0]!.value).toBe("b5");
  });
});
