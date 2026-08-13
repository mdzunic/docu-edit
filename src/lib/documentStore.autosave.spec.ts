import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import type { PersistedRecord } from "./documentStore";

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
// after must flush that pending macrotask first, or its deferred call
// picks up the reconfigured mock instead of the one active when it was
// issued.
function flushMacrotask(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
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
    await store.saveDraft("doc-1", { section: 0, block: 0, offset: 3 }, 120);
    const draft = putRecord.mock.calls.at(-1)![0];
    expect(draft.id).toBe("draft:doc-1");
    expect(draft.docId).toBe("doc-1");
    expect(draft.model.id).toBe("doc-1");
    expect(draft.caret).toEqual({ section: 0, block: 0, offset: 3 });
    expect(draft.scroll).toBe(120);
    expect(typeof draft.savedAt).toBe("number");
  });

  it("rejects on write failure WITHOUT latching the SPEC-003 banner", async () => {
    const { store, putRecord } = await loadStore();
    store.getOrCreateDocument("doc-2");
    await flushMacrotask();
    putRecord.mockRejectedValue(
      new DOMException("quota", "QuotaExceededError"),
    );
    await expect(store.saveDraft("doc-2", null, 0)).rejects.toBeTruthy();
    expect(store.isStorageAvailable()).toBe(true);
  });

  it("rejects for an unknown document id", async () => {
    const { store } = await loadStore();
    await expect(store.saveDraft("nope", null, 0)).rejects.toBeTruthy();
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
    await store.commitDocument("doc-5", { section: 0, block: 1, offset: 7 }, 33);
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
    await flushMacrotask();
    putRecord.mockRejectedValue(new Error("full"));
    await expect(store.commitDocument("doc-6", null, 0)).rejects.toBeTruthy();
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
