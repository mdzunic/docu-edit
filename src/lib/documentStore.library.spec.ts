import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import type { StoredDocumentRecord } from "./documentStore";
import { createDocument } from "./documentModel";

// SPEC-003 — the persistent document library. Each test gets a fresh module
// instance (module state is the store) with the IndexedDB seam mocked, so
// both the happy path and the storage-failure path are exercised
// deterministically. The pre-SPEC-003 in-memory behaviour is still covered
// by documentStore.spec.ts.

function makeIdbMock(overrides: Record<string, unknown> = {}) {
  return {
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    ...overrides,
  };
}

async function loadStore(overrides: Record<string, unknown> = {}) {
  const idb = makeIdbMock(overrides);
  vi.doMock("./idb", () => idb);
  const store = await import("./documentStore");
  return { store, idb };
}

function makeRecord(
  id: string,
  title: string,
  lastModified: number,
): StoredDocumentRecord {
  const model = createDocument(id);
  model.title = title;
  return { id, format: "new", lastModified, model };
}

// Flush the fire-and-forget persist chain: persist()/persistMedia() now
// yield a real macrotask (idleScheduler's MessageChannel-based
// yieldToMainThread, SPEC-047 AC-15/AC-16) before their putRecord call, so
// a microtask-only flush no longer catches up with it. Several real ticks
// (not just one) so every hop in yield -> putRecord -> .catch has had a
// turn regardless of exactly how MessageChannel tasks and timer callbacks
// interleave in this environment.
async function flushMicrotasks() {
  for (let i = 0; i < 5; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.doUnmock("./idb");
});

describe("initLibrary", () => {
  it("hydrates the store from persisted records", async () => {
    const persisted = makeRecord("persisted-id", "From storage", 1000);
    const { store } = await loadStore({
      loadAllRecords: vi.fn().mockResolvedValue([persisted]),
    });
    await store.initLibrary();
    expect(store.getDocument("persisted-id")?.title).toBe("From storage");
    expect(store.listDocuments()).toEqual([
      {
        id: "persisted-id",
        title: "From storage",
        format: "new",
        lastModified: 1000,
      },
    ]);
    expect(store.isStorageAvailable()).toBe(true);
  });

  it("is idempotent — loads from storage only once", async () => {
    const { store, idb } = await loadStore();
    await store.initLibrary();
    await store.initLibrary();
    expect(idb.loadAllRecords).toHaveBeenCalledTimes(1);
  });

  it("resolves (not rejects) and marks storage unavailable when loading fails", async () => {
    const { store } = await loadStore({
      loadAllRecords: vi.fn().mockRejectedValue(new Error("no idb")),
    });
    await expect(store.initLibrary()).resolves.toBeUndefined();
    expect(store.isStorageAvailable()).toBe(false);
  });

  it("marks storage unavailable in environments without IndexedDB", async () => {
    // Real idb module, jsdom has no indexedDB global.
    vi.resetModules();
    const store = await import("./documentStore");
    await store.initLibrary();
    expect(store.isStorageAvailable()).toBe(false);
  });
});

describe("write-through persistence", () => {
  it("persists a newly created document keyed by its id (AC-1)", async () => {
    const { store, idb } = await loadStore();
    store.getOrCreateDocument("fresh-id");
    await flushMicrotasks();
    expect(idb.putRecord).toHaveBeenCalledWith(
      expect.objectContaining({ id: "fresh-id", format: "new" }),
    );
  });

  it("keeps working in-memory and flips the status when a write fails (AC-21, AC-22)", async () => {
    const quota = new DOMException("quota", "QuotaExceededError");
    const { store } = await loadStore({
      putRecord: vi.fn().mockRejectedValue(quota),
    });
    const listener = vi.fn();
    store.onStorageStatusChange(listener);

    const doc = store.getOrCreateDocument("quota-id");
    await flushMicrotasks();

    expect(store.isStorageAvailable()).toBe(false);
    expect(listener).toHaveBeenCalledTimes(1);
    // The document is still fully usable in-memory.
    expect(store.getDocument("quota-id")).toBe(doc);
    store.renameDocument("quota-id", "Still works");
    expect(store.getDocument("quota-id")?.title).toBe("Still works");
  });

  it("notifies status listeners only on the first failure, and honours unsubscribe", async () => {
    const { store } = await loadStore({
      putRecord: vi.fn().mockRejectedValue(new Error("fail")),
    });
    const first = vi.fn();
    const second = vi.fn();
    store.onStorageStatusChange(first);
    const unsubscribe = store.onStorageStatusChange(second);
    unsubscribe();

    store.getOrCreateDocument("a");
    await flushMicrotasks();
    store.getOrCreateDocument("b");
    await flushMicrotasks();

    expect(first).toHaveBeenCalledTimes(1);
    expect(second).not.toHaveBeenCalled();
  });
});

describe("listDocuments", () => {
  it("sorts by last modified, newest first (AC-7)", async () => {
    const { store } = await loadStore();
    const now = vi.spyOn(Date, "now");
    now.mockReturnValue(1000);
    store.getOrCreateDocument("older");
    now.mockReturnValue(2000);
    store.getOrCreateDocument("newer");

    expect(store.listDocuments().map((entry) => entry.id)).toEqual([
      "newer",
      "older",
    ]);

    // Touching the older document bumps it back to the top.
    now.mockReturnValue(3000);
    store.touchDocument("older");
    expect(store.listDocuments().map((entry) => entry.id)).toEqual([
      "older",
      "newer",
    ]);
  });

  it("exposes title, format and lastModified per entry (AC-4, AC-5, AC-6)", async () => {
    const { store } = await loadStore();
    vi.spyOn(Date, "now").mockReturnValue(42_000);
    const doc = store.getOrCreateDocument("entry-id");
    doc.title = "My report";

    expect(store.listDocuments()).toEqual([
      {
        id: "entry-id",
        title: "My report",
        format: "new",
        lastModified: 42_000,
      },
    ]);
  });
});

describe("uniqueTitle", () => {
  it("returns the title unchanged when free", async () => {
    const { store } = await loadStore();
    expect(store.uniqueTitle("Report", ["Other"])).toBe("Report");
  });

  it("appends a count when taken", async () => {
    const { store } = await loadStore();
    expect(store.uniqueTitle("Report", ["Report"])).toBe("Report (2)");
  });

  it("skips to the first free count", async () => {
    const { store } = await loadStore();
    expect(store.uniqueTitle("Report", ["Report", "Report (2)"])).toBe(
      "Report (3)",
    );
  });
});

describe("renameDocument", () => {
  it("applies the trimmed title, bumps lastModified and persists", async () => {
    const { store, idb } = await loadStore();
    const now = vi.spyOn(Date, "now");
    now.mockReturnValue(1000);
    store.getOrCreateDocument("rename-id");
    now.mockReturnValue(2000);

    expect(store.renameDocument("rename-id", "  Budget  ")).toBe("Budget");
    expect(store.getDocument("rename-id")?.title).toBe("Budget");
    expect(store.listDocuments()[0]?.lastModified).toBe(2000);
    await flushMicrotasks();
    expect(idb.putRecord).toHaveBeenLastCalledWith(
      expect.objectContaining({ id: "rename-id", lastModified: 2000 }),
    );
  });

  it("deduplicates against other documents' titles", async () => {
    const { store } = await loadStore();
    const other = store.getOrCreateDocument("other-id");
    other.title = "Report";
    store.getOrCreateDocument("rename-id");

    expect(store.renameDocument("rename-id", "Report")).toBe("Report (2)");
  });

  it("does not treat the document's own title as a duplicate", async () => {
    const { store } = await loadStore();
    const doc = store.getOrCreateDocument("self-id");
    doc.title = "Report";
    expect(store.renameDocument("self-id", "Report")).toBe("Report");
  });

  it("rejects empty, whitespace-only and over-long titles", async () => {
    const { store } = await loadStore();
    const doc = store.getOrCreateDocument("invalid-id");
    const before = doc.title;
    expect(store.renameDocument("invalid-id", "")).toBeNull();
    expect(store.renameDocument("invalid-id", "   ")).toBeNull();
    expect(
      store.renameDocument("invalid-id", "x".repeat(store.MAX_TITLE_LENGTH + 1)),
    ).toBeNull();
    expect(doc.title).toBe(before);
  });

  it("returns null for an unknown id", async () => {
    const { store } = await loadStore();
    expect(store.renameDocument("missing", "Title")).toBeNull();
  });
});

describe("duplicateDocument", () => {
  it("creates an independent copy titled '<title> (copy)' (AC-16, AC-17, AC-18)", async () => {
    const { store, idb } = await loadStore();
    const now = vi.spyOn(Date, "now");
    now.mockReturnValue(1000);
    const original = store.getOrCreateDocument("orig-id");
    original.title = "Plan";
    original.sections[0]!.blocks.push({
      type: "paragraph",
      runs: [{ text: "hello", formatting: {} }],
    });

    now.mockReturnValue(2000);
    const copy = store.duplicateDocument("orig-id")!;

    expect(copy.id).not.toBe("orig-id");
    expect(copy.model.id).toBe(copy.id);
    expect(copy.model.title).toBe("Plan (copy)");
    expect(copy.format).toBe("new");
    expect(copy.lastModified).toBe(2000);
    // Content is copied…
    expect(copy.model.sections).toEqual(original.sections);
    // …but the copy is independent: mutating it leaves the original alone.
    copy.model.sections[0]!.blocks.push({ type: "paragraph", runs: [] });
    expect(original.sections[0]!.blocks).toHaveLength(2);
    // The copy sorts to the top and is persisted.
    expect(store.listDocuments()[0]?.id).toBe(copy.id);
    await flushMicrotasks();
    expect(idb.putRecord).toHaveBeenCalledWith(
      expect.objectContaining({ id: copy.id }),
    );
  });

  it("returns undefined for an unknown id", async () => {
    const { store } = await loadStore();
    expect(store.duplicateDocument("missing")).toBeUndefined();
  });
});

describe("deleteDocument", () => {
  it("removes the document from the session and from storage (AC-15)", async () => {
    const { store, idb } = await loadStore();
    store.getOrCreateDocument("doomed-id");
    store.deleteDocument("doomed-id");

    expect(store.getDocument("doomed-id")).toBeUndefined();
    expect(store.listDocuments()).toEqual([]);
    await flushMicrotasks();
    expect(idb.deleteRecord).toHaveBeenCalledWith("doomed-id");
  });

  it("ignores unknown ids without touching storage", async () => {
    const { store, idb } = await loadStore();
    store.deleteDocument("never-existed");
    await flushMicrotasks();
    expect(idb.deleteRecord).not.toHaveBeenCalled();
  });
});
