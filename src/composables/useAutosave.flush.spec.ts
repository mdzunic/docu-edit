// SPEC-060 — the close flush primitive. flushForClose is the commit boundary:
// it awaits a committed write (so the library never reads a stale row) and
// reports whether it is safe to leave. On a write failure it stays "error" and
// returns false; while a recovery bar is open it keeps the draft and leaves
// without committing (AC-26). The idb seam is mocked, as in the autosave view
// spec, so persistence is observable without a real IndexedDB.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h } from "vue";
import { mount } from "@vue/test-utils";

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

async function microtasks() {
  for (let i = 0; i < 20; i += 1) {
    await Promise.resolve();
  }
}

async function setup(id: string) {
  const store = await import("../lib/documentStore");
  const { useAutosave } = await import("./useAutosave");
  store.getOrCreateDocument(id);
  let api: ReturnType<typeof useAutosave> | null = null;
  const Harness = defineComponent({
    setup() {
      api = useAutosave({
        getCaret: () => ({ section: 0, block: 0, offset: 0 }),
        getScroll: () => 0,
        isActive: () => true,
        applyRestored: () => {},
      });
      return () => h("div");
    },
  });
  mount(Harness);
  await api!.beginSession(id);
  await microtasks();
  return { api: api!, store };
}

/** The committed (non-draft, non-media) records written to storage. */
function committedWrites(id: string) {
  return putRecord.mock.calls
    .map(([record]) => record)
    .filter((record) => record.id === id);
}

describe("useAutosave.flushForClose (SPEC-060)", () => {
  it("commits the current state and reports it is safe to leave (AC-7)", async () => {
    const { api } = await setup("commit-doc");
    putRecord.mockClear();
    const ok = await api.flushForClose();
    expect(ok).toBe(true);
    expect(committedWrites("commit-doc").length).toBeGreaterThanOrEqual(1);
    expect(api.saveState.value).toBe("saved");
  });

  it("reports it is not safe to leave when the committed write fails (AC-12)", async () => {
    const { api } = await setup("fail-doc");
    putRecord.mockRejectedValue(new DOMException("quota", "QuotaExceededError"));
    const ok = await api.flushForClose();
    expect(ok).toBe(false);
    expect(api.saveState.value).toBe("error");
  });

  it("keeps the draft and leaves without committing while a recovery bar is open (AC-26)", async () => {
    // A draft newer than the committed record opens the recovery bar.
    getRecord.mockImplementation(async (key: string) =>
      key.startsWith("draft:")
        ? {
            id: key,
            savedAt: Date.now() + 10_000,
            model: { id: "recover-doc", title: "d", sections: [] },
            caret: null,
            scroll: 0,
          }
        : undefined,
    );
    const { api } = await setup("recover-doc");
    expect(api.recovery.value).not.toBeNull();
    putRecord.mockClear();
    const ok = await api.flushForClose();
    expect(ok).toBe(true);
    // No committed write happened — the draft survives for the next open.
    expect(committedWrites("recover-doc")).toHaveLength(0);
    expect(api.recovery.value).not.toBeNull();
  });
});

describe("useAutosave.commitNow (SPEC-138)", () => {
  it("cancels the pending draft tick once the commit covers every edit — no draft stamped newer than the commit it copies", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    try {
      const { api } = await setup("commit-now-doc");
      api.markDirty();
      putRecord.mockClear();
      await api.commitNow();
      expect(committedWrites("commit-now-doc").length).toBeGreaterThanOrEqual(1);
      expect(api.saveState.value).toBe("saved");
      putRecord.mockClear();
      // Past the debounce and the 30 s deadline the tick would have had.
      await vi.advanceTimersByTimeAsync(31_000);
      await microtasks();
      expect(putRecord.mock.calls.map(([record]) => record.id).filter((key: string) => key.startsWith("draft:"))).toEqual([]);
      // A later edit still drafts on the cadence.
      api.markDirty();
      await vi.advanceTimersByTimeAsync(31_000);
      await vi.waitFor(() => expect(putRecord.mock.calls.map(([record]) => record.id)).toContain("draft:commit-now-doc"));
    } finally {
      vi.useRealTimers();
    }
  });
});
