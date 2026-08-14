import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";

// SPEC-003 AC-23 — the persistent warning banner. Fresh modules per test so
// the store's latched storage status starts clean; the idb seam is mocked to
// steer it into each state.
async function loadWithIdb(overrides: Record<string, unknown> = {}) {
  vi.doMock("./../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    ...overrides,
  }));
  const store = await import("../lib/documentStore");
  const { default: StorageWarningBanner } = await import(
    "./StorageWarningBanner.vue"
  );
  return { store, StorageWarningBanner };
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.doUnmock("./../lib/idb");
});

describe("StorageWarningBanner", () => {
  it("renders nothing while storage works", async () => {
    const { store, StorageWarningBanner } = await loadWithIdb();
    await store.initLibrary();
    const wrapper = mount(StorageWarningBanner);
    expect(wrapper.find('[data-testid="storage-warning"]').exists()).toBe(
      false,
    );
  });

  it("shows the exact warning when storage is unavailable at startup (AC-23)", async () => {
    const { store, StorageWarningBanner } = await loadWithIdb({
      loadAllRecords: vi.fn().mockRejectedValue(new Error("no idb")),
    });
    await store.initLibrary();
    const wrapper = mount(StorageWarningBanner);
    expect(wrapper.get('[data-testid="storage-warning"]').text()).toBe(
      "Changes will not be saved in this browser",
    );
  });

  it("appears reactively when a write fails mid-session (quota exceeded)", async () => {
    const { store, StorageWarningBanner } = await loadWithIdb({
      putRecord: vi
        .fn()
        .mockRejectedValue(new DOMException("quota", "QuotaExceededError")),
    });
    await store.initLibrary();
    const wrapper = mount(StorageWarningBanner);
    expect(wrapper.find('[data-testid="storage-warning"]').exists()).toBe(
      false,
    );

    store.getOrCreateDocument("over-quota");
    // persist() now yields a real macrotask (idleScheduler's
    // MessageChannel-based yieldToMainThread, SPEC-047 AC-15/AC-16) before
    // its putRecord call, so flushPromises()'s single tick isn't enough to
    // observe the resulting storage-status flip — a few real ticks are.
    for (let i = 0; i < 5; i++) {
      await flushPromises();
    }

    expect(wrapper.get('[data-testid="storage-warning"]').text()).toBe(
      "Changes will not be saved in this browser",
    );
  });
});
