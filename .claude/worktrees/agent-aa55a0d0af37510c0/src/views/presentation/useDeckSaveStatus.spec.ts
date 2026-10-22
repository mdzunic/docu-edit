import { describe, it, expect, vi, beforeEach } from "vitest";
import { defineComponent, nextTick } from "vue";
import { mount } from "@vue/test-utils";

// The store is stubbed so the storage latch can be driven directly — the point
// of the readout is that it reports real writes and a real failure, and both
// need to be forced rather than waited for.
let available = true;
const listeners = new Set<() => void>();
vi.mock("../../lib/documentStore", () => ({
  isStorageAvailable: () => available,
  onStorageStatusChange: (listener: () => void) => {
    listeners.add(listener);
    return () => listeners.delete(listener);
  },
}));

const { useDeckSaveStatus } = await import("./useDeckSaveStatus");

function harness() {
  let api!: ReturnType<typeof useDeckSaveStatus>;
  const wrapper = mount(
    defineComponent({
      setup() {
        api = useDeckSaveStatus();
        return () => null;
      },
    }),
  );
  return { api, wrapper };
}

beforeEach(() => {
  available = true;
  listeners.clear();
});

describe("useDeckSaveStatus", () => {
  it("rests on saved", () => {
    const { api } = harness();
    expect(api.saveState.value).toBe("saved");
    expect(api.saveStateText.value).toBe("All changes saved");
  });

  it("reports saving only while a write is actually in flight", async () => {
    const { api } = harness();
    let settle!: () => void;
    api.trackWrite(new Promise<void>((resolve) => (settle = resolve)));

    expect(api.saveState.value).toBe("saving");
    expect(api.saveStateText.value).toBe("Saving…");

    settle();
    await nextTick();
    await Promise.resolve();
    expect(api.saveState.value).toBe("saved");
  });

  it("stays saving until the LAST of several overlapping writes settles", async () => {
    const { api } = harness();
    let first!: () => void;
    let second!: () => void;
    api.trackWrite(new Promise<void>((resolve) => (first = resolve)));
    api.trackWrite(new Promise<void>((resolve) => (second = resolve)));

    first();
    await nextTick();
    await Promise.resolve();
    // A burst of commits overlaps; one landing does not mean the deck is saved.
    expect(api.saveState.value).toBe("saving");

    second();
    await nextTick();
    await Promise.resolve();
    expect(api.saveState.value).toBe("saved");
  });

  it("reports the failure the store has already latched", async () => {
    const { api } = harness();
    let settle!: () => void;
    api.trackWrite(new Promise<void>((resolve) => (settle = resolve)));

    available = false; // what markStorageUnavailable() leaves behind
    settle();
    await nextTick();
    await Promise.resolve();

    expect(api.saveState.value).toBe("error");
    expect(api.saveStateText.value).toContain("Not saved");
  });

  it("picks up a failure signalled while nothing is being written", async () => {
    const { api } = harness();
    available = false;
    for (const listener of listeners) listener();
    await nextTick();
    expect(api.saveState.value).toBe("error");
  });

  it("ignores a store that hands back no promise (a stubbed touchDocument)", () => {
    const { api } = harness();
    api.trackWrite(undefined);
    expect(api.saveState.value).toBe("saved");
  });
});
