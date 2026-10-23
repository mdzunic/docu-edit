import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { asParagraph, createDocument } from "../lib/documentModel";
import { setDomCaret } from "../lib/domCaret";
import {
  AUTOSAVE_DEBOUNCE_MS,
  AUTOSAVE_RETRY_MS,
} from "../lib/core/autosaveScheduler";

// SPEC-010 — the autosave loop through the component: typing drives the
// save state, drafts land under the draft key, failures surface in the app
// bar, and a newer draft on open triggers the recovery bar. Fresh modules
// per test (the store's map must start clean); the idb seam is mocked.
// jsdom has no Web Locks API, so the edit lock degrades to always-acquired
// and the editor mounts editable.

const putRecord = vi.fn();
const getRecord = vi.fn();
const deleteRecord = vi.fn();

beforeEach(() => {
  vi.useFakeTimers();
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
  vi.doMock("../lib/documents", () => ({ documentExists: () => true }));
});

let wrapper: VueWrapper | null = null;
// Bumped when a test ends: a test that dies on Vitest's timeout cannot cancel
// its in-flight mount, whose import resolves under the NEXT test and would
// otherwise mount a second editor and overwrite `wrapper` (afterEach then
// unmounts the wrong one and the stale editor swallows later tests' input).
let generation = 0;

afterEach(() => {
  generation += 1;
  wrapper?.unmount();
  wrapper = null;
  vi.doUnmock("../lib/idb");
  vi.doUnmock("../lib/documents");
  vi.useRealTimers();
});

async function microtasks() {
  for (let i = 0; i < 20; i += 1) {
    await Promise.resolve();
  }
}

/** Waits for the app bar's save state to read `text`. A draft write yields a
 * real macrotask before its putRecord call (documentStore's
 * putRecordYielding, SPEC-047 AC-15/AC-16), which microtasks() cannot wait
 * out under fake timers — whether it had landed by then was a race the
 * event loop decided, lost under load. vi.waitFor polls on real timers. */
async function saveStateReads(w: VueWrapper, text: string): Promise<void> {
  await vi.waitFor(() => expect(w.get('[data-testid="save-state"]').text()).toBe(text));
}

async function mountAt(path: string) {
  const gen = generation;
  const { default: DocumentView } = await import("./DocumentView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  if (gen !== generation) {
    throw new Error("mountAt: the test ended before DocumentView finished loading — not mounting into the next test");
  }
  wrapper = mount(DocumentView, {
    global: { plugins: [router] },
    attachTo: document.body,
  });
  await microtasks();
  return wrapper;
}

/** Places the DOM caret at a character offset inside a rendered paragraph. */
function placeCaret(root: Element, block: number, offset: number) {
  setDomCaret(root, { section: 0, block, offset });
}

/** Dispatches a beforeinput as the browser would for typing. */
function type(root: Element, block: number, offset: number, data: string) {
  placeCaret(root, block, offset);
  const event = new InputEvent("beforeinput", {
    bubbles: true,
    cancelable: true,
    data,
  });
  Object.defineProperty(event, "inputType", { value: "insertText" });
  root.querySelector(`[data-block="${block}"]`)!.dispatchEvent(event);
}

describe("DocumentView autosave", () => {
  it("typing shows 'Saving…', then a debounced draft write and 'All changes saved' (AC-1, AC-3, AC-4)", async () => {
    const w = await mountAt("/doc/autosave-doc");
    const root = w.get('[data-testid="document-content"]').element;
    expect(w.get('[data-testid="save-state"]').text()).toBe("All changes saved");

    type(root, 0, 0, "Hello");
    await microtasks();
    expect(root.textContent).toContain("Hello");
    expect(w.get('[data-testid="save-state"]').text()).toBe("Saving…");
    expect(
      putRecord.mock.calls.some(([record]) => record.id === "draft:autosave-doc"),
    ).toBe(false);

    await vi.advanceTimersByTimeAsync(AUTOSAVE_DEBOUNCE_MS);
    await saveStateReads(w, "All changes saved");
    const draft = putRecord.mock.calls
      .map(([record]) => record)
      .find((record) => record.id === "draft:autosave-doc");
    expect(draft).toBeDefined();
    expect(draft.caret).toEqual({ section: 0, block: 0, offset: 5 });
    expect(w.get('[data-testid="save-state"]').text()).toBe("All changes saved");
  });

  it("typing renders through the model, not the browser's default edit", async () => {
    const w = await mountAt("/doc/typing-doc");
    const root = w.get('[data-testid="document-content"]').element;
    type(root, 0, 0, "abc");
    await microtasks();
    type(root, 0, 2, "X");
    await microtasks();
    expect(root.textContent).toContain("abXc");
  });

  it("a failed draft write shows 'Not saved — browser storage is full' with a download link, and clears on recovery (AC-5, AC-18, AC-19)", async () => {
    const w = await mountAt("/doc/quota-doc");
    const root = w.get('[data-testid="document-content"]').element;
    putRecord.mockImplementation(async (record) => {
      if (record.id.startsWith("draft:")) {
        throw new DOMException("quota", "QuotaExceededError");
      }
    });

    type(root, 0, 0, "words");
    await vi.advanceTimersByTimeAsync(AUTOSAVE_DEBOUNCE_MS);
    await saveStateReads(w, "Not saved — browser storage is full");
    expect(w.get('[data-testid="save-state"]').text()).toBe(
      "Not saved — browser storage is full",
    );
    expect(w.get('[data-testid="save-state-download"]').text()).toBe(
      "Download a copy",
    );

    // Storage recovers: the retry succeeds and the failure state clears.
    putRecord.mockResolvedValue(undefined);
    await vi.advanceTimersByTimeAsync(AUTOSAVE_RETRY_MS);
    await saveStateReads(w, "All changes saved");
    expect(w.get('[data-testid="save-state"]').text()).toBe("All changes saved");
    expect(w.find('[data-testid="save-state-download"]').exists()).toBe(false);
  });

  it("prompts on leave only while unsaved; clean close commits (AC-6, AC-7)", async () => {
    const w = await mountAt("/doc/leave-doc");
    const root = w.get('[data-testid="document-content"]').element;

    // Clean: no prompt.
    let event = new Event("beforeunload", { cancelable: true });
    expect(window.dispatchEvent(event)).toBe(true);

    // Dirty (debounce still pending): prompt.
    type(root, 0, 0, "unsaved");
    await microtasks();
    event = new Event("beforeunload", { cancelable: true });
    expect(window.dispatchEvent(event)).toBe(false);

    // Saved again: no prompt, and the clean close commits the record.
    await vi.advanceTimersByTimeAsync(AUTOSAVE_DEBOUNCE_MS);
    await saveStateReads(w, "All changes saved");
    event = new Event("beforeunload", { cancelable: true });
    expect(window.dispatchEvent(event)).toBe(true);
    // commitDocument's write now yields a real macrotask (idleScheduler's
    // MessageChannel-based yieldToMainThread, SPEC-047 AC-15/AC-16) before
    // its putRecord call — a microtask-only flush (microtasks()) never
    // catches up with it under vi.useFakeTimers(), so this polls with
    // vi.waitFor instead, which is fake-timer-aware. Waits for the commit
    // specifically (offset 7), not just any "leave-doc" write, since the
    // doc's initial creation already wrote one at offset 0.
    await vi.waitFor(() => {
      const found = putRecord.mock.calls
        .map(([record]) => record)
        .filter((record) => record.id === "leave-doc")
        .at(-1);
      expect(found?.caret).toEqual({ section: 0, block: 0, offset: 7 });
    });
    const committed = putRecord.mock.calls
      .map(([record]) => record)
      .filter((record) => record.id === "leave-doc")
      .at(-1);
    expect(committed.caret).toEqual({ section: 0, block: 0, offset: 7 });
    expect(deleteRecord).toHaveBeenCalledWith("draft:leave-doc");
  });

  it("a draft newer than the committed record opens with the recovery bar showing the draft content (AC-11)", async () => {
    const draftModel = createDocument("rec-doc");
    asParagraph(draftModel.sections[0].blocks[0]).runs = [
      { text: "Recovered text", formatting: {} },
    ];
    getRecord.mockImplementation(async (key) =>
      key === "draft:rec-doc"
        ? {
            id: "draft:rec-doc",
            docId: "rec-doc",
            savedAt: Date.now() + 60_000,
            model: draftModel,
            caret: { section: 0, block: 0, offset: 14 },
            scroll: 0,
          }
        : undefined,
    );
    const w = await mountAt("/doc/rec-doc");
    const bar = w.get('[data-testid="recovery-bar"]');
    expect(bar.text()).toContain("Recovered unsaved changes from");
    expect(w.get('[data-testid="document-content"]').text()).toContain(
      "Recovered text",
    );
  });

  it("'Keep' commits the draft content and the bar disappears (AC-11)", async () => {
    const draftModel = createDocument("keep-doc");
    asParagraph(draftModel.sections[0].blocks[0]).runs = [
      { text: "Keep me", formatting: {} },
    ];
    getRecord.mockImplementation(async (key) =>
      key === "draft:keep-doc"
        ? {
            id: "draft:keep-doc",
            docId: "keep-doc",
            savedAt: Date.now() + 60_000,
            model: draftModel,
            caret: null,
            scroll: 0,
          }
        : undefined,
    );
    const w = await mountAt("/doc/keep-doc");
    await w.get('[data-testid="recovery-keep"]').trigger("click");
    await microtasks();

    expect(w.find('[data-testid="recovery-bar"]').exists()).toBe(false);
    expect(w.get('[data-testid="document-content"]').text()).toContain("Keep me");
    // See the "prompts on leave" test above: commitDocument's write now
    // yields a real macrotask before its putRecord call, so this polls
    // with vi.waitFor rather than relying on microtasks() alone. Waits for
    // the post-Keep commit specifically, not just any "keep-doc" write,
    // since the doc's initial (empty) creation already wrote one.
    await vi.waitFor(() => {
      const found = putRecord.mock.calls
        .map(([record]) => record)
        .filter((record) => record.id === "keep-doc")
        .at(-1);
      expect(found && asParagraph(found.model.sections[0].blocks[0]).runs[0]?.text).toBe("Keep me");
    });
    const committed = putRecord.mock.calls
      .map(([record]) => record)
      .filter((record) => record.id === "keep-doc")
      .at(-1);
    expect(asParagraph(committed.model.sections[0].blocks[0]).runs[0].text).toBe("Keep me");
    expect(deleteRecord).toHaveBeenCalledWith("draft:keep-doc");
  });

  it("'Discard' restores the committed state, deletes the draft and the bar disappears (AC-12, AC-13)", async () => {
    const draftModel = createDocument("discard-doc");
    asParagraph(draftModel.sections[0].blocks[0]).runs = [
      { text: "Crashed edits", formatting: {} },
    ];
    getRecord.mockImplementation(async (key) =>
      key === "draft:discard-doc"
        ? {
            id: "draft:discard-doc",
            docId: "discard-doc",
            savedAt: Date.now() + 60_000,
            model: draftModel,
            caret: null,
            scroll: 0,
          }
        : undefined,
    );
    const w = await mountAt("/doc/discard-doc");
    expect(w.get('[data-testid="document-content"]').text()).toContain(
      "Crashed edits",
    );

    await w.get('[data-testid="recovery-discard"]').trigger("click");
    await microtasks();

    expect(w.find('[data-testid="recovery-bar"]').exists()).toBe(false);
    // The committed state here is the freshly created blank document.
    expect(w.get('[data-testid="document-content"]').text()).not.toContain(
      "Crashed edits",
    );
    expect(deleteRecord).toHaveBeenCalledWith("draft:discard-doc");
    expect(w.get('[data-testid="save-state"]').text()).toBe("All changes saved");
  });
});
