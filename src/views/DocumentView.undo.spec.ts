import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { nextTick } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { setDomCaret, setDomSelection, modelRangeFromSelection } from "../lib/domCaret";
import { TYPING_GROUP_PAUSE_MS } from "../lib/undoHistory";

// SPEC-015 — undo/redo through the component: keyboard shortcuts and the
// toolbar buttons, typing grouped into one step and broken by a caret move,
// a pause or a non-typing command, formatting commands always their own
// step, the restored selection, redo cleared by a new edit, and history
// scoped per document and cleared when it closes. The grouping/bounding
// mechanics themselves are unit-tested directly in lib/undoHistory.spec.ts;
// this only exercises the DOM wiring. Fake timers throughout (AC-11's pause
// is time-based); same mocked-idb/always-acquired-lock seams as the other
// DocumentView specs.

beforeEach(() => {
  vi.useFakeTimers();
  vi.resetModules();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    getRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  vi.doMock("../lib/documents", () => ({ documentExists: () => true }));
});

let wrapper: VueWrapper | null = null;
// Bumped when a test ends. A test that dies on Vitest's timeout cannot cancel
// its in-flight mountAt: the dynamic import keeps going, resolves under the
// NEXT test, mounts a second editor into document.body and overwrites
// `wrapper` — so afterEach unmounts the wrong one and the stale editor keeps
// receiving every later test's input (one timed-out first test cascaded into
// 11 failures here). mountAt compares this stamp after its awaits and refuses
// to mount once its test is over.
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
  return { wrapper, router };
}

function contentRoot(): HTMLElement {
  return document.querySelector('[data-testid="document-content"]')!;
}

function blockEl(block: number): Element {
  return contentRoot().querySelector(`[data-block="${block}"]`)!;
}

function placeCaret(block: number, offset: number) {
  setDomCaret(contentRoot(), { section: 0, block, offset });
}

/** Types one character as the browser would: caret placed, then a
 * cancelable beforeinput the component intercepts. Awaits the re-render
 * (Vue's nextTick, which the component's own restoreCaretAfterRender also
 * awaits) so the DOM reflects the new text before the next character's
 * caret is placed against it. */
async function typeChar(block: number, offset: number, data: string) {
  placeCaret(block, offset);
  const event = new InputEvent("beforeinput", { bubbles: true, cancelable: true, data });
  Object.defineProperty(event, "inputType", { value: "insertText" });
  blockEl(block).dispatchEvent(event);
  await nextTick();
}

async function typeText(block: number, offset: number, text: string) {
  for (let i = 0; i < text.length; i += 1) {
    await typeChar(block, offset + i, text[i]);
  }
}

/** Types over whatever selection is currently set in the DOM (selectRange),
 * unlike typeChar which always collapses to a caret first. */
async function typeOverSelection(block: number, data: string) {
  const event = new InputEvent("beforeinput", { bubbles: true, cancelable: true, data });
  Object.defineProperty(event, "inputType", { value: "insertText" });
  blockEl(block).dispatchEvent(event);
  await nextTick();
}

/** A native caret move (arrow key, click) — jsdom does not move the DOM
 * selection on its own, so this places it directly and fires the event the
 * component listens for, same as DocumentView.formatting.spec.ts. */
function moveCaret(block: number, offset: number) {
  setDomCaret(contentRoot(), { section: 0, block, offset });
  document.dispatchEvent(new Event("selectionchange"));
}

function selectRange(block: number, from: number, to: number) {
  setDomSelection(
    contentRoot(),
    { section: 0, block, offset: from },
    { section: 0, block, offset: to },
  );
  document.dispatchEvent(new Event("selectionchange"));
}

function fireKey(init: KeyboardEventInit) {
  contentRoot().dispatchEvent(
    new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
  );
}

const undoButton = () => document.querySelector('[data-testid="undo-button"]') as HTMLButtonElement;
const redoButton = () => document.querySelector('[data-testid="redo-button"]') as HTMLButtonElement;

async function undoShortcut() {
  fireKey({ key: "z", ctrlKey: true });
  await microtasks();
}

async function redoShortcut() {
  fireKey({ key: "z", ctrlKey: true, shiftKey: true });
  await microtasks();
}

describe("toolbar buttons (AC-4..8)", () => {
  it("both start disabled with a generic tooltip", async () => {
    await mountAt("/doc/undo-buttons");
    expect(undoButton().disabled).toBe(true);
    expect(redoButton().disabled).toBe(true);
    expect(undoButton().title).toBe("Undo (Ctrl+Z)");
    expect(redoButton().title).toBe("Redo (Ctrl+Shift+Z)");
  });

  it("undo enables and names the action after an edit; redo enables after undoing", async () => {
    await mountAt("/doc/undo-buttons-2");
    await typeText(0, 0, "hi");
    await microtasks();

    expect(undoButton().disabled).toBe(false);
    expect(undoButton().title).toBe("Undo typing (Ctrl+Z)");
    expect(redoButton().disabled).toBe(true);

    await undoButton().click();
    await microtasks();
    expect(undoButton().disabled).toBe(true);
    expect(redoButton().disabled).toBe(false);
    expect(redoButton().title).toBe("Redo typing (Ctrl+Shift+Z)");
    expect(contentRoot().textContent).toBe("");
  });

  it("the redo button replays the undone edit", async () => {
    await mountAt("/doc/undo-buttons-3");
    await typeText(0, 0, "hi");
    await microtasks();
    await undoButton().click();
    await microtasks();

    await redoButton().click();
    await microtasks();
    expect(contentRoot().textContent).toBe("hi");
    expect(redoButton().disabled).toBe(true);
    expect(undoButton().disabled).toBe(false);
  });
});

describe("keyboard shortcuts (AC-1, AC-2, AC-3)", () => {
  it("Ctrl+Z undoes typed text", async () => {
    await mountAt("/doc/shortcut-undo");
    await typeText(0, 0, "abc");
    await microtasks();
    await undoShortcut();
    expect(contentRoot().textContent).toBe("");
  });

  it("Ctrl+Shift+Z redoes", async () => {
    await mountAt("/doc/shortcut-redo");
    await typeText(0, 0, "abc");
    await microtasks();
    await undoShortcut();
    await redoShortcut();
    expect(contentRoot().textContent).toBe("abc");
  });

  it("Ctrl+Y redoes (Windows convention)", async () => {
    await mountAt("/doc/shortcut-ctrl-y");
    await typeText(0, 0, "abc");
    await microtasks();
    await undoShortcut();

    fireKey({ key: "y", ctrlKey: true });
    await microtasks();
    expect(contentRoot().textContent).toBe("abc");
  });
});

describe("typing grouping (AC-9..13)", () => {
  it("consecutive typed characters undo as a single step", async () => {
    await mountAt("/doc/group-typing");
    await typeText(0, 0, "hello");
    await microtasks();

    await undoShortcut();
    expect(contentRoot().textContent).toBe("");
    expect(undoButton().disabled).toBe(true);
  });

  it("a caret move breaks the typing group", async () => {
    await mountAt("/doc/group-caret-move");
    await typeText(0, 0, "ab");
    await microtasks();
    moveCaret(0, 0);
    await typeText(0, 0, "cd");
    await microtasks();
    expect(contentRoot().textContent).toBe("cdab");

    await undoShortcut();
    expect(contentRoot().textContent).toBe("ab");
    await undoShortcut();
    expect(contentRoot().textContent).toBe("");
  });

  it("a 2 s pause breaks the typing group", async () => {
    await mountAt("/doc/group-pause");
    await typeText(0, 0, "ab");
    await microtasks();
    await vi.advanceTimersByTimeAsync(TYPING_GROUP_PAUSE_MS);
    await typeText(0, 2, "cd");
    await microtasks();
    expect(contentRoot().textContent).toBe("abcd");

    await undoShortcut();
    expect(contentRoot().textContent).toBe("ab");
    await undoShortcut();
    expect(contentRoot().textContent).toBe("");
  });

  it("a non-typing command breaks the typing group and is itself its own step", async () => {
    const mounted = (await mountAt("/doc/group-non-typing")).wrapper;
    await typeText(0, 0, "ab");
    await microtasks();

    await mounted.get('[data-testid="align-center"]').trigger("click");
    await microtasks();
    await typeText(0, 2, "cd");
    await microtasks();
    expect(contentRoot().textContent).toBe("abcd");

    await undoShortcut();
    expect(contentRoot().textContent).toBe("ab");
    expect(undoButton().title).toBe("Undo alignment (Ctrl+Z)");
    await undoShortcut();
    await undoShortcut();
    expect(contentRoot().textContent).toBe("");
    expect(undoButton().disabled).toBe(true);
  });

  it("two formatting commands over the same selection never merge (AC-13)", async () => {
    await mountAt("/doc/group-formatting");
    await typeText(0, 0, "hi");
    await microtasks();
    selectRange(0, 0, 2);
    fireKey({ key: "b", ctrlKey: true });
    await microtasks();
    selectRange(0, 0, 2);
    fireKey({ key: "i", ctrlKey: true });
    await microtasks();

    expect(undoButton().title).toBe("Undo italic (Ctrl+Z)");
    await undoShortcut();
    expect(undoButton().title).toBe("Undo bold (Ctrl+Z)");
    await undoShortcut();
    expect(undoButton().title).toBe("Undo typing (Ctrl+Z)");
  });
});

describe("selection restore (AC-14)", () => {
  it("undoing a typed replacement restores the selection that was replaced", async () => {
    await mountAt("/doc/selection-restore");
    await typeText(0, 0, "Hello world");
    await microtasks();

    selectRange(0, 6, 11);
    await typeOverSelection(0, "X");
    await microtasks();
    expect(contentRoot().textContent).toBe("Hello X");

    await undoShortcut();
    expect(contentRoot().textContent).toBe("Hello world");
    const restored = modelRangeFromSelection(contentRoot(), document.getSelection());
    expect(restored).toEqual({
      start: { section: 0, block: 0, offset: 6 },
      end: { section: 0, block: 0, offset: 11 },
    });
  });
});

describe("redo stack (AC-15)", () => {
  it("a new edit after undo clears the redo stack", async () => {
    await mountAt("/doc/redo-clear");
    await typeText(0, 0, "abc");
    await microtasks();
    await undoShortcut();
    expect(redoButton().disabled).toBe(false);

    await typeText(0, 0, "x");
    await microtasks();
    expect(redoButton().disabled).toBe(true);

    await redoShortcut();
    expect(contentRoot().textContent).toBe("x");
  });
});

describe("per-document history and clearing on close (AC-18, AC-19)", () => {
  it("a freshly opened document starts with empty history, and closing discards it", async () => {
    const { router } = await mountAt("/doc/undo-doc-a");
    await typeText(0, 0, "hello");
    await microtasks();
    expect(undoButton().disabled).toBe(false);

    await router.push("/doc/undo-doc-b");
    await microtasks();
    expect(undoButton().disabled).toBe(true);
    expect(redoButton().disabled).toBe(true);

    await router.push("/doc/undo-doc-a");
    await microtasks();
    // The content survived (it's persisted), but the undo history did not.
    expect(contentRoot().textContent).toBe("hello");
    expect(undoButton().disabled).toBe(true);
  });
});
