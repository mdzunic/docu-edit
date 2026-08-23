import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { nextTick } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { setDomCaret, setDomSelection } from "../lib/domCaret";
import type { Paragraph } from "../lib/documentModel";
import type { CaretPosition } from "../lib/textEditing";
import { asParagraph } from "../lib/documentModel";

// SPEC-012 — character formatting through the component: the toolbar and
// keyboard shortcuts drive the same command API exercised directly in
// lib/commands.spec.ts, so these tests focus on the wiring — pressed/
// indeterminate button state tracking the live selection, the collapsed-
// caret pending format surviving typing but not navigation, and the
// selection staying live across repeated toggles. Same seams as
// DocumentView.editing.spec.ts: fresh modules, mocked idb, jsdom's Web
// Locks gap degrading the edit lock to always-acquired.

beforeEach(() => {
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
});

async function microtasks() {
  for (let i = 0; i < 20; i += 1) {
    await Promise.resolve();
  }
}

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

async function mountWithContent(id: string, blocks: Paragraph[]) {
  const gen = generation;
  const store = await import("../lib/documentStore");
  const model = store.getOrCreateDocument(id);
  model.sections = [{ blocks }];
  const { default: DocumentView } = await import("./DocumentView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/doc/${id}`);
  if (gen !== generation) {
    throw new Error("mountWithContent: the test ended before DocumentView finished loading — not mounting into the next test");
  }
  wrapper = mount(DocumentView, {
    global: { plugins: [router] },
    attachTo: document.body,
  });
  await microtasks();
  await wrapper.vm.$nextTick();
  return wrapper;
}

async function storedRuns(id: string) {
  const store = await import("../lib/documentStore");
  const model = store.getDocument(id);
  return model ? model.sections[0].blocks.map((block) => asParagraph(block).runs) : [];
}

function contentRoot(): HTMLElement {
  return document.querySelector('[data-testid="document-content"]')!;
}

function formatButton(name: string): HTMLElement {
  return document.querySelector(`[data-testid="format-${name}"]`)!;
}

function fireKey(init: KeyboardEventInit) {
  contentRoot().dispatchEvent(
    new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
  );
}

function fireBeforeInput(inputType: string, data: string | null = null) {
  const event = new InputEvent("beforeinput", {
    bubbles: true,
    cancelable: true,
    data: data ?? undefined,
  });
  Object.defineProperty(event, "inputType", { value: inputType });
  contentRoot().dispatchEvent(event);
}

/** Moves the DOM caret/selection the way a real click or arrow key would,
 * firing the same event the app listens for — jsdom does not dispatch
 * selectionchange on its own for Selection API calls. */
function moveCaret(caret: CaretPosition) {
  setDomCaret(contentRoot(), caret);
  document.dispatchEvent(new Event("selectionchange"));
}

function moveSelection(anchor: CaretPosition, focus: CaretPosition) {
  setDomSelection(contentRoot(), anchor, focus);
  document.dispatchEvent(new Event("selectionchange"));
}

describe("toolbar buttons (AC-1, AC-2)", () => {
  it("renders Bold, Italic, Underline and Strikethrough with name+shortcut tooltips", async () => {
    await mountWithContent("toolbar-buttons", [paragraph(["hello"])]);
    expect(formatButton("bold").getAttribute("title")).toBe("Bold (Ctrl+B)");
    expect(formatButton("italic").getAttribute("title")).toBe("Italic (Ctrl+I)");
    expect(formatButton("underline").getAttribute("title")).toBe("Underline (Ctrl+U)");
    expect(formatButton("strikethrough").getAttribute("title")).toBe(
      "Strikethrough (Ctrl+Shift+X)",
    );
  });
});

describe("keyboard shortcuts on a selection (AC-3..AC-8)", () => {
  it("Ctrl+B bolds exactly the selection, splitting runs at the boundaries", async () => {
    const id = "shortcut-bold";
    await mountWithContent(id, [paragraph(["Hello world"])]);
    setDomSelection(
      contentRoot(),
      { section: 0, block: 0, offset: 0 },
      { section: 0, block: 0, offset: 5 },
    );
    fireKey({ key: "b", ctrlKey: true });
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [
        { text: "Hello", formatting: { bold: true } },
        { text: " world", formatting: {} },
      ],
    ]);
  });

  it("Ctrl+I, Ctrl+U and Ctrl+Shift+X apply italic, underline and strikethrough", async () => {
    const id = "shortcut-others";
    await mountWithContent(id, [paragraph(["Hello"])]);
    const select = () =>
      setDomSelection(
        contentRoot(),
        { section: 0, block: 0, offset: 0 },
        { section: 0, block: 0, offset: 5 },
      );

    select();
    fireKey({ key: "i", ctrlKey: true });
    await nextTick();
    select();
    fireKey({ key: "u", ctrlKey: true });
    await nextTick();
    select();
    fireKey({ key: "x", ctrlKey: true, shiftKey: true });
    await nextTick();

    expect(await storedRuns(id)).toEqual([
      [{ text: "Hello", formatting: { italic: true, underline: true, strikethrough: true } }],
    ]);
  });

  it("the Bold toolbar button applies the same command as the shortcut", async () => {
    const id = "toolbar-click";
    await mountWithContent(id, [paragraph(["Hello"])]);
    setDomSelection(
      contentRoot(),
      { section: 0, block: 0, offset: 0 },
      { section: 0, block: 0, offset: 5 },
    );
    await wrapper!.get('[data-testid="format-bold"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { bold: true } }]]);
  });
});

describe("toggle behaviour (AC-16, AC-17) and selection persistence (AC-13, AC-14, AC-15)", () => {
  it("applying a format twice removes it", async () => {
    const id = "toggle-twice";
    await mountWithContent(id, [paragraph(["Hello"])]);
    setDomSelection(
      contentRoot(),
      { section: 0, block: 0, offset: 0 },
      { section: 0, block: 0, offset: 5 },
    );
    fireKey({ key: "b", ctrlKey: true });
    await nextTick();
    fireKey({ key: "b", ctrlKey: true });
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: {} }]]);
  });

  it("a mixed selection shows indeterminate; the same selection's first click bolds it all, the second clears it", async () => {
    const id = "mixed-selection";
    await mountWithContent(id, [paragraph(["bo", { bold: true }], ["ld text"])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 9 });
    await nextTick();
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("mixed");

    // First click: the whole (still-selected) range becomes bold.
    await wrapper!.get('[data-testid="format-bold"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "bold text", formatting: { bold: true } }]]);
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("true");

    // The selection was preserved (not collapsed) by the first click, so a
    // second click on the toolbar — with no re-selecting in between — sees
    // the same range and clears it everywhere.
    await wrapper!.get('[data-testid="format-bold"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "bold text", formatting: {} }]]);
  });
});

describe("collapsed-caret pending format (AC-9, AC-10, AC-11)", () => {
  it("sets a pending format that the next typed characters carry", async () => {
    const id = "pending-apply";
    await mountWithContent(id, [paragraph(["hello"])]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 5 });
    fireKey({ key: "b", ctrlKey: true });
    await nextTick();
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("true");

    fireBeforeInput("insertText", "A");
    await nextTick();
    fireBeforeInput("insertText", "B");
    await nextTick();

    expect(await storedRuns(id)).toEqual([
      [
        { text: "hello", formatting: {} },
        { text: "AB", formatting: { bold: true } },
      ],
    ]);
  });

  it("moving the caret discards the pending format", async () => {
    const id = "pending-discard";
    await mountWithContent(id, [paragraph(["hello world"])]);
    // A real navigation (click/arrow key) fires selectionchange, which is
    // how the app learns the caret is at 5 — needed so the later move away
    // from 5 is actually detectable as a change.
    moveCaret({ section: 0, block: 0, offset: 5 });
    await nextTick();
    fireKey({ key: "b", ctrlKey: true });
    await nextTick();
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("true");

    moveCaret({ section: 0, block: 0, offset: 0 });
    await nextTick();
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("false");

    fireBeforeInput("insertText", "X");
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [{ text: "Xhello world", formatting: {} }],
    ]);
  });
});

describe("toolbar pressed state tracks the caret/selection (AC-12, AC-13)", () => {
  it("shows pressed when the caret sits inside already-bold text", async () => {
    const id = "caret-pressed";
    await mountWithContent(id, [paragraph(["bold text", { bold: true }])]);
    moveCaret({ section: 0, block: 0, offset: 4 });
    await nextTick();
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("true");
  });

  it("shows pressed when the entire selection is bold", async () => {
    const id = "selection-pressed";
    await mountWithContent(id, [paragraph(["bold text", { bold: true }])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 9 });
    await nextTick();
    expect(formatButton("bold").getAttribute("aria-pressed")).toBe("true");
  });
});
