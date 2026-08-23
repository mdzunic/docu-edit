import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { nextTick } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { setDomCaret, setDomSelection } from "../lib/domCaret";
import type { Paragraph } from "../lib/documentModel";
import { asParagraph } from "../lib/documentModel";

// SPEC-011 — the editing engine through the component: IME composition
// commits through the command API without corrupting the model, keyboard
// document/page jumps move the caret, Ctrl/Cmd+A selects the document, a
// click on the page canvas resolves to the nearest text position, and the
// caret status live region announces movement. Same seams as the autosave
// component spec: fresh modules per test, the idb layer mocked, and jsdom's
// missing Web Locks degrading the edit lock to always-acquired.

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
  delete (document as unknown as Record<string, unknown>).caretRangeFromPoint;
});

async function microtasks() {
  for (let i = 0; i < 20; i += 1) {
    await Promise.resolve();
  }
}

function paragraph(texts: string[]): Paragraph {
  return {
    type: "paragraph",
    runs: texts.map((text) => ({ text, formatting: {} })),
  };
}

/** Seeds the store with a document and mounts the editor on it. */
async function mountWithContent(id: string, paragraphs: string[][]) {
  const gen = generation;
  const store = await import("../lib/documentStore");
  const model = store.getOrCreateDocument(id);
  model.sections = [{ blocks: paragraphs.map(paragraph) }];
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

async function storedText(id: string): Promise<string[]> {
  const store = await import("../lib/documentStore");
  const { blockText } = await import("../lib/textEditing");
  const model = store.getDocument(id);
  return model ? model.sections.flatMap((s) => s.blocks.map(blockText)) : [];
}

function contentRoot(): HTMLElement {
  return document.querySelector('[data-testid="document-content"]')!;
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

function fireComposition(type: string, data = "") {
  contentRoot().dispatchEvent(
    new CompositionEvent(type, { bubbles: true, data }),
  );
}

function fireKey(init: KeyboardEventInit) {
  contentRoot().dispatchEvent(
    new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
  );
}

function caretStatusText(): string {
  return document
    .querySelector('[data-testid="caret-status"]')!
    .textContent!.trim();
}

describe("IME composition (AC-27, AC-28, AC-29)", () => {
  it("commits the composed text into the model exactly once", async () => {
    const w = await mountWithContent("ime-commit", [["hello "]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 6 });
    fireComposition("compositionstart");
    fireComposition("compositionupdate", "に");
    fireComposition("compositionend", "日本語");
    await w.vm.$nextTick();

    expect(await storedText("ime-commit")).toEqual(["hello 日本語"]);
    expect(contentRoot().textContent).toBe("hello 日本語");
    expect(caretStatusText()).toBe("Paragraph 1 of 1, character 10");
  });

  it("does not touch the model while composition is in flight (AC-28)", async () => {
    await mountWithContent("ime-preedit", [["hello "]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 6 });
    fireComposition("compositionstart");
    const preedit = new InputEvent("beforeinput", { bubbles: true, data: "に" });
    Object.defineProperty(preedit, "inputType", { value: "insertCompositionText" });
    Object.defineProperty(preedit, "isComposing", { value: true });
    contentRoot().dispatchEvent(preedit);
    fireBeforeInput("deleteContentBackward");

    expect(await storedText("ime-preedit")).toEqual(["hello "]);
  });

  it("replaces the selection that was live when composition started", async () => {
    const w = await mountWithContent("ime-replace", [["hello world"]]);
    setDomSelection(
      contentRoot(),
      { section: 0, block: 0, offset: 6 },
      { section: 0, block: 0, offset: 11 },
    );
    fireComposition("compositionstart");
    fireComposition("compositionend", "地球");
    await w.vm.$nextTick();

    expect(await storedText("ime-replace")).toEqual(["hello 地球"]);
  });

  it("a cancelled composition leaves the model untouched", async () => {
    const w = await mountWithContent("ime-cancel", [["hello"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 5 });
    fireComposition("compositionstart");
    fireComposition("compositionend", "");
    await w.vm.$nextTick();

    expect(await storedText("ime-cancel")).toEqual(["hello"]);
  });

  it("discards browser-composed DOM nodes when re-rendering the paragraph", async () => {
    // Composition into an empty paragraph makes the browser insert a text
    // node Vue does not know about; the commit must not duplicate it.
    const w = await mountWithContent("ime-remount", [["before"], []]);
    const empty = contentRoot().querySelector('[data-block="1"]')!;
    setDomCaret(contentRoot(), { section: 0, block: 1, offset: 0 });
    fireComposition("compositionstart");
    empty.insertBefore(document.createTextNode("日本語"), empty.firstChild);
    fireComposition("compositionend", "日本語");
    await w.vm.$nextTick();
    await microtasks();
    await w.vm.$nextTick();

    expect(await storedText("ime-remount")).toEqual(["before", "日本語"]);
    expect(contentRoot().textContent).toBe("before日本語");
  });

  it("dead-key input produces the single composed character (AC-30)", async () => {
    const w = await mountWithContent("dead-key", [["caf"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 3 });
    fireComposition("compositionstart");
    fireComposition("compositionupdate", "´");
    fireComposition("compositionend", "é");
    await w.vm.$nextTick();

    expect(await storedText("dead-key")).toEqual(["café"]);
    expect(contentRoot().textContent).toBe("café");
  });
});

describe("keyboard document navigation (AC-13, AC-14, AC-15, AC-19)", () => {
  it("Ctrl+End and Ctrl+Home jump to the document's ends", async () => {
    await mountWithContent("nav-ends", [["first"], ["second words"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 2 });
    fireKey({ key: "End", ctrlKey: true });
    await nextTick();
    expect(caretStatusText()).toBe("Paragraph 2 of 2, character 13");

    fireKey({ key: "Home", ctrlKey: true });
    await nextTick();
    expect(caretStatusText()).toBe("Paragraph 1 of 2, character 1");
  });

  it("Shift+Ctrl+End extends the selection to the document end (AC-15)", async () => {
    await mountWithContent("nav-extend", [["first"], ["second words"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 2 });
    fireKey({ key: "End", ctrlKey: true, shiftKey: true });
    await nextTick();
    expect(caretStatusText()).toBe(
      "Selected from paragraph 1, character 3 to paragraph 2, character 13",
    );
    expect(document.getSelection()!.isCollapsed).toBe(false);
  });

  it("PageDown and PageUp move the caret (AC-14)", async () => {
    // jsdom has no layout, so paging degrades to its snap behaviour: the
    // position one page beyond the content is the document's end/start.
    await mountWithContent("nav-page", [["first"], ["second words"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 2 });
    fireKey({ key: "PageDown" });
    await nextTick();
    expect(caretStatusText()).toBe("Paragraph 2 of 2, character 13");

    fireKey({ key: "PageUp" });
    await nextTick();
    expect(caretStatusText()).toBe("Paragraph 1 of 2, character 1");
  });

  it("Shift+PageDown extends the selection while paging (AC-15)", async () => {
    await mountWithContent("nav-page-extend", [["first"], ["second words"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 2 });
    fireKey({ key: "PageDown", shiftKey: true });
    await nextTick();
    expect(caretStatusText()).toBe(
      "Selected from paragraph 1, character 3 to paragraph 2, character 13",
    );
  });

  it("Ctrl+A selects the whole document (AC-19)", async () => {
    await mountWithContent("nav-all", [["first"], ["second words"]]);
    setDomCaret(contentRoot(), { section: 0, block: 1, offset: 4 });
    fireKey({ key: "a", ctrlKey: true });
    await nextTick();
    expect(caretStatusText()).toBe(
      "Selected from paragraph 1, character 1 to paragraph 2, character 13",
    );
    const { modelRangeFromSelection } = await import("../lib/domCaret");
    expect(modelRangeFromSelection(contentRoot(), document.getSelection())).toEqual({
      start: { section: 0, block: 0, offset: 0 },
      end: { section: 0, block: 1, offset: 12 },
    });
  });
});

describe("page canvas clicks (AC-1)", () => {
  it("a click outside the editable content places the caret at the nearest text position", async () => {
    await mountWithContent("click-margin", [["first"], ["second words"]]);
    const textNode = contentRoot().querySelectorAll("span")[1].firstChild!;
    (document as unknown as Record<string, unknown>).caretRangeFromPoint = () => {
      const range = document.createRange();
      range.setStart(textNode, 3);
      return range;
    };
    // SPEC-028: pages render inside the editable root, so "outside the
    // editable content" is the pages wrapper (page gaps and overlays).
    const page = document.querySelector('[data-testid="document-pages"]')!;
    page.dispatchEvent(
      new MouseEvent("mousedown", { bubbles: true, cancelable: true, clientX: 5, clientY: 40 }),
    );
    await nextTick();
    expect(caretStatusText()).toBe("Paragraph 2 of 2, character 4");
  });
});

describe("command-driven mutations through the surface (AC-4, AC-7, AC-8, AC-9)", () => {
  it("typing updates the rendered text and the stored model together (AC-4)", async () => {
    const w = await mountWithContent("type-sync", [["helo"]]);
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 3 });
    fireBeforeInput("insertText", "l");
    await w.vm.$nextTick();

    expect(contentRoot().textContent).toBe("hello");
    expect(await storedText("type-sync")).toEqual(["hello"]);
  });

  it("Enter splits and carries the paragraph formatting (AC-8, AC-9)", async () => {
    const w = await mountWithContent("split-fmt", [["centered text"]]);
    const store = await import("../lib/documentStore");
    const model = store.getDocument("split-fmt")!;
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { alignment: "center" },
    };
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 8 });
    fireBeforeInput("insertParagraph");
    await w.vm.$nextTick();

    const after = store.getDocument("split-fmt")!;
    expect(after.sections[0].blocks).toHaveLength(2);
    expect(asParagraph(after.sections[0].blocks[0]).formatting).toEqual({ alignment: "center" });
    expect(asParagraph(after.sections[0].blocks[1]).formatting).toEqual({ alignment: "center" });
    expect(await storedText("split-fmt")).toEqual(["centered", " text"]);
  });

  it("Backspace at a paragraph start merges paragraphs (AC-7)", async () => {
    const w = await mountWithContent("merge-back", [["first"], ["second"]]);
    setDomCaret(contentRoot(), { section: 0, block: 1, offset: 0 });
    fireBeforeInput("deleteContentBackward");
    await w.vm.$nextTick();

    expect(await storedText("merge-back")).toEqual(["firstsecond"]);
  });
});

describe("assistive technology surface (AC-32, AC-33, AC-34)", () => {
  it("exposes a multiline textbox whose content is the document text", async () => {
    await mountWithContent("a11y", [["readable text"]]);
    const root = contentRoot();
    expect(root.getAttribute("role")).toBe("textbox");
    expect(root.getAttribute("aria-multiline")).toBe("true");
    expect(root.getAttribute("aria-label")).toBe("Document content");
    expect(root.textContent).toBe("readable text");
  });

  it("announces caret movement through a polite live region", async () => {
    await mountWithContent("a11y-live", [["one"], ["two"]]);
    const status = document.querySelector('[data-testid="caret-status"]')!;
    expect(status.getAttribute("role")).toBe("status");
    expect(status.getAttribute("aria-live")).toBe("polite");

    fireKey({ key: "End", ctrlKey: true });
    await nextTick();
    expect(status.textContent!.trim()).toBe("Paragraph 2 of 2, character 4");
  });
});
