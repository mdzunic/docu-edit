import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { nextTick } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { setDomCaret, setDomSelection } from "../lib/domCaret";
import { DOC_MODEL_MIME, serializeFragment } from "../lib/clipboard";
import type { Paragraph } from "../lib/documentModel";

// SPEC-016 — cut/copy/paste through the component: the ClipboardEvent
// handlers (three flavours out, three routes in), the plain-text paste
// flag, the paste options chip's single-undo-step re-application, the
// context menu wiring and the image gate. The fragment/import mechanics are
// unit-tested in lib/clipboard.spec.ts and lib/htmlPaste.spec.ts; this only
// exercises the DOM wiring. Same mocked-idb/always-acquired-lock seams as
// the other DocumentView specs. Clipboard events are plain Events with a
// Map-backed clipboardData expando — jsdom ships no DataTransfer.

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
  delete (document as unknown as Record<string, unknown>).execCommand;
  delete (navigator as unknown as Record<string, unknown>).clipboard;
});

async function microtasks() {
  for (let i = 0; i < 20; i += 1) {
    await Promise.resolve();
  }
}

function paragraph(runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

async function mountWithBlocks(id: string, blocks: Paragraph[]) {
  const gen = generation;
  const store = await import("../lib/documentStore");
  const model = store.getOrCreateDocument(id);
  model.sections = [{ blocks }];
  const { default: DocumentView } = await import("./DocumentView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/doc/${id}`);
  if (gen !== generation) {
    throw new Error("mountWithBlocks: the test ended before DocumentView finished loading — not mounting into the next test");
  }
  wrapper = mount(DocumentView, {
    global: { plugins: [router] },
    attachTo: document.body,
  });
  await microtasks();
  return store;
}

function contentRoot(): HTMLElement {
  return document.querySelector('[data-testid="document-content"]')!;
}

function placeCaret(block: number, offset: number) {
  setDomCaret(contentRoot(), { section: 0, block, offset });
}

function selectRange(block: number, from: number, to: number, toBlock = block) {
  setDomSelection(
    contentRoot(),
    { section: 0, block, offset: from },
    { section: 0, block: toBlock, offset: to },
  );
  document.dispatchEvent(new Event("selectionchange"));
}

interface FakeClipboardData {
  setData(type: string, value: string): void;
  getData(type: string): string;
  items: Array<{ kind: string; type: string }>;
  store: Map<string, string>;
}

function fakeClipboardData(
  initial: Record<string, string> = {},
  items: Array<{ kind: string; type: string }> = [],
): FakeClipboardData {
  const store = new Map(Object.entries(initial));
  return {
    store,
    items,
    setData: (type, value) => {
      store.set(type, value);
    },
    getData: (type) => store.get(type) ?? "",
  };
}

async function dispatchClipboard(type: string, data: FakeClipboardData) {
  const event = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", { value: data });
  contentRoot().dispatchEvent(event);
  await nextTick();
  await nextTick();
}

function fireKey(init: KeyboardEventInit) {
  contentRoot().dispatchEvent(
    new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
  );
}

function blockTexts(): string[] {
  return Array.from(contentRoot().querySelectorAll(".doc-paragraph")).map(
    (el) => el.textContent ?? "",
  );
}

describe("copy and cut", () => {
  it("copy writes plain text, HTML and the internal payload (AC-2, AC-7..9)", async () => {
    await mountWithBlocks("clip-copy", [paragraph([["hello ", {}], ["world", { bold: true }]])]);
    selectRange(0, 0, 11);
    const data = fakeClipboardData();
    await dispatchClipboard("copy", data);
    expect(data.getData("text/plain")).toBe("hello world");
    expect(data.getData("text/html")).toBe("<p>hello <b>world</b></p>");
    const payload = JSON.parse(data.getData(DOC_MODEL_MIME));
    expect(payload.mimeType).toBe(DOC_MODEL_MIME);
    expect(payload.version).toBe(1);
    expect(payload.blocks).toEqual([
      {
        type: "paragraph",
        runs: [
          { text: "hello ", formatting: {} },
          { text: "world", formatting: { bold: true } },
        ],
      },
    ]);
    // Copying edits nothing.
    expect(contentRoot().textContent).toBe("hello world");
  });

  it("cut removes the selection, fills the clipboard and undoes in one step (AC-1, AC-36..38)", async () => {
    await mountWithBlocks("clip-cut", [paragraph([["hello world", {}]])]);
    selectRange(0, 0, 6);
    const data = fakeClipboardData();
    await dispatchClipboard("cut", data);
    expect(data.getData("text/plain")).toBe("hello ");
    expect(contentRoot().textContent).toBe("world");
    const undoButton = document.querySelector('[data-testid="undo-button"]')!;
    expect(undoButton.getAttribute("title")).toContain("Undo cut");
    fireKey({ key: "z", ctrlKey: true });
    await nextTick();
    expect(contentRoot().textContent).toBe("hello world");
  });

  it("a collapsed selection copies nothing", async () => {
    await mountWithBlocks("clip-copy-empty", [paragraph([["hello", {}]])]);
    placeCaret(0, 2);
    document.dispatchEvent(new Event("selectionchange"));
    const data = fakeClipboardData();
    await dispatchClipboard("copy", data);
    expect(data.store.size).toBe(0);
  });
});

describe("paste", () => {
  it("an internal payload restores formatting and structure exactly (AC-3, AC-10..13)", async () => {
    await mountWithBlocks("clip-paste-internal", [paragraph([["", {}]])]);
    placeCaret(0, 0);
    const fragment: Paragraph[] = [
      { type: "paragraph", runs: [{ text: "Title", formatting: {} }], headingLevel: 1 },
      {
        type: "paragraph",
        runs: [{ text: "bold", formatting: { bold: true } }],
        formatting: { alignment: "center" },
      },
    ];
    const data = fakeClipboardData({
      [DOC_MODEL_MIME]: serializeFragment(fragment),
      "text/html": "<p>decoy</p>",
      "text/plain": "decoy",
    });
    await dispatchClipboard("paste", data);
    expect(blockTexts()).toEqual(["Title", "bold"]);
    expect(contentRoot().querySelector("h1")).not.toBeNull();
    const bold = contentRoot().querySelectorAll(".doc-paragraph")[1];
    expect((bold as HTMLElement).style.textAlign).toBe("center");
    expect((bold.querySelector("span") as HTMLElement).style.fontWeight).toBe("bold");
  });

  it("pasting into a selection replaces it (AC-34)", async () => {
    await mountWithBlocks("clip-paste-replace", [paragraph([["one two three", {}]])]);
    selectRange(0, 4, 7);
    const data = fakeClipboardData({ "text/plain": "TWO" });
    await dispatchClipboard("paste", data);
    expect(contentRoot().textContent).toBe("one TWO three");
  });

  it("pasting paragraphs mid-paragraph splits it correctly (AC-35)", async () => {
    await mountWithBlocks("clip-paste-split", [paragraph([["helloworld", {}]])]);
    placeCaret(0, 5);
    const fragment: Paragraph[] = [
      { type: "paragraph", runs: [{ text: "A", formatting: {} }] },
      { type: "paragraph", runs: [{ text: "B", formatting: {} }] },
    ];
    const data = fakeClipboardData({ [DOC_MODEL_MIME]: serializeFragment(fragment) });
    await dispatchClipboard("paste", data);
    expect(blockTexts()).toEqual(["helloA", "Bworld"]);
  });

  it("external HTML is imported through the sanitiser (AC-14..26, AC-39..44)", async () => {
    await mountWithBlocks("clip-paste-html", [paragraph([["", {}]])]);
    placeCaret(0, 0);
    const data = fakeClipboardData({
      "text/html":
        '<script>window.__xss = 1</script><h2 class="x" style="position:fixed">Head</h2>' +
        '<p><b>bold</b> <a href="javascript:evil()">link</a></p>',
      "text/plain": "fallback",
    });
    await dispatchClipboard("paste", data);
    expect(blockTexts()).toEqual(["Head", "bold link"]);
    expect(contentRoot().querySelector("h2")).not.toBeNull();
    expect(contentRoot().querySelector("script, a, [onclick]")).toBeNull();
    expect(document.body.innerHTML).not.toContain("javascript:evil");
  });

  it("Ctrl/Cmd+Shift+V routes the next paste down the plain-text path (AC-27, AC-28)", async () => {
    await mountWithBlocks("clip-paste-plain", [paragraph([["", {}]])]);
    placeCaret(0, 0);
    fireKey({ key: "v", ctrlKey: true, shiftKey: true });
    const data = fakeClipboardData({
      "text/html": "<p><b>rich</b></p>",
      "text/plain": "rich",
    });
    await dispatchClipboard("paste", data);
    expect(contentRoot().textContent).toBe("rich");
    // The HTML flavour's bold never reached the model.
    const span = contentRoot().querySelector(".doc-paragraph span") as HTMLElement;
    expect(span.style.fontWeight).toBe("");
  });

  it("an image-only paste is gated behind the unsupported toast (AC-46)", async () => {
    await mountWithBlocks("clip-paste-image", [paragraph([["keep", {}]])]);
    placeCaret(0, 4);
    const data = fakeClipboardData({}, [{ kind: "file", type: "image/png" }]);
    await dispatchClipboard("paste", data);
    await microtasks();
    const toast = document.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("Images aren't supported yet");
    expect(contentRoot().textContent).toBe("keep");
  });
});

describe("paste options chip", () => {
  const richClipboard = () =>
    fakeClipboardData({
      [DOC_MODEL_MIME]: serializeFragment([
        {
          type: "paragraph",
          runs: [{ text: "Fancy", formatting: { bold: true } }],
          headingLevel: 2,
        },
        { type: "paragraph", runs: [{ text: "plain", formatting: {} }] },
      ]),
    });

  it("appears after a paste with the three options (AC-29..32)", async () => {
    await mountWithBlocks("clip-chip", [paragraph([["", {}]])]);
    placeCaret(0, 0);
    await dispatchClipboard("paste", richClipboard());
    const chip = document.querySelector('[data-testid="paste-options-chip"]');
    expect(chip).not.toBeNull();
    expect(chip!.textContent).toContain("Keep source formatting");
    expect(chip!.textContent).toContain("Match destination formatting");
    expect(chip!.textContent).toContain("Keep text only");
  });

  it("re-applies the paste as one undo step (AC-33)", async () => {
    await mountWithBlocks("clip-chip-reapply", [paragraph([["seed", {}]])]);
    placeCaret(0, 4);
    await dispatchClipboard("paste", richClipboard());
    expect(blockTexts()).toEqual(["seedFancy", "plain"]);
    (
      document.querySelector('[data-testid="paste-option-text"]') as HTMLElement
    ).click();
    await microtasks();
    await nextTick();
    expect(blockTexts()).toEqual(["seedFancy", "plain"]);
    const span = contentRoot().querySelector(".doc-paragraph span") as HTMLElement;
    expect(span.style.fontWeight).toBe("");
    // One undo returns to the pre-paste document — the paste and the
    // re-application collapsed into a single step.
    fireKey({ key: "z", ctrlKey: true });
    await nextTick();
    expect(blockTexts()).toEqual(["seed"]);
  });

  it("is dismissed by the next edit", async () => {
    await mountWithBlocks("clip-chip-dismiss", [paragraph([["", {}]])]);
    placeCaret(0, 0);
    await dispatchClipboard("paste", richClipboard());
    expect(document.querySelector('[data-testid="paste-options-chip"]')).not.toBeNull();
    placeCaret(1, 5);
    const event = new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: "x" });
    Object.defineProperty(event, "inputType", { value: "insertText" });
    contentRoot().querySelector('[data-block="1"]')!.dispatchEvent(event);
    await nextTick();
    expect(document.querySelector('[data-testid="paste-options-chip"]')).toBeNull();
  });
});

describe("context menu", () => {
  it("opens on right-click with cut, copy and paste entries (AC-4..6)", async () => {
    await mountWithBlocks("clip-menu", [paragraph([["hello", {}]])]);
    contentRoot().dispatchEvent(
      new MouseEvent("contextmenu", { bubbles: true, cancelable: true }),
    );
    await nextTick();
    const menu = document.querySelector('[data-testid="editor-context-menu"]');
    expect(menu).not.toBeNull();
    expect(document.querySelector('[data-testid="context-menu-cut"]')).not.toBeNull();
    expect(document.querySelector('[data-testid="context-menu-copy"]')).not.toBeNull();
    expect(document.querySelector('[data-testid="context-menu-paste"]')).not.toBeNull();
  });

  it("cut and copy entries drive execCommand, which replays into the handlers (AC-4, AC-5)", async () => {
    await mountWithBlocks("clip-menu-copy", [paragraph([["hello", {}]])]);
    const execCommand = vi.fn().mockReturnValue(true);
    (document as unknown as Record<string, unknown>).execCommand = execCommand;
    contentRoot().dispatchEvent(
      new MouseEvent("contextmenu", { bubbles: true, cancelable: true }),
    );
    await nextTick();
    (document.querySelector('[data-testid="context-menu-copy"]') as HTMLElement).click();
    await nextTick();
    expect(execCommand).toHaveBeenCalledWith("copy");
    expect(document.querySelector('[data-testid="editor-context-menu"]')).toBeNull();

    contentRoot().dispatchEvent(
      new MouseEvent("contextmenu", { bubbles: true, cancelable: true }),
    );
    await nextTick();
    (document.querySelector('[data-testid="context-menu-cut"]') as HTMLElement).click();
    await nextTick();
    expect(execCommand).toHaveBeenCalledWith("cut");
  });

  it("the paste entry reads the async clipboard and pastes (AC-6)", async () => {
    await mountWithBlocks("clip-menu-paste", [paragraph([["", {}]])]);
    placeCaret(0, 0);
    document.dispatchEvent(new Event("selectionchange"));
    Object.defineProperty(navigator, "clipboard", {
      value: { readText: async () => "from menu" },
      configurable: true,
    });
    contentRoot().dispatchEvent(
      new MouseEvent("contextmenu", { bubbles: true, cancelable: true }),
    );
    await nextTick();
    (document.querySelector('[data-testid="context-menu-paste"]') as HTMLElement).click();
    await microtasks();
    await nextTick();
    expect(contentRoot().textContent).toBe("from menu");
  });
});
