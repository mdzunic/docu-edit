import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { setDomCaret } from "../lib/domCaret";
import type { Paragraph } from "../lib/documentModel";

// SPEC-026 — the smart-quote / dash / decade autocorrect substitutions and the
// non-breaking-space shortcut, driven through the component so the two-step
// undo behaviour and the "typed input only" rule are exercised end-to-end.
// Same mounting seams as DocumentView.editing.spec.ts.

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    getRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  vi.doMock("../lib/documents", () => ({ documentExists: () => true }));
});

let wrapper: VueWrapper | null = null;

afterEach(() => {
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

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text ? [{ text, formatting: {} }] : [] };
}

async function mountWithContent(id: string, text: string) {
  const store = await import("../lib/documentStore");
  const model = store.getOrCreateDocument(id);
  model.sections = [{ blocks: [paragraph(text)] }];
  const { default: DocumentView } = await import("./DocumentView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/doc/${id}`);
  wrapper = mount(DocumentView, {
    global: { plugins: [router] },
    attachTo: document.body,
  });
  await microtasks();
  await wrapper.vm.$nextTick();
  return wrapper;
}

async function storedText(id: string): Promise<string> {
  const store = await import("../lib/documentStore");
  const { blockText } = await import("../lib/textEditing");
  const model = store.getDocument(id);
  return model ? blockText(model.sections[0].blocks[0]) : "";
}

function contentRoot(): HTMLElement {
  return document.querySelector('[data-testid="document-content"]')!;
}

function fireInput(inputType: string, data: string | null = null) {
  const event = new InputEvent("beforeinput", {
    bubbles: true,
    cancelable: true,
    data: data ?? undefined,
  });
  Object.defineProperty(event, "inputType", { value: inputType });
  contentRoot().dispatchEvent(event);
}

async function typeAt(offset: number, ch: string) {
  setDomCaret(contentRoot(), { section: 0, block: 0, offset });
  fireInput("insertText", ch);
  await wrapper!.vm.$nextTick();
  await microtasks();
}

function fireKey(init: KeyboardEventInit) {
  contentRoot().dispatchEvent(
    new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
  );
}

async function undo() {
  fireKey({ key: "z", ctrlKey: true });
  await wrapper!.vm.$nextTick();
  await microtasks();
}

describe("dash substitution (AC-16, AC-17)", () => {
  it("turns -- into an en dash and --- into an em dash", async () => {
    await mountWithContent("dash", "a");
    await typeAt(1, "-");
    await typeAt(2, "-");
    expect(await storedText("dash")).toBe("a–");
    await typeAt(2, "-");
    expect(await storedText("dash")).toBe("a—");
  });

  it("does not fire when the char before the pair is a dash", async () => {
    await mountWithContent("dash-guard", "a--");
    await typeAt(3, "-");
    expect(await storedText("dash-guard")).toBe("a---");
  });

  it("one undo restores the literal -- (AC-20, AC-21)", async () => {
    await mountWithContent("dash-undo", "a");
    await typeAt(1, "-");
    await typeAt(2, "-");
    expect(await storedText("dash-undo")).toBe("a–");
    await undo();
    expect(await storedText("dash-undo")).toBe("a--");
  });
});

describe("smart quotes (AC-18, AC-19)", () => {
  it("opens a single quote at a paragraph start", async () => {
    await mountWithContent("quote-open", "");
    await typeAt(0, "'");
    expect(await storedText("quote-open")).toBe("‘");
  });

  it("closes a single quote after a letter (don’t)", async () => {
    await mountWithContent("quote-close", "don");
    await typeAt(3, "'");
    expect(await storedText("quote-close")).toBe("don’");
  });

  it("double quotes too, and one undo restores the straight quote (AC-21)", async () => {
    await mountWithContent("dquote", "");
    await typeAt(0, '"');
    expect(await storedText("dquote")).toBe("“");
    await undo();
    expect(await storedText("dquote")).toBe('"');
  });
});

describe("decade idiom (AC-50)", () => {
  it("flips the opening single quote before a digit and undoes with the digit", async () => {
    await mountWithContent("decade", "");
    await typeAt(0, "'");
    expect(await storedText("decade")).toBe("‘");
    await typeAt(1, "9");
    expect(await storedText("decade")).toBe("’9");
    await undo();
    expect(await storedText("decade")).toBe("‘");
  });
});

describe("settings (AC-24, AC-27, AC-28)", () => {
  it("dashes off leaves -- literal", async () => {
    const w = await mountWithContent("dash-off", "a");
    await w.get('[data-testid="autocorrect-menu"]').trigger("click");
    await w.get('[data-testid="autocorrect-dashes"]').trigger("click");
    await typeAt(1, "-");
    await typeAt(2, "-");
    expect(await storedText("dash-off")).toBe("a--");
    expect(JSON.parse(localStorage.getItem("docu-edit:autocorrect")!)).toEqual({
      smartQuotes: true,
      dashes: false,
    });
  });

  it("never runs on paste", async () => {
    await mountWithContent("paste", "");
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 0 });
    fireInput("insertFromPaste", "--");
    await wrapper!.vm.$nextTick();
    await microtasks();
    expect(await storedText("paste")).toBe("--");
  });
});

describe("non-breaking space shortcut (AC-11)", () => {
  it("inserts U+00A0 without autocorrecting", async () => {
    await mountWithContent("nbsp", "a");
    setDomCaret(contentRoot(), { section: 0, block: 0, offset: 1 });
    fireKey({ key: " ", ctrlKey: true, shiftKey: true });
    await wrapper!.vm.$nextTick();
    await microtasks();
    expect(await storedText("nbsp")).toBe("a ");
  });
});
