import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { nextTick } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { setDomCaret, setDomSelection } from "../lib/domCaret";
import type { Paragraph } from "../lib/documentModel";
import type { CaretPosition } from "../lib/textEditing";

// SPEC-013 — the font family/size, colour, highlight and sub/superscript
// controls through the component: value/blank state tracking the live
// selection, and each control applying through the same command API that
// lib/commands.formatting.spec.ts exercises directly. Same seams as
// DocumentView.formatting.spec.ts (SPEC-012).

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

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

async function mountWithContent(id: string, blocks: Paragraph[]) {
  const store = await import("../lib/documentStore");
  const model = store.getOrCreateDocument(id);
  model.sections = [{ blocks }];
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

async function storedRuns(id: string) {
  const store = await import("../lib/documentStore");
  const model = store.getDocument(id);
  return model ? model.sections[0].blocks.map((block) => block.runs) : [];
}

function contentRoot(): HTMLElement {
  return document.querySelector('[data-testid="document-content"]')!;
}

function byTestId<T extends HTMLElement>(testId: string): T {
  return document.querySelector(`[data-testid="${testId}"]`)! as T;
}

function moveCaret(caret: CaretPosition) {
  setDomCaret(contentRoot(), caret);
  document.dispatchEvent(new Event("selectionchange"));
}

function moveSelection(anchor: CaretPosition, focus: CaretPosition) {
  setDomSelection(contentRoot(), anchor, focus);
  document.dispatchEvent(new Event("selectionchange"));
}

function fireKey(init: KeyboardEventInit) {
  contentRoot().dispatchEvent(
    new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
  );
}

async function commitSizeInput(text: string) {
  const input = byTestId<HTMLInputElement>("font-size-input");
  input.value = text;
  await wrapper!.get('[data-testid="font-size-input"]').trigger("input");
  await wrapper!.get('[data-testid="font-size-input"]').trigger("keydown", { key: "Enter" });
  await nextTick();
}

describe("the toolbar has all four controls plus sub/superscript (AC-1..4, AC-24)", () => {
  it("renders them when the document is editable", async () => {
    await mountWithContent("controls-present", [paragraph(["hello"])]);
    expect(byTestId("font-family-select")).toBeTruthy();
    expect(byTestId("font-size-input")).toBeTruthy();
    expect(byTestId("font-size-decrease")).toBeTruthy();
    expect(byTestId("font-size-increase")).toBeTruthy();
    expect(byTestId("text-color-button")).toBeTruthy();
    expect(byTestId("highlight-button")).toBeTruthy();
    expect(byTestId("format-subscript").getAttribute("title")).toBe("Subscript (Ctrl+,)");
    expect(byTestId("format-superscript").getAttribute("title")).toBe("Superscript (Ctrl+.)");
  });

  it("offers the document group and standard families, each option in its own typeface (AC-5..8)", async () => {
    await mountWithContent("font-list", [
      paragraph(["exotic", { fontFamily: "Aptos Narrow" }]),
    ]);
    const docGroup = byTestId<HTMLOptGroupElement>("font-family-doc-group");
    expect(docGroup.label).toBe("Fonts in this document");
    const docFamilies = Array.from(docGroup.querySelectorAll("option")).map((o) => o.value);
    // The document default plus the (locally uninstalled) family a run uses.
    expect(docFamilies).toEqual(["Aptos Narrow", "Calibri"]);
    const allGroup = byTestId<HTMLOptGroupElement>("font-family-all-group");
    const families = Array.from(allGroup.querySelectorAll("option")).map((o) => o.value);
    for (const family of ["Calibri", "Cambria", "Times New Roman", "Arial", "Georgia", "Verdana", "Courier New", "Liberation Serif", "Carlito"]) {
      expect(families).toContain(family);
    }
    for (const option of allGroup.querySelectorAll("option")) {
      expect(option.style.fontFamily).toContain(option.value);
    }
  });
});

describe("the controls reflect the caret's formatting (AC-22)", () => {
  it("shows the caret run's family, size, colour and highlight", async () => {
    await mountWithContent("caret-values", [
      paragraph(["styled", { fontFamily: "Georgia", fontSizePt: 18, color: "#ff0000", highlight: "yellow" }]),
    ]);
    moveCaret({ section: 0, block: 0, offset: 3 });
    await nextTick();
    expect(byTestId<HTMLSelectElement>("font-family-select").value).toBe("Georgia");
    expect(byTestId<HTMLInputElement>("font-size-input").value).toBe("18");
    expect(byTestId("text-color-current").getAttribute("data-color")).toBe("#ff0000");
    expect(byTestId("highlight-current").getAttribute("data-color")).toBe("yellow");
  });

  it("falls back to the document defaults for unformatted text", async () => {
    await mountWithContent("caret-defaults", [paragraph(["plain"])]);
    moveCaret({ section: 0, block: 0, offset: 2 });
    await nextTick();
    expect(byTestId<HTMLSelectElement>("font-family-select").value).toBe("Calibri");
    expect(byTestId<HTMLInputElement>("font-size-input").value).toBe("11");
    expect(byTestId("text-color-current").getAttribute("data-color")).toBe("#000000");
    expect(byTestId("highlight-current").getAttribute("data-color")).toBe("none");
  });

  it("shows blank/indeterminate for a mixed selection (AC-23)", async () => {
    await mountWithContent("mixed-values", [
      paragraph(
        ["one", { fontFamily: "Georgia", fontSizePt: 18, color: "#ff0000", highlight: "yellow" }],
        ["two", { fontFamily: "Arial", fontSizePt: 12, color: "#0000ff", highlight: "cyan" }],
      ),
    ]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 6 });
    await nextTick();
    expect(byTestId<HTMLSelectElement>("font-family-select").value).toBe("");
    expect(byTestId<HTMLInputElement>("font-size-input").value).toBe("");
    expect(byTestId("text-color-current").getAttribute("data-color")).toBe("mixed");
    expect(byTestId("highlight-current").getAttribute("data-color")).toBe("mixed");
  });
});

describe("applying formats from the controls", () => {
  it("the family dropdown formats exactly the selection", async () => {
    const id = "apply-family";
    await mountWithContent(id, [paragraph(["Hello world"])]);
    moveSelection({ section: 0, block: 0, offset: 6 }, { section: 0, block: 0, offset: 11 });
    await nextTick();
    const select = byTestId<HTMLSelectElement>("font-family-select");
    select.value = "Georgia";
    await wrapper!.get('[data-testid="font-family-select"]').trigger("change");
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [
        { text: "Hello ", formatting: {} },
        { text: "world", formatting: { fontFamily: "Georgia" } },
      ],
    ]);
  });

  it("the size input commits on Enter and accepts half-point values (AC-9, AC-10)", async () => {
    const id = "apply-size";
    await mountWithContent(id, [paragraph(["Hello"])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 5 });
    await nextTick();
    await commitSizeInput("11.5");
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { fontSizePt: 11.5 } }]]);
  });

  it("out-of-range size input clamps and shows the tooltip (AC-12)", async () => {
    const id = "clamp-size";
    await mountWithContent(id, [paragraph(["Hello"])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 5 });
    await nextTick();
    expect(document.querySelector('[data-testid="font-size-tooltip"]')).toBeNull();
    await commitSizeInput("500");
    expect(byTestId("font-size-tooltip").textContent).toContain("between 1 and 409");
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { fontSizePt: 409 } }]]);
  });

  it("the +/- buttons step by half a point (AC-11)", async () => {
    const id = "step-size";
    await mountWithContent(id, [paragraph(["Hello", { fontSizePt: 12 }])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 5 });
    await nextTick();
    await wrapper!.get('[data-testid="font-size-increase"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { fontSizePt: 12.5 } }]]);
    await wrapper!.get('[data-testid="font-size-decrease"]').trigger("click");
    await nextTick();
    await wrapper!.get('[data-testid="font-size-decrease"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { fontSizePt: 11.5 } }]]);
  });

  it("a palette colour applies and lands in the recent list (AC-16, AC-17)", async () => {
    const id = "apply-color";
    await mountWithContent(id, [paragraph(["Hello"])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 5 });
    await nextTick();
    await wrapper!.get('[data-testid="text-color-button"]').trigger("click");
    const swatch = document.querySelector(
      '[data-testid="text-color-swatch"][data-color="#ff0000"]',
    )!;
    expect(swatch).toBeTruthy();
    await wrapper!.get('[data-testid="text-color-swatch"][data-color="#ff0000"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { color: "#ff0000" } }]]);

    // Reopen the menu: the applied colour is now offered under Recent.
    await wrapper!.get('[data-testid="text-color-button"]').trigger("click");
    const recent = byTestId("text-color-recent");
    expect(
      recent.querySelector('[data-testid="text-color-recent-swatch"][data-color="#ff0000"]'),
    ).toBeTruthy();
  });

  it("the custom hex input applies any colour (AC-18)", async () => {
    const id = "apply-custom-color";
    await mountWithContent(id, [paragraph(["Hello"])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 5 });
    await nextTick();
    await wrapper!.get('[data-testid="text-color-button"]').trigger("click");
    const input = byTestId<HTMLInputElement>("text-color-custom-input");
    input.value = "rgb(255, 136, 0)";
    await wrapper!.get('[data-testid="text-color-custom-input"]').trigger("input");
    await wrapper!.get('[data-testid="text-color-custom-apply"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { color: "#ff8800" } }]]);
  });

  it("the highlight picker offers 15 colours plus No colour and applies both (AC-19)", async () => {
    const id = "apply-highlight";
    await mountWithContent(id, [paragraph(["Hello", { highlight: "yellow" }])]);
    moveSelection({ section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 5 });
    await nextTick();
    await wrapper!.get('[data-testid="highlight-button"]').trigger("click");
    expect(document.querySelectorAll('[data-testid="highlight-swatch"]')).toHaveLength(15);
    await wrapper!.get('[data-testid="highlight-swatch"][data-color="cyan"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { highlight: "cyan" } }]]);

    await wrapper!.get('[data-testid="highlight-button"]').trigger("click");
    await wrapper!.get('[data-testid="highlight-none"]').trigger("click");
    await nextTick();
    // "No colour" is an explicit false, not an absence (SPEC-013).
    expect(await storedRuns(id)).toEqual([[{ text: "Hello", formatting: { highlight: false } }]]);
  });
});

describe("subscript and superscript (AC-24..27)", () => {
  it("Ctrl+, applies subscript to the selection and Ctrl+. flips it to superscript", async () => {
    const id = "vert-align";
    await mountWithContent(id, [paragraph(["H2O"])]);
    moveSelection({ section: 0, block: 0, offset: 1 }, { section: 0, block: 0, offset: 2 });
    await nextTick();
    fireKey({ key: ",", ctrlKey: true });
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [
        { text: "H", formatting: {} },
        { text: "2", formatting: { vertAlign: "subscript" } },
        { text: "O", formatting: {} },
      ],
    ]);
    expect(byTestId("format-subscript").getAttribute("aria-pressed")).toBe("true");
    expect(byTestId("format-superscript").getAttribute("aria-pressed")).toBe("false");

    // Mutually exclusive: superscript replaces the subscript (AC-25).
    fireKey({ key: ".", ctrlKey: true });
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [
        { text: "H", formatting: {} },
        { text: "2", formatting: { vertAlign: "superscript" } },
        { text: "O", formatting: {} },
      ],
    ]);
    expect(byTestId("format-subscript").getAttribute("aria-pressed")).toBe("false");
    expect(byTestId("format-superscript").getAttribute("aria-pressed")).toBe("true");

    // Toggling the active one off clears it.
    fireKey({ key: ".", ctrlKey: true });
    await nextTick();
    expect(await storedRuns(id)).toEqual([[{ text: "H2O", formatting: {} }]]);
  });

  it("the toolbar buttons drive the same toggle", async () => {
    const id = "vert-align-buttons";
    await mountWithContent(id, [paragraph(["xn"])]);
    moveSelection({ section: 0, block: 0, offset: 1 }, { section: 0, block: 0, offset: 2 });
    await nextTick();
    await wrapper!.get('[data-testid="format-superscript"]').trigger("click");
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [
        { text: "x", formatting: {} },
        { text: "n", formatting: { vertAlign: "superscript" } },
      ],
    ]);
  });

  it("a collapsed caret sets a pending sub/superscript the next typed text carries", async () => {
    const id = "vert-align-pending";
    await mountWithContent(id, [paragraph(["x"])]);
    moveCaret({ section: 0, block: 0, offset: 1 });
    await nextTick();
    fireKey({ key: ".", ctrlKey: true });
    await nextTick();
    expect(byTestId("format-superscript").getAttribute("aria-pressed")).toBe("true");
    const event = new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: "2" });
    Object.defineProperty(event, "inputType", { value: "insertText" });
    contentRoot().dispatchEvent(event);
    await nextTick();
    expect(await storedRuns(id)).toEqual([
      [
        { text: "x", formatting: {} },
        { text: "2", formatting: { vertAlign: "superscript" } },
      ],
    ]);
  });
});
