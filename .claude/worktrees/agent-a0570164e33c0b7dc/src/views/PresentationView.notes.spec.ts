// SPEC-083 — the speaker-notes panel on /slides/:id: the toolbar toggle and its
// persisted (per-device) open state, the Cmd/Ctrl+Shift+H shortcut, the
// per-slide plain-text binding, and burst-coalesced undo. The idb seam is
// mocked so the store runs in-memory; localStorage is cleared per test so the
// preference starts at its default (collapsed).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import type { PresentationModel } from "../lib/presentation/model";

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    loadAllTemplateRecords: vi.fn().mockResolvedValue([]),
    putTemplateRecord: vi.fn().mockResolvedValue(undefined),
    deleteTemplateRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

afterEach(() => {
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
});

async function mountWith(model: PresentationModel) {
  const store = await import("../lib/documentStore");
  store.storePresentation(model);
  const { default: PresentationView } = await import("./PresentationView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/slides/${model.id}`);
  const wrapper = mount(PresentationView, { attachTo: document.body, global: { plugins: [router] } });
  await flushPromises();
  await flushPromises();
  return { wrapper, router, store };
}

async function fixtureDeck(): Promise<PresentationModel> {
  const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
}

const FIXTURE_NOTES_TEXT = "Open with the Q3 highlight\nPause before the revenue chart";

describe("PresentationView notes panel", () => {
  it("is collapsed by default and the toolbar toggle opens it (AC-5)", async () => {
    const { wrapper } = await mountWith(await fixtureDeck());
    expect(wrapper.find('[data-testid="slides-notes-panel"]').exists()).toBe(false);
    const toggle = wrapper.get('[data-testid="slides-notes-toggle"]');
    expect(toggle.attributes("aria-pressed")).toBe("false");
    await toggle.trigger("click");
    expect(wrapper.find('[data-testid="slides-notes-panel"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="slides-notes-toggle"]').attributes("aria-pressed")).toBe("true");
  });

  it("persists the open choice across mounts via the presentationNotesOpen preference (AC-8)", async () => {
    const deck = await fixtureDeck();
    const first = await mountWith(deck);
    await first.wrapper.get('[data-testid="slides-notes-toggle"]').trigger("click");
    first.wrapper.unmount();
    // A fresh mount reads the persisted preference and opens the panel.
    const second = await mountWith(deck);
    expect(second.wrapper.find('[data-testid="slides-notes-panel"]').exists()).toBe(true);
  });

  it("Cmd/Ctrl+Shift+H toggles the panel (AC-6)", async () => {
    const { wrapper } = await mountWith(await fixtureDeck());
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "h", ctrlKey: true, shiftKey: true }));
    await flushPromises();
    expect(wrapper.find('[data-testid="slides-notes-panel"]').exists()).toBe(true);
  });

  it("binds the editor to the current slide's notes and keeps them per slide (AC-9/10)", async () => {
    const { wrapper } = await mountWith(await fixtureDeck());
    await wrapper.get('[data-testid="slides-notes-toggle"]').trigger("click");
    const editor = () => wrapper.get('[data-testid="slides-notes-editor"]');
    expect((editor().element as HTMLTextAreaElement).value).toBe(FIXTURE_NOTES_TEXT);
    expect(editor().attributes("aria-label")).toBe("Speaker notes for slide 1");

    // Slide 2 has no notes — the editor is empty with the prompt.
    await wrapper.findAll('[data-testid="filmstrip-slide"]')[1].trigger("click");
    expect((editor().element as HTMLTextAreaElement).value).toBe("");
    expect(editor().attributes("placeholder")).toBe("Click to add speaker notes");
    expect(editor().attributes("aria-label")).toBe("Speaker notes for slide 2");

    // Type on slide 2, switch away and back — the text is kept per slide.
    await editor().setValue("Slide two note");
    await wrapper.findAll('[data-testid="filmstrip-slide"]')[0].trigger("click");
    expect((editor().element as HTMLTextAreaElement).value).toBe(FIXTURE_NOTES_TEXT);
    await wrapper.findAll('[data-testid="filmstrip-slide"]')[1].trigger("click");
    expect((editor().element as HTMLTextAreaElement).value).toBe("Slide two note");
  });

  it("marks the slide dirty on edit and coalesces into one undo burst (AC-12/14)", async () => {
    const { wrapper } = await mountWith(await fixtureDeck());
    await wrapper.get('[data-testid="slides-notes-toggle"]').trigger("click");
    const editor = wrapper.get('[data-testid="slides-notes-editor"]');
    await editor.setValue("Rewrite one");
    await editor.setValue("Rewrite two");
    const undo = wrapper.get('[data-testid="presentation-undo"]');
    expect(undo.attributes("disabled")).toBeUndefined();
    // One burst → one undo returns to the original fixture notes.
    await undo.trigger("click");
    expect((wrapper.get('[data-testid="slides-notes-editor"]').element as HTMLTextAreaElement).value).toBe(
      FIXTURE_NOTES_TEXT,
    );
  });
});
