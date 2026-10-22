// SPEC-079 — the /slides/:id view: the filmstrip is a proper listbox with
// keyboard navigation and aria labelling, the stage announces changes, the
// read-only note is gone, and a preparation-milestone record upgrades on
// open. The idb seam is mocked so the store runs in-memory.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import { createBlankPresentation } from "../lib/presentation/slides";
import type { PresentationModel } from "../lib/presentation/model";

beforeEach(() => {
  vi.resetModules();
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
  const wrapper = mount(PresentationView, { global: { plugins: [router] } });
  await flushPromises();
  await flushPromises();
  return { wrapper, router, store };
}

async function freshDeck(): Promise<PresentationModel> {
  const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
}

describe("PresentationView filmstrip", () => {
  it("renders a listbox of options with aria-selected and labels (AC-49..52)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    const listbox = wrapper.get('[data-testid="slide-filmstrip"]');
    expect(listbox.attributes("role")).toBe("listbox");
    const options = wrapper.findAll('[data-testid="filmstrip-slide"]');
    expect(options).toHaveLength(4);
    expect(options[0].attributes("role")).toBe("option");
    expect(options[0].attributes("aria-selected")).toBe("true");
    expect(options[1].attributes("aria-selected")).toBe("false");
    expect(options[0].attributes("aria-label")).toBe("Slide 1 of 4 — Quarterly Review");
    // A slide with no title is labelled "untitled".
    expect(options[2].attributes("aria-label")).toBe("Slide 3 of 4 — untitled");
  });

  it("changes the current slide via click and ArrowDown/Home/End (AC-45..47)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    const listbox = wrapper.get('[data-testid="slide-filmstrip"]');
    const selected = () =>
      wrapper.findAll('[data-testid="filmstrip-slide"]').findIndex((o) => o.attributes("aria-selected") === "true");

    await listbox.trigger("keydown", { key: "ArrowDown" });
    expect(selected()).toBe(1);
    await listbox.trigger("keydown", { key: "End" });
    expect(selected()).toBe(3);
    await listbox.trigger("keydown", { key: "Home" });
    expect(selected()).toBe(0);

    await wrapper.findAll('[data-testid="filmstrip-slide"]')[2].trigger("click");
    expect(selected()).toBe(2);
  });

  it("navigates with PageUp/PageDown outside a text input (AC-48)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    const selected = () =>
      wrapper.findAll('[data-testid="filmstrip-slide"]').findIndex((o) => o.attributes("aria-selected") === "true");

    window.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown" }));
    await flushPromises();
    expect(selected()).toBe(1);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "PageUp" }));
    await flushPromises();
    expect(selected()).toBe(0);
    wrapper.unmount();
  });

  it("announces the current slide via an aria-live region (AC-53)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    const live = wrapper.get('[data-testid="slide-live"]');
    expect(live.attributes("aria-live")).toBe("polite");
    expect(live.text()).toBe("Slide 1 of 4");
  });

  it("drops the preparation milestone's read-only note (AC-66)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    expect(wrapper.find('[data-testid="presentation-note"]').exists()).toBe(false);
  });

  it("upgrades a preparation-milestone record on open (AC-54..57)", async () => {
    const fresh = await freshDeck();
    // The prep milestone stored outline-shaped slides and no slideWidthCm.
    const legacy = {
      kind: "presentation",
      id: fresh.id,
      title: "mini-deck",
      format: "pptx",
      slides: [{ title: "Quarterly Review", lines: ["FY26 results and outlook"] }],
      opaque: fresh.opaque,
    } as unknown as PresentationModel;

    const { wrapper, store } = await mountWith(legacy);
    // The upgrade re-emits + re-parses the package (async); poll until done.
    for (let i = 0; i < 100; i++) {
      if (typeof store.getPresentation(fresh.id)?.slideWidthCm === "number") {
        break;
      }
      await new Promise((resolve) => setTimeout(resolve, 10));
      await flushPromises();
    }
    // The stored model is replaced in place with the upgraded (frame) model.
    const upgraded = store.getPresentation(fresh.id);
    expect(upgraded && typeof upgraded.slideWidthCm).toBe("number");
    expect(wrapper.findAll('[data-testid="filmstrip-slide"]')).toHaveLength(4);
  });
});

const options = (wrapper: Awaited<ReturnType<typeof mountWith>>["wrapper"]) =>
  wrapper.findAll('[data-testid="filmstrip-slide"]');

describe("PresentationView slide management (SPEC-080)", () => {
  it("adds a slide via the Add slide button (AC-16/25)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation());
    expect(options(wrapper)).toHaveLength(1);
    await wrapper.get('[data-testid="presentation-add-slide"]').trigger("click");
    await flushPromises();
    expect(options(wrapper)).toHaveLength(2);
  });

  it("undoes and redoes via the toolbar buttons, which enable/disable (AC-40/41/42)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation());
    expect(wrapper.get('[data-testid="presentation-undo"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="presentation-add-slide"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="presentation-undo"]').attributes("disabled")).toBeUndefined();

    await wrapper.get('[data-testid="presentation-undo"]').trigger("click");
    await flushPromises();
    expect(options(wrapper)).toHaveLength(1);
    await wrapper.get('[data-testid="presentation-redo"]').trigger("click");
    await flushPromises();
    expect(options(wrapper)).toHaveLength(2);
  });

  it("undoes and redoes via Cmd/Ctrl+Z and Shift+Cmd/Ctrl+Z (AC-40/41)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation());
    await wrapper.get('[data-testid="presentation-add-slide"]').trigger("click");
    await flushPromises();
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true }));
    await flushPromises();
    expect(options(wrapper)).toHaveLength(1);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, shiftKey: true }));
    await flushPromises();
    expect(options(wrapper)).toHaveLength(2);
    wrapper.unmount();
  });

  it("opens the context menu with Shift+F10 and deletes a slide (AC-24/21)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation());
    await wrapper.get('[data-testid="presentation-add-slide"]').trigger("click");
    await flushPromises();
    // Shift+F10 on the focusable listbox opens the current thumbnail's menu.
    await wrapper.get('[data-testid="slide-filmstrip"]').trigger("keydown", { key: "F10", shiftKey: true });
    expect(wrapper.find('[data-testid="slide-context-menu"]').exists()).toBe(true);
    await wrapper.get('[data-testid="slide-menu-delete"]').trigger("click");
    await flushPromises();
    expect(options(wrapper)).toHaveLength(1);
  });

  it("shows an empty state with an Add slide action at zero slides (AC-35)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation());
    await wrapper.get('[data-testid="presentation-slide-menu"]').trigger("click");
    await wrapper.get('[data-testid="slide-menu-delete"]').trigger("click");
    await flushPromises();
    expect(options(wrapper)).toHaveLength(0);
    expect(wrapper.find('[data-testid="presentation-empty"]').exists()).toBe(true);
    await wrapper.get('[data-testid="presentation-empty-add"]').trigger("click");
    await flushPromises();
    expect(options(wrapper)).toHaveLength(1);
  });

  it("offers 'Download a copy as…' the other format (AC-13)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation()); // pptx-born
    expect(wrapper.find('[data-testid="presentation-download-menu-popup"]').exists()).toBe(false);
    await wrapper.get('[data-testid="presentation-download-menu"]').trigger("click");
    expect(wrapper.get('[data-testid="download-copy-as"]').text()).toContain(".odp");
  });

  it("persists mutations write-through so a reload shows them (AC-46/47)", async () => {
    const { wrapper, store } = await mountWith(createBlankPresentation());
    const id = store.listDocuments()[0].id;
    const before = store.getStoredRecord(id)?.lastModified ?? 0;
    await wrapper.get('[data-testid="presentation-add-slide"]').trigger("click");
    await flushPromises();
    expect(store.getPresentation(id)?.slides).toHaveLength(2);
    expect(store.getStoredRecord(id)?.lastModified).toBeGreaterThanOrEqual(before);
  });
});

// Scope to the main stage — the filmstrip thumbnails render the same frames
// read-only (no selection/editing wiring).
const stage = (wrapper: Awaited<ReturnType<typeof mountWith>>["wrapper"]) =>
  wrapper.get('[data-testid="stage-workspace"]');

describe("PresentationView text editing (SPEC-081)", () => {
  // SPEC-082 unified every frame kind under data-testid="slide-frame"; a
  // blank deck's frames are all text placeholders, so this still returns them.
  const frames = (wrapper: Awaited<ReturnType<typeof mountWith>>["wrapper"]) =>
    stage(wrapper).findAll('[data-testid="slide-frame"]');

  it("shows an empty-placeholder prompt by role, outside editing (AC-15)", async () => {
    const { wrapper } = await mountWith(createBlankPresentation());
    const prompts = stage(wrapper).findAll('[data-testid="slide-placeholder-prompt"]');
    expect(prompts).toHaveLength(2);
    expect(prompts[0].text()).toBe("Click to add title");
    expect(prompts[1].text()).toBe("Click to add text");
  });

  it("selects a frame on click and deselects on a surface click (AC-4/5)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    await frames(wrapper)[0].trigger("mousedown");
    expect(frames(wrapper)[0].classes()).toContain("is-selected");
    // SPEC-096: a click on empty stage deselects on RELEASE — the press
    // starts the marquee gesture, and a release still under the drag
    // threshold is the plain deselecting click (SPEC-082's behaviour kept).
    await stage(wrapper).get(".stage-surface").trigger("pointerdown", { clientX: 5, clientY: 5 });
    await stage(wrapper).get(".stage-surface").trigger("mousedown");
    window.dispatchEvent(new MouseEvent("pointerup", { clientX: 5, clientY: 5 }));
    await flushPromises();
    expect(frames(wrapper)[0].classes()).not.toContain("is-selected");
  });

  it("Tab cycles frame selection in z-order (AC-6)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    await frames(wrapper)[0].trigger("mousedown");
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab" }));
    await flushPromises();
    expect(frames(wrapper)[1].classes()).toContain("is-selected");
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true }));
    await flushPromises();
    expect(frames(wrapper)[0].classes()).toContain("is-selected");
    wrapper.unmount();
  });

  it("double-click enters an editing session; Escape ends it (AC-1/2/3)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    await frames(wrapper)[0].trigger("dblclick");
    await flushPromises();
    expect(wrapper.find('[data-testid="slide-text-edit"]').exists()).toBe(true);
    expect(frames(wrapper)[0].classes()).toContain("is-editing");
    await wrapper.get('[data-testid="slide-text-edit"]').trigger("keydown", { key: "Escape" });
    await flushPromises();
    expect(wrapper.find('[data-testid="slide-text-edit"]').exists()).toBe(false);
  });

  it("the toolbar disables when nothing is targeted and enables on selection (AC-25)", async () => {
    const { wrapper } = await mountWith(await freshDeck());
    expect(wrapper.get('[data-testid="slides-bold"]').attributes("disabled")).toBeDefined();
    await frames(wrapper)[0].trigger("mousedown");
    expect(wrapper.get('[data-testid="slides-bold"]').attributes("disabled")).toBeUndefined();
  });

  it("applies bold to a selected frame and reflects it, as one undo entry (AC-24/25/29)", async () => {
    const { wrapper, store } = await mountWith(await freshDeck());
    const id = store.listDocuments()[0].id;
    await frames(wrapper)[0].trigger("mousedown");
    expect(wrapper.get('[data-testid="slides-bold"]').attributes("aria-pressed")).toBe("false");
    await wrapper.get('[data-testid="slides-bold"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="slides-bold"]').attributes("aria-pressed")).toBe("true");
    const title = store.getPresentation(id)!.slides[0].frames.find((f) => f.kind === "text" && f.role === "title");
    expect(title && title.kind === "text" && title.paragraphs[0].runs.every((r) => r.bold)).toBe(true);
    expect(store.getPresentation(id)!.slides[0].dirty).toBe(true);
    // One undo entry: undoing reverts the bold.
    await wrapper.get('[data-testid="presentation-undo"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="slides-bold"]').attributes("aria-pressed")).toBe("false");
    wrapper.unmount();
  });
});

// SPEC-104 — the print-and-export entry points reach one dialog: the Download
// menu's "PDF (.pdf)…" / "Images (.png)…" open it on the matching tab, and
// Cmd/Ctrl+P opens it on the PDF tab; the editor still selects a hidden slide.
async function hiddenDeck(): Promise<PresentationModel> {
  const { buildPptxHiddenFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  return readPresentationBytes(await buildPptxHiddenFixtureBytes(), "pptx", "hidden");
}

describe("PresentationView print & export (SPEC-104)", () => {
  it("the Download menu opens the dialog on the PDF or Images tab", async () => {
    const { wrapper } = await mountWith(await hiddenDeck());
    await wrapper.get('[data-testid="presentation-download-menu"]').trigger("click");
    await wrapper.get('[data-testid="presentation-export-pdf"]').trigger("click");
    await flushPromises();
    expect(wrapper.find('[data-testid="deck-print-dialog"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="deck-print-tab-pdf"]').attributes("aria-selected")).toBe("true");
    await wrapper.get('[data-testid="deck-print-cancel"]').trigger("click");
    expect(wrapper.find('[data-testid="deck-print-dialog"]').exists()).toBe(false);
    await wrapper.get('[data-testid="presentation-download-menu"]').trigger("click");
    await wrapper.get('[data-testid="presentation-export-images"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="deck-print-tab-images"]').attributes("aria-selected")).toBe("true");
    wrapper.unmount();
  });

  it("Cmd/Ctrl+P opens the dialog on the PDF tab", async () => {
    const { wrapper } = await mountWith(await hiddenDeck());
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "p", ctrlKey: true, cancelable: true }));
    await flushPromises();
    expect(wrapper.find('[data-testid="deck-print-dialog"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="deck-print-tab-pdf"]').attributes("aria-selected")).toBe("true");
    wrapper.unmount();
  });

  it("marks the hidden thumbnail and still selects it — PageDown reaches it too", async () => {
    const { wrapper } = await mountWith(await hiddenDeck());
    const thumbs = wrapper.findAll('[data-testid="filmstrip-slide"]');
    expect(thumbs[1].classes()).toContain("filmstrip-hidden");
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "PageDown" }));
    await flushPromises();
    expect(wrapper.findAll('[data-testid="filmstrip-slide"]')[1].attributes("aria-selected")).toBe("true");
    wrapper.unmount();
  });
});
