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

describe("PresentationView slide management (SPEC-080)", () => {
  const options = (wrapper: Awaited<ReturnType<typeof mountWith>>["wrapper"]) =>
    wrapper.findAll('[data-testid="filmstrip-slide"]');

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
