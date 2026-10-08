// SPEC-099 — the slide sorter component: the listbox of cards under section
// headers, the shared multi-selection, section rename/collapse/menu, the card
// menu reusing the filmstrip's slide-menu items, one drop indicator while
// dragging, and the odp note.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h, ref, shallowRef, triggerRef } from "vue";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import type { PresentationModel, Slide, TextFrame } from "../../lib/presentation/model";
import { addSection } from "../../lib/presentation/sections";
import { usePreferences } from "../../composables/usePreferences";
import { useSlideSorter, type SlideSorter as Controller } from "../../views/presentation/useSlideSorter";
import SlideSorter from "./SlideSorter.vue";

function slide(title: string): Slide {
  const frame: TextFrame = {
    id: crypto.randomUUID(),
    kind: "text",
    role: "title",
    xCm: 1,
    yCm: 1,
    widthCm: 20,
    heightCm: 3,
    paragraphs: [{ runs: [{ text: title }] }],
  };
  return { id: crypto.randomUUID(), frames: [frame] };
}

function deck(format: "pptx" | "odp" = "pptx"): PresentationModel {
  const model: PresentationModel = {
    kind: "presentation",
    id: "d",
    title: "Deck",
    format,
    slides: ["A", "B", "C", "D"].map(slide),
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  };
  addSection(model, 0, "Intro");
  addSection(model, 2, "Body");
  delete model.structureDirty;
  return model;
}

let mounted: VueWrapper | undefined;

function mountSorter(model = deck()) {
  const presentation = shallowRef<PresentationModel | null>(model);
  const currentIndex = ref(0);
  const record = vi.fn();
  // The view's commit: a fresh slides array, the same model, triggerRef.
  const commit = vi.fn(() => {
    model.slides = model.slides.slice();
    triggerRef(presentation);
  });
  let controller!: Controller;
  const wrapper = mount(
    defineComponent({
      setup() {
        controller = useSlideSorter({ model: presentation, currentIndex, record, commit, reportError: vi.fn() });
        return () =>
          h(SlideSorter, {
            controller,
            widthCm: model.slideWidthCm,
            heightCm: model.slideHeightCm,
            currentIndex: currentIndex.value,
          });
      },
    }),
    { attachTo: document.body },
  );
  mounted = wrapper;
  return { wrapper, controller, model, record, currentIndex };
}

const { setPreference } = usePreferences();

beforeEach(() => {
  setPreference("presentationSorterOpen", true);
});
afterEach(() => {
  mounted?.unmount();
  mounted = undefined;
  setPreference("presentationSorterOpen", false);
});

describe("SlideSorter", () => {
  it("renders a multiselectable listbox of numbered cards under section headers (AC-3)", () => {
    const { wrapper } = mountSorter();
    expect(wrapper.find('[data-testid="slides-sorter"]').exists()).toBe(true);
    const listbox = wrapper.get('[role="listbox"]');
    expect(listbox.attributes("aria-multiselectable")).toBe("true");
    const cards = wrapper.findAll('[data-testid="sorter-slide"]');
    expect(cards).toHaveLength(4);
    expect(cards.map((card) => card.attributes("aria-selected"))).toEqual(["true", "false", "false", "false"]);
    expect(cards[1].text()).toContain("2");
    expect(wrapper.findAll('[data-testid="sorter-section-name"]').map((name) => name.text())).toEqual(["Intro", "Body"]);
    expect(wrapper.findAll('[data-testid="slide-stage"]')).toHaveLength(4);
    expect(wrapper.get('[data-testid="sorter-live"]').attributes("role")).toBe("status");
  });

  it("click, Shift+click and Cmd/Ctrl+click drive the shared selection (AC-4)", async () => {
    const { wrapper, controller } = mountSorter();
    const cards = () => wrapper.findAll('[data-testid="sorter-slide"]');
    await cards()[1].trigger("click");
    await cards()[3].trigger("click", { shiftKey: true });
    expect(controller.selection.indexes.value).toEqual([1, 2, 3]);
    expect(cards().map((card) => card.attributes("aria-selected"))).toEqual(["false", "true", "true", "true"]);
    expect(wrapper.get('[data-testid="sorter-live"]').text()).toBe("3 slides selected");
    await cards()[2].trigger("click", { ctrlKey: true });
    expect(controller.selection.indexes.value).toEqual([1, 3]);
  });

  it("renames a section: Enter commits a trimmed name, an empty one is rejected, Escape cancels (AC-6)", async () => {
    const { wrapper, model } = mountSorter();
    await wrapper.findAll('[data-testid="sorter-section-name"]')[1].trigger("dblclick");
    const input = wrapper.get('[data-testid="section-rename-input"]');
    await input.setValue("   ");
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.find('[data-testid="section-rename-input"]').exists()).toBe(true);
    expect(input.attributes("aria-invalid")).toBe("true");
    await input.trigger("keydown", { key: "Escape" });
    expect(wrapper.find('[data-testid="section-rename-input"]').exists()).toBe(false);
    expect(model.sections![1].name).toBe("Body");
    await wrapper.findAll('[data-testid="sorter-section-name"]')[1].trigger("dblclick");
    await wrapper.get('[data-testid="section-rename-input"]').setValue("  Results ");
    await wrapper.get('[data-testid="section-rename-input"]').trigger("keydown", { key: "Enter" });
    expect(model.sections![1].name).toBe("Results");
  });

  it("collapsing a section hides its cards and keeps its slide count (AC-6)", async () => {
    const { wrapper } = mountSorter();
    const toggle = wrapper.findAll('[data-testid="sorter-section-toggle"]')[0];
    expect(toggle.attributes("aria-expanded")).toBe("true");
    await toggle.trigger("click");
    expect(wrapper.findAll('[data-testid="sorter-section-toggle"]')[0].attributes("aria-expanded")).toBe("false");
    expect(wrapper.findAll('[data-testid="sorter-slide"]')).toHaveLength(2);
    expect(wrapper.findAll('[data-testid="sorter-section-count"]')[0].text()).toBe("2 slides");
  });

  it("the section menu holds the six section commands; move/remove act on the run (AC-6)", async () => {
    const { wrapper, model } = mountSorter();
    await wrapper.findAll('[data-testid="sorter-section-menu"]')[1].trigger("click");
    const testids = wrapper.findAll('[role="menuitem"]').map((item) => item.attributes("data-testid"));
    expect(testids).toEqual([
      "section-add",
      "section-rename",
      "section-move-up",
      "section-move-down",
      "section-remove",
      "section-remove-with-slides",
    ]);
    await wrapper.get('[data-testid="section-move-up"]').trigger("click");
    expect(model.sections!.map((s) => s.name)).toEqual(["Body", "Intro"]);
    await wrapper.findAll('[data-testid="sorter-section-menu"]')[0].trigger("click");
    await wrapper.get('[data-testid="section-remove-with-slides"]').trigger("click");
    expect(model.sections!.map((s) => s.name)).toEqual(["Intro"]);
    expect(model.slides).toHaveLength(2);
  });

  it("each card's menu reuses the filmstrip's slide-menu items and emits the same command (AC-5)", async () => {
    const { wrapper } = mountSorter();
    await wrapper.findAll('[data-testid="sorter-slide"]')[2].trigger("contextmenu");
    expect(wrapper.findAll('[role="menuitem"]').map((item) => item.attributes("data-testid"))).toEqual([
      "slide-menu-new",
      "slide-menu-new-with-layout",
      "slide-menu-duplicate",
      "slide-menu-layout",
      "slide-menu-delete",
      "slide-menu-move-up",
      "slide-menu-move-down",
    ]);
    await wrapper.get('[data-testid="slide-menu-duplicate"]').trigger("click");
    const sorter = wrapper.findComponent(SlideSorter);
    expect(sorter.emitted("command")?.[0]?.[0]).toEqual({ action: "duplicate", index: 2 });
  });

  it("a drag shows one drop indicator and drops the selection as a block (AC-5)", async () => {
    const { wrapper, model, record } = mountSorter();
    const ids = model.slides.map((s) => s.id);
    const cards = () => wrapper.findAll('[data-testid="sorter-slide"]');
    await cards()[0].trigger("click");
    await cards()[1].trigger("click", { shiftKey: true });
    await cards()[0].trigger("dragstart", { dataTransfer: { setData: () => {}, effectAllowed: "" } });
    await cards()[3].trigger("dragover");
    expect(wrapper.findAll('[data-testid="sorter-drop-indicator"]')).toHaveLength(1);
    await wrapper.get('[role="listbox"]').trigger("drop");
    expect(record).toHaveBeenCalledWith("Move slides");
    expect(model.slides.map((s) => s.id)).toEqual([ids[2], ids[0], ids[1], ids[3]]);
    expect(wrapper.findAll('[data-testid="sorter-drop-indicator"]')).toHaveLength(0);
  });

  it("dropping outside cancels", async () => {
    const { wrapper, record } = mountSorter();
    const cards = () => wrapper.findAll('[data-testid="sorter-slide"]');
    await cards()[0].trigger("dragstart", { dataTransfer: { setData: () => {}, effectAllowed: "" } });
    await cards()[2].trigger("dragover");
    await cards()[0].trigger("dragend");
    expect(wrapper.findAll('[data-testid="sorter-drop-indicator"]')).toHaveLength(0);
    expect(record).not.toHaveBeenCalled();
  });

  it("says once that Impress will not display an odp deck's sections (AC-8)", () => {
    const { wrapper } = mountSorter(deck("odp"));
    expect(wrapper.findAll('[data-testid="sorter-odp-note"]')).toHaveLength(1);
    wrapper.unmount();
    const pptx = mountSorter(deck("pptx"));
    expect(pptx.wrapper.find('[data-testid="sorter-odp-note"]').exists()).toBe(false);
  });
});

describe("SlideSorter — Add section", () => {
  it("opens the new section's name for editing, seeded with its name", async () => {
    const { wrapper, model } = mountSorter();
    await wrapper.findAll('[data-testid="sorter-slide"]')[3].trigger("click");
    await wrapper.get('[data-testid="sorter-add-section"]').trigger("click");
    const input = wrapper.get('[data-testid="section-rename-input"]');
    expect((input.element as HTMLInputElement).value).toBe("New section");
    await input.setValue("Appendix");
    await input.trigger("keydown", { key: "Enter" });
    expect(model.sections!.map((s) => s.name)).toEqual(["Intro", "Body", "Appendix"]);
  });

  // QA: the new section's header row mounts with its input inside it, so a
  // focus() made as the input is created lands on a detached element and the
  // keyboard stays on the button — Enter then added a second section.
  it("moves keyboard focus into the new section's name, selected, from the toolbar button", async () => {
    const { wrapper, model } = mountSorter();
    await wrapper.findAll('[data-testid="sorter-slide"]')[3].trigger("click");
    const button = wrapper.get('[data-testid="sorter-add-section"]');
    (button.element as HTMLButtonElement).focus();
    await button.trigger("click");
    await flushPromises();
    const input = wrapper.get('[data-testid="section-rename-input"]').element as HTMLInputElement;
    expect(document.activeElement).toBe(input);
    expect([input.selectionStart, input.selectionEnd]).toEqual([0, "New section".length]);
    expect(model.sections).toHaveLength(3);
  });

  it("moves keyboard focus into the new section's name from the section menu's Add section", async () => {
    const { wrapper, model } = mountSorter();
    await wrapper.findAll('[data-testid="sorter-slide"]')[3].trigger("click");
    await wrapper.findAll('[data-testid="sorter-section-menu"]')[1].trigger("click");
    await flushPromises();
    const item = wrapper.get('[data-testid="section-add"]');
    expect(document.activeElement).toBe(item.element);
    await item.trigger("click");
    await flushPromises();
    const input = wrapper.get('[data-testid="section-rename-input"]').element as HTMLInputElement;
    expect(document.activeElement).toBe(input);
    expect(model.sections).toHaveLength(3);
  });
});
