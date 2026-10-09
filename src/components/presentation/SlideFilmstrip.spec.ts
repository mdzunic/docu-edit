// SPEC-080 — the filmstrip's slide-management surface: the Add slide button,
// the per-slide context menu (right-click + ⋯ affordance) and its intents,
// and pointer drag-reorder with a drop indicator (dropping outside cancels).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SlideFilmstrip from "./SlideFilmstrip.vue";
import { ref } from "vue";
import SlideThumb from "./SlideThumb.vue";
import { useSlideSelection } from "../../views/presentation/useSlideSelection";

function slides(n: number) {
  return Array.from({ length: n }, (_, i) => ({ id: `s${i}`, frames: [] }));
}

function mountStrip(n = 3, currentIndex = 0) {
  return mount(SlideFilmstrip, {
    props: { slides: slides(n), widthCm: 33.867, heightCm: 19.05, currentIndex },
  });
}

const dragData = { dataTransfer: { setData: () => {}, effectAllowed: "" } };

describe("SlideFilmstrip management", () => {
  it("emits add from the Add slide button (AC-16)", async () => {
    const wrapper = mountStrip();
    await wrapper.get('[data-testid="presentation-add-slide"]').trigger("click");
    expect(wrapper.emitted("add")).toBeTruthy();
  });

  it("opens the context menu on right-click and emits the chosen command (AC-17/19)", async () => {
    const wrapper = mountStrip();
    await wrapper.findAll('[data-testid="filmstrip-slide"]')[1].trigger("contextmenu");
    expect(wrapper.find('[data-testid="slide-context-menu"]').exists()).toBe(true);
    await wrapper.get('[data-testid="slide-menu-new"]').trigger("click");
    expect(wrapper.emitted("command")?.[0]?.[0]).toEqual({ action: "new", index: 1 });
  });

  it("opens the context menu from the ⋯ affordance (AC-18)", async () => {
    const wrapper = mountStrip();
    await wrapper.findAll('[data-testid="presentation-slide-menu"]')[0].trigger("click");
    expect(wrapper.find('[data-testid="slide-context-menu"]').exists()).toBe(true);
    // Seven items: new / new-with-layout / duplicate / change-layout / delete /
    // move-up / move-down (SPEC-089 added the two layout entries).
    expect(wrapper.findAll('[data-testid="slide-context-menu"] [role="menuitem"]')).toHaveLength(7);
  });

  it("drag-reorders with a visible drop indicator (AC-38)", async () => {
    const wrapper = mountStrip();
    const opts = wrapper.findAll('[data-testid="filmstrip-slide"]');
    await opts[0].trigger("dragstart", dragData);
    await opts[2].trigger("dragover");
    expect(opts[2].classes()).toContain("drop-target");
    await opts[2].trigger("drop");
    expect(wrapper.emitted("reorder")?.[0]?.[0]).toEqual({ from: 0, to: 2 });
  });

  it("cancels a reorder when the drag ends outside a slide (AC-39)", async () => {
    const wrapper = mountStrip();
    const opts = wrapper.findAll('[data-testid="filmstrip-slide"]');
    await opts[0].trigger("dragstart", dragData);
    await opts[0].trigger("dragend");
    expect(wrapper.emitted("reorder")).toBeFalsy();
  });
});

describe("SlideFilmstrip — SPEC-099 shared selection", () => {
  function harnessed(n = 4) {
    const currentIndex = ref(0);
    const selection = useSlideSelection({
      count: () => n,
      currentIndex,
      select: (index) => {
        currentIndex.value = index;
      },
    });
    const wrapper = mount(SlideFilmstrip, {
      props: { slides: slides(n), widthCm: 33.867, heightCm: 19.05, currentIndex: 0, selection, selectedIndexes: [0] },
    });
    return { wrapper, selection, currentIndex };
  }

  it("renders every thumbnail through the shared SlideThumb", () => {
    const wrapper = mount(SlideFilmstrip, { props: { slides: slides(3), widthCm: 33.867, heightCm: 19.05, currentIndex: 0 } });
    expect(wrapper.findAllComponents(SlideThumb)).toHaveLength(3);
  });

  it("marks every selected index aria-selected (selectedIndexes)", () => {
    const wrapper = mount(SlideFilmstrip, {
      props: { slides: slides(4), widthCm: 33.867, heightCm: 19.05, currentIndex: 1, selectedIndexes: [1, 2] },
    });
    const opts = wrapper.findAll('[data-testid="filmstrip-slide"]');
    expect(opts.map((opt) => opt.attributes("aria-selected"))).toEqual(["false", "true", "true", "false"]);
  });

  it("routes clicks and the selection keys through the shared selection", async () => {
    const { wrapper, selection, currentIndex } = harnessed();
    expect(wrapper.get('[role="listbox"]').attributes("aria-multiselectable")).toBe("true");
    const opts = wrapper.findAll('[data-testid="filmstrip-slide"]');
    await opts[1].trigger("click");
    await opts[3].trigger("click", { shiftKey: true });
    expect(selection.indexes.value).toEqual([1, 2, 3]);
    expect(currentIndex.value).toBe(3);
    await wrapper.get('[role="listbox"]').trigger("keydown", { key: "Escape" });
    expect(selection.indexes.value).toEqual([3]);
    await wrapper.get('[role="listbox"]').trigger("keydown", { key: "a", ctrlKey: true });
    expect(selection.indexes.value).toEqual([0, 1, 2, 3]);
    await wrapper.get('[role="listbox"]').trigger("keydown", { key: "Home" });
    expect(currentIndex.value).toBe(0);
    // The keyboard path no longer emits select itself — the selection did it.
    expect(wrapper.emitted("select")).toBeFalsy();
  });

  it("keeps the context menu and reorder emit unchanged", async () => {
    const { wrapper } = harnessed();
    await wrapper.findAll('[data-testid="filmstrip-slide"]')[2].trigger("contextmenu");
    await wrapper.get('[data-testid="slide-menu-delete"]').trigger("click");
    expect(wrapper.emitted("command")?.[0]?.[0]).toEqual({ action: "delete", index: 2 });
  });
});
