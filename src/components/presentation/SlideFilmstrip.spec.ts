// SPEC-080 — the filmstrip's slide-management surface: the Add slide button,
// the per-slide context menu (right-click + ⋯ affordance) and its intents,
// and pointer drag-reorder with a drop indicator (dropping outside cancels).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SlideFilmstrip from "./SlideFilmstrip.vue";

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
    expect(wrapper.findAll('[data-testid="slide-context-menu"] [role="menuitem"]')).toHaveLength(5);
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
