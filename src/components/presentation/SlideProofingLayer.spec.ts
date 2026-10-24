// @vitest-environment jsdom
// SPEC-098 — the stage's proofing layer: find highlights under SPEC-017's
// testids (`.current` on the current hit), squiggles under SPEC-035's, and a
// right-click on a squiggle emitted as `spell-context` — hit-tested on the
// surface, since every rect is inert.
import { mount } from "@vue/test-utils";
import { afterEach, describe, expect, it } from "vitest";
import SlideProofingLayer from "./SlideProofingLayer.vue";

/** The layer inside a stand-in surface (unscaled: client px = surface px). */
function mountLayer(props: Record<string, unknown>) {
  const surface = document.createElement("div");
  document.body.appendChild(surface);
  const wrapper = mount(SlideProofingLayer, { props, attachTo: surface });
  // The layer listens on its parent (the stage surface in the app).
  return { wrapper, surface: wrapper.element.parentElement as HTMLElement };
}

afterEach(() => {
  document.body.innerHTML = "";
});

describe("SlideProofingLayer", () => {
  it("paints find highlights with .current on the current hit", () => {
    const { wrapper } = mountLayer({
      highlightRects: [
        { left: 10, top: 20, width: 30, height: 12, current: false },
        { left: 50, top: 20, width: 30, height: 12, current: true },
      ],
    });
    const rects = wrapper.findAll('[data-testid="find-highlights"] [data-testid="find-highlight"]');
    expect(rects).toHaveLength(2);
    expect(rects[0].classes()).not.toContain("current");
    expect(rects[1].classes()).toContain("current");
    expect(rects[1].attributes("style")).toContain("left: 50px");
    expect(wrapper.find('[data-testid="spell-underlines"]').exists()).toBe(false);
  });

  it("paints squiggles that keep a constant on-screen size at any stage scale", () => {
    const { wrapper } = mountLayer({ underlineRects: [{ left: 5, top: 6, width: 40, height: 14, issue: 0, word: "teh" }], scale: 0.5 });
    const underline = wrapper.get('[data-testid="spell-underlines"] [data-testid="spell-underline"]');
    expect(underline.attributes("data-word")).toBe("teh");
    expect(underline.attributes("style")).toContain("background-size: 12px 8px");
  });

  it("emits spell-context for a right-click on a squiggle and leaves other right-clicks alone", () => {
    const { wrapper, surface } = mountLayer({ underlineRects: [{ left: 0, top: 0, width: 40, height: 14, issue: 3, word: "teh" }] });
    const onWord = new MouseEvent("contextmenu", { clientX: 20, clientY: 7, bubbles: true, cancelable: true });
    surface.dispatchEvent(onWord);
    expect(onWord.defaultPrevented).toBe(true);
    expect(wrapper.emitted("spell-context")).toEqual([[3, 20, 7]]);
    const elsewhere = new MouseEvent("contextmenu", { clientX: 200, clientY: 200, bubbles: true, cancelable: true });
    surface.dispatchEvent(elsewhere);
    expect(elsewhere.defaultPrevented).toBe(false);
    expect(wrapper.emitted("spell-context")).toHaveLength(1);
  });

  it("keeps a right press on a squiggle from reaching the stage below (a session stays open)", () => {
    const { surface } = mountLayer({ underlineRects: [{ left: 0, top: 0, width: 40, height: 14, issue: 0, word: "teh" }] });
    const text = document.createElement("span");
    surface.appendChild(text);
    let reached = 0;
    text.addEventListener("mousedown", () => (reached += 1));
    text.dispatchEvent(new MouseEvent("mousedown", { button: 2, clientX: 5, clientY: 5, bubbles: true }));
    expect(reached).toBe(0);
    text.dispatchEvent(new MouseEvent("mousedown", { button: 0, clientX: 5, clientY: 5, bubbles: true }));
    expect(reached).toBe(1);
  });
});
