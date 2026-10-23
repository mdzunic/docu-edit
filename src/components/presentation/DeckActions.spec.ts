// SPEC-100 — the deck app bar's Present control gains a menu beside it,
// PresentationDownloadButton's split-button pattern: the ▾ opens a menu of the
// show and "Presenter view", `present` is emitted with the mode, and the menu
// is keyboard-operable (ArrowDown opens it on its first item, arrows/Home/End
// move, Escape closes it back onto the caret without reaching the editor's
// window-level shortcuts). The plain Present button still emits the show.
import { afterEach, describe, expect, it, vi } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { nextTick } from "vue";
import DeckActions from "./DeckActions.vue";

const wrappers: VueWrapper[] = [];
afterEach(() => {
  for (const wrapper of wrappers.splice(0)) wrapper.unmount();
});

function mountActions(canPresent = true) {
  const wrapper = mount(DeckActions, { props: { canPresent }, attachTo: document.body });
  wrappers.push(wrapper);
  return wrapper;
}

const MENU = '[data-testid="presentation-present-menu-popup"]';

describe("DeckActions Present menu", () => {
  it("Present still starts the show", async () => {
    const wrapper = mountActions();
    await wrapper.get('[data-testid="slides-present"]').trigger("click");
    expect(wrapper.emitted("present")).toEqual([["show"]]);
  });

  it("the ▾ opens a menu offering the show and Presenter view, each emitting its mode", async () => {
    const wrapper = mountActions();
    const caret = wrapper.get('[data-testid="presentation-present-menu"]');
    expect(caret.attributes("aria-haspopup")).toBe("menu");
    expect(caret.attributes("aria-expanded")).toBe("false");
    await caret.trigger("click");
    expect(caret.attributes("aria-expanded")).toBe("true");
    const menu = wrapper.get(MENU);
    expect(menu.attributes("role")).toBe("menu");
    const items = menu.findAll('[role="menuitem"]');
    expect(items.map((item) => item.text())).toEqual(["From current slide", "Presenter view"]);
    // The spec's testid and the work order's both resolve to the item.
    expect(wrapper.get('[data-testid="presentation-present-presenter"]').text()).toBe("Presenter view");
    await wrapper.get('[data-testid="presentation-presenter"]').trigger("click");
    expect(wrapper.emitted("present")).toEqual([["presenter"]]);
    expect(wrapper.find(MENU).exists()).toBe(false);
    await caret.trigger("click");
    await wrapper.get('[data-testid="presentation-present-show"]').trigger("click");
    expect(wrapper.emitted("present")).toEqual([["presenter"], ["show"]]);
  });

  it("is keyboard-operable: ArrowDown opens on the first item, arrows/Home/End move, Escape closes onto the caret", async () => {
    const wrapper = mountActions();
    const caret = wrapper.get('[data-testid="presentation-present-menu"]');
    await caret.trigger("keydown", { key: "ArrowDown" });
    await nextTick();
    const items = () => wrapper.findAll('[role="menuitem"]').map((item) => item.element);
    expect(document.activeElement).toBe(items()[0]);
    const menu = wrapper.get(MENU);
    await menu.trigger("keydown", { key: "ArrowDown" });
    expect(document.activeElement).toBe(items()[1]);
    await menu.trigger("keydown", { key: "ArrowDown" }); // wraps
    expect(document.activeElement).toBe(items()[0]);
    await menu.trigger("keydown", { key: "End" });
    expect(document.activeElement).toBe(items()[1]);
    await menu.trigger("keydown", { key: "Home" });
    expect(document.activeElement).toBe(items()[0]);
    await menu.trigger("keydown", { key: "ArrowUp" }); // wraps back
    expect(document.activeElement).toBe(items()[1]);
    const windowKeys = vi.fn();
    window.addEventListener("keydown", windowKeys);
    await menu.trigger("keydown", { key: "Escape" });
    window.removeEventListener("keydown", windowKeys);
    expect(windowKeys).not.toHaveBeenCalled(); // the editor's shortcuts never saw it
    expect(wrapper.find(MENU).exists()).toBe(false);
    expect(document.activeElement).toBe(caret.element);
  });

  it("Tab closes the menu; the backdrop click closes it too", async () => {
    const wrapper = mountActions();
    const caret = wrapper.get('[data-testid="presentation-present-menu"]');
    await caret.trigger("click");
    await wrapper.get(MENU).trigger("keydown", { key: "Tab" });
    expect(wrapper.find(MENU).exists()).toBe(false);
    await caret.trigger("click");
    await wrapper.get(".menu-backdrop").trigger("click");
    expect(wrapper.find(MENU).exists()).toBe(false);
    expect(wrapper.emitted("present")).toBeUndefined();
  });

  it("the menu is disabled with Present on an empty deck", () => {
    const wrapper = mountActions(false);
    expect(wrapper.get('[data-testid="slides-present"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="presentation-present-menu"]').attributes("disabled")).toBeDefined();
  });
});
