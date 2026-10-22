import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import { h } from "vue";
import UiMenuSurface from "./UiMenuSurface.vue";

// The dismissal listeners attach when `open` flips true, so every test mounts
// closed and opens with setProps — the same order every consumer follows.
function mountSurface(props: Record<string, unknown> = {}) {
  return mount(UiMenuSurface, {
    props: { open: false, ...props },
    slots: {
      default: () => h("button", { type: "button", "data-testid": "inside" }, "item"),
    },
    attachTo: document.body,
  });
}

// jsdom has no PointerEvent constructor; listeners match on the type string,
// so a MouseEvent named "pointerdown" is what these tests (and happy paths in
// other suites) deliver.
function pointerdownOn(el: Element) {
  el.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
}

describe("UiMenuSurface (SPEC-058)", () => {
  it("dismisses on a pointerdown outside the surface", async () => {
    const wrapper = mountSurface();
    await wrapper.setProps({ open: true });
    pointerdownOn(document.body);
    expect(wrapper.emitted("dismiss")).toHaveLength(1);
    wrapper.unmount();
  });

  it("ignores a pointerdown inside the surface", async () => {
    const wrapper = mountSurface();
    await wrapper.setProps({ open: true });
    pointerdownOn(wrapper.get('[data-testid="inside"]').element);
    expect(wrapper.emitted("dismiss")).toBeUndefined();
    wrapper.unmount();
  });

  it("ignores a pointerdown on the anchor, leaving the toggle to the trigger's own click", async () => {
    const anchor = document.createElement("button");
    document.body.appendChild(anchor);
    const wrapper = mountSurface({ anchor });
    await wrapper.setProps({ open: true });
    pointerdownOn(anchor);
    expect(wrapper.emitted("dismiss")).toBeUndefined();
    wrapper.unmount();
    anchor.remove();
  });

  it("ignores a pointerdown inside another open surface (a child context menu is a DOM sibling)", async () => {
    // StylePicker's shape: the style list and the per-entry context menu are
    // two sibling surfaces, both open. Dismissing the list on a pointerdown in
    // the context menu unmounts the context menu before its item's click can
    // fire — the SPEC-019 update/reset-style regressions.
    const list = mountSurface();
    await list.setProps({ open: true });
    const child = mountSurface();
    await child.setProps({ open: true });
    pointerdownOn(child.get('[data-testid="inside"]').element);
    expect(list.emitted("dismiss")).toBeUndefined();
    expect(child.emitted("dismiss")).toBeUndefined();
    child.unmount();
    list.unmount();
  });

  it("keepMounted keeps the closed surface in the DOM, hidden, and dismissal still works open", async () => {
    // MenuBar's contract: SPEC-020 reads the format painter's aria-checked
    // off a closed Format menu, so closed popups must stay queryable.
    const wrapper = mountSurface({ keepMounted: true });
    expect(wrapper.get('[data-testid="inside"]').isVisible()).toBe(false);
    await wrapper.setProps({ open: true });
    expect(wrapper.get('[data-testid="inside"]').isVisible()).toBe(true);
    pointerdownOn(document.body);
    expect(wrapper.emitted("dismiss")).toHaveLength(1);
    wrapper.unmount();
  });

  it("dismisses on a document-level Escape (mouse-opened popups keep focus on the page)", async () => {
    const wrapper = mountSurface();
    await wrapper.setProps({ open: true });
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(wrapper.emitted("dismiss")).toHaveLength(1);
    wrapper.unmount();
  });

  it("returns focus to the anchor on close while focus is still inside", async () => {
    const anchor = document.createElement("button");
    document.body.appendChild(anchor);
    const wrapper = mountSurface({ anchor });
    await wrapper.setProps({ open: true });
    (wrapper.get('[data-testid="inside"]').element as HTMLElement).focus();
    await wrapper.setProps({ open: false });
    expect(document.activeElement).toBe(anchor);
    wrapper.unmount();
    anchor.remove();
  });

  it("never moves focus with returnFocus=false (consumers with their own focus discipline)", async () => {
    const anchor = document.createElement("button");
    document.body.appendChild(anchor);
    const wrapper = mountSurface({ anchor, returnFocus: false });
    await wrapper.setProps({ open: true });
    (wrapper.get('[data-testid="inside"]').element as HTMLElement).focus();
    await wrapper.setProps({ open: false });
    expect(document.activeElement).not.toBe(anchor);
    wrapper.unmount();
    anchor.remove();
  });
});
