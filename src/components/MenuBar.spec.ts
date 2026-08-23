import { describe, it, expect, afterEach } from "vitest";
import { mount } from "@vue/test-utils";
import { h } from "vue";
import MenuBar from "./MenuBar.vue";
import { setTierWidthForTest } from "../composables/useResponsiveTier";

const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
];

// A minimal command button that MenuBar will find by role, mirroring what
// MenuItem/MenuCheckItem render.
function item(testid: string, opts: { checkbox?: boolean; disabled?: boolean; keepOpen?: boolean } = {}) {
  return h("button", {
    type: "button",
    role: opts.checkbox ? "menuitemcheckbox" : "menuitem",
    "data-testid": testid,
    disabled: opts.disabled || undefined,
    tabindex: "-1",
    ...(opts.keepOpen ? { "data-menu-keep-open": "" } : {}),
  });
}

function mountBar(slots?: Record<string, () => unknown>) {
  return mount(MenuBar, {
    props: { menus: MENUS },
    slots: slots ?? {
      file: () => [item("file-a"), item("file-b")],
      edit: () => [item("edit-a")],
      view: () => [item("view-a", { checkbox: true })],
    },
    attachTo: document.body,
  });
}

describe("MenuBar (SPEC-049)", () => {
  it("exposes a role=menubar with a menuitem title per menu, in order (AC-1, AC-10)", () => {
    const wrapper = mountBar();
    const bar = wrapper.get('[data-testid="menu-bar"]');
    expect(bar.attributes("role")).toBe("menubar");
    const titles = wrapper.findAll("[data-menu-title]");
    expect(titles.map((t) => t.text())).toEqual(["File", "Edit", "View"]);
    for (const title of titles) {
      expect(title.attributes("role")).toBe("menuitem");
      expect(title.attributes("aria-haspopup")).toBe("menu");
      expect(title.attributes("aria-expanded")).toBe("false");
    }
    expect(titles[0].attributes("aria-controls")).toBe("menu-file");
  });

  it("is a single tab stop: exactly one title is tabbable (AC-8)", () => {
    const wrapper = mountBar();
    const tabbable = wrapper.findAll("[data-menu-title]").filter((t) => t.attributes("tabindex") === "0");
    expect(tabbable).toHaveLength(1);
    expect(tabbable[0].text()).toBe("File");
  });

  it("opens a menu on title click and reports the expanded state (AC-10)", async () => {
    const wrapper = mountBar();
    const fileTitle = wrapper.get('[data-testid="menu-title-file"]');
    await fileTitle.trigger("click");
    expect(fileTitle.attributes("aria-expanded")).toBe("true");
    const popup = wrapper.get('[data-testid="menu-file"]');
    expect(popup.attributes("role")).toBe("menu");
    // A second click closes it.
    await fileTitle.trigger("click");
    expect(fileTitle.attributes("aria-expanded")).toBe("false");
  });

  it("Enter on a title opens the menu and focuses the first item (AC-8)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-title-file"]').trigger("keydown", { key: "Enter" });
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("true");
    await wrapper.vm.$nextTick();
    expect(document.activeElement?.getAttribute("data-testid")).toBe("file-a");
    wrapper.unmount();
  });

  it("ArrowRight/ArrowLeft move the tab stop between menus (AC-8)", async () => {
    const wrapper = mountBar();
    const titles = wrapper.findAll("[data-menu-title]");
    await titles[0].trigger("keydown", { key: "ArrowRight" });
    expect(wrapper.findAll("[data-menu-title]")[1].attributes("tabindex")).toBe("0");
    await titles[1].trigger("keydown", { key: "ArrowLeft" });
    expect(wrapper.findAll("[data-menu-title]")[0].attributes("tabindex")).toBe("0");
    // Wraps around.
    await titles[0].trigger("keydown", { key: "ArrowLeft" });
    expect(wrapper.findAll("[data-menu-title]")[2].attributes("tabindex")).toBe("0");
  });

  it("Escape closes an open menu and returns focus to the title (AC-8)", async () => {
    const wrapper = mountBar();
    const fileTitle = wrapper.get('[data-testid="menu-title-file"]');
    await fileTitle.trigger("keydown", { key: "Enter" });
    await wrapper.vm.$nextTick();
    await wrapper.get('[data-testid="menu-file"]').trigger("keydown", { key: "Escape" });
    expect(fileTitle.attributes("aria-expanded")).toBe("false");
    await wrapper.vm.$nextTick();
    expect(document.activeElement?.getAttribute("data-testid")).toBe("menu-title-file");
    wrapper.unmount();
  });

  it("Escape closes a mouse-opened menu even when focus stayed on the page (SPEC-058 AC-39)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("true");
    // The title prevents mousedown focus, so a pointer open leaves focus
    // wherever it was — only UiMenuSurface's document listener can see this.
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    await wrapper.vm.$nextTick();
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("false");
    wrapper.unmount();
  });

  it("a pointerdown outside the bar closes the open menu (SPEC-058 AC-39)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
    document.body.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    await wrapper.vm.$nextTick();
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("false");
    wrapper.unmount();
  });

  it("clicking the open menu's own title closes it — the full pointer sequence, not just click", async () => {
    // The title is the surface's anchor: its pointerdown must NOT dismiss the
    // popup, or the click that follows would toggle a menu that is already
    // closed and reopen it — a menu that cannot be closed from its own title.
    const wrapper = mountBar();
    const fileTitle = wrapper.get('[data-testid="menu-title-file"]');
    await fileTitle.trigger("click");
    expect(fileTitle.attributes("aria-expanded")).toBe("true");
    await fileTitle.trigger("pointerdown");
    await fileTitle.trigger("click");
    expect(fileTitle.attributes("aria-expanded")).toBe("false");
    wrapper.unmount();
  });

  it("opens a neighbour on pointer cross once a menu is open (AC-9)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
    // Moving the pointer onto Edit opens it without a click.
    await wrapper.get('[data-testid="menu-title-edit"]').trigger("mouseenter");
    expect(wrapper.get('[data-testid="menu-title-edit"]').attributes("aria-expanded")).toBe("true");
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("false");
  });

  it("does not cross-open when no menu is open yet (AC-9)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-title-edit"]').trigger("mouseenter");
    expect(wrapper.get('[data-testid="menu-title-edit"]').attributes("aria-expanded")).toBe("false");
  });

  it("closes the menu after a command is chosen", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
    await wrapper.get('[data-testid="file-a"]').trigger("click");
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("false");
  });

  it("closes the menu after a plain button command (the table grid picker's cells)", async () => {
    const wrapper = mount(MenuBar, {
      props: { menus: MENUS },
      slots: {
        file: () => [
          h("div", { class: "grid" }, [
            h("button", { type: "button", "data-testid": "grid-cell", tabindex: "-1" }),
          ]),
        ],
      },
      attachTo: document.body,
    });
    await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
    await wrapper.get('[data-testid="grid-cell"]').trigger("click");
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("false");
    wrapper.unmount();
  });

  it("keeps the menu open for a data-menu-keep-open command", async () => {
    const wrapper = mount(MenuBar, {
      props: { menus: MENUS },
      slots: {
        file: () => [item("file-keep", { keepOpen: true }), item("file-plain")],
      },
      attachTo: document.body,
    });
    await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
    await wrapper.get('[data-testid="file-keep"]').trigger("click");
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("true");
    wrapper.unmount();
  });

  it("skips disabled items when moving focus within a menu (AC-7)", async () => {
    const wrapper = mount(MenuBar, {
      props: { menus: MENUS },
      slots: {
        file: () => [item("file-a"), item("file-disabled", { disabled: true }), item("file-c")],
      },
      attachTo: document.body,
    });
    await wrapper.get('[data-testid="menu-title-file"]').trigger("keydown", { key: "Enter" });
    await wrapper.vm.$nextTick();
    expect(document.activeElement?.getAttribute("data-testid")).toBe("file-a");
    await wrapper.get('[data-testid="menu-file"]').trigger("keydown", { key: "ArrowDown" });
    // The disabled item is skipped.
    expect(document.activeElement?.getAttribute("data-testid")).toBe("file-c");
    wrapper.unmount();
  });
});

// SPEC-068 AC-7/8 — at the compact tier the bar is one "Menu" button opening a
// full-height sheet with every menu's commands. The sheet is teleported to
// body, so queries run against the document.
describe("MenuBar compact tier (SPEC-068)", () => {
  afterEach(() => {
    setTierWidthForTest(1440);
    document.querySelectorAll('[data-testid="menu-sheet"]').forEach((el) => el.remove());
  });

  it("collapses to a single Menu button (AC-7) that opens a sheet with all six menus (AC-8)", async () => {
    setTierWidthForTest(390);
    const wrapper = mountBar();
    // No docked menubar at compact; just the Menu button.
    expect(wrapper.find('[role="menubar"]').exists()).toBe(false);
    const button = wrapper.get('[data-testid="menu-button"]');
    expect(button.text()).toBe("Menu");
    expect(button.attributes("aria-expanded")).toBe("false");
    expect(document.querySelector('[data-testid="menu-sheet"]')).toBeNull();

    await button.trigger("click");
    expect(button.attributes("aria-expanded")).toBe("true");
    const sheet = document.querySelector('[data-testid="menu-sheet"]');
    expect(sheet).not.toBeNull();
    // Every menu's slot content appears, under its label, in bar order.
    expect(sheet?.querySelector('[data-testid="menu-sheet-section-file"]')).not.toBeNull();
    expect(sheet?.querySelector('[data-testid="menu-sheet-section-edit"]')).not.toBeNull();
    expect(sheet?.querySelector('[data-testid="menu-sheet-section-view"]')).not.toBeNull();
    expect(sheet?.querySelector('[data-testid="file-a"]')).not.toBeNull();
    wrapper.unmount();
  });

  it("closes the sheet when a command is tapped", async () => {
    setTierWidthForTest(390);
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-button"]').trigger("click");
    const item = document.querySelector('[data-testid="file-a"]');
    (item as HTMLElement).click();
    await wrapper.vm.$nextTick();
    expect(document.querySelector('[data-testid="menu-sheet"]')).toBeNull();
    wrapper.unmount();
  });

  it("dismisses the sheet when the tier widens past compact", async () => {
    setTierWidthForTest(390);
    const wrapper = mountBar();
    await wrapper.get('[data-testid="menu-button"]').trigger("click");
    expect(document.querySelector('[data-testid="menu-sheet"]')).not.toBeNull();
    setTierWidthForTest(1200);
    await wrapper.vm.$nextTick();
    expect(document.querySelector('[data-testid="menu-sheet"]')).toBeNull();
    // And the docked menu bar is back.
    expect(wrapper.find('[role="menubar"]').exists()).toBe(true);
    wrapper.unmount();
  });
});
