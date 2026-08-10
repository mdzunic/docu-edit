import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import { h } from "vue";
import MenuBar from "./MenuBar.vue";

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
