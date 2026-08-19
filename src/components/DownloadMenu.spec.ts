import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import DownloadMenu from "./DownloadMenu.vue";

describe("DownloadMenu", () => {
  it("renders a Download button that starts with the menu closed (AC-1)", () => {
    const wrapper = mount(DownloadMenu);
    const button = wrapper.get('[data-testid="download-menu"]');
    expect(button.text()).toBe("Download");
    expect(button.attributes("aria-haspopup")).toBe("menu");
    expect(button.attributes("aria-expanded")).toBe("false");
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
  });

  it("opens the menu with a 'Word (.docx)' option (AC-1)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    const menu = wrapper.get('[data-testid="download-menu-list"]');
    expect(menu.attributes("role")).toBe("menu");
    const item = wrapper.get('[data-testid="download-docx"]');
    expect(item.text()).toBe("Word (.docx)");
    expect(item.attributes("role")).toBe("menuitem");
  });

  it("opens the menu with an 'OpenDocument Text (.odt)' option after the .docx one (SPEC-007 AC-1)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    const item = wrapper.get('[data-testid="download-odt"]');
    expect(item.text()).toBe("OpenDocument Text (.odt)");
    expect(item.attributes("role")).toBe("menuitem");
    const items = wrapper.findAll('[role="menuitem"]').map((option) => option.text());
    // "Markdown (.md)" appended by SPEC-050; "PDF (.pdf)" by SPEC-034 AC-14.
    expect(items).toEqual(["Word (.docx)", "OpenDocument Text (.odt)", "Markdown (.md)", "PDF (.pdf)"]);
  });

  it("opens the menu with a 'Markdown (.md)' option that emits download('md') (SPEC-050)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    const item = wrapper.get('[data-testid="download-md"]');
    expect(item.text()).toBe("Markdown (.md)");
    expect(item.attributes("role")).toBe("menuitem");
    await item.trigger("click");
    expect(wrapper.emitted("download")).toEqual([["md"]]);
  });

  it("opens the menu with a 'PDF (.pdf)' option that emits download('pdf') (SPEC-034 AC-14)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    const item = wrapper.get('[data-testid="download-pdf"]');
    expect(item.text()).toBe("PDF (.pdf)");
    expect(item.attributes("role")).toBe("menuitem");
    await item.trigger("click");
    expect(wrapper.emitted("download")).toEqual([["pdf"]]);
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
  });

  it("emits download('odt') and closes when the .odt option is chosen (SPEC-007 AC-2)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    await wrapper.get('[data-testid="download-odt"]').trigger("click");
    expect(wrapper.emitted("download")).toEqual([["odt"]]);
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
  });

  it("emits download('docx') and closes when the option is chosen (AC-2)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    await wrapper.get('[data-testid="download-docx"]').trigger("click");
    expect(wrapper.emitted("download")).toEqual([["docx"]]);
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
  });

  it("Escape closes the menu without emitting", async () => {
    // Dismissal now lives in UiMenuSurface, which listens on the document (so
    // a mouse-opened popup still closes on Escape even when focus stayed on the
    // page — SPEC-058 AC-39). Attach to the document and fire a real Escape
    // there, the way the browser delivers it.
    const wrapper = mount(DownloadMenu, { attachTo: document.body });
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(true);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    await wrapper.vm.$nextTick();
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
    expect(wrapper.emitted("download")).toBeUndefined();
    wrapper.unmount();
  });
});
