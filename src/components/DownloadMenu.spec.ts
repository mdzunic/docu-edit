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

  it("emits download('docx') and closes when the option is chosen (AC-2)", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    await wrapper.get('[data-testid="download-docx"]').trigger("click");
    expect(wrapper.emitted("download")).toEqual([["docx"]]);
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
  });

  it("Escape closes the menu without emitting", async () => {
    const wrapper = mount(DownloadMenu);
    await wrapper.get('[data-testid="download-menu"]').trigger("click");
    await wrapper.get('[data-testid="download-menu"]').trigger("keydown.escape");
    expect(wrapper.find('[data-testid="download-menu-list"]').exists()).toBe(false);
    expect(wrapper.emitted("download")).toBeUndefined();
  });
});
