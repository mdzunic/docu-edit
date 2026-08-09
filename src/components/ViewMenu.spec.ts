import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import ViewMenu from "./ViewMenu.vue";

describe("ViewMenu", () => {
  it("renders a View button, menu closed by default (AC-13)", () => {
    const wrapper = mount(ViewMenu);
    const button = wrapper.get('[data-testid="view-menu-button"]');
    expect(button.text()).toBe("View ▾");
    expect(button.attributes("aria-haspopup")).toBe("menu");
    expect(button.attributes("aria-expanded")).toBe("false");
    expect(wrapper.find('[data-testid="view-menu"]').exists()).toBe(false);
  });

  it("opens a 'Document outline' checkbox item reflecting the open state", async () => {
    const wrapper = mount(ViewMenu, { props: { outlineOpen: true } });
    await wrapper.get('[data-testid="view-menu-button"]').trigger("click");
    const menu = wrapper.get('[data-testid="view-menu"]');
    expect(menu.attributes("role")).toBe("menu");
    const item = wrapper.get('[data-testid="view-menu-outline"]');
    expect(item.attributes("role")).toBe("menuitemcheckbox");
    expect(item.attributes("aria-checked")).toBe("true");
  });

  it("shows unchecked when the pane is closed", async () => {
    const wrapper = mount(ViewMenu, { props: { outlineOpen: false } });
    await wrapper.get('[data-testid="view-menu-button"]').trigger("click");
    expect(wrapper.get('[data-testid="view-menu-outline"]').attributes("aria-checked")).toBe("false");
  });

  it("emits toggle-outline and closes when the item is chosen (AC-13)", async () => {
    const wrapper = mount(ViewMenu, { props: { outlineOpen: false } });
    await wrapper.get('[data-testid="view-menu-button"]').trigger("click");
    await wrapper.get('[data-testid="view-menu-outline"]').trigger("click");
    expect(wrapper.emitted("toggle-outline")).toEqual([[]]);
    expect(wrapper.find('[data-testid="view-menu"]').exists()).toBe(false);
  });
});
