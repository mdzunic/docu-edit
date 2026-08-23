import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import MenuItem from "./MenuItem.vue";

describe("MenuItem", () => {
  it("shows the accelerator when enabled", () => {
    const wrapper = mount(MenuItem, { props: { label: "Print", shortcut: "Ctrl+P" } });
    expect(wrapper.get(".menu-shortcut").text()).toBe("Ctrl+P");
    expect(wrapper.find('[data-testid="menu-item-reason"]').exists()).toBe(false);
  });

  it("shows the stated reason in place of the shortcut when disabled with one (SPEC-068 AC-13)", () => {
    const wrapper = mount(MenuItem, {
      props: {
        label: "Mail merge…",
        shortcut: "Ctrl+M",
        disabled: true,
        reason: "not available on this screen size",
      },
    });
    const button = wrapper.get("button");
    expect(button.attributes("disabled")).toBeDefined();
    const reason = wrapper.get('[data-testid="menu-item-reason"]');
    expect(reason.text()).toBe("not available on this screen size");
    expect(button.attributes("title")).toBe("not available on this screen size");
    // The reason replaces the shortcut so the row is not double-annotated.
    expect(wrapper.find(".menu-shortcut").exists()).toBe(false);
  });

  it("does not show a reason while enabled, even if one is passed", () => {
    const wrapper = mount(MenuItem, {
      props: { label: "Mail merge…", reason: "not available on this screen size" },
    });
    expect(wrapper.find('[data-testid="menu-item-reason"]').exists()).toBe(false);
  });
});
