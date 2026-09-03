// MenuCheckItem (SPEC-049 AC-11): a menuitemcheckbox whose mark mirrors the
// live state, including the ARIA tri-state "mixed" a group toggle such as
// View → Rulers reports when only some of its members are on.
import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import MenuCheckItem from "./MenuCheckItem.vue";

describe("MenuCheckItem", () => {
  it.each([
    [true, "true", "✓"],
    [false, "false", ""],
    ["mixed", "mixed", "–"],
  ])("renders checked=%s as aria-checked=%s with the mark %s", (checked, aria, mark) => {
    const wrapper = mount(MenuCheckItem, { props: { label: "Rulers", checked } });
    const button = wrapper.get("button");
    expect(button.attributes("role")).toBe("menuitemcheckbox");
    expect(button.attributes("aria-checked")).toBe(aria);
    expect(wrapper.get(".menu-check").text()).toBe(mark);
  });

  it("emits select on click", async () => {
    const wrapper = mount(MenuCheckItem, { props: { label: "Rulers", checked: true } });
    await wrapper.get("button").trigger("click");
    expect(wrapper.emitted("select")).toHaveLength(1);
  });

  // SPEC-076 AC-18: a disabled toggle states why, exactly as MenuItem does —
  // a check item is never silently greyed out.
  it("states the reason it is disabled, in the item and as a title", () => {
    const wrapper = mount(MenuCheckItem, {
      props: {
        label: "Repeat header row",
        disabled: true,
        reason: "the caret is not in a table",
      },
    });
    expect(wrapper.get("[data-testid='menu-item-reason']").text()).toBe("the caret is not in a table");
    expect(wrapper.get("button").attributes("title")).toBe("the caret is not in a table");
  });

  it("states nothing while it is enabled", () => {
    const wrapper = mount(MenuCheckItem, {
      props: { label: "Repeat header row", reason: "the caret is not in a table", shortcut: "Ctrl+R" },
    });
    expect(wrapper.find("[data-testid='menu-item-reason']").exists()).toBe(false);
    expect(wrapper.get("button").attributes("title")).toBeUndefined();
    expect(wrapper.get(".menu-shortcut").text()).toBe("Ctrl+R");
  });
});
