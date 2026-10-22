import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import FontSizeControl from "./FontSizeControl.vue";
import { COMMON_FONT_SIZES } from "../lib/fonts";

// SPEC-061 AC-21: the font-size stepper is one control — an editable numeric
// field, −/+ steppers and a dropdown of common sizes.
describe("FontSizeControl", () => {
  it("commits a typed size on Enter (AC-21)", async () => {
    const wrapper = mount(FontSizeControl, { props: { value: 11 } });
    const input = wrapper.get('[data-testid="font-size-input"]');
    await input.setValue("28");
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("change")?.at(-1)).toEqual([28]);
  });

  it("steps up and down by the 0.5pt grid (AC-21)", async () => {
    const wrapper = mount(FontSizeControl, { props: { value: 11 } });
    await wrapper.get('[data-testid="font-size-increase"]').trigger("click");
    expect(wrapper.emitted("change")?.at(-1)).toEqual([11.5]);
    await wrapper.get('[data-testid="font-size-decrease"]').trigger("click");
    expect(wrapper.emitted("change")?.at(-1)).toEqual([10.5]);
  });

  it("opens a dropdown of the common sizes and picking one emits change (AC-21)", async () => {
    const wrapper = mount(FontSizeControl, { props: { value: 11 } });
    expect(wrapper.find('[data-testid="font-size-menu"]').exists()).toBe(false);
    await wrapper.get('[data-testid="font-size-menu-button"]').trigger("click");
    const menu = wrapper.get('[data-testid="font-size-menu"]');
    const options = menu.findAll('[role="menuitemradio"]');
    expect(options).toHaveLength(COMMON_FONT_SIZES.length);
    await wrapper.get('[data-testid="font-size-option-24"]').trigger("click");
    expect(wrapper.emitted("change")?.at(-1)).toEqual([24]);
    // The menu closes after a pick.
    expect(wrapper.find('[data-testid="font-size-menu"]').exists()).toBe(false);
  });

  it("marks the current size checked in the dropdown (AC-21)", async () => {
    const wrapper = mount(FontSizeControl, { props: { value: 12 } });
    await wrapper.get('[data-testid="font-size-menu-button"]').trigger("click");
    const current = wrapper.get('[data-testid="font-size-option-12"]');
    expect(current.attributes("aria-checked")).toBe("true");
  });

  it("clamps typed entry to a supplied min/max (SPEC-081 AC-18: 6–96)", async () => {
    const wrapper = mount(FontSizeControl, { props: { value: 18, min: 6, max: 96 } });
    const input = wrapper.get('[data-testid="font-size-input"]');
    await input.setValue("200");
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("change")?.at(-1)).toEqual([96]);
    await input.setValue("2");
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("change")?.at(-1)).toEqual([6]);
  });
});
