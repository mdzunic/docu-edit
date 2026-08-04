import { describe, it, expect, vi } from "vitest";
import { mount } from "@vue/test-utils";
import FormatToolbar from "./FormatToolbar.vue";

const OFF_STATES = { bold: "off", italic: "off", underline: "off", strikethrough: "off" };

describe("FormatToolbar", () => {
  it("renders Bold, Italic, Underline and Strikethrough buttons with tooltips (AC-1, AC-2)", () => {
    const wrapper = mount(FormatToolbar, { props: { states: OFF_STATES } });
    const bold = wrapper.get('[data-testid="format-bold"]');
    expect(bold.attributes("title")).toBe("Bold (Ctrl+B)");
    const italic = wrapper.get('[data-testid="format-italic"]');
    expect(italic.attributes("title")).toBe("Italic (Ctrl+I)");
    const underline = wrapper.get('[data-testid="format-underline"]');
    expect(underline.attributes("title")).toBe("Underline (Ctrl+U)");
    const strikethrough = wrapper.get('[data-testid="format-strikethrough"]');
    expect(strikethrough.attributes("title")).toBe("Strikethrough (Ctrl+Shift+X)");
  });

  it("reflects an 'off' state as aria-pressed=false", () => {
    const wrapper = mount(FormatToolbar, { props: { states: OFF_STATES } });
    expect(wrapper.get('[data-testid="format-bold"]').attributes("aria-pressed")).toBe("false");
  });

  it("reflects an 'on' state as aria-pressed=true (AC-12, AC-13)", () => {
    const wrapper = mount(FormatToolbar, {
      props: { states: { ...OFF_STATES, bold: "on" } },
    });
    expect(wrapper.get('[data-testid="format-bold"]').attributes("aria-pressed")).toBe("true");
  });

  it("reflects a 'mixed' state as aria-pressed=mixed (AC-14)", () => {
    const wrapper = mount(FormatToolbar, {
      props: { states: { ...OFF_STATES, italic: "mixed" } },
    });
    expect(wrapper.get('[data-testid="format-italic"]').attributes("aria-pressed")).toBe("mixed");
  });

  it("emits toggle with the format key on click", async () => {
    const wrapper = mount(FormatToolbar, { props: { states: OFF_STATES } });
    await wrapper.get('[data-testid="format-underline"]').trigger("click");
    expect(wrapper.emitted("toggle")).toEqual([["underline"]]);
  });

  it("prevents mousedown default so a live text selection is not lost on click", () => {
    const wrapper = mount(FormatToolbar, { props: { states: OFF_STATES } });
    const event = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    const prevented = vi.spyOn(event, "preventDefault");
    wrapper.get('[data-testid="format-bold"]').element.dispatchEvent(event);
    expect(prevented).toHaveBeenCalled();
  });

  it("disables the buttons and does not emit toggle when disabled", async () => {
    const wrapper = mount(FormatToolbar, { props: { states: OFF_STATES, disabled: true } });
    const bold = wrapper.get('[data-testid="format-bold"]');
    expect(bold.attributes("disabled")).toBeDefined();
    await bold.trigger("click");
    expect(wrapper.emitted("toggle")).toBeUndefined();
  });
});
