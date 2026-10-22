// SPEC-058 primitive, extended for SPEC-060 AC-9: an icon-only button carries
// its label as both the accessible name and the tooltip, and its busy state
// marks the control aria-busy and blocks a second activation while the action
// it triggered is in flight.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import UiIconButton from "./UiIconButton.vue";

describe("UiIconButton", () => {
  it("labels the button and its tooltip from `label`", () => {
    const wrapper = mount(UiIconButton, {
      props: { icon: "close", label: "Close document" },
    });
    const button = wrapper.get("button");
    expect(button.attributes("aria-label")).toBe("Close document");
    expect(button.attributes("title")).toBe("Close document");
  });

  it("is not busy or disabled by default", () => {
    const wrapper = mount(UiIconButton, {
      props: { icon: "close", label: "Close document" },
    });
    const button = wrapper.get("button");
    expect(button.attributes("aria-busy")).toBeUndefined();
    expect(button.attributes("disabled")).toBeUndefined();
  });

  it("marks the busy state aria-busy and disables the control (AC-9)", () => {
    const wrapper = mount(UiIconButton, {
      props: { icon: "close", label: "Close document", busy: true },
    });
    const button = wrapper.get("button");
    expect(button.attributes("aria-busy")).toBe("true");
    expect(button.attributes("disabled")).toBeDefined();
    expect(button.classes()).toContain("is-busy");
  });
});
