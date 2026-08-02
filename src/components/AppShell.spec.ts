import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import AppShell from "./AppShell.vue";

describe("AppShell", () => {
  it("renders the app bar with the product name and document title", () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "My doc" } });
    expect(wrapper.get('[data-testid="app-bar"]').isVisible()).toBe(true);
    expect(wrapper.get('[data-testid="product-name"]').text()).toBe(
      "docu-edit",
    );
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe("My doc");
  });

  it("renders a toolbar strip and a canvas area", () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
    expect(wrapper.get('[data-testid="toolbar"]').attributes("role")).toBe(
      "toolbar",
    );
    expect(wrapper.get('[data-testid="canvas"]').element.tagName).toBe("MAIN");
  });

  it("places slotted content into the toolbar and canvas", () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "t" },
      slots: {
        toolbar: "<button>Bold</button>",
        default: "<p>canvas content</p>",
      },
    });
    expect(wrapper.get('[data-testid="toolbar"]').text()).toContain("Bold");
    expect(wrapper.get('[data-testid="canvas"]').text()).toContain(
      "canvas content",
    );
  });
});
