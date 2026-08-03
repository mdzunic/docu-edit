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

  it("does not enter edit mode on click when titleEditable is unset", async () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    expect(wrapper.find('[data-testid="document-title-input"]').exists()).toBe(
      false,
    );
  });

  it("clicking the title enters edit mode with the current title (AC-6)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get<HTMLInputElement>(
      '[data-testid="document-title-input"]',
    );
    expect(input.element.value).toBe("My doc");
  });

  it("emits update:documentTitle on Enter (AC-7)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("New title");
    await input.trigger("keydown.enter");
    expect(wrapper.emitted("update:documentTitle")).toEqual([["New title"]]);
    expect(wrapper.find('[data-testid="document-title-input"]').exists()).toBe(
      false,
    );
  });

  it("emits update:documentTitle on blur (AC-7)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Blurred title");
    await input.trigger("blur");
    expect(wrapper.emitted("update:documentTitle")).toEqual([
      ["Blurred title"],
    ]);
  });

  it("Escape cancels editing without emitting (AC-8)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Discarded");
    await input.trigger("keydown.escape");
    expect(wrapper.emitted("update:documentTitle")).toBeUndefined();
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      "My doc",
    );
  });

  it("shows no format badge when formatBadge is unset", () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
    expect(wrapper.find('[data-testid="format-badge"]').exists()).toBe(false);
  });

  it("shows the format badge text when set (AC-30)", () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "t", formatBadge: "DOCX" },
    });
    expect(wrapper.get('[data-testid="format-badge"]').text()).toBe("DOCX");
  });

  it("places app-bar-actions slot content inside the app bar (SPEC-006 AC-1)", () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "t" },
      slots: { "app-bar-actions": "<button>Download</button>" },
    });
    expect(wrapper.get('[data-testid="app-bar"]').text()).toContain("Download");
  });
});
