import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import HomeView from "./HomeView.vue";

describe("HomeView", () => {
  it("titles the app bar 'No document'", () => {
    const wrapper = mount(HomeView);
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      "No document",
    );
  });

  it("shows the 'No document open' empty state in the canvas", () => {
    const wrapper = mount(HomeView);
    expect(wrapper.get('[data-testid="canvas"]').text()).toContain(
      "No document open",
    );
  });

  it("shows the 'New document' and 'Open document' buttons", () => {
    const wrapper = mount(HomeView);
    const newButton = wrapper.get('[data-testid="new-document"]');
    const openButton = wrapper.get('[data-testid="open-document"]');
    expect(newButton.text()).toBe("New document");
    expect(openButton.text()).toBe("Open document");
    expect(newButton.isVisible()).toBe(true);
    expect(openButton.isVisible()).toBe(true);
  });
});
