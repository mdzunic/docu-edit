import { describe, it, expect, afterEach } from "vitest";
import { mount } from "@vue/test-utils";
import DocumentStatusBar from "./DocumentStatusBar.vue";
import { setTierWidthForTest } from "../composables/useResponsiveTier";

const BASE_PROPS = {
  caretPage: 1,
  pageCount: 3,
  wordCountText: "120 words",
  editable: true,
  trackChangesRecording: false,
};

describe("DocumentStatusBar responsive reduction (SPEC-068)", () => {
  afterEach(() => setTierWidthForTest(1440));

  it("keeps the page indicator and zoom, plus word count/language/track-changes, at the full tier", () => {
    setTierWidthForTest(1440);
    const wrapper = mount(DocumentStatusBar, { props: BASE_PROPS });
    expect(wrapper.find('[data-testid="page-indicator"]').exists()).toBe(true);
    expect(wrapper.findComponent({ name: "ZoomControl" }).exists()).toBe(true);
    expect(wrapper.find('[data-testid="word-count-button"]').exists()).toBe(true);
    expect(wrapper.findComponent({ name: "LanguageStatusControl" }).exists()).toBe(true);
    expect(wrapper.find('[data-testid="status-bar-track-changes"]').exists()).toBe(true);
  });

  it("reduces to just the page indicator and zoom at the compact tier (SPEC-068)", () => {
    setTierWidthForTest(390);
    const wrapper = mount(DocumentStatusBar, { props: BASE_PROPS });
    // Kept.
    expect(wrapper.find('[data-testid="page-indicator"]').exists()).toBe(true);
    expect(wrapper.findComponent({ name: "ZoomControl" }).exists()).toBe(true);
    // Dropped into the menu sheet / palette.
    expect(wrapper.find('[data-testid="word-count-button"]').exists()).toBe(false);
    expect(wrapper.findComponent({ name: "LanguageStatusControl" }).exists()).toBe(false);
    expect(wrapper.find('[data-testid="status-bar-track-changes"]').exists()).toBe(false);
  });
});
