import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import FindBar from "./FindBar.vue";

function mountBar(props: Record<string, unknown> = {}) {
  return mount(FindBar, {
    props: {
      visible: true,
      query: "cat",
      matchCount: 3,
      currentIndex: 0,
      ...props,
    },
    attachTo: document.body,
  });
}

describe("FindBar", () => {
  it("renders the query field, count, navigation and close (AC-1..4)", () => {
    const wrapper = mountBar();
    expect(wrapper.get('[data-testid="find-query"]').attributes("value")).toBe("cat");
    expect(wrapper.get('[data-testid="find-count"]').text()).toBe("1 of 3");
    expect(wrapper.find('[data-testid="find-prev"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="find-next"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="find-close"]').exists()).toBe(true);
  });

  it("states its scope to the user (AC-28)", () => {
    const wrapper = mountBar();
    const scope = wrapper.get('[data-testid="find-scope"]');
    expect(scope.text()).toBe("Body text only");
    expect(scope.attributes("title")).toContain("Headers, footers and footnotes are excluded");
  });

  it("shows the three search options as toggle buttons (AC-20..22)", () => {
    const wrapper = mountBar({ matchCase: true, wholeWords: false, useRegex: false });
    expect(wrapper.get('[data-testid="find-match-case"]').attributes("aria-pressed")).toBe("true");
    expect(wrapper.get('[data-testid="find-whole-words"]').attributes("aria-pressed")).toBe("false");
    expect(wrapper.get('[data-testid="find-regex"]').attributes("aria-pressed")).toBe("false");
  });

  it("emits option changes when a toggle is clicked", async () => {
    const wrapper = mountBar({ matchCase: false });
    await wrapper.get('[data-testid="find-match-case"]').trigger("click");
    expect(wrapper.emitted("update:matchCase")).toEqual([[true]]);
  });

  it("steps forward on Enter and backward on Shift+Enter (AC-8, AC-9)", async () => {
    const wrapper = mountBar();
    const query = wrapper.get('[data-testid="find-query"]');
    await query.trigger("keydown", { key: "Enter" });
    await query.trigger("keydown", { key: "Enter", shiftKey: true });
    expect(wrapper.emitted("next")).toHaveLength(1);
    expect(wrapper.emitted("prev")).toHaveLength(1);
  });

  it("closes on Escape (AC-11)", async () => {
    const wrapper = mountBar();
    await wrapper.get('[data-testid="find-query"]').trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("close")).toHaveLength(1);
  });

  it("reveals the replace row when expanded (AC-14, AC-15)", () => {
    const collapsed = mountBar({ expanded: false });
    expect(collapsed.find('[data-testid="find-replace-input"]').exists()).toBe(false);
    const expanded = mountBar({ expanded: true });
    expect(expanded.find('[data-testid="find-replace-input"]').exists()).toBe(true);
    expect(expanded.find('[data-testid="find-replace"]').exists()).toBe(true);
    expect(expanded.find('[data-testid="find-replace-all"]').exists()).toBe(true);
  });

  it("shows an inline error and disables the search actions for an invalid regex (AC-24, AC-25)", () => {
    const wrapper = mountBar({ expanded: true, error: "Invalid regular expression", matchCount: 0, currentIndex: -1 });
    const status = wrapper.get('[data-testid="find-status"]');
    expect(status.text()).toBe("Invalid regular expression");
    expect(status.attributes("role")).toBe("alert");
    expect(wrapper.get('[data-testid="find-prev"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="find-next"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="find-replace"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="find-replace-all"]').attributes("disabled")).toBeDefined();
  });

  it("disables navigation and replace but keeps Replace all enabled with no matches (AC-17)", () => {
    const wrapper = mountBar({ expanded: true, query: "zzz", matchCount: 0, currentIndex: -1 });
    expect(wrapper.get('[data-testid="find-count"]').text()).toBe("0 of 0");
    expect(wrapper.get('[data-testid="find-status"]').text()).toBe("No results");
    expect(wrapper.get('[data-testid="find-prev"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="find-replace"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="find-replace-all"]').attributes("disabled")).toBeUndefined();
  });

  it("disables every action for an empty query (AC-16)", () => {
    const wrapper = mountBar({ expanded: true, query: "", matchCount: 0, currentIndex: -1 });
    expect(wrapper.get('[data-testid="find-replace-all"]').attributes("disabled")).toBeDefined();
  });

  it("shows a transient wrap hint when provided (AC-10)", () => {
    const wrapper = mountBar({ hint: "Wrapped to top" });
    expect(wrapper.get('[data-testid="find-status"]').text()).toBe("Wrapped to top");
  });

  it("emits replace and replace-all from the buttons (AC-16, AC-17)", async () => {
    const wrapper = mountBar({ expanded: true });
    await wrapper.get('[data-testid="find-replace"]').trigger("click");
    await wrapper.get('[data-testid="find-replace-all"]').trigger("click");
    expect(wrapper.emitted("replace")).toHaveLength(1);
    expect(wrapper.emitted("replace-all")).toHaveLength(1);
  });
});
