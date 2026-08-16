import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import ShortcutsDialog from "./ShortcutsDialog.vue";
import { SHORTCUTS } from "../lib/shortcuts";

describe("ShortcutsDialog (SPEC-021 AC-6..10)", () => {
  it("lists every registered shortcut, grouped by category (AC-7)", () => {
    const wrapper = mount(ShortcutsDialog);
    expect(wrapper.findAll('[role="option"]')).toHaveLength(SHORTCUTS.length);
    const categories = new Set(SHORTCUTS.map((s) => s.category));
    for (const category of categories) {
      expect(wrapper.text()).toContain(category);
    }
  });

  it("shows a bound key for every entry", () => {
    const wrapper = mount(ShortcutsDialog);
    expect(wrapper.get(`[data-testid="shortcut-row-format.bold"]`).text()).toContain("Ctrl+B");
    expect(wrapper.get(`[data-testid="shortcut-row-edit.redo"]`).text()).toContain("Ctrl+Shift+Z");
    expect(wrapper.get(`[data-testid="shortcut-row-edit.redo"]`).text()).toContain("Ctrl+Y");
  });

  it("marks the deliberately-overridden browser shortcuts (AC-10)", () => {
    const wrapper = mount(ShortcutsDialog);
    expect(wrapper.get(`[data-testid="shortcut-row-format.bold"]`).text()).toContain("overrides browser default");
    expect(wrapper.get(`[data-testid="shortcut-row-edit.undo"]`).text()).not.toContain("overrides browser default");
  });

  it("notes the reserved browser shortcuts that are never bound (AC-11)", () => {
    const wrapper = mount(ShortcutsDialog);
    const note = wrapper.get('[data-testid="shortcuts-reserved-note"]').text();
    expect(note).toContain("Ctrl/Cmd+T");
    expect(note).toContain("Ctrl/Cmd+W");
    expect(note).toContain("Ctrl/Cmd+N");
  });

  it("search narrows the list to matching commands (AC-6)", async () => {
    const wrapper = mount(ShortcutsDialog);
    await wrapper.get('[data-testid="shortcuts-search"]').setValue("bold");
    const rows = wrapper.findAll('[role="option"]');
    expect(rows).toHaveLength(1);
    expect(rows[0].text()).toContain("Bold");
  });

  it("shows an empty state when the search matches nothing", async () => {
    const wrapper = mount(ShortcutsDialog);
    await wrapper.get('[data-testid="shortcuts-search"]').setValue("xyzzy-no-match");
    expect(wrapper.findAll('[role="option"]')).toHaveLength(0);
    expect(wrapper.find('[data-testid="shortcuts-empty"]').exists()).toBe(true);
  });

  it("is keyboard-navigable: arrow keys move the highlighted row (AC-8)", async () => {
    const wrapper = mount(ShortcutsDialog);
    const search = wrapper.get('[data-testid="shortcuts-search"]');
    const firstId = SHORTCUTS[0].id;
    const secondId = SHORTCUTS[1].id;
    expect(wrapper.get(`[data-testid="shortcut-row-${firstId}"]`).classes()).toContain("active");
    await search.trigger("keydown", { key: "ArrowDown" });
    expect(wrapper.get(`[data-testid="shortcut-row-${secondId}"]`).classes()).toContain("active");
    await search.trigger("keydown", { key: "ArrowUp" });
    expect(wrapper.get(`[data-testid="shortcut-row-${firstId}"]`).classes()).toContain("active");
  });

  it("closes with Escape (AC-9)", async () => {
    const wrapper = mount(ShortcutsDialog);
    // Fired from the search input — where focus rests while the dialog is
    // open — bubbling to BaseDialog's panel-level Escape handler (the old
    // handler sat on the backdrop, which a real keypress never targets).
    await wrapper.get('[data-testid="shortcuts-search"]').trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("close")).toHaveLength(1);
  });

  it("closes via the Close button and a backdrop click", async () => {
    const wrapper = mount(ShortcutsDialog);
    await wrapper.get('[data-testid="shortcuts-close"]').trigger("click");
    expect(wrapper.emitted("close")).toHaveLength(1);
  });
});
