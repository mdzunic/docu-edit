import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import CommandPalette from "./CommandPalette.vue";
import { COMMANDS } from "../lib/commandCatalogue";

const documentCommands = COMMANDS.filter((c) => c.scope === "document");

function mountPalette(overrides: Record<string, unknown> = {}) {
  return mount(CommandPalette, {
    attachTo: document.body,
    props: {
      commands: documentCommands,
      recentIds: [],
      commonIds: ["edit.undo", "format.bold"],
      ...overrides,
    },
  });
}

describe("CommandPalette (SPEC-062)", () => {
  it("is a role=dialog with a combobox/listbox search pattern (AC-27, AC-28, AC-29)", () => {
    const wrapper = mountPalette();
    expect(wrapper.find('[role="dialog"]').exists()).toBe(true);
    const input = wrapper.get('[role="combobox"]');
    expect(input.attributes("aria-expanded")).toBe("true");
    expect(input.attributes("aria-controls")).toBe("command-palette-listbox");
    expect(wrapper.find('#command-palette-listbox[role="listbox"]').exists()).toBe(true);
    expect(wrapper.findAll('[role="option"]').length).toBeGreaterThan(0);
  });

  it("announces the result count in a live region (AC-30)", async () => {
    const wrapper = mountPalette();
    await wrapper.get('[data-testid="command-palette-input"]').setValue("bold");
    const live = wrapper.get('[data-testid="command-palette-count"]');
    expect(live.attributes("aria-live")).toBe("polite");
    const count = wrapper.findAll('[role="option"]').length;
    expect(live.text()).toContain(String(count));
  });

  it("empty query shows recents first, then common commands (AC-20, AC-21)", () => {
    const wrapper = mountPalette({ recentIds: ["tools.word-count"], commonIds: ["edit.undo"] });
    const ids = wrapper.findAll('[role="option"]').map((o) => o.attributes("data-command-id"));
    expect(ids[0]).toBe("tools.word-count");
    expect(ids).toContain("edit.undo");
  });

  it("filters by subsequence and highlights the matched characters (AC-9, AC-11)", async () => {
    const wrapper = mountPalette();
    await wrapper.get('[data-testid="command-palette-input"]').setValue("bold");
    const first = wrapper.findAll('[role="option"]')[0];
    expect(first.attributes("data-command-id")).toBe("format.bold");
    expect(first.find("mark.palette-hit").exists()).toBe(true);
  });

  it("shows the menu path and the shortcut label from shortcutLabel (AC-6, AC-8)", async () => {
    const wrapper = mountPalette();
    await wrapper.get('[data-testid="command-palette-input"]').setValue("footnote");
    const row = wrapper.findAll('[role="option"]').find((o) => o.attributes("data-command-id") === "insert.footnote");
    expect(row?.text()).toContain("Insert →");
    // insert.footnote is bound (Ctrl+Alt+F) — its shortcut label must show.
    expect(row?.find(".palette-shortcut").text().length).toBeGreaterThan(0);
  });

  it("moves the highlight with Up/Down and runs the active row on Enter (AC-12, AC-13)", async () => {
    const wrapper = mountPalette();
    const input = wrapper.get('[data-testid="command-palette-input"]');
    await input.setValue("align");
    const options = wrapper.findAll('[role="option"]');
    const secondId = options[1].attributes("data-command-id");
    await input.trigger("keydown", { key: "ArrowDown" });
    expect(wrapper.get('[role="combobox"]').attributes("aria-activedescendant")).toBe("command-palette-option-1");
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("run")?.[0]).toEqual([secondId]);
  });

  it("lists a command that does not apply as disabled with a reason, not hidden (AC-18, AC-19)", async () => {
    const wrapper = mountPalette({ availability: { "edit.undo": "nothing to undo" } });
    await wrapper.get('[data-testid="command-palette-input"]').setValue("undo");
    const row = wrapper.findAll('[role="option"]').find((o) => o.attributes("data-command-id") === "edit.undo");
    expect(row).toBeDefined();
    expect(row?.attributes("aria-disabled")).toBe("true");
    expect(row?.text()).toContain("nothing to undo");
  });

  it("does not run a disabled row on Enter", async () => {
    const wrapper = mountPalette({
      commands: [documentCommands.find((c) => c.id === "edit.undo")],
      availability: { "edit.undo": "nothing to undo" },
    });
    const input = wrapper.get('[data-testid="command-palette-input"]');
    await input.setValue("undo");
    // The only row is disabled; Enter must be a no-op.
    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("run")).toBeUndefined();
  });

  it("merges per-document virtual style rows into the same search (AC-22, AC-23)", async () => {
    const wrapper = mountPalette({
      extraRows: [
        { id: "style-apply:Heading1", label: "Apply style: Heading 1", path: ["Styles"], scope: "document" },
      ],
    });
    await wrapper.get('[data-testid="command-palette-input"]').setValue("apply style heading 1");
    const ids = wrapper.findAll('[role="option"]').map((o) => o.attributes("data-command-id"));
    expect(ids).toContain("style-apply:Heading1");
  });

  it("disabled rows sort after enabled ones (AC-10 tail rule)", async () => {
    const wrapper = mountPalette({ availability: { "format.bold": "needs a selection" } });
    await wrapper.get('[data-testid="command-palette-input"]').setValue("b");
    const options = wrapper.findAll('[role="option"]');
    const boldIndex = options.findIndex((o) => o.attributes("data-command-id") === "format.bold");
    const enabledCount = options.filter((o) => o.attributes("aria-disabled") !== "true").length;
    // format.bold (disabled) must fall into the disabled tail.
    expect(boldIndex).toBeGreaterThanOrEqual(enabledCount);
  });
});
