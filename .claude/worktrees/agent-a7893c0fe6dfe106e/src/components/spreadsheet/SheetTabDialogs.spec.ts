// SPEC-121 — the Delete-sheet confirmation and the Unhide dialog, both on
// BaseDialog: exact message and buttons, Cancel first in focus, and the
// Unhide listbox's multi-select by pointer and keyboard.
import { afterEach, describe, expect, it } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import SheetTabDeleteDialog from "./SheetTabDeleteDialog.vue";
import SheetUnhideDialog from "./SheetUnhideDialog.vue";

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

describe("SheetTabDeleteDialog", () => {
  it("names the sheet, focuses Cancel, and emits confirm / cancel (Escape cancels)", async () => {
    const wrapper = mount(SheetTabDeleteDialog, { props: { name: "Budget" }, attachTo: document.body });
    mounted.push(wrapper);
    await flushPromises();
    expect(wrapper.get('[data-testid="sheet-tab-delete-dialog"]').get("h2").text()).toBe("Delete sheet");
    expect(wrapper.get(".dialog-message").text()).toBe('Delete "Budget"? Its cells and anything on it will be removed.');
    const cancel = wrapper.get('[data-testid="sheet-tab-delete-cancel"]');
    expect(document.activeElement).toBe(cancel.element);
    expect(wrapper.get('[data-testid="sheet-tab-delete-confirm"]').classes()).toContain("variant-danger");
    await wrapper.get('[data-testid="sheet-tab-delete-confirm"]').trigger("click");
    await cancel.trigger("click");
    await wrapper.get('[role="dialog"]').trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("confirm")).toHaveLength(1);
    expect(wrapper.emitted("cancel")).toHaveLength(2);
  });
});

describe("SheetUnhideDialog", () => {
  const SHEETS = [
    { index: 1, name: "Archive" },
    { index: 4, name: "Notes" },
    { index: 6, name: "Old" },
  ];

  async function open() {
    const wrapper = mount(SheetUnhideDialog, { props: { sheets: SHEETS }, attachTo: document.body });
    mounted.push(wrapper);
    await flushPromises();
    return wrapper;
  }

  it("lists the hidden sheets in a multi-select listbox; Unhide waits for a choice and confirms in list order", async () => {
    const wrapper = await open();
    const list = wrapper.get('[role="listbox"]');
    expect(list.attributes("aria-multiselectable")).toBe("true");
    expect(list.attributes("aria-label")).toBe("Hidden sheets");
    const options = wrapper.findAll('[data-testid="sheet-unhide-option"]');
    expect(options.map((option) => option.text())).toEqual(["Archive", "Notes", "Old"]);
    expect(options.map((option) => option.attributes("aria-selected"))).toEqual(["false", "false", "false"]);
    const confirm = wrapper.get('[data-testid="sheet-unhide-confirm"]');
    expect(confirm.attributes("disabled")).toBeDefined();
    await options[2].trigger("click");
    await options[0].trigger("click");
    await options[1].trigger("click");
    await options[1].trigger("click");
    expect(options.map((option) => option.attributes("aria-selected"))).toEqual(["true", "false", "true"]);
    expect(confirm.attributes("disabled")).toBeUndefined();
    await confirm.trigger("click");
    expect(wrapper.emitted("confirm")).toEqual([[[1, 6]]]);
    await wrapper.get('[data-testid="sheet-unhide-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });

  it("moves the active option with the arrows (roving tabindex), toggles on Space and confirms on Enter", async () => {
    const wrapper = await open();
    const options = () => wrapper.findAll('[data-testid="sheet-unhide-option"]');
    expect(document.activeElement).toBe(options()[0].element);
    const list = wrapper.get('[role="listbox"]');
    await list.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("confirm")).toBeUndefined();
    await list.trigger("keydown", { key: "ArrowDown" });
    await flushPromises();
    expect(options().map((option) => option.attributes("tabindex"))).toEqual(["-1", "0", "-1"]);
    expect(document.activeElement).toBe(options()[1].element);
    await list.trigger("keydown", { key: " " });
    await list.trigger("keydown", { key: "End" });
    await list.trigger("keydown", { key: " " });
    await list.trigger("keydown", { key: "ArrowUp" });
    await list.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("confirm")).toEqual([[[4, 6]]]);
  });
});
