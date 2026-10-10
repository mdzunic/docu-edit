// SPEC-121 — the workbook menu bar's Edit menu (Undo/Redo naming the step,
// with the registry's shortcuts, disabled when nothing applies) and File ▸
// "Download a copy as …" (disabled with its reason), plus the state defaults.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SpreadsheetMenuBar from "./SpreadsheetMenuBar.vue";
import { asSpreadsheetMenuState, SHEET_MENU_COMMANDS } from "../../lib/spreadsheet/menuState";
import { shortcutLabel } from "../../lib/shortcuts";

describe("SpreadsheetMenuBar (SPEC-121)", () => {
  it("puts Edit between File and View, with Undo/Redo naming their step and routing edit.undo/edit.redo", async () => {
    const wrapper = mount(SpreadsheetMenuBar, {
      props: { state: { canUndo: true, undoLabel: "Rename sheet", canRedo: false } },
    });
    expect(wrapper.findAll("[data-menu-title]").map((title) => title.text())).toEqual(["File", "Edit", "View", "Format", "Data"]);
    const undo = wrapper.get('[data-testid="sheet-menu-undo"]');
    expect(undo.text()).toContain("Undo Rename sheet");
    expect(undo.get(".menu-shortcut").text()).toBe(shortcutLabel("edit.undo"));
    const redo = wrapper.get('[data-testid="sheet-menu-redo"]');
    expect(redo.get(".menu-item-label").text()).toBe("Redo");
    expect(redo.get(".menu-shortcut").text()).toBe(shortcutLabel("edit.redo"));
    expect(redo.attributes("disabled")).toBeDefined();
    await undo.trigger("click");
    expect(wrapper.emitted("command")).toEqual([["edit.undo"]]);
  });

  it("offers File ▸ Download a copy as … after Download, disabled with its reason", async () => {
    const wrapper = mount(SpreadsheetMenuBar, {
      props: { state: { copyLabel: "Download a copy as OpenDocument Spreadsheet (.ods)…" } },
    });
    const items = wrapper.get('[data-testid="menu-file"]').findAll('[role="menuitem"]');
    expect(items.map((item) => item.attributes("data-testid"))).toEqual([
      "sheet-menu-download",
      "sheet-menu-download-copy",
      "sheet-menu-close",
    ]);
    await items[1].trigger("click");
    expect(wrapper.emitted("command")).toEqual([[SHEET_MENU_COMMANDS.downloadCopy]]);
    await wrapper.setProps({ state: { copyDisabledReason: "Not available for this workbook yet" } });
    const copy = wrapper.get('[data-testid="sheet-menu-download-copy"]');
    expect(copy.attributes("disabled")).toBeDefined();
    expect(copy.attributes("title")).toBe("Not available for this workbook yet");
    expect(copy.get(".menu-item-label").text()).toBe("Download a copy as …");
  });

  it("defaults the new state fields to nothing to undo and an enabled generic copy", () => {
    expect(asSpreadsheetMenuState(undefined)).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoLabel: "",
      redoLabel: "",
      copyLabel: "Download a copy as …",
      copyDisabledReason: "",
    });
    expect(SHEET_MENU_COMMANDS).toMatchObject({ undo: "edit.undo", redo: "edit.redo", downloadCopy: "file.download-copy" });
  });
});
