// SPEC-121 — the workbook menu bar's Edit menu (Undo/Redo naming the step,
// with the registry's shortcuts, disabled when nothing applies) and File ▸
// "Download a copy as …" (disabled with its reason), plus the state defaults.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SpreadsheetMenuBar from "./SpreadsheetMenuBar.vue";
import { asSpreadsheetMenuState, FILTER_OFF_REASON, SHEET_MENU_COMMANDS } from "../../lib/spreadsheet/menuState";
import { shortcutLabel } from "../../lib/shortcuts";
import { SHEET_SHORTCUTS } from "../../views/spreadsheet/sheetShortcuts";

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

describe("SpreadsheetMenuBar's Data menu (SPEC-128)", () => {
  it("has Recalculate, a separator, the three sorts, a separator, then Filter, Clear filters and Reapply filter, each routing its command", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { filterOn: true } } });
    const menu = wrapper.get('[data-testid="menu-data"]');
    const rows = menu.findAll('[role^="menuitem"], hr').map((row) => (row.element.tagName === "HR" ? "—" : row.attributes("data-testid")));
    expect(rows).toEqual([
      "sheet-menu-recalculate",
      "—",
      "sheet-menu-sort-asc",
      "sheet-menu-sort-desc",
      "sheet-menu-sort-range",
      "—",
      "sheet-menu-filter",
      "sheet-menu-clear-filters",
      "sheet-menu-reapply-filter",
      // SPEC-130: Data validation… after its own separator, below the filters.
      "—",
      "sheet-menu-data-validation",
    ]);
    expect(menu.findAll(".menu-item-label").map((label) => label.text())).toEqual([
      "Recalculate",
      "Sort A to Z",
      "Sort Z to A",
      "Sort range…",
      "Filter",
      "Clear filters",
      "Reapply filter",
      "Data validation…",
    ]);
    const filter = wrapper.get('[data-testid="sheet-menu-filter"]');
    expect(filter.attributes("role")).toBe("menuitemcheckbox");
    expect(filter.attributes("aria-checked")).toBe("true");
    expect(filter.text()).toContain(shortcutLabel("sheet.toggleFilter", false, SHEET_SHORTCUTS));
    expect(wrapper.get('[data-testid="sheet-menu-reapply-filter"]').text()).toContain(shortcutLabel("sheet.reapplyFilter", false, SHEET_SHORTCUTS));
    for (const [testid, command] of [
      ["sheet-menu-sort-asc", SHEET_MENU_COMMANDS.sortAscending],
      ["sheet-menu-sort-desc", SHEET_MENU_COMMANDS.sortDescending],
      ["sheet-menu-sort-range", SHEET_MENU_COMMANDS.sortRange],
      ["sheet-menu-filter", SHEET_MENU_COMMANDS.filter],
      ["sheet-menu-clear-filters", SHEET_MENU_COMMANDS.clearFilters],
      ["sheet-menu-reapply-filter", SHEET_MENU_COMMANDS.reapplyFilter],
    ]) {
      await wrapper.get(`[data-testid="${testid}"]`).trigger("click");
      expect(wrapper.emitted("command")?.at(-1)).toEqual([command]);
    }
  });

  it("without an autofilter Filter is unchecked and Clear filters and Reapply filter are disabled, saying 'Turn on Filter first'", () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: {} } });
    expect(wrapper.get('[data-testid="sheet-menu-filter"]').attributes("aria-checked")).toBe("false");
    for (const testid of ["sheet-menu-clear-filters", "sheet-menu-reapply-filter"]) {
      const item = wrapper.get(`[data-testid="${testid}"]`);
      expect(item.attributes("disabled")).toBeDefined();
      expect(item.get('[data-testid="menu-item-reason"]').text()).toBe(FILTER_OFF_REASON);
    }
    expect(FILTER_OFF_REASON).toBe("Turn on Filter first");
    expect(asSpreadsheetMenuState(undefined).filterOn).toBe(false);
  });
});

describe("SpreadsheetMenuBar's Format menu (SPEC-129)", () => {
  it("offers Conditional formatting… after Cell styles, routing format.conditional-formatting and disabled without a grid sheet", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: true } } });
    const menu = wrapper.get('[data-testid="menu-format"]');
    const ids = menu.findAll("[data-testid]").map((row) => row.attributes("data-testid"));
    expect(ids.indexOf("sheet-menu-conditional-formatting")).toBeGreaterThan(ids.indexOf("sheet-menu-cell-styles"));
    expect(ids.indexOf("sheet-menu-conditional-formatting")).toBeLessThan(ids.indexOf("sheet-menu-clear"));
    const item = wrapper.get('[data-testid="sheet-menu-conditional-formatting"]');
    expect(item.text()).toContain("Conditional formatting…");
    await item.trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.conditionalFormatting]);
    const noGrid = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: false } } });
    expect(noGrid.get('[data-testid="sheet-menu-conditional-formatting"]').attributes("disabled")).toBeDefined();
  });
});

describe("SpreadsheetMenuBar's Data menu (SPEC-130)", () => {
  it("offers Data validation… after a separator below the filter items, routing data.validation and disabled without a grid sheet", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: true } } });
    const menu = wrapper.get('[data-testid="menu-data"]');
    const rows = menu.findAll('[role^="menuitem"], hr').map((row) => (row.element.tagName === "HR" ? "—" : row.attributes("data-testid")));
    expect(rows.slice(-3)).toEqual(["sheet-menu-reapply-filter", "—", "sheet-menu-data-validation"]);
    const item = wrapper.get('[data-testid="sheet-menu-data-validation"]');
    expect(item.attributes("role")).toBe("menuitem");
    expect(item.get(".menu-item-label").text()).toBe("Data validation…");
    expect(item.attributes("disabled")).toBeUndefined();
    await item.trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.dataValidation]);
    expect(SHEET_MENU_COMMANDS.dataValidation).toBe("data.validation");
    const foreign = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: false } } });
    expect(foreign.get('[data-testid="sheet-menu-data-validation"]').attributes("disabled")).toBeDefined();
  });
});
