// SPEC-121 — the workbook menu bar's Edit menu (Undo/Redo naming the step,
// with the registry's shortcuts, disabled when nothing applies) and File ▸
// "Download a copy as …" (disabled with its reason), plus the state defaults.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SpreadsheetMenuBar from "./SpreadsheetMenuBar.vue";
import { asSpreadsheetMenuState, FILTER_OFF_REASON, SHEET_MENU_COMMANDS } from "../../lib/spreadsheet/menuState";
import { isMacPlatform } from "../../lib/core/platform";
import { shortcutLabel } from "../../lib/shortcuts";
import { SHEET_SHORTCUTS } from "../../views/spreadsheet/sheetShortcuts";

describe("SpreadsheetMenuBar (SPEC-121)", () => {
  it("puts Edit between File and View, with Undo/Redo naming their step and routing edit.undo/edit.redo", async () => {
    const wrapper = mount(SpreadsheetMenuBar, {
      props: { state: { canUndo: true, undoLabel: "Rename sheet", canRedo: false } },
    });
    // SPEC-131 puts Insert after View, then Arrange.
    expect(wrapper.findAll("[data-menu-title]").map((title) => title.text())).toEqual(["File", "Edit", "View", "Insert", "Arrange", "Format", "Formulas", "Data"]);
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
    // SPEC-136: Page setup…, Print area ▸ (its flyout holds Set / Clear) and Print… come before Close.
    expect(items.map((item) => item.attributes("data-testid"))).toEqual([
      "sheet-menu-download",
      "sheet-menu-download-copy",
      "sheet-menu-page-setup",
      "sheet-menu-print-area",
      "sheet-menu-print",
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

describe("SpreadsheetMenuBar's Insert and Arrange menus (SPEC-131)", () => {
  it("Insert offers Image… and Shapes, routing insert.image and insert.shape, disabled where drawings cannot go in", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { canInsertDrawings: true } } });
    const menu = wrapper.get('[data-testid="menu-insert"]');
    expect(menu.findAll('[role="menuitem"]').map((item) => item.attributes("data-testid"))).toEqual([
      "sheet-menu-insert-image",
      "sheet-menu-insert-shape",
      "sheet-menu-insert-chart",
      "sheet-menu-insert-table",
      "sheet-menu-insert-link",
      "sheet-menu-insert-note",
      "sheet-menu-insert-comment",
      "sheet-menu-insert-page-break",
      "sheet-menu-remove-page-break",
      "sheet-menu-reset-page-breaks",
    ]);
    // SPEC-136 ends the menu with the page breaks.
    expect(menu.findAll(".menu-item-label").map((label) => label.text())).toEqual([
      "Image…",
      "Shapes",
      "Chart",
      "Table…",
      "Link…",
      "Note",
      "Comment",
      "Page break",
      "Remove page break",
      "Reset all page breaks",
    ]);
    await wrapper.get('[data-testid="sheet-menu-insert-image"]').trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertImage]);
    await wrapper.get('[data-testid="sheet-menu-insert-shape"]').trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertShape]);
    // SPEC-132: Insert ▸ Chart opens the chart gallery.
    await wrapper.get('[data-testid="sheet-menu-insert-chart"]').trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertChart]);
    expect(SHEET_MENU_COMMANDS.insertChart).toBe("insert.chart");
    const csv = mount(SpreadsheetMenuBar, { props: { state: { canInsertDrawings: false } } });
    for (const testid of ["sheet-menu-insert-image", "sheet-menu-insert-shape", "sheet-menu-insert-chart"]) {
      expect(csv.get(`[data-testid="${testid}"]`).attributes("disabled")).toBeDefined();
    }
  });

  it("SPEC-133: Insert ▸ Table… routes insert.table, has no shortcut and is disabled off a grid sheet", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { canInsertDrawings: false } } });
    const item = wrapper.get('[data-testid="sheet-menu-insert-table"]');
    expect(item.attributes("disabled")).toBeUndefined();
    expect(item.find(".menu-shortcut").exists()).toBe(false);
    await item.trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertTable]);
    expect(SHEET_MENU_COMMANDS.insertTable).toBe("insert.table");
    const foreign = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: false } } });
    expect(foreign.get('[data-testid="sheet-menu-insert-table"]').attributes("disabled")).toBeDefined();
  });

  it("SPEC-135: Insert ▸ Link… routes insert.link with Ctrl/Cmd+K and is disabled off a grid sheet", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: true } } });
    const item = wrapper.get('[data-testid="sheet-menu-insert-link"]');
    expect(item.attributes("disabled")).toBeUndefined();
    expect(item.get(".menu-shortcut").text()).toBe(shortcutLabel("sheet.insertLink", isMacPlatform(), SHEET_SHORTCUTS));
    await item.trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertLink]);
    expect(SHEET_MENU_COMMANDS.insertLink).toBe("insert.link");
    const foreign = mount(SpreadsheetMenuBar, { props: { state: { hasGrid: false } } });
    expect(foreign.get('[data-testid="sheet-menu-insert-link"]').attributes("disabled")).toBeDefined();
  });

  it("SPEC-135: Insert ▸ Note (Shift+F2) and Insert ▸ Comment (Ctrl+Alt+M) route insert.note and insert.comment; Comment is disabled with its reason in an .ods workbook", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: {} } });
    const note = wrapper.get('[data-testid="sheet-menu-insert-note"]');
    const comment = wrapper.get('[data-testid="sheet-menu-insert-comment"]');
    expect(note.get(".menu-shortcut").text()).toBe(shortcutLabel("sheet.note", isMacPlatform(), SHEET_SHORTCUTS));
    expect(comment.get(".menu-shortcut").text()).toBe(shortcutLabel("sheet.newComment", isMacPlatform(), SHEET_SHORTCUTS));
    await note.trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertNote]);
    await comment.trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.insertComment]);
    expect(SHEET_MENU_COMMANDS).toMatchObject({ insertNote: "insert.note", insertComment: "insert.comment" });
    const ods = mount(SpreadsheetMenuBar, { props: { state: { canComment: false, commentDisabledReason: "OpenDocument spreadsheets support notes only" } } });
    const odsComment = ods.get('[data-testid="sheet-menu-insert-comment"]');
    expect(odsComment.attributes("disabled")).toBeDefined();
    expect(odsComment.attributes("title")).toBe("OpenDocument spreadsheets support notes only");
    expect(ods.get('[data-testid="sheet-menu-insert-note"]').attributes("disabled")).toBeUndefined();
    const csv = mount(SpreadsheetMenuBar, { props: { state: { canAnnotate: false, canComment: false } } });
    expect(csv.get('[data-testid="sheet-menu-insert-note"]').attributes("disabled")).toBeDefined();
    expect(csv.get('[data-testid="sheet-menu-show-all-notes"]').attributes("disabled")).toBeDefined();
  });

  it("SPEC-135: View ends with Comments and Show all notes, checkboxes routing view.comments and view.show-all-notes", async () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: {} } });
    const view = wrapper.get('[data-testid="menu-view"]');
    const items = view.findAll('[role="menuitemcheckbox"]').map((item) => item.attributes("data-testid"));
    expect(items.slice(-2)).toEqual(["sheet-menu-comments", "sheet-menu-show-all-notes"]);
    expect(wrapper.get('[data-testid="sheet-menu-comments"]').attributes("aria-checked")).toBe("false");
    await wrapper.get('[data-testid="sheet-menu-comments"]').trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.comments]);
    await wrapper.get('[data-testid="sheet-menu-show-all-notes"]').trigger("click");
    expect(wrapper.emitted("command")?.at(-1)).toEqual([SHEET_MENU_COMMANDS.showAllNotes]);
    const on = mount(SpreadsheetMenuBar, { props: { state: { commentsOpen: true, showAllNotes: true } } });
    expect(on.get('[data-testid="sheet-menu-comments"]').attributes("aria-checked")).toBe("true");
    expect(on.get('[data-testid="sheet-menu-show-all-notes"]').attributes("aria-checked")).toBe("true");
    expect(asSpreadsheetMenuState(undefined)).toMatchObject({ canAnnotate: true, canComment: true, commentDisabledReason: "", commentsOpen: false, showAllNotes: false });
  });

  it("Arrange offers the four z-order moves and Alt Text…, disabled without a selected drawing and at the z-order's ends", async () => {
    const none = mount(SpreadsheetMenuBar, { props: { state: {} } });
    const menu = none.get('[data-testid="menu-arrange"]');
    const rows = menu.findAll('[role="menuitem"], hr').map((row) => (row.element.tagName === "HR" ? "—" : row.attributes("data-testid")));
    expect(rows).toEqual(["sheet-menu-arrange-front", "sheet-menu-arrange-forward", "sheet-menu-arrange-backward", "sheet-menu-arrange-back", "—", "sheet-menu-arrange-alt-text"]);
    expect(menu.findAll(".menu-item-label").map((label) => label.text())).toEqual(["Bring to Front", "Bring Forward", "Send Backward", "Send to Back", "Alt Text…"]);
    for (const row of menu.findAll('[role="menuitem"]')) {
      expect(row.attributes("disabled")).toBeDefined();
    }
    const middle = mount(SpreadsheetMenuBar, { props: { state: { drawingSelected: true } } });
    for (const [testid, command] of [
      ["sheet-menu-arrange-front", SHEET_MENU_COMMANDS.arrangeFront],
      ["sheet-menu-arrange-forward", SHEET_MENU_COMMANDS.arrangeForward],
      ["sheet-menu-arrange-backward", SHEET_MENU_COMMANDS.arrangeBackward],
      ["sheet-menu-arrange-back", SHEET_MENU_COMMANDS.arrangeBack],
      ["sheet-menu-arrange-alt-text", SHEET_MENU_COMMANDS.altText],
    ]) {
      const item = middle.get(`[data-testid="${testid}"]`);
      expect(item.attributes("disabled")).toBeUndefined();
      await item.trigger("click");
      expect(middle.emitted("command")?.at(-1)).toEqual([command]);
    }
    const front = mount(SpreadsheetMenuBar, { props: { state: { drawingSelected: true, drawingAtFront: true } } });
    expect(front.get('[data-testid="sheet-menu-arrange-front"]').attributes("disabled")).toBeDefined();
    expect(front.get('[data-testid="sheet-menu-arrange-forward"]').attributes("disabled")).toBeDefined();
    expect(front.get('[data-testid="sheet-menu-arrange-back"]').attributes("disabled")).toBeUndefined();
    const back = mount(SpreadsheetMenuBar, { props: { state: { drawingSelected: true, drawingAtBack: true } } });
    expect(back.get('[data-testid="sheet-menu-arrange-backward"]').attributes("disabled")).toBeDefined();
    expect(back.get('[data-testid="sheet-menu-arrange-back"]').attributes("disabled")).toBeDefined();
    expect(back.get('[data-testid="sheet-menu-arrange-front"]').attributes("disabled")).toBeUndefined();
  });

  it("defaults: drawings can go in, none is selected", () => {
    expect(asSpreadsheetMenuState(undefined)).toMatchObject({ canInsertDrawings: true, drawingSelected: false, drawingAtFront: false, drawingAtBack: false });
    expect(SHEET_MENU_COMMANDS).toMatchObject({ insertImage: "insert.image", insertShape: "insert.shape", arrangeBack: "arrange.back", altText: "arrange.alt-text" });
  });
});

describe("SpreadsheetMenuBar while the workbook is read-only (SPEC-138)", () => {
  it("disables every item that changes the workbook and keeps what views, downloads, prints or closes it", () => {
    const wrapper = mount(SpreadsheetMenuBar, { props: { state: { readOnly: true, filterOn: true, frozen: true, drawingSelected: true, canInsertPageBreak: true, hasPrintArea: true } } });
    const disabled = (testid: string) => wrapper.get(`[data-testid="${testid}"]`).attributes("disabled") !== undefined;
    // (A submenu's items render once it opens; they carry the same rule.)
    for (const testid of [
      "sheet-menu-page-setup", "sheet-menu-paste-special", "sheet-menu-insert-image", "sheet-menu-insert-chart", "sheet-menu-insert-table",
      "sheet-menu-insert-link", "sheet-menu-insert-note", "sheet-menu-insert-page-break", "sheet-menu-arrange-front", "sheet-menu-arrange-alt-text",
      "sheet-menu-format-cells", "sheet-menu-conditional-formatting", "sheet-menu-row-height", "sheet-menu-name-manager",
      "sheet-menu-sort-asc", "sheet-menu-filter", "sheet-menu-clear-filters", "sheet-menu-reapply-filter", "sheet-menu-data-validation",
    ]) {
      expect(disabled(testid), testid).toBe(true);
    }
    for (const testid of ["sheet-menu-download", "sheet-menu-print", "sheet-menu-version-history", "sheet-menu-close", "sheet-menu-gridlines", "sheet-menu-show-formulas", "sheet-menu-page-breaks", "sheet-menu-recalculate"]) {
      expect(disabled(testid), testid).toBe(false);
    }
    // With a filter on, "Turn on Filter first" is not why Clear filters is off.
    expect(wrapper.get('[data-testid="sheet-menu-clear-filters"]').text()).not.toContain(FILTER_OFF_REASON);
    expect(asSpreadsheetMenuState(undefined).readOnly).toBe(false);
  });
});
