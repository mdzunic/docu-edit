// SPEC-135's useCellLinks: the link dialog (opened on the active cell by
// Ctrl/Cmd+K from the grid, prefilled; a dangerous scheme and a bad
// reference refused inline; Save as ONE "Insert link" / "Edit link" step that
// writes the text through the cell-entry path and gives automatic-coloured
// cells the hyperlink look; Remove), the chip under the active cell, the
// open rule when following (external, refused, a place, a missing place),
// HYPERLINK() formula cells, and Ctrl/Cmd+press over a linked cell's text.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import { getCell } from "../../lib/spreadsheet/address";
import { LINK_DESTINATION_MISSING_MESSAGE, LINK_NOT_OPENABLE_MESSAGE } from "../../lib/spreadsheet/annotations";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import type { CellHyperlink, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { singleCell } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";
import { FORMULA_ERROR_MESSAGE } from "./cellEntry";
import { isAutomaticColor, isLinkedCell, LINK_REF_INVALID_MESSAGE, LINK_SCHEME_NOT_ALLOWED_MESSAGE, linkDialogResult, normaliseAddress, placeFields } from "./cellLinks";
import { HIDDEN_SHEET_MESSAGE } from "./nameBox";
import { linksOf, useCellLinks } from "./useCellLinks";
import type { SheetViewFields } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const LINKS: CellHyperlink[] = [
  { linkId: "h-A6", range: { top: 5, left: 0, bottom: 5, right: 1 }, target: { kind: "url", href: "https://example.com/report" } },
  { linkId: "h-A7", range: { top: 6, left: 0, bottom: 6, right: 0 }, target: { kind: "url", href: "https://example.com/" }, tooltip: "Open the example site" },
  { linkId: "h-B7", range: { top: 6, left: 1, bottom: 6, right: 1 }, target: { kind: "place", location: "Data!A1" } },
  { linkId: "h-A8", range: { top: 7, left: 0, bottom: 7, right: 0 }, target: { kind: "url", href: "mailto:team@example.com" } },
  { linkId: "h-B8", range: { top: 7, left: 1, bottom: 7, right: 1 }, target: { kind: "url", href: "javascript:alert(1)" } },
];

function workbook(): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  const summary: Sheet = {
    ...base,
    id: "s1",
    name: "Summary",
    rows: [
      { index: 0, cells: [{ col: 0, value: "Metric" }] },
      { index: 1, cells: [{ col: 1, value: 1200 }] },
      { index: 5, cells: [{ col: 0, value: "Report" }, { col: 1, value: "Q3" }] },
      { index: 6, cells: [{ col: 0, value: "Example site" }, { col: 1, value: "See data" }, { col: 2, value: "Details" }] },
      { index: 7, cells: [{ col: 0, value: "Email the team" }, { col: 1, value: "Run script" }] },
      { index: 8, cells: [{ col: 3, value: "Go", formula: 'HYPERLINK("#Data!A10","Go")' }, { col: 4, value: "Red", styleId: 1 }] },
    ],
    hyperlinks: LINKS.map((link) => structuredClone(link)),
  };
  const data: Sheet = { ...base, id: "data", name: "Data", rows: [{ index: 9, cells: [{ col: 0, value: "Target" }] }] };
  const secret: Sheet = { ...base, id: "secret", name: "Secret", visibility: "hidden", rows: [] };
  model.sheets = [summary, data, secret];
  model.styles = [...model.styles!, { font: { color: "#ff0000" } }];
  model.names = [
    { name: "Target", formula: "Data!$A$10" },
    { name: "Hidden", formula: "Data!$A$1", hidden: true },
    { name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Summary!$A$1:$D$10" },
    { name: "OnSecret", formula: "Secret!$A$1" },
  ];
  return model;
}

function setup(options: { editing?: boolean } = {}) {
  const scope = effectScope();
  const model = workbook();
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  const viewState = {
    update: (sheetId: string, patch: SheetViewFields) => void (fields.value = { ...fields.value, [sheetId]: { ...fields.value[sheetId], ...patch } }),
    syncInto: () => {},
  };
  const selection = computed(() => fields.value[activeSheet.value?.id ?? ""]?.selection ?? singleCell({ row: 0, col: 0 }));
  const select = (top: number, left: number, bottom = top, right = left) =>
    viewState.update(activeSheet.value!.id, { selection: { active: { row: top, col: left }, ranges: [{ top, left, bottom, right }], activeRangeIndex: 0 } });
  const toast = vi.fn();
  const openWindow = vi.fn();
  const copyText = vi.fn();
  const focusCell = vi.fn();
  const openEdit = vi.fn(() => true);
  const editing = ref(options.editing ?? false);
  const revealTick = ref(0);
  const result = scope.run(() => {
    const undo = useSpreadsheetUndo({ model: wb, recalc: (_changed, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
    const links = useCellLinks({
      book: { workbook: wb, activeIndex, activeSheet, selectSheet: (index) => void (activeIndex.value = index) },
      selection: { selection, formatContext: computed(() => ({ locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" as const })) },
      viewState,
      undo,
      cell: { editing: computed(() => editing.value), focusCell, revealTick, openEdit },
      locale: () => SPREADSHEET_LOCALES["en-US"],
      toast,
      openWindow,
      copyText,
      measurer: fixedMeasurer(7, 20),
    });
    return { undo, links };
  })!;
  const summary = () => wb.value!.sheets[0]!;
  return { ...result, wb, model, activeIndex, selection, select, toast, openWindow, copyText, focusCell, openEdit, editing, revealTick, summary, stop: () => scope.stop() };
}

const flush = async () => {
  await nextTick();
  await nextTick();
  await nextTick();
};

describe("cellLinks — the dialog's rules", () => {
  it("refuses a dangerous scheme with LinkDialog.vue's own string", () => {
    const fields = { text: "", kind: "web" as const, address: "  java\tscript:alert(1)", sheet: "Summary", ref: "A1", name: "", tooltip: "" };
    expect(linkDialogResult(fields)).toEqual({ error: LINK_SCHEME_NOT_ALLOWED_MESSAGE });
    expect(linkDialogResult({ ...fields, address: "data:text/html,x" })).toEqual({ error: LINK_SCHEME_NOT_ALLOWED_MESSAGE });
    const source = readFileSync(join(dirname(fileURLToPath(import.meta.url)), "../../components/LinkDialog.vue"), "utf8");
    expect(source).toContain(`"${LINK_SCHEME_NOT_ALLOWED_MESSAGE}"`);
  });

  it("validates a place's reference with parseRange and builds canonical text; a chosen name wins", () => {
    const place = { text: "x", kind: "place" as const, address: "", sheet: "My Sheet", ref: "b2:c3", name: "", tooltip: "" };
    expect(linkDialogResult(place)).toEqual({ text: "x", target: { kind: "place", location: "'My Sheet'!B2:C3" }, tooltip: "" });
    expect(linkDialogResult({ ...place, ref: "B" })).toEqual({ error: LINK_REF_INVALID_MESSAGE });
    expect(linkDialogResult({ ...place, ref: "" })).toBeNull();
    expect(linkDialogResult({ ...place, ref: "nonsense!", name: "Target" })).toEqual({ text: "x", target: { kind: "place", location: "Target" }, tooltip: "" });
  });

  it("an empty address saves nothing; an empty text shows the target; a bare domain or e-mail address gets its scheme", () => {
    const web = { text: "", kind: "web" as const, address: "", sheet: "S", ref: "A1", name: "", tooltip: " tip " };
    expect(linkDialogResult(web)).toBeNull();
    expect(linkDialogResult({ ...web, address: "example.com/a" })).toEqual({ text: "https://example.com/a", target: { kind: "url", href: "https://example.com/a" }, tooltip: "tip" });
    expect(normaliseAddress("team@example.com")).toBe("mailto:team@example.com");
    expect(normaliseAddress("ftp://x.example")).toBe("ftp://x.example");
    expect(normaliseAddress("docs/report.pdf")).toBe("docs/report.pdf");
  });

  it("splits a place location into the dialog's sheet and reference, or a name", () => {
    expect(placeFields("'My Sheet'!$A$1:$B$2", "Summary")).toEqual({ sheet: "My Sheet", ref: "$A$1:$B$2", name: "" });
    expect(placeFields("Target", "Summary")).toEqual({ sheet: "Summary", ref: "A1", name: "Target" });
    expect(placeFields("#REF!", "Summary")).toEqual({ sheet: "Summary", ref: "A1", name: "#REF!" });
  });

  it("reads no colour and an untransformed Text 1 as automatic", () => {
    expect(isAutomaticColor(undefined)).toBe(true);
    expect(isAutomaticColor({ theme: "dk1" })).toBe(true);
    expect(isAutomaticColor({ theme: "dk1", tint: 0.5 })).toBe(false);
    expect(isAutomaticColor("#000000")).toBe(false);
    expect(isAutomaticColor({ theme: "hlink" })).toBe(false);
  });

  it("a merge's paint-only copy is never linked", () => {
    const at = (row: number, col: number) => row === 1 && col === 2;
    expect(isLinkedCell({ row: 1, col: 2, paintOnly: false }, at)).toBe(true);
    expect(isLinkedCell({ row: 1, col: 2, paintOnly: true }, at)).toBe(false);
    expect(isLinkedCell({ row: 1, col: 3, paintOnly: false }, at)).toBe(false);
    expect(isLinkedCell({ row: 1, col: 2, paintOnly: false }, undefined)).toBe(false);
  });
});

describe("useCellLinks — the dialog", () => {
  it("opens on the active cell with its display text, the visible sheets and the visible names", () => {
    const t = setup();
    t.select(1, 1);
    t.links.openDialog();
    const state = t.links.dialog.value!;
    expect(state.editing).toBe(false);
    expect(state.fields).toEqual({ text: "1200", kind: "web", address: "", sheet: "Summary", ref: "A1", name: "", tooltip: "" });
    expect(state.sheets).toEqual(["Summary", "Data"]);
    expect(state.names.map((name) => name.value)).toEqual(["OnSecret", "Target"]);
    expect(t.links.canSave.value).toBe(false);
    t.stop();
  });

  it("opens with Ctrl/Cmd+K from the grid only, never while a cell is edited", () => {
    const t = setup();
    const grid = document.createElement("div");
    grid.dataset.testid = "spreadsheet-grid";
    const editor = document.createElement("div");
    grid.append(editor);
    document.body.append(grid);
    const press = (target: Element) => target.dispatchEvent(new KeyboardEvent("keydown", { key: "k", code: "KeyK", ctrlKey: true, metaKey: true, bubbles: true, cancelable: true }));
    press(document.body);
    expect(t.links.dialog.value).toBeNull();
    t.editing.value = true;
    press(editor);
    expect(t.links.dialog.value).toBeNull();
    t.editing.value = false;
    press(editor);
    expect(t.links.dialog.value).not.toBeNull();
    grid.remove();
    t.stop();
  });

  it("shows the refusal of a dangerous scheme inline and keeps Save off", () => {
    const t = setup();
    t.select(6, 2);
    t.links.openDialog();
    t.links.setField("address", "javascript:alert(1)");
    expect(t.links.error.value).toBe(LINK_SCHEME_NOT_ALLOWED_MESSAGE);
    expect(t.links.canSave.value).toBe(false);
    expect(t.links.save()).toBe(false);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("saves a place link as ONE 'Insert link' step: the hyperlink, the hyperlink look on an automatic colour, the dialog shut", () => {
    const t = setup();
    t.select(6, 2);
    t.links.openDialog();
    t.links.setField("kind", "place");
    t.links.setField("sheet", "Data");
    t.links.setField("ref", "a10");
    expect(t.links.canSave.value).toBe(true);
    expect(t.links.save()).toBe(true);
    expect(t.links.dialog.value).toBeNull();
    expect(t.undo.undoLabel.value).toBe("Insert link");
    const link = t.summary().hyperlinks!.find((each) => each.range.top === 6 && each.range.left === 2)!;
    expect(link.target).toEqual({ kind: "place", location: "Data!A10" });
    expect(t.summary().annotationsDirty).toBe(true);
    expect(getCell(t.summary(), 6, 2)?.value).toBe("Details");
    const font = cellStyle(t.wb.value!, effectiveStyleId(t.summary(), 6, 2)).font;
    expect(font?.color).toEqual({ theme: "hlink" });
    expect(font?.underline).toBe("single");
    t.undo.undo();
    expect(t.summary().hyperlinks!.some((each) => each.range.left === 2)).toBe(false);
    expect(cellStyle(t.wb.value!, effectiveStyleId(t.summary(), 6, 2)).font?.color).toEqual({ theme: "dk1" });
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("writes a changed text through the cell entry in the same step, and never restyles an explicit colour", () => {
    const t = setup();
    t.select(8, 4, 8, 5);
    t.links.openDialog();
    t.links.setField("text", "42");
    t.links.setField("address", "https://example.com/red");
    t.links.save();
    expect(getCell(t.summary(), 8, 4)?.value).toBe(42);
    const link = t.summary().hyperlinks!.find((each) => each.linkId === "h-E9")!;
    expect(link.range).toEqual({ top: 8, left: 4, bottom: 8, right: 5 });
    expect(cellStyle(t.wb.value!, effectiveStyleId(t.summary(), 8, 4)).font?.color).toBe("#ff0000");
    expect(cellStyle(t.wb.value!, effectiveStyleId(t.summary(), 8, 5)).font?.color).toEqual({ theme: "hlink" });
    t.undo.undo();
    expect(getCell(t.summary(), 8, 4)?.value).toBe("Red");
    expect(t.summary().hyperlinks).toEqual(LINKS);
    t.stop();
  });

  it("keeps the dialog open with the cell entry's error and changes nothing", () => {
    const t = setup();
    t.select(6, 2);
    t.links.openDialog();
    t.links.setField("text", "=1+");
    t.links.setField("address", "https://example.com/");
    expect(t.links.save()).toBe(false);
    expect(t.links.error.value).toBe(FORMULA_ERROR_MESSAGE);
    expect(t.undo.canUndo.value).toBe(false);
    t.links.setField("text", "Details");
    expect(t.links.error.value).toBe("");
    t.stop();
  });

  it("edits a stored link as 'Edit link' over its own range, prefilled, keeping its id and look; Remove is 'Remove link'", () => {
    const t = setup();
    t.select(5, 0);
    t.links.openDialog();
    expect(t.links.dialog.value!.editing).toBe(true);
    expect(t.links.dialog.value!.fields).toMatchObject({ text: "Report", kind: "web", address: "https://example.com/report" });
    t.links.setField("tooltip", "The report");
    t.links.save();
    expect(t.undo.undoLabel.value).toBe("Edit link");
    const edited = t.summary().hyperlinks!.find((each) => each.linkId === "h-A6")!;
    expect(edited).toEqual({ linkId: "h-A6", range: { top: 5, left: 0, bottom: 5, right: 1 }, target: { kind: "url", href: "https://example.com/report" }, tooltip: "The report" });
    expect(cellStyle(t.wb.value!, effectiveStyleId(t.summary(), 5, 0)).font?.color).toEqual({ theme: "dk1" });
    t.select(6, 0);
    t.links.openDialog();
    expect(t.links.dialog.value!.fields.tooltip).toBe("Open the example site");
    t.links.remove();
    expect(t.undo.undoLabel.value).toBe("Remove link");
    expect(t.summary().hyperlinks!.some((each) => each.linkId === "h-A7")).toBe(false);
    expect(t.links.dialog.value).toBeNull();
    t.stop();
  });

  it("prefills a place link's sheet and reference", () => {
    const t = setup();
    t.select(6, 1);
    t.links.openDialog();
    expect(t.links.dialog.value!.fields).toMatchObject({ text: "See data", kind: "place", sheet: "Data", ref: "A1", name: "" });
    t.links.cancel();
    expect(t.links.dialog.value).toBeNull();
    expect(t.focusCell).toHaveBeenCalled();
    t.stop();
  });
});

describe("useCellLinks — following", () => {
  it("opens http(s) and mailto in a new unlinked context and refuses any other scheme or a relative path", async () => {
    const t = setup();
    const host = { sheetId: "s1", row: 0, col: 0 };
    await t.links.follow({ kind: "url", href: "https://example.com/", subaddress: "top" }, host);
    expect(t.openWindow).toHaveBeenLastCalledWith("https://example.com/#top");
    await t.links.follow({ kind: "url", href: "mailto:team@example.com" }, host);
    expect(t.openWindow).toHaveBeenLastCalledWith("mailto:team@example.com");
    for (const href of ["javascript:alert(1)", "file:///etc/passwd", "docs/report.pdf"]) {
      t.toast.mockClear();
      await t.links.follow({ kind: "url", href }, host);
      expect(t.toast).toHaveBeenCalledWith(LINK_NOT_OPENABLE_MESSAGE);
    }
    expect(t.openWindow).toHaveBeenCalledTimes(2);
    t.stop();
  });

  it("the default opener is window.open with noopener,noreferrer", async () => {
    const scope = effectScope();
    const spy = vi.spyOn(window, "open").mockReturnValue(null);
    const wb = shallowRef<SpreadsheetModel | null>(workbook());
    const links = scope.run(() =>
      useCellLinks({
        book: { workbook: wb, activeIndex: ref(0), activeSheet: computed(() => wb.value!.sheets[0]!), selectSheet: () => {} },
        selection: { selection: computed(() => singleCell({ row: 0, col: 0 })), formatContext: computed(() => ({ locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" as const })) },
        viewState: { update: () => {}, syncInto: () => {} },
        undo: { record: () => {}, commit: () => {} },
        cell: { editing: computed(() => false), focusCell: () => {}, revealTick: ref(0), openEdit: () => true },
        locale: () => SPREADSHEET_LOCALES["en-US"],
        toast: () => {},
      }),
    )!;
    await links.follow({ kind: "url", href: "https://example.com/" }, { sheetId: "s1", row: 0, col: 0 });
    expect(spy).toHaveBeenCalledWith("https://example.com/", "_blank", "noopener,noreferrer");
    spy.mockRestore();
    scope.stop();
  });

  it("a place activates its sheet, selects its range and reveals it; a name selects its range", async () => {
    const t = setup();
    await t.links.follow({ kind: "place", location: "Data!A10:B11" }, { sheetId: "s1", row: 6, col: 1 });
    await flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value).toEqual({ active: { row: 9, col: 0 }, ranges: [{ top: 9, left: 0, bottom: 10, right: 1 }], activeRangeIndex: 0 });
    expect(t.revealTick.value).toBe(1);
    expect(t.focusCell).toHaveBeenCalled();
    t.activeIndex.value = 0;
    await t.links.follow({ kind: "place", location: "Target" }, { sheetId: "s1", row: 0, col: 0 });
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value.active).toEqual({ row: 9, col: 0 });
    expect(t.toast).not.toHaveBeenCalled();
    t.stop();
  });

  it("an unresolvable place — a missing sheet, #REF!, an unknown name — says its destination doesn't exist; a hidden sheet says so", async () => {
    const t = setup();
    for (const location of ["Nowhere!A1", "#REF!", "NoSuchName"]) {
      t.toast.mockClear();
      await t.links.follow({ kind: "place", location }, { sheetId: "s1", row: 0, col: 0 });
      expect(t.toast, location).toHaveBeenCalledWith(LINK_DESTINATION_MISSING_MESSAGE);
    }
    t.toast.mockClear();
    await t.links.follow({ kind: "place", location: "Secret!A1" }, { sheetId: "s1", row: 0, col: 0 });
    expect(t.toast).toHaveBeenCalledWith(HIDDEN_SHEET_MESSAGE);
    expect(t.activeIndex.value).toBe(0);
    t.stop();
  });
});

describe("useCellLinks — the chip, HYPERLINK() cells and the pointer", () => {
  it("shows the active cell's link until dismissed, never while a cell is edited or the dialog is open", () => {
    const t = setup();
    expect(t.links.chip.value).toBeNull();
    t.select(5, 1);
    expect(t.links.chip.value).toEqual({ label: "https://example.com/report", external: true });
    t.select(6, 1);
    expect(t.links.chip.value).toEqual({ label: "Data!A1", external: false });
    t.editing.value = true;
    expect(t.links.chip.value).toBeNull();
    t.editing.value = false;
    t.links.openDialog();
    expect(t.links.chip.value).toBeNull();
    t.links.cancel();
    t.links.dismissChip();
    expect(t.links.chip.value).toBeNull();
    t.select(7, 0);
    expect(t.links.chip.value).toEqual({ label: "mailto:team@example.com", external: true });
    t.stop();
  });

  it("the chip's Open follows, Copy copies the label, Edit opens the dialog, Remove removes a stored link", async () => {
    const t = setup();
    t.select(6, 1);
    t.links.openFromChip();
    await flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value.active).toEqual({ row: 0, col: 0 });
    t.activeIndex.value = 0;
    t.select(6, 0);
    t.links.copyFromChip();
    expect(t.copyText).toHaveBeenCalledWith("https://example.com/");
    t.links.editFromChip();
    expect(t.links.dialog.value?.editing).toBe(true);
    t.links.cancel();
    t.links.removeFromChip();
    expect(t.undo.undoLabel.value).toBe("Remove link");
    expect(t.summary().hyperlinks!.some((each) => each.linkId === "h-A7")).toBe(false);
    t.stop();
  });

  it("treats a top-level HYPERLINK() call as linked — data-link, chip, Edit in the formula — without storing it", async () => {
    const t = setup();
    expect(t.links.isLinked(8, 3)).toBe(true);
    expect(t.links.isLinked(6, 2)).toBe(false);
    expect(t.links.isLinked(5, 1)).toBe(true);
    t.select(8, 3);
    expect(t.links.chip.value).toEqual({ label: "Data!A10", external: false });
    t.links.editFromChip();
    expect(t.openEdit).toHaveBeenCalled();
    t.links.removeFromChip();
    expect(t.links.chip.value).toBeNull();
    expect(t.undo.canUndo.value).toBe(false);
    expect(t.summary().hyperlinks).toEqual(LINKS);
    t.links.openFromChip();
    await flush();
    expect(t.activeIndex.value).toBe(1);
    expect(t.selection.value.active).toEqual({ row: 9, col: 0 });
    t.stop();
  });

  it("Ctrl/Cmd+press over a linked cell's text follows; elsewhere, unmodified or with Shift it is left to the selection", () => {
    const t = setup();
    const cell = document.createElement("div");
    cell.dataset.testid = "sheet-cell";
    cell.dataset.ref = "B8";
    const text = document.createElement("span");
    text.className = "sheet-cell-link";
    cell.append(text);
    document.body.append(cell);
    const press = (target: Element, init: MouseEventInit) => {
      const event = new MouseEvent("pointerdown", { button: 0, bubbles: true, cancelable: true, ...init });
      Object.defineProperty(event, "target", { value: target });
      return { taken: t.links.pressFollows(event), prevented: event.defaultPrevented };
    };
    const mod = { ctrlKey: true, metaKey: true };
    expect(press(cell, mod).taken).toBe(false);
    expect(press(text, {}).taken).toBe(false);
    expect(press(text, { ...mod, shiftKey: true }).taken).toBe(false);
    expect(t.toast).not.toHaveBeenCalled();
    expect(press(text, mod)).toEqual({ taken: true, prevented: true });
    expect(t.toast).toHaveBeenCalledWith(LINK_NOT_OPENABLE_MESSAGE);
    expect(t.openWindow).not.toHaveBeenCalled();
    cell.dataset.ref = "C7";
    expect(press(text, mod).taken).toBe(false);
    cell.remove();
    t.stop();
  });

  it("linksOf narrows the grid context's links, null when not wired", () => {
    const t = setup();
    expect(linksOf({ links: t.links })).toBe(t.links);
    expect(linksOf({})).toBeNull();
    expect(linksOf(null)).toBeNull();
    t.stop();
  });
});
