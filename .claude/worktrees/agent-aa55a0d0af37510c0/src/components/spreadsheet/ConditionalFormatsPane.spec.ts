// SPEC-129 "The side panel" — the conditional-formatting rules manager
// (`cf-manager`) over a real model and the view's real undo stack: rules
// listed in priority order with their descriptions and previews, the scope,
// adding through a preset (directly, or through the editor for a value
// preset — the e2e's path), editing, deleting, moving, clearing, "Applies
// to" (an invalid entry saves nothing) and Stop if true, each action ONE undo
// step labelled "Conditional formatting" and a no-op none.
import { afterEach, describe, expect, it } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { effectScope, shallowRef, triggerRef } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { defaultIconThresholds, type ConditionalFormat } from "../../lib/spreadsheet/conditionalFormats";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useConditionalFormatManager } from "../../views/spreadsheet/useConditionalFormatManager";
import { useSpreadsheetUndo } from "../../views/spreadsheet/useSpreadsheetUndo";
import ConditionalFormatsPane from "./ConditionalFormatsPane.vue";

const range = (top: number, left: number, bottom = top, right = left): GridRange => ({ top, left, bottom, right });
const LIGHT_RED = { fill: { kind: "solid" as const, color: "#FFC7CE" }, font: { color: "#9C0006" } };

function seeded(format: SpreadsheetModel["format"] = "xlsx"): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.format = format;
  model.dxfs = [LIGHT_RED];
  const formats: ConditionalFormat[] = [
    { id: "f1", ranges: [range(1, 1, 12, 1)], rules: [{ id: "r1", priority: 1, kind: "cellIs", operator: "greaterThan", formulas: ["100"], dxfId: 0 }] },
    {
      id: "f2",
      ranges: [range(1, 3, 12, 3)],
      rules: [{ id: "r2", priority: 2, kind: "iconSet", icons: { set: "3Arrows", thresholds: defaultIconThresholds(3), reverse: false, showValue: true } }],
    },
  ];
  model.sheets = model.sheets.map((sheet) => ({ ...sheet, conditionalFormats: formats }));
  return model;
}

const scopes: (() => void)[] = [];
const wrappers: VueWrapper[] = [];
afterEach(() => {
  wrappers.splice(0).forEach((wrapper) => wrapper.unmount());
  scopes.splice(0).forEach((stop) => stop());
});

function setup(model = seeded()) {
  const book = shallowRef<SpreadsheetModel | null>(model);
  const selection = shallowRef<{ ranges: GridRange[]; active: CellRef }>({ ranges: [range(1, 1, 12, 1)], active: { row: 1, col: 1 } });
  const scope = effectScope();
  scopes.push(() => scope.stop());
  const undo = scope.run(() =>
    useSpreadsheetUndo({ model: book, recalc: (_changed, done) => done([]), onApplied: () => triggerRef(book), activateSheet: () => undefined }),
  )!;
  const manager = scope.run(() =>
    useConditionalFormatManager({
      model: () => book.value,
      sheet: () => book.value?.sheets[book.value.activeSheetIndex] ?? null,
      selection: () => selection.value,
      undo,
      locale: () => SPREADSHEET_LOCALES["en-US"],
    }),
  )!;
  const wrapper = mount(ConditionalFormatsPane, { props: { manager }, attachTo: document.body });
  wrappers.push(wrapper);
  const formats = () => book.value!.sheets[0]!.conditionalFormats;
  const descriptions = () => wrapper.findAll('[data-testid="cf-rule-description"]').map((each) => each.text());
  const click = async (testid: string) => {
    await wrapper.get(`[data-testid="${testid}"]`).trigger("click");
  };
  const selectRule = async (index: number) => {
    await wrapper.findAll('[data-testid="cf-rule"]')[index]!.trigger("click");
  };
  return { book, selection, undo, manager, wrapper, formats, descriptions, click, selectRule };
}

describe("ConditionalFormatsPane (SPEC-129)", () => {
  it("lists the sheet's rules in priority order with their previews and descriptions", async () => {
    const { wrapper, descriptions, selection } = setup();
    expect(wrapper.get('[data-testid="cf-rule-list"]').attributes("role")).toBe("listbox");
    const rules = wrapper.findAll('[data-testid="cf-rule"]');
    expect(rules.map((rule) => rule.attributes("role"))).toEqual(["option", "option"]);
    expect(descriptions()).toEqual(["Cell value > 100", "Icon set (3 Arrows)"]);
    expect(rules[0]!.get('[data-testid="cf-rule-preview"]').attributes("data-kind")).toBe("format");
    expect(rules[0]!.get('[data-testid="cf-rule-preview"]').text()).toBe("AaBbCc");
    expect(rules[1]!.findAll('[data-testid="sheet-cf-icon"]')).toHaveLength(3);
    expect((rules[0]!.get('[data-testid="cf-rule-applies-to"]').element as HTMLInputElement).value).toBe("B2:B13");
    // Only formatting rules offer Stop if true.
    expect(rules[0]!.find('[data-testid="cf-rule-stop-if-true"]').exists()).toBe(true);
    expect(rules[1]!.find('[data-testid="cf-rule-stop-if-true"]').exists()).toBe(false);
    // "Current selection" lists the rules whose ranges meet the selection.
    await wrapper.get('[data-testid="cf-scope"]').setValue("selection");
    expect(descriptions()).toEqual(["Cell value > 100"]);
    selection.value = { ranges: [range(4, 3)], active: { row: 4, col: 3 } };
    await wrapper.vm.$nextTick();
    expect(descriptions()).toEqual(["Icon set (3 Arrows)"]);
  });

  it("adds a direct preset over the current selection at priority 1, selected, as one undo step", async () => {
    const { wrapper, selection, click, formats, descriptions, undo, book } = setup();
    selection.value = { ranges: [range(4, 4, 5, 4)], active: { row: 4, col: 4 } };
    await click("cf-add");
    expect(wrapper.find('[data-testid="cf-preset-data-bar-blue-gradient"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="cf-preset-icon-set-3Arrows"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="cf-preset-color-scale-6"]').exists()).toBe(true);
    await click("cf-preset-above-average");
    expect(wrapper.find('[data-testid="cf-presets"]').exists()).toBe(false);
    // A press outside the open menu closes it without adding anything.
    await click("cf-add");
    document.body.dispatchEvent(new Event("pointerdown", { bubbles: true }));
    await wrapper.vm.$nextTick();
    expect(wrapper.find('[data-testid="cf-presets"]').exists()).toBe(false);
    expect(descriptions()).toEqual(["Above average", "Cell value > 100", "Icon set (3 Arrows)"]);
    const added = formats().find((format) => format.ranges[0]!.top === 4)!;
    expect(added.ranges).toEqual([range(4, 4, 5, 4)]);
    expect(added.rules[0]).toMatchObject({ kind: "aboveAverage", priority: 1, dxfId: 0 });
    expect(wrapper.findAll('[data-testid="cf-rule"]')[0]!.attributes("aria-selected")).toBe("true");
    expect(undo.undoLabel.value).toBe("Conditional formatting");
    undo.undo();
    await wrapper.vm.$nextTick();
    expect(descriptions()).toEqual(["Cell value > 100", "Icon set (3 Arrows)"]);
    expect(book.value!.dxfs).toHaveLength(1);
  });

  it("adds through the editor: Greater than… 120, green, Save — then moves it below the file's rule", async () => {
    const { wrapper, click, formats, descriptions, book, undo } = setup();
    await click("cf-add");
    await click("cf-preset-greater-than");
    expect(wrapper.find('[data-testid="cf-editor"]').exists()).toBe(true);
    expect((wrapper.get('[data-testid="cf-operator"]').element as HTMLSelectElement).value).toBe("greaterThan");
    await wrapper.get('[data-testid="cf-value-1"]').setValue("120");
    const format = wrapper.get('[data-testid="cf-format"]');
    const green = format.findAll("option").find((option) => option.text() === "Green fill with dark green text")!;
    await format.setValue(green.attributes("value"));
    await click("cf-save");
    expect(wrapper.find('[data-testid="cf-editor"]').exists()).toBe(false);
    expect(descriptions()).toEqual(["Cell value > 120", "Cell value > 100", "Icon set (3 Arrows)"]);
    const rule = formats().flatMap((each) => each.rules).find((each) => each.kind === "cellIs" && each.formulas[0] === "120")!;
    expect(rule.priority).toBe(1);
    expect(book.value!.dxfs[rule.dxfId!]).toEqual({ fill: { kind: "solid", color: "#C6EFCE" }, font: { color: "#006100" } });
    expect(wrapper.findAll('[data-testid="cf-rule"]')[0]!.attributes("aria-selected")).toBe("true");
    expect(undo.undoLabel.value).toBe("Conditional formatting");
    await click("cf-move-down");
    expect(descriptions()).toEqual(["Cell value > 100", "Cell value > 120", "Icon set (3 Arrows)"]);
    expect(wrapper.findAll('[data-testid="cf-rule"]')[1]!.attributes("aria-selected")).toBe("true");
    await click("cf-move-up");
    expect(descriptions()).toEqual(["Cell value > 120", "Cell value > 100", "Icon set (3 Arrows)"]);
    // At the top, Move up is off — and a no-op records nothing.
    expect(wrapper.get('[data-testid="cf-move-up"]').attributes("disabled")).toBeDefined();
  });

  it("edits the selected rule in place (same id), and Cancel/Back change nothing", async () => {
    const { wrapper, click, selectRule, descriptions, formats, undo } = setup();
    await selectRule(0);
    await click("cf-edit");
    expect((wrapper.get('[data-testid="cf-value-1"]').element as HTMLInputElement).value).toBe("100");
    await click("cf-editor-back");
    expect(undo.canUndo.value).toBe(false);
    await click("cf-edit");
    await wrapper.get('[data-testid="cf-value-1"]').setValue("150");
    await click("cf-save");
    expect(descriptions()).toEqual(["Cell value > 150", "Icon set (3 Arrows)"]);
    expect(formats()[0]!.rules[0]).toMatchObject({ id: "r1", priority: 1, dxfId: 0, formulas: ["150"] });
    expect(undo.undoLabel.value).toBe("Conditional formatting");
  });

  it("deletes the selected rule, selecting its neighbour; undo restores it", async () => {
    const { wrapper, click, selectRule, descriptions, undo } = setup();
    await selectRule(0);
    await click("cf-delete");
    expect(descriptions()).toEqual(["Icon set (3 Arrows)"]);
    expect(wrapper.findAll('[data-testid="cf-rule"]')[0]!.attributes("aria-selected")).toBe("true");
    undo.undo();
    await wrapper.vm.$nextTick();
    expect(descriptions()).toEqual(["Cell value > 100", "Icon set (3 Arrows)"]);
  });

  it("clears rules from the selected cells, then from the entire sheet", async () => {
    const { wrapper, click, selection, formats, descriptions } = setup();
    selection.value = { ranges: [range(1, 1, 4, 1)], active: { row: 1, col: 1 } };
    await click("cf-clear-selection");
    expect(wrapper.get('[data-testid="cf-clear-selection"]').text()).toBe("Clear rules from selected cells");
    expect(formats()[0]!.ranges).toEqual([range(5, 1, 12, 1)]);
    await click("cf-clear-sheet");
    expect(wrapper.get('[data-testid="cf-clear-sheet"]').text()).toBe("Clear rules from entire sheet");
    expect(formats()).toEqual([]);
    expect(descriptions()).toEqual([]);
  });

  it("commits Applies to on change; an invalid range is an inline error and saves nothing", async () => {
    const { wrapper, formats, undo } = setup();
    const input = () => wrapper.findAll('[data-testid="cf-rule-applies-to"]')[0]!;
    await input().setValue("B2:B5, D1");
    expect(formats()[0]!.ranges).toEqual([range(1, 1, 4, 1), range(0, 3)]);
    expect(undo.undoLabel.value).toBe("Conditional formatting");
    undo.undo();
    await wrapper.vm.$nextTick();
    await input().setValue("B2:");
    expect(wrapper.get('[data-testid="cf-rule-applies-to-error"]').text()).toContain("B2:");
    expect(formats()[0]!.ranges).toEqual([range(1, 1, 12, 1)]);
    expect(undo.canUndo.value).toBe(false);
    // The same ranges again are no step at all.
    await input().setValue("B2:B13");
    expect(wrapper.find('[data-testid="cf-rule-applies-to-error"]').exists()).toBe(false);
    expect(undo.canUndo.value).toBe(false);
  });

  it("closes the editor and drops the selection when another sheet is shown", async () => {
    const model = seeded();
    model.sheets = [...model.sheets, { ...model.sheets[0]!, id: "other", name: "Other", conditionalFormats: [] }];
    const { wrapper, click, selectRule, book, manager } = setup(model);
    await selectRule(0);
    await click("cf-edit");
    expect(wrapper.find('[data-testid="cf-editor"]').exists()).toBe(true);
    book.value!.activeSheetIndex = 1;
    triggerRef(book);
    await wrapper.vm.$nextTick();
    await wrapper.vm.$nextTick();
    expect(wrapper.find('[data-testid="cf-editor"]').exists()).toBe(false);
    expect(manager.selectedRuleId.value).toBeNull();
    expect(wrapper.find('[data-testid="cf-empty"]').exists()).toBe(true);
  });

  it("sets Stop if true on a formatting rule as one step", async () => {
    const { wrapper, formats, undo } = setup();
    await wrapper.get('[data-testid="cf-rule-stop-if-true"]').setValue(true);
    expect(formats()[0]!.rules[0]!.stopIfTrue).toBe(true);
    expect(undo.undoLabel.value).toBe("Conditional formatting");
    await wrapper.get('[data-testid="cf-rule-stop-if-true"]').setValue(false);
    expect(formats()[0]!.rules[0]!.stopIfTrue).toBeUndefined();
  });
});
