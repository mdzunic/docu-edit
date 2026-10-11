// SPEC-129 "The rule editor" — `cf-editor` over useConditionalFormatEditor:
// an unparseable value or formula is an inline error and saves nothing;
// "Use a formula…" is typed relative to the active cell and stored relative
// to the format's origin; "Reverse icon order" is absent for an `.ods`
// workbook; "Custom format…" opens FormatCellsDialog in its differential
// mode and its OK becomes the rule's dxf; the live preview follows the
// chosen format.
import { afterEach, describe, expect, it, vi } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { effectScope } from "vue";
import { DEFAULT_THEME } from "../../lib/core/theme";
import { defaultIconThresholds, type CfRuleDraft } from "../../lib/spreadsheet/conditionalFormats";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { useConditionalFormatEditor, type CfEditorSave } from "../../views/spreadsheet/useConditionalFormatEditor";
import ConditionalFormatEditor from "./ConditionalFormatEditor.vue";

const scopes: (() => void)[] = [];
const wrappers: VueWrapper[] = [];
afterEach(() => {
  wrappers.splice(0).forEach((wrapper) => wrapper.unmount());
  scopes.splice(0).forEach((stop) => stop());
});

function setup(draft: CfRuleDraft, options: { format?: string; active?: { row: number; col: number } } = {}) {
  const onSave = vi.fn<(save: CfEditorSave) => boolean>(() => true);
  const scope = effectScope();
  scopes.push(() => scope.stop());
  const editor = scope.run(() =>
    useConditionalFormatEditor({
      theme: () => DEFAULT_THEME,
      locale: () => SPREADSHEET_LOCALES["en-US"],
      dateSystem: () => "1900",
      format: () => options.format ?? "xlsx",
      onSave,
    }),
  )!;
  editor.start({ ruleId: null, draft, ranges: [{ top: 1, left: 1, bottom: 12, right: 1 }], active: options.active ?? { row: 1, col: 1 } });
  const wrapper = mount(ConditionalFormatEditor, { props: { editor }, attachTo: document.body });
  wrappers.push(wrapper);
  return { editor, wrapper, onSave };
}

const ICONS: CfRuleDraft = { kind: "iconSet", icons: { set: "3Arrows", thresholds: defaultIconThresholds(3), reverse: false, showValue: true } };

describe("ConditionalFormatEditor (SPEC-129)", () => {
  it("offers Excel's six rule types", () => {
    const { wrapper } = setup({ kind: "cellIs", operator: "greaterThan", formulas: [""] });
    expect(wrapper.get('[data-testid="cf-rule-type"]').findAll("option").map((option) => option.text())).toEqual([
      "Format all cells based on their values",
      "Format only cells that contain",
      "Format only top or bottom ranked values",
      "Format only values that are above or below average",
      "Format only unique or duplicate values",
      "Use a formula to determine which cells to format",
    ]);
  });

  it("an unparseable formula is an inline error and saves nothing", async () => {
    const { wrapper, onSave } = setup({ kind: "expression", formula: "" });
    await wrapper.get('[data-testid="cf-value-1"]').setValue("=SUM(B2");
    await wrapper.get('[data-testid="cf-save"]').trigger("click");
    expect(wrapper.get('[data-testid="cf-editor-error"]').text()).toBe("This formula isn't valid");
    expect(wrapper.get('[data-testid="cf-value-1"]').attributes("aria-invalid")).toBe("true");
    expect(onSave).not.toHaveBeenCalled();
    expect(wrapper.find('[data-testid="cf-editor"]').exists()).toBe(true);
  });

  it("an empty or unparseable cell value is an inline error and saves nothing", async () => {
    const { wrapper, onSave } = setup({ kind: "cellIs", operator: "between", formulas: ["", ""] });
    await wrapper.get('[data-testid="cf-value-1"]').setValue("1");
    await wrapper.get('[data-testid="cf-save"]').trigger("click");
    expect(wrapper.get('[data-testid="cf-editor-error"]').text()).toBe("Enter a value");
    await wrapper.get('[data-testid="cf-value-2"]').setValue("=1+");
    await wrapper.get('[data-testid="cf-save"]').trigger("click");
    expect(wrapper.get('[data-testid="cf-editor-error"]').text()).toBe("This formula isn't valid");
    expect(onSave).not.toHaveBeenCalled();
  });

  it("stores a formula typed at the active cell relative to the format's origin", async () => {
    const { wrapper, onSave } = setup({ kind: "expression", formula: "" }, { active: { row: 4, col: 2 } });
    await wrapper.get('[data-testid="cf-value-1"]').setValue("=C5>$Z$1");
    await wrapper.get('[data-testid="cf-save"]').trigger("click");
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onSave.mock.calls[0]![0].draft).toEqual({ kind: "expression", formula: "B2>$Z$1" });
    expect(onSave.mock.calls[0]![0].dxf).toEqual({ fill: { kind: "solid", color: "#FFC7CE" }, font: { color: "#9C0006" } });
  });

  it("hides Reverse icon order for an .ods workbook", async () => {
    expect(setup(ICONS).wrapper.find('[data-testid="cf-icon-reverse"]').exists()).toBe(true);
    const ods = setup(ICONS, { format: "ods" }).wrapper;
    expect(ods.find('[data-testid="cf-icon-set"]').exists()).toBe(true);
    expect(ods.findAll('[data-testid^="cf-icon-threshold-"]').length).toBeGreaterThan(0);
    expect(ods.find('[data-testid="cf-icon-reverse"]').exists()).toBe(false);
  });

  it("saves an icon set's edited threshold and a colour scale with no dxf", async () => {
    const { wrapper, onSave } = setup(ICONS);
    await wrapper.get('[data-testid="cf-icon-threshold-2"]').setValue("80");
    await wrapper.get('[data-testid="cf-icon-reverse"]').setValue(true);
    await wrapper.get('[data-testid="cf-save"]').trigger("click");
    const saved = onSave.mock.calls[0]![0];
    expect(saved.dxf).toBeNull();
    expect(saved.draft).toMatchObject({ kind: "iconSet", icons: { set: "3Arrows", reverse: true } });
    expect(saved.draft.kind === "iconSet" && saved.draft.icons.thresholds[2]).toEqual({ at: { type: "percent", value: "80" }, gte: true });
  });

  it("Custom format… opens the differential Format Cells dialog, whose OK becomes the rule's format", async () => {
    const { wrapper, onSave } = setup({ kind: "cellIs", operator: "greaterThan", formulas: ["5"] });
    expect(wrapper.find('[data-testid="format-cells-dialog"]').exists()).toBe(false);
    await wrapper.get('[data-testid="cf-format"]').setValue("custom");
    expect(wrapper.find('[data-testid="format-cells-dialog"]').exists()).toBe(true);
    expect(wrapper.findAll('[role="tab"]').map((tab) => tab.attributes("data-testid"))).toEqual(
      ["number", "font", "border", "fill"].map((tab) => `format-cells-tab-${tab}`),
    );
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    expect(wrapper.find('[data-testid="format-cells-font-family"]').exists()).toBe(false);
    await wrapper.get('[data-testid="format-cells-font-style"]').setValue("bold");
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    expect(wrapper.find('[data-testid="format-cells-dialog"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="cf-preview"]').attributes("data-kind")).toBe("format");
    await wrapper.get('[data-testid="cf-save"]').trigger("click");
    expect(onSave.mock.calls[0]![0].dxf).toMatchObject({ fill: { kind: "solid", color: "#FFC7CE" }, font: { color: "#9C0006", bold: true } });
  });

  it("the live preview follows the chosen preset", async () => {
    const { wrapper } = setup({ kind: "cellIs", operator: "greaterThan", formulas: ["5"] });
    const preview = () => wrapper.get('[data-testid="cf-preview"] .cf-preview-format').attributes("style") ?? "";
    expect(preview()).toContain("rgb(255, 199, 206)");
    await wrapper.get('[data-testid="cf-format"]').setValue("2");
    expect(preview()).toContain("rgb(198, 239, 206)");
  });
});
