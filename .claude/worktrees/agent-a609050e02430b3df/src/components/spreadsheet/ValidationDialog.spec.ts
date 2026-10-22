// SPEC-130 "The dialog" — `validation-dialog` over useValidationDialog and
// the view's real undo stack: three tabs and every control and label; the
// active cell's rule loaded (a literal list as its items, a relative custom
// formula shown relative to the active cell); the mixed-selection notice;
// OK applying to the whole selection as ONE "Data validation" step with the
// formula stored relative to the selection's origin; Clear All, and "Any
// value" with no input message, removing the rules; each pinned error shown
// in `validation-dialog-error` with nothing saved; Cancel and Escape saving
// nothing; no step for an unchanged OK; nothing while a cell is edited.
import { afterEach, describe, expect, it } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { defaultValidationSettings, type DataValidation } from "../../lib/spreadsheet/validation";
import { useSpreadsheetUndo } from "../../views/spreadsheet/useSpreadsheetUndo";
import { useValidationDialog } from "../../views/spreadsheet/useValidationDialog";
import ValidationDialog from "./ValidationDialog.vue";

const scopes: (() => void)[] = [];
const wrappers: VueWrapper[] = [];
afterEach(() => {
  wrappers.splice(0).forEach((wrapper) => wrapper.unmount());
  scopes.splice(0).forEach((stop) => stop());
});

function range(top: number, left: number, bottom = top, right = left): GridRange {
  return { top, left, bottom, right };
}

/** B2:B10 — a stop list with an input message. */
const PRIORITY: DataValidation = {
  ...defaultValidationSettings(),
  id: "priority",
  ranges: [range(1, 1, 9, 1)],
  type: "list",
  formula1: '"Low,Medium,High"',
  input: { show: true, title: "Priority", text: "Pick one" },
};

/** F2:F10 — a custom rule relative to its origin F2. */
const EMAIL: DataValidation = {
  ...defaultValidationSettings(),
  id: "email",
  ranges: [range(1, 5, 9, 5)],
  type: "custom",
  formula1: 'ISNUMBER(SEARCH("@",F2))',
  error: { show: true, style: "warning", title: "Email", text: "Type an address" },
};

function setup(options: { rules?: DataValidation[]; ranges?: GridRange[]; active?: CellRef } = {}) {
  const model = createBlankWorkbook();
  model.sheets[0]!.validations = options.rules ?? [];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const editingCell = ref(false);
  const active = options.active ?? { row: 1, col: 1 };
  const selection = shallowRef({ ranges: options.ranges ?? [range(active.row, active.col)], active });
  const scope = effectScope();
  scopes.push(() => scope.stop());
  const { undo, dialog } = scope.run(() => {
    const stack = useSpreadsheetUndo({
      model: workbook,
      recalc: (_changed, done) => done([]),
      onApplied: () => triggerRef(workbook),
      activateSheet: () => undefined,
    });
    return {
      undo: stack,
      dialog: useValidationDialog({
        model: () => workbook.value,
        sheet: () => workbook.value?.sheets[0] ?? null,
        selection: () => selection.value,
        undo: stack,
        formatContext: () => ({ locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" }),
        busy: () => editingCell.value,
      }),
    };
  })!;
  dialog.show();
  const wrapper = mount(ValidationDialog, { props: { dialog }, attachTo: document.body });
  wrappers.push(wrapper);
  const rules = () => workbook.value!.sheets[0]!.validations;
  const labelOf = (testid: string) => {
    const id = wrapper.get(`[data-testid="${testid}"]`).attributes("id");
    return wrapper.get(`label[for="${id}"]`).text();
  };
  return { model, workbook, undo, dialog, wrapper, rules, labelOf, editingCell, selection };
}

describe("ValidationDialog (SPEC-130)", () => {
  it("is titled 'Data validation' with three tabs, Criteria first", async () => {
    const { wrapper } = setup();
    const panel = wrapper.get('[role="dialog"]');
    expect(panel.get("h2").text()).toBe("Data validation");
    expect(wrapper.find('[data-testid="validation-dialog"]').exists()).toBe(true);
    const tabs = wrapper.findAll('[role="tab"]');
    expect(tabs.map((tab) => tab.attributes("data-testid"))).toEqual(["validation-tab-criteria", "validation-tab-input", "validation-tab-error"]);
    expect(tabs.map((tab) => tab.text())).toEqual(["Criteria", "Input message", "Error alert"]);
    expect(tabs.map((tab) => tab.attributes("aria-selected"))).toEqual(["true", "false", "false"]);
    expect(tabs.map((tab) => tab.attributes("tabindex"))).toEqual(["0", "-1", "-1"]);
    expect(wrapper.get('[role="tabpanel"]').attributes("aria-labelledby")).toBe("validation-tab-criteria");
    expect(wrapper.get('[data-testid="validation-clear-all"]').text()).toBe("Clear All");
    expect(wrapper.get('[data-testid="validation-cancel"]').text()).toBe("Cancel");
    expect(wrapper.get('[data-testid="validation-ok"]').text()).toBe("OK");

    // Arrow keys move between the tabs.
    await tabs[0]!.trigger("keydown", { key: "ArrowRight" });
    expect(wrapper.get('[data-testid="validation-tab-input"]').attributes("aria-selected")).toBe("true");
    expect(document.activeElement?.getAttribute("data-testid")).toBe("validation-tab-input");
    await wrapper.get('[data-testid="validation-tab-input"]').trigger("keydown", { key: "End" });
    expect(wrapper.get('[data-testid="validation-tab-error"]').attributes("aria-selected")).toBe("true");
    await wrapper.get('[data-testid="validation-tab-error"]').trigger("keydown", { key: "ArrowRight" });
    expect(wrapper.get('[data-testid="validation-tab-criteria"]').attributes("aria-selected")).toBe("true");
  });

  it("has the Criteria controls, each value labelled for the type and operator", async () => {
    const { wrapper, labelOf } = setup();
    const allow = wrapper.get('[data-testid="validation-allow"]');
    expect(labelOf("validation-allow")).toBe("Allow");
    expect(allow.findAll("option").map((option) => option.text())).toEqual([
      "Any value",
      "Whole number",
      "Decimal",
      "List",
      "Date",
      "Time",
      "Text length",
      "Custom",
    ]);
    expect(wrapper.find('[data-testid="validation-operator"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="validation-value-1"]').exists()).toBe(false);
    expect(labelOf("validation-ignore-blank")).toBe("Ignore blank");
    expect(wrapper.find('[data-testid="validation-in-cell-dropdown"]').exists()).toBe(false);

    await allow.setValue("whole");
    const operator = wrapper.get('[data-testid="validation-operator"]');
    expect(operator.findAll("option").map((option) => option.text())).toEqual([
      "between",
      "not between",
      "equal to",
      "not equal to",
      "greater than",
      "less than",
      "greater than or equal to",
      "less than or equal to",
    ]);
    expect([labelOf("validation-value-1"), labelOf("validation-value-2")]).toEqual(["Minimum", "Maximum"]);
    await operator.setValue("equal");
    expect(labelOf("validation-value-1")).toBe("Value");
    expect(wrapper.find('[data-testid="validation-value-2"]').exists()).toBe(false);
    await operator.setValue("greaterThan");
    expect(labelOf("validation-value-1")).toBe("Minimum");
    await operator.setValue("lessThanOrEqual");
    expect(labelOf("validation-value-1")).toBe("Maximum");
    await allow.setValue("textLength");
    expect(labelOf("validation-value-1")).toBe("Maximum");
    await wrapper.get('[data-testid="validation-operator"]').setValue("notEqual");
    expect(labelOf("validation-value-1")).toBe("Length");
    await allow.setValue("date");
    await wrapper.get('[data-testid="validation-operator"]').setValue("between");
    expect([labelOf("validation-value-1"), labelOf("validation-value-2")]).toEqual(["Start date", "End date"]);
    await allow.setValue("time");
    expect([labelOf("validation-value-1"), labelOf("validation-value-2")]).toEqual(["Start time", "End time"]);
    await wrapper.get('[data-testid="validation-operator"]').setValue("equal");
    expect(labelOf("validation-value-1")).toBe("Time");
    await allow.setValue("list");
    expect(wrapper.find('[data-testid="validation-operator"]').exists()).toBe(false);
    expect(labelOf("validation-value-1")).toBe("Source");
    expect(labelOf("validation-in-cell-dropdown")).toBe("In-cell dropdown");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-in-cell-dropdown"]').element.checked).toBe(true);
    await allow.setValue("custom");
    expect(labelOf("validation-value-1")).toBe("Formula");
    expect(wrapper.find('[data-testid="validation-in-cell-dropdown"]').exists()).toBe(false);
  });

  it("has the Input message and Error alert controls with their limits", async () => {
    const { wrapper, labelOf } = setup();
    await wrapper.get('[data-testid="validation-tab-input"]').trigger("click");
    expect(labelOf("validation-show-input")).toBe("Show input message when cell is selected");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-show-input"]').element.checked).toBe(true);
    expect(labelOf("validation-input-title")).toBe("Title");
    expect(wrapper.get('[data-testid="validation-input-title"]').attributes("maxlength")).toBe("32");
    expect(labelOf("validation-input-text")).toBe("Input message");
    expect(wrapper.get('[data-testid="validation-input-text"]').element.tagName).toBe("TEXTAREA");
    expect(wrapper.get('[data-testid="validation-input-text"]').attributes("maxlength")).toBe("255");

    await wrapper.get('[data-testid="validation-tab-error"]').trigger("click");
    expect(wrapper.find('[data-testid="validation-input-title"]').exists()).toBe(false);
    expect(labelOf("validation-show-error")).toBe("Show error alert after invalid data is entered");
    expect(labelOf("validation-error-style")).toBe("Style");
    expect(wrapper.get('[data-testid="validation-error-style"]').findAll("option").map((option) => option.text())).toEqual(["Stop", "Warning", "Information"]);
    expect(labelOf("validation-error-title")).toBe("Title");
    expect(wrapper.get('[data-testid="validation-error-title"]').attributes("maxlength")).toBe("32");
    expect(labelOf("validation-error-text")).toBe("Error message");
    expect(wrapper.get('[data-testid="validation-error-text"]').attributes("maxlength")).toBe("225");
  });

  it("loads the active cell's literal list and its messages", async () => {
    const { wrapper } = setup({ rules: [PRIORITY], active: { row: 4, col: 1 } });
    expect(wrapper.get<HTMLSelectElement>('[data-testid="validation-allow"]').element.value).toBe("list");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-value-1"]').element.value).toBe("Low,Medium,High");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-in-cell-dropdown"]').element.checked).toBe(true);
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-ignore-blank"]').element.checked).toBe(true);
    await wrapper.get('[data-testid="validation-tab-input"]').trigger("click");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-input-title"]').element.value).toBe("Priority");
    expect(wrapper.get<HTMLTextAreaElement>('[data-testid="validation-input-text"]').element.value).toBe("Pick one");
    expect(wrapper.find('[data-testid="validation-mixed-notice"]').exists()).toBe(false);
  });

  it("shows a relative custom formula relative to the active cell", async () => {
    const { wrapper } = setup({ rules: [EMAIL], ranges: [range(1, 5, 9, 5)], active: { row: 6, col: 5 } });
    expect(wrapper.get<HTMLSelectElement>('[data-testid="validation-allow"]').element.value).toBe("custom");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-value-1"]').element.value).toBe('=ISNUMBER(SEARCH("@",F7))');
    await wrapper.get('[data-testid="validation-tab-error"]').trigger("click");
    expect(wrapper.get<HTMLSelectElement>('[data-testid="validation-error-style"]').element.value).toBe("warning");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-error-title"]').element.value).toBe("Email");
    expect(wrapper.get<HTMLTextAreaElement>('[data-testid="validation-error-text"]').element.value).toBe("Type an address");
  });

  it("says when the selection's cells disagree", () => {
    const partly = setup({ rules: [PRIORITY], ranges: [range(8, 1, 11, 1)], active: { row: 8, col: 1 } });
    expect(partly.wrapper.get('[data-testid="validation-mixed-notice"]').text()).toBe(
      "The selection contains cells with different validation settings — OK applies these settings to all of them",
    );
    const two = setup({ rules: [PRIORITY, EMAIL], ranges: [range(1, 1, 1, 5)], active: { row: 1, col: 1 } });
    expect(two.wrapper.find('[data-testid="validation-mixed-notice"]').exists()).toBe(true);
    const inside = setup({ rules: [PRIORITY], ranges: [range(2, 1, 3, 1)], active: { row: 2, col: 1 } });
    expect(inside.wrapper.find('[data-testid="validation-mixed-notice"]').exists()).toBe(false);
  });

  it("OK applies a List with Source 'Yes,No' to the whole selection as one 'Data validation' step", async () => {
    const { wrapper, dialog, rules, undo, model } = setup({ ranges: [range(1, 3, 3, 3)], active: { row: 1, col: 3 } });
    await wrapper.get('[data-testid="validation-allow"]').setValue("list");
    await wrapper.get('[data-testid="validation-value-1"]').setValue("Yes,No");
    await wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(dialog.open.value).toBe(false);
    expect(rules()).toHaveLength(1);
    expect(rules()[0]).toMatchObject({ ranges: [range(1, 3, 3, 3)], type: "list", formula1: '"Yes,No"', showDropdown: true, allowBlank: true });
    expect(model.sheets[0]!.dirty).toBe(true);
    expect(undo.undoLabel.value).toBe("Data validation");
    undo.undo();
    expect(rules()).toEqual([]);
    expect(undo.canUndo.value).toBe(false);
  });

  it("stores a formula typed at the active cell relative to the selection's top-left", async () => {
    const { wrapper, rules, undo } = setup({ ranges: [range(1, 3, 9, 3)], active: { row: 4, col: 3 } });
    await wrapper.get('[data-testid="validation-allow"]').setValue("custom");
    await wrapper.get('[data-testid="validation-value-1"]').setValue("=D5>$Z$1");
    await wrapper.get('[data-testid="validation-value-1"]').trigger("keydown", { key: "Enter" });
    expect(rules()).toEqual([expect.objectContaining({ ranges: [range(1, 3, 9, 3)], type: "custom", formula1: "D2>$Z$1" })]);
    expect(undo.undoLabel.value).toBe("Data validation");
  });

  it("replaces the rules of the cells it covers, merging into an identical rule", async () => {
    const { wrapper, rules } = setup({ rules: [PRIORITY], ranges: [range(4, 1, 5, 1)], active: { row: 4, col: 1 } });
    await wrapper.get('[data-testid="validation-tab-error"]').trigger("click");
    await wrapper.get('[data-testid="validation-error-style"]').setValue("information");
    await wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(rules().map((rule) => [rule.ranges, rule.error.style])).toEqual([
      [[range(1, 1, 3, 1), range(6, 1, 9, 1)], "stop"],
      [[range(4, 1, 5, 1)], "information"],
    ]);
  });

  it("Clear All resets the fields, so OK removes the rules", async () => {
    const { wrapper, dialog, rules, undo } = setup({ rules: [PRIORITY, EMAIL], ranges: [range(1, 1, 9, 1)], active: { row: 1, col: 1 } });
    await wrapper.get('[data-testid="validation-clear-all"]').trigger("click");
    expect(wrapper.get<HTMLSelectElement>('[data-testid="validation-allow"]').element.value).toBe("any");
    expect(wrapper.find('[data-testid="validation-value-1"]').exists()).toBe(false);
    await wrapper.get('[data-testid="validation-tab-input"]').trigger("click");
    expect(wrapper.get<HTMLInputElement>('[data-testid="validation-input-title"]').element.value).toBe("");
    expect(dialog.open.value).toBe(true);
    await wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(rules()).toEqual([EMAIL]);
    expect(undo.undoLabel.value).toBe("Data validation");
  });

  it("OK with Any value and no input message removes the rules; an input message keeps one", async () => {
    const removed = setup({ rules: [PRIORITY], ranges: [range(1, 1, 9, 1)] });
    await removed.wrapper.get('[data-testid="validation-allow"]').setValue("any");
    await removed.wrapper.get('[data-testid="validation-tab-input"]').trigger("click");
    await removed.wrapper.get('[data-testid="validation-input-title"]').setValue("");
    await removed.wrapper.get('[data-testid="validation-input-text"]').setValue("");
    await removed.wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(removed.rules()).toEqual([]);

    const kept = setup({ rules: [PRIORITY], ranges: [range(1, 1, 9, 1)] });
    await kept.wrapper.get('[data-testid="validation-allow"]').setValue("any");
    await kept.wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(kept.rules()).toEqual([expect.objectContaining({ type: "any", input: { show: true, title: "Priority", text: "Pick one" } })]);
    expect(kept.rules()[0]!.formula1).toBeUndefined();
  });

  it.each([
    ["custom", "=1+", "The formula contains an error"],
    ["whole", "10|1", "The Maximum must be greater than or equal to the Minimum"],
    ["textLength", "-1|5", "Enter a whole number of 0 or more"],
    ["list", "x".repeat(256), "The list source can't be longer than 255 characters"],
    ["list", "=A1:B2", "The list source must be a delimited list, or a reference to a single row or column"],
  ])("shows '%s' %s's error, verbatim, and saves nothing", async (type, values, message) => {
    const { wrapper, dialog, rules, undo } = setup({ rules: [PRIORITY] });
    await wrapper.get('[data-testid="validation-allow"]').setValue(type);
    const [first, second] = values.split("|");
    await wrapper.get('[data-testid="validation-value-1"]').setValue(first!);
    if (second !== undefined) {
      await wrapper.get('[data-testid="validation-value-2"]').setValue(second);
    }
    await wrapper.get('[data-testid="validation-tab-error"]').trigger("click");
    await wrapper.get('[data-testid="validation-ok"]').trigger("click");
    const error = wrapper.get('[data-testid="validation-dialog-error"]');
    expect(error.text()).toBe(message);
    expect(error.attributes("role")).toBe("alert");
    // Back on Criteria, the value in error marked.
    expect(wrapper.get('[data-testid="validation-tab-criteria"]').attributes("aria-selected")).toBe("true");
    expect(wrapper.find('[aria-invalid="true"]').exists()).toBe(true);
    expect(dialog.open.value).toBe(true);
    expect(rules()).toEqual([PRIORITY]);
    expect(undo.canUndo.value).toBe(false);
    // Editing a value clears the message.
    await wrapper.get('[data-testid="validation-value-1"]').setValue("1");
    expect(wrapper.find('[data-testid="validation-dialog-error"]').exists()).toBe(false);
  });

  it("Cancel and Escape close with nothing saved", async () => {
    const cancelled = setup({ rules: [PRIORITY] });
    await cancelled.wrapper.get('[data-testid="validation-allow"]').setValue("whole");
    await cancelled.wrapper.get('[data-testid="validation-cancel"]').trigger("click");
    expect(cancelled.dialog.open.value).toBe(false);
    expect(cancelled.rules()).toEqual([PRIORITY]);
    expect(cancelled.undo.canUndo.value).toBe(false);

    const escaped = setup();
    await escaped.wrapper.get('[data-testid="validation-allow"]').setValue("list");
    await escaped.wrapper.get('[data-testid="validation-value-1"]').setValue("a,b");
    await escaped.wrapper.get('[role="dialog"]').trigger("keydown", { key: "Escape" });
    expect(escaped.dialog.open.value).toBe(false);
    expect(escaped.rules()).toEqual([]);
  });

  it("records nothing for an OK that changes nothing", async () => {
    const unchanged = setup({ rules: [PRIORITY], ranges: [range(3, 1)], active: { row: 3, col: 1 } });
    await unchanged.wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(unchanged.dialog.open.value).toBe(false);
    expect(unchanged.rules()[0]).toBe(PRIORITY);
    expect(unchanged.undo.canUndo.value).toBe(false);

    const empty = setup();
    const sheets = empty.model.sheets;
    await empty.wrapper.get('[data-testid="validation-ok"]').trigger("click");
    expect(empty.dialog.open.value).toBe(false);
    expect(empty.undo.canUndo.value).toBe(false);
    expect(empty.model.sheets).toBe(sheets);
  });

  it("neither opens nor applies while a cell is being edited", async () => {
    const { wrapper, dialog, rules, editingCell } = setup();
    await wrapper.get('[data-testid="validation-allow"]').setValue("list");
    await wrapper.get('[data-testid="validation-value-1"]').setValue("a,b");
    editingCell.value = true;
    await nextTick();
    expect(wrapper.get('[data-testid="validation-ok"]').attributes("disabled")).toBeDefined();
    expect(dialog.ok()).toBe(false);
    expect(rules()).toEqual([]);
    dialog.cancel();
    dialog.show();
    expect(dialog.open.value).toBe(false);
    editingCell.value = false;
    dialog.show();
    expect(dialog.open.value).toBe(true);
  });
});
