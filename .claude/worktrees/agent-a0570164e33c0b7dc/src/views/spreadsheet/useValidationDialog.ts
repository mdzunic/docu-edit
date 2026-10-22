// The Data validation dialog's state (SPEC-130 "The dialog"):
// ValidationDialog.vue (`validation-dialog`, opened by Data ▸ "Data
// validation…") binds to it. `show` opens it on the selection current then:
// the ACTIVE cell's rule (Excel's defaults when it has none) in the fields
// validationFields.ts maps — formulas shown relative to the active cell —
// and `mixed` (`validation-mixed-notice`) when the selection's cells disagree.
// `ok` turns the fields into settings (a value in error is the dialog's
// inline `error` — one of its five verbatim messages — and nothing is saved)
// and applies them to the WHOLE selection as ONE undo step labelled
// DATA_VALIDATION_LABEL through the view's undo stack (record → command →
// commit): setValidation, or clearValidation when the result says nothing
// ("Any value" with no input message — what "Clear All" leaves). OK on
// unchanged fields over a selection that agrees records nothing, nor does
// a clear that removes nothing (the stack pops an unchanged step). Cancel
// and Escape close with nothing saved. While a cell is being edited (`busy`)
// the dialog neither opens nor applies, and another sheet on show closes it.
import { computed, ref, shallowRef, watch } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import {
  clearValidation,
  DATA_VALIDATION_LABEL,
  isEmptyValidation,
  setValidation,
  usesOperator,
  validationAt,
  VALIDATION_OPERATORS,
  VALIDATION_TYPES,
} from "../../lib/spreadsheet/validation";
import {
  defaultValidationFields,
  disjointRanges,
  fieldsFromRule,
  sameValidationFields,
  selectionDisagrees,
  settingsFromFields,
  valueLabels,
  VALIDATION_ERROR_STYLE_OPTIONS,
  type ValidationFieldKey,
  type ValidationFields,
} from "./validationFields";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export type ValidationDialogTab = "criteria" | "input" | "error";

/** The tabs, in order (`validation-tab-<id>`). */
export const VALIDATION_DIALOG_TABS: readonly { id: ValidationDialogTab; label: string }[] = [
  { id: "criteria", label: "Criteria" },
  { id: "input", label: "Input message" },
  { id: "error", label: "Error alert" },
];

export interface ValidationDialogDeps {
  model: () => SpreadsheetModel | null;
  /** The sheet on show (only a grid sheet has rules). */
  sheet: () => Sheet | null;
  /** The grid selection as it is now: its ranges and active cell. */
  selection: () => { ranges: GridRange[]; active: CellRef } | null;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  /** The workbook's format context in the user's locale (values are read
   * and shown in it). */
  formatContext: () => FormatContext;
  /** The view state written into the model before a step (useSortFilter's
   * order). */
  syncInto?: (model: SpreadsheetModel) => void;
  /** True while a cell is being edited. */
  busy?: () => boolean;
}

type BooleanField = "allowBlank" | "showDropdown" | "inputShow" | "errorShow";
type TextField = "value1" | "value2" | "inputTitle" | "inputText" | "errorTitle" | "errorText";

const BOOLEAN_FIELDS: ReadonlySet<string> = new Set<BooleanField>(["allowBlank", "showDropdown", "inputShow", "errorShow"]);
const TEXT_FIELDS: ReadonlySet<string> = new Set<TextField>(["value1", "value2", "inputTitle", "inputText", "errorTitle", "errorText"]);

export function useValidationDialog(deps: ValidationDialogDeps) {
  const open = ref(false);
  const tab = ref<ValidationDialogTab>("criteria");
  const fields = shallowRef<ValidationFields>(defaultValidationFields());
  const error = shallowRef<{ field: ValidationFieldKey; message: string } | null>(null);
  const mixed = ref(false);
  /** What `show` opened on: the sheet, the selection's ranges (disjoint) and
   * active cell, and the fields as loaded. */
  let target: { sheetId: string; ranges: GridRange[]; active: CellRef; loaded: ValidationFields } | null = null;

  const gridSheet = (): Sheet | null => {
    const sheet = deps.sheet();
    return sheet && sheet.kind === "grid" ? sheet : null;
  };
  const busy = computed(() => deps.busy?.() === true);
  /** The value fields' labels (null: hidden) for the type and operator. */
  const labels = computed(() => valueLabels(fields.value.type, fields.value.operator));
  const hasOperator = computed(() => usesOperator(fields.value.type));

  function close(): void {
    open.value = false;
    target = null;
    error.value = null;
  }

  // Another sheet on show: the dialog's target is gone.
  watch(
    () => gridSheet()?.id,
    () => {
      if (open.value) {
        close();
      }
    },
  );

  /** Data ▸ "Data validation…": opens on the current selection. */
  function show(): void {
    const sheet = gridSheet();
    const selection = deps.selection();
    if (!deps.model() || !sheet || !selection || selection.ranges.length === 0 || busy.value) {
      return;
    }
    const ranges = disjointRanges(selection.ranges);
    const active = { ...selection.active };
    const loaded = fieldsFromRule(validationAt(sheet, active), active, deps.formatContext());
    target = { sheetId: sheet.id, ranges, active, loaded };
    fields.value = loaded;
    mixed.value = selectionDisagrees(sheet, ranges);
    error.value = null;
    tab.value = "criteria";
    open.value = true;
  }

  /** One control changed: the template passes its value (a checkbox's
   * checked state, a select's or text field's value). A new Allow type
   * starts its values empty. */
  function set(key: string, value: string | boolean): void {
    const current = fields.value;
    const type = key === "type" ? VALIDATION_TYPES.find((each) => each === value) : undefined;
    const operator = key === "operator" ? VALIDATION_OPERATORS.find((each) => each === value) : undefined;
    const style = key === "errorStyle" ? VALIDATION_ERROR_STYLE_OPTIONS.find((each) => each.id === value)?.id : undefined;
    let next: ValidationFields;
    if (type !== undefined) {
      next = { ...current, type, value1: "", value2: "" };
      if (type !== "list") {
        delete next.odsListOrder;
      }
    } else if (operator !== undefined) {
      next = { ...current, operator };
    } else if (style !== undefined) {
      next = { ...current, errorStyle: style };
    } else if (BOOLEAN_FIELDS.has(key) && typeof value === "boolean") {
      next = { ...current, [key as BooleanField]: value };
    } else if (TEXT_FIELDS.has(key) && typeof value === "string") {
      next = { ...current, [key as TextField]: value };
    } else {
      return;
    }
    fields.value = next;
    if (error.value) {
      error.value = null;
    }
  }

  function showTab(next: string): void {
    const found = VALIDATION_DIALOG_TABS.find((each) => each.id === next);
    if (found) {
      tab.value = found.id;
    }
  }

  /** `validation-clear-all`: every field back to Excel's defaults, so OK
   * removes the rules. */
  function clearAll(): void {
    fields.value = defaultValidationFields();
    error.value = null;
  }

  /** `validation-ok` (see the module header). Returns whether the dialog
   * closed. */
  function ok(): boolean {
    const model = deps.model();
    const sheet = gridSheet();
    const at = target;
    if (!open.value || !model || !sheet || !at || sheet.id !== at.sheetId || busy.value) {
      return false;
    }
    if (!mixed.value && sameValidationFields(fields.value, at.loaded)) {
      close();
      return true;
    }
    const first = at.ranges[0]!;
    const result = settingsFromFields(fields.value, { active: at.active, origin: { row: first.top, col: first.left }, format: deps.formatContext() });
    if ("error" in result) {
      error.value = { field: result.field, message: result.error };
      tab.value = "criteria";
      return false;
    }
    deps.syncInto?.(model);
    deps.undo.record(DATA_VALIDATION_LABEL, { sheetId: sheet.id });
    if (isEmptyValidation(result.settings)) {
      clearValidation(model, sheet.id, at.ranges);
    } else {
      setValidation(model, sheet.id, at.ranges, result.settings);
    }
    deps.undo.commit();
    close();
    return true;
  }

  return {
    open,
    tab,
    fields,
    error,
    mixed,
    busy,
    labels,
    hasOperator,
    show,
    set,
    showTab,
    clearAll,
    ok,
    cancel: close,
  };
}

export type ValidationDialog = ReturnType<typeof useValidationDialog>;

/** Narrows a component's runtime-typed `dialog` prop (the .vue parser bars
 * casts). */
export function asValidationDialog(value: unknown): ValidationDialog {
  return value as ValidationDialog;
}
