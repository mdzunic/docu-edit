// The conditional-formatting rules manager's state (SPEC-129 "The side
// panel"): ConditionalFormatsPane.vue (`cf-manager`) binds to it. It lists
// the active sheet's rules in priority order — every rule ("This sheet") or
// those whose ranges intersect the selection ("Current selection") — keeps
// the selected rule, and runs every user action as ONE undo step labelled
// "Conditional formatting" through the view's undo stack (record → command →
// commit; a step that changes nothing records nothing, the stack's own
// rule), which is also what invalidates the grid's evaluators
// (useConditionalFormats). A preset applies to the selection current when it
// is used; a value preset opens the rule editor (useConditionalFormatEditor)
// instead, whose save lands here. Dxfs are interned inside the step, so an
// undo past a new one restores the shorter table.
import { computed, ref, shallowRef, watch } from "vue";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import {
  addConditionalFormat,
  clearRulesInRanges,
  CONDITIONAL_FORMATTING_LABEL,
  deleteRule,
  isFormattingRule,
  moveRule,
  rulesByPriority,
  setFormatRanges,
  setStopIfTrue,
  sheetFormats,
  updateRule,
  type CfRule,
  type CfRuleDraft,
  type ConditionalFormat,
} from "../../lib/spreadsheet/conditionalFormats";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { internDxf } from "../../lib/spreadsheet/styles";
import { describeRule } from "./cfDescriptions";
import { addPreset, CF_FORMAT_PRESETS, EDITOR_PRESETS } from "./cfPresets";
import { rulePreview, type CfPreview } from "./cfPreview";
import { formatAppliesTo, parseAppliesTo } from "./cfRangesInput";
import { useConditionalFormatEditor, type CfEditorSave } from "./useConditionalFormatEditor";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export type CfScope = "selection" | "sheet";

/** `cf-scope`'s options. */
export const CF_SCOPES: readonly { id: CfScope; label: string }[] = [
  { id: "selection", label: "Current selection" },
  { id: "sheet", label: "This sheet" },
];

export interface ConditionalFormatManagerDeps {
  model: () => SpreadsheetModel | null;
  /** The sheet on show (only a grid sheet has rules to manage). */
  sheet: () => Sheet | null;
  /** The grid selection as it is now: its ranges and active cell. */
  selection: () => { ranges: GridRange[]; active: CellRef } | null;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  locale: () => SpreadsheetLocale;
  /** The view state written into the model before a step (useSortFilter's
   * order). */
  syncInto?: (model: SpreadsheetModel) => void;
  /** True while a cell is being edited: steps are refused. */
  busy?: () => boolean;
}

/** One `cf-rule` option. */
export interface CfRuleEntry {
  rule: CfRule;
  format: ConditionalFormat;
  description: string;
  preview: CfPreview;
  appliesTo: string;
  /** A formatting kind: offers `cf-rule-stop-if-true`. */
  formatting: boolean;
}

function intersects(a: GridRange, b: GridRange): boolean {
  return a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
}

/** A rule as a draft (what the editor opens on and the commands take). */
function draftOf(rule: CfRule): CfRuleDraft {
  const { id: _id, priority: _priority, ...draft } = rule;
  return draft;
}

export function useConditionalFormatManager(deps: ConditionalFormatManagerDeps) {
  const scope = ref<CfScope>("sheet");
  const selectedRuleId = ref<string | null>(null);
  const presetsOpen = ref(false);
  /** An invalid "Applies to" entry's message, by format id. */
  const appliesToErrors = shallowRef<Record<string, string>>({});

  const gridSheet = (): Sheet | null => {
    const sheet = deps.sheet();
    return sheet && sheet.kind === "grid" ? sheet : null;
  };
  const theme = () => deps.model()?.theme ?? DEFAULT_THEME;

  const editor = useConditionalFormatEditor({
    theme,
    locale: deps.locale,
    dateSystem: () => deps.model()?.dateSystem ?? "1900",
    format: () => deps.model()?.format ?? "xlsx",
    onSave: saveFromEditor,
  });

  // Another sheet on show: the editor and the presets close and nothing
  // stays selected (a rule id belongs to its sheet).
  watch(
    () => gridSheet()?.id,
    () => {
      editor.cancel();
      presetsOpen.value = false;
      selectedRuleId.value = null;
      appliesToErrors.value = {};
    },
  );

  /** The sheet's rules in priority order, within the scope. */
  const entries = computed<CfRuleEntry[]>(() => {
    const model = deps.model();
    const sheet = gridSheet();
    if (!model || !sheet) {
      return [];
    }
    const selection = scope.value === "selection" ? (deps.selection()?.ranges ?? []) : null;
    return rulesByPriority(sheetFormats(sheet))
      .filter(({ format }) => selection === null || format.ranges.some((range) => selection.some((each) => intersects(range, each))))
      .map(({ format, rule }) => ({
        rule,
        format,
        description: describeRule(rule),
        preview: rulePreview(rule, rule.dxfId === undefined ? undefined : model.dxfs[rule.dxfId], theme()),
        appliesTo: formatAppliesTo(format.ranges),
        formatting: isFormattingRule(rule),
      }));
  });

  const selected = computed(() => entries.value.find((entry) => entry.rule.id === selectedRuleId.value) ?? null);
  const selectedIndex = computed(() => entries.value.findIndex((entry) => entry.rule.id === selectedRuleId.value));
  const canMoveUp = computed(() => selectedIndex.value > 0);
  const canMoveDown = computed(() => selectedIndex.value >= 0 && selectedIndex.value < entries.value.length - 1);

  /** One user action as ONE "Conditional formatting" undo step. */
  function step<T>(mutate: (model: SpreadsheetModel, sheetId: string) => T): T | null {
    const model = deps.model();
    const sheet = gridSheet();
    if (!model || !sheet || deps.busy?.()) {
      return null;
    }
    deps.syncInto?.(model);
    deps.undo.record(CONDITIONAL_FORMATTING_LABEL, { sheetId: sheet.id });
    const result = mutate(model, sheet.id);
    deps.undo.commit();
    return result;
  }

  function select(ruleId: string): void {
    selectedRuleId.value = ruleId;
  }

  /** The listbox's ArrowUp/ArrowDown: the neighbouring rule. */
  function selectBy(delta: number): void {
    const list = entries.value;
    const next = list[Math.min(list.length - 1, Math.max(0, selectedIndex.value + delta))];
    if (next) {
      selectedRuleId.value = next.rule.id;
    }
  }

  function setScope(value: string): void {
    scope.value = value === "selection" ? "selection" : "sheet";
  }

  /** Adds a rule over `ranges` at priority 1 and selects it. */
  function addRule(ranges: GridRange[], draft: CfRuleDraft, dxf: CfEditorSave["dxf"]): void {
    const added = step((model, sheetId) => {
      const full: CfRuleDraft = dxf ? { ...draft, dxfId: internDxf(model, dxf) } : draft;
      return addConditionalFormat(model, sheetId, ranges, full);
    });
    if (added) {
      selectedRuleId.value = added.ruleId;
    }
  }

  /** A `cf-preset-*` item: a value preset opens the editor on the current
   * selection; the others apply to it directly. */
  function applyPreset(testid: string): void {
    presetsOpen.value = false;
    const preset = addPreset(testid);
    const selection = deps.selection();
    if (!preset || !selection || selection.ranges.length === 0 || !gridSheet()) {
      return;
    }
    const draft = preset.draft();
    if (EDITOR_PRESETS.has(testid)) {
      editor.start({ ruleId: null, draft, ranges: selection.ranges, active: selection.active });
      return;
    }
    addRule(selection.ranges, draft, isFormattingRule(draft) ? CF_FORMAT_PRESETS[0]!.dxf : null);
  }

  /** `cf-edit`: the selected rule in the editor (a rule kept from the file
   * is not editable). */
  function editSelected(): void {
    const entry = selected.value;
    const model = deps.model();
    if (!entry || !model || entry.rule.kind === "preserved") {
      return;
    }
    const dxf = entry.rule.dxfId === undefined ? undefined : model.dxfs[entry.rule.dxfId];
    const first = entry.format.ranges[0];
    const active = deps.selection()?.active ?? (first ? { row: first.top, col: first.left } : { row: 0, col: 0 });
    editor.start({ ruleId: entry.rule.id, draft: draftOf(entry.rule), ranges: entry.format.ranges, active, ...(dxf ? { dxf } : {}) });
  }

  /** The editor's save: one step that interns the dxf and adds or replaces
   * the rule; the saved rule is then the selected one. */
  function saveFromEditor(save: CfEditorSave): boolean {
    if (save.ruleId === null) {
      addRule(save.ranges, save.draft, save.dxf);
      return true;
    }
    const ruleId = save.ruleId;
    const done = step((model, sheetId) => {
      const sheet = model.sheets.find((each) => each.id === sheetId);
      const old = sheet ? rulesByPriority(sheetFormats(sheet)).find((entry) => entry.rule.id === ruleId)?.rule : undefined;
      if (!old) {
        return false;
      }
      const draft: CfRuleDraft = { ...save.draft, ...(old.stopIfTrue ? { stopIfTrue: true as const } : {}) };
      if (save.dxf) {
        draft.dxfId = internDxf(model, save.dxf);
      }
      updateRule(model, sheetId, ruleId, draft);
      return true;
    });
    if (done) {
      selectedRuleId.value = ruleId;
    }
    return done === true;
  }

  /** `cf-delete`: the selected rule; its neighbour is selected next. */
  function deleteSelected(): void {
    const index = selectedIndex.value;
    const entry = selected.value;
    if (!entry) {
      return;
    }
    const neighbour = entries.value[index + 1] ?? entries.value[index - 1] ?? null;
    if (step((model, sheetId) => deleteRule(model, sheetId, entry.rule.id)) !== null) {
      selectedRuleId.value = neighbour ? neighbour.rule.id : null;
    }
  }

  function move(direction: "up" | "down"): void {
    const entry = selected.value;
    if (entry && (direction === "up" ? canMoveUp.value : canMoveDown.value)) {
      step((model, sheetId) => moveRule(model, sheetId, entry.rule.id, direction));
    }
  }

  /** `cf-rule-applies-to` committed: invalid text is an inline error and
   * nothing is saved. */
  function setAppliesTo(formatId: string, text: string): void {
    const parsed = parseAppliesTo(text);
    const { [formatId]: _old, ...rest } = appliesToErrors.value;
    if ("error" in parsed) {
      appliesToErrors.value = { ...rest, [formatId]: parsed.error };
      return;
    }
    appliesToErrors.value = rest;
    const format = entries.value.find((entry) => entry.format.id === formatId)?.format;
    if (!format || formatAppliesTo(format.ranges) === formatAppliesTo(parsed.ranges)) {
      return;
    }
    step((model, sheetId) => setFormatRanges(model, sheetId, formatId, parsed.ranges));
  }

  function setStop(ruleId: string, stop: boolean): void {
    step((model, sheetId) => setStopIfTrue(model, sheetId, ruleId, stop));
  }

  /** `cf-clear-selection`: rules removed from the selected cells only. */
  function clearSelection(): void {
    const ranges = deps.selection()?.ranges ?? [];
    if (ranges.length > 0) {
      step((model, sheetId) => clearRulesInRanges(model, sheetId, ranges));
    }
  }

  /** `cf-clear-sheet`: every rule on the sheet. */
  function clearSheet(): void {
    const sheet = gridSheet();
    const ranges = sheet ? sheetFormats(sheet).flatMap((format) => format.ranges) : [];
    if (ranges.length > 0) {
      step((model, sheetId) => clearRulesInRanges(model, sheetId, ranges));
      selectedRuleId.value = null;
    }
  }

  return {
    scope,
    entries,
    selectedRuleId,
    selected,
    canMoveUp,
    canMoveDown,
    presetsOpen,
    appliesToErrors,
    editor,
    hasSheet: computed(() => gridSheet() !== null),
    select,
    selectBy,
    setScope,
    applyPreset,
    editSelected,
    deleteSelected,
    moveUp: (): void => move("up"),
    moveDown: (): void => move("down"),
    setAppliesTo,
    setStop,
    clearSelection,
    clearSheet,
  };
}

export type ConditionalFormatManager = ReturnType<typeof useConditionalFormatManager>;

/** Narrows a component's runtime-typed `manager` prop (the .vue parser bars
 * casts). */
export function asConditionalFormatManager(value: unknown): ConditionalFormatManager {
  return value as ConditionalFormatManager;
}
