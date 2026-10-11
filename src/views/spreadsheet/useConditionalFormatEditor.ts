// The conditional-format rule editor's state (SPEC-129 "The rule editor"):
// ConditionalFormatEditor.vue (`cf-editor`) binds to it. `start` opens it on
// a draft — a preset's or an existing rule's — over the ranges it will apply
// to; `save` turns the fields into a CfRuleDraft through cfEditorFields.ts
// (an unparseable value is an inline `error` and nothing is saved) and hands
// it, with the chosen dxf for a formatting rule, to the manager's `onSave`,
// which records the undo step. `cf-format` picks one of Excel's six format
// presets or "Custom format…", which opens SPEC-123's FormatCellsDialog in
// its differential mode on this editor's own dialog instance.
import { computed, ref, shallowRef } from "vue";
import type { DocumentTheme, ModelColor } from "../../lib/core/theme";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { ICON_SET_NAMES, ODS_ONLY_ICON_SETS, type CfRuleDraft, type Dxf, type IconSetName } from "../../lib/spreadsheet/conditionalFormats";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { cellColorHex, dxfKey } from "../../lib/spreadsheet/styles";
import { defaultFields, draftFromFields, fieldsFromDraft, takesFormat, withIconSet, type CfEditorContext, type CfEditorFields, type CfFieldKey } from "./cfEditorFields";
import { CF_FORMAT_PRESETS } from "./cfPresets";
import { rulePreview } from "./cfPreview";
import { useFormatCellsDialog } from "./useFormatCellsDialog";

/** `cf-format`'s value for "Custom format…" (the presets are their index). */
export const CUSTOM_FORMAT = "custom";

export interface CfEditorSave {
  /** The rule being edited; null for a new rule. */
  ruleId: string | null;
  ranges: GridRange[];
  draft: CfRuleDraft;
  /** The formatting rule's dxf; null for a colour scale, data bar or icon set. */
  dxf: Dxf | null;
}

export interface ConditionalFormatEditorDeps {
  theme: () => DocumentTheme;
  locale: () => SpreadsheetLocale;
  dateSystem: () => "1900" | "1904";
  /** The workbook's format: an `.ods` workbook has no reversed icon order. */
  format: () => string;
  /** The manager's save: records the step; false keeps the editor open. */
  onSave: (save: CfEditorSave) => boolean;
}

export interface CfEditorStart {
  ruleId: string | null;
  draft: CfRuleDraft;
  dxf?: Dxf;
  ranges: GridRange[];
  /** The cell the user types relative formulas against. */
  active: CellRef;
}

/** The preset index whose dxf equals `dxf`, else "Custom format…". */
function formatChoiceOf(dxf: Dxf): string {
  const key = dxfKey(dxf);
  const index = CF_FORMAT_PRESETS.findIndex((preset) => dxfKey(preset.dxf) === key);
  return index < 0 ? CUSTOM_FORMAT : String(index);
}

export function useConditionalFormatEditor(deps: ConditionalFormatEditorDeps) {
  const open = ref(false);
  const ruleId = ref<string | null>(null);
  const ranges = shallowRef<GridRange[]>([]);
  const active = shallowRef<CellRef>({ row: 0, col: 0 });
  const fields = shallowRef<CfEditorFields>(defaultFields());
  const dxf = shallowRef<Dxf>(CF_FORMAT_PRESETS[0]!.dxf);
  const formatChoice = ref("0");
  const error = shallowRef<{ field: CfFieldKey; message: string } | null>(null);
  // The differential Format Cells dialog — this editor's own instance, never
  // useFormatting's (whose OK applies a cell style).
  const formatDialog = useFormatCellsDialog({ target: () => null, apply: () => undefined });

  function context(): CfEditorContext {
    const first = ranges.value[0];
    return {
      origin: first ? { row: first.top, col: first.left } : active.value,
      active: active.value,
      locale: deps.locale(),
      dateSystem: deps.dateSystem(),
    };
  }

  function start(options: CfEditorStart): void {
    ruleId.value = options.ruleId;
    ranges.value = options.ranges.map((range) => ({ ...range }));
    active.value = { ...options.active };
    fields.value = fieldsFromDraft(options.draft, context());
    dxf.value = options.dxf ?? CF_FORMAT_PRESETS[0]!.dxf;
    formatChoice.value = options.dxf ? formatChoiceOf(options.dxf) : "0";
    error.value = null;
    open.value = true;
  }

  /** One field changed (the template passes the control's value, so the
   * numeric ones are read back as numbers). */
  function set(key: keyof CfEditorFields, value: unknown): void {
    const next = key === "stdDev" ? Number(value) : value;
    fields.value = { ...fields.value, [key]: next } as CfEditorFields;
    if (error.value) {
      error.value = null;
    }
  }

  function setStop(index: number, key: "type" | "value" | "color", value: string): void {
    set(
      "stops",
      fields.value.stops.map((stop, at) => (at === index ? { ...stop, [key]: value } : stop)),
    );
  }

  /** `gte` arrives as the `≥`/`>` select's "true"/"false". */
  function setIconThreshold(index: number, key: "type" | "value" | "gte", value: string): void {
    const next = key === "gte" ? value === "true" : value;
    set(
      "iconThresholds",
      fields.value.iconThresholds.map((each, at) => (at === index ? { ...each, [key]: next } : each)),
    );
  }

  function setBar(key: "color" | "gradient" | "negativeColor" | "axis", value: string | boolean): void {
    set("bar", { ...fields.value.bar, [key]: value });
  }

  /** A model colour as the colour input's hex. */
  function colorHex(color: ModelColor | undefined): string {
    return cellColorHex(color, deps.theme()) || "#000000";
  }

  function setIconSet(name: string): void {
    fields.value = withIconSet(fields.value, name as IconSetName);
  }

  /** `cf-format` changed: a preset applies its dxf; "Custom format…" opens
   * the differential dialog on the current one. */
  function chooseFormat(value: string): void {
    formatChoice.value = value;
    if (value === CUSTOM_FORMAT) {
      formatDialog.showDifferential(dxf.value, { theme: deps.theme(), locale: deps.locale() }, (next) => {
        dxf.value = next;
      });
      return;
    }
    const preset = CF_FORMAT_PRESETS[Number(value)];
    if (preset) {
      dxf.value = preset.dxf;
    }
  }

  const isOds = computed(() => deps.format() === "ods");
  /** `cf-icon-set`'s choices: LibreOffice's smilies only in an `.ods`. */
  const iconSets = computed(() => ICON_SET_NAMES.filter((name) => isOds.value || !ODS_ONLY_ICON_SETS.has(name)));
  const formatting = computed(() => takesFormat(fields.value));

  /** The live `cf-preview`: the pending rule as the list will show it. */
  const preview = computed(() => {
    const result = draftFromFields(fields.value, context());
    const body: CfRuleDraft = "draft" in result ? result.draft : { kind: "cellIs", operator: "greaterThan", formulas: [] };
    return rulePreview(body, formatting.value ? dxf.value : undefined, deps.theme());
  });

  function close(): void {
    open.value = false;
    formatDialog.cancel();
  }

  /** `cf-save`: an invalid field is an inline error and nothing is saved. */
  function save(): boolean {
    const result = draftFromFields(fields.value, context());
    if ("error" in result) {
      error.value = { field: result.field, message: result.error };
      return false;
    }
    const saved = deps.onSave({ ruleId: ruleId.value, ranges: ranges.value, draft: result.draft, dxf: formatting.value ? dxf.value : null });
    if (saved) {
      close();
    }
    return saved;
  }

  return {
    open,
    ruleId,
    ranges,
    fields,
    dxf,
    formatChoice,
    error,
    formatDialog,
    isOds,
    iconSets,
    formatting,
    preview,
    start,
    set,
    setStop,
    setIconThreshold,
    setBar,
    setIconSet,
    colorHex,
    chooseFormat,
    save,
    cancel: close,
  };
}

export type ConditionalFormatEditor = ReturnType<typeof useConditionalFormatEditor>;

/** Narrows a component's runtime-typed `editor` prop (the .vue parser bars
 * casts). */
export function asConditionalFormatEditor(value: unknown): ConditionalFormatEditor {
  return value as ConditionalFormatEditor;
}
