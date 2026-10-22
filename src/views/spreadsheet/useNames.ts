// SPEC-133's names in the workbook view ("The Name Box", "View wiring"):
// the Name Box, Formulas ▸ Name Manager… and its editor (nameBox.ts: what an
// entry means; names.ts: the commands).
// - Name Box: typing fills `draft`; Enter `submit`s it through
//   interpretNameBox — a jump selects the target's ranges (switching sheet;
//   view state only) and refocuses the grid; a new name is defined for the
//   selection (absoluteRangesFormula) as ONE step "Define name"; a refusal
//   toasts and changes nothing; Escape (`cancel`) restores the address;
//   `options` are the dropdown's. Nothing happens while a cell is edited.
// - Name Manager (`managerOpen`): rows, Delete (ONE step "Delete name"),
//   New/Edit; Ctrl/Cmd+F3 (`sheet.nameManager`) opens it from the grid only.
// - Editor: name, scope (fixed when editing, as in Excel), comment and
//   refers-to, validated live (validateName, then the formula must parse);
//   OK is ONE step "Define name" or "Edit name" (updateName rewrites every
//   formula using a renamed name). Relative parts are shown relative to the
//   active cell and stored relative to A1. Refers-to's point mode (SPEC-122):
//   `startPick` collapses the dialogs while the grid's selection is written
//   in as an absolute, sheet-qualified reference; `endPick` returns.
import { computed, nextTick, onScopeDispose, ref, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { DefinedName, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { absoluteRangesFormula, defineName, deleteName, NAME_PROBLEM_MESSAGES, updateName, validateName, visibleNames } from "../../lib/spreadsheet/names";
import { canPointAt, placePointReference } from "../../lib/spreadsheet/pointMode";
import { inGrid } from "./headerMenu";
import {
  FORMULA_PROBLEM_MESSAGE,
  interpretNameBox,
  nameBoxOptions,
  nameJump,
  nameKey,
  nameManagerRows,
  shownFormula,
  storedFormula,
  tableJump,
  WORKBOOK_SCOPE_LABEL,
  type NameBoxOutcome,
  type NameHost,
} from "./nameBox";
import type { CellEditing } from "./useCellEditing";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo, UndoScope } from "./useSpreadsheetUndo";

/** The names commands' undo labels. */
export const NAME_UNDO_LABELS = { define: "Define name", edit: "Edit name", delete: "Delete name" } as const;

export interface NamesDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
  };
  viewState: Pick<SheetViewState, "update" | "syncInto">;
  selection: Pick<SheetSelectionController, "selection">;
  /** useWorkbookEditing's undo stack and cell entry (busy, focus, reveal). */
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing" | "focusCell" | "revealTick"> };
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
}

/** The editor's fields; `scope` is "" for the workbook, else a sheet id. */
export interface NameEditorFields {
  name: string;
  scope: string;
  comment: string;
  refersTo: string;
}

export interface NameEditorState {
  mode: "new" | "edit";
  /** The edited name's nameKey. */
  key?: string;
  fields: NameEditorFields;
}

export function useNames(deps: NamesDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const busy = (): boolean => deps.editing.cell.editing.value;
  /** The active cell, where names evaluate. */
  function host(): NameHost | null {
    const sheet = deps.book.activeSheet.value;
    const active = deps.selection.selection.value.active;
    return sheet ? { sheetId: sheet.id, row: active.row, col: active.col } : null;
  }
  const definedByKey = (m: SpreadsheetModel, key: string | undefined | null): DefinedName | undefined =>
    key ? visibleNames(m).find((name) => nameKey(name) === key) : undefined;

  /** One undo step: the view state synced, recorded, changed, committed. */
  function mutate(label: string, scope: UndoScope, change: (m: SpreadsheetModel) => void): void {
    const m = model();
    if (!m) {
      return;
    }
    deps.viewState.syncInto(m);
    deps.editing.undo.record(label, scope);
    change(m);
    deps.editing.undo.commit();
  }

  // --- the Name Box ------------------------------------------------------------------------------
  /** What is typed in the box; null shows the address. */
  const draft = ref<string | null>(null);

  /** Selects `ranges` on `sheetId`: the top-left revealed first, then the ranges. */
  async function jump(sheetId: string, ranges: GridRange[]): Promise<void> {
    const index = model()?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
    const first = ranges[0];
    if (index < 0 || !first) {
      return;
    }
    if (index !== deps.book.activeIndex.value) {
      deps.book.selectSheet(index);
    }
    const active = { row: first.top, col: first.left };
    deps.viewState.update(sheetId, { selection: { active, ranges: [{ top: first.top, left: first.left, bottom: first.top, right: first.left }], activeRangeIndex: 0 } });
    await nextTick();
    deps.editing.cell.revealTick.value += 1;
    await nextTick();
    deps.viewState.update(sheetId, { selection: { active, ranges: ranges.map((range) => ({ ...range })), activeRangeIndex: 0 } });
    deps.editing.cell.focusCell();
  }

  /** Defines a workbook-scope name for the active sheet's selection. */
  function defineForSelection(name: string): void {
    const sheet = deps.book.activeSheet.value;
    if (sheet?.kind !== "grid") {
      return;
    }
    const formula = absoluteRangesFormula(sheet.name, deps.selection.selection.value.ranges);
    mutate(NAME_UNDO_LABELS.define, { sheetId: sheet.id }, (m) => defineName(m, { name, formula }));
  }

  async function apply(outcome: NameBoxOutcome): Promise<void> {
    if (outcome.kind === "refuse") {
      toast(outcome.message);
    } else if (outcome.kind === "define") {
      defineForSelection(outcome.name);
      deps.editing.cell.focusCell();
    } else {
      await jump(outcome.sheetId, outcome.ranges);
    }
  }

  /** Enter in the box (see the module header). */
  async function submit(): Promise<void> {
    const text = draft.value;
    draft.value = null;
    const m = model();
    const at = host();
    if (text === null || text.trim() === "" || !m || !at || busy()) {
      deps.editing.cell.focusCell();
      return;
    }
    await apply(interpretNameBox(m, text, at));
  }

  /** Escape (`refocus`: back to the grid) or a blur: the address again. */
  function cancel(refocus = false): void {
    draft.value = null;
    if (refocus) {
      deps.editing.cell.focusCell();
    }
  }

  const options = computed(() => {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    return m && sheet ? nameBoxOptions(m, sheet.id) : [];
  });

  /** A dropdown option chosen: its range selected. */
  async function pick(key: string): Promise<void> {
    draft.value = null;
    const m = model();
    const at = host();
    const option = options.value.find((each) => each.key === key);
    if (!m || !at || !option || busy()) {
      return;
    }
    const name = option.kind === "name" ? definedByKey(m, key) : undefined;
    const outcome = option.kind === "table" ? tableJump(m, option.name) : name ? nameJump(m, name, at) : null;
    if (outcome) {
      await apply(outcome);
    }
  }

  // --- the Name Manager -------------------------------------------------------------------------
  const managerOpen = ref(false);
  const selectedKey = ref<string | null>(null);
  const rows = computed(() => {
    const m = model();
    const at = host();
    return managerOpen.value && m && at ? nameManagerRows(m, at) : [];
  });

  function openManager(): void {
    if (!model() || busy()) {
      return;
    }
    managerOpen.value = true;
    selectedKey.value = rows.value[0]?.key ?? null;
  }

  function closeManager(): void {
    managerOpen.value = false;
    editor.value = null;
    endPick();
    deps.editing.cell.focusCell();
  }

  /** Delete: the selected name, ONE step; the next row is selected. */
  function deleteSelected(): void {
    const m = model();
    const target = m ? definedByKey(m, selectedKey.value) : undefined;
    if (!target) {
      return;
    }
    const at = rows.value.findIndex((row) => row.key === selectedKey.value);
    mutate(NAME_UNDO_LABELS.delete, "workbook", (live) => deleteName(live, target));
    const left = rows.value;
    selectedKey.value = left[Math.min(at, left.length - 1)]?.key ?? null;
  }

  // --- the editor ---------------------------------------------------------------------------------
  const editor = shallowRef<NameEditorState | null>(null);
  const scopes = computed(() => [
    { value: "", label: WORKBOOK_SCOPE_LABEL },
    ...(model()?.sheets ?? []).filter((sheet) => sheet.kind === "grid").map((sheet) => ({ value: sheet.id, label: sheet.name })),
  ]);

  function openNew(): void {
    const sheet = deps.book.activeSheet.value;
    const refersTo = sheet?.kind === "grid" ? `=${absoluteRangesFormula(sheet.name, deps.selection.selection.value.ranges)}` : "=";
    editor.value = { mode: "new", fields: { name: "", scope: "", comment: "", refersTo } };
  }

  function openEdit(): void {
    const m = model();
    const at = host();
    const target = m ? definedByKey(m, selectedKey.value) : undefined;
    if (!target || !at) {
      return;
    }
    const fields = { name: target.name, scope: target.scopeSheetId ?? "", comment: target.comment ?? "", refersTo: shownFormula(target, at) };
    editor.value = { mode: "edit", key: nameKey(target), fields };
  }

  /** A field typed or chosen; an edited name's scope stays fixed. */
  function setField(field: keyof NameEditorFields, value: string): void {
    const open = editor.value;
    if (open && !(field === "scope" && open.mode === "edit")) {
      editor.value = { ...open, fields: { ...open.fields, [field]: value } };
    }
  }

  /** The live validation message ("" when OK may run). */
  const editorError = computed(() => {
    const open = editor.value;
    const m = model();
    const at = host();
    if (!open || !m || !at) {
      return "";
    }
    const problem = validateName(m, open.fields.name.trim(), open.fields.scope || undefined, definedByKey(m, open.key));
    if (problem) {
      return NAME_PROBLEM_MESSAGES[problem];
    }
    return storedFormula(open.fields.refersTo, at) === null ? FORMULA_PROBLEM_MESSAGE : "";
  });

  /** OK: ONE step "Define name" / "Edit name"; false when refused. */
  function confirmEditor(): boolean {
    const open = editor.value;
    const m = model();
    const at = host();
    const formula = open && at ? storedFormula(open.fields.refersTo, at) : null;
    if (!open || !m || formula === null || editorError.value !== "") {
      return false;
    }
    const name = open.fields.name.trim();
    const comment = open.fields.comment.trim();
    const scopeSheetId = open.fields.scope || undefined;
    const fields: DefinedName = { name, formula, ...(scopeSheetId ? { scopeSheetId } : {}), ...(comment ? { comment } : {}) };
    if (open.mode === "new") {
      mutate(NAME_UNDO_LABELS.define, "workbook", (live) => defineName(live, fields));
    } else {
      const target = definedByKey(m, open.key);
      if (!target) {
        return false;
      }
      if (target.name !== name || target.formula !== formula || (target.comment ?? "") !== comment || target.formulaUnparsed) {
        const { formulaUnparsed: _unparsed, comment: _comment, ...kept } = target;
        mutate(NAME_UNDO_LABELS.edit, "workbook", (live) => updateName(live, target, { ...kept, ...fields }));
      }
    }
    selectedKey.value = nameKey(fields);
    editor.value = null;
    endPick();
    return true;
  }

  function cancelEditor(): void {
    editor.value = null;
    endPick();
  }

  // --- Refers to's point mode ------------------------------------------------------------------------
  const picking = ref(false);
  let pointBase: { text: string; span: { start: number; end: number } } | null = null;

  /** "Select cells": `caret` is the field's caret when it was clicked. */
  function startPick(caret?: number): void {
    const open = editor.value;
    if (!open) {
      return;
    }
    const text = open.fields.refersTo;
    const at = Math.min(caret ?? text.length, text.length);
    pointBase = canPointAt(text, at) ? { text, span: { start: at, end: at } } : { text: "=", span: { start: 1, end: 1 } };
    picking.value = true;
  }

  function endPick(): void {
    picking.value = false;
    pointBase = null;
  }

  watch([() => deps.selection.selection.value, () => deps.book.activeSheet.value], ([selection, sheet]) => {
    if (picking.value && pointBase && sheet?.kind === "grid") {
      setField("refersTo", placePointReference(pointBase.text, pointBase.span, absoluteRangesFormula(sheet.name, selection.ranges)).text);
    }
  });

  /** Ctrl/Cmd+F3 from the grid. */
  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || busy()) {
      return;
    }
    if (findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id === "sheet.nameManager") {
      event.preventDefault();
      openManager();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    draft,
    input: (text: string): void => void (draft.value = text),
    submit,
    cancel,
    options,
    pick,
    managerOpen,
    rows,
    selectedKey,
    select: (key: string): void => void (selectedKey.value = key),
    openManager,
    closeManager,
    deleteSelected,
    editor,
    editorError,
    scopes,
    openNew,
    openEdit,
    setField,
    confirmEditor,
    cancelEditor,
    picking,
    startPick,
    endPick,
    onKeydown,
  };
}

export type Names = ReturnType<typeof useNames>;

/** Narrows a component's runtime-typed `names` prop (the .vue parser bars casts). */
export function asNames(value: unknown): Names {
  return value as Names;
}

/** The grid context's `names`, when the view wired them (null: none). */
export function namesOf(context: unknown): Names | null {
  const value = context !== null && typeof context === "object" ? Reflect.get(context, "names") : undefined;
  return value ? (value as Names) : null;
}
