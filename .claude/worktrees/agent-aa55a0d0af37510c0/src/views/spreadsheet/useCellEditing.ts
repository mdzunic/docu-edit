// Cell entry (SPEC-122 "Editor keys", "Editor DOM", "Protection"): THIS
// composable owns the editing state machine — Ready, Enter, Edit (and Point,
// usePointMode's, shown while a reference is being pointed at) — and the
// draft (`text`, `caret`) the in-cell editor (CellEditor.vue) and the
// editable formula bar share; both render it and report input through
// useDraftSurface, so typing in either updates both.
//
// - Ready: the first printable character or IME composition lands natively
//   in the (empty) editor and opens Enter mode with that input replacing the
//   cell's content; F2 (or a double-click, or a click in the formula bar)
//   opens Edit mode on the cell's edit text, caret at the end (or at the
//   clicked character); Delete/Backspace clear the selection's contents;
//   Enter/Tab move (SPEC-117's cycle within a multi-cell selection); every
//   other key is SPEC-117's navigation, on the editor proxy unchanged.
// - Enter/Edit: Enter/Shift+Enter commit and move down/up, Tab/Shift+Tab
//   right/left; Enter after a run of Tabs returns to the run's starting
//   column one row down (any other navigation, a click included, forgets
//   the run); inside a multi-cell selection Enter and Tab cycle within it
//   without collapsing it. Enter-mode arrows point when usePointMode allows,
//   else commit and move; Edit-mode arrows move the caret. Escape cancels;
//   Alt/Option+Enter inserts a line break; Ctrl/Cmd+Enter fills the
//   selection ("Fill"); F2 toggles Enter ↔ Edit; F4 cycles `$`. Cmd/Ctrl+Z/Y
//   undo the draft's own text edits — never the workbook (the editor's
//   handler takes them first, so useSpreadsheetUndo's listener never fires
//   mid-edit). Enter, Tab, arrows and Escape are ignored while an IME
//   composition is active.
// - Commit (cellCommit.ts over cellEntry.ts): parse, validate, one undo
//   step, refit, move; a formula error keeps the editor open with the message
//   and the caret at the error; an unchanged text records nothing.
//   Protection and arrays are checked before the editor opens and before
//   Delete or Fill (a toast). SPEC-130: then the data-validation entry gate
//   (`entryGate`) — for a commit, a Fill and a dropdown pick (`enterValue`).
// - Sheets: a tab clicked while pointing shows that sheet and keeps editing;
//   commit or cancel returns to the edited sheet.
import { computed, nextTick, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import { MAX_TEXT_LENGTH } from "../../lib/spreadsheet/formula/values";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { createCommitPath, type EntryHooks } from "./cellCommit";
import { createDraft, type SurfaceKind } from "./cellDraft";
import { createCellMoves, editKey, readyKey, type EditMode, type EditorKeyActions, type Move } from "./cellEditorKeys";
import { canEditCell, cellEditText, clearTargets, refusalFor, writeClear } from "./cellEntry";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { usePointMode } from "./usePointMode";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";
import type { EntryGate } from "./validationGate";

export type { EditMode } from "./cellEditorKeys";
export type { DraftSurface, SurfaceKind } from "./cellDraft";

export interface CellEditingDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
  };
  selection: Pick<SheetSelectionController, "selection" | "set" | "formatContext" | "styleAt">;
  sheetMetrics: Pick<SheetMetrics, "nav">;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  locale: () => SpreadsheetLocale;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
  /** Row autofit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
  /** Writes the view state (each sheet's selection) into the live model
   * (useSheetViewState's syncInto) — SPEC-121's rule for every command:
   * before a write records its undo step, so the step holds the selection
   * it was made from, and after a commit's move, since applying a step
   * re-reads the view state from the model. */
  syncViewState?: (model: SpreadsheetModel) => void;
  /** SPEC-130: the data-validation entry gate (useDataValidation's
   * entryGateOf); absent = every entry is written unchecked. */
  entryGate?: EntryGate;
}

const MODE_LABELS: Readonly<Record<EditMode, string>> = { ready: "Ready", enter: "Enter", edit: "Edit" };

export function useCellEditing(deps: CellEditingDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  let measurer = deps.measurer;
  const mode = ref<EditMode>("ready");
  const draft = createDraft();
  /** The cell being edited, and its sheet's selection when the editor
   * opened (what Ctrl/Cmd+Enter fills — not the selection of another sheet
   * shown while pointing). */
  const target = shallowRef<{ sheetId: string; at: CellRef; ranges: GridRange[] } | null>(null);
  /** Bumped when a key moved the active cell, so the grid reveals it. */
  const revealTick = ref(0);
  let original = "";
  let composing = false;
  let keyFilter: ((event: KeyboardEvent) => boolean) | null = null;
  /** SPEC-133: the typed-entry hooks (useSheetTables installs them). */
  let entryHooks: EntryHooks | null = null;

  const editing = computed(() => mode.value !== "ready");
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  /** The view state into the live model (see `syncViewState`). */
  const syncViewState = (): void => {
    const m = model();
    if (m) {
      deps.syncViewState?.(m);
    }
  };
  const point = usePointMode({
    text: () => draft.text.value,
    caret: () => draft.caret.value,
    setDraft: draft.set,
    editing: () => editing.value,
    editedSheet: () => model()?.sheets.find((sheet) => sheet.id === target.value?.sheetId) ?? null,
    shownSheet: () => deps.book.activeSheet.value,
    shownSelection: () => deps.selection.selection.value,
    nav: () => deps.sheetMetrics.nav(),
  });
  const moves = createCellMoves({
    sheetId: () => deps.book.activeSheet.value?.id ?? "",
    selection: () => deps.selection.selection.value,
    set: (next) => deps.selection.set(next),
    nav: () => deps.sheetMetrics.nav(),
  });

  /** "Ready", "Enter", "Edit" or "Point" — the status bar's mode. */
  const modeLabel = computed(() => (editing.value && point.pointing.value ? "Point" : MODE_LABELS[mode.value]));

  /** Opens the editor on the active cell, unless protection or an array
   * refuses it (a toast). */
  function open(next: EditMode, typed: string | null, at: number | "end"): boolean {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    if (!m || !sheet || sheet.kind !== "grid") {
      return false;
    }
    const active = deps.selection.selection.value.active;
    const refusal = refusalFor(m, deps.book.activeIndex.value, [{ top: active.row, left: active.col, bottom: active.row, right: active.col }]);
    if (refusal) {
      toast(refusal);
      return false;
    }
    original = cellEditText(m, sheet, active, deps.selection.styleAt(sheet, active.row, active.col), deps.selection.formatContext.value);
    target.value = { sheetId: sheet.id, at: { ...active }, ranges: deps.selection.selection.value.ranges };
    mode.value = next;
    const text = typed ?? original;
    draft.reset(text, at === "end" ? text.length : at);
    return true;
  }

  /** A host's input, read back: opens Enter mode from Ready (the input
   * replaces the content); false when refused (the host re-renders). */
  function input(next: string, at: number): boolean {
    if (mode.value === "ready") {
      return open("enter", next, at);
    }
    if (next.length > MAX_TEXT_LENGTH) {
      return false;
    }
    point.end();
    draft.set(next, at);
    return true;
  }

  /** An IME composition starts: from Ready it opens Enter mode (the
   * composition replaces the content once it ends). */
  function compositionStart(): boolean {
    composing = true;
    return mode.value === "ready" ? open("enter", "", 0) : true;
  }

  /** F2, a double-click (`point`: the caret goes to the character there),
   * a click in the formula bar (`at`). */
  function openEdit(options: { at?: number; point?: { x: number; y: number } } = {}): boolean {
    if (editing.value) {
      mode.value = "edit";
      point.end();
      return true;
    }
    if (!open("edit", null, options.at ?? "end")) {
      return false;
    }
    const where = options.point;
    if (where) {
      void nextTick(() => {
        const offset = draft.caretAt(where.x, where.y);
        if (offset !== null) {
          draft.caret.value = offset;
          draft.revision.value += 1;
        }
      });
    }
    return true;
  }

  /** Back to Ready, on the edited sheet. */
  function finish(): void {
    const sheetId = target.value?.sheetId;
    mode.value = "ready";
    target.value = null;
    path.forget();
    composing = false;
    point.end();
    draft.reset();
    const index = model()?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
    if (index >= 0 && index !== deps.book.activeIndex.value) {
      draft.requestFocus();
      deps.book.selectSheet(index);
    }
  }

  /** Back to Ready after a write, applying the commit's move. */
  function done(move: Move | null): void {
    finish();
    if (move && moves.apply(move)) {
      revealTick.value += 1;
    }
    syncViewState();
  }

  /** The commit path: commit, Fill and a dropdown pick (cellCommit.ts). */
  const path = createCommitPath({
    model,
    target: () => target.value,
    draft,
    original: () => original,
    editing: () => editing.value,
    locale: deps.locale,
    format: () => deps.selection.formatContext.value,
    undo: deps.undo,
    measurer: () => (measurer ??= createCanvasCellMeasurer()),
    toast,
    syncViewState,
    open: (text) => open("enter", text, "end"),
    finish,
    done,
    edit: () => {
      mode.value = "edit";
      point.end();
    },
    entryGate: deps.entryGate,
    entryHooks: () => entryHooks,
  });
  const { commit, fill } = path;

  /** Delete/Backspace in Ready (and Format ▸ Clear contents): see clearTargets. */
  function clearSelection(): void {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    if (!m || sheet?.kind !== "grid") {
      return;
    }
    const ranges = clearTargets(sheet, deps.selection.selection.value.ranges);
    const refusal = refusalFor(m, deps.book.activeIndex.value, ranges);
    if (refusal) {
      toast(refusal);
      return;
    }
    measurer ??= createCanvasCellMeasurer();
    syncViewState();
    writeClear({ model: m, sheetIndex: deps.book.activeIndex.value, undo: deps.undo, measurer, format: deps.selection.formatContext.value }, ranges);
  }

  const actions: EditorKeyActions = {
    mode,
    point,
    moves,
    revealed: () => {
      revealTick.value += 1;
    },
    openEdit: () => openEdit(),
    focusCell: draft.focusCell,
    clearSelection,
    cancel: finish,
    commit,
    fill,
    insertText: (text) => {
      point.end();
      draft.replaceSelection(text, MAX_TEXT_LENGTH);
    },
    historyStep: (back) => {
      if (draft.step(back)) {
        point.end();
      }
    },
  };

  /** The editor hosts' keydown. True when the key is the editor's: the
   * caller stops it there (no grid navigation, no window shortcut). */
  function onKeydown(event: KeyboardEvent): boolean {
    if (event.isComposing || composing || event.keyCode === 229) {
      return editing.value || composing;
    }
    draft.syncCaret();
    if (editing.value && keyFilter?.(event)) {
      return true; // SPEC-127: the function assists' keys come first
    }
    return editing.value ? editKey(event, actions) : readyKey(event, actions);
  }

  /** A press on a grid cell: "native" inside the editor (the caret moves;
   * Enter mode becomes Edit), "handled" when it pointed (or a failed commit
   * kept the editor open), "select" when the grid selects as usual. */
  function pressCell(cell: CellRef, shift: boolean, inEditor: boolean): "native" | "handled" | "select" {
    moves.forget();
    if (!editing.value) {
      return "select";
    }
    if (inEditor) {
      mode.value = "edit";
      point.end();
      return "native";
    }
    if (point.press(cell, shift)) {
      return "handled";
    }
    return commit(null) ? "select" : "handled";
  }

  /** A sheet tab chosen (SheetTabs' select): while pointing the sheet is
   * shown and editing goes on; otherwise an edit commits first. */
  function selectSheet(index: number): void {
    if (editing.value && point.allowed()) {
      draft.requestFocus();
      deps.book.selectSheet(index);
      draft.focusCell();
      return;
    }
    if (!editing.value || commit(null)) {
      deps.book.selectSheet(index);
    }
  }

  return {
    mode,
    modeLabel,
    editing,
    draft,
    text: draft.text,
    caret: draft.caret,
    error: draft.error,
    revision: draft.revision,
    target,
    revealTick,
    point,
    /** Whether the edited cell is on the sheet on show (Ready: true). */
    onShownSheet: computed(() => target.value === null || target.value.sheetId === deps.book.activeSheet.value?.id),
    input,
    compositionStart,
    compositionEnd: (): void => void (composing = false),
    replaceRange: draft.replaceRange,
    /** SPEC-127's Insert Function from Ready: edits the active cell with
     * `text`, caret at the end; false when refused (a toast). */
    startEditing: (text: string): boolean => open("enter", text, "end") && (draft.focusCell(), true),
    /** Whether protection lets the active cell be edited (SPEC-127's fx). */
    canEditActiveCell: (): boolean => canEditCell(model(), deps.book.activeIndex.value, deps.book.activeSheet.value, deps.selection.selection.value.active),
    /** Keys a host's assist takes before the editor's own table. */
    setKeyFilter: (filter: ((event: KeyboardEvent) => boolean) | null): void => void (keyFilter = filter),
    /** SPEC-133: installs the typed-entry hooks (cellCommit.ts). */
    setEntryHooks: (hooks: EntryHooks | null): void => void (entryHooks = hooks),
    openEdit,
    commit: (): boolean => commit(null),
    cancel: finish,
    fill,
    /** SPEC-130's dropdown pick (cellCommit.ts). */
    enterValue: path.enterValue,
    clearSelection,
    /** Replaces the focused host's selection (paste, Alt+Enter). */
    insertText: actions.insertText,
    onKeydown,
    pressCell,
    /** A press on a header or the corner: an edit commits first (false:
     * it could not, the editor stays open). */
    pressOutside: (): boolean => (moves.forget(), !editing.value || commit(null)),
    pointDrag: point.drag,
    pointRelease: point.release,
    selectSheet,
    register: (kind: SurfaceKind, surface: Parameters<typeof draft.register>[1]): (() => void) => draft.register(kind, surface),
    focusCell: draft.focusCell,
  };
}

export type CellEditing = ReturnType<typeof useCellEditing>;

/** Narrows a component's runtime-typed `editor` prop (the .vue parser bars
 * casts). */
export function asCellEditing(value: unknown): CellEditing {
  return value as CellEditing;
}
