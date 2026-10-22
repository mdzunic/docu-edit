// Notes on cells, as the grid shows and edits them (SPEC-135 "The grid shows
// annotations", the hover popovers, the note commands). The model and its
// commands are lib/spreadsheet/annotations.ts'; this composable is the view
// side of the notes and of the thread popover:
// - the indicators: `marksAt(row, col)` — a note's red corner triangle
//   (`sheet-note-indicator`) and a thread's purple marker
//   (`sheet-thread-indicator`, dimmed when resolved) on the sheet on show,
//   which SheetCellLayer draws without restyling the cell;
// - the active cell's annotation text for assistive technology
//   (`activeDescription` — SheetNoteLayer's visually hidden element, which
//   useSheetGrid names in the active cell's aria-describedby);
// - hovering: an annotated cell under the pointer for NOTE_HOVER_DELAY_MS
//   shows its popover (`popover` — SheetNotePopover: a note's author then its
//   text, a thread's opening message, "N replies" and `sheet-thread-open`,
//   which opens the comments panel on that thread through `openThread`); the
//   pointer may travel from the cell into the popover (a short grace before
//   it hides);
// - the note boxes (`boxes`): every `visible` note on the sheet on show, or
//   every note while View ▸ "Show all notes" is on (`showAll` — session-only,
//   never written), at the default placement and the note's `box` size;
// - the note editor (`editor`, `draft`): Insert ▸ Note, Shift+F2
//   (`sheet.note`, from the grid only, never while a cell is edited) and the
//   cell context menu open a plain textarea on the active cell — on its note
//   when it has one; Escape, Ctrl/Cmd+Enter and leaving it commit (Excel
//   keeps a note's text on Escape). An unchanged draft changes nothing (a
//   rich note keeps its runs), an emptied one deletes the note;
// - "Delete note" and "Show note" / "Hide note" on the active cell.
// Every command is ONE SPEC-121 undo step on the note's sheet — view state
// synced, `record(label, { sheetId })`, the annotations.ts command,
// `commit()` — labelled NOTE_UNDO_LABELS'. The author of a new note is the
// settable comment identity (loadCommentAuthor). Notes are not offered on a
// CSV or TSV workbook (nothing could keep them) nor off a grid sheet, and
// are refused, with `noteDisabledReason`, on an .xlsx sheet whose source's
// notes the reader could not model (`blocked` — AC-11: they stay verbatim,
// so a new one could not be saved).
// Annotation texts are plain strings everywhere; the components render them
// with text interpolation only.
import { computed, onScopeDispose, ref, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { loadCommentAuthor } from "../../lib/core/commentIdentity";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import { formatA1, type CellRef } from "../../lib/spreadsheet/address";
import { addNote, deleteNote, editNote, NOTE_DEFAULT_BOX, NOTE_HOVER_DELAY_MS, NOTE_UNDO_LABELS, noteAt, setNoteVisible, sheetNotes, threadAt } from "../../lib/spreadsheet/annotations";
import { isCsvFormat, type Sheet, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import { annotationDescription, annotationMarks, marksKey, noteBoxSize, noteBoxViews, notePopoverView, type AnnotationMarks, type NoteBoxView, type NotePopoverView } from "./cellNoteViews";
import { inGrid } from "./headerMenu";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export { repliesLabel } from "./cellNoteViews";

/** How long the popover stays once the pointer leaves its cell — room to
 * reach the popover (and its `sheet-thread-open`). */
export const NOTE_HIDE_GRACE_MS = 300;

/** The id of the visually hidden element holding the active cell's
 * annotation text (the active cell's aria-describedby names it). */
export const NOTE_DESCRIPTION_ID = "sheet-annotation-description";

/** The open note editor: the cell, the note it edits (null: a new one) and
 * the box it shows in. */
export interface NoteEditorState {
  sheetId: string;
  ref: CellRef;
  noteId: string | null;
  widthPx: number;
  heightPx: number;
}

export interface CellNotesDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  /** The active cell the commands act on. */
  selection: Pick<SheetSelectionController, "selection">;
  /** Syncs the view state into the model before a step is recorded. */
  syncInto?: (model: SpreadsheetModel) => void;
  /** Whether a cell is being edited: notes are inert meanwhile. */
  busy?: () => boolean;
  /** `sheet-thread-open`: the comments panel on that thread. */
  openThread?: (sheetId: string, threadId: string) => void;
  /** Focus back to the grid (the in-cell editor) once the note editor
   * commits from the keyboard. */
  focusGrid?: () => void;
  /** The author of a new note — default: the settable comment identity. */
  author?: () => string;
  /** Why the sheet on show's notes cannot be changed ("" when they can) —
   * useAnnotationEditing's notesReason: its source broke the subset. */
  blocked?: () => string;
}

const sameRef = (a: CellRef, b: CellRef): boolean => a.row === b.row && a.col === b.col;

export function useCellNotes(deps: CellNotesDeps) {
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const busy = (): boolean => deps.busy?.() === true;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet?.kind === "grid" ? sheet : null;
  };

  /** Why notes cannot be added on the sheet on show although it could hold
   * them — its source's notes are kept as read ("" when they can). */
  const noteDisabledReason = computed(() => {
    const m = model();
    return m !== null && !isCsvFormat(m.format) && gridSheet() !== null ? (deps.blocked?.() ?? "") : "";
  });
  /** Whether notes can be added here: a grid sheet of a workbook that can
   * keep them (not CSV or TSV) whose source's notes were modelled. */
  const canAnnotate = computed(() => {
    const m = model();
    return m !== null && !isCsvFormat(m.format) && gridSheet() !== null && noteDisabledReason.value === "";
  });

  // --- indicators and the active cell's description ----------------------------------------------

  /** The sheet on show's marks by marksKey. */
  const marks = computed(() => {
    const sheet = gridSheet();
    return sheet ? annotationMarks(sheet) : new Map<string, AnnotationMarks>();
  });

  /** A cell's marks on the sheet on show (null: none). */
  function marksAt(row: number, col: number): AnnotationMarks | null {
    return marks.value.get(marksKey(row, col)) ?? null;
  }

  const active = (): CellRef => deps.selection.selection.value.active;
  /** The active cell's note (absent: none). */
  const activeNote = computed(() => {
    const sheet = gridSheet();
    return sheet ? noteAt(sheet, active()) : undefined;
  });
  const activeDescription = computed(() => {
    const sheet = gridSheet();
    return sheet ? annotationDescription(noteAt(sheet, active()), threadAt(sheet, active())) : "";
  });

  // --- show all, the boxes -----------------------------------------------------------------------

  const showAll = ref(false);
  function toggleShowAll(): void {
    showAll.value = !showAll.value;
  }

  const editor = shallowRef<NoteEditorState | null>(null);
  const draft = ref("");

  /** The notes shown as boxes on the sheet on show — not the one in the editor. */
  const boxes = computed((): NoteBoxView[] => {
    const sheet = gridSheet();
    const open = editor.value;
    return sheet ? noteBoxViews(sheet, showAll.value, open && open.sheetId === sheet.id ? open.noteId : null) : [];
  });

  // --- hovering ------------------------------------------------------------------------------------

  const hovered = shallowRef<{ sheetId: string; ref: CellRef } | null>(null);
  /** The cell the pointer rests on, waiting out the delay. */
  let pending: { sheetId: string; ref: CellRef } | null = null;
  let showTimer: ReturnType<typeof setTimeout> | null = null;
  let hideTimer: ReturnType<typeof setTimeout> | null = null;

  function clearShow(): void {
    if (showTimer !== null) {
      clearTimeout(showTimer);
      showTimer = null;
    }
    pending = null;
  }
  function clearHide(): void {
    if (hideTimer !== null) {
      clearTimeout(hideTimer);
      hideTimer = null;
    }
  }
  function scheduleHide(): void {
    if (hovered.value !== null && hideTimer === null) {
      hideTimer = setTimeout(() => {
        hideTimer = null;
        hovered.value = null;
      }, NOTE_HIDE_GRACE_MS);
    }
  }
  /** Hides the popover now (a press, a scroll, the editor opening). */
  function dismissPopover(): void {
    clearShow();
    clearHide();
    hovered.value = null;
  }

  /** The pointer over a cell of the sheet on show (null: over no cell). */
  function hoverCell(cell: CellRef | null): void {
    const sheet = gridSheet();
    if (hovered.value && cell && sheet && hovered.value.sheetId === sheet.id && sameRef(hovered.value.ref, cell)) {
      clearShow();
      clearHide();
      return;
    }
    if (pending && cell && sameRef(pending.ref, cell)) {
      return;
    }
    clearShow();
    scheduleHide();
    if (!cell || !sheet || busy() || editor.value || marksAt(cell.row, cell.col) === null) {
      return;
    }
    pending = { sheetId: sheet.id, ref: { row: cell.row, col: cell.col } };
    showTimer = setTimeout(() => {
      showTimer = null;
      const at = pending;
      pending = null;
      clearHide();
      hovered.value = at;
    }, NOTE_HOVER_DELAY_MS);
  }

  /** The pointer entering (true) or leaving (false) the popover. */
  function hoverPopover(inside: boolean): void {
    if (inside) {
      clearShow();
      clearHide();
    } else {
      scheduleHide();
    }
  }

  const popover = computed((): NotePopoverView | null => {
    const at = hovered.value;
    const sheet = gridSheet();
    return at && sheet && sheet.id === at.sheetId ? notePopoverView(sheet, at.ref, showAll.value) : null;
  });

  /** `sheet-thread-open`: the comments panel on the popover's thread. */
  function openPopoverThread(): void {
    const view = popover.value;
    if (view?.thread) {
      dismissPopover();
      deps.openThread?.(view.sheetId, view.thread.threadId);
    }
  }

  // --- commands ----------------------------------------------------------------------------------

  /** One undo step on `sheetId`. */
  function run(label: string, sheetId: string, command: (m: SpreadsheetModel) => void): void {
    const m = model();
    if (!m) {
      return;
    }
    deps.syncInto?.(m);
    // record() first: it settles a running recalculation, which can replace
    // the sheet the command copies.
    deps.undo.record(label, { sheetId });
    command(m);
    deps.undo.commit();
  }

  /** Opens the note editor on `at` (default: the active cell) — on its note
   * when it has one. False when notes are not offered here. */
  function openEditor(at?: CellRef): boolean {
    const sheet = gridSheet();
    if (!sheet || !canAnnotate.value || busy()) {
      return false;
    }
    commitEditor(false);
    const target = at ?? active();
    const note = noteAt(sheet, target);
    dismissPopover();
    draft.value = note?.text ?? "";
    editor.value = { sheetId: sheet.id, ref: { row: target.row, col: target.col }, noteId: note?.noteId ?? null, ...(note ? noteBoxSize(note) : NOTE_DEFAULT_BOX) };
    return true;
  }

  function input(text: string): void {
    draft.value = text;
  }

  /** Commits the open editor (Escape, Ctrl/Cmd+Enter, leaving it): adds,
   * edits or — emptied — deletes the note; an unchanged draft changes
   * nothing. `refocus`: hand focus back to the grid (a key committed). */
  function commitEditor(refocus = true): void {
    const open = editor.value;
    if (!open) {
      return;
    }
    editor.value = null;
    const text = draft.value;
    const sheet = model()?.sheets.find((each) => each.id === open.sheetId);
    const note = open.noteId === null || !sheet ? undefined : sheetNotes(sheet).find((each) => each.noteId === open.noteId);
    if (sheet && note && text !== note.text) {
      const label = text.trim() === "" ? NOTE_UNDO_LABELS.delete : NOTE_UNDO_LABELS.edit;
      run(label, open.sheetId, (m) => void editNote(m, open.sheetId, note.noteId, text));
    } else if (sheet && !note && text.trim() !== "") {
      const author = (deps.author ?? loadCommentAuthor)();
      run(NOTE_UNDO_LABELS.add, open.sheetId, (m) => void addNote(m, open.sheetId, open.ref, text, author));
    }
    if (refocus) {
      deps.focusGrid?.();
    }
  }

  /** "Delete note" on the active cell. */
  function deleteActive(): void {
    const sheet = gridSheet();
    const note = activeNote.value;
    if (sheet && note && !busy()) {
      run(NOTE_UNDO_LABELS.delete, sheet.id, (m) => void deleteNote(m, sheet.id, note.noteId));
    }
  }

  /** "Show note" / "Hide note" on the active cell. */
  function toggleActiveVisible(): void {
    const sheet = gridSheet();
    const note = activeNote.value;
    if (sheet && note && !busy()) {
      const visible = note.visible !== true;
      run(visible ? NOTE_UNDO_LABELS.show : NOTE_UNDO_LABELS.hide, sheet.id, (m) => void setNoteVisible(m, sheet.id, note.noteId, visible));
    }
  }

  // The sheet on show changing commits the editor and drops the popover.
  watch(
    () => deps.book.activeSheet.value?.id,
    () => {
      commitEditor(false);
      dismissPopover();
    },
  );

  /** Shift+F2 from the grid: the note editor on the active cell. */
  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || busy()) {
      return;
    }
    if (findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id === "sheet.note" && openEditor()) {
      event.preventDefault();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
    clearShow();
    clearHide();
  });

  /** The editor's cell as A1 (its accessible name), "" when closed. */
  const editorAddress = computed(() => (editor.value ? formatA1(editor.value.ref) : ""));

  return {
    canAnnotate,
    noteDisabledReason,
    marksAt,
    activeNote,
    activeDescription,
    showAll,
    toggleShowAll,
    boxes,
    popover,
    hoverCell,
    hoverPopover,
    dismissPopover,
    openPopoverThread,
    editor,
    editorAddress,
    draft,
    openEditor,
    input,
    commitEditor,
    deleteActive,
    toggleActiveVisible,
    onKeydown,
  };
}

export type CellNotes = ReturnType<typeof useCellNotes>;

/** Narrows a component's runtime-typed `controller` prop (the .vue parser
 * bars casts). */
export function asCellNotes(value: unknown): CellNotes {
  return value as CellNotes;
}
