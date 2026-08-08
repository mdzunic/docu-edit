// Footnotes and endnotes (SPEC-031): the note being edited (the same
// one-live-editable-surface shape as the header/footer band), inserting a note,
// its per-stream settings dialog, and note-edit caret handling (including IME
// composition). Note content is addressed through withNoteModel's synthetic
// single-section coordinates (section 0); applyNoteEdit folds a command result
// back into the real document and commits it — applyEdit's note sibling.
import { nextTick, ref } from "vue";
import type { Ref } from "vue";
import { modelRangeFromSelection, setDomCaret, setDomSelection } from "../../lib/domCaret";
import { blockLength, noCaret } from "../../lib/textEditing";
import {
  caretSelection,
  commitNoteModel,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentStart,
  insertLineBreak,
  insertText,
  insertNote,
  isCollapsed,
  noPendingFormat,
  noSelection,
  setNoteSettings,
  splitBlock,
  withNoteModel,
} from "../../lib/commands";
import { recordEdit as recordHistoryEdit } from "../../lib/undoHistory";
import { replaceDocumentModel } from "../../lib/documentStore";
import {
  asNoteKind,
  asNoteNumberFormat,
  asNoteRestart,
} from "../../lib/documentModel";
import { collectNoteRefs, createNoteId, effectiveNoteSettings, noteById } from "../../lib/notes";
import type { CaretPosition } from "../../lib/textEditing";
import type { ModelSelection } from "../../lib/commands";
import type { DocumentModel } from "../../lib/documentModel";

export interface NotesDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  editable: Readonly<Ref<boolean>>;
  documentId: Readonly<Ref<string>>;
  activeHeaderFooter: Ref<any>;
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  applyEdit: (
    result: { model: DocumentModel; selection: ModelSelection } | null,
    beforeSelection: ModelSelection | null,
    label: string,
  ) => void;
  applyCaretMove: (target?: CaretPosition, extend?: boolean) => void;
  scrollModelCaretIntoView: (position?: CaretPosition, always?: boolean) => Promise<void>;
  clearCellSelection: () => void;
  clearImageSelection: () => void;
  /** Leaves header/footer editing before entering a note (one live surface). */
  exitHeaderFooterEditing: () => void;
  getCaret: () => CaretPosition | null;
  markDirty: () => void;
  refreshUndoRedoState: () => void;
}

export function useNotes(deps: NotesDeps) {
  const {
    doc,
    contentRoot,
    editable,
    documentId,
    activeHeaderFooter,
    currentSelection,
    applyEdit,
    applyCaretMove,
    scrollModelCaretIntoView,
    clearCellSelection,
    clearImageSelection,
    exitHeaderFooterEditing,
    getCaret,
    markDirty,
    refreshUndoRedoState,
  } = deps;

  const activeNote = ref<{ noteId: string; kind: string } | null>(null);
  let noteCaret: CaretPosition | null = noCaret();
  let noteSelectionAnchor: CaretPosition | null = noCaret();
  let notePendingFormat: ReturnType<typeof noPendingFormat> = null;
  let noteComposing = false;
  let noteCompositionRange: ModelSelection | null = noSelection();
  const noteRemountEpoch = ref(0);
  const noteSettingsOpen = ref(false);
  const noteSettingsInitial = ref({});
  void noteSettingsInitial;

  /** The one DOM element currently rendering as an editable note body. */
  function noteRootEl(): HTMLElement | null {
    const found = contentRoot.value?.querySelector("[data-note-editable]");
    return found instanceof HTMLElement ? found : null;
  }

  function currentNoteSelection(): ModelSelection | null {
    const root = noteRootEl();
    if (root) {
      const range = modelRangeFromSelection(root, document.getSelection());
      if (range) {
        return { anchor: range.start, focus: range.end };
      }
    }
    return noteCaret
      ? { anchor: noteSelectionAnchor ?? noteCaret, focus: noteCaret }
      : null;
  }

  /** Leaves note editing and returns focus to the body. */
  function exitNoteEditing(): void {
    if (!activeNote.value) {
      return;
    }
    activeNote.value = null;
    noteCaret = noCaret();
    noteSelectionAnchor = noCaret();
    notePendingFormat = noPendingFormat();
    void nextTick().then(() => (contentRoot.value as HTMLElement | null)?.focus());
  }

  /** Enters note editing (AC-2, AC-4, AC-15). */
  function enterNoteEditing(noteId = "", kind = asNoteKind("footnote"), atEnd = false) {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    clearCellSelection();
    clearImageSelection();
    exitHeaderFooterEditing();
    activeNote.value = { noteId, kind };
    const note = noteById(model, noteId);
    const lastBlock = note ? Math.max(0, note.blocks.length - 1) : 0;
    const lastLength = note ? blockLength(note.blocks[lastBlock]) : 0;
    noteCaret = atEnd
      ? { section: 0, block: lastBlock, offset: lastLength }
      : { section: 0, block: 0, offset: 0 };
    noteSelectionAnchor = noCaret();
    notePendingFormat = noPendingFormat();
    void nextTick().then(() => {
      const root = noteRootEl();
      if (root) {
        root.focus();
        setDomCaret(root, noteCaret ?? { section: 0, block: 0, offset: 0 });
        root.scrollIntoView({ block: "nearest" });
      }
    });
  }

  /** Folds a note command's result (built against withNoteModel) back into the
   * real document and commits it — applyEdit's note sibling. */
  function applyNoteEdit(
    resultModel: DocumentModel,
    resultSelection: ModelSelection,
    label = "",
  ): void {
    const active = activeNote.value;
    const before = doc.value;
    if (!active || !before || !resultSelection) {
      return;
    }
    const nextModel = commitNoteModel(before, active.noteId, resultModel);
    const bodyFiller = caretSelection(getCaret() ?? documentStart());
    recordHistoryEdit(documentId.value, {
      before: { model: before, selection: bodyFiller },
      after: { model: nextModel, selection: bodyFiller },
      label,
      groupable: false,
    });
    doc.value = replaceDocumentModel(documentId.value, nextModel);
    noteCaret = resultSelection.focus;
    noteSelectionAnchor = isCollapsed(resultSelection) ? noCaret() : resultSelection.anchor;
    markDirty();
    refreshUndoRedoState();
    void nextTick().then(() => {
      const root = noteRootEl();
      if (!root || !noteCaret) {
        return;
      }
      if (noteSelectionAnchor) {
        setDomSelection(root, noteSelectionAnchor, noteCaret);
      } else {
        setDomCaret(root, noteCaret);
      }
    });
  }

  function handleNoteBeforeInput(event: InputEvent) {
    const type = event.inputType;
    if (noteComposing || event.isComposing || type === "insertCompositionText") {
      return;
    }
    event.preventDefault();
    const active = activeNote.value;
    const model = doc.value;
    if (!active || !model || !editable.value) {
      return;
    }
    const selection = currentNoteSelection();
    if (!selection) {
      return;
    }
    const synthetic = withNoteModel(model, active.noteId);
    let result: { model: DocumentModel; selection: ModelSelection } | null = null;
    let label = "";
    if (type === "insertText" || type === "insertFromPaste") {
      const text = event.data ?? (event.dataTransfer ? event.dataTransfer.getData("text/plain") : "");
      if (text) {
        result = insertText(synthetic, selection, text, notePendingFormat);
        label = "typing";
      }
    } else if (type === "insertParagraph") {
      result = splitBlock(synthetic, selection);
      label = "new paragraph";
    } else if (type === "insertLineBreak") {
      result = insertLineBreak(synthetic, selection);
      label = "line break";
    } else if (type === "deleteContentBackward") {
      result = deleteBackward(synthetic, selection);
      label = "delete";
    } else if (type === "deleteContentForward") {
      result = deleteForward(synthetic, selection);
      label = "delete";
    } else if (type === "deleteByCut" || type === "deleteContent") {
      result = isCollapsed(selection) ? null : deleteRange(synthetic, selection);
      label = "delete";
    }
    if (result) {
      notePendingFormat = noPendingFormat();
      applyNoteEdit(result.model, result.selection, label);
    }
  }

  function handleNoteKeyDown(event: KeyboardEvent) {
    if (event.key === "Escape") {
      event.stopPropagation();
      exitNoteEditing();
    }
  }

  function handleNoteCompositionStart() {
    if (!noteRootEl() || !editable.value) {
      return;
    }
    noteComposing = true;
    noteCompositionRange = currentNoteSelection();
  }

  function handleNoteCompositionEnd(event: CompositionEvent) {
    noteComposing = false;
    const range = noteCompositionRange;
    noteCompositionRange = null;
    const active = activeNote.value;
    const model = doc.value;
    if (!active || !model || !editable.value || !range) {
      return;
    }
    const data = event.data ?? "";
    if (!data) {
      return;
    }
    const synthetic = withNoteModel(model, active.noteId);
    const result = insertText(synthetic, range, data, null);
    applyNoteEdit(result.model, result.selection, "typing");
    noteRemountEpoch.value += 1;
  }

  /** Clicking a read-only note body activates editing it (AC-6). */
  function handleNoteActivate(noteId = "", kind = "") {
    enterNoteEditing(noteId, asNoteKind(kind), true);
  }

  /** Clicking a note's number jumps back to its reference mark (AC-16). */
  function handleNoteNumberClick(noteId = "") {
    const model = doc.value;
    if (!model) {
      return;
    }
    const kind = noteById(model, noteId)?.kind ?? "footnote";
    const ref = collectNoteRefs(model, kind).find((entry) => entry.noteId === noteId);
    if (!ref) {
      return;
    }
    // Leave note editing inline (not through exitNoteEditing, whose deferred
    // refocus would race the caret placement below).
    activeNote.value = null;
    noteCaret = noCaret();
    noteSelectionAnchor = noCaret();
    notePendingFormat = noPendingFormat();
    const target = ref.cell
      ? { section: ref.section, block: ref.block, offset: ref.offset + 1, cell: ref.cell }
      : { section: ref.section, block: ref.block, offset: ref.offset + 1 };
    void nextTick().then(() => {
      (contentRoot.value as HTMLElement | null)?.focus();
      applyCaretMove(target);
      void scrollModelCaretIntoView(target, true);
    });
  }

  /** "Insert → Footnote/Endnote" (AC-1, AC-3): inserts the reference mark at
   * the caret, then moves the caret into the new note's area (AC-2, AC-4). */
  function handleInsertNote(kind = "") {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value || activeHeaderFooter.value || activeNote.value) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const noteId = createNoteId();
    const noteKind = asNoteKind(kind);
    applyEdit(insertNote(model, selection, noteKind, noteId), selection, `insert ${kind}`);
    enterNoteEditing(noteId, noteKind);
  }

  /** The footnote settings dialog (AC-17..20). */
  function openNoteSettings() {
    const model = doc.value;
    if (!model) {
      return;
    }
    noteSettingsInitial.value = {
      kind: activeNote.value?.kind ?? "footnote",
      footnote: effectiveNoteSettings(model, "footnote"),
      endnote: effectiveNoteSettings(model, "endnote"),
    };
    noteSettingsOpen.value = true;
  }

  function handleNoteSettingsApply(
    payload = { kind: "footnote", numberFormat: "decimal", startAt: 1, restart: "continuous" },
  ) {
    noteSettingsOpen.value = false;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      setNoteSettings(model, selection, asNoteKind(payload.kind), {
        numberFormat: asNoteNumberFormat(payload.numberFormat),
        startAt: payload.startAt,
        restart: asNoteRestart(payload.restart),
      }),
      selection,
      "note settings",
    );
  }

  function handleNoteSettingsCancel() {
    noteSettingsOpen.value = false;
  }

  // --- note-mirror formatting (used by useFormatting's triple-mirrored
  // toolbar commands) ------------------------------------------------

  function getNotePendingFormat(): ReturnType<typeof noPendingFormat> {
    return notePendingFormat;
  }
  function setNotePendingFormat(pending: ReturnType<typeof noPendingFormat>): void {
    notePendingFormat = pending;
  }

  return {
    activeNote,
    noteRemountEpoch,
    noteSettingsOpen,
    noteSettingsInitial,
    noteRootEl,
    currentNoteSelection,
    applyNoteEdit,
    getNotePendingFormat,
    setNotePendingFormat,
    exitNoteEditing,
    enterNoteEditing,
    handleNoteActivate,
    handleNoteNumberClick,
    handleInsertNote,
    handleNoteBeforeInput,
    handleNoteKeyDown,
    handleNoteCompositionStart,
    handleNoteCompositionEnd,
    openNoteSettings,
    handleNoteSettingsApply,
    handleNoteSettingsCancel,
  };
}
