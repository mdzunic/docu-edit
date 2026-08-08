// Headers, footers and page numbers (SPEC-029). The one-live-editable-surface
// band (mirroring notes), header/footer/page options dialogs, and page-number
// insertion. `applyHeaderFooterEdit` is applyEdit's header/footer sibling: it
// folds a command result built against withHeaderFooterModel's synthetic
// single-section model back into the real document and commits it.
//
// Headers and notes each call the other's `exit*` at runtime (one live surface
// at a time), so headers receives `getExitNoteEditing` — a late-binding getter
// that resolves once useNotes is set up — to break the two-way dependency.
import { nextTick, ref } from "vue";
import type { Ref } from "vue";
import { noCaret } from "../../lib/textEditing";
import { modelRangeFromSelection, setDomCaret, setDomSelection } from "../../lib/domCaret";
import {
  caretSelection,
  commitHeaderFooterModel,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentStart,
  insertLineBreak,
  insertPageNumberCommand,
  insertText,
  isCollapsed,
  noPendingFormat,
  noSelection,
  setHeaderFooterOptions,
  splitBlock,
  withHeaderFooterModel,
} from "../../lib/commands";
import { recordEdit as recordHistoryEdit } from "../../lib/undoHistory";
import { replaceDocumentModel } from "../../lib/documentStore";
import {
  asHeaderFooterPart,
  asPageNumberFormat,
  asParagraphAlignment,
  noActiveHeaderFooter,
  type HeaderFooterKind,
  type HeaderFooterPart,
} from "../../lib/documentModel";
import { resolveHeaderFooterKind } from "../../lib/pageNumbering";
import type { CaretPosition } from "../../lib/textEditing";
import type { ModelSelection } from "../../lib/commands";
import type { DocumentModel } from "../../lib/documentModel";
import type { PageLayout } from "../../lib/pagination";

export interface HeadersFootersDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  editable: Readonly<Ref<boolean>>;
  documentId: Readonly<Ref<string>>;
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  applyEdit: (
    result: { model: DocumentModel; selection: ModelSelection } | null,
    beforeSelection: ModelSelection | null,
    label: string,
  ) => void;
  /** These are declared later in the view than this composable runs, so they
   * are passed as lazy getters evaluated only at call time. */
  getPaginationPages: () => PageLayout[];
  getPageNumbers: () => number[];
  getCaretSectionIndex: () => number;
  pageOfCaret: (caret: CaretPosition) => number;
  clearCellSelection: () => void;
  /** Late-binding: images is set up later than this composable, so its
   * clearImageSelection is resolved through a getter. */
  getClearImageSelection: () => () => void;
  /** Late-binding: the note composable's exitNoteEditing (set up later). */
  getExitNoteEditing: () => () => void;
  getCaret: () => CaretPosition | null;
  markDirty: () => void;
  refreshUndoRedoState: () => void;
}

export function useHeadersFooters(deps: HeadersFootersDeps) {
  const {
    doc,
    contentRoot,
    editable,
    documentId,
    currentSelection,
    applyEdit,
    getPaginationPages,
    getPageNumbers,
    getCaretSectionIndex,
    pageOfCaret,
    clearCellSelection,
    getClearImageSelection,
    getExitNoteEditing,
    getCaret,
    markDirty,
    refreshUndoRedoState,
  } = deps;

  const activeHeaderFooter = ref<{
    sectionIndex: number;
    part: HeaderFooterPart;
    kind: HeaderFooterKind;
    pageIndex: number;
  } | null>(noActiveHeaderFooter());
  let hfCaret: CaretPosition | null = noCaret();
  let hfSelectionAnchor: CaretPosition | null = noCaret();
  let hfPendingFormat: ReturnType<typeof noPendingFormat> = null;
  let hfComposing = false;
  let hfCompositionRange: ModelSelection | null = noSelection();
  const hfRemountEpoch = ref(0);
  const headerFooterOptionsOpen = ref(false);
  const headerFooterOptionsInitial = ref({});
  const pageNumberDialogOpen = ref(false);

  /** The one DOM element currently rendering as an editable header/footer band. */
  function hfRootEl(): HTMLElement | null {
    const found = contentRoot.value?.querySelector("[data-hf-editable]");
    return found instanceof HTMLElement ? found : null;
  }

  /** The header/footer content's live selection, falling back to hfCaret when
   * the DOM selection is not inside the band (a toolbar control took focus). */
  function currentHfSelection(): ModelSelection | null {
    const root = hfRootEl();
    if (root) {
      const range = modelRangeFromSelection(root, document.getSelection());
      if (range) {
        return { anchor: range.start, focus: range.end };
      }
    }
    return hfCaret ? { anchor: hfSelectionAnchor ?? hfCaret, focus: hfCaret } : null;
  }

  /** Leaves header/footer editing and returns focus to the body (AC-7, AC-8). */
  function exitHeaderFooterEditing(): void {
    if (!activeHeaderFooter.value) {
      return;
    }
    activeHeaderFooter.value = null;
    hfCaret = noCaret();
    hfSelectionAnchor = noCaret();
    void nextTick().then(() => (contentRoot.value as HTMLElement | null)?.focus());
  }

  /** Enters header/footer editing for the given page's band (AC-1, AC-2). */
  function enterHeaderFooterEditing(pageIndex = 0, part = asHeaderFooterPart("header")) {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    const pages = getPaginationPages();
    const page = pages[pageIndex];
    if (!page) {
      return;
    }
    const pageNumber = getPageNumbers()[pageIndex] ?? 1;
    const kind = resolveHeaderFooterKind(model, pages, pageIndex, pageNumber);
    clearCellSelection();
    getClearImageSelection()();
    getExitNoteEditing()();
    activeHeaderFooter.value = { sectionIndex: page.sectionIndex, part, kind, pageIndex };
    hfCaret = { section: 0, block: 0, offset: 0 };
    hfSelectionAnchor = noCaret();
    void nextTick().then(() => {
      const root = hfRootEl();
      if (root) {
        root.focus();
        setDomCaret(root, hfCaret ?? { section: 0, block: 0, offset: 0 });
      }
    });
  }

  function handleDblclickHeader(pageIndex = 0) {
    enterHeaderFooterEditing(pageIndex, "header");
  }
  function handleDblclickFooter(pageIndex = 0) {
    enterHeaderFooterEditing(pageIndex, "footer");
  }
  function handleDblclickBody() {
    exitHeaderFooterEditing();
    getExitNoteEditing()();
  }

  /** Opens "Insert -> Header/Footer" (AC-2): the caret's own section. */
  function openInsertHeaderFooter(part = asHeaderFooterPart("header")) {
    const model = doc.value;
    if (!model) {
      return;
    }
    const pageIndex = getCaret() ? pageOfCaret(getCaret() as CaretPosition) : 0;
    enterHeaderFooterEditing(pageIndex, part);
  }

  /** Folds a header/footer command's result back into the real document and
   * commits it — applyEdit's header/footer sibling. */
  function applyHeaderFooterEdit(
    resultModel: DocumentModel,
    resultSelection: ModelSelection,
    label = "",
  ): void {
    const active = activeHeaderFooter.value;
    const before = doc.value;
    if (!active || !before || !resultSelection) {
      return;
    }
    const nextModel = commitHeaderFooterModel(before, active.sectionIndex, active.part, active.kind, resultModel);
    const bodyFiller = caretSelection(getCaret() ?? documentStart());
    recordHistoryEdit(documentId.value, {
      before: { model: before, selection: bodyFiller },
      after: { model: nextModel, selection: bodyFiller },
      label,
      groupable: false,
    });
    doc.value = replaceDocumentModel(documentId.value, nextModel);
    hfCaret = resultSelection.focus;
    hfSelectionAnchor = isCollapsed(resultSelection) ? noCaret() : resultSelection.anchor;
    markDirty();
    refreshUndoRedoState();
    void nextTick().then(() => {
      const root = hfRootEl();
      if (!root || !hfCaret) {
        return;
      }
      if (hfSelectionAnchor) {
        setDomSelection(root, hfSelectionAnchor, hfCaret);
      } else {
        setDomCaret(root, hfCaret);
      }
    });
  }

  function handleHfBeforeInput(event: InputEvent) {
    const type = event.inputType;
    if (hfComposing || event.isComposing || type === "insertCompositionText") {
      return;
    }
    event.preventDefault();
    const active = activeHeaderFooter.value;
    const model = doc.value;
    if (!active || !model || !editable.value) {
      return;
    }
    const selection = currentHfSelection();
    if (!selection) {
      return;
    }
    const synthetic = withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind);
    let result: { model: DocumentModel; selection: ModelSelection } | null = null;
    let label = "";
    if (type === "insertText" || type === "insertFromPaste") {
      const text = event.data ?? (event.dataTransfer ? event.dataTransfer.getData("text/plain") : "");
      if (text) {
        result = insertText(synthetic, selection, text, hfPendingFormat);
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
      applyHeaderFooterEdit(result.model, result.selection, label);
    }
  }

  function handleHfKeyDown(event: KeyboardEvent) {
    if (event.key === "Escape") {
      event.stopPropagation();
      exitHeaderFooterEditing();
    }
  }

  function handleHfCompositionStart() {
    const root = hfRootEl();
    if (!root || !editable.value) {
      return;
    }
    hfComposing = true;
    hfCompositionRange = currentHfSelection();
  }

  function handleHfCompositionEnd(event: CompositionEvent) {
    hfComposing = false;
    const range = hfCompositionRange;
    hfCompositionRange = null;
    const active = activeHeaderFooter.value;
    const model = doc.value;
    if (!active || !model || !editable.value || !range) {
      return;
    }
    const data = event.data ?? "";
    if (!data) {
      return;
    }
    const synthetic = withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind);
    const result = insertText(synthetic, range, data, null);
    applyHeaderFooterEdit(result.model, result.selection, "typing");
    hfRemountEpoch.value += 1;
  }

  /** "Header/footer options" (AC-11..23), prefilled from the band's or the
   * caret's section. */
  function openHeaderFooterOptions() {
    const model = doc.value;
    if (!model) {
      return;
    }
    const sectionIndex = activeHeaderFooter.value?.sectionIndex ?? getCaretSectionIndex();
    const section = model.sections[sectionIndex];
    headerFooterOptionsInitial.value = {
      sectionIndex,
      titlePage: section?.titlePage === true,
      evenAndOddHeaders: section?.evenAndOddHeaders === true,
      pageNumberStart: section?.pageNumberStart ?? null,
    };
    headerFooterOptionsOpen.value = true;
  }

  function handleHeaderFooterOptionsApply(
    payload = { sectionIndex: 0, titlePage: false, evenAndOddHeaders: false, pageNumberStart: null },
  ) {
    headerFooterOptionsOpen.value = false;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      setHeaderFooterOptions(model, selection, payload.sectionIndex, {
        titlePage: payload.titlePage,
        evenAndOddHeaders: payload.evenAndOddHeaders,
        pageNumberStart: payload.pageNumberStart,
      }),
      selection,
      "header/footer options",
    );
  }

  function handleHeaderFooterOptionsCancel() {
    headerFooterOptionsOpen.value = false;
  }

  /** "Insert page number" (AC-14..19). */
  function openPageNumberDialog() {
    if (!doc.value) {
      return;
    }
    pageNumberDialogOpen.value = true;
  }

  function handlePageNumberInsert(
    options = { part: "footer", position: "center", format: "decimal", pageXofY: false },
  ) {
    pageNumberDialogOpen.value = false;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const sectionIndex = activeHeaderFooter.value?.sectionIndex ?? getCaretSectionIndex();
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      insertPageNumberCommand(model, selection, sectionIndex, {
        part: asHeaderFooterPart(options.part),
        position: asParagraphAlignment(options.position),
        format: asPageNumberFormat(options.format),
        pageXofY: !!options.pageXofY,
      }),
      selection,
      "insert page number",
    );
  }

  function handlePageNumberCancel() {
    pageNumberDialogOpen.value = false;
  }

  // --- header-mirror formatting (used by useFormatting's triple-mirrored
  // toolbar commands) ------------------------------------------------

  function getHfPendingFormat(): ReturnType<typeof noPendingFormat> {
    return hfPendingFormat;
  }
  function setHfPendingFormat(pending: ReturnType<typeof noPendingFormat>): void {
    hfPendingFormat = pending;
  }

  return {
    activeHeaderFooter,
    hfRemountEpoch,
    headerFooterOptionsOpen,
    headerFooterOptionsInitial,
    pageNumberDialogOpen,
    hfRootEl,
    currentHfSelection,
    applyHeaderFooterEdit,
    getHfPendingFormat,
    setHfPendingFormat,
    exitHeaderFooterEditing,
    enterHeaderFooterEditing,
    handleDblclickHeader,
    handleDblclickFooter,
    handleDblclickBody,
    openInsertHeaderFooter,
    handleHfBeforeInput,
    handleHfKeyDown,
    handleHfCompositionStart,
    handleHfCompositionEnd,
    openHeaderFooterOptions,
    handleHeaderFooterOptionsApply,
    handleHeaderFooterOptionsCancel,
    openPageNumberDialog,
    handlePageNumberInsert,
    handlePageNumberCancel,
  };
}
