<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import AutocorrectMenu from "../components/AutocorrectMenu.vue";
import BookmarkDialog from "../components/BookmarkDialog.vue";
import BookmarkManagerDialog from "../components/BookmarkManagerDialog.vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import ConversionReportDialog from "../components/ConversionReportDialog.vue";
import CustomSpacingDialog from "../components/CustomSpacingDialog.vue";
import DefineListDialog from "../components/DefineListDialog.vue";
import DocImageBlock from "../components/DocImageBlock.vue";
import DocPage from "../components/DocPage.vue";
import DocTableBlock from "../components/DocTableBlock.vue";
import DocumentRuler from "../components/DocumentRuler.vue";
import DownloadMenu from "../components/DownloadMenu.vue";
import EditorContextMenu from "../components/EditorContextMenu.vue";
import HeaderFooterMenu from "../components/HeaderFooterMenu.vue";
import HeaderFooterOptionsDialog from "../components/HeaderFooterOptionsDialog.vue";
import ImageInsertControl from "../components/ImageInsertControl.vue";
import ImageOptionsDialog from "../components/ImageOptionsDialog.vue";
import ImageToolbar from "../components/ImageToolbar.vue";
import InsertBreakMenu from "../components/InsertBreakMenu.vue";
import InsertTableDialog from "../components/InsertTableDialog.vue";
import ColumnsDialog from "../components/ColumnsDialog.vue";
import SpecialCharacterDialog from "../components/SpecialCharacterDialog.vue";
import LinkChip from "../components/LinkChip.vue";
import LinkDialog from "../components/LinkDialog.vue";
import TableInsertControl from "../components/TableInsertControl.vue";
import FindBar from "../components/FindBar.vue";
import FontFamilySelect from "../components/FontFamilySelect.vue";
import FontSizeControl from "../components/FontSizeControl.vue";
import FormatToolbar from "../components/FormatToolbar.vue";
import HighlightPicker from "../components/HighlightPicker.vue";
import MultilevelListMenu from "../components/MultilevelListMenu.vue";
import NotesMenu from "../components/NotesMenu.vue";
import NoteSettingsDialog from "../components/NoteSettingsDialog.vue";
import PageNumberDialog from "../components/PageNumberDialog.vue";
import PageSetupDialog from "../components/PageSetupDialog.vue";
import ParagraphToolbar from "../components/ParagraphToolbar.vue";
import PasteOptionsChip from "../components/PasteOptionsChip.vue";
import SplitCellDialog from "../components/SplitCellDialog.vue";
import StylePicker from "../components/StylePicker.vue";
import TablePropertiesDialog from "../components/TablePropertiesDialog.vue";
import TabStopsDialog from "../components/TabStopsDialog.vue";
import TextColorPicker from "../components/TextColorPicker.vue";
import UndoRedoControls from "../components/UndoRedoControls.vue";
import { useAutosave } from "../composables/useAutosave";
import { usePagination } from "../composables/usePagination";
import { useToasts } from "../composables/useToasts";
import { documentExists } from "../lib/documents";
import { DOCX_MIME_TYPE, docxFileName, writeDocx } from "../lib/docx/writer";
import { ODT_MIMETYPE, odtFileName, writeOdt } from "../lib/odf/writer";
import { noConversionReport } from "../lib/opaque";
import {
  deleteDocument,
  getOrCreateDocument,
  replaceDocumentModel,
  touchDocument,
} from "../lib/documentStore";
import {
  asBorderLineStyle,
  asCellVerticalAlign,
  asCharacterFormatKey,
  asHeaderFooterPart,
  asImage,
  asImageWrapMode,
  asListLevelDefinitions,
  asListMarkerStyle,
  asListType,
  asMeasurementUnit,
  asNoteKind,
  asNoteNumberFormat,
  asNoteRestart,
  asTabStopAlignment,
  asTabStops,
  asPageNumberFormat,
  asParagraphAlignment,
  asParagraphBlock,
  asBreakKind,
  asSectionStart,
  asTableAlignment,
  asTableWidthMode,
  createDocument,
  createEmptyHeaderFooterContent,
  DEFAULT_TITLE,
  isBreak,
  isImage,
  isTable,
  noActiveHeaderFooter,
  noActiveNote,
  noBlocks,
  noNoteContent,
  resolveColumnWidthsCm,
  sectionColumns,
  sectionContentWidthCm,
  sectionMargins,
  sectionPageSize,
} from "../lib/documentModel";
import { detectListAutoformat, listMarkers } from "../lib/lists";
import {
  anchorCaret,
  applyLink,
  autolinkAt,
  deleteBookmark,
  insertBookmark,
  insertLinkText,
  linkAtCaret,
  linkTargetLabel,
  listBookmarks,
  listHeadings,
  noLinkExtent,
  noLinkTarget,
  removeLink,
  sanitizeHref,
  validateBookmarkName,
} from "../lib/links";
import {
  appendTableRowForTab,
  applyBorderPreset,
  arrowEscapeCaret,
  asBorderPreset,
  caretTableLocation,
  cellRangeBounds,
  deleteTable,
  deleteTableColumn,
  deleteTableRow,
  fullTableRange,
  insertTable,
  insertTableColumn,
  insertTableRow,
  mergeCells,
  MIN_COLUMN_WIDTH_CM,
  MIN_ROW_HEIGHT_CM,
  nextCellCaret,
  noCellRange,
  prevCellCaret,
  setCellPadding,
  setCellShading,
  setCellVerticalAlign,
  setColumnWidths,
  setRowMinHeight,
  setTableAlignment,
  setTableWidthMode,
  splitCell,
  tableAt,
} from "../lib/tables";
import {
  applyColumns,
  asColumnsScope,
  setColumnBoundary,
} from "../lib/columns";
import { documentFontFamilies } from "../lib/fonts";
import { loadRecentColors, noHighlight, pushRecentColor } from "../lib/colors";
import { acquireEditLock, noEditLock } from "../lib/editLock";
import {
  BLOCK_SELECTOR,
  canvasModelText,
  caretFromDomPoint,
  domPointForCaret,
  extendDomSelection,
  modelRangeFromSelection,
  modelTextLengthOf,
  noElement,
  noEventTarget,
  noParagraphTarget,
  noTableCellTarget,
  paragraphElementForCaret,
  paragraphFromEventTarget,
  setDomCaret,
  setDomSelection,
  tableCellFromEventTarget,
} from "../lib/domCaret";
import {
  compileQuery,
  findMatches,
  indexAtOrAfter,
  isQueryError,
  matchSelection,
  noHighlightRects,
  noMatches,
  replaceMatch,
  replaceMatches,
} from "../lib/findReplace";
import {
  caretAtClientPoint,
  caretClientRect,
  caretNearClientPoint,
  caretOnBoundaryLine,
  measureImageGeometry,
  measureTableHandles,
  noImageHandleGeometry,
  noTableHandleGeometry,
} from "../lib/caretGeometry";
import {
  CM_PER_PX,
  findTableElement,
  noColumnResizeDrag,
  noRowResizeDrag,
  previewColumnWidths,
  previewRowHeight,
  tableRowElement,
} from "../lib/tableResize";
import {
  blockLength,
  blockText,
  compareCaret,
  noCaret,
  paragraphAt,
  paragraphAtCaret,
  splitRuns,
} from "../lib/textEditing";
import {
  dashSubstitution,
  decadeSubstitution,
  isDigit,
  loadAutocorrectSettings,
  saveAutocorrectSettings,
  smartQuote,
} from "../lib/autocorrect";
import {
  asPasteMode,
  clipboardFlavours,
  DOC_MODEL_MIME,
  executePaste,
  extractFragment,
  fragmentText,
  IMAGES_UNSUPPORTED_MESSAGE,
  noClipboardSnapshot,
  noPastePayload,
  parseFragmentPayload,
  serializeFragment,
} from "../lib/clipboard";
import { importHtml } from "../lib/htmlPaste";
import {
  adjustIndent,
  applyFormat,
  applyListAutoformat,
  applyListDefinition,
  applyMultilevelPreset,
  applyPainterFormat,
  applyParagraphFormat,
  applyParagraphStyle,
  asPageSetupScope,
  continueListNumbering,
  blockCount,
  blockOrdinal,
  capturePainterFormat,
  caretSelection,
  changeListLevel,
  clearFormatting,
  commitHeaderFooterModel,
  commitNoteModel,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentEnd,
  documentStart,
  emptyParagraphFormatting,
  emptyParagraphPaginationState,
  emptyToolbarFormatValues,
  emptyToolbarListState,
  emptyToolbarStyleState,
  firstTouchedFormatting,
  formatState,
  hasSpaceBefore,
  INDENT_STEP_CM,
  insertBreak,
  insertLineBreak,
  insertNote,
  insertPageNumberCommand,
  insertSectionBreak,
  insertText,
  isCollapsed,
  linkHeadingNumbering,
  noCommandResult,
  noPainterFormat,
  noParagraphCommand,
  noPendingFormat,
  noSelection,
  paragraphPaginationState,
  removeListFormat,
  resetParagraphStyle,
  restartListNumbering,
  selectAll,
  selectionListDefinitionLevels,
  selectionTabStops,
  setAlignment,
  setHeaderFooterOptions,
  setListMarker,
  setListNumberingValue,
  setNoteSettings,
  setPageMargins,
  setPageSetup,
  setTabStops,
  splitBlock,
  toggleFormat,
  toggleList,
  toggleSpaceBefore,
  toggleVerticalAlign,
  toolbarFormatValues,
  toolbarListState,
  toolbarStyleState,
  touchedBlocks,
  updateStyleFromSelection,
  verticalAlignState,
  withHeaderFooterModel,
  withNoteModel,
} from "../lib/commands";
import {
  collectNoteRefs,
  createNoteId,
  effectiveNoteSettings,
  noteById,
  noteLabels,
  noteText,
} from "../lib/notes";
import { headingTag, isExternalLink, paragraphStyle, runKey, runLinkHref, runStyle } from "../lib/renderFormatting";
import { markSegments } from "../lib/marks";
import { noIdleHandle, scheduleIdleWork } from "../lib/idleScheduler";
import { PX_PER_CM, sectionColumnFlowWidthPx, wholePageItem } from "../lib/pagination";
import { noNoteArea, noNoteAreaItems } from "../lib/notesLayout";
import {
  computePageNumbers,
  fieldValueText,
  headerFooterForPage,
  resolveHeaderFooterKind,
} from "../lib/pageNumbering";
import { noRenderEntries } from "../lib/pageRender";
import {
  deleteImage,
  imageAt,
  insertImage,
  isAcceptedImageType,
  mediaDataUri,
  noImageCommand,
  noImageLocation,
  pxToCm,
  setImageAltText,
  setImageSize,
  setImageWrap,
} from "../lib/images";
import { prepareImageInsert, ImageInsertError } from "../lib/imageInsert";
import {
  asImageHandle,
  IMAGE_HANDLES,
  noImageResizeDrag,
  previewImageSize,
  resizedImageSize,
} from "../lib/imageResize";
import {
  headingStyleId,
  NORMAL_STYLE_ID,
  noStyleEntries,
  resolveNamedStyle,
  stylePickerEntries,
} from "../lib/styles";
import {
  breakGroup as breakTypingGroup,
  canRedo as canRedoDocument,
  canUndo as canUndoDocument,
  clearHistory as clearUndoHistory,
  noHistorySnapshot,
  recordEdit as recordHistoryEdit,
  redo as redoDocument,
  redoLabel as redoDocumentLabel,
  undo as undoDocument,
  undoLabel as undoDocumentLabel,
} from "../lib/undoHistory";
import { findShortcut, isOutsideDocumentInput, shortcutLabel } from "../lib/shortcuts";
import { isMacPlatform } from "../lib/platform";

const route = useRoute();
const router = useRouter();
const { addToast } = useToasts();
const documentId = computed(() => String(route.params.id ?? ""));

// A SHALLOW ref, deliberately: the model must stay a plain object because
// IndexedDB's structured clone rejects Vue's reactive proxies outright
// (DataCloneError), and autosave persists this exact object. Mutations
// therefore happen on the raw model and re-render via triggerRef — which is
// also what keeps 200-page documents cheap: no proxy per block/run. The
// initial value is computed inline (rather than starting from `null`) so TS
// infers the `DocumentModel | null` type from this expression alone —
// a generic would need TS syntax, which ESLint's espree parser for .vue
// scripts can't parse.
const doc = shallowRef(
  documentExists(documentId.value)
    ? getOrCreateDocument(documentId.value)
    : null,
);

// --- edit lock (SPEC-010: two tabs on the same document) -------------------
// "pending" until the Web Lock resolves; "busy" renders the conflict bar and
// a read-only canvas until the other tab closes and the queued lock grants.
const lockState = ref("pending");
const editable = computed(() => lockState.value === "acquired" && !!doc.value);

// --- caret and scroll (SPEC-010: restored on reopen) -----------------------
// Plain variables, not refs: nothing renders from them, they are only read
// when a draft or commit is written. noCaret() seeds the CaretPosition|null
// type, which espree bars this file from annotating directly.
let caret = noCaret();
let scrollTop = 0;
let scrollContainer = noElement();
let editLock = noEditLock();

// --- undo/redo (SPEC-015) ----------------------------------------------
// The undo/redo history itself is owned by the framework-free lib/undoHistory
// module, keyed by document id (AC-18); this view only feeds it edits and
// caret moves and mirrors its per-document state into a ref for the toolbar
// (AC-5..8). `sessionDocId` — not documentId.value — is what closeSession
// clears, the same reason useAutosave tracks its own docId closure variable:
// by the time the documentId watcher's callback runs, documentId.value
// already reflects the NEW id.
let sessionDocId = "";
const EMPTY_UNDO_REDO_STATE = { canUndo: false, undoLabel: "", canRedo: false, redoLabel: "" };
const undoRedoState = ref(EMPTY_UNDO_REDO_STATE);

function refreshUndoRedoState() {
  const id = documentId.value;
  undoRedoState.value = {
    canUndo: canUndoDocument(id),
    undoLabel: undoDocumentLabel(id) ?? "",
    canRedo: canRedoDocument(id),
    redoLabel: redoDocumentLabel(id) ?? "",
  };
}

// --- clipboard (SPEC-016) ---------------------------------------------------
// `plainPasteNext` is armed by the Ctrl/Cmd+Shift+V keydown (AC-27): where
// the browser fires a `paste` event for that combination the flag routes it
// down the plain-text path, and where it does not (macOS — see
// schedulePlainPasteFallback) the flag survives long enough for the fallback
// to notice. `pastePayload`/`pastePendingFormat`
// describe the last applied paste so the options chip can re-interpret it;
// `lastCopy` is the local snapshot backing the context menu's async paste
// (see ClipboardSnapshot in lib/clipboard.ts). Plain variables — nothing
// renders from them; the chip and menu render from the two refs below.
let plainPasteNext = false;
let pastePayload = noPastePayload();
let pastePendingFormat = noPendingFormat();
let lastCopy = noClipboardSnapshot();
const pasteChip = ref({ visible: false, left: 0, top: 0, applied: "source" });
const contextMenu = ref({ visible: false, left: 0, top: 0 });

// --- tables (SPEC-022) ------------------------------------------------------
// The rectangular cell-range selection (AC-15, AC-16): activated by
// dragging from one cell into another of the same table; the DOM text
// selection is suppressed while it is active. `cellDragOrigin` remembers
// the cell the primary button went down in until the button is released.
const cellSelection = ref(noCellRange());
let cellDragOrigin = noTableCellTarget();
// The cell a right-click opened the context menu in — the target of the
// menu's table actions (AC-17..19).
let contextMenuTableCell = noTableCellTarget();
// The list paragraph a right-click opened the context menu on — the target
// of the menu's numbering-control actions (SPEC-032 AC-14..16). Like
// contextMenuTableCell, set before the menu's visible ref flips so the
// template's render reads the fresh value.
let contextMenuListTarget = noParagraphTarget();
// The hover handles' measured geometry (AC-20, AC-21): page-relative
// rectangles for the hovered table, its columns and its rows. Null hides
// the overlay.
const tableHover = ref(noTableHandleGeometry());
// The "Insert table…" dialog for sizes beyond the 10×10 picker (AC-2).
const insertTableDialogOpen = ref(false);

// --- special characters and autocorrect (SPEC-026) --------------------------
// The "Insert → Special characters" picker dialog.
const specialCharacterDialogOpen = ref(false);
// Smart-quote/dash autocorrect settings, persisted in localStorage.
const autocorrectSettings = ref(loadAutocorrectSettings());
// Tracking state for the follow-up substitutions: the caret offset just after
// an en dash / opening single quote THIS feature produced (or -1 for none),
// and a caret offset at which a just-undone substitution must not re-fire
// until the caret moves (SPEC-026 undo behaviour). These are plain lets, not
// refs — they never drive rendering, only the next keystroke's decision.
let producedEnDashAt = -1;
let producedOpenSingleAt = -1;
let autocorrectSuppressedAt = -1;

// --- table formatting (SPEC-023) --------------------------------------------
// "Split cell…" and the table properties dialog both act on the cell/table
// the context menu was opened against — captured into these refs when the
// dialog opens, since the menu itself (and the click that opened it) is
// long closed by the time the dialog's OK button fires.
const splitCellDialogOpen = ref(false);
const tablePropertiesOpen = ref(false);
const tablePropertiesInitial = ref({});
let tableFormattingLocation = { section: 0, block: 0, row: 0, col: 0 };

// Column/row resize (AC-1..5): live-resizes the hovered table's DOM directly
// while dragging (never the model — AC-3), then commits the final widths/
// height on release. Null when no drag is active.
let columnResizeDrag = noColumnResizeDrag();
let rowResizeDrag = noRowResizeDrag();

// --- images (SPEC-024) -------------------------------------------------------
// The selected image (AC-7): set by clicking one, cleared by any edit
// (applyEdit) or a mousedown elsewhere on the canvas — mirrors cellSelection.
const imageSelection = ref(noImageLocation());
// The selected image's measured geometry (AC-8, AC-9): page-relative
// rectangle for the resize handles, recomputed on select/scroll/resize. Null
// hides the handles/floating toolbar.
const imageHandles = ref(noImageHandleGeometry());
const imageOptionsOpen = ref(false);
const imageOptionsInitial = ref({});
let imageResizeDrag = noImageResizeDrag();
let imageDragMoved = false;

function clearImageSelection() {
  if (imageSelection.value) {
    imageSelection.value = null;
    imageHandles.value = null;
  }
}

// --- headers, footers and page numbers (SPEC-029) ---------------------------
// The section/part/kind currently being edited, or null while editing the
// body (AC-1, AC-2, AC-7, AC-8) — set by double-clicking a header/footer
// band or "Insert -> Header/Footer", cleared by Escape or double-clicking
// the body. Exactly one page's band ever renders as the live contenteditable
// surface at a time (see hfRootEl); every other page of the same section
// renders the identical content read-only, which is how "edits apply to
// every page in the section at once" (AC-10) falls out for free — they all
// read the same model content, not a per-page copy.
const activeHeaderFooter = ref(noActiveHeaderFooter());
// Caret/selection inside the active header/footer's own content, addressed
// in the same synthetic single-section coordinates withHeaderFooterModel
// builds (section 0) — plain variables like the body's own caret/
// selectionAnchor, since nothing renders from them directly.
let hfCaret = noCaret();
let hfSelectionAnchor = noCaret();
// A collapsed hf caret's not-yet-typed format override, mirroring the
// body's own pendingFormat (SPEC-012/013).
let hfPendingFormat = noPendingFormat();
let hfComposing = false;
let hfCompositionRange = noSelection();
// Forces the header/footer band's entries to remount after a composition
// commit (mirrors the body's bumpBlockKey/remount, scoped down since a
// header/footer's content is small enough to always remount whole rather
// than track individual block addresses).
const hfRemountEpoch = ref(0);
const headerFooterOptionsOpen = ref(false);
const headerFooterOptionsInitial = ref({});
const pageNumberDialogOpen = ref(false);

/** The one DOM element currently rendering as an editable header/footer
 * band, if any — set by DocPage's `data-hf-editable` attribute on whichever
 * page's band `activeHeaderFooter` addresses (see DocPage.vue's
 * headerEditable/footerEditable props). A plain query instead of a Vue
 * template ref because the band lives inside one of many `v-for`-repeated
 * DocPage instances — at most one exists in the DOM at a time. */
function hfRootEl() {
  const found = contentRoot.value?.querySelector("[data-hf-editable]");
  return found instanceof HTMLElement ? found : null;
}

/** The header/footer content's live selection, mirroring the body's
 * currentSelection(root) — falls back to the tracked hfCaret/
 * hfSelectionAnchor when the DOM selection is not inside the band (a
 * toolbar control just took focus, SPEC-013's own pattern). */
function currentHfSelection() {
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
function exitHeaderFooterEditing() {
  if (!activeHeaderFooter.value) {
    return;
  }
  activeHeaderFooter.value = null;
  hfCaret = noCaret();
  hfSelectionAnchor = noCaret();
  void nextTick().then(() => contentRoot.value?.focus());
}

/** Enters header/footer editing for the given page's band (AC-1, AC-2):
 * resolves which kind (default/first/even) that page actually renders —
 * the same resolution headerFooterForPage uses for read-only display — so
 * typing lands in the content the user is looking at. */
function enterHeaderFooterEditing(pageIndex = 0, part = asHeaderFooterPart("header")) {
  const model = doc.value;
  if (!model || !editable.value) {
    return;
  }
  const pages = paginationPages.value;
  const page = pages[pageIndex];
  if (!page) {
    return;
  }
  const pageNumber = pageNumbers.value[pageIndex] ?? 1;
  const kind = resolveHeaderFooterKind(model, pages, pageIndex, pageNumber);
  clearCellSelection();
  clearImageSelection();
  exitNoteEditing();
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
  exitNoteEditing();
}

/** Opens "Insert -> Header/Footer" (AC-2): enters editing the caret's own
 * section, defaulting to the header. */
function openInsertHeaderFooter(part = asHeaderFooterPart("header")) {
  const model = doc.value;
  if (!model) {
    return;
  }
  const pages = paginationPages.value;
  const pageIndex = caret ? pagination.pageOfCaret(caret) : 0;
  enterHeaderFooterEditing(pageIndex, part);
  void pages;
}

/** Folds a header/footer command's result (built against the synthetic
 * single-section model, see withHeaderFooterModel) back into the real
 * document and commits it — applyEdit's header/footer sibling. Undo is
 * recorded against a body-caret filler selection (the body caret does not
 * move for a header/footer edit) so Ctrl/Cmd+Z still restores the prior
 * document content; the "currently editing this band" indicator itself is
 * not part of the undo history. */
function applyHeaderFooterEdit(resultModel = noDocument(), resultSelection = noSelection(), label = "") {
  const active = activeHeaderFooter.value;
  const before = doc.value;
  if (!active || !before || !resultSelection) {
    return;
  }
  const nextModel = commitHeaderFooterModel(before, active.sectionIndex, active.part, active.kind, resultModel);
  const bodyFiller = caretSelection(caret ?? documentStart());
  recordHistoryEdit(documentId.value, {
    before: { model: before, selection: bodyFiller },
    after: { model: nextModel, selection: bodyFiller },
    label,
    groupable: false,
  });
  doc.value = replaceDocumentModel(documentId.value, nextModel);
  hfCaret = resultSelection.focus;
  hfSelectionAnchor = isCollapsed(resultSelection) ? noCaret() : resultSelection.anchor;
  autosave.markDirty();
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

function handleHfBeforeInput(event = new InputEvent("beforeinput")) {
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
  let result = null;
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

function handleHfKeyDown(event = new KeyboardEvent("keydown")) {
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

function handleHfCompositionEnd(event = new CompositionEvent("compositionend")) {
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

/** "Header/footer options" (AC-11, AC-12, AC-22, AC-23): prefilled from the
 * section currently being edited, or the caret's section when opened from
 * the menu without an active header/footer. */
function openHeaderFooterOptions() {
  const model = doc.value;
  if (!model) {
    return;
  }
  const sectionIndex = activeHeaderFooter.value?.sectionIndex ?? caretSectionIndex.value;
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
  const root = contentRoot.value;
  if (!model || !root) {
    return;
  }
  const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
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

/** "Insert page number" (AC-14..19): targets the section currently being
 * edited if a header/footer is active, else the caret's own section. */
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
  const root = contentRoot.value;
  if (!model || !root) {
    return;
  }
  const sectionIndex = activeHeaderFooter.value?.sectionIndex ?? caretSectionIndex.value;
  const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
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

// --- footnotes and endnotes (SPEC-031) ---------------------------------------
// The note currently being edited, or null while editing the body — set by
// inserting a note (AC-2, AC-4), clicking a reference mark (AC-15) or
// clicking into a note's body; cleared by Escape or double-clicking the
// body. The same one-live-editable-surface shape as activeHeaderFooter:
// note content is addressed through withNoteModel's synthetic
// single-section coordinates (section 0).
const activeNote = ref(noActiveNote());
let noteCaret = noCaret();
let noteSelectionAnchor = noCaret();
let notePendingFormat = noPendingFormat();
let noteComposing = false;
let noteCompositionRange = noSelection();
const noteRemountEpoch = ref(0);
const noteSettingsOpen = ref(false);
const noteSettingsInitial = ref({});

/** The one DOM element currently rendering as an editable note body, if
 * any (see hfRootEl, the same pattern). */
function noteRootEl() {
  const found = contentRoot.value?.querySelector("[data-note-editable]");
  return found instanceof HTMLElement ? found : null;
}

function currentNoteSelection() {
  const root = noteRootEl();
  if (root) {
    const range = modelRangeFromSelection(root, document.getSelection());
    if (range) {
      return { anchor: range.start, focus: range.end };
    }
  }
  return noteCaret ? { anchor: noteSelectionAnchor ?? noteCaret, focus: noteCaret } : null;
}

/** Leaves note editing and returns focus to the body. */
function exitNoteEditing() {
  if (!activeNote.value) {
    return;
  }
  activeNote.value = null;
  noteCaret = noCaret();
  noteSelectionAnchor = noCaret();
  notePendingFormat = noPendingFormat();
  void nextTick().then(() => contentRoot.value?.focus());
}

/** Enters note editing (AC-2, AC-4, AC-15): the note's area renders as the
 * live editable surface and the caret lands at the given offset of its
 * first block (its end for a click-to-jump, its start for a fresh note —
 * both are the same place for an empty note). */
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

/** Folds a note command's result (built against withNoteModel's synthetic
 * model) back into the real document and commits it — applyEdit's note
 * sibling, mirroring applyHeaderFooterEdit's undo/selection handling. */
function applyNoteEdit(resultModel = noDocument(), resultSelection = noSelection(), label = "") {
  const active = activeNote.value;
  const before = doc.value;
  if (!active || !before || !resultSelection) {
    return;
  }
  const nextModel = commitNoteModel(before, active.noteId, resultModel);
  const bodyFiller = caretSelection(caret ?? documentStart());
  recordHistoryEdit(documentId.value, {
    before: { model: before, selection: bodyFiller },
    after: { model: nextModel, selection: bodyFiller },
    label,
    groupable: false,
  });
  doc.value = replaceDocumentModel(documentId.value, nextModel);
  noteCaret = resultSelection.focus;
  noteSelectionAnchor = isCollapsed(resultSelection) ? noCaret() : resultSelection.anchor;
  autosave.markDirty();
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

function handleNoteBeforeInput(event = new InputEvent("beforeinput")) {
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
  let result = null;
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

function handleNoteKeyDown(event = new KeyboardEvent("keydown")) {
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

function handleNoteCompositionEnd(event = new CompositionEvent("compositionend")) {
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
  exitNoteEditing();
  const target = ref.cell
    ? { section: ref.section, block: ref.block, offset: ref.offset + 1, cell: ref.cell }
    : { section: ref.section, block: ref.block, offset: ref.offset + 1 };
  void nextTick().then(() => {
    contentRoot.value?.focus();
    applyCaretMove(target);
    void scrollModelCaretIntoView(target, true);
  });
}

/** "Insert → Footnote/Endnote" and Ctrl/Cmd+Alt+F / Ctrl/Cmd+Alt+D (AC-1,
 * AC-3): inserts the superscript reference mark at the caret, then moves
 * the caret into the new note's own area (AC-2, AC-4). */
function handleInsertNote(kind = "") {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value || activeHeaderFooter.value || activeNote.value) {
    return;
  }
  const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
  const noteId = createNoteId();
  const noteKind = asNoteKind(kind);
  applyEdit(insertNote(model, selection, noteKind, noteId), selection, `insert ${kind}`);
  enterNoteEditing(noteId, noteKind);
}

/** The footnote settings dialog (AC-17..20), prefilled with both streams'
 * effective settings. */
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
  const root = contentRoot.value;
  if (!model || !root) {
    return;
  }
  const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
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

/** The document's own effective save format (the same rule DownloadMenu's
 * default and handleDownload use) — decides whether an inserted SVG is
 * converted to PNG (AC-2) or kept as-is (AC-3). */
function effectiveFormat() {
  return doc.value && doc.value.format === "odt" ? "odt" : "docx";
}

function remeasureImageHandles() {
  const location = imageSelection.value;
  const pageEl = pageRef.value;
  const root = contentRoot.value;
  if (!location || !pageEl || !root) {
    return;
  }
  const imageEl = root.querySelector(
    `.doc-image-anchor[data-section="${location.section}"][data-block="${location.block}"]`,
  );
  imageHandles.value = imageEl ? measureImageGeometry(pageEl, imageEl) : null;
}

function handleImageSelect(section = 0, block = 0) {
  clearCellSelection();
  imageSelection.value = { section, block };
  void nextTick().then(remeasureImageHandles);
}

const HANDLE_SIZE = 8;

/** One resize handle's page-relative position, centred on its edge/corner
 * of the measured image rect (AC-8). */
function imageHandleStyle(handle = "se") {
  const geometry = imageHandles.value;
  if (!geometry) {
    return {};
  }
  const xFrac = handle.includes("w") ? 0 : handle.includes("e") ? 1 : 0.5;
  const yFrac = handle.includes("n") ? 0 : handle.includes("s") ? 1 : 0.5;
  return {
    left: `${geometry.left + geometry.width * xFrac - HANDLE_SIZE / 2}px`,
    top: `${geometry.top + geometry.height * yFrac - HANDLE_SIZE / 2}px`,
    cursor: `${handle}-resize`,
  };
}

/** The floating image toolbar's fixed (viewport) position (AC-9): above the
 * image when there's room, otherwise just below it. */
function imageToolbarStyle() {
  const geometry = imageHandles.value;
  const pageEl = pageRef.value;
  if (!geometry || !pageEl) {
    return { left: 0, top: 0 };
  }
  const pageRect = pageEl.getBoundingClientRect();
  const viewportTop = pageRect.top + geometry.top;
  const above = viewportTop - 40;
  return {
    left: pageRect.left + geometry.left,
    top: above >= 0 ? above : pageRect.top + geometry.top + geometry.height + 8,
  };
}

/** Runs an image command (setImageWrap/setImageSize/…) against the
 * currently selected image and commits it — the shared apply step every
 * image toolbar/dialog/resize action funnels through. */
function applyImageEdit(mutate = noImageCommand(), label = "") {
  const model = doc.value;
  const root = contentRoot.value;
  const location = imageSelection.value;
  if (!mutate || !model || !root || !location) {
    return;
  }
  const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
  const result = mutate(model, selection, location);
  applyEdit(result, selection, label);
}

function handleImageWrap(wrap = "inline") {
  const location = imageSelection.value;
  if (!location) {
    return;
  }
  applyImageEdit((model, selection) => setImageWrap(model, selection, location, asImageWrapMode(wrap)), "wrap image");
}

function handleImageDelete() {
  const location = imageSelection.value;
  if (!location) {
    return;
  }
  applyImageEdit((model, selection) => deleteImage(model, selection, location), "delete image");
  clearImageSelection();
}

function openImageOptions() {
  const model = doc.value;
  const location = imageSelection.value;
  const image = model && location ? imageAt(model, location) : null;
  if (!image) {
    return;
  }
  imageOptionsInitial.value = {
    altText: image.altText || "",
    widthCm: Number(image.widthCm.toFixed(2)),
    heightCm: Number(image.heightCm.toFixed(2)),
    naturalWidthCm: Number(pxToCm(image.naturalWidthPx).toFixed(2)),
    naturalHeightCm: Number(pxToCm(image.naturalHeightPx).toFixed(2)),
  };
  imageOptionsOpen.value = true;
}

function handleImageOptionsApply(payload = { altText: "", widthCm: 0, heightCm: 0 }) {
  imageOptionsOpen.value = false;
  const location = imageSelection.value;
  const model = doc.value;
  const root = contentRoot.value;
  if (!location || !model || !root) {
    return;
  }
  const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
  let result = setImageAltText(model, selection, location, payload.altText);
  result = setImageSize(result.model, result.selection, location, payload.widthCm, payload.heightCm);
  applyEdit(result, selection, "image options");
  root.focus();
}

/** Reads a File chosen/dropped/pasted and inserts it at the caret (AC-1),
 * showing validation/conversion failures as a toast — never the console
 * (AC-34). */
async function insertImageFromFile(file = new File([], ""), atSelection = noSelection()) {
  const active = activeHeaderFooter.value;
  const model = doc.value;
  const root = contentRoot.value;
  const selection = active
    ? (atSelection ?? currentHfSelection() ?? caretSelection({ section: 0, block: 0, offset: 0 }))
    : (atSelection ?? (root ? currentSelection(root) : null) ?? (model ? caretSelection(caret ?? documentStart()) : null));
  if (!model || !selection) {
    return;
  }
  let prepared;
  try {
    prepared = await prepareImageInsert(file, effectiveFormat());
  } catch (error) {
    if (error instanceof ImageInsertError) {
      addToast(error.message);
      return;
    }
    throw error;
  }
  const current = doc.value;
  if (!current) {
    return;
  }
  if (active) {
    // Header/footer images (SPEC-029 AC-6): inserted into the active band's
    // synthetic model — insertImage's own media-dedup writes into that
    // model's `media`, which commitHeaderFooterModel keeps (see its own
    // doc comment) exactly like any other model-level side effect.
    const synthetic = withHeaderFooterModel(current, active.sectionIndex, active.part, active.kind);
    const result = insertImage(synthetic, selection, prepared.media, prepared.naturalWidthPx, prepared.naturalHeightPx);
    applyHeaderFooterEdit(result.model, result.selection, "insert image");
    return;
  }
  const result = insertImage(current, selection, prepared.media, prepared.naturalWidthPx, prepared.naturalHeightPx);
  applyEdit(result, selection, "insert image");
}

function handleInsertImageFile(file = new File([], "")) {
  void insertImageFromFile(file);
}

// --- image resize handles (SPEC-024 AC-8, AC-10..12) ------------------------

function startImageResize(event = new MouseEvent("mousedown"), handle = "se") {
  const location = imageSelection.value;
  const model = doc.value;
  const root = contentRoot.value;
  if (!location || !model || !root || !editable.value) {
    return;
  }
  const image = imageAt(model, location);
  const imageEl = root.querySelector(
    `.doc-image-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-image`,
  );
  if (!image || !(imageEl instanceof HTMLElement)) {
    return;
  }
  event.preventDefault();
  imageDragMoved = false;
  imageResizeDrag = {
    section: location.section,
    block: location.block,
    handle: asImageHandle(handle),
    startClientX: event.clientX,
    startClientY: event.clientY,
    start: { widthCm: image.widthCm, heightCm: image.heightCm },
    imageEl,
  };
  window.addEventListener("mousemove", handleImageResizeMove);
  window.addEventListener("mouseup", finishImageResize);
}

function handleImageResizeMove(event = new MouseEvent("mousemove")) {
  const drag = imageResizeDrag;
  if (!drag) {
    return;
  }
  imageDragMoved = true;
  const deltaXCm = (event.clientX - drag.startClientX) * CM_PER_PX;
  const deltaYCm = (event.clientY - drag.startClientY) * CM_PER_PX;
  const size = resizedImageSize(drag.start, drag.handle, deltaXCm, deltaYCm, event.shiftKey);
  previewImageSize(drag.imageEl, size.widthCm, size.heightCm);
}

/** Release (AC-10..12): the live-previewed size becomes the one model
 * update this whole drag makes — mirrors finishColumnResize's contract. */
function finishImageResize() {
  window.removeEventListener("mousemove", handleImageResizeMove);
  window.removeEventListener("mouseup", finishImageResize);
  const drag = imageResizeDrag;
  imageResizeDrag = null;
  if (!drag || !imageDragMoved) {
    return;
  }
  const widthCm = parseFloat(drag.imageEl.style.width) || drag.start.widthCm;
  const heightCm = parseFloat(drag.imageEl.style.height) || drag.start.heightCm;
  const location = { section: drag.section, block: drag.block };
  applyImageEdit((model, selection) => setImageSize(model, selection, location, widthCm, heightCm), "resize image");
}

/** Whether the active cell-range selection spans more than one cell — the
 * context menu's "Merge cells" gate (SPEC-023 AC-6). */
const canMergeCells = computed(() => {
  const range = cellSelection.value;
  if (!range) {
    return false;
  }
  const bounds = cellRangeBounds(range);
  return bounds.top !== bounds.bottom || bounds.left !== bounds.right;
});

function clearCellSelection() {
  if (cellSelection.value) {
    cellSelection.value = noCellRange();
  }
}

/** The table cell an event target sits in — see lib/domCaret. */
function tableCellAtTarget(eventTarget = noEventTarget()) {
  const root = contentRoot.value;
  return root ? tableCellFromEventTarget(root, eventTarget) : null;
}

// --- character formatting (SPEC-012) ----------------------------------------
// `selectionAnchor` is set only when the tracked selection is a real (non-
// collapsed) range, so restoreCaretAfterRender knows to re-select rather
// than place a plain caret — needed because setFormat/toggleFormat, unlike
// every other command, do not collapse the selection (AC-13, AC-14).
// `pendingFormat` is the collapsed-caret format set by a toolbar click or
// shortcut with nothing selected (AC-9): it overrides the next inserted
// text's formatting (AC-10) until a real caret move discards it (AC-11).
let selectionAnchor = noCaret();
let pendingFormat = noPendingFormat();

const EMPTY_FORMAT_STATES = {
  bold: "off",
  italic: "off",
  underline: "off",
  strikethrough: "off",
  subscript: "off",
  superscript: "off",
};
const formatStates = ref(EMPTY_FORMAT_STATES);

// --- font/colour/highlight controls (SPEC-013) ------------------------------
// The four value controls' display state (family, size, colour, highlight
// with per-control mixed flags), recomputed alongside formatStates.
const formatValues = ref(emptyToolbarFormatValues());

// --- paragraph formatting (SPEC-014) ----------------------------------------
// What the paragraph toolbar and ruler reflect, all read from the FIRST
// touched paragraph (AC-17's label rule, applied uniformly): the pressed
// alignment button, the space-before toggle label, the direction toggle,
// and the full formatting the ruler positions its markers from.
const EMPTY_PARAGRAPH_STATE = { alignment: "left", spaceBefore: false, direction: "ltr" };
const paragraphState = ref(EMPTY_PARAGRAPH_STATE);
const paragraphFormatting = ref(emptyParagraphFormatting());
// The list buttons' pressed states and the current marker style, from the
// same selection the rest of the toolbar reflects (SPEC-018).
const listState = ref(emptyToolbarListState());
// --- paragraph styles (SPEC-019) --------------------------------------------
// The style picker's current entry (the first touched paragraph's named
// style, AC-1) and its menu entries — the document's own styles, or the
// built-in set for a blank document (AC-2, AC-8). doc.value is replaced on
// every edit, so the list follows catalogue changes (AC-13).
const styleState = ref(emptyToolbarStyleState());
const styleEntries = computed(() => (doc.value ? stylePickerEntries(doc.value) : noStyleEntries()));
// The custom-spacing dialog (AC-12..14): open flag plus the prefill taken
// from the first touched paragraph at the moment it opens.
const customSpacingOpen = ref(false);
const customSpacingInitial = ref({ mode: "multiple", value: 1, beforePt: 0, afterPt: 0 });
// --- multilevel lists and tab stops (SPEC-032) ------------------------------
// The "Define new multilevel list…" dialog: open flag plus the nine level
// definitions it prefills from (the selection's own definition, or the
// decimal preset).
const defineListOpen = ref(false);
const defineListInitial = ref(asListLevelDefinitions([]));
// The Tabs dialog (AC-25): open flag plus the selection's current stops.
const tabsDialogOpen = ref(false);
const tabsDialogInitial = ref(asTabStops([]));
// The Page setup dialog (SPEC-027 AC-1..7): open flag plus the prefill
// taken from the document's own page geometry/unit at the moment it opens.
const pageSetupOpen = ref(false);
const pageSetupInitial = ref({
  pageSize: { widthCm: 21, heightCm: 29.7 },
  margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
  measurementUnit: "cm",
  applyTo: "wholeDocument",
});
// The Format → Columns dialog (SPEC-030): open flag plus the prefill taken
// from the caret's section columns and content width at the moment it opens.
const columnsDialogOpen = ref(false);
const columnsDialogInitial = ref({
  columns: { count: 1, equalWidth: true, gapCm: 1.27, separator: false },
  contentWidthCm: 16,
  applyTo: "thisSection",
});
// Recently applied text colours, shared across documents (SPEC-013 AC).
const recentColors = ref(loadRecentColors());
// The families used by the open document, for the dropdown's "Fonts in
// this document" group — doc.value is replaced on every edit, so this
// recomputes as families come and go.
const documentFonts = computed(() => (doc.value ? documentFontFamilies(doc.value) : []));

// --- format painter (SPEC-020) ----------------------------------------------
// "" (unarmed), "once" (a single click — disarms itself after one applied
// selection, AC-13) or "repeat" (a double-click — stays armed until Escape or
// another button click, AC-14..17). A ref because the canvas's cursor and
// the button's pressed state both render from it; `painterFormat` is the
// copied formatting itself, plain because nothing renders from it directly.
const painterMode = ref("");
let painterFormat = noPainterFormat();
// The selection maybeApplyPainter last acted on (applied to, or saw
// collapsed and skipped) — guards against its own applyEdit's caret-restore
// echo (setDomSelection re-selecting the same range still raises a
// selectionchange) re-triggering the same application over and over.
let lastPainterSelection = noSelection();

// --- hyperlinks and bookmarks (SPEC-025) ------------------------------------
// The link/bookmark dialogs capture the selection they will apply to when
// they open (their own inputs take the focus); the chip tracks the link
// extent the caret currently sits in. Plain vars where nothing renders.
const linkDialogOpen = ref(false);
const linkDialogInitial = ref({ text: "", url: "", anchor: "" });
let linkDialogSelection = noSelection();
const linkChip = ref({ visible: false, left: 0, top: 0, label: "", external: false });
let linkChipExtent = noLinkExtent();
const bookmarkDialogOpen = ref(false);
let bookmarkDialogSelection = noSelection();
const bookmarkManagerOpen = ref(false);

// --- IME composition (SPEC-011) --------------------------------------------
// While the user composes (Japanese/Chinese/Korean input, dead keys) the
// browser owns the DOM: preedit updates are neither cancelable nor modelled.
// The model is committed exactly once, on compositionend, replacing the
// selection snapshotted at compositionstart.
let composing = false;
let compositionRange = noSelection();

// The screen-reader announcement for the current caret/selection (SPEC-011
// AC: "announces caret movement") — rendered into a visually hidden
// aria-live region below the canvas.
const caretStatus = ref("");

// Forces paragraph re-mounts by folding into the v-for keys: `key` names the
// single block whose key carries `count` (bumped after an IME commit so the
// browser-composed DOM nodes are discarded for a fresh model render), and
// `epoch` re-keys the whole canvas (the recovery hatch when a composition
// left the DOM diverged from the model).
const remount = ref({ key: "", count: 0, epoch: 0 });

const contentRoot = useTemplateRef("contentRoot");

const autosave = useAutosave({
  getCaret: () => caret,
  getScroll: () => scrollTop,
  isActive: () => editable.value,
  applyRestored: (restored) => {
    doc.value = restored.model;
    caret = restored.caret;
    scrollTop = restored.scroll;
    void focusAndRestore();
  },
});
const { saveState, saveStateText, recovery, keepRecovered, discardRecovered } = autosave;

const recoveryTime = computed(() =>
  recovery.value ? new Date(recovery.value.savedAt).toLocaleTimeString() : "",
);

async function focusAndRestore() {
  await nextTick();
  attachScrollListener();
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  root.focus();
  caret = caret ?? { section: 0, block: 0, offset: 0 };
  setDomCaret(root, caret);
  updateCaretStatus({ start: caret, end: caret });
  // After focus/caret placement, which may themselves scroll.
  if (scrollContainer) {
    scrollContainer.scrollTop = scrollTop;
  }
}

// --- session lifecycle -----------------------------------------------------

function openSession(id = "") {
  lockState.value = "pending";
  sessionDocId = id;
  editLock = acquireEditLock(id, (state) => {
    if (documentId.value !== id || !doc.value) {
      return;
    }
    lockState.value = state === "acquired" ? "acquired" : "busy";
    if (state === "acquired") {
      void autosave.beginSession(id);
      refreshUndoRedoState();
    }
  });
}

function closeSession() {
  // An in-app close (navigation, unmount) is a clean close: commit, unless
  // this tab never owned the document.
  autosave.endSession(lockState.value === "acquired");
  if (editLock) {
    editLock.release();
    editLock = null;
  }
  lockState.value = "pending";
  caret = noCaret();
  scrollTop = 0;
  // The next document opens scrolled to the top: realize only its head
  // until the viewport refresh sees the real scroll position.
  viewportPageRange.value = { start: 0, end: 2 };
  anchorPage.value = -1;
  caretPage.value = 1;
  dismissPasteChip();
  closeContextMenu();
  resetFindState();
  clearCellSelection();
  cellDragOrigin = null;
  tableHover.value = null;
  insertTableDialogOpen.value = false;
  specialCharacterDialogOpen.value = false;
  producedEnDashAt = -1;
  producedOpenSingleAt = -1;
  plainPasteNext = false;
  // AC-19: the undo/redo history does not outlive the document being open.
  if (sessionDocId) {
    clearUndoHistory(sessionDocId);
    sessionDocId = "";
  }
  undoRedoState.value = EMPTY_UNDO_REDO_STATE;
}

watch(documentId, (id) => {
  closeSession();
  doc.value = documentExists(id) ? getOrCreateDocument(id) : null;
  if (doc.value) {
    openSession(id);
  }
});

function attachScrollListener() {
  const root = contentRoot.value;
  const container = root ? root.closest("main") : null;
  if (container && container !== scrollContainer) {
    if (scrollContainer) {
      scrollContainer.removeEventListener("scroll", handleScroll);
    }
    scrollContainer = container;
    scrollContainer.addEventListener("scroll", handleScroll, { passive: true });
  }
}

function handleScroll() {
  if (scrollContainer) {
    scrollTop = scrollContainer.scrollTop;
  }
  scheduleViewportRefresh();
  // Re-measure the (viewport-limited) highlight overlay for the new scroll
  // position; rAF-throttled inside scheduleFindMeasure (SPEC-017).
  if (findVisible.value) {
    scheduleFindMeasure();
  }
}

function handleSelectionChange() {
  const root = contentRoot.value;
  // During composition the DOM temporarily contains preedit text, so model
  // offsets computed from it would be garbage — the caret is re-derived at
  // compositionend instead.
  if (!root || !editable.value || composing) {
    return;
  }
  const range = modelRangeFromSelection(root, document.getSelection());
  if (range) {
    // A genuine external move (arrow keys, a native click/drag inside the
    // canvas) discards the pending format (AC-11 of SPEC-012) and breaks the
    // undo history's typing group (SPEC-015 AC-10). Our own programmatic
    // placements (restoreCaretAfterRender after an edit) don't count: they
    // set `caret` to the new position *before* touching the DOM, so by the
    // time this fires the observed range already matches it.
    const genuineMove =
      !caret || compareCaret(range.end, caret) !== 0 || compareCaret(range.start, range.end) !== 0;
    if (pendingFormat && genuineMove) {
      pendingFormat = noPendingFormat();
    }
    if (genuineMove) {
      breakTypingGroup(documentId.value);
    }
    caret = range.end;
    // Track the anchor of a real selection so a toolbar interaction that
    // moves focus into its own input (the size field, the custom colour
    // field) can still apply to the range the user selected (SPEC-013).
    selectionAnchor =
      compareCaret(range.start, range.end) === 0 ? noCaret() : range.start;
    updateCaretStatus(range);
    void refreshLinkChip();
    // The format painter's target (SPEC-020 AC-12): skipped while a text
    // drag is in progress (handleCanvasMouseUp applies on release instead,
    // once the range has settled) — see maybeApplyPainter's own guard
    // against re-triggering off its own caret-restore echo.
    if (!mouseButtonDown) {
      maybeApplyPainter();
    }
  }
}

onMounted(() => {
  document.addEventListener("selectionchange", handleSelectionChange);
  document.addEventListener("keydown", handleGlobalShortcut);
  attachScrollListener();
  refreshViewportRange();
  if (doc.value) {
    openSession(documentId.value);
  }
});

onBeforeUnmount(() => {
  closeSession();
  cancelFindScheduled();
  document.removeEventListener("selectionchange", handleSelectionChange);
  document.removeEventListener("keydown", handleGlobalShortcut);
  window.removeEventListener("mousemove", handleColumnResizeMove);
  window.removeEventListener("mouseup", finishColumnResize);
  window.removeEventListener("mousemove", handleRowResizeMove);
  window.removeEventListener("mouseup", finishRowResize);
  if (scrollContainer) {
    scrollContainer.removeEventListener("scroll", handleScroll);
    scrollContainer = null;
  }
});

// --- the editing engine (SPEC-011) -----------------------------------------
// Fully controlled: every beforeinput is cancelled and the supported ones
// re-applied through the command API in lib/commands, which returns a new
// model state that replaces doc.value (and the store's model, so autosave
// persists it). The DOM re-renders from the model and the caret is
// re-placed — the model never drifts from the DOM, the DOM is never the
// source of truth, and unsupported input types are simply inert. The one
// exception is IME composition (see the composition handlers below).

/** The current selection in model coordinates: the live DOM selection when
 * it is inside the canvas, else the tracked selection/caret — the tracked
 * anchor covers toolbar interactions whose own inputs took the focus
 * (SPEC-013's size and custom-colour fields). */
function currentSelection(root = noElement()) {
  if (!root) {
    return null;
  }
  const range = modelRangeFromSelection(root, document.getSelection());
  if (range) {
    return { anchor: range.start, focus: range.end };
  }
  return caret ? { anchor: selectionAnchor ?? caret, focus: caret } : null;
}

/** Applies a command result: the new model becomes the document (in the
 * view and in the store, so autosave and downloads see it), and the caret
 * moves to the command's post-edit position — or, for the character-
 * formatting commands, the whole (non-collapsed) selection is restored
 * (SPEC-012 AC-13, AC-14).
 *
 * `beforeSelection` is the selection the command was given — exactly the
 * state undo must restore the caret/selection to (SPEC-015 AC-14) — and
 * `label`/`groupable` describe the edit for the undo history (AC-7..13):
 * every call site names its own action, so recording is never skipped by
 * accident the way a defaulted label could be. */
function applyEdit(
  result = noCommandResult(),
  beforeSelection = noSelection(),
  label = "",
  groupable = false,
  keepPasteChip = false,
) {
  if (!result || !beforeSelection) {
    return;
  }
  // Any edit other than the paste (or its chip re-application, which passes
  // keepPasteChip) supersedes the paste options chip (SPEC-016).
  if (!keepPasteChip) {
    dismissPasteChip();
  }
  // An edit invalidates the cell-range highlight and the hover handles'
  // measured geometry (SPEC-022). The selected image (SPEC-024) is
  // re-selected below if the same location still holds an image after the
  // edit — true for the image's own toolbar/resize/options actions, so the
  // toolbar stays open across them; false for an edit elsewhere (typing,
  // deleting the image, …), which deselects as SPEC-022's table handles do.
  clearCellSelection();
  tableHover.value = null;
  const priorImageSelection = imageSelection.value;
  clearImageSelection();
  const before = doc.value;
  if (before) {
    recordHistoryEdit(documentId.value, {
      before: { model: before, selection: beforeSelection },
      after: { model: result.model, selection: result.selection },
      label,
      groupable,
    });
  }
  doc.value = replaceDocumentModel(documentId.value, result.model);
  caret = result.selection.focus;
  selectionAnchor = isCollapsed(result.selection) ? noCaret() : result.selection.anchor;
  autosave.markDirty();
  updateCaretStatus({ start: result.selection.anchor, end: result.selection.focus });
  refreshUndoRedoState();
  void restoreCaretAfterRender();
  void keepCaretVisible();
  void refreshLinkChip();
  if (
    priorImageSelection &&
    isImage(doc.value.sections[priorImageSelection.section]?.blocks[priorImageSelection.block])
  ) {
    imageSelection.value = priorImageSelection;
    void nextTick().then(remeasureImageHandles);
  }
  // An edit made with the find bar open shifts match positions; refresh off
  // the keystroke path so typing stays responsive (SPEC-017 AC-31).
  scheduleFindRefresh();
}

/** Ctrl/Cmd+Z / the toolbar's Undo button (AC-1, AC-4). */
function performUndo() {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  applyHistorySnapshot(undoDocument(documentId.value));
}

/** Ctrl/Cmd+Shift+Z, Ctrl+Y on Windows / the toolbar's Redo button (AC-2,
 * AC-3, AC-4). */
function performRedo() {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  applyHistorySnapshot(redoDocument(documentId.value));
}

/** Restores a snapshot from the undo/redo history: unlike applyEdit, no new
 * edit is recorded (undo/redo isn't itself an undoable edit) and any
 * pending format is discarded, since the selection is moving. */
function applyHistorySnapshot(snapshot = noHistorySnapshot()) {
  const root = contentRoot.value;
  if (!snapshot || !root) {
    return;
  }
  dismissPasteChip();
  doc.value = replaceDocumentModel(documentId.value, snapshot.model);
  caret = snapshot.selection.focus;
  selectionAnchor = isCollapsed(snapshot.selection) ? noCaret() : snapshot.selection.anchor;
  pendingFormat = noPendingFormat();
  // Undo/redo moves the caret and rewrites text, so a pending substitution
  // follow-up context (en dash / opening quote) can no longer apply — clearing
  // it stops a stale trigger firing after an undo (SPEC-026 AC-22).
  producedEnDashAt = -1;
  producedOpenSingleAt = -1;
  autosave.markDirty();
  updateCaretStatus({ start: snapshot.selection.anchor, end: snapshot.selection.focus });
  refreshUndoRedoState();
  root.focus();
  void restoreCaretAfterRender();
  // Undo/redo can move the caret a long way — undoing an inserted page
  // break pulls it back a whole page — so follow it, the way the break
  // insertion it reverses does.
  void scrollModelCaretIntoView(caret);
  scheduleFindRefresh();
}

/** Inserts a literal string at the current caret WITHOUT running autocorrect
 * — the character picker (AC-30) and the non-breaking-space shortcut (AC-11)
 * both go through here, never the typed-input path. */
function insertSpecialCharacter(text = "") {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value || !text) {
    return;
  }
  const selection =
    currentSelection(root) ?? caretSelection(caret ?? documentStart());
  applyEdit(
    insertText(model, selection, text, pendingFormat),
    selection,
    "insert character",
    false,
  );
  // A deliberate insert is not a substitution follow-up context.
  producedEnDashAt = -1;
  producedOpenSingleAt = -1;
}

function openSpecialCharacters() {
  specialCharacterDialogOpen.value = true;
}

function handleSpecialCharacterInsert(text = "") {
  insertSpecialCharacter(text);
}

function toggleAutocorrectSmartQuotes() {
  const next = {
    ...autocorrectSettings.value,
    smartQuotes: !autocorrectSettings.value.smartQuotes,
  };
  autocorrectSettings.value = next;
  saveAutocorrectSettings(next);
}

function toggleAutocorrectDashes() {
  const next = {
    ...autocorrectSettings.value,
    dashes: !autocorrectSettings.value.dashes,
  };
  autocorrectSettings.value = next;
  saveAutocorrectSettings(next);
}

/** The code point immediately before `index` in `text`, as a string (a full
 * surrogate pair when it is a non-BMP character), or null at the start. */
function codePointBefore(text = "", index = 0) {
  if (index <= 0) {
    return null;
  }
  const code = text.charCodeAt(index - 1);
  if (code >= 0xdc00 && code <= 0xdfff && index >= 2) {
    return text.slice(index - 2, index);
  }
  return text[index - 1];
}

/** Replaces [replacement.start, replacement.end) within the caret's paragraph
 * with replacement.text, returning a CommandResult whose caret sits after the
 * inserted text. Used to record a substitution as its own edit. */
function replaceInParagraph(
  model = createDocument(""),
  atCaret = documentStart(),
  replacement = { start: 0, end: 0, text: "" },
) {
  const deleted = deleteRange(model, {
    anchor: { ...atCaret, offset: replacement.start },
    focus: { ...atCaret, offset: replacement.end },
  });
  return insertText(deleted.model, deleted.selection, replacement.text, null);
}

/** Runs the smart-quote/dash/decade substitutions after a single character was
 * plain-inserted (SPEC-026). `plainResult` is that plain insert. Returns true
 * when it recorded its own undo entries (the caller must then stop); false to
 * let the caller apply the plain insert normally. Only ever reached from the
 * typed-input branch below (AC-27..31). */
function applyAutocorrect(
  selection = noSelection(),
  typed = "",
  plainResult = noCommandResult(),
) {
  if (!selection || !plainResult) {
    return false;
  }
  const settings = autocorrectSettings.value;
  const caretAfter = plainResult.selection.focus;
  const blockAfter = paragraphAtCaret(plainResult.model, caretAfter);
  if (!blockAfter) {
    return false;
  }
  const text = blockText(blockAfter);
  const off = caretAfter.offset;

  // Decade idiom (AC-50): a digit typed right after our opening single quote
  // flips it to a closing quote as part of the digit's OWN edit, so one undo
  // removes the digit and reverts the flip together.
  if (isDigit(typed) && producedOpenSingleAt >= 0) {
    const rep = decadeSubstitution(text, off, typed, producedOpenSingleAt);
    if (rep) {
      const flipped = replaceInParagraph(plainResult.model, caretAfter, rep);
      applyEdit(
        { model: flipped.model, selection: caretSelection(caretAfter) },
        selection,
        "typing",
        true,
      );
      producedOpenSingleAt = -1;
      producedEnDashAt = -1;
      return true;
    }
  }

  // Smart quotes (AC-18, AC-19): plain insert first, then the curly quote as a
  // separate undo entry so one Ctrl/Cmd+Z restores the straight quote (AC-21).
  if (settings.smartQuotes && (typed === "'" || typed === '"')) {
    const quote = smartQuote(typed, codePointBefore(text, off - 1));
    if (quote) {
      applyEdit(plainResult, selection, "typing", true);
      breakTypingGroup(documentId.value);
      const subResult = replaceInParagraph(plainResult.model, caretAfter, {
        start: off - 1,
        end: off,
        text: quote,
      });
      applyEdit(subResult, caretSelection(caretAfter), "smart quote", false);
      producedOpenSingleAt =
        quote === "‘" ? subResult.selection.focus.offset : -1;
      producedEnDashAt = -1;
      return true;
    }
  }

  // Dash substitution (AC-16, AC-17): same two-step undo shape.
  if (settings.dashes && typed === "-") {
    const sub = dashSubstitution(
      text,
      off,
      producedEnDashAt >= 0 ? producedEnDashAt : null,
    );
    if (sub) {
      applyEdit(plainResult, selection, "typing", true);
      breakTypingGroup(documentId.value);
      const subResult = replaceInParagraph(
        plainResult.model,
        caretAfter,
        sub.replacement,
      );
      applyEdit(
        subResult,
        caretSelection(caretAfter),
        sub.producesEnDash ? "en dash" : "em dash",
        false,
      );
      producedEnDashAt = sub.producesEnDash
        ? subResult.selection.focus.offset
        : -1;
      producedOpenSingleAt = -1;
      return true;
    }
  }

  // An ordinary typed character with no substitution invalidates any pending
  // follow-up context.
  producedEnDashAt = -1;
  producedOpenSingleAt = -1;
  return false;
}

function handleBeforeInput(event = new InputEvent("beforeinput")) {
  const type = event.inputType;
  if (composing || event.isComposing || type === "insertCompositionText") {
    // The browser is composing into the DOM; insertCompositionText is not
    // cancelable and the model must not change until compositionend.
    return;
  }
  event.preventDefault();
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  // The paragraph a collapsed caret sits in — what the list-specific Enter
  // and Backspace behaviours below decide on (SPEC-018). Resolved through
  // the caret's cell path when it sits inside a table (SPEC-022).
  const focusBlock = paragraphAtCaret(model, selection.focus);
  let result = null;
  // Undo grouping (SPEC-015 AC-9, AC-12): only literal typed characters
  // group into one step. A paste is bulk content, not "typed characters",
  // so it stays its own step even though it shares insertText's mechanics.
  let label = "";
  let groupable = false;
  if (type === "insertText" || type === "insertFromPaste") {
    const text =
      event.data ??
      (event.dataTransfer ? event.dataTransfer.getData("text/plain") : "");
    if (text) {
      result = insertText(model, selection, text, pendingFormat);
      label = type === "insertText" ? "typing" : "paste";
      groupable = type === "insertText";
      if (type === "insertText" && text === " " && isCollapsed(selection)) {
        // List autoformat (SPEC-018 AC-14..17): the typed space completed a
        // trigger prefix that is the paragraph's entire text. Two separate
        // undo steps — the typing, then the conversion — so Ctrl/Cmd+Z
        // right after restores the literal prefix (AC-16).
        const caretAfter = result.selection.focus;
        const blockAfter = paragraphAtCaret(result.model, caretAfter);
        const prefix = blockAfter ? blockText(blockAfter) : "";
        const autoformatType =
          blockAfter && !blockAfter.list && caretAfter.offset === prefix.length
            ? detectListAutoformat(prefix)
            : null;
        if (autoformatType) {
          applyEdit(result, selection, label, groupable);
          const autoformatSelection = caretSelection(caretAfter);
          applyEdit(
            applyListAutoformat(result.model, autoformatSelection, autoformatType),
            autoformatSelection,
            autoformatType === "bullet" ? "bulleted list" : "numbered list",
            false,
          );
          return;
        }
      }
      // Autocorrect substitutions run on typed input only (SPEC-026 AC-27):
      // a single typed character may become a smart quote, an en/em dash, or
      // trigger the decade-idiom flip. Never on paste (this branch also
      // handles insertFromPaste, hence the type guard), the picker or load.
      if (type === "insertText" && text.length === 1) {
        if (applyAutocorrect(selection, text, result)) {
          return;
        }
      }
    }
  } else if (type === "insertParagraph") {
    if (
      isCollapsed(selection) &&
      focusBlock &&
      focusBlock.list &&
      blockLength(focusBlock) === 0
    ) {
      // Enter on an empty list item (SPEC-018): outdent one level (AC-7),
      // or at level 1 convert it to a normal paragraph (AC-8).
      result =
        focusBlock.list.level > 1
          ? changeListLevel(model, selection, -1)
          : removeListFormat(model, selection);
      label = "list";
    } else {
      result = splitBlock(model, selection);
      label = "new paragraph";
    }
  } else if (type === "insertLineBreak") {
    // Shift+Enter: a forced line break, stored as a literal "\n" that
    // pre-wrap rendering displays without a new paragraph.
    result = insertLineBreak(model, selection);
    label = "line break";
  } else if (type === "deleteContentBackward") {
    if (isCollapsed(selection) && selection.focus.offset === 0 && focusBlock?.list) {
      // Backspace at the start of a list item removes only the list
      // formatting, deleting nothing (SPEC-018 AC-9).
      result = removeListFormat(model, selection);
      label = "list";
    } else {
      result = deleteBackward(model, selection);
      label = "delete";
    }
  } else if (type === "deleteContentForward") {
    result = deleteForward(model, selection);
    label = "delete";
  } else if (type === "deleteByCut" || type === "deleteContent") {
    // The clipboard copy half of a cut happened in the earlier `cut` event;
    // cancelling beforeinput only leaves the deletion to us.
    result = isCollapsed(selection) ? null : deleteRange(model, selection);
    label = "delete";
  }
  if (result) {
    applyEdit(result, selection, label, groupable);
    // Autolink (SPEC-025 AC-5, AC-6): a completed space or Enter may have
    // finished a bare URL just before the split/insertion point.
    if (type === "insertText" && event.data === " " && label === "typing") {
      maybeAutolink(result, false);
    } else if (type === "insertParagraph" && label === "new paragraph") {
      maybeAutolink(result, true);
    }
  }
}

async function restoreCaretAfterRender() {
  await nextTick();
  const root = contentRoot.value;
  if (!root || !caret) {
    return;
  }
  if (!paragraphElementForCaret(root, caret)) {
    // The caret's paragraph landed on a virtualized page (an undo/redo or
    // command can move it far): refresh the page trackers — which realizes
    // that page — and wait for the extra patch before placing.
    refreshCaretPage();
    if (selectionAnchor) {
      anchorPage.value = pagination.pageOfCaret(selectionAnchor);
    }
    await nextTick();
  }
  if (selectionAnchor) {
    setDomSelection(root, selectionAnchor, caret);
  } else {
    setDomCaret(root, caret);
  }
}

// --- IME composition and dead keys (SPEC-011) ------------------------------

function handleCompositionStart() {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  composing = true;
  // Snapshot what the composition will replace BEFORE the browser touches
  // the DOM: the selection at compositionstart, still in model coordinates.
  compositionRange = currentSelection(root);
}

function handleCompositionEnd(event = new CompositionEvent("compositionend")) {
  composing = false;
  const range = compositionRange;
  compositionRange = null;
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value || !range) {
    return;
  }
  const data = event.data ?? "";
  if (!data) {
    // Cancelled composition: the browser removed its own preedit and the
    // model was never touched — nothing to commit.
    return;
  }
  // Re-key the paragraph the composition happened in so the render below
  // mounts it fresh from the model, discarding whatever nodes the browser's
  // composition left in the DOM (Vue cannot patch nodes it did not create).
  bumpBlockKey(range.anchor.section, range.anchor.block);
  applyEdit(insertText(model, range, data, pendingFormat), range, "typing", true);
  void verifyCanvasAfterRender();
}

function bumpBlockKey(section = 0, block = 0) {
  const current = remount.value;
  remount.value = {
    key: `${section}:${block}`,
    count: current.count + 1,
    epoch: current.epoch,
  };
}

/** Belt and braces for composition edge cases: if the canvas text does not
 * match the model after the commit render, re-key every paragraph so the
 * whole canvas re-mounts from the model — the model always wins. */
async function verifyCanvasAfterRender() {
  await nextTick();
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model) {
    return;
  }
  if (virtualizing.value) {
    // Virtualized pages keep their text out of the DOM by design, so the
    // whole-canvas equality check below cannot hold; composition commits
    // already re-key their own paragraph (bumpBlockKey), which is the part
    // of this belt that matters.
    return;
  }
  const expected = model.sections
    .map((section) => section.blocks.map(blockText).join(""))
    .join("");
  // canvasModelText, not textContent: note reference marks render their
  // computed number over a one-character model placeholder, and note areas/
  // header/footer bands hold content outside the body flow (SPEC-031).
  if (canvasModelText(root) !== expected) {
    const current = remount.value;
    remount.value = { key: current.key, count: current.count, epoch: current.epoch + 1 };
    void restoreCaretAfterRender();
  }
}

// --- cut, copy and paste (SPEC-016) ----------------------------------------
// Fully controlled, like typing: the copy/cut/paste ClipboardEvents are
// cancelled and re-implemented over the model, because only these
// synchronous handlers can carry the private document-model MIME type
// (navigator.clipboard.read/write accept a fixed type allowlist). The
// native context menu's entries and Ctrl/Cmd+X/C/V both fire these same
// events; the app's own context menu drives them via execCommand (cut/copy)
// and the async clipboard API (paste).

/** Captures the selection's fragment and writes all three clipboard
 * flavours (AC-7..9): plain text, minimal HTML and the internal payload.
 * Also remembers the snapshot locally for the context menu's paste. */
function writeClipboardFlavours(data = new DataTransfer(), model = doc.value, selection = noSelection()) {
  if (!model || !selection) {
    return;
  }
  const fragment = extractFragment(model, selection);
  const flavours = clipboardFlavours(fragment);
  data.setData("text/plain", flavours.text);
  data.setData("text/html", flavours.html);
  data.setData(DOC_MODEL_MIME, flavours.payload);
  lastCopy = { text: flavours.text, html: flavours.html, fragment };
}

/** Ctrl/Cmd+C and the context menu's Copy (AC-2, AC-5). Copying mutates
 * nothing, so it works in a read-only (conflict-bar) session too. */
function handleCopy(event = new ClipboardEvent("copy")) {
  const root = contentRoot.value;
  const model = doc.value;
  const data = event.clipboardData;
  if (!root || !model || !data) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection || isCollapsed(selection)) {
    return;
  }
  event.preventDefault();
  writeClipboardFlavours(data, model, selection);
}

/** Ctrl/Cmd+X and the context menu's Cut (AC-1, AC-4): the copy flavours
 * plus deleting the selection, recorded as a single "cut" undo step
 * (AC-36..38). */
function handleCut(event = new ClipboardEvent("cut")) {
  const root = contentRoot.value;
  const model = doc.value;
  const data = event.clipboardData;
  if (!root || !model || !data || !editable.value) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection || isCollapsed(selection)) {
    return;
  }
  event.preventDefault();
  writeClipboardFlavours(data, model, selection);
  applyEdit(deleteRange(model, selection), selection, "cut");
}

function clipboardHasImage(data = new DataTransfer()) {
  const items = data.items;
  if (!items) {
    return false;
  }
  for (let i = 0; i < items.length; i++) {
    if (items[i].kind === "file" && items[i].type.startsWith("image/")) {
      return true;
    }
  }
  return false;
}

/** Ctrl/Cmd+V and the native menu's Paste (AC-3, AC-6): reads the flavours
 * off the event and routes them through performPaste. Always prevented —
 * the browser must never splice its own HTML into the canvas. */
/** The first clipboard file item of an accepted image type, or null — read
 * synchronously off the paste event (SPEC-024 AC-1), the one point a raw
 * image File is reachable without the async Clipboard API's permission
 * prompt. */
function clipboardImageFile(data = new DataTransfer()) {
  const items = data.items;
  if (!items) {
    return null;
  }
  for (let i = 0; i < items.length; i++) {
    if (items[i].kind === "file" && isAcceptedImageType(items[i].type) && typeof items[i].getAsFile === "function") {
      return items[i].getAsFile();
    }
  }
  return null;
}

function handlePaste(event = new ClipboardEvent("paste")) {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  event.preventDefault();
  const data = event.clipboardData;
  if (!data) {
    return;
  }
  const plainOnly = plainPasteNext;
  plainPasteNext = false;
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  const imageFile = clipboardImageFile(data);
  if (imageFile) {
    void insertImageFromFile(imageFile, selection);
    return;
  }
  performPaste(
    {
      internal: data.getData(DOC_MODEL_MIME),
      html: data.getData("text/html"),
      text: data.getData("text/plain"),
      hasImage: clipboardHasImage(data),
    },
    selection,
    plainOnly,
  );
}

/** Ctrl/Cmd+Shift+V's fallback path. Chromium only turns that combination
 * into a clipboard command on Windows/Linux, where it is a native paste
 * variant; on macOS it belongs to the application menu and never reaches
 * the page, so the armed flag would sit there and the command would
 * silently do nothing. Once the event has had its chance, read the
 * clipboard's text ourselves — plain text is the only flavour this command
 * wants, and the only one navigator.clipboard exposes. A refused
 * permission or an empty clipboard leaves the document untouched, exactly
 * as before. */
function schedulePlainPasteFallback() {
  window.setTimeout(() => {
    if (!plainPasteNext) {
      // The paste event arrived and consumed the flag.
      return;
    }
    plainPasteNext = false;
    void pastePlainFromClipboard();
  }, 0);
}

async function pastePlainFromClipboard() {
  let text = "";
  try {
    text = await navigator.clipboard.readText();
  } catch {
    return;
  }
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value || text === "") {
    return;
  }
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  performPaste({ internal: "", html: "", text, hasImage: false }, selection, true);
}

/** One paste, whatever triggered it: the internal payload wins (exact
 * editor-to-editor fidelity, AC-10..13), then the sanitised HTML import
 * (AC-14..26), then plain text; `plainOnly` (Ctrl/Cmd+Shift+V, AC-27)
 * forces the text path but keeps any richer fragment so the chip can still
 * offer "Keep source formatting". Image content is gated behind the
 * "Images aren't supported yet" toast until SPEC-024 (AC-45, AC-46). */
function performPaste(
  flavours = { internal: "", html: "", text: "", hasImage: false },
  selection = noSelection(),
  plainOnly = false,
) {
  const model = doc.value;
  if (!model || !selection) {
    return;
  }
  let fragment = flavours.internal ? parseFragmentPayload(flavours.internal) : null;
  let imagesDropped = false;
  if (!fragment && flavours.html) {
    const imported = importHtml(flavours.html);
    imagesDropped = imported.imagesDropped;
    if (imported.blocks.length > 0) {
      fragment = imported.blocks;
    }
  }
  if (imagesDropped || flavours.hasImage) {
    addToast(IMAGES_UNSUPPORTED_MESSAGE);
  }
  if (!fragment && !flavours.text) {
    return;
  }
  const payload = { fragment, text: flavours.text };
  const mode = plainOnly || !fragment ? "text" : "source";
  const result = executePaste(model, selection, payload, mode, pendingFormat);
  pastePayload = payload;
  pastePendingFormat = pendingFormat ? { ...pendingFormat } : noPendingFormat();
  applyEdit(result, selection, "paste", false, true);
  void openPasteChip(mode);
}

/** Shows the chip pinned under the post-paste caret (AC-29). */
async function openPasteChip(applied = "source") {
  await nextTick();
  const root = contentRoot.value;
  const rect = root && caret ? caretClientRect(root, caret) : null;
  pasteChip.value = {
    visible: true,
    left: rect ? Math.max(8, rect.left) : 16,
    top: rect ? rect.bottom + 8 : 16,
    applied,
  };
}

function dismissPasteChip() {
  if (pasteChip.value.visible) {
    pasteChip.value = { visible: false, left: 0, top: 0, applied: "source" };
  }
  pastePayload = noPastePayload();
  pastePendingFormat = noPendingFormat();
}

/** A chip choice (AC-30..32): re-applies the same paste under the chosen
 * interpretation by replacing the paste's undo step — pop it, restore the
 * pre-paste state, re-run the paste from the stored payload — so the paste
 * plus the re-application stay one undo step (AC-33). Choosing the option
 * already applied is a confirmation no-op. */
function handlePasteOption(mode = "") {
  const root = contentRoot.value;
  const payload = pastePayload;
  if (!root || !editable.value || !payload || !pasteChip.value.visible) {
    return;
  }
  const chosen = asPasteMode(mode);
  if (chosen === pasteChip.value.applied) {
    root.focus();
    return;
  }
  const snapshot = undoDocument(documentId.value);
  if (!snapshot) {
    return;
  }
  doc.value = snapshot.model;
  const result = executePaste(
    snapshot.model,
    snapshot.selection,
    payload,
    chosen,
    pastePendingFormat,
  );
  applyEdit(result, snapshot.selection, "paste", false, true);
  void openPasteChip(chosen);
  root.focus();
}

// --- hyperlinks and bookmarks (SPEC-025) -----------------------------------

/** The dialog's heading targets (AC-19, AC-21), re-read from the current
 * model each time the dialog template evaluates them. */
function documentHeadings() {
  const model = doc.value;
  return model
    ? listHeadings(model).map((heading) => ({
        anchor: heading.anchor,
        text: heading.text,
        level: heading.level,
      }))
    : [];
}

/** The dialog's bookmark targets (AC-20, AC-21). */
function documentBookmarkTargets() {
  const model = doc.value;
  return model
    ? listBookmarks(model).map((bookmark) => ({ anchor: bookmark.name, text: bookmark.name }))
    : [];
}

/** Bookmark names in document order — the manager's list (AC-23) and the
 * uniqueness set the insert dialog validates against (AC-27). */
function documentBookmarkNames() {
  const model = doc.value;
  return model ? listBookmarks(model).map((bookmark) => bookmark.name) : [];
}

/** A model selection spanning a link extent — what Edit/Remove act on. */
function extentSelection(extent = noLinkExtent()) {
  if (!extent) {
    return null;
  }
  const address = extent.address;
  const anchor = address.cell
    ? { section: address.section, block: address.block, offset: extent.from, cell: address.cell }
    : { section: address.section, block: address.block, offset: extent.from };
  const focus = address.cell
    ? { section: address.section, block: address.block, offset: extent.to, cell: address.cell }
    : { section: address.section, block: address.block, offset: extent.to };
  return { anchor, focus };
}

/** Opens the link dialog (AC-1, AC-2): prefilled with the selected text and
 * an empty URL, both fields empty without a selection — or, with a collapsed
 * caret inside an existing link, editing that link's whole extent (AC-9). */
function openLinkDialog() {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  dismissLinkChip();
  const selection = currentSelection(root) ?? caretSelection(documentStart());
  const extent = isCollapsed(selection) ? linkAtCaret(model, selection.focus) : null;
  if (extent) {
    linkDialogSelection = extentSelection(extent);
    linkDialogInitial.value = {
      text: extent.text,
      url: extent.target.href ?? "",
      anchor: extent.target.anchor ?? "",
    };
  } else {
    linkDialogSelection = selection;
    linkDialogInitial.value = {
      text: isCollapsed(selection) ? "" : fragmentText(extractFragment(model, selection)),
      url: "",
      anchor: "",
    };
  }
  linkDialogOpen.value = true;
}

/** Applies the dialog's link (AC-3, AC-19, AC-20): the selection is wrapped
 * when its text was kept, or replaced by the edited text. The dialog already
 * rejected dangerous schemes (AC-14..16); sanitizeHref here is a backstop. */
function handleLinkDialogSave(payload = { text: "", href: "", anchor: "" }) {
  const model = doc.value;
  const selection = linkDialogSelection;
  linkDialogOpen.value = false;
  linkDialogSelection = noSelection();
  if (!model || !selection) {
    return;
  }
  const anchor = String(payload.anchor ?? "");
  const href = anchor === "" ? sanitizeHref(String(payload.href ?? "")) : null;
  if (anchor === "" && href === null) {
    return;
  }
  const target = anchor !== "" ? { anchor } : { href: href ?? "" };
  const text = String(payload.text ?? "");
  const selectionText = isCollapsed(selection)
    ? ""
    : fragmentText(extractFragment(model, selection));
  const result =
    !isCollapsed(selection) && text === selectionText
      ? applyLink(model, selection, target)
      : insertLinkText(model, selection, text, target);
  applyEdit(result, selection, "insert link", false);
  contentRoot.value?.focus();
}

function handleLinkDialogCancel() {
  linkDialogOpen.value = false;
  linkDialogSelection = noSelection();
  contentRoot.value?.focus();
  void restoreCaretAfterRender();
}

function dismissLinkChip() {
  if (linkChip.value.visible) {
    linkChip.value = { visible: false, left: 0, top: 0, label: "", external: false };
  }
  linkChipExtent = noLinkExtent();
}

/** Shows the chip when a collapsed caret sits inside a link (AC-8), pinned
 * under the caret like the paste options chip; hides it everywhere else. */
async function refreshLinkChip() {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value || !caret || selectionAnchor || linkDialogOpen.value) {
    dismissLinkChip();
    return;
  }
  const at = caret;
  const extent = linkAtCaret(model, at);
  if (!extent) {
    dismissLinkChip();
    return;
  }
  await nextTick();
  if (caret !== at || !contentRoot.value) {
    return;
  }
  const rect = caretClientRect(root, at);
  if (!rect) {
    dismissLinkChip();
    return;
  }
  linkChipExtent = extent;
  linkChip.value = {
    visible: true,
    left: Math.max(8, rect.left),
    top: rect.bottom + 8,
    label: linkTargetLabel(model, extent.target),
    external: extent.target.href !== undefined,
  };
}

/** Opens a link target: an external URL in a new tab (AC-12, AC-13 — the
 * window.open feature string is the programmatic rel="noopener noreferrer";
 * the rendered <a> carries the literal attribute too), an internal anchor
 * as an in-document jump. Re-sanitized before navigating (AC-17). */
function openLinkTarget(target = noLinkTarget()) {
  if (!target) {
    return;
  }
  if (target.href !== undefined) {
    const href = sanitizeHref(target.href);
    if (href !== null) {
      window.open(href, "_blank", "noopener,noreferrer");
    }
    return;
  }
  jumpToAnchor(target.anchor ?? "");
}

/** Moves the caret to a bookmark or heading anchor and scrolls it into
 * view (AC-24, and internal link opens). */
function jumpToAnchor(anchor = "") {
  const model = doc.value;
  if (!model || anchor === "") {
    return;
  }
  const position = anchorCaret(model, anchor);
  if (!position) {
    return;
  }
  applyCaretMove(position, false);
  void scrollModelCaretIntoView(position);
}

/** Brings a model position into view and settles it a third of the way down
 * the canvas — a deliberate move to somewhere else in the document (a
 * bookmark jump, the page a break just started), where landing on the
 * viewport's very edge would show the destination without its context.
 * `always` skips the "only when it is off-screen" test, for callers that
 * know the caret changed page and want that page framed either way. */
async function scrollModelCaretIntoView(position = noCaret(), always = false) {
  await nextTick();
  const root = contentRoot.value;
  if (!root || !position || !scrollContainer) {
    return;
  }
  const rect = caretClientRect(root, position);
  if (!rect) {
    return;
  }
  const containerRect = scrollContainer.getBoundingClientRect();
  if (always || rect.top < containerRect.top || rect.bottom > containerRect.bottom) {
    scrollContainer.scrollTop += rect.top - containerRect.top - containerRect.height / 3;
  }
}

/** Keeps the caret on screen after an edit, moving the view as little as
 * possible. An edit must never leave the user typing somewhere they cannot
 * see — filling a page pushes the caret past the fold, and until this ran
 * the canvas simply stayed put — but scrolling further than the line just
 * written would lose their place in its own way, so this nudges rather than
 * reframes (contrast scrollModelCaretIntoView above). */
async function keepCaretVisible() {
  await nextTick();
  const root = contentRoot.value;
  if (!root || !caret || !scrollContainer) {
    return;
  }
  const rect = caretClientRect(root, caret);
  if (!rect) {
    return;
  }
  const view = scrollContainer.getBoundingClientRect();
  const pad = 24;
  if (rect.top < view.top + pad) {
    scrollContainer.scrollTop -= view.top + pad - rect.top;
  } else if (rect.bottom > view.bottom - pad) {
    scrollContainer.scrollTop += rect.bottom - (view.bottom - pad);
  }
}

function handleLinkChipOpen() {
  const extent = linkChipExtent;
  if (extent) {
    openLinkTarget(extent.target);
  }
}

function handleLinkChipEdit() {
  openLinkDialog();
}

/** Copies the chip's target to the clipboard (AC-10). */
function handleLinkChipCopy() {
  const extent = linkChipExtent;
  if (!extent) {
    return;
  }
  const text = extent.target.href ?? `#${extent.target.anchor ?? ""}`;
  void navigator.clipboard?.writeText(text).catch(() => {});
}

/** Removes the link the caret sits in, keeping the text and its other
 * formatting (AC-11, AC-33, AC-34). */
function handleLinkChipRemove() {
  const model = doc.value;
  if (!model || !editable.value || !caret) {
    return;
  }
  const selection = caretSelection(caret);
  applyEdit(removeLink(model, selection), selection, "remove link", false);
  contentRoot.value?.focus();
}

/** Clicks on rendered hyperlinks (AC-12). Browsers do not navigate a plain
 * click on an <a> inside a contenteditable region — the anchors are real
 * links (href/target/rel) purely for fidelity and Ctrl/Cmd+click — but
 * preventDefault keeps that guaranteed; with Ctrl/Cmd held the target opens
 * through the same sanitized path as the chip. */
function handleContentClick(event = new MouseEvent("click")) {
  const eventTarget = event.target;
  // Clicking a note reference mark jumps to its note (SPEC-031 AC-15).
  const mark = eventTarget instanceof Element ? eventTarget.closest(".doc-note-ref[data-note-id]") : null;
  if (mark) {
    event.preventDefault();
    const noteId = mark.getAttribute("data-note-id") ?? "";
    const note = doc.value ? noteById(doc.value, noteId) : undefined;
    if (note) {
      enterNoteEditing(note.noteId, note.kind, true);
    }
    return;
  }
  const anchor = eventTarget instanceof Element ? eventTarget.closest("a[data-doc-link]") : null;
  if (!anchor) {
    return;
  }
  event.preventDefault();
  if (!event.ctrlKey && !event.metaKey) {
    return;
  }
  const href = anchor.getAttribute("href") ?? "";
  if (href.startsWith("#")) {
    jumpToAnchor(href.slice(1));
  } else if (href !== "") {
    openLinkTarget({ href });
  }
}

/** Autolinks after a completed space/Enter edit (AC-5, AC-6): a bare URL
 * ending where the word ended becomes a link, as its own undo step so
 * Ctrl/Cmd+Z right after reverts just the link (AC-7). The caret stays
 * exactly where the base edit put it — after the space, or at the start of
 * the paragraph Enter created; the autolink itself never moves it. */
function maybeAutolink(result = noCommandResult(), afterEnter = false) {
  const model = doc.value;
  if (!model || !result) {
    return;
  }
  const focus = result.selection.focus;
  const cell = focus.cell;
  if (afterEnter && (focus.offset !== 0 || (cell ? cell.block === 0 : focus.block === 0))) {
    return;
  }
  const address = afterEnter
    ? cell
      ? {
          section: focus.section,
          block: focus.block,
          cell: { row: cell.row, col: cell.col, block: cell.block - 1 },
        }
      : { section: focus.section, block: focus.block - 1 }
    : cell
      ? { section: focus.section, block: focus.block, cell }
      : { section: focus.section, block: focus.block };
  const paragraph = paragraphAt(model, address);
  if (!paragraph) {
    return;
  }
  const endOffset = afterEnter ? blockLength(paragraph) : focus.offset - 1;
  if (endOffset <= 0) {
    return;
  }
  const linked = autolinkAt(model, address, endOffset);
  if (!linked) {
    return;
  }
  const keep = caretSelection(focus);
  applyEdit({ model: linked, selection: keep }, keep, "autolink", false);
}

/** Opens the "Insert bookmark" dialog (AC-22) for the current caret. */
function openBookmarkDialog() {
  const root = contentRoot.value;
  if (!root || !doc.value || !editable.value) {
    return;
  }
  bookmarkDialogSelection = currentSelection(root) ?? caretSelection(documentStart());
  bookmarkDialogOpen.value = true;
}

function handleBookmarkDialogSave(name = "") {
  const model = doc.value;
  const selection = bookmarkDialogSelection;
  bookmarkDialogOpen.value = false;
  bookmarkDialogSelection = noSelection();
  if (!model || !selection || name === "") {
    return;
  }
  if (validateBookmarkName(name, documentBookmarkNames()) !== "") {
    return;
  }
  applyEdit(insertBookmark(model, selection, name), selection, "insert bookmark", false);
  contentRoot.value?.focus();
}

function handleBookmarkDialogCancel() {
  bookmarkDialogOpen.value = false;
  bookmarkDialogSelection = noSelection();
  contentRoot.value?.focus();
  void restoreCaretAfterRender();
}

function handleBookmarkJump(name = "") {
  bookmarkManagerOpen.value = false;
  contentRoot.value?.focus();
  jumpToAnchor(String(name));
}

function handleBookmarkRemove(name = "") {
  const model = doc.value;
  if (!model || !editable.value) {
    return;
  }
  const selection = currentSelection(contentRoot.value) ?? caretSelection(documentStart());
  applyEdit(deleteBookmark(model, selection, String(name)), selection, "delete bookmark", false);
}

function closeBookmarkManager() {
  bookmarkManagerOpen.value = false;
  contentRoot.value?.focus();
}

// --- the editor's context menu (SPEC-016 AC-4..6) --------------------------

/** Right-click on the canvas: the app's own menu replaces the native one so
 * its entries can drive the clipboard pipeline (and tests can reach them). */
function handleContextMenu(event = new MouseEvent("contextmenu")) {
  if (!doc.value || !editable.value) {
    return;
  }
  event.preventDefault();
  // A right-click inside a table cell adds the row/column section to the
  // menu, targeted at that cell (SPEC-022 AC-17..19).
  contextMenuTableCell = tableCellAtTarget(event.target);
  // A right-click on a list paragraph adds the numbering-control section,
  // targeted at that paragraph (SPEC-032 AC-14..16).
  contextMenuListTarget = listParagraphAtTarget(event.target);
  contextMenu.value = { visible: true, left: event.clientX, top: event.clientY };
}

/** The list paragraph under a right-click, or null when the click landed
 * outside a list item (SPEC-032). */
function listParagraphAtTarget(eventTarget = noEventTarget()) {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model) {
    return null;
  }
  const target = paragraphFromEventTarget(root, eventTarget);
  if (!target) {
    return null;
  }
  return paragraphAt(model, target)?.list ? target : null;
}

function closeContextMenu() {
  if (contextMenu.value.visible) {
    contextMenu.value = { visible: false, left: 0, top: 0 };
  }
  contextMenuTableCell = null;
  contextMenuListTarget = null;
}

/** A numbering-control entry picked from the context menu (SPEC-032
 * AC-14..16), applied at the right-clicked paragraph. "Set numbering
 * value…" prompts for the ordinal; cancelling the prompt changes nothing. */
function handleContextMenuListAction(action = "") {
  const target = contextMenuListTarget;
  closeContextMenu();
  const model = doc.value;
  if (!target || !model || !editable.value) {
    return;
  }
  const selection = caretSelection({ ...target, offset: 0 });
  if (action === "restart-numbering") {
    applyEdit(restartListNumbering(model, selection), selection, "restart numbering");
  } else if (action === "continue-numbering") {
    applyEdit(continueListNumbering(model, selection), selection, "continue numbering");
  } else if (action === "set-numbering-value") {
    const raw = window.prompt("Set numbering value", "1");
    if (raw === null) {
      return;
    }
    const value = Number(raw);
    if (!Number.isInteger(value) || value < 0) {
      return;
    }
    applyEdit(setListNumberingValue(model, selection, value), selection, "set numbering value");
  }
}

/** Menu Cut/Copy synthesise a real cut/copy via execCommand — the only
 * synchronous route from a button click to clipboardData (and the private
 * MIME type). The events land in handleCut/handleCopy above. */
function contextMenuCut() {
  closeContextMenu();
  const root = contentRoot.value;
  if (!root) {
    return;
  }
  root.focus();
  if (typeof document.execCommand === "function") {
    document.execCommand("cut");
  }
}

function contextMenuCopy() {
  closeContextMenu();
  const root = contentRoot.value;
  if (!root) {
    return;
  }
  root.focus();
  if (typeof document.execCommand === "function") {
    document.execCommand("copy");
  }
}

/** Reads text/html, text/plain and image presence via the async clipboard
 * API — all a button click can reach. Null when the API is missing or the
 * read was denied. */
async function readAsyncClipboard() {
  const clipboard = navigator.clipboard;
  if (!clipboard) {
    return null;
  }
  try {
    if (typeof clipboard.read === "function") {
      const items = await clipboard.read();
      let html = "";
      let text = "";
      let hasImage = false;
      for (const item of items) {
        for (const type of item.types) {
          if (type === "text/html") {
            html = await (await item.getType(type)).text();
          } else if (type === "text/plain") {
            text = await (await item.getType(type)).text();
          } else if (type.startsWith("image/")) {
            hasImage = true;
          }
        }
      }
      return { html, text, hasImage };
    }
    if (typeof clipboard.readText === "function") {
      return { html: "", text: await clipboard.readText(), hasImage: false };
    }
  } catch {
    return null;
  }
  return null;
}

/** The menu's Paste (AC-6). The async API cannot carry the private MIME
 * type, so a clipboard whose text/html matches our last copy is recognised
 * as our own and pasted from the local fragment snapshot with exact
 * fidelity; anything else takes the same import paths as a keyboard
 * paste. */
async function contextMenuPaste() {
  closeContextMenu();
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  root.focus();
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  const flavours = await readAsyncClipboard();
  if (!flavours) {
    addToast("Clipboard unavailable — paste with the keyboard instead");
    return;
  }
  const own =
    lastCopy && flavours.html && flavours.html === lastCopy.html
      ? lastCopy.fragment
      : null;
  performPaste(
    {
      internal: own ? serializeFragment(own) : "",
      html: flavours.html,
      text: flavours.text,
      hasImage: flavours.hasImage,
    },
    selection,
    false,
  );
}

// --- caret movement and selection (SPEC-011) -------------------------------
// Character/word/line movement, plain Home/End and mouse selection are the
// browser's native caret behaviour over the rendered canvas — the DOM
// selection is the *presentation* of the selection, which selectionchange
// mirrors back into model coordinates. Document-level jumps and paging are
// implemented here because native behaviour is unreliable across them.

/** Announces the caret/selection to assistive technology, mirrors it into
 * the tracked caret, and refreshes the toolbar's format button states
 * (SPEC-012 AC-12..14) for the same range. */
function updateCaretStatus(range = { start: noCaret(), end: noCaret() }) {
  const model = doc.value;
  if (!model || !range.start || !range.end) {
    formatStates.value = EMPTY_FORMAT_STATES;
    formatValues.value = emptyToolbarFormatValues();
    paragraphState.value = EMPTY_PARAGRAPH_STATE;
    paragraphFormatting.value = emptyParagraphFormatting();
    listState.value = emptyToolbarListState();
    styleState.value = emptyToolbarStyleState();
    paragraphPagination.value = emptyParagraphPaginationState();
    return;
  }
  const total = blockCount(model);
  if (compareCaret(range.start, range.end) === 0) {
    caretStatus.value =
      `Paragraph ${blockOrdinal(model, range.end)} of ${total}, ` +
      `character ${range.end.offset + 1}`;
  } else {
    caretStatus.value =
      `Selected from paragraph ${blockOrdinal(model, range.start)}, ` +
      `character ${range.start.offset + 1} to paragraph ` +
      `${blockOrdinal(model, range.end)}, character ${range.end.offset + 1}`;
  }
  const selection = { anchor: range.start, focus: range.end };
  formatStates.value = {
    bold: formatState(model, selection, "bold", pendingFormat),
    italic: formatState(model, selection, "italic", pendingFormat),
    underline: formatState(model, selection, "underline", pendingFormat),
    strikethrough: formatState(model, selection, "strikethrough", pendingFormat),
    subscript: verticalAlignState(model, selection, "subscript", pendingFormat),
    superscript: verticalAlignState(model, selection, "superscript", pendingFormat),
  };
  formatValues.value = toolbarFormatValues(model, selection, pendingFormat);
  const firstFormatting = firstTouchedFormatting(model, selection);
  paragraphFormatting.value = firstFormatting;
  paragraphState.value = {
    alignment: firstFormatting.alignment ?? "left",
    spaceBefore: hasSpaceBefore(firstFormatting),
    direction: firstFormatting.direction ?? "ltr",
  };
  listState.value = toolbarListState(model, selection);
  styleState.value = toolbarStyleState(model, selection);
  paragraphPagination.value = paragraphPaginationState(model, selection);
  caretSectionIndex.value = range.end.section;
  // Both ends of a selection must stay realized (a virtualized endpoint
  // would collapse the DOM selection); a collapsed caret pins only its own
  // page via refreshCaretPage.
  anchorPage.value =
    compareCaret(range.start, range.end) === 0 ? -1 : pagination.pageOfCaret(range.start);
  refreshCaretPage();
}

/** Moves the caret (or, with extend, the selection's focus end) to a model
 * position and refreshes tracking + announcement. Explicit navigation, so
 * any pending format is discarded (SPEC-012 AC-11). */
function applyCaretMove(target = noCaret(), extend = false) {
  const root = contentRoot.value;
  if (!root || !target) {
    return;
  }
  pendingFormat = noPendingFormat();
  // Explicit navigation breaks the undo history's typing group too (SPEC-015
  // AC-10) — handleSelectionChange only catches *native* moves (arrow keys,
  // a native click/drag), not these app-driven ones, because `caret` is
  // already updated to match by the time their selectionchange fires.
  breakTypingGroup(documentId.value);
  const place = () => {
    if (extend) {
      extendDomSelection(root, target);
    } else {
      setDomCaret(root, target);
    }
    caret = target;
    const range = modelRangeFromSelection(root, document.getSelection());
    updateCaretStatus(range ?? { start: target, end: target });
  };
  if (paragraphElementForCaret(root, target)) {
    place();
    return;
  }
  // The target sits on a virtualized page: update the caret trackers first
  // (which realizes that page) and place once the patch has rendered it.
  caret = target;
  const anchor = extend && selectionAnchor ? selectionAnchor : target;
  updateCaretStatus({ start: anchor, end: target });
  void nextTick().then(place);
}

/** The caret's moving end (the DOM selection's focus), which Shift+paging
 * keeps extending; falls back to the tracked caret. */
function domFocusCaret(root = noElement()) {
  if (!root) {
    return caret;
  }
  const selection = document.getSelection();
  if (selection && selection.rangeCount > 0 && selection.focusNode) {
    const focus = caretFromDomPoint(root, selection.focusNode, selection.focusOffset);
    if (focus) {
      return focus;
    }
  }
  return caret;
}

/** PageUp/PageDown: scrolls one viewport and moves the caret to the text
 * position one viewport away from where it was, snapping to the document's
 * start/end when there is nothing further to page onto. */
async function moveCaretByPage(direction = 1, extend = false) {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model) {
    return;
  }
  const from = domFocusCaret(root) ?? { section: 0, block: 0, offset: 0 };
  const rect = caretClientRect(root, from);
  const container = scrollContainer;
  const pageHeight =
    container && container.clientHeight > 0
      ? container.clientHeight
      : window.innerHeight;
  let target = null;
  if (rect) {
    const before = container ? container.scrollTop : 0;
    if (container) {
      container.scrollTop = before + direction * pageHeight;
      // The page scrolled onto may be a virtualized empty box: realize the
      // new viewport before hit-testing for the caret's landing position.
      refreshViewportRange();
      await nextTick();
    }
    const scrolled = container ? container.scrollTop - before : 0;
    target = caretAtClientPoint(
      root,
      rect.left,
      rect.top + direction * pageHeight - scrolled,
    );
    if (target && compareCaret(target, from) === 0) {
      target = null;
    }
  }
  if (!target) {
    target = direction > 0 ? documentEnd(model) : documentStart();
  }
  applyCaretMove(target, extend);
}

/** Home/End — the start and end of the caret's *visual* line, so a wrapped
 * paragraph stops at the wrap and not at the paragraph's ends (SPEC-011
 * AC-12). Windows and Linux browsers bind these themselves and the native
 * move is left alone; macOS binds them to scrolling the view instead, which
 * moves no caret at all, so there the editor has to do it. Hit-testing the
 * caret's own line at the paragraph's left/right edge is what makes it
 * visual: the browser resolves the point to the nearest position on that
 * line, wherever the text happens to wrap. Returns true when consumed. */
function handleLineEdgeKey(event = new KeyboardEvent("keydown")) {
  const root = contentRoot.value;
  if (!root || !isMacPlatform() || (event.key !== "Home" && event.key !== "End")) {
    return false;
  }
  if (event.ctrlKey || event.metaKey || event.altKey) {
    // Cmd+Home/End is the document-start/end command in the registry.
    return false;
  }
  const from = domFocusCaret(root);
  const rect = from ? caretClientRect(root, from) : null;
  const paragraph = from ? paragraphElementForCaret(root, from) : null;
  if (!from || !rect || !paragraph) {
    return false;
  }
  const bounds = paragraph.getBoundingClientRect();
  const toStart = event.key === "Home";
  const target = caretAtClientPoint(
    root,
    toStart ? bounds.left + 1 : bounds.right - 1,
    rect.top + rect.height / 2,
  );
  if (!target) {
    return false;
  }
  event.preventDefault();
  applyCaretMove(target, event.shiftKey);
  return true;
}

// --- tables: arrow fallback, drag selection, handles (SPEC-022) -------------

/** Deterministic cell-boundary arrows (AC-13): a collapsed caret exactly at
 * the boundary its arrow points across — end of the cell for ArrowRight,
 * start for ArrowLeft, first/last visual line of the first/last paragraph
 * for ArrowUp/ArrowDown — moves straight to the neighbouring cell (or out
 * of the table) instead of relying on the browser. Returns true when the
 * event was consumed. */
function handleTableArrowKey(event = new KeyboardEvent("keydown")) {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return false;
  }
  const selection = currentSelection(root);
  if (!selection || !isCollapsed(selection) || !selection.focus.cell) {
    return false;
  }
  const focus = selection.focus;
  const cellBlock = paragraphAtCaret(model, focus);
  const cellPath = focus.cell;
  if (!cellBlock || !cellPath) {
    return false;
  }
  const table = model.sections[focus.section].blocks[focus.block];
  if (!isTable(table)) {
    return false;
  }
  const cellBlocks = table.rows[cellPath.row].cells[cellPath.col].blocks;
  const atFirstParagraph = cellPath.block === 0;
  const atLastParagraph = cellPath.block === cellBlocks.length - 1;
  let direction = "";
  if (event.key === "ArrowRight" && atLastParagraph && focus.offset >= blockLength(cellBlock)) {
    direction = "right";
  } else if (event.key === "ArrowLeft" && atFirstParagraph && focus.offset === 0) {
    direction = "left";
  } else if (event.key === "ArrowDown" && atLastParagraph && caretOnBoundaryLine(root, focus, "last")) {
    direction = "down";
  } else if (event.key === "ArrowUp" && atFirstParagraph && caretOnBoundaryLine(root, focus, "first")) {
    direction = "up";
  }
  if (direction === "") {
    return false;
  }
  event.preventDefault();
  const target = arrowEscapeCaret(
    model,
    focus,
    direction === "up" ? "up" : direction === "down" ? "down" : direction === "left" ? "left" : "right",
  );
  if (target) {
    applyCaretMove(target);
  }
  return true;
}

/** After the browser processed an arrow keydown with a collapsed caret in a
 * table cell, checks whether the caret actually moved; when it did not, the
 * model-level escape (lib/tables.arrowEscapeCaret) moves it to the
 * neighbouring cell or out of the table — the AC-14 guarantee that native
 * visual navigation cannot strand the caret. */
function scheduleTableArrowFallback(key = "") {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection || !isCollapsed(selection) || !selection.focus.cell) {
    return;
  }
  const before = selection.focus;
  const direction =
    key === "ArrowUp" ? "up" : key === "ArrowDown" ? "down" : key === "ArrowLeft" ? "left" : "right";
  window.setTimeout(() => {
    const model = doc.value;
    const currentRoot = contentRoot.value;
    if (!model || !currentRoot || !editable.value) {
      return;
    }
    const after = currentSelection(currentRoot);
    if (!after || !isCollapsed(after) || compareCaret(after.focus, before) !== 0) {
      return;
    }
    const target = arrowEscapeCaret(model, before, direction);
    if (target) {
      applyCaretMove(target);
    }
  }, 0);
}

// --- cross-column caret navigation (SPEC-030 AC-20/AC-21) -------------------

// The block-level children of a rendered column box (skipping zero-height
// section markers, which are not navigation targets).
function columnBlockChildren(columnEl = noElement()) {
  if (!columnEl) {
    return [];
  }
  return Array.from(columnEl.children).filter((child) => child.hasAttribute("data-section"));
}

// Where the caret should land when a vertical arrow crosses a column boundary,
// or null when it should not (more lines remain in the column, the caret is
// not in a column, or there is no adjacent column). Down at the bottom of a
// column targets the top of the next column (or the first column of the next
// page); Up at the top targets the bottom of the previous column (or the last
// column of the previous page).
function columnEscapeTarget(root = noElement(), focus = noCaret(), direction = "down") {
  if (!root || !focus) {
    return null;
  }
  const paragraphEl = paragraphElementForCaret(root, focus);
  const columnEl = paragraphEl?.closest(".doc-column") ?? null;
  const pageEl = columnEl?.closest(".page") ?? null;
  if (!paragraphEl || !columnEl || !pageEl) {
    return null;
  }
  const children = columnBlockChildren(columnEl);
  const dataCol = Number(columnEl.getAttribute("data-column"));
  if (direction === "down") {
    if (paragraphEl !== children[children.length - 1] || !caretOnBoundaryLine(root, focus, "last")) {
      return null;
    }
    let targetCol = pageEl.querySelector(`.doc-column[data-column="${dataCol + 1}"]`);
    if (!targetCol) {
      targetCol = pageEl.nextElementSibling?.querySelector('.doc-column[data-column="0"]') ?? null;
    }
    const first = targetCol ? columnBlockChildren(targetCol)[0] : null;
    // The top of the next column: the start of its first block (AC-20/21).
    return first ? caretAtBlockEdge(first, "start") : null;
  }
  if (paragraphEl !== children[0] || !caretOnBoundaryLine(root, focus, "first")) {
    return null;
  }
  let targetCol = null;
  if (dataCol > 0) {
    targetCol = pageEl.querySelector(`.doc-column[data-column="${dataCol - 1}"]`);
  } else {
    const prevCols = pageEl.previousElementSibling?.querySelectorAll(".doc-column");
    targetCol = prevCols && prevCols.length > 0 ? prevCols[prevCols.length - 1] : null;
  }
  const blocks = targetCol ? columnBlockChildren(targetCol) : [];
  const last = blocks[blocks.length - 1] ?? null;
  // The bottom of the previous column: the end of its last block.
  return last ? caretAtBlockEdge(last, "end") : null;
}

// The model caret at the start or end of a rendered block element, read from
// its data-section/data-block/data-slice-start attributes — deterministic,
// unlike a geometry hit test that cannot tell same-height lines of adjacent
// columns apart.
function caretAtBlockEdge(blockEl = noElement(), edge = "start") {
  if (!blockEl) {
    return null;
  }
  const section = Number(blockEl.getAttribute("data-section"));
  const block = Number(blockEl.getAttribute("data-block"));
  if (!Number.isInteger(section) || !Number.isInteger(block)) {
    return null;
  }
  const sliceStart = Number(blockEl.getAttribute("data-slice-start") ?? "0");
  // modelTextLengthOf, not textContent.length: a note reference mark's
  // rendered number stands for a one-character placeholder (SPEC-031).
  const offset = edge === "start" ? sliceStart : sliceStart + modelTextLengthOf(blockEl);
  return { section, block, offset };
}

/** Deterministic cross-column arrows (AC-20/AC-21): Down at the bottom line of
 * a column jumps to the top of the next column (or next page's first column),
 * Up at the top line jumps to the bottom of the previous column. Returns true
 * when the event was consumed. */
function handleColumnArrowKey(event = new KeyboardEvent("keydown")) {
  const root = contentRoot.value;
  if (!root || !editable.value || (event.key !== "ArrowUp" && event.key !== "ArrowDown")) {
    return false;
  }
  const selection = currentSelection(root);
  if (!selection || !isCollapsed(selection) || selection.focus.cell) {
    return false;
  }
  const target = columnEscapeTarget(root, selection.focus, event.key === "ArrowDown" ? "down" : "up");
  if (!target) {
    return false;
  }
  event.preventDefault();
  applyCaretMove(target);
  return true;
}

// After a native vertical arrow with a collapsed caret in a column made no
// progress, cross the column boundary — the guarantee columns cannot trap the
// caret, mirroring scheduleTableArrowFallback.
function scheduleColumnArrowFallback(key = "") {
  const root = contentRoot.value;
  if (!root || !editable.value || (key !== "ArrowUp" && key !== "ArrowDown")) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection || !isCollapsed(selection) || selection.focus.cell) {
    return;
  }
  const paragraphEl = paragraphElementForCaret(root, selection.focus);
  if (!paragraphEl?.closest(".doc-column")) {
    return;
  }
  const before = selection.focus;
  const direction = key === "ArrowDown" ? "down" : "up";
  window.setTimeout(() => {
    const currentRoot = contentRoot.value;
    if (!currentRoot || !editable.value) {
      return;
    }
    const after = currentSelection(currentRoot);
    if (!after || !isCollapsed(after) || compareCaret(after.focus, before) !== 0) {
      return;
    }
    const target = columnEscapeTarget(currentRoot, before, direction);
    if (target && compareCaret(target, before) !== 0) {
      applyCaretMove(target);
    }
  }, 0);
}

// Tracks the primary button across a text-selection drag (SPEC-020): the
// format painter must not paint a partial selection mid-drag, only the
// range the user settles on at mouseup — see maybeApplyPainter.
let mouseButtonDown = false;

/** Mousedown on the canvas: remembers the cell under the pointer as a
 * potential drag origin (AC-15) and dismisses any existing cell-range
 * highlight — a plain click always returns to a normal text caret. */
function handleCanvasMouseDown(event = new MouseEvent("mousedown")) {
  if (event.button !== 0) {
    return;
  }
  // Clicking the body while a note is being edited leaves note editing
  // (SPEC-031) — clicks inside the note area stop propagation before here.
  exitNoteEditing();
  mouseButtonDown = true;
  clearCellSelection();
  cellDragOrigin = editable.value ? tableCellAtTarget(event.target) : null;
  // A click outside a selected image deselects it (SPEC-024) — a click on
  // one is handled by DocImageBlock's own @select, which fires afterward.
  const imageTarget = event.target instanceof Element ? event.target.closest(".doc-image-anchor") : null;
  if (!imageTarget) {
    clearImageSelection();
  }
}

/** Mousemove with the primary button down: entering a different cell of the
 * drag origin's table activates (or extends) the rectangular cell range
 * (AC-15). The DOM text selection is collapsed away — the highlighted cells
 * are the selection's presentation (AC-16). */
function handleCanvasMouseMove(event = new MouseEvent("mousemove")) {
  const origin = cellDragOrigin;
  if (!origin || (event.buttons & 1) === 0) {
    return;
  }
  const cell = tableCellAtTarget(event.target);
  if (!cell || cell.section !== origin.section || cell.block !== origin.block) {
    return;
  }
  const active = cellSelection.value;
  if (!active && cell.row === origin.row && cell.col === origin.col) {
    return;
  }
  event.preventDefault();
  const next = {
    section: origin.section,
    block: origin.block,
    anchor: { row: origin.row, col: origin.col },
    focus: { row: cell.row, col: cell.col },
  };
  const unchanged =
    active && active.focus.row === next.focus.row && active.focus.col === next.focus.col;
  if (!unchanged) {
    cellSelection.value = next;
    const domSelection = document.getSelection();
    if (domSelection && domSelection.rangeCount > 0) {
      domSelection.removeAllRanges();
    }
  }
}

function handleCanvasMouseUp() {
  cellDragOrigin = null;
  if (mouseButtonDown) {
    mouseButtonDown = false;
    maybeApplyPainter();
  }
}

/** Whether a drag carries a file of an accepted image type — checked from
 * `dataTransfer.items` (available during dragenter/dragover, unlike the
 * actual file, which only `drop` exposes). Used to steal the drag away from
 * AppShell's window-level "drop to open a document" handling (SPEC-009) only
 * when it is actually an image (AC-1); anything else — including a .docx/
 * .odt file dragged over the canvas — falls through to that existing
 * behaviour untouched. */
function isImageDrag(event = new DragEvent("dragover")) {
  const items = event.dataTransfer?.items;
  if (!items) {
    return false;
  }
  for (let i = 0; i < items.length; i++) {
    if (items[i].kind === "file" && isAcceptedImageType(items[i].type)) {
      return true;
    }
  }
  return false;
}

function handleCanvasDragEnter(event = new DragEvent("dragenter")) {
  if (editable.value && isImageDrag(event)) {
    event.stopPropagation();
  }
}

function handleCanvasDragOver(event = new DragEvent("dragover")) {
  if (editable.value && isImageDrag(event)) {
    event.preventDefault();
    event.stopPropagation();
  }
}

/** Drag-and-drop onto the canvas (AC-1): inserts at the caret nearest the
 * drop point. */
function handleCanvasDrop(event = new DragEvent("drop")) {
  if (!editable.value) {
    return;
  }
  const file = event.dataTransfer?.files?.[0];
  if (!file || !isAcceptedImageType(file.type)) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  const root = contentRoot.value;
  const dropCaret = root ? caretAtClientPoint(root, event.clientX, event.clientY) : null;
  void insertImageFromFile(file, dropCaret ? caretSelection(dropCaret) : null);
}

/** Mouseover anywhere on the page: hovering a table measures its geometry
 * for the insert/delete handles (AC-20, AC-21); hovering the handles keeps
 * them; anything else hides them. */
function handlePageMouseOver(event = new MouseEvent("mouseover")) {
  if (!editable.value || !(event.target instanceof Element)) {
    return;
  }
  if (event.target.closest(".table-handles")) {
    return;
  }
  const table = event.target.closest("table.doc-table");
  const pageEl = pageRef.value;
  if (!table || !pageEl) {
    if (tableHover.value) {
      tableHover.value = null;
    }
    return;
  }
  tableHover.value = measureTableHandles(pageEl, table);
}

/** One handle/menu table action against an explicit target cell. `before`
 * is the current selection for undo restoration. */
function applyTableAction(action = "", target = { section: 0, block: 0, row: 0, col: 0 }) {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const location = target;
  const fallbackCaret = {
    section: location.section,
    block: location.block,
    offset: 0,
    cell: { row: location.row, col: location.col, block: 0 },
  };
  const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
  let result = noCommandResult();
  let label = "";
  if (action === "insert-row-above") {
    result = insertTableRow(model, selection, location, "above");
    label = "insert row";
  } else if (action === "insert-row-below") {
    result = insertTableRow(model, selection, location, "below");
    label = "insert row";
  } else if (action === "insert-col-left") {
    result = insertTableColumn(model, selection, location, "left");
    label = "insert column";
  } else if (action === "insert-col-right") {
    result = insertTableColumn(model, selection, location, "right");
    label = "insert column";
  } else if (action === "delete-row") {
    result = deleteTableRow(model, selection, location);
    label = "delete row";
  } else if (action === "delete-col") {
    result = deleteTableColumn(model, selection, location);
    label = "delete column";
  } else if (action === "delete-table") {
    result = deleteTable(model, selection, location);
    label = "delete table";
  }
  if (result) {
    applyEdit(result, selection, label);
    root.focus();
  }
}

/** A hover handle's click (AC-20, AC-21): boundary index `at` for inserts
 * (0 = before the first row/column), row/column index for deletes. */
function handleTableHandle(kind = "", at = 0) {
  const hover = tableHover.value;
  if (!hover) {
    return;
  }
  const base = { section: hover.section, block: hover.block, row: 0, col: 0 };
  if (kind === "insert-row") {
    applyTableAction(at === 0 ? "insert-row-above" : "insert-row-below", {
      ...base,
      row: at === 0 ? 0 : at - 1,
    });
  } else if (kind === "delete-row") {
    applyTableAction("delete-row", { ...base, row: at });
  } else if (kind === "insert-col") {
    applyTableAction(at === 0 ? "insert-col-left" : "insert-col-right", {
      ...base,
      col: at === 0 ? 0 : at - 1,
    });
  } else if (kind === "delete-col") {
    applyTableAction("delete-col", { ...base, col: at });
  }
}

// --- column/row resize (SPEC-023 AC-1..5) -----------------------------------

/** Dragging a column boundary handle (AC-1..3): live-resizes the column on
 * either side of the boundary (or, for the table's own right edge, just
 * that column) directly in the DOM; the model updates once, on release. */
function startColumnResize(event = new MouseEvent("mousedown"), colIndex = 0) {
  const hover = tableHover.value;
  const model = doc.value;
  const root = contentRoot.value;
  if (!hover || !model || !root || !editable.value) {
    return;
  }
  const table = tableAt(model, hover.section, hover.block);
  const tableEl = findTableElement(root, hover.section, hover.block);
  if (!table || !tableEl) {
    return;
  }
  event.preventDefault();
  columnResizeDrag = {
    section: hover.section,
    block: hover.block,
    colIndex,
    startX: event.clientX,
    startWidthsCm: [...table.columnWidthsCm],
    tableEl,
    liveWidthsCm: null,
  };
  window.addEventListener("mousemove", handleColumnResizeMove);
  window.addEventListener("mouseup", finishColumnResize);
}

function handleColumnResizeMove(event = new MouseEvent("mousemove")) {
  const drag = columnResizeDrag;
  if (!drag) {
    return;
  }
  const deltaCm = (event.clientX - drag.startX) * CM_PER_PX;
  const widths = [...drag.startWidthsCm];
  const i = drag.colIndex;
  if (i < widths.length - 1) {
    // An internal boundary: the neighbour absorbs exactly what this column
    // gains, so the table's overall width is unchanged (AC-1, AC-2).
    const total = widths[i] + widths[i + 1];
    const next = Math.min(Math.max(widths[i] + deltaCm, MIN_COLUMN_WIDTH_CM), total - MIN_COLUMN_WIDTH_CM);
    widths[i] = next;
    widths[i + 1] = total - next;
  } else {
    // The table's own right edge: only the last column changes.
    widths[i] = Math.max(MIN_COLUMN_WIDTH_CM, widths[i] + deltaCm);
  }
  drag.liveWidthsCm = widths;
  previewColumnWidths(drag.tableEl, widths);
}

/** Release (AC-3): the live DOM widths become the one model update this
 * whole drag makes. */
function finishColumnResize() {
  window.removeEventListener("mousemove", handleColumnResizeMove);
  window.removeEventListener("mouseup", finishColumnResize);
  const drag = columnResizeDrag;
  columnResizeDrag = null;
  const root = contentRoot.value;
  const model = doc.value;
  if (!drag || !drag.liveWidthsCm || !root || !model) {
    return;
  }
  const location = { section: drag.section, block: drag.block, row: 0, col: drag.colIndex };
  const fallbackCaret = {
    section: drag.section,
    block: drag.block,
    offset: 0,
    cell: { row: 0, col: drag.colIndex, block: 0 },
  };
  const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
  applyEdit(setColumnWidths(model, selection, location, drag.liveWidthsCm), selection, "resize column");
}

/** Dragging a row boundary handle (AC-4, AC-5): live-resizes just that row;
 * the model's minimum height updates once, on release. Starts from the
 * row's current minimum, or (unset) its rendered height, so the boundary
 * starts exactly under the pointer. */
function startRowResize(event = new MouseEvent("mousedown"), rowIndex = 0) {
  const hover = tableHover.value;
  const model = doc.value;
  const root = contentRoot.value;
  if (!hover || !model || !root || !editable.value) {
    return;
  }
  const table = tableAt(model, hover.section, hover.block);
  const tableEl = findTableElement(root, hover.section, hover.block);
  const trEl = tableEl ? tableRowElement(tableEl, rowIndex) : null;
  if (!table || !trEl) {
    return;
  }
  event.preventDefault();
  const row = table.rows[rowIndex];
  const startHeightCm = row.minHeightCm ?? trEl.getBoundingClientRect().height * CM_PER_PX;
  rowResizeDrag = {
    section: hover.section,
    block: hover.block,
    rowIndex,
    startY: event.clientY,
    startHeightCm,
    trEl,
    liveHeightCm: null,
  };
  window.addEventListener("mousemove", handleRowResizeMove);
  window.addEventListener("mouseup", finishRowResize);
}

function handleRowResizeMove(event = new MouseEvent("mousemove")) {
  const drag = rowResizeDrag;
  if (!drag) {
    return;
  }
  const deltaCm = (event.clientY - drag.startY) * CM_PER_PX;
  const heightCm = Math.max(MIN_ROW_HEIGHT_CM, drag.startHeightCm + deltaCm);
  drag.liveHeightCm = heightCm;
  previewRowHeight(drag.trEl, heightCm);
}

/** Release (AC-4): the live height becomes the row's minimum — a floor the
 * row can still grow past for its content (AC-5, see renderFormatting.ts's
 * tableRowStyle). */
function finishRowResize() {
  window.removeEventListener("mousemove", handleRowResizeMove);
  window.removeEventListener("mouseup", finishRowResize);
  const drag = rowResizeDrag;
  rowResizeDrag = null;
  const root = contentRoot.value;
  const model = doc.value;
  if (!drag || drag.liveHeightCm === null || !root || !model) {
    return;
  }
  const location = { section: drag.section, block: drag.block, row: drag.rowIndex, col: 0 };
  const fallbackCaret = {
    section: drag.section,
    block: drag.block,
    offset: 0,
    cell: { row: drag.rowIndex, col: 0, block: 0 },
  };
  const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
  applyEdit(
    setRowMinHeight(model, selection, location, drag.rowIndex, drag.liveHeightCm),
    selection,
    "resize row",
  );
}

/** The context menu's table action (AC-17..19), against the cell the
 * right-click landed in — or, for the SPEC-023 formatting actions, routed
 * to their own handlers instead of the row/column command dispatch above. */
function handleContextMenuTableAction(action = "") {
  const target = contextMenuTableCell;
  closeContextMenu();
  if (!target) {
    return;
  }
  if (action === "merge-cells") {
    handleMergeCells();
  } else if (action === "split-cell") {
    openSplitCellDialog(target);
  } else if (action === "table-properties") {
    openTableProperties(target);
  } else {
    applyTableAction(action, target);
  }
}

/** "Merge cells" (AC-6, AC-7): the active cell-range selection, or a no-op
 * without one (the menu only offers it with a multi-cell range active). */
function handleMergeCells() {
  const root = contentRoot.value;
  const model = doc.value;
  const range = cellSelection.value;
  if (!root || !model || !editable.value || !range) {
    return;
  }
  const selection = currentSelection(root) ?? caretSelection({ section: range.section, block: range.block, offset: 0 });
  applyEdit(mergeCells(model, selection, range), selection, "merge cells");
  root.focus();
}

/** "Split cell…" (AC-8, AC-9): opens the dialog against the right-clicked
 * cell (or the active range's anchor cell, if one covers it). */
function openSplitCellDialog(target = { section: 0, block: 0, row: 0, col: 0 }) {
  tableFormattingLocation = target;
  splitCellDialogOpen.value = true;
}

/** OK in the split-cell dialog: one combined edit (AC-9, AC-34). */
function handleSplitCellApply(size = { rows: 1, cols: 1 }) {
  splitCellDialogOpen.value = false;
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const location = tableFormattingLocation;
  const fallbackCaret = {
    section: location.section,
    block: location.block,
    offset: 0,
    cell: { row: location.row, col: location.col, block: 0 },
  };
  const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
  applyEdit(splitCell(model, selection, location, size.rows, size.cols), selection, "split cell");
  root.focus();
}

/** The table properties dialog (SPEC-023 AC-10..23): opened against the
 * right-clicked cell's table, prefilled from that cell's own current
 * values. Border/shading/padding/vertical-alignment apply to the active
 * cell-range selection when one covers more than one cell (AC-14),
 * otherwise to the whole table (AC-13); alignment/width always apply to the
 * whole table. */
function openTableProperties(target = { section: 0, block: 0, row: 0, col: 0 }) {
  const model = doc.value;
  if (!model) {
    return;
  }
  const table = tableAt(model, target.section, target.block);
  if (!table) {
    return;
  }
  tableFormattingLocation = target;
  const range = cellSelection.value ?? fullTableRange(model, target);
  const bounds = range ? cellRangeBounds(range) : { top: target.row, left: target.col, bottom: target.row, right: target.col };
  const cell = table.rows[bounds.top]?.cells[bounds.left];
  const scope =
    cellSelection.value && (bounds.top !== bounds.bottom || bounds.left !== bounds.right)
      ? "the selected cells"
      : "the whole table";
  const firstBorder = cell?.borders?.top ?? cell?.borders?.bottom ?? cell?.borders?.left ?? cell?.borders?.right;
  tablePropertiesInitial.value = {
    borderStyle: firstBorder?.style ?? "single",
    borderWidthPt: firstBorder?.widthPt ?? 1,
    borderColor: firstBorder && firstBorder.color !== "auto" ? firstBorder.color : "#000000",
    shadingFill: cell?.shadingFill ?? "",
    paddingCm: cell?.paddingCm ?? 0.19,
    verticalAlign: cell?.verticalAlign ?? "top",
    alignment: table.alignment ?? "left",
    widthMode: table.widthMode ?? "fixed",
    widthPercent: table.widthPercent ?? 100,
    scope,
  };
  tablePropertiesOpen.value = true;
}

/** OK in the table properties dialog: every field the dialog carries,
 * chained through the model as one combined edit — a single undo step
 * (AC-34) regardless of how many of the dialog's sections the user
 * touched. */
function handleTablePropertiesApply(
  payload = {
    borderPreset: "",
    borderStyle: "single",
    borderWidthPt: 1,
    borderColor: "#000000",
    shadingFill: "",
    paddingCm: 0.19,
    verticalAlign: "top",
    alignment: "left",
    widthMode: "fixed",
    widthPercent: 100,
  },
) {
  tablePropertiesOpen.value = false;
  const root = contentRoot.value;
  const model = doc.value;
  const location = tableFormattingLocation;
  if (!root || !model || !editable.value) {
    return;
  }
  const range = cellSelection.value ?? fullTableRange(model, location);
  if (!range) {
    return;
  }
  const fallbackCaret = {
    section: location.section,
    block: location.block,
    offset: 0,
    cell: { row: location.row, col: location.col, block: 0 },
  };
  const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
  let current = model;
  if (payload.borderPreset) {
    current = applyBorderPreset(current, selection, range, asBorderPreset(payload.borderPreset), {
      style: asBorderLineStyle(payload.borderStyle),
      widthPt: payload.borderWidthPt,
      color: payload.borderColor,
    }).model;
  }
  current = setCellShading(current, selection, range, payload.shadingFill || undefined).model;
  current = setCellPadding(current, selection, range, payload.paddingCm).model;
  current = setCellVerticalAlign(current, selection, range, asCellVerticalAlign(payload.verticalAlign)).model;
  current = setTableAlignment(current, selection, location, asTableAlignment(payload.alignment)).model;
  current = setTableWidthMode(
    current,
    selection,
    location,
    asTableWidthMode(payload.widthMode),
    payload.widthPercent,
  ).model;
  applyEdit({ model: current, selection }, selection, "table properties");
  root.focus();
}

/** The toolbar's grid picker or the dialog (AC-1..4): inserts at the caret
 * and focuses the first cell. Inside a table the insert is refused —
 * nested tables are out of scope (SPEC-022). */
function handleInsertTable(size = { rows: 0, cols: 0 }) {
  insertTableDialogOpen.value = false;
  const active = activeHeaderFooter.value;
  const model = doc.value;
  if (active) {
    // Header/footer tables (SPEC-029 AC-5): same insertTable command, run
    // against the active band's synthetic model.
    if (!model || !editable.value) {
      return;
    }
    const selection = currentHfSelection() ?? caretSelection({ section: 0, block: 0, offset: 0 });
    if (selection.focus.cell) {
      addToast("Tables can't be inserted inside a table");
      hfRootEl()?.focus();
      return;
    }
    const synthetic = withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind);
    const result = insertTable(synthetic, selection, size.rows, size.cols);
    applyHeaderFooterEdit(result.model, result.selection, "insert table");
    hfRootEl()?.focus();
    return;
  }
  const root = contentRoot.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  if (selection.focus.cell) {
    addToast("Tables can't be inserted inside a table");
    root.focus();
    return;
  }
  applyEdit(insertTable(model, selection, size.rows, size.cols), selection, "insert table");
  root.focus();
}

// --- character formatting toolbar/shortcuts (SPEC-012, SPEC-013) -----------

/** The character-formatting toolbar's header/footer counterpart (SPEC-029
 * AC-4): identical bold/italic/… and subscript/superscript logic to
 * toggleCharacterFormat, just against the active band's synthetic model
 * and its own pending-format/caret state. */
function toggleHfCharacterFormat(key = "") {
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
  if (key === "subscript" || key === "superscript") {
    if (isCollapsed(selection)) {
      const on = verticalAlignState(synthetic, selection, key, hfPendingFormat) === "on";
      hfPendingFormat = { ...hfPendingFormat, vertAlign: on ? false : key };
    } else {
      const result = toggleVerticalAlign(synthetic, selection, key);
      applyHeaderFooterEdit(result.model, result.selection, key);
    }
    hfRootEl()?.focus();
    return;
  }
  const formatKey = asCharacterFormatKey(key);
  if (isCollapsed(selection)) {
    const next = formatState(synthetic, selection, formatKey, hfPendingFormat) !== "on";
    hfPendingFormat = { ...hfPendingFormat, [formatKey]: next };
  } else {
    const result = toggleFormat(synthetic, selection, formatKey);
    applyHeaderFooterEdit(result.model, result.selection, formatKey);
  }
  hfRootEl()?.focus();
}

/** applyRunFormatPatch's header/footer counterpart (SPEC-029 AC-4). */
function applyHfRunFormatPatch(patch = noPendingFormat(), label = "formatting") {
  const active = activeHeaderFooter.value;
  const model = doc.value;
  if (!patch || !active || !model || !editable.value) {
    return;
  }
  const selection = currentHfSelection();
  if (!selection) {
    return;
  }
  const synthetic = withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind);
  if (isCollapsed(selection)) {
    hfPendingFormat = { ...hfPendingFormat, ...patch };
  } else {
    const result = applyFormat(synthetic, selection, patch);
    applyHeaderFooterEdit(result.model, result.selection, label);
  }
  hfRootEl()?.focus();
}

/** applyParagraphCommand's header/footer counterpart (SPEC-029 AC-4): the
 * same chokepoint alignment/indent/list/style toolbar actions run through
 * for body content, redirected at the active band's synthetic model. */
function applyHfParagraphCommand(command = noParagraphCommand(), label = "formatting") {
  const active = activeHeaderFooter.value;
  const model = doc.value;
  if (!command || !active || !model || !editable.value) {
    return;
  }
  const selection = currentHfSelection() ?? caretSelection({ section: 0, block: 0, offset: 0 });
  const synthetic = withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind);
  const result = command(synthetic, selection);
  applyHeaderFooterEdit(result.model, result.selection, label);
  hfRootEl()?.focus();
}

/** toggleHfCharacterFormat's note counterpart (SPEC-031 AC-6). */
function toggleNoteCharacterFormat(key = "") {
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
  if (key === "subscript" || key === "superscript") {
    if (isCollapsed(selection)) {
      const on = verticalAlignState(synthetic, selection, key, notePendingFormat) === "on";
      notePendingFormat = { ...notePendingFormat, vertAlign: on ? false : key };
    } else {
      const result = toggleVerticalAlign(synthetic, selection, key);
      applyNoteEdit(result.model, result.selection, key);
    }
    noteRootEl()?.focus();
    return;
  }
  const formatKey = asCharacterFormatKey(key);
  if (isCollapsed(selection)) {
    const next = formatState(synthetic, selection, formatKey, notePendingFormat) !== "on";
    notePendingFormat = { ...notePendingFormat, [formatKey]: next };
  } else {
    const result = toggleFormat(synthetic, selection, formatKey);
    applyNoteEdit(result.model, result.selection, formatKey);
  }
  noteRootEl()?.focus();
}

/** applyRunFormatPatch's note counterpart (SPEC-031 AC-6). */
function applyNoteRunFormatPatch(patch = noPendingFormat(), label = "formatting") {
  const active = activeNote.value;
  const model = doc.value;
  if (!patch || !active || !model || !editable.value) {
    return;
  }
  const selection = currentNoteSelection();
  if (!selection) {
    return;
  }
  const synthetic = withNoteModel(model, active.noteId);
  if (isCollapsed(selection)) {
    notePendingFormat = { ...notePendingFormat, ...patch };
  } else {
    const result = applyFormat(synthetic, selection, patch);
    applyNoteEdit(result.model, result.selection, label);
  }
  noteRootEl()?.focus();
}

/** applyParagraphCommand's note counterpart (SPEC-031 AC-6). */
function applyNoteParagraphCommand(command = noParagraphCommand(), label = "formatting") {
  const active = activeNote.value;
  const model = doc.value;
  if (!command || !active || !model || !editable.value) {
    return;
  }
  const selection = currentNoteSelection() ?? caretSelection({ section: 0, block: 0, offset: 0 });
  const synthetic = withNoteModel(model, active.noteId);
  const result = command(synthetic, selection);
  applyNoteEdit(result.model, result.selection, label);
  noteRootEl()?.focus();
}

/** Bold/italic/underline/strikethrough (SPEC-012 AC-3..6, AC-16, AC-17)
 * and subscript/superscript (SPEC-013): a real selection is formatted
 * through the command API (which keeps the selection live, see applyEdit);
 * a collapsed caret instead sets/flips the pending format the next typed
 * text will carry (AC-9, AC-10). */
function toggleCharacterFormat(key = "") {
  if (activeHeaderFooter.value) {
    toggleHfCharacterFormat(key);
    return;
  }
  if (activeNote.value) {
    toggleNoteCharacterFormat(key);
    return;
  }
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  if (key === "subscript" || key === "superscript") {
    // One vertAlign field makes the pair mutually exclusive (SPEC-013).
    if (isCollapsed(selection)) {
      const on = verticalAlignState(model, selection, key, pendingFormat) === "on";
      pendingFormat = { ...pendingFormat, vertAlign: on ? false : key };
      updateCaretStatus({ start: selection.anchor, end: selection.focus });
    } else {
      applyEdit(toggleVerticalAlign(model, selection, key), selection, key, false);
    }
    root.focus();
    return;
  }
  const formatKey = asCharacterFormatKey(key);
  if (isCollapsed(selection)) {
    const next = formatState(model, selection, formatKey, pendingFormat) !== "on";
    pendingFormat = { ...pendingFormat, [formatKey]: next };
    updateCaretStatus({ start: selection.anchor, end: selection.focus });
  } else {
    applyEdit(toggleFormat(model, selection, formatKey), selection, formatKey, false);
  }
  root.focus();
}

/** Applies a value-format patch (font family/size, colour, highlight —
 * SPEC-013): a real selection goes through the command API; a collapsed
 * caret merges the patch into the pending format the next typed text
 * carries, exactly like the boolean toggles. */
function applyRunFormatPatch(patch = noPendingFormat(), label = "formatting") {
  if (activeHeaderFooter.value) {
    applyHfRunFormatPatch(patch, label);
    return;
  }
  if (activeNote.value) {
    applyNoteRunFormatPatch(patch, label);
    return;
  }
  const root = contentRoot.value;
  const model = doc.value;
  if (!patch || !root || !model || !editable.value) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  if (isCollapsed(selection)) {
    pendingFormat = { ...pendingFormat, ...patch };
    updateCaretStatus({ start: selection.anchor, end: selection.focus });
  } else {
    applyEdit(applyFormat(model, selection, patch), selection, label, false);
  }
  root.focus();
}

function handleFontFamilySelect(family = "") {
  if (family !== "") {
    applyRunFormatPatch({ fontFamily: family }, "font family");
  }
}

function handleFontSizeChange(sizePt = 0) {
  if (sizePt > 0) {
    applyRunFormatPatch({ fontSizePt: sizePt }, "font size");
  }
}

function handleTextColorSelect(hex = "") {
  if (hex !== "") {
    applyRunFormatPatch({ color: hex }, "text color");
    recentColors.value = pushRecentColor(hex, recentColors.value);
  }
}

/** The highlight picker emits a palette name or false for "No colour" —
 * false is applied as an explicit off, not an absence (SPEC-013). */
function handleHighlightSelect(value = noHighlight()) {
  applyRunFormatPatch({ highlight: value }, "highlight");
}

// --- paragraph formatting toolbar/ruler/shortcuts (SPEC-014) ----------------

/** The selection paragraph commands target: the live/tracked selection, or
 * — before the canvas has ever held a caret — the document start, so a
 * toolbar click or ruler drag on a fresh document still lands on the first
 * paragraph. */
function paragraphSelection(root = noElement()) {
  return currentSelection(root) ?? caretSelection(documentStart());
}

/** Runs one paragraph command against the current selection and applies its
 * result; `label` names the edit for the undo history (SPEC-015 AC-7, AC-8,
 * AC-13 — every paragraph command is its own undo step); `refocus` returns
 * focus to the canvas (skipped for ruler drags, which would tear down the
 * pointer capture mid-drag). */
function applyParagraphCommand(command = noParagraphCommand(), label = "formatting", refocus = true) {
  if (activeHeaderFooter.value) {
    applyHfParagraphCommand(command, label);
    return;
  }
  if (activeNote.value) {
    applyNoteParagraphCommand(command, label);
    return;
  }
  const root = contentRoot.value;
  const model = doc.value;
  if (!command || !root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  applyEdit(command(model, selection), selection, label, false);
  if (refocus) {
    root.focus();
  }
}

function handleAlign(alignment = "") {
  applyParagraphCommand(
    (model, selection) => setAlignment(model, selection, asParagraphAlignment(alignment)),
    "alignment",
  );
}

function handleIndent(direction = 1) {
  applyParagraphCommand(
    (model, selection) => adjustIndent(model, selection, direction * INDENT_STEP_CM),
    "indent",
  );
}

// --- paragraph styles (SPEC-019) --------------------------------------------

/** Applies a style to every paragraph in the selection (AC-3) — the picker's
 * click and the Ctrl/Cmd+Alt+0..6 shortcuts (AC-5, AC-6) both land here. An
 * empty caret paragraph has no run to carry the style's run formatting, so
 * it becomes the pending format the next typed text takes. */
function handleApplyStyle(styleId = "") {
  const root = contentRoot.value;
  const model = doc.value;
  if (styleId === "" || !root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  applyEdit(applyParagraphStyle(model, selection, styleId), selection, "paragraph style", false);
  const after = doc.value;
  const block = after ? paragraphAtCaret(after, selection.focus) : undefined;
  if (after && block && block.runs.length === 0) {
    pendingFormat = { ...resolveNamedStyle(after, styleId).runFormatting };
    updateCaretStatus({ start: selection.anchor, end: selection.focus });
  }
  root.focus();
}

/** "Update <style> to match" from a picker entry's context menu (AC-12). */
function handleUpdateStyle(styleId = "") {
  if (styleId === "") {
    return;
  }
  applyParagraphCommand(
    (model, selection) => updateStyleFromSelection(model, selection, styleId),
    "update style",
  );
}

/** "Reset style": removes direct formatting from the selection (AC-14). */
function handleResetStyle() {
  applyParagraphCommand(
    (model, selection) => resetParagraphStyle(model, selection),
    "reset style",
  );
}

// --- clear formatting and format painter (SPEC-020) -------------------------

/** The "Clear formatting" toolbar button and Ctrl/Cmd+\ (AC-1, AC-2, AC-3):
 * one command result, so — like every other paragraph command — it is one
 * undo step (AC-18). */
function handleClearFormatting() {
  applyParagraphCommand(clearFormatting, "clear formatting");
}

/** Disarms the format painter (AC-15, AC-16, AC-17): dropping the copied
 * format and the armed cursor/button state. Idempotent, so Escape can call
 * it unconditionally. */
function disarmPainter() {
  painterMode.value = "";
  painterFormat = noPainterFormat();
  lastPainterSelection = noSelection();
}

/** Arms the painter from the current caret/selection (AC-9, AC-10): copies
 * its formatting and switches the canvas cursor and the button's pressed
 * state on. */
function armPainter(mode = "once") {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  painterFormat = capturePainterFormat(model, selection);
  painterMode.value = mode;
  lastPainterSelection = noSelection();
}

/** A single click on the painter button: arms single-use mode, or disarms
 * if it was already armed — including the first of the two clicks a
 * dblclick fires, which the dblclick handler below re-arms as repeat mode
 * (AC-13, AC-14, AC-16). */
function handlePainterClick() {
  if (painterMode.value !== "") {
    disarmPainter();
    return;
  }
  armPainter("once");
}

/** A double-click on the painter button arms repeated-use mode (AC-14). */
function handlePainterDblClick() {
  armPainter("repeat");
}

/** Applies the armed painter's copied formatting to the selection just made
 * (AC-12) — the settled result of a text drag (canvas mouseup) or of a
 * single-step keyboard/click selection change (handleSelectionChange).
 * Single-use mode disarms after this one application (AC-13); repeated-use
 * stays armed for the next one. A collapsed selection is not "a selection
 * made" — nothing is applied, and the painter stays armed. `lastPainterSelection`
 * skips a selection already acted on, which is what keeps this safe to call
 * from every selectionchange without special-casing our own edits' echo. */
function maybeApplyPainter() {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || painterMode.value === "" || !painterFormat) {
    return;
  }
  const selection = currentSelection(root);
  if (!selection) {
    return;
  }
  const last = lastPainterSelection;
  if (
    last &&
    compareCaret(selection.anchor, last.anchor) === 0 &&
    compareCaret(selection.focus, last.focus) === 0
  ) {
    return;
  }
  lastPainterSelection = selection;
  if (isCollapsed(selection)) {
    return;
  }
  applyEdit(applyPainterFormat(model, selection, painterFormat), selection, "format painter", false);
  if (painterMode.value === "once") {
    disarmPainter();
  }
}

/** A line-spacing preset (Single/1.15/1.5/Double) sets the multiple and
 * clears any absolute height so the preset actually takes effect (AC-11). */
function handleSpacingPreset(multiple = 1) {
  applyParagraphCommand(
    (model, selection) =>
      applyParagraphFormat(model, selection, {
        lineSpacing: multiple,
        lineSpacingPt: null,
        lineSpacingRule: null,
      }),
    "line spacing",
  );
}

function handleToggleSpaceBefore() {
  applyParagraphCommand(toggleSpaceBefore, "space before");
}

/** The bulleted/numbered toolbar toggles (SPEC-018 AC-1, AC-2). */
function handleToggleList(type = "") {
  applyParagraphCommand(
    (model, selection) => toggleList(model, selection, asListType(type)),
    type === "bullet" ? "bulleted list" : "numbered list",
  );
}

/** A marker style picked from a list button's dropdown (SPEC-018 AC-13). */
function handleListMarker(style = "") {
  applyParagraphCommand(
    (model, selection) => setListMarker(model, selection, asListMarkerStyle(style)),
    "list marker",
  );
}

// --- multilevel lists and custom tab stops (SPEC-032) -----------------------

/** A gallery preset picked from the Multilevel list menu (AC-1..4). */
function handleMultilevelPreset(key = "") {
  applyParagraphCommand(
    (model, selection) => applyMultilevelPreset(model, selection, key),
    "multilevel list",
  );
}

/** "Number headings 1–3" (AC-13). */
function handleLinkHeadingNumbering() {
  applyParagraphCommand(
    (model, selection) => linkHeadingNumbering(model, selection),
    "number headings",
  );
}

/** "Define new multilevel list…" (AC-5..12): prefilled from the selection's
 * own definition when it sits in a defined list. */
function openDefineList() {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model) {
    return;
  }
  defineListInitial.value = selectionListDefinitionLevels(model, paragraphSelection(root));
  defineListOpen.value = true;
}

function handleDefineListApply(levels = asListLevelDefinitions([])) {
  defineListOpen.value = false;
  applyParagraphCommand(
    (model, selection) => applyListDefinition(model, selection, asListLevelDefinitions(levels)),
    "list definition",
  );
}

/** The Tabs dialog (AC-25), prefilled from the first touched paragraph. */
function openTabsDialog() {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model) {
    return;
  }
  tabsDialogInitial.value = selectionTabStops(model, paragraphSelection(root));
  tabsDialogOpen.value = true;
}

function handleTabsApply(stops = asTabStops([])) {
  tabsDialogOpen.value = false;
  applyParagraphCommand(
    (model, selection) => setTabStops(model, selection, asTabStops(stops)),
    "tab stops",
  );
}

function handleToggleDirection() {
  const rtl = paragraphState.value.direction !== "rtl";
  applyParagraphCommand(
    (model, selection) => applyParagraphFormat(model, selection, { direction: rtl ? "rtl" : "ltr" }),
    "text direction",
  );
}

function openCustomSpacing() {
  const formatting = paragraphFormatting.value;
  customSpacingInitial.value =
    formatting.lineSpacingPt !== undefined
      ? {
          mode: formatting.lineSpacingRule === "exact" ? "exact" : "atLeast",
          value: formatting.lineSpacingPt,
          beforePt: formatting.spaceBeforePt ?? 0,
          afterPt: formatting.spaceAfterPt ?? 0,
        }
      : {
          mode: "multiple",
          value: formatting.lineSpacing ?? 1,
          beforePt: formatting.spaceBeforePt ?? 0,
          afterPt: formatting.spaceAfterPt ?? 0,
        };
  customSpacingOpen.value = true;
}

/** OK in the custom-spacing dialog (AC-13): one patch, applied to every
 * touched paragraph. The unused line-spacing representation is cleared so
 * the chosen mode wins regardless of what was set before. */
function handleCustomSpacingApply(payload = { mode: "multiple", value: 1, beforePt: 0, afterPt: 0 }) {
  customSpacingOpen.value = false;
  applyParagraphCommand(
    (model, selection) =>
      payload.mode === "multiple"
        ? applyParagraphFormat(model, selection, {
            lineSpacing: payload.value,
            lineSpacingPt: null,
            lineSpacingRule: null,
            spaceBeforePt: payload.beforePt,
            spaceAfterPt: payload.afterPt,
          })
        : applyParagraphFormat(model, selection, {
            lineSpacing: null,
            lineSpacingPt: payload.value,
            lineSpacingRule: payload.mode === "exact" ? "exact" : null,
            spaceBeforePt: payload.beforePt,
            spaceAfterPt: payload.afterPt,
          }),
    "paragraph spacing",
  );
}

/** A ruler marker drag (AC-23): indent updates go to every touched
 * paragraph as logical values; margin updates go to the page margins. No
 * refocus — grabbing focus mid-drag would break the pointer capture. */
function handleRulerUpdate(
  update = { kind: "", valueCm: 0, boundary: 0, deltaCm: 0, positionCm: 0, alignment: "left", index: 0 },
) {
  const valueCm = update.valueCm;
  if (update.kind === "indentStartCm") {
    applyParagraphCommand(
      (model, selection) => applyParagraphFormat(model, selection, { indentStartCm: valueCm }),
      "indent",
      false,
    );
  } else if (update.kind === "indentEndCm") {
    applyParagraphCommand(
      (model, selection) => applyParagraphFormat(model, selection, { indentEndCm: valueCm }),
      "indent",
      false,
    );
  } else if (update.kind === "indentFirstLineCm") {
    applyParagraphCommand(
      (model, selection) => applyParagraphFormat(model, selection, { indentFirstLineCm: valueCm }),
      "indent",
      false,
    );
  } else if (update.kind === "marginLeftCm") {
    applyParagraphCommand(
      (model, selection) => setPageMargins(model, selection, { leftCm: valueCm }),
      "margins",
      false,
    );
  } else if (update.kind === "marginRightCm") {
    applyParagraphCommand(
      (model, selection) => setPageMargins(model, selection, { rightCm: valueCm }),
      "margins",
      false,
    );
  } else if (update.kind === "column-boundary") {
    const boundary = update.boundary ?? 0;
    const deltaCm = update.deltaCm ?? 0;
    applyParagraphCommand(
      (model, selection) => setColumnBoundary(model, selection, boundary, deltaCm),
      "column width",
      false,
    );
  } else if (update.kind === "add-tab") {
    // A click in the ruler's tab lane lays a stop of the armed type
    // (SPEC-032 AC-24).
    const positionCm = update.positionCm ?? 0;
    const alignment = asTabStopAlignment(update.alignment ?? "left");
    applyParagraphCommand(
      (model, selection) =>
        setTabStops(model, selection, [
          ...selectionTabStops(model, selection),
          { positionCm, alignment },
        ]),
      "tab stops",
      false,
    );
  } else if (update.kind === "remove-tab") {
    // Clicking a laid handle removes its stop (SPEC-032 AC-24).
    const index = update.index ?? 0;
    applyParagraphCommand(
      (model, selection) => {
        const stops = selectionTabStops(model, selection);
        stops.splice(index, 1);
        return setTabStops(model, selection, stops);
      },
      "tab stops",
      false,
    );
  }
}

// --- page setup and ruler visibility (SPEC-027) -----------------------

/** Opens the Page setup dialog prefilled from the document's own page
 * geometry and measurement-unit preference (AC-8; absent unit reads as
 * "cm", createDefaultStyles's own convention). */
function openPageSetup() {
  const model = doc.value;
  if (!model) {
    return;
  }
  // Prefilled from the caret's section (SPEC-028 AC-20): with per-section
  // page setup, "what the dialog shows" is the geometry under the caret.
  const sectionIndex = Math.min(caretSectionIndex.value, model.sections.length - 1);
  pageSetupInitial.value = {
    pageSize: { ...sectionPageSize(model, sectionIndex) },
    margins: { ...sectionMargins(model, sectionIndex) },
    measurementUnit: model.measurementUnit ?? "cm",
    applyTo: "wholeDocument",
  };
  pageSetupOpen.value = true;
}

/** The dialog's OK (AC-5): one combined edit — page size, margins and the
 * measurement-unit preference — applied as a single undo step, which
 * re-renders the page and re-flows text immediately (doc.value's
 * replacement is what drives pageStyle/marginStyle's recompute). */
function handlePageSetupApply(
  payload = {
    pageSize: { widthCm: 21, heightCm: 29.7 },
    margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
    measurementUnit: "cm",
    applyTo: "wholeDocument",
  },
) {
  pageSetupOpen.value = false;
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  applyEdit(
    setPageSetup(
      model,
      selection,
      {
        pageSize: payload.pageSize,
        margins: payload.margins,
        measurementUnit: asMeasurementUnit(payload.measurementUnit),
      },
      asPageSetupScope(payload.applyTo ?? "wholeDocument"),
    ),
    selection,
    "page setup",
  );
  root.focus();
}

/** Format → Columns (SPEC-030 AC-1..AC-5): prefilled from the caret's
 * section columns and its content width. */
function openColumnsDialog() {
  const model = doc.value;
  if (!model) {
    return;
  }
  const sectionIndex = Math.min(caretSectionIndex.value, model.sections.length - 1);
  columnsDialogInitial.value = {
    columns: { ...sectionColumns(model, sectionIndex) },
    contentWidthCm: sectionContentWidthCm(model, sectionIndex),
    applyTo: "thisSection",
  };
  columnsDialogOpen.value = true;
}

/** The Columns dialog's OK (AC-6/AC-7): applies the layout to the caret's
 * section or wraps the selection in a continuous section, as one undo step. */
function handleColumnsApply(
  payload = {
    columns: { count: 1, equalWidth: true, gapCm: 1.27, separator: false, widthsCm: undefined },
    applyTo: "thisSection",
  },
) {
  columnsDialogOpen.value = false;
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  applyEdit(
    applyColumns(model, selection, payload.columns, asColumnsScope(payload.applyTo ?? "thisSection")),
    selection,
    "columns",
  );
  root.focus();
}

/** The "Show/hide ruler" toggle (AC-14): a view preference, not an
 * undo-able content edit — same direct-mutation + touchDocument pattern as
 * handleTitleUpdate, and triggerRef is required for the same reason (doc is
 * a shallowRef; mutating a field on its current value alone doesn't
 * re-render). */
function toggleRuler() {
  const model = doc.value;
  if (!model) {
    return;
  }
  // Absent reads as shown (documentModel.ts's own convention), so the flip
  // is "was it explicitly false" rather than a plain `!model.showRuler`.
  model.showRuler = model.showRuler === false;
  triggerRef(doc);
  touchDocument(model.id);
}

// Every key this handler acts on is looked up in the SPEC-021 registry
// (AC-2) rather than compared against literal key/modifier checks — the
// registry in lib/shortcuts.ts is the only place that spells out which key
// combination belongs to which command. Two exceptions, both intentionally
// outside the registry (see shortcuts.ts's header comment): Escape's
// dismiss-the-current-popup behaviour, and the plain Ctrl/Cmd+V bookkeeping
// below, which observes a key the browser already owns (nothing is
// prevented) rather than binding one.
function handleKeyDown(event = new KeyboardEvent("keydown")) {
  const root = contentRoot.value;
  const model = doc.value;
  // Alt combos flow through to the registry lookup — the style shortcuts
  // (SPEC-019) bind Ctrl/Cmd+Alt+0..6; every other binding requires Alt to
  // NOT be held (exact matching), so native Alt behaviours are untouched.
  if (!root || !model || !editable.value) {
    return;
  }
  if (event.key === "Escape") {
    // Escape dismisses the paste options chip, the context menu, the
    // cell-range selection (SPEC-022) and the format painter (SPEC-020
    // AC-15, AC-17); also the selected image (SPEC-024) and header/footer
    // editing (SPEC-029 AC-7).
    dismissPasteChip();
    dismissLinkChip();
    closeContextMenu();
    clearCellSelection();
    disarmPainter();
    clearImageSelection();
    exitHeaderFooterEditing();
    exitNoteEditing();
    return;
  }
  if (imageSelection.value && (event.key === "Delete" || event.key === "Backspace")) {
    // Deleting a selected image (SPEC-024 AC-30) is a dedicated keydown
    // path, not the native beforeinput deletion flow: an image click never
    // places a DOM caret/selection inside the contenteditable region for
    // that flow to act on.
    event.preventDefault();
    handleImageDelete();
    return;
  }
  if (handleLineEdgeKey(event)) {
    return;
  }
  if (
    !event.ctrlKey &&
    !event.metaKey &&
    !event.shiftKey &&
    !event.altKey &&
    (event.key === "ArrowUp" || event.key === "ArrowDown" || event.key === "ArrowLeft" || event.key === "ArrowRight")
  ) {
    // Cell-boundary arrow movement (SPEC-022 AC-13, AC-14). A caret at the
    // exact boundary in the arrow's direction moves deterministically to
    // the neighbouring cell (or out of the table); anywhere else the
    // browser's native movement handles it, backed by a scheduled check
    // that steps in only when the native move made no progress — so the
    // caret can never be trapped in a cell. Deliberately outside the
    // SPEC-021 registry, like Escape: contextual caret behaviour, not a
    // named command.
    if (handleTableArrowKey(event)) {
      return;
    }
    scheduleTableArrowFallback(event.key);
    // Cross-column caret navigation (SPEC-030 AC-20/AC-21), the same
    // deterministic-then-fallback shape as the table arrows above.
    if (handleColumnArrowKey(event)) {
      return;
    }
    scheduleColumnArrowFallback(event.key);
  }
  if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "v") {
    // A plain Ctrl/Cmd+V clears a stale plain-paste flag (armed but never
    // consumed because an empty clipboard fired no paste event).
    plainPasteNext = false;
    return;
  }
  const shortcut = findShortcut(event);
  if (!shortcut) {
    return;
  }
  switch (shortcut.id) {
    case "insert.pageBreak": {
      // Ctrl/Cmd+Enter (AC-4): a manual page break at the caret.
      event.preventDefault();
      const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
      applyEdit(insertBreak(model, selection, "page"), selection, "insert page break");
      // The caret moved to the top of a page that is very likely below the
      // fold: follow it, or the user types onto a page they cannot see.
      void scrollModelCaretIntoView(caret, true);
      break;
    }
    case "insert.footnote": {
      // Ctrl/Cmd+Alt+F (SPEC-031 AC-1).
      event.preventDefault();
      handleInsertNote("footnote");
      break;
    }
    case "insert.endnote": {
      // Ctrl/Cmd+Alt+D (SPEC-031 AC-3).
      event.preventDefault();
      handleInsertNote("endnote");
      break;
    }
    case "view.formattingMarks": {
      // Ctrl/Cmd+Shift+P (AC-12).
      event.preventDefault();
      toggleFormattingMarks();
      break;
    }
    case "nav.documentStart":
    case "nav.documentEnd": {
      // Ctrl/Cmd+Home/End: start/end of the document, Shift extends.
      event.preventDefault();
      const toStart = shortcut.id === "nav.documentStart";
      applyCaretMove(toStart ? documentStart() : documentEnd(model), event.shiftKey);
      if (scrollContainer) {
        scrollContainer.scrollTop = toStart ? 0 : scrollContainer.scrollHeight;
      }
      break;
    }
    case "edit.undo":
      // Ctrl/Cmd+Z (SPEC-015 AC-1).
      event.preventDefault();
      performUndo();
      break;
    case "edit.redo":
      // Ctrl/Cmd+Shift+Z, or Ctrl+Y — the Windows convention (SPEC-015 AC-2, AC-3).
      event.preventDefault();
      performRedo();
      break;
    case "edit.pastePlain":
      // Ctrl/Cmd+Shift+V (SPEC-016 AC-27): arm the plain-text flag and let
      // the browser fire the paste event it produces for this combo — that
      // event hands us every clipboard flavour at once.
      plainPasteNext = true;
      schedulePlainPasteFallback();
      break;
    case "insert.nonBreakingSpace":
      // Ctrl/Cmd+Shift+Space (SPEC-026 AC-11): a direct non-breaking space
      // insert, not typed input, so autocorrect never sees it.
      event.preventDefault();
      insertSpecialCharacter("\u00A0");
      break;
    case "edit.selectAll": {
      // Ctrl/Cmd+A: select the whole document, exactly to the model's extent.
      event.preventDefault();
      pendingFormat = noPendingFormat();
      const all = selectAll(model);
      caret = all.focus;
      updateCaretStatus({ start: all.anchor, end: all.focus });
      if (
        paragraphElementForCaret(root, all.anchor) &&
        paragraphElementForCaret(root, all.focus)
      ) {
        setDomSelection(root, all.anchor, all.focus);
      } else {
        // An endpoint sits on a virtualized page: updateCaretStatus above
        // realizes both endpoint pages; select once they have rendered.
        void nextTick().then(() => setDomSelection(root, all.anchor, all.focus));
      }
      break;
    }
    case "format.bold":
    case "format.italic":
    case "format.underline":
      // Ctrl/Cmd+B/I/U (AC-3, AC-4, AC-5).
      event.preventDefault();
      toggleCharacterFormat(
        shortcut.id === "format.bold" ? "bold" : shortcut.id === "format.italic" ? "italic" : "underline",
      );
      break;
    case "format.strikethrough":
      // Ctrl/Cmd+Shift+X: strikethrough (AC-6).
      event.preventDefault();
      toggleCharacterFormat("strikethrough");
      break;
    case "format.subscript":
    case "format.superscript":
      // Ctrl/Cmd+, and Ctrl/Cmd+. : subscript / superscript (SPEC-013).
      event.preventDefault();
      toggleCharacterFormat(shortcut.id === "format.subscript" ? "subscript" : "superscript");
      break;
    case "format.clearFormatting":
      // Ctrl/Cmd+\ (SPEC-020 AC-3).
      event.preventDefault();
      handleClearFormatting();
      break;
    case "paragraph.alignLeft":
    case "paragraph.alignCenter":
    case "paragraph.alignRight":
    case "paragraph.alignJustify": {
      // Ctrl/Cmd+L/E/R/J: alignment (SPEC-014 AC-3).
      event.preventDefault();
      const alignment =
        shortcut.id === "paragraph.alignLeft"
          ? "left"
          : shortcut.id === "paragraph.alignCenter"
            ? "center"
            : shortcut.id === "paragraph.alignRight"
              ? "right"
              : "justify";
      handleAlign(alignment);
      break;
    }
    case "paragraph.indent":
    case "paragraph.outdent": {
      // SPEC-014 AC-8..10. A caret collapsed at offset 0 — an empty paragraph
      // or the start of its first line — dispatches the toolbar's own indent
      // command; a wrapped continuation line's visual start sits at a non-zero
      // offset, so it falls through to the literal tab. Always prevented:
      // the browser's default would move focus out of the editor.
      event.preventDefault();
      const selection = currentSelection(root);
      if (!selection) {
        break;
      }
      // SPEC-022 AC-10..12: inside a table cell, Tab moves to the next cell
      // and Shift+Tab to the previous — taking priority over every indent
      // and list behaviour; Tab in the last cell appends a row.
      if (isCollapsed(selection) && selection.focus.cell) {
        if (event.shiftKey) {
          const prev = prevCellCaret(model, selection.focus);
          if (prev) {
            applyCaretMove(prev);
          }
        } else {
          const next = nextCellCaret(model, selection.focus);
          if (next) {
            applyCaretMove(next);
          } else {
            const location = caretTableLocation(model, selection.focus);
            if (location) {
              applyEdit(
                appendTableRowForTab(model, selection, location),
                selection,
                "add table row",
              );
            }
          }
        }
        break;
      }
      // SPEC-018 AC-3/AC-4: inside a list item — at any offset — Tab demotes
      // and Shift+Tab promotes, taking priority over the indent behaviours.
      const touchesList = touchedBlocks(model, selection).some(
        (address) => !!paragraphAt(model, address)?.list,
      );
      if (touchesList) {
        applyParagraphCommand(
          (current, target) => changeListLevel(current, target, event.shiftKey ? -1 : 1),
          "list level",
        );
      } else if (isCollapsed(selection) && selection.focus.offset === 0) {
        applyParagraphCommand(
          (current, target) => adjustIndent(current, target, (event.shiftKey ? -1 : 1) * INDENT_STEP_CM),
          "indent",
        );
      } else if (!event.shiftKey) {
        // Any other position — mid-paragraph or a non-collapsed selection —
        // inserts a literal tab character (AC-10), grouped with typing like
        // any other inserted character (SPEC-015 AC-9).
        applyEdit(insertText(model, selection, "\t", pendingFormat), selection, "typing", true);
      }
      break;
    }
    case "insert.link":
      // Ctrl/Cmd+K opens the link dialog (SPEC-025 AC-1, AC-2).
      event.preventDefault();
      openLinkDialog();
      break;
    case "nav.pageUp":
    case "nav.pageDown":
      event.preventDefault();
      void moveCaretByPage(shortcut.id === "nav.pageDown" ? 1 : -1, event.shiftKey);
      break;
    case "style.normal":
    case "style.heading1":
    case "style.heading2":
    case "style.heading3":
    case "style.heading4":
    case "style.heading5":
    case "style.heading6": {
      // Ctrl/Cmd+Alt+1..6 apply Heading 1-6, Ctrl/Cmd+Alt+0 Normal
      // (SPEC-019 AC-5, AC-6).
      event.preventDefault();
      const styleId =
        shortcut.id === "style.normal"
          ? NORMAL_STYLE_ID
          : headingStyleId(Number(shortcut.id.slice(-1)));
      handleApplyStyle(styleId);
      break;
    }
  }
}

/** Clicking anywhere on the page canvas outside the text (margins, empty
 * space) places the caret at the nearest text position. Clicks inside the
 * editable content keep the browser's native precise placement — except a
 * read-only header/footer band (SPEC-029): it is a `contenteditable="false"`
 * island inside the editable root, so the browser treats it as an atomic
 * non-editable unit rather than falling through to nearby text the way an
 * ordinary margin click does, and needs the same explicit fallback a click
 * outside the root entirely gets. */
function handlePageMouseDown(event = new MouseEvent("mousedown")) {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  const targetEl = event.target instanceof Element ? event.target : null;
  const inReadOnlyBand = !!targetEl?.closest(
    ".page-header-band:not([data-hf-editable]), .page-footer-band:not([data-hf-editable])",
  );
  if (!inReadOnlyBand && event.target instanceof Node && root.contains(event.target)) {
    return;
  }
  const target = caretNearClientPoint(root, event.clientX, event.clientY);
  if (!target) {
    return;
  }
  event.preventDefault();
  root.focus();
  applyCaretMove(target, event.shiftKey);
}

// ---------------------------------------------------------------------------

// --- pagination (SPEC-028) --------------------------------------------------
// The controller owns measurement; every doc.value replacement re-lays-out
// synchronously (flush: "sync") so the rendered pages always agree with the
// new model's block indices before Vue patches the canvas.

const pagination = usePagination(
  () => doc.value,
  () => contentRoot.value,
  (noteId) => noteLabelsMap.value.get(noteId) ?? "",
);
const paginationPages = pagination.pages;
const paginationBusy = pagination.busy;
const paginationState = pagination.state;

// --- note numbering (SPEC-031 AC-7..9, AC-17..20) ---------------------------
// Displayed note numbers, derived fresh from mark order and the numbering
// settings on every model change — never persisted, like list numbering.
// Footnotes thread the per-page assignment through for "restart each page".
const noteLabelsMap = computed(() => {
  const model = doc.value;
  if (!model) {
    return new Map();
  }
  const refPages = pagination.footnoteRefPages.value;
  const labels = noteLabels(model, "footnote", collectNoteRefs(model, "footnote"), (ref, index) => refPages[index] ?? 0);
  for (const [noteId, label] of noteLabels(model, "endnote", collectNoteRefs(model, "endnote"))) {
    labels.set(noteId, label);
  }
  return labels;
});

/** Reference-mark tooltips (AC-14): each note's plain text. */
const noteTooltipsMap = computed(() => {
  const model = doc.value;
  const tooltips = new Map();
  for (const note of model?.notes ?? []) {
    tooltips.set(note.noteId, noteText(note));
  }
  return tooltips;
});

// A mark's rendered label width differs per number, and the measurement
// cache is keyed by block object identity — so when an existing note's
// label CHANGES (an earlier note inserted/deleted renumbers it), the
// paragraphs holding those marks must re-measure. New ids (a fresh insert)
// need none: nothing measured them yet.
let knownNoteLabels = new Map();
watch(noteLabelsMap, (labels) => {
  let changed = false;
  for (const [noteId, label] of labels) {
    const known = knownNoteLabels.get(noteId);
    if (known !== undefined && known !== label) {
      changed = true;
      break;
    }
  }
  knownNoteLabels = new Map(labels);
  if (changed) {
    pagination.invalidateAll();
  }
});

watch(doc, () => pagination.schedule(), { flush: "sync", immediate: true });
onMounted(() => pagination.schedule());
onBeforeUnmount(() => pagination.stop());

const marksOn = computed(() => doc.value?.showFormattingMarks === true);
const pageCount = computed(() => Math.max(1, paginationPages.value.length));
// The displayed page number of every laid-out page (SPEC-029 AC-20..23),
// recomputed whenever pagination or a section's numbering options change.
const pageNumbers = computed(() => {
  const model = doc.value;
  return model ? computePageNumbers(paginationPages.value, model) : [];
});
const caretPage = ref(1);
const caretSectionIndex = ref(0);
const paragraphPagination = ref(emptyParagraphPaginationState());

// --- page virtualization (SPEC-028 AC-49/50) --------------------------------
// Mounting every block of a 200-page document is what used to blow the
// settle budget: each mid-pass relayout shifted page boundaries, moving
// most paragraphs to a different DocPage, and Vue unmounted + remounted
// nearly the whole canvas several times per pass. Beyond a small page
// count only pages that matter render their content — everything near the
// viewport, the caret's / selection anchor's / current find match's page,
// and pages already realized whose layout hasn't changed (sticky, so DOM
// nodes on earlier pages survive edits elsewhere — AC-52/53) — while the
// rest render as correctly-sized empty page boxes. Their blocks are
// measured off-DOM by the pagination controller's canvas metrics, which is
// why virtualization requires `canMeasureOffscreen`.

/** Documents up to this many pages always realize fully — the exact
 * pre-virtualization behaviour for ordinary documents. */
const VIRTUALIZE_MIN_PAGES = 12;
/** DocPage's inter-page margin (AC-2's visible gap), part of page-top math. */
const PAGE_GAP_PX = 12;
/** Extra realization margin around the viewport, px. */
const VIEWPORT_OVERSCAN_PX = 400;

// Starts narrow: a large document's first render (opening lands before any
// scroll/mount refresh) must realize only the top pages, not all of them —
// refreshViewportRange corrects the range from the real viewport.
const viewportPageRange = ref({ start: 0, end: 2 });
/** Page-layout objects that have rendered their content once: they keep
 * rendering while reconciliation preserves them (an unchanged page never
 * loses its DOM), and drop out naturally when an edit replaces the page. */
const realizedSticky = new WeakSet();
/** The selection anchor's page (−1 for a collapsed caret): both ends of a
 * selection must stay rendered or the DOM selection would collapse. */
const anchorPage = ref(-1);

const virtualizing = computed(
  () =>
    pagination.canMeasureOffscreen && paginationPages.value.length > VIRTUALIZE_MIN_PAGES,
);

/** The current find match's page: realized so the highlight overlay and
 * scroll-into-view can measure it (SPEC-017). Reads paginationPages so it
 * recomputes when placements do. */
const findMatchPage = computed(() => {
  void paginationPages.value;
  const match = findMatchesRef.value[findCurrentIndex.value];
  if (!match) {
    return -1;
  }
  return pagination.pageOfCaret({ section: match.section, block: match.block, offset: match.start });
});

function refreshViewportRange() {
  const container = scrollContainer;
  const pagesEl = pageRef.value;
  const layout = paginationPages.value;
  if (!container || !pagesEl || layout.length === 0) {
    return;
  }
  const pagesTop =
    pagesEl.getBoundingClientRect().top -
    container.getBoundingClientRect().top +
    container.scrollTop;
  const viewTop = container.scrollTop - pagesTop - VIEWPORT_OVERSCAN_PX - container.clientHeight;
  const viewBottom =
    container.scrollTop - pagesTop + 2 * container.clientHeight + VIEWPORT_OVERSCAN_PX;
  let top = 0;
  let start = layout.length - 1;
  let end = 0;
  for (let index = 0; index < layout.length; index += 1) {
    const bottom = top + layout[index].pageSize.heightCm * PX_PER_CM;
    if (bottom >= viewTop && top <= viewBottom) {
      if (index < start) {
        start = index;
      }
      if (index > end) {
        end = index;
      }
    }
    top = bottom + PAGE_GAP_PX;
  }
  if (start > end) {
    start = 0;
    end = Math.min(2, layout.length - 1);
  }
  const current = viewportPageRange.value;
  if (current.start !== start || current.end !== end) {
    viewportPageRange.value = { start, end };
  }
}

let viewportRefreshHandle = 0;
function scheduleViewportRefresh() {
  if (viewportRefreshHandle !== 0 || typeof window.requestAnimationFrame !== "function") {
    return;
  }
  viewportRefreshHandle = window.requestAnimationFrame(() => {
    viewportRefreshHandle = 0;
    refreshViewportRange();
  });
}
watch(paginationPages, scheduleViewportRefresh);

// Newly realized pages render blocks that may only have estimated metrics
// (tables, images — canvas metrics cover paragraphs only): re-schedule so
// the controller measures them off their fresh DOM.
watch([viewportPageRange, caretPage], () => {
  if (!virtualizing.value) {
    return;
  }
  void nextTick().then(() => pagination.schedule());
});

function refreshCaretPage() {
  if (caret) {
    caretPage.value = pagination.pageOfCaret(caret) + 1;
  } else if (caretPage.value > pageCount.value) {
    caretPage.value = pageCount.value;
  }
}
watch(paginationPages, refreshCaretPage);

/** The "Show formatting marks" toggle (AC-12): a view preference like the
 * ruler's, persisted with the document — and a re-measure trigger, because
 * the pilcrow pseudo-element can change where full lines wrap. */
function toggleFormattingMarks() {
  const model = doc.value;
  if (!model) {
    return;
  }
  model.showFormattingMarks = model.showFormattingMarks !== true;
  triggerRef(doc);
  touchDocument(model.id);
  pagination.invalidateAll();
}

// The ruler and the Page setup dialog follow the caret's section (AC-17..21)
// so a per-section page setup is edited where the caret sits.
const rulerSectionIndex = computed(() => {
  const model = doc.value;
  return model ? Math.min(caretSectionIndex.value, model.sections.length - 1) : 0;
});
const rulerPageWidthCm = computed(() =>
  doc.value ? sectionPageSize(doc.value, rulerSectionIndex.value).widthCm : 21,
);
const rulerMarginLeftCm = computed(() =>
  doc.value ? sectionMargins(doc.value, rulerSectionIndex.value).leftCm : 2.54,
);
const rulerMarginRightCm = computed(() =>
  doc.value ? sectionMargins(doc.value, rulerSectionIndex.value).rightCm : 2.54,
);
// The caret section's column boundaries for the ruler (SPEC-030 AC-14): each
// gutter's centre, measured from the content area's start; null when the
// section has a single column (no boundary handles).
const rulerColumns = computed(() => {
  const model = doc.value;
  if (!model) {
    return undefined;
  }
  const columns = sectionColumns(model, rulerSectionIndex.value);
  if (columns.count <= 1) {
    return undefined;
  }
  const widthsCm = resolveColumnWidthsCm(columns, sectionContentWidthCm(model, rulerSectionIndex.value));
  const boundaries = [];
  let offset = 0;
  for (let i = 0; i < widthsCm.length - 1; i += 1) {
    offset += widthsCm[i];
    boundaries.push({ index: i, centerCm: offset + columns.gapCm / 2 });
    offset += columns.gapCm;
  }
  return { boundaries };
});

const SECTION_MARKER_LABELS = {
  continuous: "Section Break (Continuous)",
  nextPage: "Section Break (Next Page)",
};
const BREAK_LABELS = { page: "Page Break", column: "Column Break" };

// Typed-null seeds for renderEntryFor's espree-style defaults (never hit:
// pagedViews always passes real values).
function noDocument() {
  return createDocument("unused");
}
function noFlatEntry() {
  return flatBlocks.value[0];
}

/** One page item's precomputed render entry for DocPage — or null for the
 * continuation slice of a block whose real metrics have not arrived yet
 * (it renders whole at its first slice until measured). */
function renderEntryFor(model = noDocument(), flatEntry = noFlatEntry(), item = wholePageItem(0, 0), marks = false) {
  const block = flatEntry.block;
  const column = item.columnIndex;
  if (isBreak(block)) {
    return {
      key: flatEntry.key,
      memo: [block, flatEntry.section, flatEntry.blockIndex, marks, column],
      kind: "break",
      tag: "div",
      className: "doc-break",
      uneditable: true,
      dataSection: flatEntry.section,
      dataBlock: flatEntry.blockIndex,
      breakKind: block.kind,
      markerLabel: BREAK_LABELS[block.kind] ?? "Break",
      column,
    };
  }
  if (flatEntry.isTable || flatEntry.isImage) {
    return {
      key: flatEntry.key,
      memo: [block, flatEntry.section, flatEntry.blockIndex, flatEntry.marker, flatEntry.cellSel, flatEntry.imgSel, column, model.listDefinitions],
      kind: flatEntry.isTable ? "table" : "image",
      blockProp: block,
      sectionProp: flatEntry.section,
      blockIndexProp: flatEntry.blockIndex,
      // SPEC-032: cell paragraphs derive defined-list markers/indents from
      // the model's definitions (DocTableBlock).
      listDefinitions: flatEntry.isTable ? model.listDefinitions ?? [] : undefined,
      imageSrc: flatEntry.isImage ? flatEntry.imageSrc : undefined,
      imageSelected: flatEntry.isImage ? flatEntry.imageSelected : undefined,
      column,
    };
  }
  const paragraph = flatEntry.paragraph;
  let sliceStart;
  let sliceEnd;
  if (item.startLine !== undefined) {
    const widthPx = sectionColumnFlowWidthPx(model, item.sectionIndex);
    const offsets = pagination.isMeasured(block, widthPx) ? pagination.lineOffsetsFor(block) : null;
    if (!offsets) {
      if (item.startLine > 0) {
        return null;
      }
    } else {
      sliceStart = offsets[item.startLine] ?? 0;
      sliceEnd =
        item.endLine !== undefined && item.endLine < offsets.length
          ? offsets[item.endLine]
          : undefined;
      if (sliceStart === 0 && sliceEnd === undefined) {
        sliceStart = undefined;
      }
    }
  }
  let runs = paragraph.runs;
  if (sliceStart !== undefined || sliceEnd !== undefined) {
    const from = sliceStart ?? 0;
    const tail = splitRuns(paragraph.runs, from)[1];
    runs = sliceEnd !== undefined ? splitRuns(tail, sliceEnd - from)[0] : tail;
  }
  const styleObj = { ...paragraphStyle(paragraph, model.listDefinitions ?? []) };
  const continuation = sliceStart !== undefined && sliceStart > 0;
  if (continuation) {
    // A continuation slice starts mid-paragraph: no space-before, no
    // first-line indent, no repeated list marker.
    styleObj.marginTop = "0";
    styleObj.textIndent = "0";
  }
  if (sliceEnd !== undefined) {
    styleObj.marginBottom = "0";
  }
  const sliceSig = `${sliceStart ?? ""}:${sliceEnd ?? ""}`;
  let noteSig = "";
  const renderRuns = runs.map((run, runIndex) => {
    // A note reference mark (SPEC-031): renders its computed number as a
    // caret atom standing for the one-character model placeholder (see
    // lib/domCaret.ts), with the note's text as its tooltip (AC-14).
    if (run.note) {
      const label = noteLabelsMap.value.get(run.note.noteId) ?? run.text;
      const title = noteTooltipsMap.value.get(run.note.noteId) ?? "";
      noteSig += `${run.note.noteId}=${label}|${title};`;
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting),
        text: label,
        segments: null,
        atomText: run.text,
        noteId: run.note.noteId,
        title,
      };
    }
    const href = runLinkHref(run.formatting);
    const external = href ? isExternalLink(run.formatting) : false;
    // Segments only when the run actually holds a space/tab to mark — the
    // plain variant keeps the span's single-text-node DOM shape.
    const segments = marks ? markSegments(run.text) : null;
    const marked = segments && segments.some((segment) => segment.kind !== "text");
    return {
      key: runKey(run, runIndex),
      href: href || "",
      target: external ? "_blank" : undefined,
      rel: external ? "noopener noreferrer" : undefined,
      styleObj: runStyle(run.formatting),
      text: run.text,
      segments: marked ? segments : null,
    };
  });
  return {
    key: continuation ? `${flatEntry.key}@${sliceStart}` : flatEntry.key,
    memo: [block, flatEntry.section, flatEntry.blockIndex, flatEntry.marker, sliceSig, marks, column, noteSig],
    kind: "paragraph",
    tag: headingTag(paragraph.headingLevel),
    className: "doc-paragraph",
    styleObj,
    dataSection: flatEntry.section,
    dataBlock: flatEntry.blockIndex,
    sliceStart: continuation ? String(sliceStart) : undefined,
    sliceMore: sliceEnd !== undefined ? "1" : undefined,
    listMarker: continuation ? undefined : flatEntry.marker || undefined,
    listLevel: paragraph.list ? paragraph.list.level : undefined,
    runs: renderRuns,
    needsBr: runs.length === 0,
    column,
  };
}

/** A header/footer's content rendered as DocPage entries (SPEC-029 AC-3..6,
 * AC-20, AC-21): the same paragraph/table/image shapes renderEntryFor
 * builds for body content, minus line-slicing (header/footer content never
 * splits across pages) and with any page-number/page-count field run
 * substituted for its live value on this specific page. Addressed with
 * section 0 (dataSection/dataBlock), matching withHeaderFooterModel's
 * synthetic single-section convention — only meaningful while this exact
 * band is the active editable one; read-only mirror bands ignore it. */
function headerFooterRenderEntries(blocks = noBlocks(), pageNumber = 1, totalPages = 1, marks = false) {
  const media = doc.value?.media ?? [];
  const entries = noRenderEntries();
  blocks.forEach((block, blockIndex) => {
    if (isTable(block)) {
      entries.push({
        key: `hf:${blockIndex}`,
        memo: [block, blockIndex, marks],
        kind: "table",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: blockIndex,
      });
      return;
    }
    if (isImage(block)) {
      entries.push({
        key: `hf:${blockIndex}`,
        memo: [block, blockIndex],
        kind: "image",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: blockIndex,
        imageSrc: mediaDataUri(media.find((part) => part.mediaId === asImage(block).mediaId)),
        imageSelected: false,
      });
      return;
    }
    const paragraph = asParagraphBlock(block);
    const styleObj = { ...paragraphStyle(paragraph) };
    const renderRuns = paragraph.runs.map((run, runIndex) => {
      const text = run.field
        ? fieldValueText(run.field.kind, run.field.format, pageNumber, totalPages)
        : run.text;
      const segments = marks ? markSegments(text) : null;
      const marked = segments && segments.some((segment) => segment.kind !== "text");
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting),
        text,
        segments: marked ? segments : null,
      };
    });
    entries.push({
      key: `hf:${blockIndex}`,
      memo: [block, blockIndex, pageNumber, totalPages, marks, hfRemountEpoch.value],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: "doc-paragraph",
      styleObj,
      dataSection: 0,
      dataBlock: blockIndex,
      runs: renderRuns,
      needsBr: renderRuns.length === 0,
    });
  });
  return entries;
}

/** One note's blocks (or page slices of them) as DocPage entries (SPEC-031
 * AC-6): the paragraph/table/image shapes renderEntryFor builds, addressed
 * with section 0 (withNoteModel's synthetic convention) so the note body's
 * caret mapping works while it is the live editable surface. */
function noteBlocksEntries(note = noNoteContent(), items = noNoteAreaItems(), marks = false) {
  const entries = noRenderEntries();
  if (!note) {
    return entries;
  }
  const media = doc.value?.media ?? [];
  for (const item of items) {
    const block = note.blocks[item.blockIndex];
    if (!block) {
      continue;
    }
    if (isTable(block)) {
      entries.push({
        key: `note:${note.noteId}:${item.blockIndex}`,
        memo: [block, item.blockIndex, marks],
        kind: "table",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: item.blockIndex,
      });
      continue;
    }
    if (isImage(block)) {
      entries.push({
        key: `note:${note.noteId}:${item.blockIndex}`,
        memo: [block, item.blockIndex],
        kind: "image",
        blockProp: block,
        sectionProp: 0,
        blockIndexProp: item.blockIndex,
        imageSrc: mediaDataUri(media.find((part) => part.mediaId === asImage(block).mediaId)),
        imageSelected: false,
      });
      continue;
    }
    const paragraph = asParagraphBlock(block);
    let sliceStart;
    let sliceEnd;
    if (item.startLine !== undefined) {
      const offsets = pagination.lineOffsetsFor(block);
      if (!offsets) {
        // Unmeasured split: the first slice renders whole, continuations
        // wait for measurement (the body's own rule).
        if (item.startLine > 0) {
          continue;
        }
      } else {
        sliceStart = offsets[item.startLine] ?? 0;
        sliceEnd =
          item.endLine !== undefined && item.endLine < offsets.length
            ? offsets[item.endLine]
            : undefined;
        if (sliceStart === 0 && sliceEnd === undefined) {
          sliceStart = undefined;
        }
      }
    }
    let runs = paragraph.runs;
    if (sliceStart !== undefined || sliceEnd !== undefined) {
      const from = sliceStart ?? 0;
      const tail = splitRuns(paragraph.runs, from)[1];
      runs = sliceEnd !== undefined ? splitRuns(tail, sliceEnd - from)[0] : tail;
    }
    const continuation = sliceStart !== undefined && sliceStart > 0;
    const renderRuns = runs.map((run, runIndex) => {
      const segments = marks ? markSegments(run.text) : null;
      const marked = segments && segments.some((segment) => segment.kind !== "text");
      return {
        key: runKey(run, runIndex),
        href: "",
        styleObj: runStyle(run.formatting),
        text: run.text,
        segments: marked ? segments : null,
      };
    });
    entries.push({
      key: `note:${note.noteId}:${item.blockIndex}${continuation ? `@${sliceStart}` : ""}`,
      memo: [block, item.blockIndex, `${sliceStart ?? ""}:${sliceEnd ?? ""}`, marks, noteRemountEpoch.value],
      kind: "paragraph",
      tag: headingTag(paragraph.headingLevel),
      className: "doc-paragraph",
      styleObj: { ...paragraphStyle(paragraph) },
      dataSection: 0,
      dataBlock: item.blockIndex,
      sliceStart: continuation ? String(sliceStart) : undefined,
      sliceMore: sliceEnd !== undefined ? "1" : undefined,
      runs: renderRuns,
      needsBr: renderRuns.length === 0,
    });
  }
  return entries;
}

/** One page's footnote or endnote area as DocPage's prop shape (SPEC-031):
 * per-note groups in area order, each with its number label (on the group
 * that starts the note), its rendered entries and whether it is the live
 * editable surface. The active note renders its full content unsliced
 * where it starts; its read-only continuation slices elsewhere hide. */
function noteAreaView(model = noDocument(), area = noNoteArea(), marks = false) {
  if (!area) {
    return null;
  }
  const activeId = activeNote.value?.noteId ?? null;
  const labels = noteLabelsMap.value;
  const groups = [];
  let current = null;
  for (const item of area.items) {
    if (!current || current.noteId !== item.noteId) {
      current = { noteId: item.noteId, items: noNoteAreaItems() };
      groups.push(current);
    }
    current.items.push(item);
  }
  const notes = [];
  for (const group of groups) {
    const note = noteById(model, group.noteId);
    if (!note) {
      continue;
    }
    const first = group.items[0];
    const starts = first.blockIndex === 0 && (first.startLine === undefined || first.startLine === 0);
    const isActive = editable.value && activeId === group.noteId;
    if (isActive && !starts) {
      continue; // The active note renders whole where it starts.
    }
    const items = isActive
      ? note.blocks.map((_, blockIndex) => ({ noteId: note.noteId, blockIndex }))
      : group.items;
    notes.push({
      noteId: note.noteId,
      kind: note.kind,
      label: starts ? (labels.get(note.noteId) ?? "") : "",
      editable: isActive,
      entries: noteBlocksEntries(note, items, marks),
    });
  }
  return notes.length > 0 ? { kind: area.kind, notes } : null;
}

/** The footnote area's band geometry (SPEC-031 AC-2): pinned to the bottom
 * of the content area — the planner reserved exactly this height. */
function footnoteAreaStyle(margins = { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 }, heightPx = 0) {
  return {
    left: `${margins.leftCm}cm`,
    right: `${margins.rightCm}cm`,
    bottom: `${margins.bottomCm}cm`,
    height: `${heightPx}px`,
  };
}

/** The header/footer band geometry for a page (SPEC-029): positioned
 * within the section's own top/bottom margin so it never competes with the
 * body content box (which still starts exactly at the margin, unchanged by
 * SPEC-029). Null when nothing should render there — the read-only mirror
 * bands of an unedited, content-free header/footer stay entirely out of
 * the DOM. */
function headerFooterBandStyle(margins = { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 }, part = "header") {
  const left = `${margins.leftCm}cm`;
  const right = `${margins.rightCm}cm`;
  if (part === "header") {
    return { left, right, height: `${margins.topCm}cm`, top: "0" };
  }
  return { left, right, height: `${margins.bottomCm}cm`, bottom: "0" };
}

// The rendered pages: the layout's items joined with flatBlocks' per-block
// render data, styles resolved from each page's own geometry (AC-17..19).
// Under virtualization only realized pages build entries; the rest render
// as empty boxes with the layout's exact geometry.
const pagedViews = computed(() => {
  const model = doc.value;
  if (!model) {
    return [];
  }
  const marks = marksOn.value;
  const byAddress = new Map();
  for (const entry of flatBlocks.value) {
    byAddress.set(`${entry.section}:${entry.blockIndex}`, entry);
  }
  const lastSection = model.sections.length - 1;
  const virtual = virtualizing.value;
  const range = viewportPageRange.value;
  const caretIndex = caretPage.value - 1;
  const anchorIndex = anchorPage.value;
  const findIndex = findMatchPage.value;
  const active = activeHeaderFooter.value;
  const numbers = pageNumbers.value;
  const totalPages = paginationPages.value.length;
  const pageNoteAreas = pagination.noteAreas.value;
  return paginationPages.value.map((page, pageIndex) => {
    const realized =
      !virtual ||
      (pageIndex >= range.start && pageIndex <= range.end) ||
      pageIndex === caretIndex ||
      pageIndex === anchorIndex ||
      pageIndex === findIndex ||
      realizedSticky.has(page);
    if (!realized) {
      return {
        pageStyle: {
          width: `${page.pageSize.widthCm}cm`,
          height: `${page.pageSize.heightCm}cm`,
        },
        contentStyle: {
          padding: `${page.margins.topCm}cm ${page.margins.rightCm}cm ${page.margins.bottomCm}cm ${page.margins.leftCm}cm`,
        },
        entries: [],
        columns: page.columns,
        virtual: true,
        // The header/footer bands still need real geometry even on a
        // virtualized page (SPEC-029): they are what double-clicking the
        // margin hit-tests against (AC-1) — a virtual page still occupies
        // its full on-screen box, just without body content mounted.
        headerStyle: headerFooterBandStyle(page.margins, "header"),
        footerStyle: headerFooterBandStyle(page.margins, "footer"),
        headerEntries: [],
        footerEntries: [],
        footnoteArea: null,
        endnoteArea: null,
      };
    }
    realizedSticky.add(page);
    const entries = [];
    for (const item of page.items) {
      const flatEntry = byAddress.get(`${item.sectionIndex}:${item.blockIndex}`);
      if (!flatEntry) {
        continue;
      }
      const entry = renderEntryFor(model, flatEntry, item, marks);
      if (!entry) {
        continue;
      }
      entries.push(entry);
      const sectionBlocks = model.sections[item.sectionIndex].blocks;
      if (
        marks &&
        item.sectionIndex < lastSection &&
        item.blockIndex === sectionBlocks.length - 1 &&
        entry.sliceMore === undefined
      ) {
        const kind =
          model.sections[item.sectionIndex + 1].start === "continuous" ? "continuous" : "nextPage";
        entries.push({
          key: `${flatEntry.key}:marker`,
          memo: [kind, marks, flatEntry.key],
          kind: "marker",
          tag: "div",
          className: "doc-section-marker",
          uneditable: true,
          sectionMarker: kind,
          markerLabel: SECTION_MARKER_LABELS[kind],
          column: entry.column,
        });
      }
    }
    // Header/footer bands (SPEC-029 AC-9, AC-10): every page of a section
    // renders the same resolved content — the one page `active` addresses
    // renders it as the live editable surface, every other page read-only.
    const pageNumber = numbers[pageIndex] ?? pageIndex + 1;
    const headerEditable = !!active && active.part === "header" && active.pageIndex === pageIndex;
    const footerEditable = !!active && active.part === "footer" && active.pageIndex === pageIndex;
    const headerContent = headerFooterForPage(model, paginationPages.value, pageIndex, pageNumber, "header");
    const footerContent = headerFooterForPage(model, paginationPages.value, pageIndex, pageNumber, "footer");
    const headerBlocks = headerContent?.blocks ?? (headerEditable ? createEmptyHeaderFooterContent().blocks : null);
    const footerBlocks = footerContent?.blocks ?? (footerEditable ? createEmptyHeaderFooterContent().blocks : null);
    // Note areas (SPEC-031): the footnote band pins to the content bottom;
    // the endnote area flows after the body content (AC-2, AC-4).
    const footnotePlan = pageNoteAreas.footnotes[pageIndex] ?? null;
    const footnoteView = footnotePlan ? noteAreaView(model, footnotePlan, marks) : null;
    const endnoteView = noteAreaView(model, pageNoteAreas.endnotes[pageIndex] ?? null, marks);
    return {
      pageStyle: {
        width: `${page.pageSize.widthCm}cm`,
        height: `${page.pageSize.heightCm}cm`,
      },
      contentStyle: {
        padding: `${page.margins.topCm}cm ${page.margins.rightCm}cm ${page.margins.bottomCm}cm ${page.margins.leftCm}cm`,
      },
      entries,
      columns: page.columns,
      virtual: false,
      // Note editing does not dim the body (matching Word): the note area
      // is the only live surface, but the body stays a legible, clickable
      // reference — a body mousedown exits note editing.
      dimmed: !!active,
      headerEditable,
      footerEditable,
      headerStyle: headerFooterBandStyle(page.margins, "header"),
      footerStyle: headerFooterBandStyle(page.margins, "footer"),
      headerEntries: headerBlocks ? headerFooterRenderEntries(headerBlocks, pageNumber, totalPages, marks) : [],
      footerEntries: footerBlocks ? headerFooterRenderEntries(footerBlocks, pageNumber, totalPages, marks) : [],
      footnoteArea:
        footnotePlan && footnoteView
          ? { ...footnoteView, style: footnoteAreaStyle(page.margins, footnotePlan.heightPx) }
          : null,
      endnoteArea: endnoteView,
    };
  });
});

/** Insert → Break menu picks and Ctrl/Cmd+Enter (AC-4..9). */
function handleInsertBreak(kind = "") {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  applyEdit(insertBreak(model, selection, asBreakKind(kind)), selection, "insert break");
  root.focus();
  void scrollModelCaretIntoView(caret, true);
}

function handleInsertSectionBreak(start = "") {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = paragraphSelection(root);
  applyEdit(
    insertSectionBreak(model, selection, asSectionStart(start)),
    selection,
    "insert section break",
  );
  root.focus();
  void scrollModelCaretIntoView(caret, true);
}

/** The paragraph pagination menu's toggles (AC-22..28): each key flips
 * against the first touched paragraph's current state, applied to every
 * touched paragraph. */
function handleTogglePagination(key = "") {
  applyParagraphCommand((model, selection) => {
    const state = paragraphPaginationState(model, selection);
    if (key === "keepNext") {
      return applyParagraphFormat(model, selection, { keepNext: !state.keepNext });
    }
    if (key === "keepLines") {
      return applyParagraphFormat(model, selection, { keepLines: !state.keepLines });
    }
    if (key === "pageBreakBefore") {
      return applyParagraphFormat(model, selection, { pageBreakBefore: !state.pageBreakBefore });
    }
    return applyParagraphFormat(model, selection, { widowControl: !state.widowControl });
  }, "paragraph pagination");
}

const formatBadge = computed(() => doc.value?.format?.toUpperCase() ?? "");

// Sections flattened to one list so a single v-for can carry v-memo (Vue
// bars v-memo on elements nested inside another v-for). The memo skips
// re-rendering paragraphs whose block object is unchanged — textEditing
// replaces edited blocks — which is what keeps typing responsive on very
// large documents.
const flatBlocks = computed(() => {
  if (!doc.value) {
    return [];
  }
  // Keys carry the remount state (see `remount`): bumping it renames a key,
  // which makes Vue mount that paragraph fresh instead of patching DOM the
  // browser's IME composition may have rearranged underneath it.
  const bump = remount.value;
  // List markers are derived from the whole block sequence (SPEC-018 AC-10,
  // AC-11), so they are computed here per render rather than stored — and
  // each entry's marker joins its v-memo key, because an edit elsewhere in
  // a list renumbers items whose block object never changed.
  const markers = listMarkers(doc.value);
  // The active cell-range highlight joins table entries' v-memo keys the
  // same way (SPEC-022 AC-16): the highlighted table's block object never
  // changes when the drag extends.
  const range = cellSelection.value;
  // The selected image (SPEC-024) joins its own entry's v-memo key the same
  // way, so a wrap-mode/selection change re-renders it.
  const selectedImage = imageSelection.value;
  const media = doc.value.media ?? [];
  return doc.value.sections.flatMap((section, sectionIndex) =>
    section.blocks.map((block, blockIndex) => {
      const address = `${sectionIndex}:${blockIndex}`;
      const suffix = bump.key === address ? `:${bump.count}` : "";
      const table = isTable(block);
      const image = isImage(block);
      const cellSel =
        table && range && range.section === sectionIndex && range.block === blockIndex
          ? `${range.anchor.row}:${range.anchor.col}:${range.focus.row}:${range.focus.col}`
          : "";
      const imgSelected =
        image && !!selectedImage && selectedImage.section === sectionIndex && selectedImage.block === blockIndex;
      return {
        key: `${bump.epoch}:${address}${suffix}`,
        block,
        // The paragraph-branch view of the block — see asParagraphBlock;
        // the template's table/image branches never read it.
        paragraph: asParagraphBlock(block),
        isTable: table,
        isImage: image,
        imageSrc: image ? mediaDataUri(media.find((part) => part.mediaId === asImage(block).mediaId)) : "",
        imageSelected: imgSelected,
        section: sectionIndex,
        blockIndex,
        marker: markers[sectionIndex][blockIndex],
        cellSel,
        imgSel: imgSelected ? "1" : "",
      };
    }),
  );
});
const downloadFormat = computed(() =>
  doc.value && doc.value.format === "odt" ? "odt" : "docx",
);

function handleTitleUpdate(newTitle = "") {
  if (!doc.value) {
    return;
  }
  const trimmed = newTitle.trim();
  doc.value.title = trimmed.length > 0 ? trimmed : DEFAULT_TITLE;
  triggerRef(doc);
  // The assignment above already mutated the stored model (doc.value IS the
  // raw stored model); touching records the modification time and persists.
  touchDocument(doc.value.id);
}

async function handleDownload(format = "") {
  if (!doc.value || (format !== "docx" && format !== "odt")) {
    return;
  }
  const odt = format === "odt";
  const { bytes, dropped } = odt ? await writeOdt(doc.value) : await writeDocx(doc.value);
  const fileName = odt ? odtFileName(doc.value.title) : docxFileName(doc.value.title);
  // A transient anchor pointed at an object URL is the download mechanism
  // that needs no new browser permissions; revocation is deferred so the
  // browser has started streaming the blob before the URL disappears.
  const url = URL.createObjectURL(new Blob([bytes], { type: odt ? ODT_MIMETYPE : DOCX_MIME_TYPE }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
  // A download is SPEC-003's explicit save: promote the state to committed.
  void autosave.commitNow();
  // SPEC-008: content that could not be preserved by a cross-format save is
  // reported once, after the download has been triggered, unless the user
  // opted out for this document.
  if (dropped.length > 0 && !doc.value.suppressConversionReport) {
    conversionReport.value = dropped;
  }
}

const conversionReport = ref(noConversionReport());

function closeConversionReport(dontShowAgain = false) {
  conversionReport.value = null;
  if (dontShowAgain && doc.value) {
    doc.value.suppressConversionReport = true;
    touchDocument(doc.value.id);
  }
}

const confirmingDelete = ref(false);

function deleteOpenDocument() {
  if (!doc.value) {
    return;
  }
  const { id, title } = doc.value;
  confirmingDelete.value = false;
  deleteDocument(id);
  addToast(`Deleted ${title}`);
  router.push("/");
}

// --- find and replace (SPEC-017) -------------------------------------------
// The pure logic (compiling a query, locating matches, rewriting the model)
// lives in lib/findReplace.ts; this view owns the DOM-facing half: the find
// bar's reactive state, the on-screen highlight overlay, scrolling the current
// match into view, and applying replaces through the same applyEdit path as
// every other edit so undo/autosave treat them identically.
//
// The performance contract (AC-30, AC-31 — first results within 300 ms and no
// typing block on a 200-page document) shapes the whole design: finding
// matches is a cheap string scan, but *measuring* a DOM rectangle for every
// match is not, so the overlay only ever measures matches whose paragraph is
// on screen (plus the current one), and a document edit while the bar is open
// reschedules that work on a debounce/rAF instead of running it synchronously
// in the keystroke path.
const findVisible = ref(false);
const findExpanded = ref(false);
const findQuery = ref("");
const findReplaceText = ref("");
const findMatchCase = ref(false);
const findWholeWords = ref(false);
const findUseRegex = ref(false);
const findError = ref("");
const findHint = ref("");
const findCurrentIndex = ref(-1);
// A shallow ref of the match list and one of the overlay rectangles: shallow
// because both are wholesale-replaced, never mutated in place, and can grow
// large (thousands of matches) — a deep proxy per entry would be wasteful.
const findMatchesRef = shallowRef(noMatches());
const findHighlightRects = shallowRef(noHighlightRects());
const pageRef = useTemplateRef("pageRef");

// The caret to restore when Escape closes the bar (AC-13), a per-block index
// of the matches for viewport-limited measuring, and the scheduling handles
// that keep highlight work off the keystroke path. Plain variables — nothing
// renders from them. window.setTimeout/requestAnimationFrame return numbers in
// the DOM lib, so 0 is a safe "nothing scheduled" sentinel.
let caretBeforeFind = noCaret();
let findMatchesByBlock = new Map();
let findIdleHandle = noIdleHandle();
let findDebounceTimer = 0;
let findHintTimer = 0;

function currentFindOptions() {
  return {
    matchCase: findMatchCase.value,
    wholeWords: findWholeWords.value,
    regex: findUseRegex.value,
  };
}

/** The registry shortcuts that act regardless of where focus sits in the
 * document view — Ctrl/Cmd+F/H (which work from the canvas, not just once
 * the bar is already open) and Ctrl/Cmd+P (print) — so this is bound on
 * `document` rather than the canvas element itself. Guarded like every
 * other registry shortcut (SPEC-021 AC-12..14): none of them act while
 * focus is inside a dialog, the find bar, or another text input. */
function handleGlobalShortcut(event = new KeyboardEvent("keydown")) {
  if (!doc.value || !editable.value || event.altKey) {
    return;
  }
  // Escape disarms the format painter regardless of where focus is (AC-17)
  // — handleKeyDown below also disarms it when Escape is pressed with the
  // canvas focused, so this is deliberately unguarded by isOutsideDocumentInput.
  if (event.key === "Escape") {
    disarmPainter();
  }
  if (isOutsideDocumentInput(event.target)) {
    return;
  }
  const shortcut = findShortcut(event);
  if (shortcut?.id === "find.open") {
    event.preventDefault();
    openFind(false);
  } else if (shortcut?.id === "find.openReplace") {
    event.preventDefault();
    openFind(true);
  } else if (shortcut?.id === "file.print") {
    event.preventDefault();
    window.print();
  }
}

function openFind(expand = false) {
  if (!doc.value || !editable.value) {
    return;
  }
  if (!findVisible.value) {
    // Snapshot the caret only on a fresh open, so Escape can put it back
    // exactly where it was (AC-13) even after navigation moved the selection.
    caretBeforeFind = caret;
  }
  findVisible.value = true;
  if (expand) {
    findExpanded.value = true;
  }
  runSearch();
}

/** Escape / the close button (AC-4, AC-11, AC-12): hide the bar, clear the
 * highlights, and — for Escape — return the caret to where it was (AC-13). */
function closeFind(restoreCaret = false) {
  if (!findVisible.value) {
    return;
  }
  findVisible.value = false;
  findExpanded.value = false;
  findError.value = "";
  findHint.value = "";
  setFindMatches(noMatches(), -1);
  findHighlightRects.value = noHighlightRects();
  cancelFindScheduled();
  const root = contentRoot.value;
  if (root && editable.value) {
    root.focus();
    if (restoreCaret && caretBeforeFind) {
      caret = caretBeforeFind;
      setDomCaret(root, caret);
      updateCaretStatus({ start: caret, end: caret });
    }
  }
}

function toggleFindExpand() {
  findExpanded.value = !findExpanded.value;
}

/** Tears the bar down when the document session ends (navigation/unmount).
 * The three search options are deliberately kept — they persist for the
 * session (AC-23), across documents too. */
function resetFindState() {
  findVisible.value = false;
  findExpanded.value = false;
  findQuery.value = "";
  findReplaceText.value = "";
  findError.value = "";
  findHint.value = "";
  setFindMatches(noMatches(), -1);
  findHighlightRects.value = noHighlightRects();
  cancelFindScheduled();
}

function setFindMatches(matches = noMatches(), index = -1) {
  findMatchesRef.value = matches;
  findCurrentIndex.value = index;
  const byBlock = new Map();
  for (const match of matches) {
    const key = `${match.section}:${match.block}`;
    const list = byBlock.get(key);
    if (list) {
      list.push(match);
    } else {
      byBlock.set(key, [match]);
    }
  }
  findMatchesByBlock = byBlock;
}

/** Recomputes matches for the current query/options (AC-5's "as the user
 * types"). Cheap enough to run synchronously even on a large document (AC-30);
 * only the DOM measurement is deferred. `preserveAnchor`, when given, keeps the
 * current match near a model position across a change. */
function runSearch(preserveAnchor = noCaret()) {
  const model = doc.value;
  if (!model) {
    return;
  }
  findHint.value = "";
  const compiled = compileQuery(findQuery.value, currentFindOptions());
  if (compiled === null) {
    findError.value = "";
    setFindMatches(noMatches(), -1);
    scheduleFindMeasure();
    return;
  }
  if (isQueryError(compiled)) {
    // AC-24, AC-25: an invalid regex shows an inline error and does not run —
    // no matches, no highlights.
    findError.value = compiled.error;
    setFindMatches(noMatches(), -1);
    scheduleFindMeasure();
    return;
  }
  findError.value = "";
  const matches = findMatches(model, compiled);
  let index = matches.length > 0 ? 0 : -1;
  if (preserveAnchor && matches.length > 0) {
    index = indexAtOrAfter(matches, preserveAnchor);
  }
  setFindMatches(matches, index);
  scheduleFindMeasure();
  if (index >= 0) {
    scheduleScrollCurrentIntoView();
  }
}

function findNext() {
  stepFind(1);
}

function findPrev() {
  stepFind(-1);
}

/** Enter / Shift+Enter and the prev/next buttons (AC-3, AC-8, AC-9), wrapping
 * at either end with a transient hint (AC-10). */
function stepFind(direction = 1) {
  const matches = findMatchesRef.value;
  if (matches.length === 0) {
    return;
  }
  let index = findCurrentIndex.value;
  if (index < 0) {
    index = 0;
  } else {
    const moved = index + direction;
    if (moved < 0) {
      index = matches.length - 1;
      showFindHint("Wrapped to bottom");
    } else if (moved >= matches.length) {
      index = 0;
      showFindHint("Wrapped to top");
    } else {
      index = moved;
    }
  }
  findCurrentIndex.value = index;
  scheduleFindMeasure();
  scheduleScrollCurrentIntoView();
}

function showFindHint(text = "") {
  findHint.value = text;
  window.clearTimeout(findHintTimer);
  findHintTimer = window.setTimeout(() => {
    findHint.value = "";
  }, 1500);
}

function currentFindSelection() {
  const matches = findMatchesRef.value;
  const index = findCurrentIndex.value;
  if (index >= 0 && index < matches.length) {
    return matchSelection(matches[index]);
  }
  return caret ? caretSelection(caret) : noSelection();
}

/** The "Replace" button (AC-16): replaces the current match, taking its first
 * run's formatting (AC-26), then advances the current match to the next one so
 * repeated clicks walk the document (AC-8). Focus stays in the bar — applyEdit
 * places the DOM caret without stealing focus — so the button can be clicked
 * again immediately. */
function replaceCurrent() {
  const model = doc.value;
  const matches = findMatchesRef.value;
  const index = findCurrentIndex.value;
  if (!model || !editable.value || findError.value !== "" || index < 0 || index >= matches.length) {
    return;
  }
  const match = matches[index];
  const before = matchSelection(match);
  const anchor = {
    section: match.section,
    block: match.block,
    offset: match.start + findReplaceText.value.length,
  };
  applyEdit(replaceMatch(model, match, findReplaceText.value), before, "replace");
  refreshFindAfterEdit(anchor);
}

/** The "Replace all" button (AC-17): one edit for every match (AC-19's single
 * undo step), reporting the count (AC-18). Enabled even with no matches so it
 * can report "Replaced 0 occurrences". */
function replaceAll() {
  const model = doc.value;
  if (!model || !editable.value || findQuery.value === "" || findError.value !== "") {
    return;
  }
  const compiled = compileQuery(findQuery.value, currentFindOptions());
  if (compiled === null || isQueryError(compiled)) {
    return;
  }
  const matches = findMatches(model, compiled);
  const count = matches.length;
  if (count > 0) {
    const before = currentFindSelection() ?? caretSelection(documentStart());
    applyEdit(replaceMatches(model, matches, findReplaceText.value), before, "replace all");
  }
  addToast(`Replaced ${count} occurrences`);
  refreshFindAfterEdit(noCaret());
}

/** Recomputes matches after a model edit (a replace, or typing in the canvas
 * with the bar open), keeping the current match near `anchor`. */
function refreshFindAfterEdit(anchor = noCaret()) {
  const model = doc.value;
  if (!model || !findVisible.value) {
    return;
  }
  const compiled = compileQuery(findQuery.value, currentFindOptions());
  if (compiled === null || isQueryError(compiled)) {
    setFindMatches(noMatches(), -1);
    scheduleFindMeasure();
    return;
  }
  const matches = findMatches(model, compiled);
  const index = anchor
    ? indexAtOrAfter(matches, anchor)
    : matches.length > 0
      ? 0
      : -1;
  setFindMatches(matches, index);
  scheduleFindMeasure();
  if (index >= 0) {
    scheduleScrollCurrentIntoView();
  }
}

function currentFindAnchor() {
  const matches = findMatchesRef.value;
  const index = findCurrentIndex.value;
  if (index >= 0 && index < matches.length) {
    return matchSelection(matches[index]).anchor;
  }
  return caret;
}

/** Called from every model-edit path while the bar is open. Debounced so a
 * burst of keystrokes reschedules a single recompute rather than rescanning
 * and re-measuring on every character — the crux of AC-31 (typing stays
 * responsive even with thousands of matches highlighted). */
function scheduleFindRefresh() {
  if (!findVisible.value) {
    return;
  }
  window.clearTimeout(findDebounceTimer);
  findDebounceTimer = window.setTimeout(() => {
    refreshFindAfterEdit(currentFindAnchor());
  }, 120);
}

function scheduleFindMeasure() {
  if (findIdleHandle) {
    return;
  }
  // Highlight measurement forces a full reflow — on a 200-page canvas
  // (SPEC-028) that is the most expensive read in the find path, and the
  // match COUNT never depends on it. Idle scheduling keeps it out of the
  // keystroke-to-first-result window (SPEC-017 AC-30/31).
  findIdleHandle = scheduleIdleWork(
    () => {
      findIdleHandle = null;
      measureFindHighlights();
    },
    { timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 },
  );
}

function scheduleScrollCurrentIntoView() {
  window.requestAnimationFrame(scrollCurrentMatchIntoView);
}

function cancelFindScheduled() {
  if (findIdleHandle) {
    findIdleHandle.cancel();
    findIdleHandle = null;
  }
  window.clearTimeout(findDebounceTimer);
  window.clearTimeout(findHintTimer);
}

function findViewportBounds() {
  if (scrollContainer) {
    const rect = scrollContainer.getBoundingClientRect();
    return { top: rect.top, bottom: rect.bottom };
  }
  return { top: 0, bottom: window.innerHeight };
}

/** The set of "section:block" keys for paragraphs currently on (or just off)
 * screen — the only ones whose matches are worth measuring. Walks pages
 * first (SPEC-028): off-screen pages are skipped whole, so a 200-page
 * document reads a few hundred rects here, not one per paragraph. */
function visibleFindBlockKeys(root = noElement(), bounds = { top: 0, bottom: 0 }) {
  const keys = new Set();
  if (!root) {
    return keys;
  }
  const margin = 200;
  const pageElements = root.querySelectorAll('[data-testid="document-page"]');
  const scopes = [];
  if (pageElements.length > 0) {
    for (const pageElement of pageElements) {
      const rect = pageElement.getBoundingClientRect();
      if (rect.bottom < bounds.top - margin || rect.top > bounds.bottom + margin) {
        continue;
      }
      scopes.push(pageElement);
    }
  } else {
    scopes.push(root);
  }
  for (const scope of scopes) {
    for (const element of scope.querySelectorAll(BLOCK_SELECTOR)) {
      const rect = element.getBoundingClientRect();
      if (rect.bottom < bounds.top - margin || rect.top > bounds.bottom + margin) {
        continue;
      }
      keys.add(`${element.getAttribute("data-section")}:${element.getAttribute("data-block")}`);
    }
  }
  return keys;
}

function matchClientRects(root = noElement(), match = { section: 0, block: 0, start: 0, end: 0 }) {
  if (!root) {
    return [];
  }
  const startPoint = domPointForCaret(root, {
    section: match.section,
    block: match.block,
    offset: match.start,
  });
  const endPoint = domPointForCaret(root, {
    section: match.section,
    block: match.block,
    offset: match.end,
  });
  if (!startPoint || !endPoint) {
    return [];
  }
  const range = root.ownerDocument.createRange();
  try {
    range.setStart(startPoint.node, startPoint.offset);
    range.setEnd(endPoint.node, endPoint.offset);
  } catch {
    return [];
  }
  return Array.from(range.getClientRects());
}

/** Rebuilds the overlay rectangles, measuring only on-screen matches (plus the
 * current one). Runs off the keystroke path via rAF. */
function measureFindHighlights() {
  const root = contentRoot.value;
  const pageEl = pageRef.value;
  const matches = findMatchesRef.value;
  if (!findVisible.value || !root || !pageEl || matches.length === 0) {
    findHighlightRects.value = noHighlightRects();
    return;
  }
  const pageRect = pageEl.getBoundingClientRect();
  const visibleKeys = visibleFindBlockKeys(root, findViewportBounds());
  const currentIndex = findCurrentIndex.value;
  const rects = noHighlightRects();
  for (const [key, blockMatches] of findMatchesByBlock) {
    if (!visibleKeys.has(key)) {
      continue;
    }
    for (const match of blockMatches) {
      pushMatchRects(rects, root, pageRect, match, false, currentIndex, matches);
    }
  }
  // The current match is always drawn distinctly, even if it fell outside the
  // measured band before its scroll-into-view lands (AC-6).
  if (currentIndex >= 0 && currentIndex < matches.length) {
    const current = matches[currentIndex];
    if (!visibleKeys.has(`${current.section}:${current.block}`)) {
      pushMatchRects(rects, root, pageRect, current, true, currentIndex, matches);
    }
  }
  findHighlightRects.value = rects;
}

function pushMatchRects(
  rects = noHighlightRects(),
  root = noElement(),
  pageRect = new DOMRect(),
  match = { section: 0, block: 0, start: 0, end: 0 },
  forceCurrent = false,
  currentIndex = -1,
  matches = noMatches(),
) {
  const isCurrent = forceCurrent || matches[currentIndex] === match;
  for (const rect of matchClientRects(root, match)) {
    rects.push({
      left: rect.left - pageRect.left,
      top: rect.top - pageRect.top,
      width: rect.width,
      height: rect.height,
      current: isCurrent,
    });
  }
}

/** Scrolls the current match into view (AC-7); a following scroll event
 * re-measures the overlay for the new viewport. */
function scrollCurrentMatchIntoView() {
  const root = contentRoot.value;
  const matches = findMatchesRef.value;
  const index = findCurrentIndex.value;
  if (!root || !scrollContainer || index < 0 || index >= matches.length) {
    return;
  }
  const rects = matchClientRects(root, matches[index]);
  if (rects.length === 0) {
    return;
  }
  const rect = rects[0];
  const view = scrollContainer.getBoundingClientRect();
  const pad = 60;
  if (rect.top < view.top + pad) {
    scrollContainer.scrollTop -= view.top + pad - rect.top;
  } else if (rect.bottom > view.bottom - pad) {
    scrollContainer.scrollTop += rect.bottom - (view.bottom - pad);
  }
}

// Re-run the search whenever the query or an option changes while the bar is
// open (AC-5, AC-20..23). Keeping the current match near where it was avoids
// the view jumping to the top of the document on an option toggle.
watch([findQuery, findMatchCase, findWholeWords, findUseRegex], () => {
  if (findVisible.value) {
    runSearch(currentFindAnchor());
  }
});
</script>

<template>
  <AppShell
    v-if="doc"
    :document-title="doc.title"
    :title-editable="editable"
    :format-badge="formatBadge"
    @update:document-title="handleTitleUpdate"
  >
    <template #app-bar-actions>
      <div class="app-bar-cluster">
        <span
          class="save-state"
          data-testid="save-state"
          :data-save-state="saveState"
        >{{ saveStateText }}</span>
        <a
          v-if="saveState === 'error'"
          class="save-state-download"
          data-testid="save-state-download"
          href="#"
          @click.prevent="handleDownload(downloadFormat)"
        >Download a copy</a>
        <DownloadMenu @download="handleDownload" />
      </div>
    </template>
    <template #toolbar>
      <UndoRedoControls
        v-if="editable"
        :can-undo="undoRedoState.canUndo"
        :undo-label="undoRedoState.undoLabel"
        :can-redo="undoRedoState.canRedo"
        :redo-label="undoRedoState.redoLabel"
        @undo="performUndo"
        @redo="performRedo"
      />
      <StylePicker
        v-if="editable"
        :entries="styleEntries"
        :current-id="styleState.id"
        :current-name="styleState.name"
        @apply="handleApplyStyle"
        @update-style="handleUpdateStyle"
        @reset-style="handleResetStyle"
      />
      <FontFamilySelect
        v-if="editable"
        :value="formatValues.fontFamily"
        :document-fonts="documentFonts"
        @select="handleFontFamilySelect"
      />
      <!-- ?? undefined: the mixed state's null must reach the control as
           "prop absent" so its own null default (= blank) applies. -->
      <FontSizeControl
        v-if="editable"
        :value="formatValues.fontSizePt ?? undefined"
        @change="handleFontSizeChange"
      />
      <FormatToolbar
        v-if="editable"
        :states="formatStates"
        @toggle="toggleCharacterFormat"
      />
      <ParagraphToolbar
        v-if="editable"
        :alignment="paragraphState.alignment"
        :space-before="paragraphState.spaceBefore"
        :direction="paragraphState.direction"
        :bulleted="listState.bulleted"
        :numbered="listState.numbered"
        :list-marker="listState.marker"
        :pagination="paragraphPagination"
        @align="handleAlign"
        @indent="handleIndent"
        @spacing-preset="handleSpacingPreset"
        @custom-spacing="openCustomSpacing"
        @toggle-space-before="handleToggleSpaceBefore"
        @toggle-direction="handleToggleDirection"
        @toggle-list="handleToggleList"
        @list-marker="handleListMarker"
        @toggle-pagination="handleTogglePagination"
      />
      <MultilevelListMenu
        v-if="editable"
        @preset="handleMultilevelPreset"
        @link-headings="handleLinkHeadingNumbering"
        @define="openDefineList"
      />
      <TableInsertControl
        v-if="editable"
        @insert="handleInsertTable"
        @open-dialog="insertTableDialogOpen = true"
      />
      <InsertBreakMenu
        v-if="editable"
        @insert-break="handleInsertBreak"
        @insert-section-break="handleInsertSectionBreak"
      />
      <ImageInsertControl
        v-if="editable"
        @insert-file="handleInsertImageFile"
      />
      <HeaderFooterMenu
        v-if="editable"
        data-testid="header-footer-menu-wrapper"
        @insert-header="openInsertHeaderFooter('header')"
        @insert-footer="openInsertHeaderFooter('footer')"
        @insert-page-number="openPageNumberDialog"
        @options="openHeaderFooterOptions"
      />
      <NotesMenu
        v-if="editable"
        @insert-footnote="handleInsertNote('footnote')"
        @insert-endnote="handleInsertNote('endnote')"
        @settings="openNoteSettings"
      />
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="insert-special-character-button"
        title="Special characters"
        aria-label="Insert special character"
        @mousedown.prevent
        @click="openSpecialCharacters"
      >
        Ω
      </button>
      <AutocorrectMenu
        v-if="editable"
        :smart-quotes="autocorrectSettings.smartQuotes"
        :dashes="autocorrectSettings.dashes"
        @toggle-smart-quotes="toggleAutocorrectSmartQuotes"
        @toggle-dashes="toggleAutocorrectDashes"
      />
      <TextColorPicker
        v-if="editable"
        :value="formatValues.color"
        :mixed="formatValues.colorMixed"
        :recent="recentColors"
        @select="handleTextColorSelect"
      />
      <HighlightPicker
        v-if="editable"
        :value="formatValues.highlight"
        :mixed="formatValues.highlightMixed"
        @select="handleHighlightSelect"
      />
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="clear-formatting-button"
        :title="`Clear formatting (${shortcutLabel('format.clearFormatting')})`"
        :aria-label="`Clear formatting (${shortcutLabel('format.clearFormatting')})`"
        @mousedown.prevent
        @click="handleClearFormatting"
      >
        Clear formatting
      </button>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="format-painter-button"
        title="Format painter (double-click to keep it on)"
        aria-label="Format painter (double-click to keep it on)"
        :aria-pressed="painterMode !== ''"
        @mousedown.prevent
        @click="handlePainterClick"
        @dblclick="handlePainterDblClick"
      >
        Format painter
      </button>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="toolbar-insert-link"
        :title="`Insert link (${shortcutLabel('insert.link')})`"
        :aria-label="`Insert link (${shortcutLabel('insert.link')})`"
        @mousedown.prevent
        @click="openLinkDialog"
      >
        Link
      </button>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="toolbar-insert-bookmark"
        title="Insert bookmark"
        aria-label="Insert bookmark"
        @mousedown.prevent
        @click="openBookmarkDialog"
      >
        Bookmark
      </button>
      <button
        type="button"
        class="toolbar-button"
        data-testid="toolbar-bookmarks"
        title="Bookmarks"
        aria-label="Bookmarks"
        @click="bookmarkManagerOpen = true"
      >
        Bookmarks
      </button>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="page-setup-button"
        title="Page setup"
        aria-label="Page setup"
        @mousedown.prevent
        @click="openPageSetup"
      >
        Page setup
      </button>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="columns-button"
        title="Columns"
        aria-label="Columns"
        @mousedown.prevent
        @click="openColumnsDialog"
      >
        Columns
      </button>
      <button
        type="button"
        class="toolbar-button"
        data-testid="toggle-ruler-button"
        title="Show/hide ruler"
        aria-label="Show/hide ruler"
        :aria-pressed="doc.showRuler !== false"
        @mousedown.prevent
        @click="toggleRuler"
      >
        Ruler
      </button>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="delete-document"
        @click="confirmingDelete = true"
      >
        Delete
      </button>
    </template>
    <div
      v-if="lockState === 'busy'"
      class="notice-bar"
      role="status"
      data-testid="conflict-bar"
    >
      This document is open in another tab
    </div>
    <div
      v-if="recovery"
      class="notice-bar"
      role="status"
      data-testid="recovery-bar"
    >
      <span>Recovered unsaved changes from {{ recoveryTime }}</span>
      <button
        type="button"
        class="notice-button"
        data-testid="recovery-keep"
        @click="keepRecovered"
      >
        Keep
      </button>
      <button
        type="button"
        class="notice-button"
        data-testid="recovery-discard"
        @click="discardRecovered"
      >
        Discard
      </button>
    </div>
    <div class="page-wrap">
      <div
        v-if="findVisible"
        class="find-bar-dock"
      >
        <FindBar
          :visible="findVisible"
          :expanded="findExpanded"
          :query="findQuery"
          :replace-text="findReplaceText"
          :match-case="findMatchCase"
          :whole-words="findWholeWords"
          :use-regex="findUseRegex"
          :match-count="findMatchesRef.length"
          :current-index="findCurrentIndex"
          :error="findError"
          :hint="findHint"
          @update:query="findQuery = $event"
          @update:replace-text="findReplaceText = $event"
          @update:match-case="findMatchCase = $event"
          @update:whole-words="findWholeWords = $event"
          @update:use-regex="findUseRegex = $event"
          @toggle-expand="toggleFindExpand"
          @next="findNext"
          @prev="findPrev"
          @replace="replaceCurrent"
          @replace-all="replaceAll"
          @close="closeFind(true)"
        />
      </div>
      <DocumentRuler
        v-if="doc.showRuler !== false"
        :page-width-cm="rulerPageWidthCm"
        :margin-left-cm="rulerMarginLeftCm"
        :margin-right-cm="rulerMarginRightCm"
        :formatting="paragraphFormatting"
        :columns="rulerColumns"
        :tab-stops="paragraphFormatting.tabStops || []"
        :disabled="!editable"
        @update="handleRulerUpdate"
        @open-tabs-dialog="openTabsDialog"
      />
      <div
        ref="pageRef"
        class="pages"
        data-testid="document-pages"
        :data-pagination-state="paginationState"
        @mousedown="handlePageMouseDown"
        @mouseover="handlePageMouseOver"
      >
        <div
          v-if="findVisible"
          class="find-highlight-layer"
          data-testid="find-highlights"
          aria-hidden="true"
        >
          <div
            v-for="(rect, index) in findHighlightRects"
            :key="index"
            class="find-highlight"
            :class="{ current: rect.current }"
            :data-current="rect.current ? 'true' : 'false'"
            :style="{
              left: `${rect.left}px`,
              top: `${rect.top}px`,
              width: `${rect.width}px`,
              height: `${rect.height}px`,
            }"
          />
        </div>
        <div
          v-if="tableHover"
          class="table-handles"
          data-testid="table-handles"
        >
          <!-- Column handles above the table: an insert button at every
               column boundary, a delete button over every column centre
               (AC-20, AC-21). -->
          <button
            v-for="boundary in tableHover.cols.length + 1"
            :key="`ci${boundary}`"
            type="button"
            class="table-handle insert"
            :data-testid="`table-handle-insert-col-${boundary - 1}`"
            :aria-label="`Insert column at position ${boundary}`"
            :style="{
              left: `${(boundary - 1 < tableHover.cols.length
                ? tableHover.cols[boundary - 1].left
                : tableHover.left + tableHover.width) - 8}px`,
              top: `${tableHover.top - 20}px`,
            }"
            @mousedown.prevent
            @click="handleTableHandle('insert-col', boundary - 1)"
          >
            +
          </button>
          <button
            v-for="(col, colIndex) in tableHover.cols"
            :key="`cd${colIndex}`"
            type="button"
            class="table-handle delete"
            :data-testid="`table-handle-delete-col-${colIndex}`"
            :aria-label="`Delete column ${colIndex + 1}`"
            :style="{
              left: `${col.left + col.width / 2 - 8}px`,
              top: `${tableHover.top - 20}px`,
            }"
            @mousedown.prevent
            @click="handleTableHandle('delete-col', colIndex)"
          >
            ×
          </button>
          <!-- Row handles left of the table: same pattern. -->
          <button
            v-for="boundary in tableHover.rows.length + 1"
            :key="`ri${boundary}`"
            type="button"
            class="table-handle insert"
            :data-testid="`table-handle-insert-row-${boundary - 1}`"
            :aria-label="`Insert row at position ${boundary}`"
            :style="{
              left: `${tableHover.left - 20}px`,
              top: `${(boundary - 1 < tableHover.rows.length
                ? tableHover.rows[boundary - 1].top
                : tableHover.top + tableHover.height) - 8}px`,
            }"
            @mousedown.prevent
            @click="handleTableHandle('insert-row', boundary - 1)"
          >
            +
          </button>
          <button
            v-for="(row, rowIndex) in tableHover.rows"
            :key="`rd${rowIndex}`"
            type="button"
            class="table-handle delete"
            :data-testid="`table-handle-delete-row-${rowIndex}`"
            :aria-label="`Delete row ${rowIndex + 1}`"
            :style="{
              left: `${tableHover.left - 20}px`,
              top: `${row.top + row.height / 2 - 8}px`,
            }"
            @mousedown.prevent
            @click="handleTableHandle('delete-row', rowIndex)"
          >
            ×
          </button>
          <!-- Column/row resize strips (SPEC-023 AC-1..5): a thin draggable
               strip along each column's right edge (spanning the table's
               height) and each row's bottom edge (spanning its width). -->
          <div
            v-for="(col, colIndex) in tableHover.cols"
            :key="`cr${colIndex}`"
            class="table-resize-handle column"
            :data-testid="`table-resize-col-${colIndex}`"
            :style="{
              left: `${col.left + col.width - 3}px`,
              top: `${tableHover.top}px`,
              height: `${tableHover.height}px`,
            }"
            @mousedown="startColumnResize($event, colIndex)"
          />
          <div
            v-for="(row, rowIndex) in tableHover.rows"
            :key="`rr${rowIndex}`"
            class="table-resize-handle row"
            :data-testid="`table-resize-row-${rowIndex}`"
            :style="{
              left: `${tableHover.left}px`,
              top: `${row.top + row.height - 3}px`,
              width: `${tableHover.width}px`,
            }"
            @mousedown="startRowResize($event, rowIndex)"
          />
        </div>
        <div
          v-if="imageHandles"
          class="image-handles"
          data-testid="image-handles"
        >
          <div
            v-for="handle in IMAGE_HANDLES"
            :key="handle"
            class="image-resize-handle"
            :data-testid="`image-resize-${handle}`"
            :style="imageHandleStyle(handle)"
            @mousedown="startImageResize($event, handle)"
          />
        </div>
        <div
          ref="contentRoot"
          class="document-content"
          :class="{ 'painter-armed': painterMode !== '', 'marks-on': marksOn }"
          data-testid="document-content"
          role="textbox"
          aria-multiline="true"
          aria-label="Document content"
          :aria-readonly="editable ? undefined : 'true'"
          :contenteditable="editable && !activeHeaderFooter && !activeNote"
          spellcheck="false"
          @beforeinput="handleBeforeInput"
          @keydown="handleKeyDown"
          @compositionstart="handleCompositionStart"
          @compositionend="handleCompositionEnd"
          @copy="handleCopy"
          @cut="handleCut"
          @paste="handlePaste"
          @contextmenu="handleContextMenu"
          @mousedown="handleCanvasMouseDown"
          @mousemove="handleCanvasMouseMove"
          @mouseup="handleCanvasMouseUp"
          @click="handleContentClick"
          @dragenter="handleCanvasDragEnter"
          @dragover="handleCanvasDragOver"
          @drop="handleCanvasDrop"
        >
          <DocPage
            v-for="(pageView, pageIndex) in pagedViews"
            :key="pageIndex"
            :page-style="pageView.pageStyle"
            :content-style="pageView.contentStyle"
            :entries="pageView.entries"
            :columns="pageView.columns"
            :virtual="pageView.virtual"
            :cell-range="cellSelection || undefined"
            :dimmed="pageView.dimmed"
            :header-style="pageView.headerStyle ?? undefined"
            :footer-style="pageView.footerStyle ?? undefined"
            :header-entries="pageView.headerEntries"
            :footer-entries="pageView.footerEntries"
            :header-editable="pageView.headerEditable"
            :footer-editable="pageView.footerEditable"
            :footnote-area="pageView.footnoteArea ?? undefined"
            :endnote-area="pageView.endnoteArea ?? undefined"
            @image-select="handleImageSelect"
            @dblclick-header="handleDblclickHeader(pageIndex)"
            @dblclick-footer="handleDblclickFooter(pageIndex)"
            @dblclick-body="handleDblclickBody"
            @hf-beforeinput="handleHfBeforeInput"
            @hf-keydown="handleHfKeyDown"
            @hf-compositionstart="handleHfCompositionStart"
            @hf-compositionend="handleHfCompositionEnd"
            @note-activate="handleNoteActivate"
            @note-number-click="handleNoteNumberClick"
            @note-beforeinput="handleNoteBeforeInput"
            @note-keydown="handleNoteKeyDown"
            @note-compositionstart="handleNoteCompositionStart"
            @note-compositionend="handleNoteCompositionEnd"
          />
        </div>
      </div>
      <div
        class="visually-hidden"
        role="status"
        aria-live="polite"
        data-testid="caret-status"
      >
        {{ caretStatus }}
      </div>
    </div>
    <template #status-bar>
      <span data-testid="page-indicator">Page {{ caretPage }} of {{ pageCount }}</span>
      <span
        v-if="paginationBusy"
        data-testid="pagination-busy"
      >Paginating…</span>
    </template>
    <LinkChip
      v-if="linkChip.visible"
      :left="linkChip.left"
      :top="linkChip.top"
      :label="linkChip.label"
      :external="linkChip.external"
      @open="handleLinkChipOpen"
      @edit="handleLinkChipEdit"
      @copy="handleLinkChipCopy"
      @remove="handleLinkChipRemove"
      @dismiss="dismissLinkChip"
    />
    <LinkDialog
      v-if="linkDialogOpen"
      :initial-text="linkDialogInitial.text"
      :initial-url="linkDialogInitial.url"
      :initial-anchor="linkDialogInitial.anchor"
      :headings="documentHeadings()"
      :bookmarks="documentBookmarkTargets()"
      @save="handleLinkDialogSave"
      @cancel="handleLinkDialogCancel"
    />
    <BookmarkDialog
      v-if="bookmarkDialogOpen"
      :existing-names="documentBookmarkNames()"
      @save="handleBookmarkDialogSave"
      @cancel="handleBookmarkDialogCancel"
    />
    <BookmarkManagerDialog
      v-if="bookmarkManagerOpen"
      :names="documentBookmarkNames()"
      @jump="handleBookmarkJump"
      @remove="handleBookmarkRemove"
      @close="closeBookmarkManager"
    />
    <ConfirmDeleteDialog
      v-if="confirmingDelete"
      :document-title="doc.title"
      @confirm="deleteOpenDocument"
      @cancel="confirmingDelete = false"
    />
    <ConversionReportDialog
      v-if="conversionReport"
      :items="conversionReport"
      @close="closeConversionReport"
    />
    <CustomSpacingDialog
      v-if="customSpacingOpen"
      :initial="customSpacingInitial"
      @apply="handleCustomSpacingApply"
      @cancel="customSpacingOpen = false"
    />
    <DefineListDialog
      v-if="defineListOpen"
      :initial="defineListInitial"
      @apply="handleDefineListApply"
      @cancel="defineListOpen = false"
    />
    <TabStopsDialog
      v-if="tabsDialogOpen"
      :initial="tabsDialogInitial"
      @apply="handleTabsApply"
      @cancel="tabsDialogOpen = false"
    />
    <PasteOptionsChip
      v-if="pasteChip.visible"
      :left="pasteChip.left"
      :top="pasteChip.top"
      :applied="pasteChip.applied"
      @select="handlePasteOption"
      @dismiss="dismissPasteChip"
    />
    <ImageToolbar
      v-if="imageSelection && imageHandles"
      :left="imageToolbarStyle().left"
      :top="imageToolbarStyle().top"
      :wrap="imageAt(doc, imageSelection)?.wrap || 'inline'"
      @wrap="handleImageWrap"
      @options="openImageOptions"
      @delete="handleImageDelete"
    />
    <EditorContextMenu
      v-if="contextMenu.visible"
      :left="contextMenu.left"
      :top="contextMenu.top"
      :table-actions="!!contextMenuTableCell"
      :can-merge-cells="canMergeCells"
      :list-actions="!!contextMenuListTarget"
      @cut="contextMenuCut"
      @copy="contextMenuCopy"
      @paste="contextMenuPaste"
      @table-action="handleContextMenuTableAction"
      @list-action="handleContextMenuListAction"
      @dismiss="closeContextMenu"
    />
    <SplitCellDialog
      v-if="splitCellDialogOpen"
      @apply="handleSplitCellApply"
      @cancel="splitCellDialogOpen = false"
    />
    <TablePropertiesDialog
      v-if="tablePropertiesOpen"
      :initial="tablePropertiesInitial"
      @apply="handleTablePropertiesApply"
      @cancel="tablePropertiesOpen = false"
    />
    <PageSetupDialog
      v-if="pageSetupOpen"
      :initial="pageSetupInitial"
      @apply="handlePageSetupApply"
      @cancel="pageSetupOpen = false"
    />
    <ColumnsDialog
      v-if="columnsDialogOpen"
      :initial="columnsDialogInitial"
      @apply="handleColumnsApply"
      @cancel="columnsDialogOpen = false"
    />
    <InsertTableDialog
      v-if="insertTableDialogOpen"
      @apply="handleInsertTable"
      @cancel="insertTableDialogOpen = false"
    />
    <PageNumberDialog
      v-if="pageNumberDialogOpen"
      :initial="{ part: activeHeaderFooter?.part ?? 'header' }"
      @apply="handlePageNumberInsert"
      @cancel="handlePageNumberCancel"
    />
    <HeaderFooterOptionsDialog
      v-if="headerFooterOptionsOpen"
      :initial="headerFooterOptionsInitial"
      @apply="handleHeaderFooterOptionsApply"
      @cancel="handleHeaderFooterOptionsCancel"
    />
    <NoteSettingsDialog
      v-if="noteSettingsOpen"
      :initial="noteSettingsInitial"
      @apply="handleNoteSettingsApply"
      @cancel="handleNoteSettingsCancel"
    />
    <SpecialCharacterDialog
      v-if="specialCharacterDialogOpen"
      @insert="handleSpecialCharacterInsert"
      @cancel="specialCharacterDialogOpen = false"
    />
    <ImageOptionsDialog
      v-if="imageOptionsOpen"
      :initial="imageOptionsInitial"
      @apply="handleImageOptionsApply"
      @cancel="imageOptionsOpen = false"
    />
  </AppShell>
  <AppShell
    v-else
    document-title="No document"
  >
    <div
      class="not-found"
      data-testid="not-found"
    >
      <p class="not-found-message">
        Document not found
      </p>
      <RouterLink to="/">
        Back to home
      </RouterLink>
    </div>
  </AppShell>
</template>

<style scoped>
.toolbar-button {
  font: inherit;
  font-size: 0.875rem;
  padding: 0.25rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

/* The format painter button's armed state (SPEC-020 AC-11, AC-14). */
.toolbar-button[aria-pressed="true"] {
  background: #d2e3fc;
  border-color: #8ab4f8;
}

.app-bar-cluster {
  display: flex;
  align-items: center;
  gap: 0.75rem;
}

.save-state {
  font-size: 0.8125rem;
  color: rgb(255 255 255 / 0.85);
  white-space: nowrap;
}

.save-state-download {
  font-size: 0.8125rem;
  color: #fff;
}

/* Sticky so the conflict/recovery notice stays visible while the canvas
   scrolls; they are mutually exclusive, so a single top slot suffices. */
.notice-bar {
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  align-items: center;
  gap: 0.75rem;
  padding: 0.5rem 1rem;
  background: #e8f0fe;
  border-bottom: 1px solid #aecbfa;
  color: #174ea6;
  font-size: 0.875rem;
}

.notice-button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.125rem 0.75rem;
  border: 1px solid #aecbfa;
  border-radius: 4px;
  background: #fff;
  color: #174ea6;
  cursor: pointer;
}

.page-wrap {
  position: relative;
  padding: 2rem 1rem;
}

/* A zero-height sticky rail pinned to the top of the scrolling canvas; the
   find bar (position: absolute) floats at its top-right and stays put while
   the document scrolls. */
.find-bar-dock {
  position: sticky;
  top: 0;
  z-index: 25;
  height: 0;
}

/* The highlight overlay sits inside the (positioned) .page, behind the text —
   .page-margin renders after it, so the text paints on top and stays
   readable. pointer-events: none lets clicks reach the text underneath. */
.find-highlight-layer {
  position: absolute;
  inset: 0;
  z-index: 0;
  pointer-events: none;
}

.find-highlight {
  position: absolute;
  background: rgb(255 214 0 / 0.45);
  border-radius: 1px;
}

.find-highlight.current {
  background: rgb(255 145 0 / 0.65);
  outline: 1px solid #f57c00;
}

/* The paginated canvas (SPEC-028): one relative wrapper hosting every page
   box plus the page-relative overlays (find highlights, table/image
   handles), so overlay coordinates keep working across pages. The page
   surfaces themselves live in DocPage.vue. */
.pages {
  position: relative;
}

.document-content {
  color: #202124;
  outline: none;
}

/* The format painter's armed cursor (SPEC-020 AC-11) — a generic keyword
   cursor rather than a custom image asset, matching the rest of the app's
   cursor usage (DocumentRuler's resize handles, etc.). */
.document-content.painter-armed {
  cursor: copy;
}

/* The table hover handles (SPEC-022 AC-20, AC-21): a page-relative overlay
   of small insert (+) and delete (×) buttons above the columns and beside
   the rows of the hovered table. The layer itself ignores the pointer so
   text editing underneath is unaffected; only the buttons are clickable. */
.table-handles {
  position: absolute;
  inset: 0;
  z-index: 15;
  pointer-events: none;
}

.table-handle {
  position: absolute;
  pointer-events: auto;
  width: 16px;
  height: 16px;
  padding: 0;
  border: 1px solid #bdbdbd;
  border-radius: 50%;
  background: #fff;
  color: #3c4043;
  font-size: 11px;
  line-height: 1;
  cursor: pointer;
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

.table-handle.insert:hover {
  background: #d2e3fc;
  border-color: #1a73e8;
  color: #1a73e8;
}

.table-handle.delete:hover {
  background: #fce8e6;
  border-color: #c5221f;
  color: #c5221f;
}

/* Column/row resize strips (SPEC-023 AC-1..5): thin transparent bands
   along a column's right edge or a row's bottom edge, widened just enough
   to be a comfortable drag target without covering neighbouring cells. */
.table-resize-handle {
  position: absolute;
  pointer-events: auto;
  background: transparent;
}

.table-resize-handle.column {
  width: 6px;
  cursor: col-resize;
}

.table-resize-handle.row {
  height: 6px;
  cursor: row-resize;
}

.table-resize-handle:hover {
  background: rgb(26 115 232 / 0.35);
}

/* The selected image's eight resize handles (SPEC-024 AC-8, AC-10..12) —
   same page-relative overlay layer as .table-handles, positioned per handle
   by imageHandleStyle(). */
.image-handles {
  position: absolute;
  inset: 0;
  z-index: 15;
  pointer-events: none;
}

.image-resize-handle {
  position: absolute;
  pointer-events: auto;
  width: 8px;
  height: 8px;
  border: 1px solid #1a73e8;
  border-radius: 2px;
  background: #fff;
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

/* Off-screen but rendered, so the aria-live caret announcements reach
   assistive technology without painting anything. */
.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip-path: inset(50%);
  white-space: nowrap;
}

.not-found {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.5rem;
  color: #3c4043;
}

.not-found-message {
  margin: 0;
  font-size: 1.125rem;
}
</style>
