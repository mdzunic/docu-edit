<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import MenuBar from "../components/MenuBar.vue";
import MenuItem from "../components/MenuItem.vue";
import MenuCheckItem from "../components/MenuCheckItem.vue";
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
import SpellContextMenu from "../components/SpellContextMenu.vue";
import LanguageStatusControl from "../components/LanguageStatusControl.vue";
import CustomDictionaryDialog from "../components/CustomDictionaryDialog.vue";
import { SUPPORTED_LANGUAGES } from "../lib/language";
import HeaderFooterOptionsDialog from "../components/HeaderFooterOptionsDialog.vue";
import ImageOptionsDialog from "../components/ImageOptionsDialog.vue";
import ImageToolbar from "../components/ImageToolbar.vue";
import ShapeToolbar from "../components/ShapeToolbar.vue";
import ShapeGallery from "../components/ShapeGallery.vue";
import InsertTableDialog from "../components/InsertTableDialog.vue";
import ColumnsDialog from "../components/ColumnsDialog.vue";
import CommentsPane from "../components/CommentsPane.vue";
import TrackChangesPane from "../components/TrackChangesPane.vue";
import AccessibilityPane from "../components/AccessibilityPane.vue";
import VersionHistoryPane from "../components/VersionHistoryPane.vue";
import SpecialCharacterDialog from "../components/SpecialCharacterDialog.vue";
import LinkChip from "../components/LinkChip.vue";
import LinkDialog from "../components/LinkDialog.vue";
import TocOptionsDialog from "../components/TocOptionsDialog.vue";
import FindBar from "../components/FindBar.vue";
import FontFamilySelect from "../components/FontFamilySelect.vue";
import FontSizeControl from "../components/FontSizeControl.vue";
import FormatToolbar from "../components/FormatToolbar.vue";
import HighlightPicker from "../components/HighlightPicker.vue";
import NoteSettingsDialog from "../components/NoteSettingsDialog.vue";
import OutlinePane from "../components/OutlinePane.vue";
import PageNumberDialog from "../components/PageNumberDialog.vue";
import PageSetupDialog from "../components/PageSetupDialog.vue";
import ParagraphToolbar from "../components/ParagraphToolbar.vue";
import PasteOptionsChip from "../components/PasteOptionsChip.vue";
import PrintPreviewDialog from "../components/PrintPreviewDialog.vue";
import MergeFieldPickerDialog from "../components/MergeFieldPickerDialog.vue";
import MergeRunDialog from "../components/MergeRunDialog.vue";
import PropertiesDialog from "../components/PropertiesDialog.vue";
import RenameDialog from "../components/RenameDialog.vue";
import SplitCellDialog from "../components/SplitCellDialog.vue";
import StatisticsDialog from "../components/StatisticsDialog.vue";
import StylePicker from "../components/StylePicker.vue";
import TablePropertiesDialog from "../components/TablePropertiesDialog.vue";
import TabStopsDialog from "../components/TabStopsDialog.vue";
import TextColorPicker from "../components/TextColorPicker.vue";
import UndoRedoControls from "../components/UndoRedoControls.vue";
import { useAutosave } from "../composables/useAutosave";
import { usePagination } from "../composables/usePagination";
import { useToasts } from "../composables/useToasts";
import { useClipboard } from "./document/useClipboard";
import { asCommentFilter, useComments } from "./document/useComments";
import { asChangeEntries, useTrackChanges } from "./document/useTrackChanges";
import { useVersionHistory } from "./document/useVersionHistory";
import { useAccessibilityCheck } from "./document/useAccessibilityCheck";
import { useFindReplace } from "./document/useFindReplace";
import { usePdfExport } from "./document/usePdfExport";
import { usePrintPreview } from "./document/usePrintPreview";
import { useFormatPainter } from "./document/useFormatPainter";
import { useHeadersFooters } from "./document/useHeadersFooters";
import { useImages } from "./document/useImages";
import { useShapes } from "./document/useShapes";
import { useLinks } from "./document/useLinks";
import { useNotes } from "./document/useNotes";
import { useTables } from "./document/useTables";
import { usePaginationView } from "./document/usePaginationView";
import { usePageNavigation } from "./document/usePageNavigation";
import { useWordCount } from "./document/useWordCount";
import { useDocumentProperties } from "./document/useDocumentProperties";
import { noSpellContextMenu, noSpellUnderlineRect, useSpellCheck } from "./document/useSpellCheck";
import { useDocumentLanguage } from "./document/useDocumentLanguage";
import { documentExists } from "../lib/documents";
import {
  activeHeadingAnchor,
  blockKey,
  buildOutline,
  emptyKeySet,
  hiddenBlockAddresses,
  moveHeading,
} from "../lib/outline";
import { loadOutlinePaneOpen, saveOutlinePaneOpen } from "../lib/outlinePane";
import { DOCX_MIME_TYPE, docxFileName, writeDocx } from "../lib/docx/writer";
import { ODT_MIMETYPE, odtFileName, writeOdt } from "../lib/odf/writer";
import { MD_MIME_TYPE, mdFileName, writeMd } from "../lib/md/writer";
import { noConversionReport } from "../lib/opaque";
import { convertLiteralChipsToRuns, hasConvertibleLiteralChips } from "../lib/mergeChips";
import { namedArrayFieldList, namedFieldList, noParsedDataSource } from "../lib/merge";
import {
  deleteDocument,
  getOrCreateDocument,
  replaceDocumentModel,
  touchDocument,
} from "../lib/documentStore";
import { saveAsTemplate } from "../lib/templateStore";
import {
  asBorderLineStyle,
  asCellVerticalAlign,
  asCharacterFormatKey,
  asHeaderFooterPart,
  asImage,
  asImageWrapMode,
  asShape,
  asShapeGeometry,
  asListLevelDefinitions,
  asListMarkerStyle,
  asListType,
  asMeasurementUnit,
  asNoteKind,
  asNoteNumberFormat,
  asNoteRestart,
  asTabStopAlignment,
  asTabStops,
  asTrackChangesDisplayMode,
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
  isShape,
  isTable,
  isToc,
  asToc,
  asTocOptions,
  noActiveHeaderFooter,
  noActiveNote,
  noBlocks,
  noNamedStyles,
  noNoteContent,
  resolveColumnWidthsCm,
  sectionColumns,
  sectionContentWidthCm,
  sectionMargins,
  sectionPageSize,
} from "../lib/documentModel";
import {
  defaultTocOptions,
  insertTableOfContents,
  noTocSelection,
  tocHeadingsStale,
  tocPageNumbersStale,
  updateTocEntirely,
  updateTocPageNumbersOnly,
} from "../lib/toc";
import { detectListAutoformat, listMarkers } from "../lib/lists";
import { GALLERY_PRESETS } from "../lib/listDefinitions";
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
  PICKER_MAX_ROWS,
  PICKER_MAX_COLS,
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
import { ceilingWarningMessage, checkSupportedCeiling, withinCeiling } from "../lib/supportedCeiling";
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
import { applyCasePattern, detectCasePattern } from "../lib/spellcheck/suggestions";
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
  applyCharacterStyle,
  applyFormat,
  applyListAutoformat,
  applyListDefinition,
  applyMultilevelPreset,
  applyPainterFormat,
  applyParagraphFormat,
  applyParagraphStyle,
  applyTableStyle,
  asPageSetupScope,
  createStyle,
  deleteStyle,
  modifyStyle,
  noStyleFormInput,
  toolbarCharacterStyleState,
  continueListNumbering,
  blockCount,
  blockOrdinal,
  capturePainterFormat,
  caretFormatting,
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
  insertMergeChip,
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
  ACCEPTED_IMAGE_ACCEPT_ATTR,
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
import { asNudgeDirection, noShapeLocation, shapeAt } from "../lib/shapes";
import {
  asImageHandle,
  IMAGE_HANDLES,
  noImageResizeDrag,
  previewImageSize,
  resizedImageSize,
} from "../lib/imageResize";
import {
  allStyles,
  headingStyleId,
  isStyleInUse,
  NORMAL_STYLE_ID,
  noStyleEntries,
  noStylesPaneEntries,
  resolveNamedStyle,
  styleById,
  stylePickerEntries,
  stylesPaneEntries,
  styleTypeOf,
} from "../lib/styles";
import {
  importStyles,
  noStyleImportPlan,
  noStyleImportResolutionMap,
  noStyleImportSummary,
  planStyleImport,
  readStylesFromFile,
  styleImportConflictRows,
  styleImportResolutionsFromMap,
} from "../lib/customStyles";
import StylesPane from "../components/StylesPane.vue";
import StyleEditorDialog from "../components/StyleEditorDialog.vue";
import StyleDeleteDialog from "../components/StyleDeleteDialog.vue";
import ImportStylesDialog from "../components/ImportStylesDialog.vue";
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
import { applyTrackChanges } from "../lib/trackChangesDiff";
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
// Whether this tab owns the document, regardless of version-history preview
// — the version history pane's own actions (restore, name, delete) stay
// enabled while previewing; only the canvas itself goes read-only.
const documentActionsEnabled = computed(() => lockState.value === "acquired" && !!doc.value);

// --- supported ceiling (SPEC-047: a document past the ceiling opens
// read-only rather than the editor pretending it can hold arbitrarily large
// content) -------------------------------------------------------------
// Counted once per document, at the door ("counted after parse and before
// pagination" — the spec's own definition), NOT re-checked on every edit: a
// live re-check would permanently lock a document the moment one paste
// carried it over the limit, with undo — the way back under — disabled by
// the very same flag. A document that grows past the ceiling mid-session
// stays editable until it is next opened; the door check catches it then.
// (Evaluating per-document rather than per-edit also keeps the O(blocks)
// count walk off the keystroke path.)
const ceilingResult = ref(withinCeiling());
watch(
  doc,
  (model, prev) => {
    if (!model) {
      ceilingResult.value = withinCeiling();
      return;
    }
    if (prev && prev.id === model.id) {
      return; // the same document being edited — the door check already ran
    }
    ceilingResult.value = checkSupportedCeiling(model);
  },
  { immediate: true },
);
const pastSupportedCeiling = computed(() => ceilingResult.value.exceeded);
const ceilingWarning = computed(() => ceilingWarningMessage(ceilingResult.value));

// Find and print are read commands (SPEC-047 AC-27: "opening, scrolling,
// find, print, PDF export and download still work" while a too-large
// document is read-only) — this stays what `editable` used to mean, before
// the ceiling factored in, for the one handler (handleGlobalShortcut) that
// must keep reacting to Ctrl/Cmd+F/H/P over a too-large document.
const findAndPrintEnabled = computed(
  () => documentActionsEnabled.value && !versionHistoryPreviewActive.value,
);
const editable = computed(() => findAndPrintEnabled.value && !pastSupportedCeiling.value);

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

// --- feature lifecycle hooks ------------------------------------------------
// Extracted feature composables register callbacks the backbone runs at the
// points they care about (session close, unmount, scroll, post-edit find
// refresh), so the backbone need not import feature modules and there are no
// import cycles. Registration callbacks are idempotent per feature.
// No TS annotations here — ESLint parses .vue scripts with espree, which
// cannot parse them (see the other espree notes in this file). A dummy
// no-op keeps `findRefreshScheduler` function-typed via inference.
// The callbacks are typed by inference — a `() => {}` seed makes each Set a
// `Set<() => void>` and each param a `() => void` (see note above: espree).
const sessionCloseHooks = new Set([() => {}]);
const unmountHooks = new Set([() => {}]);
const scrollHooks = new Set([() => {}]);
let findRefreshScheduler = () => {};

function onSessionClose(cb = () => {}) {
  sessionCloseHooks.add(cb);
}
function onUnmount(cb = () => {}) {
  unmountHooks.add(cb);
}
function onScroll(cb = () => {}) {
  scrollHooks.add(cb);
}
function registerFindRefresh(fn = () => {}) {
  findRefreshScheduler = fn;
}

const contextMenu = ref({ visible: false, left: 0, top: 0 });
const spellContextMenu = ref(noSpellContextMenu());

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
const repeatRowDialogOpen = ref(false);
let tableFormattingLocation = { section: 0, block: 0, row: 0, col: 0 };
// The column/row resize drag state now lives inside useTables (it owns the
// live-resize handlers and their listener lifecycle).




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
// --- custom styles manager (SPEC-042) ---------------------------------------
// The caret's active character style (AC-10's second readout, alongside
// styleState above) and the Styles pane/dialogs' own state.
const characterStyleState = ref(emptyToolbarStyleState());
const stylesPaneOpen = ref(false);
const stylesPaneEntriesList = computed(() => (doc.value ? stylesPaneEntries(doc.value) : noStylesPaneEntries()));
const styleEditorOpen = ref(false);
const styleEditorMode = ref("new");
const styleEditorType = ref("paragraph");
const styleEditorTargetId = ref("");
const styleDeleteTargetId = ref("");
const styleDeleteTargetInUse = ref(false);
const importDialogOpen = ref(false);
const importPhase = ref("pick");
const importFileName = ref("");
const importError = ref("");
const importSourceStyles = ref(noNamedStyles());
const importPlanState = ref(noStyleImportPlan());
const importResolutions = ref(noStyleImportResolutionMap());
const importSummaryState = ref(noStyleImportSummary());
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
const pageRef = useTemplateRef("pageRef");

// The format painter (SPEC-020) lives in useFormatPainter; handleSelectionChange
// and the canvas mouseup dispatch into its maybeApplyPainter.
const {
  painterMode,
  disarmPainter,
  handlePainterClick,
  handlePainterDblClick,
  maybeApplyPainter,
} = useFormatPainter({
  doc,
  contentRoot,
  editable,
  currentSelection,
  applyEdit,
});

// --- headers, footers and page numbers (SPEC-029) ----------------------------
// Lives in useHeadersFooters (the one live-editable band + page-number dialogs).
// It runs before images/notes — which consume its `activeHeaderFooter` /
// `applyHeaderFooterEdit` — so the mutual deps (clearImageSelection,
// exitNoteEditing, and the later pagination refs) come in as lazy getters.
const {
  activeHeaderFooter,
  hfRemountEpoch,
  headerFooterOptionsOpen,
  headerFooterOptionsInitial,
  pageNumberDialogOpen,
  currentHfSelection,
  applyHeaderFooterEdit,
  hfRootEl,
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
} = useHeadersFooters({
  doc,
  contentRoot,
  editable,
  documentId,
  currentSelection,
  applyEdit,
  getPaginationPages: () => paginationPages.value,
  getPageNumbers: () => pageNumbers.value,
  getCaretSectionIndex: () => caretSectionIndex.value,
  pageOfCaret: (caretPos) => pagination.pageOfCaret(caretPos),
  clearCellSelection,
  getClearImageSelection: () => clearImageSelection,
  getExitNoteEditing: () => exitNoteEditing,
  getCaret: () => caret,
  markDirty: () => noteEdit(),
  refreshUndoRedoState,
});

// --- images (SPEC-024) -------------------------------------------------------
// Image selection/handles/options and insertion + resize live in useImages; the
// canvas and applyEdit reach them through these returned bindings.
const {
  imageSelection,
  imageHandles,
  imageOptionsOpen,
  imageOptionsInitial,
  clearImageSelection,
  handleImageSelect,
  imageHandleStyle,
  imageToolbarStyle,
  remeasureImageHandles,
  handleImageWrap,
  handleImageDelete,
  openImageOptions,
  handleImageOptionsApply,
  insertImageFromFile,
  handleInsertImageFile,
  startImageResize,
  handleImageResizeMove,
  finishImageResize,
} = useImages({
  doc,
  contentRoot,
  pageRef,
  editable,
  currentSelection,
  applyEdit,
  clearCellSelection,
  addToast,
  activeHeaderFooter,
  currentHfSelection,
  applyHeaderFooterEdit,
  getCaret: () => caret,
});

// --- shapes and text boxes (SPEC-051) ---------------------------------------
// Shape selection/handles/toolbar and draw-mode insertion + resize/rotate
// live in useShapes; the canvas and applyEdit reach them through these
// returned bindings, the same shape useImages follows.
const {
  shapeSelection,
  shapeMultiSelection,
  shapeHandles,
  shapeAnchorMarker,
  drawGeometry,
  editingShapeText,
  shapeTextRemountEpoch,
  clearShapeSelection,
  handleShapeSelect,
  handleShapeEditText,
  exitShapeTextEditing,
  handleShapeTextBeforeInput,
  handleShapeTextKeyDown,
  handleShapeTextCompositionStart,
  handleShapeTextCompositionEnd,
  shapeAtClientPoint,
  canGroupSelection,
  shapeHandleStyle,
  rotationHandleStyle,
  shapeToolbarStyle,
  remeasureShapeHandles,
  handleShapeDelete,
  handleShapeDuplicate,
  handleShapeNudge,
  handleShapeFill,
  handleShapeOutline,
  handleShapeRotationInput,
  handleShapeWrap,
  handleShapeAltText,
  handleShapeZOrder,
  handleGroupSelected,
  handleUngroupSelected,
  currentGroupMembers,
  startShapeResize,
  startShapeMove,
  startShapeRotate,
  armShapeDraw,
  disarmShapeDraw,
  startShapeDraw,
  insertTextBoxAction,
} = useShapes({
  doc,
  contentRoot,
  pageRef,
  editable,
  documentId,
  currentSelection,
  applyEdit,
  clearCellSelection,
  getCaret: () => caret,
  markDirty: () => noteEdit(),
  refreshUndoRedoState,
});

// Footnotes and endnotes (SPEC-031) live in useNotes; it provides `activeNote` /
// `noteRemountEpoch` that usePaginationView renders the note areas from.
const {
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
} = useNotes({
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
  getCaret: () => caret,
  markDirty: () => noteEdit(),
  refreshUndoRedoState,
});

// Cut/copy/paste and the paste-options chip (SPEC-016) live in useClipboard;
// the plain-paste flag and chip state it owns are wired into handleKeyDown and
// the session-close hook (resetClipboard) here.
const {
  pasteChip,
  handleCopy,
  handleCut,
  handlePaste,
  handlePasteOption,
  dismissPasteChip,
  armPlainPaste,
  clearPlainPaste,
  performPaste,
  getLastCopy,
} = useClipboard({
  doc,
  contentRoot,
  editable,
  documentId,
  currentSelection,
  applyEdit,
  addToast,
  getCaret: () => caret,
  getPendingFormat: () => pendingFormat,
  insertImageFromFile,
  onSessionClose,
});

// Live word count and the statistics dialog (SPEC-036) live in useWordCount;
// updateCaretStatus below feeds it every selection change, and its own
// throttled watch on `doc` keeps the document-wide total live.
const {
  statusBarText: wordCountStatusText,
  announcement: wordCountAnnouncement,
  setSelection: setWordCountSelection,
  statisticsOpen,
  includeNotes: statisticsIncludeNotes,
  dialogStats: statisticsStats,
  refreshDialogStats: refreshStatisticsStats,
  openStatistics,
  closeStatistics,
} = useWordCount({ doc, onUnmount });

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
  onCleanClose: () => noteVersionHistoryCommitPoint(),
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
      beginVersionHistorySession();
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
  sessionCloseHooks.forEach((cb) => cb());
  clearCellSelection();
  cellDragOrigin = null;
  tableHover.value = null;
  insertTableDialogOpen.value = false;
  specialCharacterDialogOpen.value = false;
  producedEnDashAt = -1;
  producedOpenSingleAt = -1;
  // The clipboard's own state (plain-paste flag, chip, payload) is reset by the
  // clipboard composable's session-close hook, not here.
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
  // Feature composables re-measure scroll-sensitive overlays (e.g. the find
  // highlight overlay) by registering a scroll hook — see onScroll.
  scrollHooks.forEach((cb) => cb());
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
  unmountHooks.forEach((cb) => cb());
  document.removeEventListener("selectionchange", handleSelectionChange);
  document.removeEventListener("keydown", handleGlobalShortcut);
  // Column/row resize listener cleanup is owned by useTables' onUnmount hook.
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
  skipTrackChanges = false,
) {
  if (!result || !beforeSelection) {
    return;
  }
  // The read-only backstop for a too-large document (SPEC-047 AC-27):
  // `editable` gates the canvas, the toolbar and the shortcut registry, but
  // several composables reach applyEdit through their own UI (the status
  // bar's language control, the accessibility pane's Fix buttons, …) — one
  // gate here covers every such path, present and future.
  if (pastSupportedCeiling.value) {
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
  clearTocSelection();
  const before = doc.value;
  // SPEC-039: with recording on, rewrite the command's result into tracked-
  // change markup before it ever reaches undo history or the model — every
  // command becomes trackable with no per-command wiring (see
  // trackChangesDiff.ts's header comment). `skipTrackChanges` is for the
  // accept/reject commands themselves, whose result IS the final state a
  // reviewer chose, not a new edit to re-diff.
  if (before && !skipTrackChanges && isRecording()) {
    result = {
      model: applyTrackChanges(before, result.model, trackChangesAuthorName.value, new Date().toISOString()),
      selection: result.selection,
    };
  }
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
  noteEdit();
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
  findRefreshScheduler?.();
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
  noteEdit();
  updateCaretStatus({ start: snapshot.selection.anchor, end: snapshot.selection.focus });
  refreshUndoRedoState();
  root.focus();
  void restoreCaretAfterRender();
  // Undo/redo can move the caret a long way — undoing an inserted page
  // break pulls it back a whole page — so follow it, the way the break
  // insertion it reverses does.
  void scrollModelCaretIntoView(caret);
  findRefreshScheduler?.();
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

/** Whether a caret's own block address (ignoring any cell path) is a TOC
 * atom (SPEC-033) — used to catch every edit attempt at its boundary before
 * the ordinary per-inputType dispatch below. */
function caretAtToc(model = noDocument(), caret = noCaret()) {
  if (!caret || caret.cell) {
    return false;
  }
  return isToc(model.sections[caret.section]?.blocks[caret.block]);
}

/** Whether a keydown would type or delete content — a plain printable
 * character, or Enter/Backspace/Delete without a modifier — the set
 * handleKeyDown's selected-TOC guard (SPEC-033 AC-15) intercepts. Excludes
 * navigation, Tab and every modifier combo, so Ctrl/Cmd+C (copy) and arrow
 * keys around a selected TOC still work normally. */
function isContentKey(event = new KeyboardEvent("keydown")) {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return false;
  }
  if (event.key === "Enter" || event.key === "Backspace" || event.key === "Delete") {
    return true;
  }
  return event.key.length === 1;
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
  // A TOC is not directly editable (SPEC-033 AC-14, AC-15, AC-24): a caret
  // sitting at its atom boundary (either end of the selection) shows the
  // update hint instead of silently doing nothing, the way Image/BreakBlock
  // atoms already do for every other kind of edit.
  if (caretAtToc(model, selection.anchor) || caretAtToc(model, selection.focus)) {
    addToast('This table of contents isn’t editable — use "Update table of contents" instead.');
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


// --- hyperlinks and bookmarks (SPEC-025) -----------------------------------

/** The dialog's heading targets (AC-19, AC-21), re-read from the current
 * model each time the dialog template evaluates them. */











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








// --- the editor's context menu (SPEC-016 AC-4..6) --------------------------

/** Right-click on the canvas: the app's own menu replaces the native one so
 * its entries can drive the clipboard pipeline (and tests can reach them). */
function handleContextMenu(event = new MouseEvent("contextmenu")) {
  if (!doc.value || !editable.value) {
    return;
  }
  // A right-click on a misspelled word opens the suggestions menu instead of
  // this one (SPEC-035 AC-2). The underline layer is inert, so the click
  // arrives here rather than on the overlay — see spellRectAtPoint.
  const spellRect = spellRectAtPoint(event.clientX, event.clientY);
  if (spellRect) {
    handleSpellContextMenu(event, spellRect);
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

// --- the spell-check context menu (SPEC-035 AC-2, AC-3) ---------------------

/** The underline rect under a point, in client coordinates, or null when the
 * point is not over a flagged word. The underline layer cannot do this hit
 * test itself: its rects cover whole glyph boxes and paint above the text, so
 * making them pointer-interactive would swallow ordinary clicks, double-clicks
 * and drags (see the .spell-underline rule). The rects are stored relative to
 * the .pages box — the same space useSpellCheck measured them in. */
function spellRectAtPoint(clientX = 0, clientY = 0) {
  const pageEl = pageRef.value;
  if (!pageEl) {
    return null;
  }
  const pageRect = pageEl.getBoundingClientRect();
  const x = clientX - pageRect.left;
  const y = clientY - pageRect.top;
  for (const rect of spellUnderlineRects.value) {
    if (x >= rect.left && x <= rect.left + rect.width && y >= rect.top && y <= rect.top + rect.height) {
      return rect;
    }
  }
  return null;
}

/** Right-click on a spell-check underline: opens the suggestions menu
 * instead of the regular editing menu. Reached from handleContextMenu, which
 * hit-tests the click, because the underline layer itself is inert. */
function handleSpellContextMenu(event = new MouseEvent("contextmenu"), rect = noSpellUnderlineRect()) {
  if (!editable.value) {
    return;
  }
  event.preventDefault();
  spellContextMenu.value = {
    visible: true,
    left: event.clientX,
    top: event.clientY,
    section: rect.section,
    block: rect.block,
    start: rect.start,
    end: rect.end,
    word: rect.word,
    lang: rect.lang,
    suggestions: suggestionsFor(rect.word, rect.lang),
  };
}

function closeSpellContextMenu() {
  spellContextMenu.value = noSpellContextMenu();
}

/** Choosing a suggestion (SPEC-035 AC-4, AC-4.1): replaceMatch preserves
 * the word's character formatting; applyCasePattern reshapes the
 * suggestion to the misspelling's own capitalisation pattern first. */
function applySpellSuggestion(suggestion = "") {
  const state = spellContextMenu.value;
  closeSpellContextMenu();
  const model = doc.value;
  if (!model || !editable.value) {
    return;
  }
  const match = { section: state.section, block: state.block, start: state.start, end: state.end };
  const shaped = applyCasePattern(suggestion, detectCasePattern(state.word));
  const before = caretSelection({ ...match, offset: match.start });
  applyEdit(replaceMatch(model, match, shaped), before, "spelling suggestion");
}

/** "Ignore" (SPEC-035 AC-5): this occurrence only, for the session. */
function spellIgnore() {
  const state = spellContextMenu.value;
  closeSpellContextMenu();
  ignoreOccurrence({
    section: state.section,
    block: state.block,
    start: state.start,
    end: state.end,
    word: state.word,
    lang: state.lang,
  });
}

/** "Ignore all" (SPEC-035 AC-6): every occurrence of the word, for the
 * session. */
function spellIgnoreAll() {
  const state = spellContextMenu.value;
  closeSpellContextMenu();
  ignoreAllOccurrences(state.word);
}

/** "Add to dictionary" (SPEC-035 AC-7): persists across sessions. */
function spellAddToDictionary() {
  const state = spellContextMenu.value;
  closeSpellContextMenu();
  addWordToDictionary(state.word, state.lang);
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
  const ownIfLastCopy = getLastCopy();
  const own =
    ownIfLastCopy && flavours.html && flavours.html === ownIfLastCopy.html
      ? ownIfLastCopy.fragment
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
  setWordCountSelection(range.start, range.end);
  setLanguageSelection(range.start, range.end);
  // The caret's focus end drives the outline's active-heading highlight
  // (SPEC-037 AC-3, AC-4) — a reactive mirror of the plain `caret` variable.
  caretPosition.value = range.end;
  if (!model || !range.start || !range.end) {
    formatStates.value = EMPTY_FORMAT_STATES;
    formatValues.value = emptyToolbarFormatValues();
    paragraphState.value = EMPTY_PARAGRAPH_STATE;
    paragraphFormatting.value = emptyParagraphFormatting();
    listState.value = emptyToolbarListState();
    styleState.value = emptyToolbarStyleState();
    characterStyleState.value = emptyToolbarStyleState();
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
  // SPEC-042 AC-10: the character style readout, alongside the paragraph
  // style readout above — both recomputed here so pure caret movement (not
  // just an edit) keeps the Styles pane's caret readout live.
  characterStyleState.value = toolbarCharacterStyleState(model, selection);
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

// Deterministic caret navigation (page up/down, line Home/End, cell-boundary
// and column-boundary arrow escapes) lives in usePageNavigation; handleKeyDown
// dispatches into its handlers.
const {
  moveCaretByPage,
  handleLineEdgeKey,
  handleTableArrowKey,
  scheduleTableArrowFallback,
  handleColumnArrowKey,
  scheduleColumnArrowFallback,
} = usePageNavigation({
  doc,
  contentRoot,
  editable,
  currentSelection,
  applyCaretMove,
  getCaret: () => caret,
  getScrollContainer: () => scrollContainer,
  refreshViewportRange,
});

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
  // Draw mode (SPEC-051 AC-3, AC-4): a gallery pick armed drawGeometry, so
  // this mousedown draws the shape instead of doing anything else the
  // canvas would normally do with a click.
  if (drawGeometry.value) {
    startShapeDraw(event);
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
  // Same for a selected shape (SPEC-051) — a click on one is handled by
  // DocShapeBlock's own @select, which fires afterward. An unfilled,
  // textless shape's proxy is pointer-events: none over its whole box (AC-11
  // — its own interior is not part of its painted region), so `event.target`
  // never lands inside it; shapeAtClientPoint recovers that case by testing
  // every shape's actual painted region (outline included) at this point.
  const shapeTarget = event.target instanceof Element ? event.target.closest(".doc-shape-anchor") : null;
  let shapeMouseDownLocation = noShapeLocation();
  if (shapeTarget instanceof HTMLElement) {
    const section = Number(shapeTarget.dataset.section);
    const block = Number(shapeTarget.dataset.block);
    if (Number.isInteger(section) && Number.isInteger(block)) {
      shapeMouseDownLocation = { section, block };
    }
  } else {
    const recovered = shapeAtClientPoint(event.clientX, event.clientY);
    if (recovered) {
      handleShapeSelect(recovered.section, recovered.block, event.shiftKey);
      shapeMouseDownLocation = recovered;
    } else {
      clearShapeSelection();
    }
  }
  // A drag on the shape's own body moves (and, crossing a paragraph
  // boundary, re-anchors) it (AC-28) — armed here rather than from
  // DocShapeBlock's own click handler so it also covers a shape that was not
  // already selected (mousedown's native hit-test target is the same
  // element a plain click would land on) and the AC-11 recovery case above.
  // Shift is reserved for the multi-select toggle (AC-34), not a drag.
  if (shapeMouseDownLocation && !event.shiftKey) {
    startShapeMove(event, shapeMouseDownLocation);
  }
  // Same for a selected TOC (SPEC-033) — a click on one is handled by
  // DocTocBlock's own @action "select", which fires afterward.
  const tocTarget = event.target instanceof Element ? event.target.closest(".doc-toc") : null;
  if (!tocTarget) {
    clearTocSelection();
  }
}

/** Dblclick recovery for a shape DocShapeBlock's own proxy couldn't catch
 * natively (AC-11's pointer-events: none case, same reasoning as
 * handleCanvasMouseDown's shapeAtClientPoint call) — otherwise AC-17's
 * "double-clicking a shape puts the caret inside it, populating Shape.text
 * if absent" would be unreachable by mouse for an unfilled, textless shape
 * whose only clickable region is its outline. */
function handleCanvasDblClick(event = new MouseEvent("dblclick")) {
  const shapeTarget = event.target instanceof Element ? event.target.closest(".doc-shape-anchor") : null;
  if (shapeTarget) {
    return;
  }
  const recovered = shapeAtClientPoint(event.clientX, event.clientY);
  if (recovered) {
    handleShapeEditText(recovered.section, recovered.block);
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

// --- tables (SPEC-022, SPEC-023) ---------------------------------------------
// Row/column insert-delete, merge/split, the properties dialog, and live column
// /row resize live in useTables; the hover handles' geometry and the cell-range
// drag selection stay with the canvas event hub above.
const {
  applyTableAction,
  handleTableHandle,
  startColumnResize,
  handleColumnResizeMove,
  finishColumnResize,
  startRowResize,
  handleRowResizeMove,
  finishRowResize,
  handleContextMenuTableAction,
  handleMergeCells,
  openSplitCellDialog,
  handleSplitCellApply,
  openTableProperties,
  handleTablePropertiesApply,
  handleInsertTable,
  confirmRepeatRow,
} = useTables({
  doc,
  contentRoot,
  editable,
  currentSelection,
  paragraphSelection,
  applyEdit,
  addToast,
  activeHeaderFooter,
  currentHfSelection,
  applyHeaderFooterEdit,
  hfRootEl,
  closeContextMenu,
  tableHover,
  cellSelection,
  insertTableDialogOpen,
  splitCellDialogOpen,
  tablePropertiesOpen,
  repeatRowDialogOpen,
  tablePropertiesInitial,
  getTableFormattingLocation: () => tableFormattingLocation,
  setTableFormattingLocation: (location) => {
    tableFormattingLocation = location;
  },
  getContextMenuTableCell: () => contextMenuTableCell,
  onUnmount,
});

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
      const on = verticalAlignState(synthetic, selection, key, getHfPendingFormat()) === "on";
      setHfPendingFormat({ ...getHfPendingFormat(), vertAlign: on ? false : key });
    } else {
      const result = toggleVerticalAlign(synthetic, selection, key);
      applyHeaderFooterEdit(result.model, result.selection, key);
    }
    hfRootEl()?.focus();
    return;
  }
  const formatKey = asCharacterFormatKey(key);
  if (isCollapsed(selection)) {
    const next = formatState(synthetic, selection, formatKey, getHfPendingFormat()) !== "on";
    setHfPendingFormat({ ...getHfPendingFormat(), [formatKey]: next });
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
    setHfPendingFormat({ ...getHfPendingFormat(), ...patch });
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
      const on = verticalAlignState(synthetic, selection, key, getNotePendingFormat()) === "on";
      setNotePendingFormat({ ...getNotePendingFormat(), vertAlign: on ? false : key });
    } else {
      const result = toggleVerticalAlign(synthetic, selection, key);
      applyNoteEdit(result.model, result.selection, key);
    }
    noteRootEl()?.focus();
    return;
  }
  const formatKey = asCharacterFormatKey(key);
  if (isCollapsed(selection)) {
    const next = formatState(synthetic, selection, formatKey, getNotePendingFormat()) !== "on";
    setNotePendingFormat({ ...getNotePendingFormat(), [formatKey]: next });
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
    setNotePendingFormat({ ...getNotePendingFormat(), ...patch });
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

// --- custom styles manager (SPEC-042) ---------------------------------------

function toggleStylesPane() {
  stylesPaneOpen.value = !stylesPaneOpen.value;
}

function closeStylesPane() {
  stylesPaneOpen.value = false;
}

/** "New style" (AC-3): seeded from the caret's current formatting, and — for
 * a paragraph/list style — its current paragraph style as the "based on"
 * starting point. */
function openNewStyle(type = "paragraph") {
  styleEditorMode.value = "new";
  styleEditorType.value = type;
  styleEditorTargetId.value = "";
  styleEditorOpen.value = true;
}

/** "Modify style" (AC-4). */
function openModifyStyle(styleId = "") {
  const model = doc.value;
  const target = model ? styleById(model, styleId) : undefined;
  styleEditorMode.value = "modify";
  styleEditorType.value = target ? styleTypeOf(target) : "paragraph";
  styleEditorTargetId.value = styleId;
  styleEditorOpen.value = true;
}

function closeStyleEditor() {
  styleEditorOpen.value = false;
}

/** The style being modified, or an empty style for "new" (fields default to
 * the caret's own formatting via styleEditorInitialProps below). */
const styleEditorInitialProps = computed(() => {
  const model = doc.value;
  if (styleEditorMode.value === "modify") {
    const target = model ? styleById(model, styleEditorTargetId.value) : undefined;
    const border = target?.cellBorders?.top;
    return {
      initialName: target?.name ?? "",
      initialType: target ? styleTypeOf(target) : "paragraph",
      initialBasedOn: target?.basedOn ?? "",
      initialNext: target?.next ?? "",
      initialFontFamily: target?.runFormatting?.fontFamily ?? "",
      initialFontSizePt: target?.runFormatting?.fontSizePt,
      initialBold: target?.runFormatting?.bold === true,
      initialItalic: target?.runFormatting?.italic === true,
      initialUnderline: target?.runFormatting?.underline === true,
      initialColor: target?.runFormatting?.color ?? "",
      initialAlignment: target?.paragraphFormatting?.alignment ?? "",
      initialListPresetKey: "decimal",
      initialCellShadingFill: target?.cellShadingFill ?? "",
      initialCellBorderStyle: border?.style ?? "",
      initialCellBorderWidthPt: border?.widthPt ?? 1,
      initialCellBorderColor: border?.color ?? "#000000",
    };
  }
  const caretNow = caretPosition.value;
  const caretRun = model && caretNow ? caretFormatting(model, caretNow) : {};
  const caretParagraph = model && caretNow ? paragraphAtCaret(model, caretNow) : undefined;
  return {
    initialName: "",
    initialType: styleEditorType.value,
    initialBasedOn: caretParagraph?.styleId ?? "",
    initialNext: "",
    initialFontFamily: caretRun.fontFamily ?? "",
    initialFontSizePt: caretRun.fontSizePt,
    initialBold: caretRun.bold === true,
    initialItalic: caretRun.italic === true,
    initialUnderline: caretRun.underline === true,
    initialColor: caretRun.color ?? "",
    initialAlignment: caretParagraph?.formatting?.alignment ?? "",
    initialListPresetKey: "decimal",
    initialCellShadingFill: "",
    initialCellBorderStyle: "",
    initialCellBorderWidthPt: 1,
    initialCellBorderColor: "#000000",
  };
});

const styleEditorBasedOnOptions = computed(() => {
  const model = doc.value;
  if (!model) {
    return [];
  }
  return allStyles(model)
    .filter((entry) => entry.name !== undefined && entry.id !== styleEditorTargetId.value)
    .map((entry) => ({ id: entry.id, name: entry.name ?? entry.id }));
});

const styleEditorExistingNames = computed(() => {
  const model = doc.value;
  if (!model) {
    return [];
  }
  return allStyles(model)
    .filter((entry) => entry.name !== undefined && entry.id !== styleEditorTargetId.value)
    .map((entry) => (entry.name ?? entry.id).trim().toLowerCase());
});

function saveStyleEditor(payload = noStyleFormInput()) {
  if (styleEditorMode.value === "new") {
    applyParagraphCommand((model, selection) => createStyle(model, selection, payload), "new style");
  } else {
    applyParagraphCommand(
      (model, selection) => modifyStyle(model, selection, styleEditorTargetId.value, payload),
      "modify style",
    );
  }
  styleEditorOpen.value = false;
}

/** "Delete style" (AC-7): the pane's own Delete button opens the confirm
 * dialog directly — the in-use warning (AC-8) is computed from the live
 * document here, not left to the dialog to guess. */
function openDeleteStyle(styleId = "") {
  const model = doc.value;
  styleDeleteTargetId.value = styleId;
  styleDeleteTargetInUse.value = model ? isStyleInUse(model, styleId) : false;
}

function cancelDeleteStyle() {
  styleDeleteTargetId.value = "";
}

function confirmDeleteStyle() {
  applyParagraphCommand(
    (model, selection) => deleteStyle(model, selection, styleDeleteTargetId.value),
    "delete style",
  );
  styleDeleteTargetId.value = "";
}

const styleDeleteTargetName = computed(() => {
  const model = doc.value;
  const target = model ? styleById(model, styleDeleteTargetId.value) : undefined;
  return target?.name ?? styleDeleteTargetId.value;
});

/** Applying a style from the pane (AC-9): character/table styles use their
 * own commands; paragraph/list styles reuse the picker's own apply path
 * (which also links a list style's numbering, SPEC-032's styleLinkFor). */
function applyStyleFromPane(payload = { id: "", type: "paragraph" }) {
  if (payload.type === "character") {
    applyParagraphCommand(
      (model, selection) => applyCharacterStyle(model, selection, payload.id),
      "character style",
    );
  } else if (payload.type === "table") {
    applyParagraphCommand((model, selection) => applyTableStyle(model, selection, payload.id), "table style");
  } else {
    handleApplyStyle(payload.id);
  }
}

// --- import styles from another document (SPEC-042 AC-19..24) --------------

function openImportDialog() {
  importDialogOpen.value = true;
  importPhase.value = "pick";
  importFileName.value = "";
  importError.value = "";
  importSourceStyles.value = noNamedStyles();
  importPlanState.value = noStyleImportPlan();
  importResolutions.value = noStyleImportResolutionMap();
  importSummaryState.value = noStyleImportSummary();
}

function closeImportDialog() {
  importDialogOpen.value = false;
}

const importConflictRows = computed(() =>
  importPlanState.value ? styleImportConflictRows(importPlanState.value.conflicts) : [],
);
const importNonConflictingCount = computed(() => importPlanState.value?.nonConflicting.length ?? 0);

/** Reads the chosen file's styles and plans the merge (AC-19, AC-20, AC-21). */
async function handleImportChooseFile(file = new File([], "")) {
  const model = doc.value;
  if (!model) {
    return;
  }
  importError.value = "";
  try {
    const styles = await readStylesFromFile(file);
    importFileName.value = file.name;
    importSourceStyles.value = styles;
    importPlanState.value = planStyleImport(model, styles);
    importResolutions.value = noStyleImportResolutionMap();
    importPhase.value = "plan";
  } catch (error) {
    importError.value = error instanceof Error ? error.message : "Could not read styles from that file";
  }
}

function setImportResolution(payload = { sourceId: "", resolution: "keep" }) {
  importResolutions.value = { ...importResolutions.value, [payload.sourceId]: payload.resolution };
}

/** Applies the plan (AC-22, AC-23) and shows the summary (AC-24). */
function confirmImportStyles() {
  const model = doc.value;
  const root = contentRoot.value;
  if (!model || !root) {
    return;
  }
  const selection = paragraphSelection(root);
  const resolutions = styleImportResolutionsFromMap(importResolutions.value);
  const result = importStyles(model, importSourceStyles.value, resolutions);
  applyEdit({ model: result.model, selection }, selection, "import styles", false);
  importSummaryState.value = result.summary;
  importPhase.value = "summary";
}

// --- mail-merge: convert literal {{...}} text (SPEC-046) --------------------

/** Whether "Convert {{…}} text to merge fields" has anything to do — offered
 * when a document holds at least one well-formed `{{…}}` literal outside an
 * existing chip (AC-20, "Literal text is not a chip"). */
const canConvertLiteralChips = computed(() => (doc.value ? hasConvertibleLiteralChips(doc.value) : false));

/** Scans the whole document and converts every literal `{{…}}` match into a
 * chip in one undo step — explicit and undoable, never automatic on open
 * (SPEC-046 "Literal text is not a chip"). */
function convertLiteralChips() {
  const model = doc.value;
  const root = contentRoot.value;
  if (!model || !root) {
    return;
  }
  const selection = paragraphSelection(root);
  const { model: next, converted } = convertLiteralChipsToRuns(model);
  if (converted === 0) {
    return;
  }
  applyEdit({ model: next, selection }, selection, "convert merge fields", false);
  addToast(converted === 1 ? "Converted 1 merge field" : `Converted ${converted} merge fields`);
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
    clearShapeSelection();
    disarmShapeDraw();
    clearTocSelection();
    exitHeaderFooterEditing();
    exitNoteEditing();
    return;
  }
  if (shapeSelection.value && !editingShapeText.value) {
    // A selected shape's keyboard interactions (SPEC-051 AC-40..43): a
    // dedicated keydown path, the same reason images.ts's own image-delete
    // path is one — a shape click never places a DOM caret/selection inside
    // the contenteditable region.
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      handleShapeDelete();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "d") {
      event.preventDefault();
      handleShapeDuplicate();
      return;
    }
    const direction =
      event.key === "ArrowUp" ? "up" :
      event.key === "ArrowDown" ? "down" :
      event.key === "ArrowLeft" ? "left" :
      event.key === "ArrowRight" ? "right" : "";
    if (direction) {
      event.preventDefault();
      handleShapeNudge(asNudgeDirection(direction), event.ctrlKey || event.metaKey);
      return;
    }
  }
  if (tocSelection.value && (event.key === "Delete" || event.key === "Backspace")) {
    // Deleting a selected TOC (SPEC-033 AC-16) is the same dedicated
    // keydown path SPEC-024 uses for a selected image: a click never
    // places a DOM caret/selection inside it for the native beforeinput
    // deletion flow to act on. Removes the whole atom; the headings it was
    // built from are separate blocks and are untouched.
    event.preventDefault();
    handleTocDelete();
    return;
  }
  if (tocSelection.value && isContentKey(event)) {
    // Every other content-producing key while a TOC is selected (SPEC-033
    // AC-14, AC-15): the hint fires here instead of letting the keystroke
    // reach whatever the native caret still sits on elsewhere.
    event.preventDefault();
    addToast('This table of contents isn’t editable — use "Update table of contents" instead.');
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
    clearPlainPaste();
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
    case "insert.comment": {
      // Ctrl/Cmd+Alt+M (SPEC-038 AC-1).
      event.preventDefault();
      openCommentComposer();
      break;
    }
    case "view.formattingMarks": {
      // Ctrl/Cmd+Shift+P (AC-12).
      event.preventDefault();
      toggleFormattingMarks();
      break;
    }
    case "edit.toggleTrackChanges": {
      // Ctrl/Cmd+Shift+E (SPEC-039 AC-1).
      event.preventDefault();
      toggleTrackChangesRecording();
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
      armPlainPaste();
      break;
    case "insert.nonBreakingSpace":
      // Ctrl/Cmd+Shift+Space (SPEC-026 AC-11): a direct non-breaking space
      // insert, not typed input, so autocorrect never sees it.
      event.preventDefault();
      insertSpecialCharacter("\u00A0");
      break;
    case "edit.selectAll": {
      // Ctrl/Cmd+A: select the whole document, exactly to the model's extent.
      // Shared with the Edit-menu "Select all" command (SPEC-049).
      event.preventDefault();
      menuSelectAll();
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

// --- document outline navigator (SPEC-037) ---------------------------------
// The caret's focus end, mirrored reactively (the live `caret` is a plain
// variable) so the pane's active-heading highlight tracks caret and scroll
// (AC-3, AC-4). Updated in updateCaretStatus.
const caretPosition = ref(noCaret());
// Per-user open/closed state, persisted to localStorage (AC-14) and open by
// default for a new user (AC-15).
const outlinePaneOpen = ref(loadOutlinePaneOpen());
watch(outlinePaneOpen, (open) => saveOutlinePaneOpen(open));
// Transient view state, never part of the model: which pane tree nodes are
// collapsed (AC-7, AC-8), and which body headings are collapsed in the
// document body (AC-16), keyed by "section:block". Body collapse hides blocks
// from layout only — it is never written to the file (AC-17).
const outlinePaneCollapsed = ref(emptyKeySet());
const collapsedHeadings = ref(emptyKeySet());

const outlineTree = computed(() => (doc.value ? buildOutline(doc.value) : []));
const activeOutlineAnchor = computed(() =>
  doc.value ? activeHeadingAnchor(doc.value, caretPosition.value) : null,
);
const hiddenBodyBlocks = computed(() =>
  doc.value ? hiddenBlockAddresses(doc.value, collapsedHeadings.value) : emptyKeySet(),
);

// --- pagination (SPEC-028) --------------------------------------------------
// The controller owns measurement; every doc.value replacement re-lays-out
// synchronously (flush: "sync") so the rendered pages always agree with the
// new model's block indices before Vue patches the canvas.

const pagination = usePagination(
  () => doc.value,
  () => contentRoot.value,
  (noteId) => noteLabelsMap.value.get(noteId) ?? "",
  () => hiddenBodyBlocks.value,
);
// A body-heading collapse/expand changes only the hidden set, not the model,
// so re-lay-out explicitly when it changes (AC-16). Empty-to-empty is not a
// change: the computed mints a fresh Set object whenever `doc` is replaced
// (opening a document), and relaying out for that fused a second full
// relayout into the open mount task (SPEC-047 AC-16) for a hidden set that
// is still "nothing hidden".
watch(hiddenBodyBlocks, (now, prev) => {
  if (now.size === 0 && prev !== undefined && prev.size === 0) {
    return;
  }
  pagination.schedule();
});
const paginationPages = pagination.pages;
const paginationBusy = pagination.busy;
const paginationState = pagination.state;

// --- open progress indicator (SPEC-047 AC-3..5): shown from this
// document's mount until pagination first reports settled, distinct from
// paginationBusy above (which reappears on every later edit-triggered
// relayout too — this one only ever fires once per document open). ------
const documentOpening = ref(true);
watch(documentId, () => {
  documentOpening.value = true;
});
watch(paginationState, (state) => {
  if (state === "clean") {
    documentOpening.value = false;
  }
});
onMounted(() => {
  // A document whose pagination never needs to measure (nothing realized
  // yet misses the cache) never transitions through "measuring", so the
  // watcher above never observes a change to react to — checked once,
  // after the initial schedule() call above has had a chance to run.
  void nextTick(() => {
    if (paginationState.value === "clean") {
      documentOpening.value = false;
    }
  });
});

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
// A measurement kick, not a model change: the immediate sync watcher above
// already laid out during setup (before `contentRoot` existed), so the mount
// kick only needs to start measuring against the now-present DOM — see
// scheduleMeasure's own comment (SPEC-047 AC-16).
onMounted(() => pagination.scheduleMeasure());
onBeforeUnmount(() => pagination.stop());

const marksOn = computed(() => doc.value?.showFormattingMarks === true);
const pageCount = computed(() => Math.max(1, paginationPages.value.length));
// The displayed page number of every laid-out page (SPEC-029 AC-20..23),
// recomputed whenever pagination or a section's numbering options change.
const pageNumbers = computed(() => {
  const model = doc.value;
  return model ? computePageNumbers(paginationPages.value, model) : [];
});

// Print preview and printing (SPEC-034): the preview overlays the canvas
// (which is itself the preview surface), owns the page-range state and the
// realize-everything gate both printing and PDF export share.
const {
  printPreviewOpen,
  printing,
  forceRealizeAll: printForceRealize,
  rangeMode: printRangeMode,
  rangeText: printRangeText,
  rangeError: printRangeError,
  canPrint,
  printablePageCount,
  printHiddenPages,
  openPrintPreview,
  closePrintPreview,
  confirmPrint,
  realizeAllPages,
} = usePrintPreview({ doc, pagination, pagesRoot: pageRef });

// PDF export (SPEC-034): reads the realized canvas into a client-built PDF.
const { pdfExporting, pdfProgress, exportPdf } = usePdfExport({
  doc,
  pagination,
  pagesRoot: pageRef,
  realizeAllPages,
  endRealizeAll: () => {
    if (!printPreviewOpen.value) {
      printForceRealize.value = false;
    }
  },
  addToast,
});

// Formatting marks are a screen affordance: the class that paints them
// drops while the print preview is open and while printing (SPEC-034
// AC-13). The model's showFormattingMarks setting itself is untouched.
const marksVisible = computed(
  () => marksOn.value && !printPreviewOpen.value && !printing.value,
);

// Hyperlinks, bookmarks and TOC navigation (SPEC-025, SPEC-033) live in
// useLinks; it needs pageNumbers (above) and the session's caret/scroll anchors.
const {
  linkDialogOpen,
  linkDialogInitial,
  linkChip,
  bookmarkDialogOpen,
  bookmarkManagerOpen,
  tocSelection,
  tocOptionsDialogOpen,
  documentHeadings,
  documentBookmarkTargets,
  documentBookmarkNames,
  openLinkDialog,
  handleLinkDialogSave,
  handleLinkDialogCancel,
  dismissLinkChip,
  refreshLinkChip,
  openLinkTarget,
  jumpToAnchor,
  handleLinkChipOpen,
  handleLinkChipEdit,
  handleLinkChipCopy,
  handleLinkChipRemove,
  openBookmarkDialog,
  handleBookmarkDialogSave,
  handleBookmarkDialogCancel,
  handleBookmarkJump,
  handleBookmarkRemove,
  closeBookmarkManager,
  maybeAutolink,
  tocPageNumberFor,
  openTocDialog,
  handleTocInsert,
  handleTocCancel,
  handleTocDelete,
  handleTocAction,
  handleTocUpdatePageNumbers,
  handleTocUpdateEntire,
  clearTocSelection,
} = useLinks({
  doc,
  contentRoot,
  editable,
  currentSelection,
  paragraphSelection,
  applyEdit,
  applyCaretMove,
  scrollModelCaretIntoView,
  restoreCaretAfterRender,
  activeHeaderFooter,
  pageOfCaret: (caretPos) => pagination.pageOfCaret(caretPos),
  pageNumbers,
  clearCellSelection,
  clearImageSelection,
  getCaret: () => caret,
  getSelectionAnchor: () => selectionAnchor,
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
    // Print preview / printing / PDF export need every page's real content
    // in the DOM (SPEC-034) — virtualization pauses while they run.
    !printForceRealize.value &&
    pagination.canMeasureOffscreen && paginationPages.value.length > VIRTUALIZE_MIN_PAGES,
);

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
  // Realization changed, the model didn't: measure the freshly rendered
  // blocks without forcing a synchronous relayout of an unchanged layout
  // (SPEC-047 AC-16 — see scheduleMeasure's own comment).
  void nextTick().then(() => pagination.scheduleMeasure());
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


// --- find and replace (SPEC-017) -------------------------------------------
// The DOM-facing find/replace logic is extracted into useFindReplace; it owns
// the find bar's reactive state, the highlight overlay and the post-edit
// refresh (wired into applyEdit via registerFindRefresh, and into session
// close / unmount / scroll via the hooks above).
const {
  findVisible,
  findExpanded,
  findQuery,
  findReplaceText,
  findMatchCase,
  findWholeWords,
  findUseRegex,
  findError,
  findHint,
  findCurrentIndex,
  findMatchesRef,
  findHighlightRects,
  findMatchPage,
  openFind,
  closeFind,
  toggleFindExpand,
  findNext,
  findPrev,
  replaceCurrent,
  replaceAll,
} = useFindReplace({
  doc,
  contentRoot,
  pageRef,
  editable,
  canSearch: findAndPrintEnabled,
  paginationPages,
  pageOfCaret: (caretPos) => pagination.pageOfCaret(caretPos),
  applyEdit,
  updateCaretStatus,
  addToast,
  getCaret: () => caret,
  setCaret: (next) => {
    caret = next;
  },
  getScrollContainer: () => scrollContainer,
  registerRefresh: registerFindRefresh,
  onSessionClose,
  onUnmount,
  onScroll: (cb) => scrollHooks.add(cb),
});

const {
  spellUnderlineRects,
  customDictionaryList,
  ignoreOccurrence,
  ignoreAllOccurrences,
  addWordToDictionary,
  removeWordFromDictionary,
  suggestionsFor,
} = useSpellCheck({
  doc,
  contentRoot,
  pageRef,
  getScrollContainer: () => scrollContainer,
  onSessionClose,
  onUnmount,
  onScroll: (cb) => scrollHooks.add(cb),
});

// --- comments (SPEC-038) -----------------------------------------------------
const {
  commentsPaneOpen,
  toggleCommentsPane,
  commentsOnlyView,
  toggleCommentsOnlyView,
  commentFilter,
  commentAuthorFilter,
  commentAuthorName,
  activeThreadId: activeCommentThreadId,
  openThreads: openCommentThreads,
  resolvedThreads: resolvedCommentThreads,
  commentAuthors,
  filteredThreads: filteredCommentThreads,
  commentComposerOpen,
  openCommentComposer,
  cancelCommentComposer,
  submitComment,
  submitReply: submitCommentReply,
  saveEditedMessage: saveEditedCommentMessage,
  deleteMessage: deleteCommentMessage,
  resolveThread: resolveCommentThread,
  reopenThread: reopenCommentThread,
  deleteThread: deleteCommentThread,
  focusThreadFromHighlight,
  focusThread: focusCommentThread,
  commentHighlightRects,
  commentThreadTops,
  scheduleCommentsMeasure,
} = useComments({
  doc,
  contentRoot,
  pagesRoot: pageRef,
  editable,
  currentSelection,
  applyEdit,
  applyCaretMove,
  scrollModelCaretIntoView,
  getCaret: () => caret,
  addToast,
  onSessionClose,
  onUnmount,
  onScroll: (cb) => scrollHooks.add(cb),
});

// --- track changes (SPEC-039) ------------------------------------------
const {
  trackChangesRecording,
  toggleTrackChangesRecording,
  isRecording,
  trackChangesAuthorName,
  trackChangesDisplayMode,
  setTrackChangesDisplayMode,
  trackChangesPaneOpen,
  toggleTrackChangesPane,
  changes: trackedChanges,
  acceptChange,
  rejectChange,
  acceptAll: acceptAllChanges,
  rejectAll: rejectAllChanges,
  goToNextChange,
  goToPreviousChange,
  goToChangeById,
} = useTrackChanges({
  doc,
  applyEdit,
  getCaret: () => caret,
  applyCaretMove,
  scrollModelCaretIntoView,
  onSessionClose,
});

const {
  versionHistoryPaneOpen,
  toggleVersionHistoryPane,
  versions: versionHistoryEntries,
  pausedMessage: versionHistoryPausedMessage,
  namedOnlyFilter: versionHistoryNamedOnlyFilter,
  previewActive: versionHistoryPreviewActive,
  previewVersionId: versionHistoryPreviewVersionId,
  beginVersionHistorySession,
  noteChange: noteVersionHistoryChange,
  noteCommitPoint: noteVersionHistoryCommitPoint,
  selectVersion: selectVersionHistoryVersion,
  exitPreview: exitVersionHistoryPreview,
  restoreVersion: restoreVersionHistoryVersion,
  nameVersion: nameVersionHistoryVersion,
  removeVersionName: removeVersionHistoryName,
  deleteVersion: deleteVersionHistoryVersion,
  downloadVersion: downloadVersionHistoryVersion,
} = useVersionHistory({
  doc,
  documentId,
  applyEdit,
  getCaret: () => caret,
  onSessionClose,
});

// --- accessibility checker (SPEC-043) ---------------------------------------
const {
  accessibilityPaneOpen,
  accessibilityIssues,
  accessibilityScanning,
  openAccessibilityPane,
  closeAccessibilityPane,
  goToAccessibilityIssue,
  applyAccessibilityFix,
  rescanAccessibility,
} = useAccessibilityCheck({
  doc,
  applyEdit,
  applyCaretMove,
  scrollModelCaretIntoView,
  imageSelection,
  openImageOptions,
  onSessionClose,
});

/** The Image options dialog's Apply (SPEC-024) also commits the accessibility
 * checker's "Add alt text" fix (AC-27, AC-28) — its own Apply handler has no
 * reason to know the checker exists, so this thin wrapper is what refreshes
 * the pane's list once the user actually applies the alt text they typed. */
function handleImageOptionsApplyAndRescan(payload = { altText: "", widthCm: 0, heightCm: 0 }) {
  handleImageOptionsApply(payload);
  if (accessibilityPaneOpen.value) {
    rescanAccessibility();
  }
}

// --- document properties and metadata (SPEC-041) ---------------------------
const {
  propertiesOpen,
  propertiesFileSize,
  openProperties,
  closeProperties,
  updateDocumentProperties,
  updateAuthorIdentity,
} = useDocumentProperties({
  doc,
  touchDocument,
  commentAuthorName,
  trackChangesAuthorName,
});

// Write paths that bypass applyEdit still respect the too-large read-only
// mode (SPEC-047 AC-27): restoring a version replaces the model wholesale,
// and the Properties dialog mutates metadata directly — both are editing
// commands, which "are unavailable" past the ceiling. (Reading properties,
// like reading the document, stays allowed — only the writes are gated.)
function restoreVersionRespectingCeiling(versionId = "") {
  if (pastSupportedCeiling.value) {
    return;
  }
  void restoreVersionHistoryVersion(versionId);
}
function updatePropertiesRespectingCeiling(next = {}) {
  if (pastSupportedCeiling.value) {
    return;
  }
  updateDocumentProperties(next);
}
function updateAuthorIdentityRespectingCeiling(name = "") {
  if (pastSupportedCeiling.value) {
    return;
  }
  updateAuthorIdentity(name);
}
// The Properties dialog's Statistics tab renders useWordCount's dialogStats
// (SPEC-041 AC-4). Since the worker move (SPEC-047 AC-12) those stats are
// computed on demand rather than by a doc-tracking computed, so this dialog
// must request them like the Statistics dialog does — on open, and again if
// the document is replaced under an open dialog (multi-tab refresh).
watch(propertiesOpen, (open) => {
  if (open) {
    void refreshStatisticsStats();
  }
});
watch(doc, () => {
  if (propertiesOpen.value) {
    void refreshStatisticsStats();
  }
});

// Every markDirty() call site also feeds the version-history cadence timer
// (SPEC-040 AC-1) — the same activity signal, one order of magnitude
// coarser (see versionCadenceScheduler.ts's header comment).
function noteEdit() {
  autosave.markDirty();
  noteVersionHistoryChange();
}

// SPEC-040 AC-11..13: while previewing an old version, the canvas must show
// its tracked-change diff markup regardless of the live document's own
// track-changes display-mode setting.
const effectiveTrackChangesDisplayMode = computed(() =>
  versionHistoryPreviewActive.value ? asTrackChangesDisplayMode("all") : trackChangesDisplayMode.value,
);

const {
  statusBarLanguageLabel,
  currentLanguageTag,
  selectionActive: languageSelectionActive,
  selectionNoProof,
  setSelection: setLanguageSelection,
  setLanguage,
  setNoProof,
} = useDocumentLanguage({ doc, applyEdit });

const customDictionaryDialogOpen = ref(false);

function openCustomDictionaryDialog() {
  customDictionaryDialogOpen.value = true;
}

function closeCustomDictionaryDialog() {
  customDictionaryDialogOpen.value = false;
}

function handleCustomDictionaryAdd(word = "", lang = "") {
  addWordToDictionary(word, lang);
}

function handleCustomDictionaryRemove(word = "", lang = "") {
  removeWordFromDictionary(word, lang);
}

const commentsDockRef = useTemplateRef("commentsDockRef");
watch(commentsDockRef, () => scheduleCommentsMeasure(commentsDockRef.value));

// Seed for espree-style handler defaults that never get real values here
// (applyHeaderFooterEdit/applyNoteEdit's `= noDocument()` params, tocPageNumberFor).
function noDocument() {
  return createDocument("unused");
}

// The paginated render data is derived by usePaginationView (it owns flatBlocks
// and the renderEntryFor / note-area / header-band builders). Kept in a
// composable because it is the framework-free counterpart to the editing core —
// nothing here edits the model, it only says what the canvas shows.
const { pagedViews } = usePaginationView({
  doc,
  pagination,
  noteLabelsMap,
  noteTooltipsMap,
  pageNumbers,
  marksOn,
  virtualizing,
  viewportPageRange,
  caretPage,
  anchorPage,
  findMatchPage,
  activeHeaderFooter,
  activeNote,
  editable,
  cellSelection,
  imageSelection,
  shapeSelection,
  shapeMultiSelection,
  editingShapeText,
  collapsedHeadings,
  remount,
  realizedSticky,
  hfRemountEpoch,
  noteRemountEpoch,
  tocPageNumberFor,
  trackChangesDisplayMode: effectiveTrackChangesDisplayMode,
});

// --- outline navigator handlers (SPEC-037) ---------------------------------
/** View menu toggle / persistence of the pane's open state (AC-13, AC-14). */
function toggleOutlinePane() {
  outlinePaneOpen.value = !outlinePaneOpen.value;
}

/** Clicking or Enter on an outline entry (AC-5, AC-6, AC-19): jump to the
 * heading — place the caret at its start and scroll it into view. Reuses the
 * same anchor-jump the TOC and internal links use. */
function handleOutlineNavigate(anchor = "") {
  jumpToAnchor(anchor);
}

/** Collapse/expand a node in the pane (AC-7, AC-8) — pane-only view state. */
function handleOutlineToggleCollapse(anchor = "") {
  const next = new Set(outlinePaneCollapsed.value);
  if (next.has(anchor)) {
    next.delete(anchor);
  } else {
    next.add(anchor);
  }
  outlinePaneCollapsed.value = next;
}

/** Drop of a dragged outline entry (AC-9, AC-10, AC-11): move the heading and
 * all its content as ONE undo step. moveHeading returns null for a rejected
 * drop (different level, non-sibling, onto itself), so no edit is recorded. */
function handleOutlineMove(payload = { source: "", target: "", position: "before" }) {
  const model = doc.value;
  if (!model || !editable.value) {
    return;
  }
  const before = currentSelection(contentRoot.value);
  // Narrow the payload's position to the DropPosition literal union (espree
  // bars a type annotation on the handler's own parameter).
  const position = payload.position === "after" ? "after" : "before";
  const result = moveHeading(model, payload.source, payload.target, position);
  if (result) {
    applyEdit(result, before, "move heading");
  }
}

/** The body chevron next to a heading (AC-16): collapse/expand its content.
 * Keyed by "section:block"; purely transient (AC-17). When collapsing the
 * section the caret sits in, move the caret to the heading so it never hides
 * behind collapsed content. */
function handleBodyHeadingToggle(section = 0, block = 0) {
  const model = doc.value;
  if (!model) {
    return;
  }
  const key = blockKey(section, block);
  const next = new Set(collapsedHeadings.value);
  if (next.has(key)) {
    next.delete(key);
  } else {
    next.add(key);
  }
  collapsedHeadings.value = next;
  const hidden = hiddenBlockAddresses(model, next);
  if (caret && hidden.has(blockKey(caret.section, caret.block))) {
    applyCaretMove({ section, block, offset: 0 });
  }
}

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

const downloadFormat = computed(() => {
  if (!doc.value) {
    return "docx";
  }
  if (doc.value.format === "odt") {
    return "odt";
  }
  if (doc.value.format === "md") {
    return "md";
  }
  return "docx";
});

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
  // PDF export (SPEC-034 AC-14): a rendered-output snapshot, not a save —
  // it runs its own pipeline and never touches autosave or the model.
  if (format === "pdf") {
    void exportPdf();
    return;
  }
  if (!doc.value || (format !== "docx" && format !== "odt" && format !== "md")) {
    return;
  }
  // The dcterms:created/dc:date (SPEC-041 AC-2) stamp: `created` is set only
  // the first time (kept stable on every later save, matching what General
  // tab shows once it has a value), `modified` refreshes every save. Both
  // .docx/.odt writers read `model.properties.created`/`modified` — see
  // docProps.ts/meta.ts's buildCorePropertiesXml/buildMetaXml — so this
  // must be resolved into the model *before* writing, then mirrored back so
  // the in-memory model (and so the Properties dialog, without a reopen)
  // agrees with what was just written. Markdown never reads `properties`
  // (SPEC-050: front matter is carried verbatim, never interpreted), but the
  // stamp is still recorded so a later save as .docx/.odt keeps it.
  const modified = new Date();
  const properties = {
    ...doc.value.properties,
    created: doc.value.properties?.created ?? modified.toISOString(),
    modified: modified.toISOString(),
  };
  doc.value.properties = properties;
  const { bytes, dropped } =
    format === "odt"
      ? await writeOdt(doc.value, { modified })
      : format === "md"
        ? await writeMd(doc.value)
        : await writeDocx(doc.value, { modified });
  triggerRef(doc);
  touchDocument(doc.value.id);
  const fileName =
    format === "odt" ? odtFileName(doc.value.title) : format === "md" ? mdFileName(doc.value.title) : docxFileName(doc.value.title);
  // A transient anchor pointed at an object URL is the download mechanism
  // that needs no new browser permissions; revocation is deferred so the
  // browser has started streaming the blob before the URL disappears.
  const mimeType = format === "odt" ? ODT_MIMETYPE : format === "md" ? MD_MIME_TYPE : DOCX_MIME_TYPE;
  const url = URL.createObjectURL(new Blob([bytes], { type: mimeType }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
  // A download is SPEC-003's explicit save: promote the state to committed,
  // and SPEC-040's other commit-point version-history trigger (AC-2).
  void autosave.commitNow().then(() => noteVersionHistoryCommitPoint());
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

const savingAsTemplate = ref(false);

// --- mail-merge chip insertion (SPEC-046 AC-12, AC-15, AC-30) ---------------
// The attached data source (if any) is shared between "Insert -> Merge
// field...", "Table -> Repeat row for each..." and the merge run dialog —
// attaching it once populates every named field list, and running the merge
// re-parses the same file rather than trusting stale state across a session.
const mergeDataSource = shallowRef(noParsedDataSource());
const mergeFieldPickerOpen = ref(false);
const mergeRunDialogOpen = ref(false);

function attachMergeDataSource(parsed = noParsedDataSource()) {
  mergeDataSource.value = parsed;
}
const mergeFieldPickerFields = computed(() => (doc.value ? namedFieldList(mergeDataSource.value, doc.value) : []));
const repeatRowFieldOptions = computed(() => (doc.value ? namedArrayFieldList(mergeDataSource.value, doc.value) : []));

function insertMergeFieldFromPicker(name = "") {
  mergeFieldPickerOpen.value = false;
  const root = contentRoot.value;
  const model = doc.value;
  if (!name || !root || !model) {
    return;
  }
  const selection = paragraphSelection(root);
  const result = insertMergeChip(model, selection, { kind: "merge", name });
  applyEdit(result, selection, "insert merge field");
}

function saveCurrentAsTemplate(name = "") {
  savingAsTemplate.value = false;
  if (!doc.value) {
    return;
  }
  const record = saveAsTemplate(doc.value, name);
  addToast(`Saved template "${record.title}"`);
}

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

/** The registry shortcuts that act regardless of where focus sits in the
 * document view — Ctrl/Cmd+F/H (which work from the canvas, not just once
 * the bar is already open) and Ctrl/Cmd+P (print) — so this is bound on
 * `document` rather than the canvas element itself. Guarded like every
 * other registry shortcut (SPEC-021 AC-12..14): none of them act while
 * focus is inside a dialog, the find bar, or another text input. */
function handleGlobalShortcut(event = new KeyboardEvent("keydown")) {
  if (!doc.value || !findAndPrintEnabled.value || event.altKey) {
    return;
  }
  // Escape disarms the format painter regardless of where focus is (AC-17)
  // — handleKeyDown below also disarms it when Escape is pressed with the
  // canvas focused, so this is deliberately unguarded by isOutsideDocumentInput.
  if (event.key === "Escape") {
    disarmPainter();
    if (printPreviewOpen.value) {
      closePrintPreview();
    }
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
    // SPEC-034 AC-1 (superseding SPEC-021's immediate window.print()):
    // Ctrl/Cmd+P opens the in-app print preview; the native dialog opens
    // from its Print button.
    event.preventDefault();
    openPrintPreview();
  }
}

// --- Menu bar wiring (SPEC-049) --------------------------------------------
// The menu bar (File/Edit/View/Insert/Format/Tools, in that order) hosts the
// non-continuous-use commands relocated off the toolbar strip. It is a stable
// map of the product: a command that does not apply is disabled, never hidden
// (AC-7); the toolbar keeps only the continuous-use controls (AC-3).
const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
  { id: "insert", label: "Insert" },
  { id: "format", label: "Format" },
  { id: "tools", label: "Tools" },
];

// Text colour and highlight stay in the toolbar (continuous-use) but are also
// listed in the Format menu, so the menus remain a complete map — the spec's
// one deliberate toolbar+menu exception. The menu entries open the real
// toolbar pickers through these refs instead of duplicating the swatch popups
// (which would collide on their fixed data-testids).
const textColorPickerRef = useTemplateRef("textColorPickerRef");
const highlightPickerRef = useTemplateRef("highlightPickerRef");

function openTextColorMenu() {
  textColorPickerRef.value?.openMenu();
}

function openHighlightMenu() {
  highlightPickerRef.value?.openMenu();
}

// The Insert menu's "Image…" command drives the same hidden file picker the
// old toolbar button did (SPEC-024).
const menuImageInput = useTemplateRef("menuImageInput");

function triggerImageInsert() {
  if (editable.value) {
    menuImageInput.value?.click();
  }
}

function handleMenuImageChange(event = new Event("change")) {
  const target = event.target;
  const file = target instanceof HTMLInputElement ? target.files?.[0] : undefined;
  if (file) {
    handleInsertImageFile(file);
  }
  if (target instanceof HTMLInputElement) {
    // Clear so picking the same file twice still fires a change event.
    target.value = "";
  }
}

// The Insert menu's table-size grid (SPEC-022): the same quick-pick the old
// toolbar dropdown offered, now a section of the Insert menu.
const TABLE_GRID_ROWS = Array.from({ length: PICKER_MAX_ROWS }, (_, index) => index + 1);
const TABLE_GRID_COLS = Array.from({ length: PICKER_MAX_COLS }, (_, index) => index + 1);
const tableGridHover = ref({ rows: 0, cols: 0 });

function hoverTableGrid(rows = 0, cols = 0) {
  tableGridHover.value = { rows, cols };
}

function chooseTableGrid(rows = 0, cols = 0) {
  handleInsertTable({ rows, cols });
  // The pick closes the menu, so clear the highlight the pointer left behind:
  // reopening Insert should show an unlit grid, not the last size chosen.
  hoverTableGrid(0, 0);
}

/** "Insert -> Shape" gallery pick (SPEC-051 AC-3): arms draw mode so the
 * next drag/click on the canvas draws the chosen geometry. */
function handleShapePick(geometry = "rect") {
  armShapeDraw(asShapeGeometry(geometry));
}

// The multilevel-list presets (SPEC-032), absorbed into the Format menu.
const MULTILEVEL_PRESETS = GALLERY_PRESETS;

// Edit-menu Select all / Paste without formatting: menu-invocable forms of
// the Ctrl+A / Ctrl+Shift+V commands (their keyboard paths call these too).
function menuSelectAll() {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model) {
    return;
  }
  pendingFormat = noPendingFormat();
  const all = selectAll(model);
  caret = all.focus;
  updateCaretStatus({ start: all.anchor, end: all.focus });
  if (paragraphElementForCaret(root, all.anchor) && paragraphElementForCaret(root, all.focus)) {
    setDomSelection(root, all.anchor, all.focus);
  } else {
    void nextTick().then(() => setDomSelection(root, all.anchor, all.focus));
  }
}

async function menuPastePlain() {
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
  performPaste({ internal: "", html: "", text: flavours.text, hasImage: false }, selection, true);
}

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
    <template #menubar>
      <MenuBar
        :menus="MENUS"
        aria-label="Document menu"
      >
        <template #file>
          <MenuItem
            label="Print"
            testid="file-menu-print"
            :shortcut="shortcutLabel('file.print')"
            @select="openPrintPreview"
          />
          <MenuItem
            label="Page setup"
            testid="page-setup-button"
            :disabled="!editable"
            @select="openPageSetup"
          />
          <MenuItem
            label="Properties…"
            testid="file-menu-properties"
            @select="openProperties"
          />
          <MenuItem
            label="Version history"
            testid="file-menu-version-history"
            @select="toggleVersionHistoryPane"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Save as template"
            testid="file-menu-save-as-template"
            @select="savingAsTemplate = true"
          />
          <MenuItem
            label="Mail merge…"
            testid="file-menu-mail-merge"
            @select="mergeRunDialogOpen = true"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Delete document"
            testid="delete-document"
            :disabled="!editable"
            @select="confirmingDelete = true"
          />
        </template>

        <template #edit>
          <MenuItem
            label="Undo"
            testid="edit-menu-undo"
            :shortcut="shortcutLabel('edit.undo')"
            :disabled="!editable || !undoRedoState.canUndo"
            @select="performUndo"
          />
          <MenuItem
            label="Redo"
            testid="edit-menu-redo"
            :shortcut="shortcutLabel('edit.redo')"
            :disabled="!editable || !undoRedoState.canRedo"
            @select="performRedo"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Cut"
            testid="edit-menu-cut"
            :disabled="!editable"
            @select="contextMenuCut"
          />
          <MenuItem
            label="Copy"
            testid="edit-menu-copy"
            @select="contextMenuCopy"
          />
          <MenuItem
            label="Paste"
            testid="edit-menu-paste"
            :disabled="!editable"
            @select="contextMenuPaste"
          />
          <MenuItem
            label="Paste without formatting"
            testid="edit-menu-paste-plain"
            :shortcut="shortcutLabel('edit.pastePlain')"
            :disabled="!editable"
            @select="menuPastePlain"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Select all"
            testid="edit-menu-select-all"
            :shortcut="shortcutLabel('edit.selectAll')"
            @select="menuSelectAll"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Find and replace"
            testid="edit-menu-find-replace"
            :shortcut="shortcutLabel('find.openReplace')"
            @select="openFind(true)"
          />
          <hr class="menu-separator">
          <div class="menu-section-label">
            Autocorrect
          </div>
          <MenuCheckItem
            label="Smart quotes"
            testid="autocorrect-smart-quotes"
            :checked="autocorrectSettings.smartQuotes"
            :disabled="!editable"
            @select="toggleAutocorrectSmartQuotes"
          />
          <MenuCheckItem
            label="Em and en dashes"
            testid="autocorrect-dashes"
            :checked="autocorrectSettings.dashes"
            :disabled="!editable"
            @select="toggleAutocorrectDashes"
          />
        </template>

        <template #view>
          <MenuCheckItem
            label="Ruler"
            testid="toggle-ruler-button"
            :checked="doc.showRuler !== false"
            @select="toggleRuler"
          />
          <MenuCheckItem
            label="Outline pane"
            testid="view-menu-outline"
            :checked="outlinePaneOpen"
            @select="toggleOutlinePane"
          />
          <MenuCheckItem
            label="Styles pane"
            testid="view-menu-styles-pane"
            :checked="stylesPaneOpen"
            @select="toggleStylesPane"
          />
          <MenuCheckItem
            label="Comments pane"
            testid="view-menu-comments-pane"
            :checked="commentsPaneOpen"
            @select="toggleCommentsPane"
          />
          <MenuCheckItem
            label="Comments-only view"
            testid="view-menu-comments-only"
            :checked="commentsOnlyView"
            @select="toggleCommentsOnlyView"
          />
          <hr class="menu-separator">
          <MenuCheckItem
            label="Track changes"
            testid="view-menu-track-changes"
            :shortcut="shortcutLabel('edit.toggleTrackChanges')"
            :checked="trackChangesRecording"
            :disabled="!editable"
            @select="toggleTrackChangesRecording"
          />
          <MenuCheckItem
            label="Track changes pane"
            testid="view-menu-track-changes-pane"
            :checked="trackChangesPaneOpen"
            @select="toggleTrackChangesPane"
          />
        </template>

        <template #insert>
          <div class="menu-section-label">
            Table
          </div>
          <div class="menu-table-grid-wrap">
            <div
              class="menu-table-grid"
              data-testid="insert-table-grid"
              @mouseleave="hoverTableGrid(0, 0)"
            >
              <div
                v-for="rows in TABLE_GRID_ROWS"
                :key="rows"
                class="menu-table-grid-row"
              >
                <button
                  v-for="cols in TABLE_GRID_COLS"
                  :key="cols"
                  type="button"
                  class="menu-table-grid-cell"
                  :class="{ active: rows <= tableGridHover.rows && cols <= tableGridHover.cols }"
                  :data-testid="`table-grid-cell-${rows}-${cols}`"
                  :aria-label="`Insert ${rows} by ${cols} table`"
                  :disabled="!editable"
                  tabindex="-1"
                  @mouseenter="hoverTableGrid(rows, cols)"
                  @mousedown.prevent
                  @click="chooseTableGrid(rows, cols)"
                />
              </div>
            </div>
            <div
              class="menu-table-grid-label"
              data-testid="insert-table-grid-label"
            >
              {{ tableGridHover.rows > 0 ? `${tableGridHover.rows} × ${tableGridHover.cols}` : "Insert table" }}
            </div>
          </div>
          <MenuItem
            label="Insert table…"
            testid="insert-table-dialog-item"
            :disabled="!editable"
            @select="insertTableDialogOpen = true"
          />
          <MenuItem
            label="Image…"
            testid="insert-image-button"
            :disabled="!editable"
            @select="triggerImageInsert"
          />
          <hr class="menu-separator">
          <div class="menu-section-label">
            Shape
          </div>
          <ShapeGallery @pick="handleShapePick" />
          <MenuItem
            label="Text box"
            testid="insert-textbox-button"
            :disabled="!editable"
            @select="insertTextBoxAction"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Page break"
            testid="break-page"
            :shortcut="shortcutLabel('insert.pageBreak')"
            :disabled="!editable"
            @select="handleInsertBreak('page')"
          />
          <MenuItem
            label="Column break"
            testid="break-column"
            :disabled="!editable"
            @select="handleInsertBreak('column')"
          />
          <MenuItem
            label="Section break (next page)"
            testid="break-section-next"
            :disabled="!editable"
            @select="handleInsertSectionBreak('nextPage')"
          />
          <MenuItem
            label="Section break (continuous)"
            testid="break-section-continuous"
            :disabled="!editable"
            @select="handleInsertSectionBreak('continuous')"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Header"
            testid="hf-menu-header"
            :disabled="!editable"
            @select="openInsertHeaderFooter('header')"
          />
          <MenuItem
            label="Footer"
            testid="hf-menu-footer"
            :disabled="!editable"
            @select="openInsertHeaderFooter('footer')"
          />
          <MenuItem
            label="Page number"
            testid="hf-menu-page-number"
            :disabled="!editable"
            @select="openPageNumberDialog"
          />
          <MenuItem
            label="Header and footer options"
            testid="hf-menu-options"
            :disabled="!editable"
            @select="openHeaderFooterOptions"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Merge field…"
            testid="insert-menu-merge-field"
            :disabled="!editable"
            @select="mergeFieldPickerOpen = true"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Footnote"
            testid="notes-menu-footnote"
            :shortcut="shortcutLabel('insert.footnote')"
            :disabled="!editable"
            @select="handleInsertNote('footnote')"
          />
          <MenuItem
            label="Endnote"
            testid="notes-menu-endnote"
            :shortcut="shortcutLabel('insert.endnote')"
            :disabled="!editable"
            @select="handleInsertNote('endnote')"
          />
          <MenuItem
            label="Footnote and endnote settings"
            testid="notes-menu-settings"
            :disabled="!editable"
            @select="openNoteSettings"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Table of contents"
            testid="insert-toc-button"
            :disabled="!editable"
            @select="openTocDialog"
          />
          <MenuItem
            label="Special character"
            testid="insert-special-character-button"
            :disabled="!editable"
            @select="openSpecialCharacters"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Link"
            testid="toolbar-insert-link"
            :shortcut="shortcutLabel('insert.link')"
            :disabled="!editable"
            @select="openLinkDialog"
          />
          <MenuItem
            label="Bookmark"
            testid="toolbar-insert-bookmark"
            :disabled="!editable"
            @select="openBookmarkDialog"
          />
          <MenuItem
            label="Bookmark list"
            testid="toolbar-bookmarks"
            @select="bookmarkManagerOpen = true"
          />
          <MenuItem
            label="Comment"
            testid="toolbar-insert-comment"
            :shortcut="shortcutLabel('insert.comment')"
            :disabled="!editable"
            @select="openCommentComposer"
          />
        </template>

        <template #format>
          <MenuItem
            label="Styles pane"
            testid="format-menu-styles-pane"
            @select="toggleStylesPane"
          />
          <MenuItem
            label="Import styles from another document…"
            testid="format-menu-import-styles"
            :disabled="!editable"
            @select="openImportDialog"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Clear formatting"
            testid="clear-formatting-button"
            :shortcut="shortcutLabel('format.clearFormatting')"
            :disabled="!editable"
            @select="handleClearFormatting"
          />
          <!-- Format painter keeps its double-click-to-lock gesture
               (SPEC-020 AC-14): data-menu-keep-open holds the menu open
               across the click/click/dblclick sequence, and its armed state
               shows as a live check (SPEC-049 AC-11). -->
          <button
            type="button"
            role="menuitemcheckbox"
            class="menu-item"
            data-testid="format-painter-button"
            data-menu-keep-open
            title="Format painter (double-click to keep it on)"
            :aria-checked="painterMode !== '' ? 'true' : 'false'"
            :disabled="!editable"
            tabindex="-1"
            @mousedown.prevent
            @click="handlePainterClick"
            @dblclick="handlePainterDblClick"
          >
            <span
              class="menu-check"
              aria-hidden="true"
            >{{ painterMode !== '' ? '✓' : '' }}</span>
            <span class="menu-item-label">Format painter</span>
          </button>
          <hr class="menu-separator">
          <MenuItem
            label="Columns"
            testid="columns-button"
            :disabled="!editable"
            @select="openColumnsDialog"
          />
          <hr class="menu-separator">
          <div class="menu-section-label">
            Multilevel list
          </div>
          <MenuItem
            v-for="preset in MULTILEVEL_PRESETS"
            :key="preset.key"
            :label="preset.label"
            :testid="`multilevel-preset-${preset.key}`"
            :disabled="!editable"
            @select="handleMultilevelPreset(preset.key)"
          />
          <MenuItem
            label="Number headings 1–3"
            testid="multilevel-link-headings"
            :disabled="!editable"
            @select="handleLinkHeadingNumbering"
          />
          <MenuItem
            label="Define new list…"
            testid="multilevel-define-list"
            :disabled="!editable"
            @select="openDefineList"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Text colour"
            testid="format-menu-text-color"
            :disabled="!editable"
            @select="openTextColorMenu"
          />
          <MenuItem
            label="Highlight colour"
            testid="format-menu-highlight"
            :disabled="!editable"
            @select="openHighlightMenu"
          />
        </template>

        <template #tools>
          <MenuItem
            label="Word count"
            testid="tools-menu-word-count"
            @select="openStatistics"
          />
          <MenuItem
            label="Custom dictionary"
            testid="tools-menu-dictionary"
            @select="openCustomDictionaryDialog"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Check accessibility"
            testid="tools-menu-accessibility"
            @select="openAccessibilityPane"
          />
          <hr class="menu-separator">
          <MenuItem
            label="Convert {{…}} text to merge fields"
            testid="tools-menu-convert-merge-fields"
            :disabled="!editable || !canConvertLiteralChips"
            @select="convertLiteralChips"
          />
        </template>
      </MenuBar>
      <!-- The Insert menu's "Image…" command drives this hidden picker
           (SPEC-024). -->
      <input
        ref="menuImageInput"
        type="file"
        class="menu-image-input"
        data-testid="insert-image-file-input"
        :accept="ACCEPTED_IMAGE_ACCEPT_ATTR"
        @change="handleMenuImageChange"
      >
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
      <TextColorPicker
        v-if="editable"
        ref="textColorPickerRef"
        :value="formatValues.color"
        :mixed="formatValues.colorMixed"
        :recent="recentColors"
        @select="handleTextColorSelect"
      />
      <HighlightPicker
        v-if="editable"
        ref="highlightPickerRef"
        :value="formatValues.highlight"
        :mixed="formatValues.highlightMixed"
        @select="handleHighlightSelect"
      />
      <button
        v-if="editable"
        type="button"
        class="track-changes-toolbar-toggle"
        data-testid="toolbar-track-changes-toggle"
        :aria-pressed="trackChangesRecording"
        :title="shortcutLabel('edit.toggleTrackChanges')"
        @click="toggleTrackChangesRecording"
      >
        Track changes: {{ trackChangesRecording ? "On" : "Off" }}
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
      v-if="pastSupportedCeiling"
      class="notice-bar"
      role="status"
      data-testid="too-large-bar"
    >
      {{ ceilingWarning }}
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
    <div class="editor-layout">
      <div
        v-if="outlinePaneOpen"
        class="outline-dock"
      >
        <OutlinePane
          :outline="outlineTree"
          :collapsed="outlinePaneCollapsed"
          :active-anchor="activeOutlineAnchor ?? undefined"
          @navigate="handleOutlineNavigate"
          @toggle="handleOutlineToggleCollapse"
          @move="handleOutlineMove"
        />
      </div>
      <div
        v-if="stylesPaneOpen"
        class="styles-dock print-exclude"
      >
        <StylesPane
          :entries="stylesPaneEntriesList"
          :caret-paragraph-style-id="styleState.id"
          :caret-paragraph-style-name="styleState.name"
          :caret-character-style-id="characterStyleState.id"
          :caret-character-style-name="characterStyleState.name"
          @close="closeStylesPane"
          @new-style="openNewStyle"
          @modify-style="openModifyStyle"
          @delete-style="openDeleteStyle"
          @apply-style="applyStyleFromPane"
        />
      </div>
      <div
        v-if="commentsPaneOpen"
        ref="commentsDockRef"
        class="comments-dock print-exclude"
      >
        <CommentsPane
          :open-threads="openCommentThreads"
          :resolved-threads="resolvedCommentThreads"
          :filtered-threads="filteredCommentThreads"
          :comments-only-view="commentsOnlyView"
          :thread-tops="commentThreadTops"
          :active-thread-id="activeCommentThreadId ?? undefined"
          :author-name="commentAuthorName"
          :comment-filter="commentFilter"
          :comment-author-filter="commentAuthorFilter"
          :authors="commentAuthors"
          :composer-open="commentComposerOpen"
          :editable="editable"
          @close="toggleCommentsPane"
          @toggle-comments-only="toggleCommentsOnlyView"
          @update:comment-filter="commentFilter = asCommentFilter($event)"
          @update:author-filter="commentAuthorFilter = $event"
          @update:author-name="commentAuthorName = $event"
          @submit-comment="submitComment"
          @cancel-composer="cancelCommentComposer"
          @focus-thread="focusCommentThread"
          @submit-reply="submitCommentReply"
          @save-edit="saveEditedCommentMessage"
          @delete-message="deleteCommentMessage"
          @resolve="resolveCommentThread"
          @reopen="reopenCommentThread"
          @delete-thread="deleteCommentThread"
        />
      </div>
      <div
        v-if="trackChangesPaneOpen"
        class="track-changes-dock print-exclude"
      >
        <TrackChangesPane
          :changes="trackedChanges"
          :display-mode="trackChangesDisplayMode"
          :recording="trackChangesRecording"
          :author-name="trackChangesAuthorName"
          :editable="editable"
          @close="toggleTrackChangesPane"
          @toggle-recording="toggleTrackChangesRecording"
          @update:display-mode="setTrackChangesDisplayMode"
          @update:author-name="trackChangesAuthorName = $event"
          @accept="acceptChange"
          @reject="rejectChange"
          @accept-all="acceptAllChanges"
          @reject-all="rejectAllChanges"
          @next="goToNextChange"
          @previous="goToPreviousChange"
          @focus-change="goToChangeById"
        />
      </div>
      <div
        v-if="versionHistoryPaneOpen"
        class="version-history-dock print-exclude"
      >
        <VersionHistoryPane
          :versions="versionHistoryEntries"
          :paused-message="versionHistoryPausedMessage ?? ''"
          :named-only-filter="versionHistoryNamedOnlyFilter"
          :preview-version-id="versionHistoryPreviewVersionId ?? ''"
          :editable="documentActionsEnabled && !pastSupportedCeiling"
          @close="toggleVersionHistoryPane"
          @update:named-only-filter="versionHistoryNamedOnlyFilter = $event"
          @select="selectVersionHistoryVersion"
          @exit-preview="exitVersionHistoryPreview"
          @restore="restoreVersionRespectingCeiling"
          @name="nameVersionHistoryVersion"
          @remove-name="removeVersionHistoryName"
          @delete="deleteVersionHistoryVersion"
          @download="downloadVersionHistoryVersion"
        />
      </div>
      <div
        v-if="accessibilityPaneOpen"
        class="accessibility-dock print-exclude"
      >
        <AccessibilityPane
          :issues="accessibilityIssues"
          :scanning="accessibilityScanning"
          @close="closeAccessibilityPane"
          @select="goToAccessibilityIssue"
          @fix="applyAccessibilityFix"
        />
      </div>
      <div class="page-wrap">
        <div
          v-if="findVisible"
          class="find-bar-dock print-exclude"
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
            class="find-highlight-layer print-exclude"
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
            v-if="commentHighlightRects.length > 0"
            class="comment-highlight-layer print-exclude"
            data-testid="comment-highlights"
          >
            <div
              v-for="(rect, index) in commentHighlightRects"
              :key="index"
              class="comment-highlight"
              :class="{ active: rect.active }"
              :data-testid="`comment-highlight-${rect.threadId}`"
              :style="{
                left: `${rect.left}px`,
                top: `${rect.top}px`,
                width: `${rect.width}px`,
                height: `${rect.height}px`,
              }"
              @click="focusThreadFromHighlight(rect.threadId)"
            />
          </div>
          <!-- The rects address a paragraph offset range, but they are NOT
               blocks: the address is spelled data-spell-* rather than
               data-section/data-block because that exact pair is
               BLOCK_SELECTOR (lib/domCaret.ts), the selector caret mapping,
               pagination and page navigation use to enumerate real block
               elements. Naming them the same way makes every flagged word
               masquerade as an extra slice of its own paragraph. -->
          <div
            v-if="spellUnderlineRects.length > 0"
            class="spell-underline-layer print-exclude"
            data-testid="spell-underlines"
            aria-hidden="true"
          >
            <div
              v-for="(rect, index) in spellUnderlineRects"
              :key="index"
              class="spell-underline"
              data-testid="spell-underline"
              :data-spell-section="rect.section"
              :data-spell-block="rect.block"
              :data-spell-start="rect.start"
              :data-spell-end="rect.end"
              :data-word="rect.word"
              :data-lang="rect.lang"
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
            class="table-handles print-exclude"
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
            class="image-handles print-exclude"
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
            v-if="shapeHandles"
            class="shape-handles print-exclude"
            data-testid="shape-handles"
          >
            <div
              v-for="handle in IMAGE_HANDLES"
              :key="handle"
              class="shape-resize-handle"
              :data-testid="`shape-resize-${handle}`"
              :style="shapeHandleStyle(handle)"
              @mousedown="startShapeResize($event, handle)"
            />
            <div
              class="shape-rotation-handle"
              data-testid="shape-rotation-handle"
              :style="rotationHandleStyle()"
              @mousedown="startShapeRotate($event)"
            />
          </div>
          <div
            v-if="shapeAnchorMarker"
            class="shape-anchor-marker print-exclude"
            data-testid="shape-anchor-marker"
            :style="{ left: `${shapeAnchorMarker.left}px`, top: `${shapeAnchorMarker.top}px` }"
          />
          <div
            ref="contentRoot"
            class="document-content"
            :class="{ 'painter-armed': painterMode !== '', 'marks-on': marksVisible }"
            data-testid="document-content"
            role="textbox"
            aria-multiline="true"
            aria-label="Document content"
            :aria-readonly="editable ? undefined : 'true'"
            :contenteditable="editable && !activeHeaderFooter && !activeNote && !editingShapeText"
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
            @dblclick="handleCanvasDblClick"
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
              :print-hidden="printHiddenPages.has(pageIndex)"
              @image-select="handleImageSelect"
              @shape-select="handleShapeSelect"
              @shape-edit-text="handleShapeEditText"
              @shape-text-beforeinput="handleShapeTextBeforeInput"
              @shape-text-keydown="handleShapeTextKeyDown"
              @shape-text-compositionstart="handleShapeTextCompositionStart"
              @shape-text-compositionend="handleShapeTextCompositionEnd"
              @toc-action="handleTocAction"
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
              @outline-toggle="handleBodyHeadingToggle"
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
        <div
          class="visually-hidden"
          role="status"
          aria-live="polite"
          data-testid="word-count-status"
        >
          {{ wordCountAnnouncement }}
        </div>
      </div>
    </div>
    <template #status-bar>
      <span data-testid="page-indicator">Page {{ caretPage }} of {{ pageCount }}</span>
      <span
        v-if="documentOpening"
        role="status"
        data-testid="document-opening"
      >Opening…</span>
      <span
        v-if="paginationBusy"
        data-testid="pagination-busy"
      >Paginating…</span>
      <button
        type="button"
        class="word-count-button"
        data-testid="word-count-button"
        @click="openStatistics"
      >
        {{ wordCountStatusText }}
      </button>
      <LanguageStatusControl
        :current-language-tag="currentLanguageTag"
        :current-language-label="statusBarLanguageLabel"
        :language-options="SUPPORTED_LANGUAGES"
        :selection-active="languageSelectionActive"
        :selection-no-proof="selectionNoProof"
        @select="setLanguage"
        @manage-dictionary="openCustomDictionaryDialog"
        @toggle-no-proof="setNoProof"
      />
      <button
        v-if="editable"
        type="button"
        class="track-changes-status-indicator"
        data-testid="status-bar-track-changes"
        :aria-pressed="trackChangesRecording"
        @click="toggleTrackChangesRecording"
      >
        Track changes: {{ trackChangesRecording ? "On" : "Off" }}
      </button>
    </template>
    <LinkChip
      v-if="linkChip.visible"
      class="print-exclude"
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
    <RenameDialog
      v-if="savingAsTemplate && doc"
      :current-title="doc.title"
      heading="Save as template"
      confirm-label="Save"
      @save="saveCurrentAsTemplate"
      @cancel="savingAsTemplate = false"
    />
    <ConversionReportDialog
      v-if="conversionReport"
      :items="conversionReport"
      @close="closeConversionReport"
    />
    <MergeFieldPickerDialog
      v-if="mergeFieldPickerOpen"
      heading="Insert merge field"
      :fields="mergeFieldPickerFields"
      @insert="insertMergeFieldFromPicker"
      @cancel="mergeFieldPickerOpen = false"
    />
    <MergeFieldPickerDialog
      v-if="repeatRowDialogOpen"
      heading="Repeat row for each…"
      :fields="repeatRowFieldOptions"
      @insert="confirmRepeatRow"
      @cancel="repeatRowDialogOpen = false"
    />
    <MergeRunDialog
      v-if="mergeRunDialogOpen && doc"
      :model="doc"
      :initial-source="mergeDataSource ?? undefined"
      @attach-source="attachMergeDataSource"
      @close="mergeRunDialogOpen = false"
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
      class="print-exclude"
      :left="pasteChip.left"
      :top="pasteChip.top"
      :applied="pasteChip.applied"
      @select="handlePasteOption"
      @dismiss="dismissPasteChip"
    />
    <ImageToolbar
      v-if="imageSelection && imageHandles"
      class="print-exclude"
      :left="imageToolbarStyle().left"
      :top="imageToolbarStyle().top"
      :wrap="imageAt(doc, imageSelection)?.wrap || 'inline'"
      @wrap="handleImageWrap"
      @options="openImageOptions"
      @delete="handleImageDelete"
    />
    <ShapeToolbar
      v-if="shapeSelection && shapeHandles && doc && shapeAt(doc, shapeSelection)"
      class="print-exclude"
      :left="shapeToolbarStyle().left"
      :top="shapeToolbarStyle().top"
      :fill="shapeAt(doc, shapeSelection)?.fill || ''"
      :outline-color="shapeAt(doc, shapeSelection)?.outline?.color || ''"
      :outline-width-pt="shapeAt(doc, shapeSelection)?.outline?.widthPt ?? 1"
      :outline-dash="shapeAt(doc, shapeSelection)?.outline?.dash || 'solid'"
      :rotation-deg="shapeAt(doc, shapeSelection)?.rotationDeg ?? 0"
      :wrap="shapeAt(doc, shapeSelection)?.wrap || 'inline'"
      :alt-text="shapeAt(doc, shapeSelection)?.altText || ''"
      :grouped="shapeAt(doc, shapeSelection)?.groupOrdinal !== undefined"
      :can-group="canGroupSelection()"
      @fill="handleShapeFill"
      @outline="handleShapeOutline"
      @rotation="handleShapeRotationInput"
      @wrap="handleShapeWrap"
      @alt-text="handleShapeAltText"
      @z-order="handleShapeZOrder"
      @group="handleGroupSelected(currentGroupMembers())"
      @ungroup="handleUngroupSelected"
      @delete="handleShapeDelete"
      @duplicate="handleShapeDuplicate"
    />
    <EditorContextMenu
      v-if="contextMenu.visible"
      class="print-exclude"
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
    <SpellContextMenu
      v-if="spellContextMenu.visible"
      class="print-exclude"
      :left="spellContextMenu.left"
      :top="spellContextMenu.top"
      :word="spellContextMenu.word"
      :suggestions="spellContextMenu.suggestions"
      @suggestion="applySpellSuggestion"
      @ignore="spellIgnore"
      @ignore-all="spellIgnoreAll"
      @add-to-dictionary="spellAddToDictionary"
      @dismiss="closeSpellContextMenu"
    />
    <CustomDictionaryDialog
      v-if="customDictionaryDialogOpen"
      :entries="customDictionaryList"
      :language-options="SUPPORTED_LANGUAGES"
      :default-language="currentLanguageTag"
      @add="handleCustomDictionaryAdd"
      @remove="handleCustomDictionaryRemove"
      @close="closeCustomDictionaryDialog"
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
    <TocOptionsDialog
      v-if="tocOptionsDialogOpen"
      :initial="defaultTocOptions()"
      @apply="handleTocInsert"
      @cancel="handleTocCancel"
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
    <StatisticsDialog
      v-if="statisticsOpen"
      :stats="statisticsStats"
      :pages="pageCount"
      :include-notes="statisticsIncludeNotes"
      @update-include-notes="statisticsIncludeNotes = $event"
      @close="closeStatistics"
    />
    <PropertiesDialog
      v-if="propertiesOpen"
      :model="doc"
      :stats="statisticsStats"
      :pages="pageCount"
      :file-size="propertiesFileSize"
      :author-identity="commentAuthorName"
      @update-properties="updatePropertiesRespectingCeiling"
      @update-author-identity="updateAuthorIdentityRespectingCeiling"
      @close="closeProperties"
    />
    <StyleEditorDialog
      v-if="styleEditorOpen"
      :mode="styleEditorMode"
      v-bind="styleEditorInitialProps"
      :based-on-options="styleEditorBasedOnOptions"
      :next-options="styleEditorBasedOnOptions"
      :existing-names="styleEditorExistingNames"
      @save="saveStyleEditor"
      @cancel="closeStyleEditor"
    />
    <StyleDeleteDialog
      v-if="styleDeleteTargetId"
      :style-name="styleDeleteTargetName"
      :in-use="styleDeleteTargetInUse"
      @confirm="confirmDeleteStyle"
      @cancel="cancelDeleteStyle"
    />
    <ImportStylesDialog
      v-if="importDialogOpen"
      :phase="importPhase"
      :file-name="importFileName"
      :error="importError"
      :non-conflicting-count="importNonConflictingCount"
      :conflicts="importConflictRows"
      :resolutions="importResolutions"
      :summary="importSummaryState ?? undefined"
      @choose-file="handleImportChooseFile"
      @set-resolution="setImportResolution"
      @import="confirmImportStyles"
      @close="closeImportDialog"
    />
    <SpecialCharacterDialog
      v-if="specialCharacterDialogOpen"
      @insert="handleSpecialCharacterInsert"
      @cancel="specialCharacterDialogOpen = false"
    />
    <ImageOptionsDialog
      v-if="imageOptionsOpen"
      :initial="imageOptionsInitial"
      @apply="handleImageOptionsApplyAndRescan"
      @cancel="imageOptionsOpen = false"
    />
    <PrintPreviewDialog
      v-if="printPreviewOpen"
      :page-count="printablePageCount"
      :range-mode="printRangeMode"
      :range-text="printRangeText"
      :error="printRangeError ?? ''"
      :can-print="canPrint"
      @update:range-mode="printRangeMode = $event"
      @update:range-text="printRangeText = $event"
      @print="confirmPrint"
      @close="closePrintPreview"
    />
    <!-- PDF export progress (AC-22): a passive status chip — pointer-events
         off, so the UI stays fully interactive underneath (AC-23). -->
    <div
      v-if="pdfExporting"
      class="pdf-export-progress print-exclude"
      role="status"
      data-testid="pdf-export-progress"
    >
      Exporting PDF…{{ pdfProgress.total > 0 ? ` page ${pdfProgress.done} of ${pdfProgress.total}` : "" }}
    </div>
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

/* The Insert menu's table-size quick-pick (SPEC-049), slotted into the menu
   popup. These rules reach it because slotted content keeps this view's
   scope id. */
.menu-image-input {
  display: none;
}

.menu-table-grid-wrap {
  padding: 0.25rem 1rem 0.5rem;
}

.menu-table-grid {
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.menu-table-grid-row {
  display: flex;
  gap: 1px;
}

.menu-table-grid-cell {
  width: 0.8125rem;
  height: 0.8125rem;
  padding: 0;
  border: 1px solid #bdbdbd;
  border-radius: 2px;
  background: #fff;
  cursor: pointer;
}

.menu-table-grid-cell.active {
  background: #bbdefb;
  border-color: #1a73e8;
}

.menu-table-grid-cell:disabled {
  cursor: default;
  opacity: 0.5;
}

.menu-table-grid-label {
  margin-top: 0.1875rem;
  font-size: 0.78rem;
  color: #5f6368;
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

/* Sits in the status bar next to the page indicator (SPEC-036 AC-1, AC-4):
   a plain-looking button (no border/background of its own) so it reads as
   status text until focused/hovered, while still being a real <button> —
   keyboard-reachable by Tab and activatable with Enter/Space with no extra
   keydown handling (AC-27). */
.word-count-button {
  font: inherit;
  font-size: 0.75rem;
  color: inherit;
  padding: 0;
  border: none;
  background: none;
  cursor: pointer;
}

.word-count-button:hover,
.word-count-button:focus-visible {
  text-decoration: underline;
}

/* SPEC-039 AC-1: a clear on/off state in both the toolbar and the status
   bar — the toolbar button reads as a toggle chip, the status-bar one
   matches word-count-button's plain-until-hovered look. */
.track-changes-toolbar-toggle {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.6rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: none;
  cursor: pointer;
}

.track-changes-toolbar-toggle[aria-pressed="true"] {
  background: #e6f4ea;
  color: #188038;
  border-color: #188038;
}

.track-changes-status-indicator {
  font: inherit;
  font-size: 0.75rem;
  color: inherit;
  padding: 0;
  border: none;
  background: none;
  cursor: pointer;
}

.track-changes-status-indicator[aria-pressed="true"] {
  text-decoration: underline;
  font-weight: 600;
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

/* The outline pane (SPEC-037) is a left sidebar that OVERLAYS the canvas
   rather than reflowing it: the pane sits in an absolutely-positioned dock on
   the left edge, while page-wrap spans the full width beneath. Keeping the
   page column full-width (and so its page centred within the whole canvas)
   preserves the SPEC-002 invariant that a blank document's page is centred in
   the editor viewport even while the sidebar is open. The pane itself stays
   position:sticky (see OutlinePane.vue) so it stays visible while the pages
   scroll; the dock stretches to the content height to give it a stick range. */
.editor-layout {
  position: relative;
  min-height: 100%;
}

/* Full-width page column; the absolute dock sits over its left edge. */
.page-wrap {
  position: relative;
  width: 100%;
  min-width: 0;
  box-sizing: border-box;
  padding: 2rem 1rem;
}

/* The sticky rail hosting the pane. absolute so it never reflows the page
   column; top:0/bottom:0 make it as tall as the content so its sticky child
   (OutlinePane) stays pinned for the whole scroll. */
.outline-dock {
  position: absolute;
  top: 0;
  left: 0;
  bottom: 0;
  width: 16rem;
  z-index: 20;
  /* The dock floats over the full-width, fully-centred page column. It must
     not swallow pointer events over the page (typing/caret, the ruler's
     left-edge tab selector, right-click) — only the outline tree itself is
     interactive; the rest passes through to the page beneath. */
  pointer-events: none;
}

.outline-dock :deep(.outline-tree) {
  pointer-events: auto;
}

/* The comments pane's own dock (SPEC-038) — the outline-dock's mirror on
   the right edge of the page column. */
.comments-dock {
  position: absolute;
  top: 0;
  right: 0;
  bottom: 0;
  width: 18rem;
  z-index: 20;
  pointer-events: none;
}

.comments-dock :deep(.comments-pane) {
  pointer-events: auto;
}

/* The Styles pane's own dock (SPEC-042) — same shape as comments-dock, on
   the same edge (the panes are not normally opened together). */
.styles-dock {
  position: absolute;
  top: 0;
  right: 0;
  bottom: 0;
  width: 18rem;
  z-index: 20;
}

/* The track-changes review pane's own dock (SPEC-039) — same shape as
   comments-dock, on the same edge (the two panes are not normally opened
   together). */
.track-changes-dock {
  position: absolute;
  top: 0;
  right: 0;
  bottom: 0;
  width: 18rem;
  z-index: 20;
  pointer-events: none;
}

.track-changes-dock :deep(.track-changes-pane) {
  pointer-events: auto;
}

/* The version history pane's own dock (SPEC-040) — same shape as the other
   two, on the same edge (not normally opened together with them). */
.version-history-dock {
  position: absolute;
  top: 0;
  right: 0;
  bottom: 0;
  width: 20rem;
  z-index: 20;
  pointer-events: none;
}

.version-history-dock :deep(.version-history-pane) {
  pointer-events: auto;
}

/* The accessibility checker's own dock (SPEC-043) — same shape as the other
   docked panes, on the same edge (not normally opened together with them). */
.accessibility-dock {
  position: absolute;
  top: 0;
  right: 0;
  bottom: 0;
  width: 20rem;
  z-index: 20;
  pointer-events: none;
}

.accessibility-dock :deep(.accessibility-pane) {
  pointer-events: auto;
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

/* The spell-check underline overlay (SPEC-035 AC-1). Unlike
   .find-highlight-layer (intentionally *behind* the text it highlights),
   this must paint ABOVE the page: DocPage's own .page root is itself
   `position: relative` with no z-index (stack level 0, CSS2.1 Appendix E
   step 6), the same tier a sibling z-index:0 layer occupies — ties within
   that tier resolve by DOM order, and .page (containing every page's
   opaque background and all its text) comes after this layer in the
   markup, so a bare z-index:0 here would paint BEHIND every page's
   background and be invisible (verified by screenshot while building this:
   z-index:0 here is empirically hidden). z-index:1 moves this layer into
   the next stacking tier (CSS2.1 step 7, positive z-index), which always
   paints after every stack-level-0 box regardless of DOM order — still well
   under the table/image handle layers' z-index:15. */
.spell-underline-layer {
  position: absolute;
  inset: 0;
  z-index: 1;
  pointer-events: none;
}

.spell-underline {
  position: absolute;
  /* Inert, and it must stay that way. Each rect covers a flagged word's whole
     glyph box (not a 3px sliver), and the layer paints above the text — so
     opting into pointer events here makes the overlay swallow every plain
     click, double-click and drag that starts on a misspelled word, leaving
     the caret unset and the selection empty. Right-clicks are hit-tested
     against these same rects in spellRectAtPoint() instead. */
  pointer-events: none;
  background-repeat: repeat-x;
  background-position: left bottom;
  background-size: 6px 4px;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='6' height='4' viewBox='0 0 6 4'%3E%3Cpath d='M0 3 Q1.5 0 3 3 T6 3' fill='none' stroke='%23d93025' stroke-width='1'/%3E%3C/svg%3E");
}

/* The comment highlight overlay (SPEC-038): same layering convention as
   .find-highlight-layer, but each rect IS clickable (AC "clicking a
   highlighted range focuses its thread"), so pointer-events stays auto on
   the individual rects while the layer itself still passes clicks through
   the gaps between them. */
.comment-highlight-layer {
  position: absolute;
  inset: 0;
  /* Unlike .find-highlight-layer (z-index: 0, never needs to catch clicks),
     these rects must win hit-testing against the text underneath — the
     same z-index band .table-handles (15) uses for its own clickable
     overlay controls. */
  z-index: 5;
  pointer-events: none;
}

.comment-highlight {
  position: absolute;
  background: rgb(26 115 232 / 0.18);
  border-bottom: 2px solid rgb(26 115 232 / 0.55);
  pointer-events: auto;
  cursor: pointer;
}

.comment-highlight.active {
  background: rgb(26 115 232 / 0.32);
  border-bottom-color: #1a73e8;
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

/* The selected shape's eight resize handles plus rotation handle (SPEC-051
   AC-10, AC-16) — same overlay convention as .image-handles. */
.shape-handles {
  position: absolute;
  inset: 0;
  /* Above .shape-toolbar (z-index: 30, ShapeToolbar.vue) — the toolbar
     renders 44px above the shape (shapeToolbarStyle) and the rotation
     handle sits 24px above it (ROTATION_HANDLE_OFFSET_PX), so their boxes
     always overlap for a normally sized toolbar; without this the handle
     is visually present but never receives the click/drag that starts a
     rotation (AC-16, AC-38). */
  z-index: 31;
  pointer-events: none;
}

.shape-resize-handle {
  position: absolute;
  pointer-events: auto;
  width: 8px;
  height: 8px;
  border: 1px solid #1a73e8;
  border-radius: 2px;
  background: #fff;
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

.shape-rotation-handle {
  position: absolute;
  pointer-events: auto;
  width: 8px;
  height: 8px;
  border: 1px solid #1a73e8;
  border-radius: 50%;
  background: #fff;
  cursor: grab;
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

/* The anchor marker (SPEC-051 AC-27): a small flag at the start of the
   paragraph a selected shape is anchored to, moving with it on re-anchor
   (AC-28) — same overlay convention as the resize/rotation handles, one
   layer below them (z-index) since it never needs to receive a pointer. */
.shape-anchor-marker {
  position: absolute;
  z-index: 14;
  pointer-events: none;
  width: 10px;
  height: 10px;
  margin-left: -5px;
  margin-top: -2px;
  background: #1a73e8;
  clip-path: polygon(0 0, 100% 50%, 0 100%);
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

/* The PDF export progress chip (SPEC-034 AC-22): floats over the canvas
   without capturing the pointer, so scrolling/clicking keep working while
   the export runs (AC-23). */
.pdf-export-progress {
  position: fixed;
  bottom: 3rem;
  left: 50%;
  transform: translateX(-50%);
  z-index: 130;
  padding: 0.5rem 1.25rem;
  border-radius: 999px;
  background: #202124;
  color: #fff;
  font-size: 0.875rem;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.3);
  pointer-events: none;
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
