<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, provide, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute } from "vue-router";
import AppShell from "../components/AppShell.vue";
import DocumentMenuBar from "../components/DocumentMenuBar.vue";
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
import { SUPPORTED_LANGUAGES } from "../lib/core/language";
import HeaderFooterOptionsDialog from "../components/HeaderFooterOptionsDialog.vue";
import ImageOptionsDialog from "../components/ImageOptionsDialog.vue";
import ImageToolbar from "../components/ImageToolbar.vue";
import ShapeToolbar from "../components/ShapeToolbar.vue";
import GroupToolbar from "../components/GroupToolbar.vue";
import ShapeDrawPreview from "../components/ShapeDrawPreview.vue";
import ChartToolbar from "../components/ChartToolbar.vue";
import ChartDataSheetDialog from "../components/ChartDataSheetDialog.vue";
import ChartOptionsDialog from "../components/ChartOptionsDialog.vue";
import InsertTableDialog from "../components/InsertTableDialog.vue";
import ColumnsDialog from "../components/ColumnsDialog.vue";
import CommentsPane from "../components/CommentsPane.vue";
import TrackChangesPane from "../components/TrackChangesPane.vue";
import AccessibilityPane from "../components/AccessibilityPane.vue";
import FontsPane from "../components/FontsPane.vue";
import FontLoadingSkeleton from "../components/FontLoadingSkeleton.vue";
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
import FieldDialog from "../components/FieldDialog.vue";
import DateTimeDialog from "../components/DateTimeDialog.vue";
import CrossReferenceDialog from "../components/CrossReferenceDialog.vue";
import CaptionDialog from "../components/CaptionDialog.vue";
import TableOfFiguresDialog from "../components/TableOfFiguresDialog.vue";
import DeleteCaptionPrompt from "../components/DeleteCaptionPrompt.vue";
import FieldToolbar from "../components/FieldToolbar.vue";
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
import TableHandlesOverlay from "../components/TableHandlesOverlay.vue";
import TablePropertiesDialog from "../components/TablePropertiesDialog.vue";
import TabStopsDialog from "../components/TabStopsDialog.vue";
import TextColorPicker from "../components/TextColorPicker.vue";
import UndoRedoControls from "../components/UndoRedoControls.vue";
import { useAutosave } from "../composables/useAutosave";
import { useToasts } from "../composables/useToasts";
import { noApplyEditOptions } from "./document/context";
import { useAutocorrect } from "./document/useAutocorrect";
import { useClipboard } from "./document/useClipboard";
import { asCommentFilter, useComments } from "./document/useComments";
import { asChangeEntries, useTrackChanges } from "./document/useTrackChanges";
import { useVersionHistory } from "./document/useVersionHistory";
import { useAccessibilityCheck } from "./document/useAccessibilityCheck";
import { useFonts } from "./document/useFonts";
import { useFeatureHooks } from "./document/useFeatureHooks";
import { useFindReplace } from "./document/useFindReplace";
import { usePdfExport } from "./document/usePdfExport";
import { usePrintPreview } from "./document/usePrintPreview";
import { useFormatPainter } from "./document/useFormatPainter";
import { useHeadersFooters } from "./document/useHeadersFooters";
import { useImages } from "./document/useImages";
import { useShapes } from "./document/useShapes";
import { useDrawingGroups } from "./document/useDrawingGroups";
import { useCharts } from "./document/useCharts";
import { useLinks } from "./document/useLinks";
import { useNotes } from "./document/useNotes";
import { useFormattingSurfaces } from "./document/useFormattingSurfaces";
import { noTablePropertiesInitial, useTables } from "./document/useTables";
import { usePaginationEngine } from "./document/usePaginationEngine";
import { usePaginationView } from "./document/usePaginationView";
import { useFields } from "./document/useFields";
import { useCaptions } from "./document/useCaptions";
import { renumberCaptions } from "../lib/captions";
import { useEquations } from "./document/useEquations";
import { usePageNavigation } from "./document/usePageNavigation";
import { useWordCount } from "./document/useWordCount";
import { useDocumentProperties } from "./document/useDocumentProperties";
import { useSpellCheck } from "./document/useSpellCheck";
import { useDocumentLanguage } from "./document/useDocumentLanguage";
import { useStyleManager } from "./document/useStyleManager";
import { useContextMenus } from "./document/useContextMenus";
import { useFileActions } from "./document/useFileActions";
import { usePageSetupDialogs } from "./document/usePageSetupDialogs";
import { useOutlinePane } from "./document/useOutlinePane";
import { useMergeFields } from "./document/useMergeFields";
import { useViewportVirtualization } from "./document/useViewportVirtualization";
import { useCanvasGestures } from "./document/useCanvasGestures";
import { documentExists } from "../lib/documents";
import {
  getOrCreateDocument,
  replaceDocumentModel,
  touchDocument,
} from "../lib/documentStore";
import { fieldAtCaret, fieldSpanningSelection } from "../lib/commands.fields";
import { isReadOnlyField, noFieldSpec } from "../lib/fields";
import {
  asHeaderFooterPart,
  asImage,
  isFieldLocked,
  asImageWrapMode,
  asShape,
  asChart,
  asNoteKind,
  asNoteNumberFormat,
  asNoteRestart,
  asTrackChangesDisplayMode,
  asPageNumberFormat,
  asParagraphBlock,
  asBreakKind,
  asSectionStart,
  createDocument,
  createEmptyHeaderFooterContent,
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
  noNoteContent,
} from "../lib/documentModel";
import { asChartType, CHART_CATALOGUE } from "../lib/core/charts";
import { asShapeGeometry } from "../lib/core/shapes";
import {
  asBorderLineStyle,
  asCellVerticalAlign,
  asTableAlignment,
  asTableWidthMode,
} from "../lib/core/tables";
import {
  defaultTocOptions,
  insertTableOfContents,
  noTocSelection,
  tocHeadingsStale,
  tocPageNumbersStale,
  updateTocEntirely,
  updateTocPageNumbersOnly,
} from "../lib/toc";
import { listMarkers } from "../lib/lists";
import { detectListAutoformat } from "../lib/core/lists";
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
  caretTableLocation,
  deleteTable,
  deleteTableColumn,
  deleteTableRow,
  fullTableRange,
  insertTable,
  insertTableColumn,
  insertTableRow,
  mergeCells,
  nextCellCaret,
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
  MIN_COLUMN_WIDTH_CM,
  MIN_ROW_HEIGHT_CM,
  noCellRange,
} from "../lib/core/tables";
import { acquireEditLock, noEditLock } from "../lib/core/editLock";
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
  noParagraphTarget,
  noTableCellTarget,
  paragraphElementForCaret,
  setDomCaret,
  setDomSelection,
} from "../lib/domCaret";
import {
  findMatches,
  indexAtOrAfter,
  isQueryError,
  matchSelection,
  noHighlightRects,
  noMatches,
} from "../lib/findReplace";
import {
  caretClientRect,
  noTableHandleGeometry,
} from "../lib/caretGeometry";
import { findTableElement, noColumnResizeDrag, noRowResizeDrag } from "../lib/tableResize";
import { previewColumnWidths, previewRowHeight, tableRowElement } from "../lib/core/tables";
import { CM_PER_PX } from "../lib/core/units";
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
} from "../lib/clipboard";
import { importHtml } from "../lib/htmlPaste";
import {
  adjustIndent,
  applyListAutoformat,
  applyPainterFormat,
  toolbarCharacterStyleState,
  blockCount,
  blockOrdinal,
  capturePainterFormat,
  caretSelection,
  changeListLevel,
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
  noCommandResult,
  noPainterFormat,
  noPendingFormat,
  noSelection,
  paragraphPaginationState,
  removeListFormat,
  setHeaderFooterOptions,
  setNoteSettings,
  splitBlock,
  toolbarFormatValues,
  toolbarListState,
  toolbarStyleState,
  touchedBlocks,
  verticalAlignState,
} from "../lib/commands";
import { effectiveNoteSettings } from "../lib/notes";
import { headingTag, isExternalLink, paragraphStyle, runKey, runLinkHref, runStyle } from "../lib/renderFormatting";
import { markSegments } from "../lib/core/marks";
import { noIdleHandle, scheduleIdleWork } from "../lib/core/idleScheduler";
import { sectionColumnFlowWidthPx, wholePageItem } from "../lib/pagination";
import { noNoteArea, noNoteAreaItems } from "../lib/notesLayout";
import { fieldValueText, headerFooterForPage, resolveHeaderFooterKind } from "../lib/pageNumbering";
import { noRenderEntries } from "../lib/pageRender";
import {
  deleteImage,
  imageAt,
  insertImage,
  noImageCommand,
  noImageLocation,
  setImageAltText,
  setImageSize,
  setImageWrap,
} from "../lib/images";
import { pxToCm } from "../lib/core/units";
import { prepareImageInsert, ImageInsertError } from "../lib/imageInsert";
import { shapeAt } from "../lib/shapes";
import { asNudgeDirection } from "../lib/core/shapes";
import { chartAt } from "../lib/charts";
import {
  asImageHandle,
  IMAGE_HANDLES,
  noImageResizeDrag,
  previewImageSize,
  resizedImageSize,
} from "../lib/imageResize";
import { headingStyleId, NORMAL_STYLE_ID } from "../lib/styles";
import StylesPane from "../components/StylesPane.vue";
import StyleEditorDialog from "../components/StyleEditorDialog.vue";
import StyleDeleteDialog from "../components/StyleDeleteDialog.vue";
import ImportStylesDialog from "../components/ImportStylesDialog.vue";
import EquationEditorHost from "../components/EquationEditorHost.vue";
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
import { isMacPlatform } from "../lib/core/platform";

const route = useRoute();
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
// The registry (session close, unmount, scroll, post-edit find refresh) lives
// in useFeatureHooks; the backbone fires the exposed Sets at the right moments.
const { sessionCloseHooks, unmountHooks, scrollHooks, onSessionClose, onUnmount, onScroll, registerFindRefresh, fireFindRefresh } =
  useFeatureHooks();

// --- tables (SPEC-022) ------------------------------------------------------
// The rectangular cell-range selection (AC-15, AC-16). The ref stays here —
// useTables' wiring and the template consume it — while the drag gesture
// that populates it (and its drag-origin tracking) lives in useCanvasGestures.
const cellSelection = ref(noCellRange());
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

// --- table formatting (SPEC-023) --------------------------------------------
// "Split cell…" and the table properties dialog both act on the cell/table
// the context menu was opened against — captured into these refs when the
// dialog opens, since the menu itself (and the click that opened it) is
// long closed by the time the dialog's OK button fires.
const splitCellDialogOpen = ref(false);
const tablePropertiesOpen = ref(false);
const tablePropertiesInitial = ref(noTablePropertiesInitial());
const repeatRowDialogOpen = ref(false);
let tableFormattingLocation = { section: 0, block: 0, row: 0, col: 0 };
// The column/row resize drag state now lives inside useTables (it owns the
// live-resize handlers and their listener lifecycle).
// canMergeCells / clearCellSelection / tableCellAtTarget live in
// useCanvasGestures, wired below with the rest of the canvas gesture hub.

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
// The field under a collapsed caret (SPEC-053 AC-14) — drives the field
// toolbar chip. Null when the selection is a range or not on a field.
const selectedField = shallowRef(noFieldSpec());

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
// style, AC-1); the picker's menu entries and its actions live in
// useStyleManager.
const styleState = ref(emptyToolbarStyleState());
// --- custom styles manager (SPEC-042) ---------------------------------------
// The caret's active character style (AC-10's second readout, alongside
// styleState above); the Styles pane/dialogs' own state lives in
// useStyleManager.
const characterStyleState = ref(emptyToolbarStyleState());
// The formatting dialogs' state (custom spacing, define-list, tabs) and the
// recent-colours list live in useFormattingSurfaces, wired below.
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

// The caret's focus end, mirrored reactively (the live `caret` is a plain
// variable) so the outline pane's active-heading highlight tracks caret and
// scroll (SPEC-037 AC-3, AC-4). Updated in updateCaretStatus.
const caretPosition = ref(noCaret());

// Forces paragraph re-mounts by folding into the v-for keys: `key` names the
// single block whose key carries `count` (bumped after an IME commit so the
// browser-composed DOM nodes are discarded for a fresh model render), and
// `epoch` re-keys the whole canvas (the recovery hatch when a composition
// left the DOM diverged from the model).
const remount = ref({ key: "", count: 0, epoch: 0 });

const contentRoot = useTemplateRef("contentRoot");
const pageRef = useTemplateRef("pageRef");

// The one object satisfying EditorContext (document/context.ts): the
// universal dependencies every feature composable may rely on. Each use*()
// call below spreads this and adds only its feature-specific extras, so
// wiring a new feature never re-plumbs the core. The function fields are
// hoisted declarations (defined further down this file), so referencing
// them here — before the composables that call them — is safe; the getters
// wrap plain variables that must be read live, never snapshotted.
const editorContext = {
  doc,
  documentId,
  contentRoot,
  pageRef,
  editable,
  currentSelection,
  applyEdit,
  applyCaretMove,
  scrollModelCaretIntoView,
  getCaret: () => caret,
  getScrollContainer: () => scrollContainer,
  // Deferred behind a lambda: the const is destructured from
  // useCanvasGestures, wired after useLinks below — every consumer calls it
  // at event/edit time (past setup), never during its deps evaluation.
  clearCellSelection: () => clearCellSelection(),
  addToast,
  markDirty: () => noteEdit(),
  refreshUndoRedoState,
  onSessionClose,
  onUnmount,
  onScroll,
};

// The format painter (SPEC-020) lives in useFormatPainter; handleSelectionChange
// and the canvas mouseup dispatch into its maybeApplyPainter.
const {
  painterMode,
  disarmPainter,
  handlePainterClick,
  handlePainterDblClick,
  maybeApplyPainter,
} = useFormatPainter({ ...editorContext });

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
  ...editorContext,
  getPaginationPages: () => paginationPages.value,
  getPageNumbers: () => pageNumbers.value,
  getCaretSectionIndex: () => caretSectionIndex.value,
  pageOfCaret: (caretPos) => pagination.pageOfCaret(caretPos),
  getClearImageSelection: () => clearImageSelection,
  getExitNoteEditing: () => exitNoteEditing,
  getUpdateFieldInSurface: () => updateFieldInSurface,
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
  ...editorContext,
  activeHeaderFooter,
  currentHfSelection,
  applyHeaderFooterEdit,
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
  drawPreview,
  movingShape,
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
  handleShapeKeys,
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
} = useShapes({ ...editorContext, getUpdateFieldInSurface: () => updateFieldInSurface });

// --- charts (SPEC-052) -------------------------------------------------------
// Chart selection/handles/toolbar, insertion at the caret and the data
// sheet/options-panel commit paths live in useCharts, trimmed from
// useShapes' own pattern (no drag/resize/rotate/multi-select-group gesture
// — see that composable's header comment).
const {
  chartSelection,
  handleChartKeys,
  chartHandles,
  editingChartData,
  editingChartOptions,
  clearChartSelection,
  remeasureChartHandles,
  handleChartSelect,
  handleChartEditData,
  openChartOptions,
  insertChartAction,
  insertChartFromTableAction,
  applyChartData,
  applyChartOptions,
  applyChartSeriesColor,
  handleChartDelete,
  handleChartDuplicate,
  chartToolbarStyle,
  chartAtClientPoint,
  chartDataSheetInitial,
  chartOptionsInitial,
} = useCharts({ ...editorContext });

// --- native drawing-group containers (SPEC-057) ------------------------------
// Container/member selection (click-to-enter), Group/Ungroup, the container
// and member gestures, the container toolbar and keyboard proxy traversal.
// clearOtherSelections keeps a group selection exclusive with the per-kind
// shape/image/chart selections.
const drawingGroups = useDrawingGroups({ ...editorContext, clearOtherSelections: () => { clearShapeSelection(); clearImageSelection(); clearChartSelection(); } });
const groupSelection = drawingGroups.groupSelection;

// Equations (SPEC-056): Insert → Equation / Display equation and Alt+=. The
// session's editing surface, op dispatch and overlay live in the composable
// and EquationEditorHost.vue; DocumentView keeps only the insert commands, the
// clicked-equation `provide`, the character-picker route, and the host tag.
const equations = useEquations({ ...editorContext });

// SPEC-056: an equation clicked in the document opens the editor. Provided to
// the paint components (DocPageBlock/DocBlockEntries) which own the click.
provide("activateEquation", equations.activateEquation);

// SPEC-056: route a character-picker glyph into an open equation (AC-14/15);
// otherwise it inserts into the document as before.
function onSpecialCharacterInsert(text = "") {
  if (equations.insertPickerGlyph(text)) {
    specialCharacterDialogOpen.value = false;
    return;
  }
  handleSpecialCharacterInsert(text);
}

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
  ...editorContext,
  activeHeaderFooter,
  clearImageSelection,
  exitHeaderFooterEditing,
  getUpdateFieldInSurface: () => updateFieldInSurface,
});

// --- character & paragraph formatting (SPEC-012, SPEC-013, SPEC-014) --------
// One implementation of the toolbar/shortcut formatting commands, their
// dialogs and the ruler handler, dispatched over the active editing surface —
// body, header/footer band (SPEC-029 AC-4) or note (SPEC-031 AC-6) — in
// useFormattingSurfaces, replacing the per-surface triads that lived here.
// Wired directly after the header/footer and note bundles that complete its
// surface table, and before useStyleManager, which passes
// applyParagraphCommand eagerly. `pendingFormat` and the toolbar readouts
// stay in this view (updateCaretStatus recomputes them); the surfaces reach
// the body's pending-format store through the get/set lambdas.
const {
  toggleCharacterFormat,
  handleFontFamilySelect,
  handleFontSizeChange,
  handleTextColorSelect,
  handleHighlightSelect,
  recentColors,
  applyParagraphCommand,
  handleAlign,
  handleIndent,
  handleClearFormatting,
  handleSpacingPreset,
  handleToggleSpaceBefore,
  handleToggleList,
  handleListMarker,
  handleMultilevelPreset,
  handleLinkHeadingNumbering,
  defineListOpen,
  defineListInitial,
  openDefineList,
  handleDefineListApply,
  tabsDialogOpen,
  tabsDialogInitial,
  tabsDialogDefaultInterval,
  openTabsDialog,
  handleTabsApply,
  handleToggleDirection,
  customSpacingOpen,
  customSpacingInitial,
  openCustomSpacing,
  handleCustomSpacingApply,
  handleRulerUpdate,
  handleTogglePagination,
} = useFormattingSurfaces({
  ...editorContext,
  paragraphSelection,
  getPendingFormat: () => pendingFormat,
  setPendingFormat: (next) => {
    pendingFormat = next;
  },
  updateCaretStatus,
  activeHeaderFooter,
  currentHfSelection,
  applyHeaderFooterEdit,
  hfRootEl,
  getHfPendingFormat,
  setHfPendingFormat,
  activeNote,
  currentNoteSelection,
  applyNoteEdit,
  noteRootEl,
  getNotePendingFormat,
  setNotePendingFormat,
  getParagraphState: () => paragraphState.value,
  getParagraphFormatting: () => paragraphFormatting.value,
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
  ...editorContext,
  getPendingFormat: () => pendingFormat,
  insertImageFromFile,
});

// Smart-quote/dash autocorrect and the special-character picker (SPEC-026)
// live in useAutocorrect; handleBeforeInput's typed-input branch dispatches
// into applyAutocorrect, and applyHistorySnapshot clears the substitution
// follow-up context through resetSubstitutionContext.
const {
  autocorrectSettings,
  toggleAutocorrectSmartQuotes,
  toggleAutocorrectDashes,
  applyAutocorrect,
  resetSubstitutionContext,
  insertSpecialCharacter,
  specialCharacterDialogOpen,
  openSpecialCharacters,
  handleSpecialCharacterInsert,
} = useAutocorrect({
  ...editorContext,
  getPendingFormat: () => pendingFormat,
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
} = useWordCount({ ...editorContext });

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
  // The first layout is deferred out of the mount task (usePaginationEngine's
  // prime kick): until it lands there are no page heights to scroll within
  // and no block DOM to place the caret in, so a session restore racing it
  // would clamp the scroll to 0 and drop the caret. Wait for the first
  // non-empty layout, then one more tick for its page DOM.
  if (paginationPages.value.length === 0) {
    await new Promise((resolve) => {
      const stopWaiting = watch(paginationPages, (pages) => {
        if (pages.length > 0) {
          stopWaiting();
          resolve(true);
        }
      });
    });
    await nextTick();
  }
  attachScrollListener();
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  // Waiting for the layout means this can now resume AFTER the user has
  // already interacted — opened a dialog, clicked a menu, reached the find
  // bar. Restoring must never steal focus from whatever they moved to:
  // only take it while focus is still neutral (nowhere, or already in the
  // canvas).
  const active = document.activeElement;
  const focusIsNeutral = !active || active === document.body || root.contains(active);
  if (focusIsNeutral) {
    root.focus();
    caret = caret ?? { section: 0, block: 0, offset: 0 };
    setDomCaret(root, caret);
    updateCaretStatus({ start: caret, end: caret });
  }
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
  // The virtualization trackers (viewportPageRange/anchorPage/caretPage)
  // reset via useViewportVirtualization's own session-close hook below.
  dismissPasteChip();
  sessionCloseHooks.forEach((cb) => cb());
  // The gesture state (cell-range selection, drag origin, hover handles)
  // resets via useCanvasGestures' own session-close hook below.
  insertTableDialogOpen.value = false;
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
    // against re-triggering off its own caret-restore echo. The drag flag
    // lives in useCanvasGestures; the getter reads it live, post-setup.
    if (!isMouseButtonDown()) {
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
 * `label` names the edit for the undo history (AC-7..13): every call site
 * names its own action, so recording is never skipped by accident the way
 * a defaulted label could be. The optional behaviours (undo grouping, the
 * paste chip's survival, the track-changes bypass) ride in `options` as
 * named flags — see ApplyEditOptions in document/context.ts. */
function applyEdit(
  result = noCommandResult(),
  beforeSelection = noSelection(),
  label = "",
  options = noApplyEditOptions(),
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
  if (!options.keepPasteChip) {
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
  clearTofSelection();
  const before = doc.value;
  // SPEC-039: with recording on, rewrite the command's result into tracked-
  // change markup before it ever reaches undo history or the model — every
  // command becomes trackable with no per-command wiring (see
  // trackChangesDiff.ts's header comment). `skipTrackChanges` is for the
  // accept/reject commands themselves, whose result IS the final state a
  // reviewer chose, not a new edit to re-diff.
  if (before && !options.skipTrackChanges && isRecording()) {
    result = {
      model: applyTrackChanges(before, result.model, trackChangesAuthorName.value, new Date().toISOString()),
      selection: result.selection,
    };
  }
  // SPEC-055 AC-15/16: reordering captioned objects renumbers in the same undo step.
  result = { model: renumberCaptions(result.model), selection: result.selection };
  if (before) {
    recordHistoryEdit(documentId.value, {
      before: { model: before, selection: beforeSelection },
      after: { model: result.model, selection: result.selection },
      label,
      groupable: options.groupable === true,
    });
  }
  doc.value = replaceDocumentModel(documentId.value, result.model);
  caret = result.selection.focus;
  selectionAnchor = isCollapsed(result.selection) ? noCaret() : result.selection.anchor;
  noteEdit();
  updateCaretStatus({ start: result.selection.anchor, end: result.selection.focus });
  refreshUndoRedoState();
  if (!options.skipCaretRestore) {
    void restoreCaretAfterRender();
  }
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
  fireFindRefresh();
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
  resetSubstitutionContext();
  noteEdit();
  updateCaretStatus({ start: snapshot.selection.anchor, end: snapshot.selection.focus });
  refreshUndoRedoState();
  root.focus();
  void restoreCaretAfterRender();
  // Undo/redo can move the caret a long way — undoing an inserted page
  // break pulls it back a whole page — so follow it, the way the break
  // insertion it reverses does.
  void scrollModelCaretIntoView(caret);
  fireFindRefresh();
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
          applyEdit(result, selection, label, { groupable });
          const autoformatSelection = caretSelection(caretAfter);
          applyEdit(
            applyListAutoformat(result.model, autoformatSelection, autoformatType),
            autoformatSelection,
            autoformatType === "bullet" ? "bulleted list" : "numbered list",
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
    applyEdit(result, selection, label, { groupable });
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
  applyEdit(insertText(model, range, data, pendingFormat), range, "typing", { groupable: true });
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

// Clicks on rendered hyperlinks and note reference marks are handled by
// useCanvasGestures' handleContentClick, wired below.

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
    selectedField.value = null;
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
  // SPEC-053 AC-14: the field chip shows for a collapsed caret on a field, or a
  // selection spanning exactly one field — what clicking a field produces (AC-7).
  const collapsed = compareCaret(range.start, range.end) === 0;
  selectedField.value = collapsed
    ? fieldAtCaret(model, range.end)
    : fieldSpanningSelection(model, range.start, range.end);
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
  ...editorContext,
  // Deferred behind a lambda: the const is destructured from
  // useViewportVirtualization, wired after pagination below — the wrapper
  // reads it at call time (a paging key handler, past setup), never during
  // this deps object's evaluation.
  refreshViewportRange: () => refreshViewportRange(),
});

// --- tables (SPEC-022, SPEC-023) ---------------------------------------------
// Row/column insert-delete, merge/split, the properties dialog, and live column
// /row resize live in useTables; the hover handles' geometry and the cell-range
// drag selection live in useCanvasGestures' canvas event hub, wired below.
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
  ...editorContext,
  paragraphSelection,
  activeHeaderFooter,
  currentHfSelection,
  applyHeaderFooterEdit,
  hfRootEl,
  // Deferred behind a lambda: the const is destructured from useContextMenus,
  // wired after useSpellCheck below — the wrapper reads it at call time (an
  // event handler, past setup), never during this deps object's evaluation.
  closeContextMenu: () => closeContextMenu(),
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
  // Deferred behind a lambda (SPEC-055): requestObjectDelete is destructured
  // from useCaptions, wired after useLinks below — invoked only at
  // delete-table time, never during this deps object's evaluation.
  requestObjectDelete: (object, kind, perform) => requestObjectDelete(object, kind, perform),
});

/** The selection paragraph commands target: the live/tracked selection, or
 * — before the canvas has ever held a caret — the document start, so a
 * toolbar click or ruler drag on a fresh document still lands on the first
 * paragraph. */
function paragraphSelection(root = noElement()) {
  return currentSelection(root) ?? caretSelection(documentStart());
}

// --- paragraph styles and the custom styles manager (SPEC-019, SPEC-042) ---
// The style picker's actions, the Styles pane, the style editor/delete
// dialogs and the import-styles flow live in useStyleManager; handleKeyDown's
// Ctrl/Cmd+Alt+0..6 shortcuts dispatch into handleApplyStyle. caretPosition
// comes in as a getter over the ref so styleEditorInitialProps re-evaluates
// when the caret moves.
const {
  styleEntries,
  stylesPaneOpen,
  stylesPaneEntriesList,
  styleEditorOpen,
  styleEditorMode,
  styleDeleteTargetId,
  styleDeleteTargetInUse,
  importDialogOpen,
  importPhase,
  importFileName,
  importError,
  importResolutions,
  importSummaryState,
  styleEditorInitialProps,
  styleEditorBasedOnOptions,
  styleEditorExistingNames,
  styleDeleteTargetName,
  importConflictRows,
  importNonConflictingCount,
  handleApplyStyle,
  handleUpdateStyle,
  handleResetStyle,
  toggleStylesPane,
  closeStylesPane,
  openNewStyle,
  openModifyStyle,
  closeStyleEditor,
  saveStyleEditor,
  openDeleteStyle,
  cancelDeleteStyle,
  confirmDeleteStyle,
  applyStyleFromPane,
  openImportDialog,
  closeImportDialog,
  handleImportChooseFile,
  setImportResolution,
  confirmImportStyles,
} = useStyleManager({
  ...editorContext,
  paragraphSelection,
  applyParagraphCommand,
  setPendingFormat: (next) => {
    pendingFormat = next;
  },
  updateCaretStatus,
  getCaretPosition: () => caretPosition.value,
});

// --- page setup, columns and ruler visibility (SPEC-027, SPEC-030) ----------
// The Page setup and Format → Columns dialogs, the "Show/hide ruler" toggle
// and the ruler's caret-section geometry live in usePageSetupDialogs.
// caretSectionIndex is declared later (with the caret-page tracking), so it
// comes in as a lazy getter — read only when a dialog opens or a ruler
// computed evaluates, after setup.
const {
  pageSetupOpen,
  pageSetupInitial,
  columnsDialogOpen,
  columnsDialogInitial,
  openPageSetup,
  handlePageSetupApply,
  openColumnsDialog,
  handleColumnsApply,
  toggleRuler,
  rulerPageWidthCm,
  rulerMarginLeftCm,
  rulerMarginRightCm,
  rulerColumns,
} = usePageSetupDialogs({
  ...editorContext,
  paragraphSelection,
  getCaretSectionIndex: () => caretSectionIndex.value,
});

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
  if (event.key === "Escape" && drawingGroups.escapeGroupSelection()) {
    // Step back out one level of group selection first (SPEC-057 AC-24).
    event.preventDefault();
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
    clearChartSelection();
    drawingGroups.clearGroupSelection();
    clearTocSelection();
    clearTofSelection();
    exitHeaderFooterEditing();
    exitNoteEditing();
    return;
  }
  if (handleShapeKeys(event)) {
    return;
  }
  if (drawingGroups.handleGroupKeys(event)) {
    return;
  }
  if (handleChartKeys(event, deleteChartWithCaptionPrompt)) {
    return;
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
    deleteImageWithCaptionPrompt();
    return;
  }
  if (tofSelection.value && (event.key === "Delete" || event.key === "Backspace")) {
    // Deleting a selected table of figures (SPEC-055 AC-35), like a selected TOC. tofSelection is
    // cleared on click-away (useCanvasGestures), so this fires only while it's genuinely selected.
    event.preventDefault();
    handleTofDelete();
    return;
  }
  if (tofSelection.value && isContentKey(event)) {
    // Typing while a table of figures is selected (SPEC-055): not editable.
    event.preventDefault();
    addToast('This table of figures isn’t editable — use "Update table of figures" instead.');
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
  // Only the event-dependent commands remain switch cases here — they read
  // shiftKey or must control preventDefault themselves, so they cannot be
  // nullary entries in the editorCommands table the default arm dispatches
  // through.
  switch (shortcut.id) {
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
    case "edit.pastePlain":
      // Ctrl/Cmd+Shift+V (SPEC-016 AC-27): arm the plain-text flag and let
      // the browser fire the paste event it produces for this combo — that
      // event hands us every clipboard flavour at once.
      armPlainPaste();
      break;
    case "paragraph.indent":
    case "paragraph.outdent": {
      // Always prevented: the browser's default Tab would move focus out of
      // the editor. The precedence table is resolved below (SPEC-054).
      event.preventDefault();
      const selection = currentSelection(root);
      if (!selection) {
        break;
      }
      // SPEC-022 AC-10..12 (precedence rule 1): in a table cell, Tab moves to
      // the next cell (Shift+Tab the previous), the last cell appending a row.
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
      // SPEC-054 Tab precedence (rules 2-4): a caret collapsed at offset 0
      // gates the list demote (rule 2, narrowed from SPEC-018's "any offset")
      // and indent (rule 3); elsewhere a literal tab (rule 4) — AC-29, AC-30.
      const collapsedAtStart = isCollapsed(selection) && selection.focus.offset === 0;
      const listAtStart =
        collapsedAtStart &&
        touchedBlocks(model, selection).some((address) => !!paragraphAt(model, address)?.list);
      if (listAtStart) {
        applyParagraphCommand(
          (current, target) => changeListLevel(current, target, event.shiftKey ? -1 : 1),
          "list level",
        );
      } else if (collapsedAtStart) {
        applyParagraphCommand(
          (current, target) => adjustIndent(current, target, (event.shiftKey ? -1 : 1) * INDENT_STEP_CM),
          "indent",
        );
      } else if (!event.shiftKey) {
        applyEdit(insertText(model, selection, "\t", pendingFormat), selection, "typing", { groupable: true });
      }
      break;
    }
    case "nav.pageUp":
    case "nav.pageDown":
      event.preventDefault();
      void moveCaretByPage(shortcut.id === "nav.pageDown" ? 1 : -1, event.shiftKey);
      break;
    default: {
      // Every other registry id is a nullary command in editorCommands —
      // the same table the menu bar dispatches through — except the
      // document-level ids handleGlobalShortcut owns (find/print keep
      // working over a read-only too-large document), which must not
      // double-dispatch from this canvas path on top of that handler.
      if (shortcut.id === "find.openReplace" || shortcut.id === "file.print") {
        break;
      }
      // Every moved keyboard command called event.preventDefault() before
      // its body, so a table hit always suppresses the browser default
      // here; edit.pastePlain stays a case above precisely because it must
      // NOT preventDefault — the browser's own paste event does the work.
      const command = editorCommands.get(shortcut.id);
      if (command) {
        event.preventDefault();
        command();
      }
    }
  }
}

// Margin/empty-space clicks (handlePageMouseDown) are handled by
// useCanvasGestures, wired below.

// ---------------------------------------------------------------------------

// --- document outline navigator (SPEC-037) ---------------------------------
// The pane's state, the outline tree and its handlers live in useOutlinePane.
// Wired BEFORE usePaginationEngine below, which takes the destructured
// hiddenBodyBlocks as an extra. caretPosition is the backbone caret mirror
// declared with the caret tracking above; jumpToAnchor is destructured from
// useLinks, wired after pagination — the lambda defers the read to call time
// (an event handler, past setup), never during this deps object's evaluation.
const {
  outlinePaneOpen,
  outlinePaneCollapsed,
  collapsedHeadings,
  outlineTree,
  activeOutlineAnchor,
  hiddenBodyBlocks,
  toggleOutlinePane,
  handleOutlineNavigate,
  handleOutlineToggleCollapse,
  handleOutlineMove,
  handleBodyHeadingToggle,
} = useOutlinePane({
  ...editorContext,
  getCaretPosition: () => caretPosition.value,
  jumpToAnchor: (anchor) => jumpToAnchor(anchor),
});

// --- pagination (SPEC-028, SPEC-031, SPEC-047) -------------------------------
// The controller's session lifecycle — construction over this document, the
// relayout triggers (model replacement, the outline's hidden set, a renumbered
// note label), the first-layout deferral and the open progress flag — lives in
// usePaginationEngine, together with the note numbering it is mutually
// recursive with (see that composable's header). Wired after useOutlinePane,
// whose hiddenBodyBlocks it lays out around, and before every consumer of
// `pagination` below.
const {
  pagination,
  paginationPages,
  paginationBusy,
  paginationState,
  documentOpening,
  noteLabelsMap,
  noteTooltipsMap,
  pageCount,
  pageNumbers,
} = usePaginationEngine({ ...editorContext, hiddenBodyBlocks });

const marksOn = computed(() => doc.value?.showFormattingMarks === true);
// "View → Field shading" (SPEC-053 AC-11): a session-local view toggle — like
// the ruler/marks toggles it decorates the screen only and is never written to
// the file. Off by default (Word's "When selected" is the closest match, but
// the toggle explicitly lets a user shade every field at once).
const fieldShadingOn = ref(false);

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
} = usePrintPreview({ ...editorContext, pagination });

// PDF export (SPEC-034): reads the realized canvas into a client-built PDF.
const { pdfExporting, pdfProgress, exportPdf } = usePdfExport({
  ...editorContext,
  pagination,
  realizeAllPages,
  endRealizeAll: () => {
    if (!printPreviewOpen.value) {
      printForceRealize.value = false;
    }
  },
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
  ...editorContext,
  paragraphSelection,
  restoreCaretAfterRender,
  activeHeaderFooter,
  pageOfCaret: (caretPos) => pagination.pageOfCaret(caretPos),
  pageNumbers,
  clearImageSelection,
  getSelectionAnchor: () => selectionAnchor,
});

// --- canvas gesture hub (SPEC-011, SPEC-020, SPEC-022, SPEC-024, SPEC-025,
// SPEC-031, SPEC-033, SPEC-051, SPEC-052) ------------------------------------
// The page/canvas mouse handlers — the cell-range drag selection, shape
// draw/move and dblclick recovery, chart recovery, image drag-and-drop,
// note-reference and link clicks, the table hover handles and margin caret
// placement — live in useCanvasGestures. Wired after useLinks, the last
// provider of its extras (every dep here is passed eagerly), and before
// useContextMenus, whose wiring passes tableCellAtTarget eagerly.
const {
  canMergeCells,
  clearCellSelection,
  tableCellAtTarget,
  isMouseButtonDown,
  handleContentClick,
  handleCanvasMouseDown,
  handleCanvasDblClick,
  handleCanvasMouseMove,
  handleCanvasMouseUp,
  handleCanvasDragEnter,
  handleCanvasDragOver,
  handleCanvasDrop,
  handlePageMouseOver,
  handlePageMouseDown,
} = useCanvasGestures({
  ...editorContext,
  cellSelection,
  tableHover,
  drawGeometry,
  startShapeDraw,
  handleShapeSelect,
  handleShapeEditText,
  shapeAtClientPoint,
  clearShapeSelection,
  startShapeMove,
  handleChartSelect,
  handleChartEditData,
  chartAtClientPoint,
  clearChartSelection,
  exitNoteEditing,
  enterNoteEditing,
  clearImageSelection,
  insertImageFromFile,
  clearTocSelection,
  clearTofSelection: () => clearTofSelection(),
  jumpToAnchor,
  // Deferred: useCaptions/useFields initialise below, so wrap the read in an arrow that resolves at click time.
  followCrossReference: (target) => followCrossReference(target),
  selectField: (target) => selectField(target),
  openLinkTarget,
  maybeApplyPainter,
});

const caretSectionIndex = ref(0);
const paragraphPagination = ref(emptyParagraphPaginationState());

// --- page virtualization (SPEC-028 AC-49/50) --------------------------------
// Which pages render real content versus correctly-sized empty page boxes
// lives in useViewportVirtualization (see its header comment). Wired after
// usePrintPreview — whose forceRealizeAll (printForceRealize) pauses it —
// and before usePaginationView, which consumes the trackers below through
// its own extras.
const {
  viewportPageRange,
  anchorPage,
  caretPage,
  realizedSticky,
  virtualizing,
  refreshViewportRange,
  scheduleViewportRefresh,
  refreshCaretPage,
} = useViewportVirtualization({
  ...editorContext,
  pagination,
  paginationPages,
  printForceRealize,
  getPageCount: () => pageCount.value,
});

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
  ...editorContext,
  canSearch: findAndPrintEnabled,
  paginationPages,
  pageOfCaret: (caretPos) => pagination.pageOfCaret(caretPos),
  updateCaretStatus,
  setCaret: (next) => {
    caret = next;
  },
  registerRefresh: registerFindRefresh,
});

const {
  spellUnderlineRects,
  customDictionaryList,
  ignoreOccurrence,
  ignoreAllOccurrences,
  addWordToDictionary,
  removeWordFromDictionary,
  suggestionsFor,
} = useSpellCheck({ ...editorContext });

// --- context menus (SPEC-016, SPEC-022, SPEC-032, SPEC-035) ------------------
// The editor's right-click menu, the spell-check suggestions menu and the
// Edit-menu Select all / Paste without formatting live in useContextMenus.
// The menu-target variables (contextMenuTableCell / contextMenuListTarget)
// stay plain lets above — the template reads them fresh on each menu render
// and useTables' wiring closes over the cell — so the composable reaches
// them through the get/set lambdas here.
const {
  contextMenu,
  spellContextMenu,
  handleContextMenu,
  closeContextMenu,
  closeSpellContextMenu,
  applySpellSuggestion,
  spellIgnore,
  spellIgnoreAll,
  spellAddToDictionary,
  handleContextMenuListAction,
  contextMenuCut,
  contextMenuCopy,
  contextMenuPaste,
  menuSelectAll,
  menuPastePlain,
} = useContextMenus({
  ...editorContext,
  tableCellAtTarget,
  setContextMenuTableCell: (cell) => {
    contextMenuTableCell = cell;
  },
  getContextMenuListTarget: () => contextMenuListTarget,
  setContextMenuListTarget: (target) => {
    contextMenuListTarget = target;
  },
  spellUnderlineRects,
  suggestionsFor,
  ignoreOccurrence,
  ignoreAllOccurrences,
  addWordToDictionary,
  getLastCopy,
  performPaste,
  setCaret: (next) => {
    caret = next;
  },
  updateCaretStatus,
  setPendingFormat: (next) => {
    pendingFormat = next;
  },
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
} = useComments({ ...editorContext });

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
} = useTrackChanges({ ...editorContext });

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
} = useVersionHistory({ ...editorContext });

// File-level document actions (title rename, download, save as template,
// delete, conversion report) live in useFileActions; the download path needs
// the PDF exporter, the autosave commit and the version-history commit point
// wired in, so this sits after all three.
const {
  formatBadge,
  downloadFormat,
  conversionReport,
  conversionReportSameFormat,
  confirmingDelete,
  savingAsTemplate,
  handleTitleUpdate,
  handleDownload,
  closeConversionReport,
  saveCurrentAsTemplate,
  deleteOpenDocument,
} = useFileActions({
  ...editorContext,
  exportPdf,
  commitNow: () => autosave.commitNow(),
  noteVersionHistoryCommitPoint,
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
  ...editorContext,
  imageSelection,
  openImageOptions,
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

// --- font fallback & embedding (SPEC-048): bundled+embedded @font-face, the
// Fonts panel, embed-on-save; relayout re-measures when a face arrives.
const { fontsPaneOpen, fontsReady, fontReport, documentFonts, embedControls, openFontsPane, closeFontsPane } = useFonts({
  ...editorContext,
  relayout: () => pagination.invalidateAll(),
});

// --- document properties and metadata (SPEC-041) ---------------------------
const {
  propertiesOpen,
  propertiesFileSize,
  openProperties,
  closeProperties,
  updateDocumentProperties,
  updateAuthorIdentity,
} = useDocumentProperties({
  ...editorContext,
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
} = useDocumentLanguage({ ...editorContext });

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
  ...editorContext,
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
  cellSelection,
  imageSelection,
  shapeSelection,
  shapeMultiSelection,
  editingShapeText,
  chartSelection,
  groupSelection,
  collapsedHeadings,
  remount,
  realizedSticky,
  hfRemountEpoch,
  noteRemountEpoch,
  tocPageNumberFor,
  trackChangesDisplayMode: effectiveTrackChangesDisplayMode,
  fieldShadingOn,
  caretPosition,
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

// --- mail-merge fields (SPEC-046) -------------------------------------------
// The attached data source, the field picker / merge-run dialogs' state and
// the literal {{...}} conversion live in useMergeFields.
const {
  mergeDataSource,
  mergeFieldPickerOpen,
  mergeRunDialogOpen,
  attachMergeDataSource,
  mergeFieldPickerFields,
  repeatRowFieldOptions,
  insertMergeFieldFromPicker,
  canConvertLiteralChips,
  convertLiteralChips,
} = useMergeFields({ ...editorContext, paragraphSelection });

// SPEC-053 fields: the Insert dialogs, F9/Update-all update paths, and the
// field context-menu/chip actions.
const {
  fieldDialogOpen,
  dateTimeDialogOpen,
  crossReferenceDialogOpen,
  nowMs: fieldNowMs,
  fieldPreviews,
  propertyNames: fieldPropertyNames,
  crossReferenceHeadings,
  crossReferenceBookmarks,
  crossReferenceCaptions,
  openFieldDialog,
  handleFieldInsert,
  openDateTimeDialog,
  handleDateTimeInsert,
  openCrossReferenceDialog,
  handleCrossReferenceInsert,
  handleUpdateField,
  updateFieldInSurface,
  handleUpdateAllFields,
  followCrossReference,
  selectField,
  contextMenuOnField,
  contextMenuFieldLocked,
  onContextMenu,
  handleFieldAction,
} = useFields({ ...editorContext, tocPageNumberFor, handleContextMenu, closeContextMenu });

const {
  captionDialogOpen,
  captionLabels,
  captionTargetSelected,
  openCaptionDialog,
  handleCaptionInsert,
  handleCaptionCancel,
  captionSettings,
  toggleAutoImages,
  toggleAutoTables,
  requestObjectDelete,
  deleteCaptionPromptOpen,
  deleteCaptionKind,
  confirmDeleteBoth,
  confirmKeepCaption, cancelDeletePrompt,
  deleteImageWithCaptionPrompt,
  deleteChartWithCaptionPrompt,
  insertImageFileWithAutoCaption, insertTableWithAutoCaption,
  tofDialogOpen,
  tofLabels,
  tofSelection,
  openTableOfFiguresDialog,
  handleTofInsert,
  handleTofCancel,
  handleTofAction,
  handleTofDelete,
  clearTofSelection,
} = useCaptions({
  ...editorContext,
  activeHeaderFooter, currentHfSelection, applyHeaderFooterEdit, hfRootEl, imageSelection,
  chartSelection,
  cellSelection,
  caretPosition,
  tofPageNumberFor: tocPageNumberFor,
  jumpToAnchor,
  handleImageDelete,
  handleChartDelete,
  insertImageFromFile, handleInsertTable,
});

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
// The menu bar itself (File/Edit/View/Insert/Format/Tools and every item)
// lives in DocumentMenuBar.vue, which renders menu state from props and
// reports picks as events. Nullary commands arrive here as string ids
// through @command and dispatch through the editorCommands table below; picks
// that carry a payload (table size, break kind, note kind, gallery picks,
// the image file, the painter's click pair) have dedicated events bound in
// the template.

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

/** "Insert -> Shape" gallery pick (SPEC-051 AC-3): arms draw mode so the
 * next drag/click on the canvas draws the chosen geometry. */
function handleShapePick(geometry = "rect") {
  armShapeDraw(asShapeGeometry(geometry));
}

/** "Insert -> Chart" gallery pick (SPEC-052 AC-2, AC-3): inserts at the
 * caret immediately, sample data and all — a chart has no draw-mode gesture
 * (unlike a shape's drag-to-size). */
function handleChartPick(chartType = "columnClustered") {
  insertChartAction(asChartType(chartType));
}

/** Whether a table is selected for "Insert -> Chart from table" (AC-21) —
 * a table block is the target of the active cell-range selection. */
const tableSelectedForChart = computed(() => {
  const model = doc.value;
  const range = cellSelection.value;
  if (!model || !range) {
    return false;
  }
  const block = model.sections[range.section]?.blocks[range.block];
  return !!block && block.type === "table";
});

/** "Insert -> Chart from table" (AC-21, AC-22): reads the selected table's
 * cells (every row/column, not just the range selection — a whole-table
 * source is the simplest, most predictable default), reports any cell that
 * could not be read as a number. */
function handleChartFromTable(chartType = "columnClustered") {
  const model = doc.value;
  const range = cellSelection.value;
  if (!model || !range) {
    return;
  }
  const table = model.sections[range.section]?.blocks[range.block];
  if (!table || table.type !== "table") {
    return;
  }
  const cellText = table.rows.map((row) => row.cells.map((cell) => cell.blocks.map((paragraph) => blockText(paragraph)).join(" ")));
  const result = insertChartFromTableAction({ section: range.section, block: range.block }, asChartType(chartType), cellText);
  if (result.unparsable.length > 0) {
    addToast(`Chart from table: ${result.unparsable.length} cell(s) could not be read as a number and were left blank.`);
  }
}

// Every nullary editor command, keyed by id — the ONE registry both the
// menubar (handleMenuCommand) and the keyboard shortcut path (handleKeyDown's
// default arm) dispatch through, so a new command is wired in exactly one
// place. Ids reuse the lib/shortcuts.ts names wherever a shortcut exists for
// the command; menu-only commands follow the same dotted style. Listed in
// the bar's menu-by-menu order, then the shortcut-only commands. A Map
// rather than an object literal: a string-keyed lookup into an inferred
// object-literal type needs an index signature, which is TS-only syntax
// espree cannot parse here — Map's get() is typed without one.
const editorCommands = new Map([
  // File
  ["file.print", openPrintPreview],
  ["file.page-setup", openPageSetup],
  ["file.properties", openProperties],
  ["file.version-history", toggleVersionHistoryPane],
  ["file.save-as-template", () => { savingAsTemplate.value = true; }],
  ["file.mail-merge", () => { mergeRunDialogOpen.value = true; }],
  ["file.delete", () => { confirmingDelete.value = true; }],
  // Edit
  // Ctrl/Cmd+Z (SPEC-015 AC-1).
  ["edit.undo", performUndo],
  // Ctrl/Cmd+Shift+Z, or Ctrl+Y — the Windows convention (SPEC-015 AC-2, AC-3).
  ["edit.redo", performRedo],
  ["edit.cut", contextMenuCut],
  ["edit.copy", contextMenuCopy],
  ["edit.paste", contextMenuPaste],
  ["edit.pastePlain", menuPastePlain],
  // Ctrl/Cmd+A: select the whole document, exactly to the model's extent —
  // shared by the Edit-menu item and the shortcut (SPEC-049).
  ["edit.selectAll", menuSelectAll],
  ["find.openReplace", () => { openFind(true); }],
  ["edit.autocorrect-smart-quotes", toggleAutocorrectSmartQuotes],
  ["edit.autocorrect-dashes", toggleAutocorrectDashes],
  // View
  ["view.ruler", toggleRuler],
  ["view.outline-pane", toggleOutlinePane],
  ["view.styles-pane", toggleStylesPane],
  ["view.comments-pane", toggleCommentsPane],
  ["view.comments-only", toggleCommentsOnlyView],
  // Ctrl/Cmd+Shift+E (SPEC-039 AC-1).
  ["edit.toggleTrackChanges", toggleTrackChangesRecording],
  ["view.track-changes-pane", toggleTrackChangesPane],
  // Insert
  ["insert.table-dialog", () => { insertTableDialogOpen.value = true; }],
  ["insert.textbox", insertTextBoxAction],
  ["insert.chart-from-table", () => { handleChartFromTable(); }],
  ["insert.header", () => { openInsertHeaderFooter("header"); }],
  ["insert.footer", () => { openInsertHeaderFooter("footer"); }],
  ["insert.page-number", openPageNumberDialog],
  ["insert.header-footer-options", openHeaderFooterOptions],
  ["insert.merge-field", () => { mergeFieldPickerOpen.value = true; }],
  ["insert.note-settings", openNoteSettings],
  // SPEC-053 fields.
  ["insert.field", openFieldDialog],
  ["insert.date-time", openDateTimeDialog],
  ["insert.cross-reference", openCrossReferenceDialog],
  ["insert.toc", openTocDialog],
  ["insert.caption", openCaptionDialog],
  ["insert.table-of-figures", openTableOfFiguresDialog],
  ["format.auto-caption-images", toggleAutoImages],
  ["format.auto-caption-tables", toggleAutoTables],
  ["insert.special-character", openSpecialCharacters],
  // Ctrl/Cmd+K opens the link dialog (SPEC-025 AC-1, AC-2).
  ["insert.link", openLinkDialog],
  ["insert.bookmark", openBookmarkDialog],
  ["insert.bookmark-list", () => { bookmarkManagerOpen.value = true; }],
  // Ctrl/Cmd+Alt+M (SPEC-038 AC-1).
  ["insert.comment", openCommentComposer],
  // Format
  ["format.styles-pane", toggleStylesPane],
  ["format.import-styles", openImportDialog],
  // Ctrl/Cmd+\ (SPEC-020 AC-3).
  ["format.clearFormatting", handleClearFormatting],
  ["format.columns", openColumnsDialog],
  ["format.multilevel-link-headings", handleLinkHeadingNumbering],
  ["format.multilevel-define-list", openDefineList],
  ["format.text-color", openTextColorMenu],
  ["format.highlight", openHighlightMenu],
  // Tools
  ["tools.word-count", openStatistics],
  ["tools.dictionary", openCustomDictionaryDialog],
  ["tools.accessibility", openAccessibilityPane],
  ["tools.fonts", openFontsPane],
  ["tools.convert-merge-fields", convertLiteralChips],
  // SPEC-053 field update paths and View toggle.
  ["tools.update-all-fields", handleUpdateAllFields],
  ["view.field-shading", () => { fieldShadingOn.value = !fieldShadingOn.value; }],
  // F9 (AC-15): update the field at the caret. Unlink/lock/unlock reach the
  // model through the field context menu / toolbar chip (handleFieldAction),
  // not through their own command ids.
  ["field.update", handleUpdateField],
  // Shortcut-only commands: no menubar item emits these ids (the menu
  // reaches page break and foot/endnotes through the insert-break /
  // insert-note payload events instead), so they arrive only from
  // handleKeyDown's registry lookup.
  ["insert.pageBreak", () => {
    // Ctrl/Cmd+Enter (AC-4): a manual page break at the caret.
    const root = contentRoot.value;
    const model = doc.value;
    if (!root || !model) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(caret ?? documentStart());
    applyEdit(insertBreak(model, selection, "page"), selection, "insert page break");
    // The caret moved to the top of a page that is very likely below the
    // fold: follow it, or the user types onto a page they cannot see.
    void scrollModelCaretIntoView(caret, true);
  }],
  // Alt+= (SPEC-056 AC-3), and Insert → Equation (AC-1).
  ["insert.equation", () => { equations.insertEquation(); }],
  // Insert → Display equation (SPEC-056 AC-2).
  ["insert.display-equation", () => { equations.insertDisplayEquationAtCaret(); }],
  // Insert → Numbered equation (SPEC-056 AC-17).
  ["insert.numbered-equation", () => { equations.insertNumberedDisplayEquationAtCaret(); }],
  // Ctrl/Cmd+Alt+F (SPEC-031 AC-1).
  ["insert.footnote", () => { handleInsertNote("footnote"); }],
  // Ctrl/Cmd+Alt+D (SPEC-031 AC-3).
  ["insert.endnote", () => { handleInsertNote("endnote"); }],
  // Ctrl/Cmd+Shift+Space (SPEC-026 AC-11): a direct non-breaking space
  // insert, not typed input, so autocorrect never sees it.
  ["insert.nonBreakingSpace", () => { insertSpecialCharacter("\u00A0"); }],
  // SPEC-054: a literal tab anywhere (AC-31/32); and the Tabs dialog.
  ["insert.tab-character", () => { insertSpecialCharacter("\t"); }],
  ["format.tabs", openTabsDialog],
  // Ctrl/Cmd+Shift+P (AC-12).
  ["view.formattingMarks", toggleFormattingMarks],
  // Ctrl/Cmd+B/I/U (AC-3, AC-4, AC-5).
  ["format.bold", () => { toggleCharacterFormat("bold"); }],
  ["format.italic", () => { toggleCharacterFormat("italic"); }],
  ["format.underline", () => { toggleCharacterFormat("underline"); }],
  // Ctrl/Cmd+Shift+X: strikethrough (AC-6).
  ["format.strikethrough", () => { toggleCharacterFormat("strikethrough"); }],
  // Ctrl/Cmd+, and Ctrl/Cmd+. : subscript / superscript (SPEC-013).
  ["format.subscript", () => { toggleCharacterFormat("subscript"); }],
  ["format.superscript", () => { toggleCharacterFormat("superscript"); }],
  // Ctrl/Cmd+L/E/R/J: alignment (SPEC-014 AC-3).
  ["paragraph.alignLeft", () => { handleAlign("left"); }],
  ["paragraph.alignCenter", () => { handleAlign("center"); }],
  ["paragraph.alignRight", () => { handleAlign("right"); }],
  ["paragraph.alignJustify", () => { handleAlign("justify"); }],
  // Ctrl/Cmd+Alt+1..6 apply Heading 1-6, Ctrl/Cmd+Alt+0 Normal
  // (SPEC-019 AC-5, AC-6).
  ["style.normal", () => { handleApplyStyle(NORMAL_STYLE_ID); }],
  ["style.heading1", () => { handleApplyStyle(headingStyleId(1)); }],
  ["style.heading2", () => { handleApplyStyle(headingStyleId(2)); }],
  ["style.heading3", () => { handleApplyStyle(headingStyleId(3)); }],
  ["style.heading4", () => { handleApplyStyle(headingStyleId(4)); }],
  ["style.heading5", () => { handleApplyStyle(headingStyleId(5)); }],
  ["style.heading6", () => { handleApplyStyle(headingStyleId(6)); }],
]);

function handleMenuCommand(id = "") {
  editorCommands.get(id)?.();
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
      <DocumentMenuBar
        :editable="editable"
        :undo-redo="undoRedoState"
        :autocorrect-smart-quotes="autocorrectSettings.smartQuotes"
        :autocorrect-dashes="autocorrectSettings.dashes"
        :show-ruler="doc.showRuler !== false"
        :outline-pane-open="outlinePaneOpen"
        :styles-pane-open="stylesPaneOpen"
        :comments-pane-open="commentsPaneOpen"
        :comments-only-view="commentsOnlyView"
        :track-changes-recording="trackChangesRecording"
        :track-changes-pane-open="trackChangesPaneOpen"
        :can-convert-literal-chips="canConvertLiteralChips"
        :table-selected-for-chart="tableSelectedForChart"
        :caption-target-selected="captionTargetSelected"
        :auto-caption-images="captionSettings.autoImages"
        :auto-caption-tables="captionSettings.autoTables"
        :painter-mode="painterMode"
        :field-shading="fieldShadingOn"
        @command="handleMenuCommand"
        @insert-table-size="(rows, cols) => insertTableWithAutoCaption({ rows, cols })"
        @insert-break="handleInsertBreak"
        @insert-section-break="handleInsertSectionBreak"
        @insert-note="handleInsertNote"
        @shape-pick="handleShapePick"
        @chart-pick="handleChartPick"
        @multilevel-preset="handleMultilevelPreset"
        @insert-image-file="insertImageFileWithAutoCaption"
        @painter-click="handlePainterClick"
        @painter-dblclick="handlePainterDblClick"
      />
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
      <FieldToolbar
        v-if="editable && selectedField"
        :locked="isFieldLocked(selectedField)"
        :read-only="isReadOnlyField(selectedField)"
        @field-action="handleFieldAction"
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
      <div
        v-if="fontsPaneOpen"
        class="fonts-dock print-exclude"
      >
        <FontsPane
          :entries="fontReport"
          :controls="embedControls"
          @close="closeFontsPane"
        />
      </div>
      <div class="page-wrap">
        <FontLoadingSkeleton v-if="!fontsReady" />
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
          <TableHandlesOverlay
            v-if="tableHover"
            :geometry="tableHover"
            @handle="handleTableHandle"
            @column-resize="startColumnResize"
            @row-resize="startRowResize"
          />
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
          <ShapeDrawPreview
            v-if="drawPreview"
            class="print-exclude"
            :shape="drawPreview.shape"
            :left="drawPreview.left"
            :top="drawPreview.top"
            :width="drawPreview.width"
            :height="drawPreview.height"
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
            @contextmenu="onContextMenu"
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
              @chart-select="handleChartSelect"
              @chart-edit-data="handleChartEditData"
              @group-select="drawingGroups.handleGroupSelect"
              @group-edit-text="drawingGroups.handleGroupEditText"
              @toc-action="handleTocAction"
              @tof-action="handleTofAction"
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
      :same-format="conversionReportSameFormat"
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
      :default-interval-cm="tabsDialogDefaultInterval"
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
      @delete="deleteImageWithCaptionPrompt"
    />
    <ShapeToolbar
      v-if="shapeSelection && shapeHandles && doc && shapeAt(doc, shapeSelection) && !movingShape"
      class="print-exclude"
      :left="shapeToolbarStyle().left"
      :top="shapeToolbarStyle().top"
      :placement="shapeToolbarStyle().placement"
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
    <GroupToolbar
      v-if="drawingGroups.groupToolbarState()"
      v-bind="drawingGroups.groupToolbarState()"
      class="print-exclude"
      @wrap="drawingGroups.setSelectedWrap"
      @rotation="drawingGroups.setSelectedRotation"
      @alt-text="drawingGroups.setSelectedAltText"
      @duplicate="drawingGroups.duplicateSelected"
      @delete="drawingGroups.deleteSelected"
      @ungroup="drawingGroups.ungroupSelectedContainer"
    />
    <ChartToolbar
      v-if="chartSelection && chartHandles && doc && chartAt(doc, chartSelection)"
      class="print-exclude"
      :left="chartToolbarStyle().left"
      :top="chartToolbarStyle().top"
      :unsupported="chartAt(doc, chartSelection)?.unsupportedChartXml !== undefined"
      @edit-data="handleChartEditData(chartSelection.section, chartSelection.block)"
      @options="openChartOptions"
      @delete="deleteChartWithCaptionPrompt"
      @duplicate="handleChartDuplicate"
    />
    <EditorContextMenu
      v-if="contextMenu.visible"
      class="print-exclude"
      :left="contextMenu.left"
      :top="contextMenu.top"
      :table-actions="!!contextMenuTableCell"
      :can-merge-cells="canMergeCells"
      :list-actions="!!contextMenuListTarget"
      :field-actions="contextMenuOnField"
      :field-locked="contextMenuFieldLocked"
      :caption-action="captionTargetSelected"
      @cut="contextMenuCut"
      @copy="contextMenuCopy"
      @paste="contextMenuPaste"
      @table-action="handleContextMenuTableAction"
      @list-action="handleContextMenuListAction"
      @field-action="handleFieldAction"
      @caption-action="() => { closeContextMenu(); openCaptionDialog(); }"
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
    <FieldDialog
      v-if="fieldDialogOpen && doc"
      :model="doc"
      :now-ms="fieldNowMs"
      :file-name="doc.title"
      :property-names="fieldPropertyNames"
      @insert="handleFieldInsert"
      @cancel="() => { fieldDialogOpen = false; }"
    />
    <DateTimeDialog
      v-if="dateTimeDialogOpen"
      :previews="fieldPreviews"
      @insert="handleDateTimeInsert"
      @cancel="() => { dateTimeDialogOpen = false; }"
    />
    <CrossReferenceDialog
      v-if="crossReferenceDialogOpen"
      :headings="crossReferenceHeadings"
      :bookmarks="crossReferenceBookmarks"
      :captions="crossReferenceCaptions"
      @insert="handleCrossReferenceInsert"
      @cancel="() => { crossReferenceDialogOpen = false; }"
    />
    <TocOptionsDialog
      v-if="tocOptionsDialogOpen"
      :initial="defaultTocOptions()"
      @apply="handleTocInsert"
      @cancel="handleTocCancel"
    />
    <CaptionDialog
      v-if="captionDialogOpen"
      :labels="captionLabels"
      @insert="handleCaptionInsert"
      @cancel="handleCaptionCancel"
    />
    <TableOfFiguresDialog
      v-if="tofDialogOpen"
      :labels="tofLabels"
      @apply="handleTofInsert"
      @cancel="handleTofCancel"
    />
    <DeleteCaptionPrompt
      v-if="deleteCaptionPromptOpen"
      :object-kind="deleteCaptionKind"
      @delete-both="confirmDeleteBoth"
      @keep-caption="confirmKeepCaption"
      @cancel="cancelDeletePrompt"
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
      @insert="onSpecialCharacterInsert"
      @cancel="specialCharacterDialogOpen = false"
    />
    <EquationEditorHost :api="equations" />
    <ImageOptionsDialog
      v-if="imageOptionsOpen"
      :initial="imageOptionsInitial"
      @apply="handleImageOptionsApplyAndRescan"
      @cancel="imageOptionsOpen = false"
    />
    <ChartDataSheetDialog
      v-if="editingChartData"
      :initial="chartDataSheetInitial"
      @preview="({ categories, series }) => applyChartData(categories, series)"
      @apply="({ categories, series }) => { applyChartData(categories, series); editingChartData = false; }"
      @cancel="editingChartData = false"
    />
    <ChartOptionsDialog
      v-if="editingChartOptions"
      :initial="chartOptionsInitial"
      @apply="(options) => { applyChartOptions(options); editingChartOptions = false; }"
      @series-color="(index, color) => applyChartSeriesColor(index, color)"
      @cancel="editingChartOptions = false"
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

/* The accessibility checker's own dock (SPEC-043), and the Fonts panel's
   (SPEC-048) — same shape as the other docked panes, on the same edge (not
   normally opened together with them). */
.accessibility-dock,
.fonts-dock {
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

.fonts-dock :deep(.fonts-pane) {
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
     same z-index band the table handles (15, TableHandlesOverlay.vue) use
     for their own clickable overlay controls. */
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

/* The selected image's eight resize handles (SPEC-024 AC-8, AC-10..12) —
   same page-relative overlay layer as the table handles
   (TableHandlesOverlay.vue), positioned per handle by imageHandleStyle(). */
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
  /* Above .shape-toolbar (z-index: 30, ShapeToolbar.vue). shapeToolbarStyle
     now clears the rotation handle by design, but the toolbar is still a
     wrapping box positioned against a single edge of the shape: any overlap
     it does land in (a shape near the viewport top, where the toolbar flips
     below) must leave the handles clickable, not merely visible — a covered
     rotation handle never receives the drag that starts a rotation (AC-16,
     AC-38). */
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
