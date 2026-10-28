<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, provide, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute } from "vue-router";
import AppShell from "../components/AppShell.vue";
import DocumentMenuBar from "../components/DocumentMenuBar.vue";
import CommandPalette from "../components/CommandPalette.vue";
import BookmarkDialog from "../components/BookmarkDialog.vue";
import BookmarkManagerDialog from "../components/BookmarkManagerDialog.vue";
import CloseWithoutSavingDialog from "../components/CloseWithoutSavingDialog.vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import ConversionReportDialog from "../components/ConversionReportDialog.vue";
import CustomSpacingDialog from "../components/CustomSpacingDialog.vue";
import DefineListDialog from "../components/DefineListDialog.vue";
import DocImageBlock from "../components/DocImageBlock.vue";
import DocPage from "../components/DocPage.vue";
import DocTableBlock from "../components/DocTableBlock.vue";
import DocumentRuler from "../components/DocumentRuler.vue";
import DocumentVerticalRuler from "../components/DocumentVerticalRuler.vue";
import PageMarkupOverlays from "../components/PageMarkupOverlays.vue";
import RulerUnitMenu from "../components/RulerUnitMenu.vue";
import EditorContextMenu from "../components/EditorContextMenu.vue";
import EditorSaveStatus from "../components/EditorSaveStatus.vue";
import SpellContextMenu from "../components/SpellContextMenu.vue";
import DocumentStatusBar from "../components/DocumentStatusBar.vue";
import CustomDictionaryDialog from "../components/CustomDictionaryDialog.vue";
import { SUPPORTED_LANGUAGES } from "../lib/core/language";
import HeaderFooterOptionsDialog from "../components/HeaderFooterOptionsDialog.vue";
import ImageOptionsDialog from "../components/ImageOptionsDialog.vue";
import ShapePositionDialog from "../components/ShapePositionDialog.vue";
import ImageToolbar from "../components/ImageToolbar.vue";
import ShapeToolbarIsland from "../components/ShapeToolbarIsland.vue";
import DrawingOverlays from "../components/DrawingOverlays.vue";
import ShapeHandlesOverlay from "../components/ShapeHandlesOverlay.vue";
import ChartToolbar from "../components/ChartToolbar.vue";
import ChartDataSheetDialog from "../components/ChartDataSheetDialog.vue";
import ChartOptionsDialog from "../components/ChartOptionsDialog.vue";
import InsertTableDialog from "../components/InsertTableDialog.vue";
import TableToolsDialogs from "../components/TableToolsDialogs.vue";
import ColumnsDialog from "../components/ColumnsDialog.vue";
import BordersShadingDialog from "../components/BordersShadingDialog.vue";
import ThemeGalleryDialog from "../components/ThemeGalleryDialog.vue";
import CommentsPane from "../components/CommentsPane.vue";
import TrackChangesPane from "../components/TrackChangesPane.vue";
import AccessibilityPane from "../components/AccessibilityPane.vue";
import FontsPane from "../components/FontsPane.vue";
import FontLoadingSkeleton from "../components/FontLoadingSkeleton.vue";
import VersionHistoryPane from "../components/VersionHistoryPane.vue";
import RecoveryBar from "../components/RecoveryBar.vue";
import SpecialCharacterDialog from "../components/SpecialCharacterDialog.vue";
import LinkChip from "../components/LinkChip.vue";
import LinkDialog from "../components/LinkDialog.vue";
import TocOptionsDialog from "../components/TocOptionsDialog.vue";
import FindBar from "../components/FindBar.vue";
import DocumentToolbar from "../components/DocumentToolbar.vue";
import ContextualToolbar from "../components/ContextualToolbar.vue";
import TextSelectionToolbar from "../components/TextSelectionToolbar.vue";
import NoteSettingsDialog from "../components/NoteSettingsDialog.vue";
import OutlinePane from "../components/OutlinePane.vue";
import SidePanel from "../components/SidePanel.vue";
import FieldDialog from "../components/FieldDialog.vue";
import DateTimeDialog from "../components/DateTimeDialog.vue";
import CrossReferenceDialog from "../components/CrossReferenceDialog.vue";
import CaptionDialog from "../components/CaptionDialog.vue";
import TableOfFiguresDialog from "../components/TableOfFiguresDialog.vue";
import DeleteCaptionPrompt from "../components/DeleteCaptionPrompt.vue";
import FieldToolbar from "../components/FieldToolbar.vue";
import PageNumberDialog from "../components/PageNumberDialog.vue";
import PageSetupDialog from "../components/PageSetupDialog.vue";
import PasteOptionsChip from "../components/PasteOptionsChip.vue";
import PrintPreviewDialog from "../components/PrintPreviewDialog.vue";
import MergeFieldPickerDialog from "../components/MergeFieldPickerDialog.vue";
import MergeRunDialog from "../components/MergeRunDialog.vue";
import PropertiesDialog from "../components/PropertiesDialog.vue";
import RenameDialog from "../components/RenameDialog.vue";
import SplitCellDialog from "../components/SplitCellDialog.vue";
import StatisticsDialog from "../components/StatisticsDialog.vue";
import TableHandlesOverlay from "../components/TableHandlesOverlay.vue";
import TablePropertiesDialog from "../components/TablePropertiesDialog.vue";
import TabStopsDialog from "../components/TabStopsDialog.vue";
import { useAutosave } from "../composables/useAutosave";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";
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
import { useFindLink } from "./document/useFindLink";
import { useSupportedCeiling } from "./document/useSupportedCeiling";
import { usePdfExport } from "./document/usePdfExport";
import { usePrintPreview } from "./document/usePrintPreview";
import { useFormatPainter } from "./document/useFormatPainter";
import { useHeadersFooters } from "./document/useHeadersFooters";
import { useImages } from "./document/useImages";
import { useShapes } from "./document/useShapes";
import { useDrawingArrange } from "./document/useDrawingArrange";
import { useConnectors } from "./document/useConnectors";
import { useCharts } from "./document/useCharts";
import { useLinks } from "./document/useLinks";
import { useTextSelectionToolbar } from "./document/useTextSelectionToolbar";
import { useNotes } from "./document/useNotes";
import { useFormattingSurfaces } from "./document/useFormattingSurfaces";
import { noTablePropertiesInitial, useTables } from "./document/useTables";
import { useTableTools } from "./document/useTableTools";
import { usePaginationEngine } from "./document/usePaginationEngine";
import { useThemeControls } from "./document/useThemeControls";
import { useCommandPalette } from "./document/useCommandPalette";
import { usePaginationView } from "./document/usePaginationView";
import { useFields } from "./document/useFields";
import { useLongPress, openContextMenuAtPoint } from "../composables/useLongPress";
import { useVisualViewportKeyboard } from "../composables/useVisualViewportKeyboard";
import { usePinchZoom } from "../composables/usePinchZoom";
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
import { useOverlayRemeasure } from "./document/useOverlayRemeasure";
import { usePageSetupDialogs } from "./document/usePageSetupDialogs";
import { useVerticalRuler } from "./document/useVerticalRuler";
import { useOutlinePane } from "./document/useOutlinePane";
import { useSidePanel } from "./document/useSidePanel";
import { useMergeFields } from "./document/useMergeFields";
import { useInsertBreaks } from "./document/useInsertBreaks";
import { useBordersShading } from "./document/useBordersShading";
import { useThemeGallery } from "./document/useThemeGallery";
import { useViewportVirtualization } from "./document/useViewportVirtualization";
import { useCanvasGestures } from "./document/useCanvasGestures";
import { useKeyboardMap } from "./document/useKeyboardMap";
import { useZoomView } from "./document/useZoomView";
import { useViewModes } from "./document/useViewModes";
import { useCaretScroll } from "./document/useCaretScroll";
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
  createDocument,
  createEmptyHeaderFooterContent,
  documentThemeOf,
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
import { CHART_CATALOGUE } from "../lib/core/charts";
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
import {
  BLOCK_SELECTOR,
  canvasModelText,
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
import { caretClientRect, noTableHandleGeometry } from "../lib/caretGeometry";
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
import { findShortcut, isOutsideDocumentInput } from "../lib/shortcuts";
import { applyTrackChanges } from "../lib/trackChangesDiff";
import { remapSelectionThroughTracking } from "../lib/trackChangesRemap";
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

// --- supported ceiling (SPEC-047): read-only past it, checked once per
// document at the door — useSupportedCeiling.
const { pastSupportedCeiling, ceilingWarning } = useSupportedCeiling(doc);

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

// SPEC-063 Print/Web + Focus (AC-32..51); also owns the Web-layout content
// width (AC-40) and re-tiles/re-fits the presenter once the layout settles
// (AC-36 — onLayoutSettled runs past setup, when zoomView exists).
// onCompactWebDefault (SPEC-068 AC-35) is deferred to nextTick so zoomView —
// wired just below — exists when the compact tier defaults this document to Web
// layout at fit-width.
const viewModes = useViewModes({ documentId, doc, onLayoutSettled: () => { zoomView.recomputeGrid(); zoomView.recomputeFit(); }, onCompactWebDefault: () => { void nextTick(() => zoomView.setFitWidth()); } });
const { viewMode, focusMode, rootClass: viewRootClass, exitFocus, webContentWidthCss } = viewModes;

// Zoom (SPEC-063): created before the shared context — its scale fraction and coordinate helpers are part of it (AC-21, AC-22).
// The page grid only tiles in Print layout, so its eligibility is gated on the view mode (AC-36).
const zoomView = useZoomView({ documentId, pageRef, getScrollContainer: () => scrollContainer, onUnmount, isGridEligible: () => !viewModes.isWebLayout.value });
const { zoom, zoomPercent, fitMode: zoomFitMode, scaled: pagesScaled, viewportStyle: pagesViewportStyle } = zoomView;

// Keeping the model caret on screen (SPEC-011/025) — reads live caret geometry
// and nudges the scroll container; part of EditorContext, so created ahead of it.
const { scrollModelCaretIntoView, keepCaretVisible } = useCaretScroll({
  contentRoot,
  getCaret: () => caret,
  getScrollContainer: () => scrollContainer,
});

// The one object satisfying EditorContext (document/context.ts): the
// universal dependencies every feature composable may rely on. Each use*()
// call below spreads this and adds only its feature-specific extras, so
// wiring a new feature never re-plumbs the core. The function fields are
// hoisted declarations (defined further down this file), so referencing
// them here — before the composables that call them — is safe; the getters
// wrap plain variables that must be read live, never snapshotted.
// SPEC-070: late-bound connector re-route hook (assigned once useConnectors
// is created below), so a shape drag re-routes its connectors live.
let remeasureConnectorsFn = () => {};

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
  zoom,
  clientToContentPx: zoomView.clientToContentPx,
  contentPxToClient: zoomView.contentPxToClient,
  onDrawingGeometryChanged: () => remeasureConnectorsFn(),
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
  snapGuides,
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
  adjustmentHandleList,
  startShapeAdjust,
  handleAdjustmentKeydown,
  shapeToolbarStyle,
  remeasureShapeHandles,
  handleShapeDelete,
  handleShapeDuplicate,
  handleShapeNudge,
  shapeToolbarState, shapeToolbarOn, shapeFormatOpen, openShapeFormat,
  shapePositionOpen,
  shapePositionInitial,
  openShapePosition,
  handleShapePositionApply,
  handleGroupSelected,
  handleUngroupSelected,
  currentGroupMembers,
  startShapeResize,
  startShapeMove,
  startShapeRotate,
  armShapeDraw,
  handleShapePick,
  disarmShapeDraw,
  startShapeDraw,
  drawFreeformTool, freeformPreviewD, handleFreeformPick, startFreeformGesture, pointEdit,
  insertTextBoxAction,
} = useShapes({ ...editorContext, getUpdateFieldInSurface: () => updateFieldInSurface });

// Align/distribute for a drawing selection (SPEC-072); commands spread below.
const drawingArrange = useDrawingArrange({ ...editorContext, shapeSelection, shapeMultiSelection });

// --- connectors (SPEC-070) --------------------------------------------------
// Two-click draw, live-measured overlay geometry, handles and formatting. One
// namespace (not destructured) to hold DocumentView's line budget.
const connectors = useConnectors({ ...editorContext });
const remeasureConnectors = connectors.remeasureConnectors;
remeasureConnectorsFn = remeasureConnectors;

// --- charts (SPEC-052) -------------------------------------------------------
// Chart selection/handles/toolbar, insertion at the caret and the data
// sheet/options-panel commit paths live in useCharts, trimmed from
// useShapes' own pattern (no drag/resize/rotate/multi-select-group gesture
// — see that composable's header comment).
const {
  chartSelection,
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
  handleChartPick,
  tableSelectedForChart,
  handleChartFromTable,
  applyChartData,
  applyChartOptions,
  applyChartSeriesColor,
  handleChartDelete,
  handleChartDuplicate,
  chartToolbarStyle,
  chartAtClientPoint,
  chartDataSheetInitial,
  chartOptionsInitial,
} = useCharts({ ...editorContext, getCellSelection: () => cellSelection.value, caretPosition });

// Equations (SPEC-056): Insert → Equation / Display equation and Alt+=. The
// session's editing surface, op dispatch and overlay live in the composable
// and EquationEditorHost.vue; DocumentView keeps only the insert commands, the
// clicked-equation `provide`, the character-picker route, and the host tag.
const equations = useEquations({ ...editorContext });

// SPEC-056: an equation clicked in the document opens the editor. Provided to
// the paint components (DocPageBlock/DocBlockEntries) which own the click.
provide("activateEquation", equations.activateEquation);
// SPEC-063 AC-20: the live zoom for canvas-painted content (shapes, charts)
// deep in the tree — they size their backing store by devicePixelRatio * zoom.
provide("documentZoom", zoom);

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
    // A text session restores a text model (SPEC-103 widened RestoredState).
    if ("kind" in restored.model) return;
    doc.value = restored.model;
    caret = restored.caret;
    scrollTop = restored.scroll;
    void focusAndRestore();
  },
  onCleanClose: () => noteVersionHistoryCommitPoint(),
});
const { displaySaveState: saveState, saveStateText, recovery, recoveryLabel, keepRecovered, discardRecovered } = autosave;

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

// Zoom/resize move every measured overlay just as a scroll does — policy in useOverlayRemeasure.
const { remeasureDrawingHandles } = useOverlayRemeasure({ ...editorContext, scrollHooks, remeasureShapeHandles, remeasureImageHandles, remeasureChartHandles, remeasureConnectors, clearTableHover: () => (tableHover.value = null) });

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
    const tracked = applyTrackChanges(before, result.model, trackChangesAuthorName.value, new Date().toISOString());
    // The command's selection is in untracked coordinates; retained deletions shift offsets (trackChangesRemap.ts).
    result = { model: tracked, selection: remapSelectionThroughTracking(result.model, tracked, result.selection) };
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
  if (!options.skipCaretScroll) {
    void keepCaretVisible();
  }
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
  // A snapshot can move (or remove) a selected drawing; handles re-find it.
  void nextTick().then(remeasureDrawingHandles);
  fireFindRefresh();
}

/** Whether a caret's own block address (ignoring any cell path) is a TOC
 * atom (SPEC-033) — used to catch every edit attempt at its boundary before
 * the ordinary per-inputType dispatch below. */
function caretAtToc(model = noDocument(), at = noCaret()) {
  if (!at || at.cell) {
    return false;
  }
  return isToc(model.sections[at.section]?.blocks[at.block]);
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
const tables = useTables({
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

// --- the Table menu (SPEC-076) ----------------------------------------------
// Header-row repetition, the style gallery, sorting and both conversions.
const tableTools = useTableTools({
  ...editorContext,
  caretPosition,
  paragraphSelection,
  getContextMenuTableCell: () => contextMenuTableCell,
  closeContextMenu: () => closeContextMenu(),
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
// The Page setup and Format → Columns dialogs, the ruler toggles (both
// together, or each alone) and the ruler's caret-section geometry live in
// usePageSetupDialogs.
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
  showVerticalRuler: showVerticalRulerPref,
  rulersCheckState,
  commands: rulerCommands,
  setMeasurementUnit,
  rulerPageWidthCm,
  rulerMarginLeftCm,
  rulerMarginRightCm,
  rulerColumns,
} = usePageSetupDialogs({
  ...editorContext,
  paragraphSelection,
  getCaretSectionIndex: () => caretSectionIndex.value,
});

// The vertical ruler (SPEC-064): its caret-section geometry, table row
// boundaries, active band, print-layout visibility and drag handler.
const { showVerticalRuler, verticalRulerProps, handleVerticalRulerUpdate } = useVerticalRuler({
  ...editorContext,
  paragraphSelection,
  getCaretSectionIndex: () => caretSectionIndex.value,
  activeHeaderFooter,
  viewMode,
  caretPosition,
});

// Right-clicking either ruler opens the measurement-unit picker (SPEC-064
// AC-10a); choosing a unit writes the same preference the Page setup dialog
// does (AC-10b).
const rulerUnitMenu = ref({ open: false, left: 0, top: 0 });
function openRulerUnitMenu(event = new MouseEvent("contextmenu")) {
  rulerUnitMenu.value = { open: true, left: event.clientX, top: event.clientY };
}
function selectRulerUnit(unit = "cm") {
  setMeasurementUnit(unit);
  rulerUnitMenu.value = { ...rulerUnitMenu.value, open: false };
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
  printBackground,
  rangeError: printRangeError,
  canPrint,
  printablePageCount,
  printHiddenPages,
  openPrintPreview,
  closePrintPreview,
  confirmPrint,
  realizeAllPages,
} = usePrintPreview({ ...editorContext, pagination, remeasureConnectors: () => remeasureConnectorsFn() });

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
  redirectRepeatedHeaderClick: tableTools.redirectRepeatedHeaderClick,
  cellSelection,
  tableHover,
  drawGeometry, startShapeDraw, drawFreeformTool, startFreeformGesture, pointEditActive: pointEdit.active, exitPointEdit: pointEdit.exit,
  handleShapeSelect,
  handleShapeEditText,
  shapeAtClientPoint,
  clearShapeSelection,
  handleConnectorMouseDown: connectors.handleConnectorMouseDown,
  handleConnectorMouseMove: connectors.handleConnectorMouseMove,
  clearConnectorSelection: connectors.clearConnectorSelection,
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
// SPEC-114: /doc/:id?find=<text>(&n=<k>) opens the find bar at that match.
useFindLink({ ...editorContext, canSearch: findAndPrintEnabled, findQuery, findMatchCase, findWholeWords, findUseRegex, openFind });

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
  closing,
  storageWarningOpen,
  handleTitleUpdate,
  handleDownload,
  commands: fileExportCommands,
  closeConversionReport,
  saveCurrentAsTemplate,
  deleteOpenDocument,
  closeOpenDocument,
  confirmCloseAnyway,
} = useFileActions({
  ...editorContext,
  exportPdf,
  commitNow: () => autosave.commitNow(),
  flushForClose: () => autosave.flushForClose(),
  noteVersionHistoryCommitPoint,
});

// --- accessibility checker (SPEC-043) ---------------------------------------
const {
  accessibilityPaneOpen,
  accessibilityIssues,
  accessibilityScanning,
  goToAccessibilityIssue,
  applyAccessibilityFix,
  rescanAccessibility,
  cancelAccessibilityScan,
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
const { fontsPaneOpen, fontsReady, fontReport, documentFonts, embedControls } = useFonts({
  ...editorContext,
  relayout: () => pagination.invalidateAll(),
});

// --- consolidated side panel (SPEC-065) --------------------------------------
// The seven auxiliary panes above each keep their own data + open ref; this
// wiring hands those open refs to useSidePanel, which owns the single docked,
// tabbed region (side/width/collapse/openPanes/activePane, persisted under one
// key) and the open/activate/close/toggle semantics. Badge counts are the live
// lengths of the relevant lists (AC-6b); the accessibility pane rescans when
// the panel opens it and cancels the scan when it closes it.
const commentBadgeCount = computed(() => openCommentThreads.value.length + resolvedCommentThreads.value.length);
const unresolvedChangeCount = computed(() => trackedChanges.value.length);
const accessibilityIssueCount = computed(() => accessibilityIssues.value.length);
const {
  side: sidePanelSide,
  widthPx: sidePanelWidth,
  collapsed: sidePanelCollapsed,
  activePane: sidePanelActivePane,
  tabs: sidePanelTabs,
  panelOpen: sidePanelOpen,
  togglePane: toggleSidePane,
  activatePane: activateSidePane,
  closePane: closeSidePane,
  closePanel: closeSidePanel,
  toggleSide: toggleSidePanelSide,
  toggleCollapsed: toggleSidePanelCollapsed,
  setWidth: setSidePanelWidth,
  commitWidth: commitSidePanelWidth,
  nudgeWidth: nudgeSidePanelWidth,
} = useSidePanel({
  ...editorContext,
  paneOpenRefs: {
    outline: outlinePaneOpen,
    styles: stylesPaneOpen,
    comments: commentsPaneOpen,
    trackChanges: trackChangesPaneOpen,
    versionHistory: versionHistoryPaneOpen,
    accessibility: accessibilityPaneOpen,
    fonts: fontsPaneOpen,
  },
  badges: {
    comments: commentBadgeCount,
    trackChanges: unresolvedChangeCount,
    accessibility: accessibilityIssueCount,
  },
  onPaneOpen: (id) => {
    if (id === "accessibility") {
      rescanAccessibility();
    }
  },
  onPaneClose: (id) => {
    if (id === "accessibility") {
      cancelAccessibilityScan();
    }
  },
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
  void nextTick().then(remeasureConnectors); // SPEC-070: re-route on any edit
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

// Width (px) the open right-docked panes claim on the toolbar's trailing edge,
// so the strip collapses its bands like narrowing the window would (AC-15); the
// left-docked (default-open) outline is excluded to keep 1024px fits (AC-36).
const toolbarReservedRight = computed(
  () =>
    (stylesPaneOpen.value ? 288 : 0) +
    (commentsPaneOpen.value ? 288 : 0) +
    (trackChangesPaneOpen.value ? 288 : 0) +
    (versionHistoryPaneOpen.value ? 320 : 0) +
    (accessibilityPaneOpen.value ? 320 : 0) +
    (fontsPaneOpen.value ? 320 : 0),
);

// The anchor point (centre-x/top/bottom of the current selection rect) the
// shared ContextualToolbar positions the contextual field toolbar against.
function fieldToolbarAnchor() {
  const selection = typeof window !== "undefined" ? window.getSelection() : null;
  const rect = selection && selection.rangeCount > 0 ? selection.getRangeAt(0).getBoundingClientRect() : null;
  return rect ? { left: rect.left + rect.width / 2, top: rect.top, bottom: rect.bottom } : { left: 0, top: 0, bottom: 0 };
}

// The contextual text-selection toolbar (SPEC-061 AC-24..30): raised over a
// non-empty body selection, suppressed while read-only, an object toolbar is
// showing, shape text is being edited, or a track-changes preview is on.
const { visible: textSelectionToolbarVisible, anchor: textSelectionAnchor, hide: hideTextSelectionToolbar } =
  useTextSelectionToolbar({
    contentRoot: () => contentRoot.value,
    // A zoom change moves the selection in client space without firing any
    // event the toolbar listens for, so the composable re-measures its anchor
    // when this changes — keeping the fixed-position toolbar (outside the scale
    // transform) registered with the selection under zoom (SPEC-063 AC-57).
    zoom: () => zoom.value,
    suppressed: () =>
      !editable.value || !!activeHeaderFooter.value || !!shapeSelection.value || !!imageSelection.value ||
      !!chartSelection.value || !!selectedField.value || editingShapeText.value || versionHistoryPreviewActive.value,
  });

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
// Insert → Break / section break (SPEC-018) — wiring in useInsertBreaks.
const { handleInsertBreak, handleInsertSectionBreak } = useInsertBreaks({ ...editorContext, paragraphSelection });

// Format → Borders and shading… (SPEC-074) — wiring in useBordersShading.
const {
  bordersDialogOpen,
  bordersDialogInitial,
  openBordersDialog,
  closeBordersDialog,
  applyBordersShading,
} = useBordersShading({ ...editorContext, paragraphSelection });
const themeGallery = useThemeGallery(editorContext); // Format → Theme… gallery (SPEC-077); wiring in useThemeGallery.

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

// SPEC-068 AC-22/23: a touch long-press on the canvas opens the same context
// menu the mouse right-click opens, by dispatching a synthetic contextmenu at
// the point (which the contenteditable's @contextmenu handler runs unchanged);
// a moving finger is a selection drag and never trips it.
useLongPress({ target: () => contentRoot.value, onLongPress: openContextMenuAtPoint });

// SPEC-068 AC-29/30/31: keep the caret above the on-screen keyboard by scrolling
// the document to the visual viewport, and restore the scroll position when the
// keyboard closes.
const { keyboardInset } = useVisualViewportKeyboard({
  getScrollContainer: () => scrollContainer,
  getCaretRect: () => { const root = contentRoot.value; return root && caret ? caretClientRect(root, caret) : null; },
});

// SPEC-068 AC-32/33/34: pinch drives the zoom transform (setZoomPercent clamps
// to the zoom band); a two-finger pan scrolls. The chrome never scales because
// only .pages carries the zoom transform.
usePinchZoom({
  getScrollContainer: () => scrollContainer,
  target: () => scrollContainer,
  getZoomPercent: () => zoomView.zoomPercent.value,
  setZoomPercent: (percent) => zoomView.setZoomPercent(percent),
});

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
  insertImageFromFile, handleInsertTable: tables.handleInsertTable,
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
  // Escape disarms the format painter regardless of focus (AC-17) — deliberately
  // unguarded by isOutsideDocumentInput; handleKeyDown also disarms it in-canvas.
  if (event.key === "Escape") {
    disarmPainter();
    if (printPreviewOpen.value) {
      closePrintPreview();
    }
    exitFocus(); // Escape leaves Focus mode (SPEC-063 AC-46); a no-op when off.
  }
  if (isOutsideDocumentInput(event.target)) {
    return;
  }
  const shortcut = findShortcut(event);
  if (shortcut?.id === "app.commandPalette") {
    event.preventDefault(); // SPEC-062: opens alongside find/print, canvas focus or not.
    openCommandPalette();
  } else if (shortcut?.id === "find.open") {
    event.preventDefault();
    openFind(false);
  } else if (shortcut?.id === "find.openReplace") {
    event.preventDefault();
    openFind(true);
  } else if (shortcut?.id === "file.print") {
    // SPEC-034 AC-1 (superseding SPEC-021's immediate window.print()): Ctrl/Cmd+P
    // opens the in-app print preview; the native dialog opens from its Print button.
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

// Text colour and highlight stay in the toolbar but are also listed in the
// Format menu; those menu entries open the real toolbar pickers through these
// refs (via DocumentToolbar) instead of duplicating the swatch popups (which
// would collide on their fixed data-testids).
const documentToolbarRef = useTemplateRef("documentToolbarRef");

function openTextColorMenu() { documentToolbarRef.value?.openTextColorMenu(); }
function openHighlightMenu() { documentToolbarRef.value?.openHighlightMenu(); }

// Every nullary editor command, keyed by id — the ONE registry both the
// menubar (handleMenuCommand) and the keyboard shortcut path (handleKeyDown's
// default arm) dispatch through, so a new command is wired in exactly one
// place. Ids reuse the lib/shortcuts.ts names wherever a shortcut exists for
// the command; menu-only commands follow the same dotted style. Listed in
// the bar's menu-by-menu order, then the shortcut-only commands. A Map
// rather than an object literal: a string-keyed lookup into an inferred
// object-literal type needs an index signature, which is TS-only syntax
// espree cannot parse here — Map's get() is typed without one.
const { themeMenu, themeCommands } = useThemeControls({ ...editorContext, invalidate: () => pagination.invalidateAll() });
// The Preferences dialog (SPEC-066) is hosted by AppShell and shared via this
// singleton flag, so File → Preferences… opens the same instance Cmd+, does.
const preferencesDialog = usePreferencesDialog();
const editorCommands = new Map([...themeCommands, // Theme command ids (SPEC-059 AC-50); wiring in useThemeControls.
  ...zoomView.commands, ...viewModes.commands, // SPEC-063 zoom + view-mode command ids (AC-11, AC-33, AC-34).
  ...fileExportCommands, // File — Export command ids; wiring in useFileActions.
  ["file.print", openPrintPreview],
  ["file.page-setup", openPageSetup],
  ["file.properties", openProperties],
  ["file.version-history", () => toggleSidePane("versionHistory")],
  ["file.save-as-template", () => { savingAsTemplate.value = true; }],
  ["file.mail-merge", () => { mergeRunDialogOpen.value = true; }],
  ["file.delete", () => { confirmingDelete.value = true; }],
  // Return to the library (SPEC-060); the route guard flushes autosave first.
  ["file.close", closeOpenDocument],
  ["app.preferences", () => preferencesDialog.open()],
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
  ...rulerCommands, // View — Rulers/Horizontal/Vertical toggles; wiring in usePageSetupDialogs.
  ["view.outline-pane", () => toggleSidePane("outline")],
  ["view.styles-pane", () => toggleSidePane("styles")],
  ["view.comments-pane", () => toggleSidePane("comments")],
  // Flip the dock between left and right (SPEC-065 AC-2).
  ["view.panel-side", toggleSidePanelSide],
  ["view.comments-only", toggleCommentsOnlyView],
  // Ctrl/Cmd+Shift+E (SPEC-039 AC-1).
  ["edit.toggleTrackChanges", toggleTrackChangesRecording],
  ["view.track-changes-pane", () => toggleSidePane("trackChanges")],
  // Insert
  ["insert.table-dialog", () => { insertTableDialogOpen.value = true; }],
  // Table (SPEC-076 AC-80): the five ids useTableTools owns.
  ...tableTools.commands,
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
  // Ctrl/Cmd+Alt+K opens the link dialog (SPEC-025; moved off Ctrl/Cmd+K by SPEC-062).
  ["insert.link", openLinkDialog],
  ["insert.bookmark", openBookmarkDialog],
  ["insert.bookmark-list", () => { bookmarkManagerOpen.value = true; }],
  // Ctrl/Cmd+Alt+M (SPEC-038 AC-1).
  ["insert.comment", openCommentComposer],
  // Format
  ["format.styles-pane", () => toggleSidePane("styles")],
  ["format.import-styles", openImportDialog],
  // Ctrl/Cmd+\ (SPEC-020 AC-3).
  ["format.clearFormatting", handleClearFormatting],
  ["format.columns", openColumnsDialog], ["format.edit-points", pointEdit.enter],
  // Format → Borders and shading… (SPEC-074 AC-6); Theme… (SPEC-077 AC-9).
  ["format.borders", openBordersDialog], ["format.theme", themeGallery.openThemeGallery],
  // Drawing align/distribute (SPEC-072); wiring in useDrawingArrange.
  ...drawingArrange.commands,
  // Format → Drawing → Position and size… (SPEC-072 AC-35).
  ["drawing.position-size", openShapePosition],
  ["drawing.shape-formatting", openShapeFormat],
  ["format.multilevel-link-headings", handleLinkHeadingNumbering],
  ["format.multilevel-define-list", openDefineList],
  ["format.text-color", openTextColorMenu],
  ["format.highlight", openHighlightMenu],
  // Tools
  ["tools.word-count", openStatistics],
  ["tools.dictionary", openCustomDictionaryDialog],
  ["tools.accessibility", () => toggleSidePane("accessibility")],
  ["tools.fonts", () => toggleSidePane("fonts")],
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
  // Ctrl/Cmd+B/I/U (AC-3..5); Shift+X strikethrough (AC-6); Ctrl/Cmd+,/. sub/superscript (SPEC-013).
  ["format.bold", () => { toggleCharacterFormat("bold"); }], ["format.italic", () => { toggleCharacterFormat("italic"); }],
  ["format.underline", () => { toggleCharacterFormat("underline"); }], ["format.strikethrough", () => { toggleCharacterFormat("strikethrough"); }],
  ["format.subscript", () => { toggleCharacterFormat("subscript"); }], ["format.superscript", () => { toggleCharacterFormat("superscript"); }],
  // Ctrl/Cmd+L/E/R/J: alignment (SPEC-014 AC-3).
  ["paragraph.alignLeft", () => { handleAlign("left"); }], ["paragraph.alignCenter", () => { handleAlign("center"); }],
  ["paragraph.alignRight", () => { handleAlign("right"); }], ["paragraph.alignJustify", () => { handleAlign("justify"); }],
  // Ctrl/Cmd+Alt+1..6 apply Heading 1-6, Ctrl/Cmd+Alt+0 Normal
  // (SPEC-019 AC-5, AC-6).
  ["style.normal", () => { handleApplyStyle(NORMAL_STYLE_ID); }], ["style.heading1", () => { handleApplyStyle(headingStyleId(1)); }],
  ["style.heading2", () => { handleApplyStyle(headingStyleId(2)); }], ["style.heading3", () => { handleApplyStyle(headingStyleId(3)); }],
  ["style.heading4", () => { handleApplyStyle(headingStyleId(4)); }], ["style.heading5", () => { handleApplyStyle(headingStyleId(5)); }],
  ["style.heading6", () => { handleApplyStyle(headingStyleId(6)); }],
]);

// SPEC-021's keyboard map (ARCH-1: extracted to views/document/useKeyboardMap).
// Wired here, at the end of setup, because its deps are read when this object
// is built and the registry above is the last of them to exist.
const { handleKeyDown } = useKeyboardMap({
  ...editorContext,
  shapeSelection, imageSelection, chartSelection, tocSelection, tofSelection,
  editingShapeText, editingChartData, editingChartOptions,
  clearShapeSelection, clearImageSelection, clearChartSelection,
  clearTocSelection, clearTofSelection,
  dismissPasteChip, dismissLinkChip, closeContextMenu,
  disarmPainter, disarmShapeDraw,
  exitHeaderFooterEditing, exitNoteEditing,
  handleShapeDelete, handleShapeDuplicate, handleShapeNudge,
  handleChartDuplicate, deleteChartWithCaptionPrompt, deleteImageWithCaptionPrompt,
  handleTocDelete, handleTofDelete,
  handleTableArrowKey, handleColumnArrowKey,
  scheduleTableArrowFallback, scheduleColumnArrowFallback,
  handleLineEdgeKey, moveCaretByPage,
  applyParagraphCommand,
  armPlainPaste, clearPlainPaste,
  getPendingFormat: () => pendingFormat,
  editorCommands,
});

// Command palette (SPEC-062): wired after editorCommands so its dispatcher
// reuses that exact table — a palette run is the same call the menu makes.
const {
  commandPaletteOpen,
  paletteCommands,
  paletteStyleRows,
  paletteAvailability,
  paletteRecentIds,
  paletteCommonIds,
  openCommandPalette,
  closeCommandPalette,
  runCommandFromPalette,
} = useCommandPalette({
  ...editorContext,
  editable,
  undoRedo: undoRedoState,
  canConvertLiteralChips,
  tableSelectedForChart,
  captionTargetSelected,
  resolvedTheme: () => themeMenu.value.resolvedTheme,
  dispatch: (id) => { editorCommands.get(id)?.(); },
  applyStyle: (styleId) => { handleApplyStyle(styleId); },
});

function handleMenuCommand(id = "") {
  // "Tools → Search commands" opens the palette (AC-4); not an editorCommands
  // entry, so the catalogue completeness test stays exact.
  if (id === "app.commandPalette") {
    openCommandPalette();
    return;
  }
  editorCommands.get(id)?.();
}

</script>

<template>
  <AppShell
    v-if="doc"
    :class="viewRootClass"
    :style="{ '--web-content-width': webContentWidthCss, '--keyboard-inset-bottom': keyboardInset + 'px' }"
    :document-title="doc.title"
    :title-editable="editable"
    :format-badge="formatBadge"
    @update:document-title="handleTitleUpdate"
  >
    <template #app-bar-actions>
      <EditorSaveStatus
        :save-state="saveState"
        :save-state-text="saveStateText"
        :download-format="downloadFormat"
        :closing="closing"
        @download="handleDownload"
        @close="closeOpenDocument"
      />
    </template>
    <template #menubar>
      <DocumentMenuBar
        :table-state="tableTools.menuState.value"
        :editable="editable"
        :undo-redo="undoRedoState"
        :autocorrect-smart-quotes="autocorrectSettings.smartQuotes"
        :autocorrect-dashes="autocorrectSettings.dashes"
        :show-ruler="rulersCheckState"
        :show-horizontal-ruler="doc.showRuler !== false"
        :show-vertical-ruler="showVerticalRulerPref"
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
        :view-mode="viewMode"
        :focus-mode="focusMode"
        :zoom-percent="zoomPercent"
        :zoom-fit-mode="zoomFitMode"
        v-bind="{ ...themeMenu, ...drawingArrange.menuBindings.value }"
        @command="handleMenuCommand"
        @drawing-align-to="drawingArrange.setAlignToReference"
        @toggle-snap-guides="drawingArrange.toggleSnapToGuides"
        @toggle-snap-grid="drawingArrange.toggleSnapToGrid"
        @insert-table-size="(rows, cols) => insertTableWithAutoCaption({ rows, cols })"
        @insert-break="handleInsertBreak"
        @insert-section-break="handleInsertSectionBreak"
        @insert-note="handleInsertNote"
        @shape-pick="handleShapePick"
        @connector-pick="connectors.armConnectorDraw"
        @freeform-pick="handleFreeformPick"
        @chart-pick="handleChartPick"
        @multilevel-preset="handleMultilevelPreset"
        @insert-image-file="insertImageFileWithAutoCaption"
        @painter-click="handlePainterClick"
        @painter-dblclick="handlePainterDblClick"
      />
    </template>
    <template #toolbar>
      <DocumentToolbar
        v-if="editable"
        ref="documentToolbarRef"
        :theme="documentThemeOf(doc)"
        :reserved-right="toolbarReservedRight"
        :undo-redo="undoRedoState"
        :style-entries="styleEntries"
        :style-state="styleState"
        :format-values="formatValues"
        :document-fonts="documentFonts"
        :recent-colors="recentColors"
        :format-states="formatStates"
        :paragraph-state="paragraphState"
        :list-state="listState"
        :paragraph-pagination="paragraphPagination"
        @undo="performUndo"
        @redo="performRedo"
        @apply-style="handleApplyStyle"
        @update-style="handleUpdateStyle"
        @reset-style="handleResetStyle"
        @font-family-select="handleFontFamilySelect"
        @font-size-change="handleFontSizeChange"
        @toggle-character-format="toggleCharacterFormat"
        @text-color-select="handleTextColorSelect"
        @highlight-select="handleHighlightSelect"
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
    <RecoveryBar
      v-if="recovery"
      :label="recoveryLabel"
      @keep="keepRecovered"
      @discard="discardRecovered"
    />
    <div class="editor-layout">
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
        <div
          v-if="doc.showRuler !== false"
          class="ruler-sizer"
          :class="{ 'ruler-sizer--scaled': pagesScaled }"
          :style="zoomView.rulerSizerStyle.value"
        >
          <DocumentRuler
            :zoom="zoom"
            :page-width-cm="rulerPageWidthCm"
            :margin-left-cm="rulerMarginLeftCm"
            :margin-right-cm="rulerMarginRightCm"
            :formatting="paragraphFormatting"
            :columns="rulerColumns"
            :tab-stops="paragraphFormatting.tabStops || []"
            :unit="doc.measurementUnit || 'cm'"
            :disabled="!editable"
            @update="handleRulerUpdate"
            @open-tabs-dialog="openTabsDialog"
            @contextmenu.prevent="openRulerUnitMenu"
          />
        </div>
        <div
          class="pages-viewport"
          :class="{ 'pages-viewport--scaled': pagesScaled }"
          :style="pagesViewportStyle"
          @wheel="zoomView.onWheel"
        >
          <div
            ref="pageRef"
            class="pages"
            :class="{ 'pages--scaled': pagesScaled }"
            data-testid="document-pages"
            :data-pagination-state="paginationState"
            @mousedown="handlePageMouseDown"
            @mouseover="handlePageMouseOver"
          >
            <PageMarkupOverlays
              :find-visible="findVisible"
              :find-highlight-rects="findHighlightRects"
              :comment-highlight-rects="commentHighlightRects"
              :spell-underline-rects="spellUnderlineRects"
              @focus-thread="focusThreadFromHighlight"
            />
            <DocumentVerticalRuler
              v-if="showVerticalRuler"
              v-bind="verticalRulerProps"
              @update="handleVerticalRulerUpdate"
              @contextmenu.prevent="openRulerUnitMenu"
            />
            <TableHandlesOverlay
              v-if="tableHover"
              :geometry="tableHover"
              @handle="tables.handleTableHandle"
              @column-resize="tables.startColumnResize"
              @row-resize="tables.startRowResize"
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
            <!-- In point-edit mode the vertex handles (PointEditOverlay) are the
                 only draggable handles: hide the box resize/rotation handles so
                 they never sit on top of a vertex handle and steal its drag
                 (SPEC-071 AC-17 — a vertex on a box corner/edge is the common
                 case for a converted preset). -->
            <ShapeHandlesOverlay
              :geometry="pointEdit.active.value ? undefined : shapeHandles || undefined"
              :adjustment-handles="adjustmentHandleList()"
              :handle-style="shapeHandleStyle"
              :rotation-style="rotationHandleStyle"
              @resize-start="startShapeResize"
              @rotate-start="startShapeRotate"
              @adjust-start="startShapeAdjust"
              @adjust-keydown="handleAdjustmentKeydown"
            />
            <DrawingOverlays
              :snap-guides="snapGuides"
              :anchor-marker="shapeAnchorMarker || undefined"
              :connectors="connectors"
              :draw-preview="drawPreview || undefined"
              :freeform-d="freeformPreviewD || ''"
              :point-edit="pointEdit"
            />
            <div
              ref="contentRoot"
              class="document-content"
              :class="{ 'painter-armed': painterMode !== '', 'marks-on': marksVisible, 'is-page-grid': zoomView.pageGrid.value }"
              :style="zoomView.gridStyle.value"
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
                :page-border-style="pageView.pageBorderStyle || undefined"
                :page-background-style="pageView.pageBackgroundStyle || undefined"
                :entries="pageView.entries"
                :columns="pageView.columns"
                :virtual="pageView.virtual"
                :cell-range="cellSelection || undefined"
                :theme="documentThemeOf(doc)"
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
                :connectors="connectors.connectorsForPage(pageIndex)"
                @image-select="handleImageSelect"
                @shape-select="handleShapeSelect"
                @shape-edit-text="handleShapeEditText"
                @shape-text-beforeinput="handleShapeTextBeforeInput"
                @shape-text-keydown="handleShapeTextKeyDown"
                @shape-text-compositionstart="handleShapeTextCompositionStart"
                @shape-text-compositionend="handleShapeTextCompositionEnd"
                @chart-select="handleChartSelect"
                @chart-edit-data="handleChartEditData"
                @toc-action="handleTocAction"
                @tof-action="handleTofAction"
                @dblclick-header="handleDblclickHeader(pageIndex, $event)"
                @dblclick-footer="handleDblclickFooter(pageIndex, $event)"
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
      <SidePanel
        v-if="sidePanelOpen"
        :side="sidePanelSide"
        :width-px="sidePanelWidth"
        :collapsed="sidePanelCollapsed"
        :tabs="sidePanelTabs"
        :active-pane="sidePanelActivePane ?? ''"
        @activate="activateSidePane"
        @close="closeSidePanel"
        @toggle-side="toggleSidePanelSide"
        @toggle-collapse="toggleSidePanelCollapsed"
        @update:width-px="setSidePanelWidth"
        @commit-width="commitSidePanelWidth"
        @nudge-width="nudgeSidePanelWidth"
      >
        <OutlinePane
          v-if="outlinePaneOpen"
          v-show="sidePanelActivePane === 'outline'"
          :outline="outlineTree"
          :collapsed="outlinePaneCollapsed"
          :active-anchor="activeOutlineAnchor ?? undefined"
          @navigate="handleOutlineNavigate"
          @toggle="handleOutlineToggleCollapse"
          @move="handleOutlineMove"
        />
        <StylesPane
          v-if="stylesPaneOpen"
          v-show="sidePanelActivePane === 'styles'"
          :entries="stylesPaneEntriesList"
          :caret-paragraph-style-id="styleState.id"
          :caret-paragraph-style-name="styleState.name"
          :caret-character-style-id="characterStyleState.id"
          :caret-character-style-name="characterStyleState.name"
          @close="closeSidePane('styles')"
          @new-style="openNewStyle"
          @modify-style="openModifyStyle"
          @delete-style="openDeleteStyle"
          @apply-style="applyStyleFromPane"
        />
        <div
          v-if="commentsPaneOpen"
          v-show="sidePanelActivePane === 'comments'"
          ref="commentsDockRef"
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
            @close="closeSidePane('comments')"
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
        <TrackChangesPane
          v-if="trackChangesPaneOpen"
          v-show="sidePanelActivePane === 'trackChanges'"
          :changes="trackedChanges"
          :display-mode="trackChangesDisplayMode"
          :recording="trackChangesRecording"
          :author-name="trackChangesAuthorName"
          :editable="editable"
          @close="closeSidePane('trackChanges')"
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
        <VersionHistoryPane
          v-if="versionHistoryPaneOpen"
          v-show="sidePanelActivePane === 'versionHistory'"
          :versions="versionHistoryEntries"
          :paused-message="versionHistoryPausedMessage ?? ''"
          :named-only-filter="versionHistoryNamedOnlyFilter"
          :preview-version-id="versionHistoryPreviewVersionId ?? ''"
          :editable="documentActionsEnabled && !pastSupportedCeiling"
          @close="closeSidePane('versionHistory')"
          @update:named-only-filter="versionHistoryNamedOnlyFilter = $event"
          @select="selectVersionHistoryVersion"
          @exit-preview="exitVersionHistoryPreview"
          @restore="restoreVersionRespectingCeiling"
          @name="nameVersionHistoryVersion"
          @remove-name="removeVersionHistoryName"
          @delete="deleteVersionHistoryVersion"
          @download="downloadVersionHistoryVersion"
        />
        <AccessibilityPane
          v-if="accessibilityPaneOpen"
          v-show="sidePanelActivePane === 'accessibility'"
          :issues="accessibilityIssues"
          :scanning="accessibilityScanning"
          @close="closeSidePane('accessibility')"
          @select="goToAccessibilityIssue"
          @fix="applyAccessibilityFix"
        />
        <FontsPane
          v-if="fontsPaneOpen"
          v-show="sidePanelActivePane === 'fonts'"
          :entries="fontReport"
          :controls="embedControls"
          @close="closeSidePane('fonts')"
        />
      </SidePanel>
    </div>
    <template #status-bar>
      <DocumentStatusBar
        :caret-page="caretPage"
        :page-count="pageCount"
        :document-opening="documentOpening"
        :pagination-busy="paginationBusy"
        :word-count-text="wordCountStatusText"
        :zoom-percent="zoomPercent"
        :zoom-fit-mode="zoomFitMode"
        :current-language-tag="currentLanguageTag"
        :current-language-label="statusBarLanguageLabel"
        :language-options="SUPPORTED_LANGUAGES"
        :language-selection-active="languageSelectionActive"
        :selection-no-proof="selectionNoProof"
        :editable="editable"
        :track-changes-recording="trackChangesRecording"
        @open-statistics="openStatistics"
        @set-language="setLanguage"
        @manage-dictionary="openCustomDictionaryDialog"
        @toggle-no-proof="setNoProof"
        @toggle-track-changes="toggleTrackChangesRecording"
        @zoom-in="zoomView.zoomIn"
        @zoom-out="zoomView.zoomOut"
        @zoom-apply-input="zoomView.applyZoomInput"
        @zoom-set-preset="zoomView.setZoomPercent"
        @zoom-fit-width="zoomView.setFitWidth"
        @zoom-fit-page="zoomView.setFitPage"
      />
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
    <CloseWithoutSavingDialog
      v-if="storageWarningOpen"
      @download="handleDownload(downloadFormat)"
      @confirm="confirmCloseAnyway"
      @cancel="storageWarningOpen = false"
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
      @insert="tables.confirmRepeatRow"
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
      :unit="doc?.measurementUnit || 'cm'"
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
    <ContextualToolbar
      v-if="textSelectionToolbarVisible"
      class="print-exclude"
      testid="text-selection-toolbar"
      label="Text selection actions"
      tone="accent"
      placement="above"
      :left="textSelectionAnchor.left"
      :top="textSelectionAnchor.top"
      :anchor-bottom="textSelectionAnchor.bottom"
      @dismiss="hideTextSelectionToolbar"
    >
      <TextSelectionToolbar
        :states="formatStates"
        @toggle="toggleCharacterFormat"
        @text-color="openTextColorMenu"
        @highlight="openHighlightMenu"
        @link="openLinkDialog"
        @comment="openCommentComposer"
      />
    </ContextualToolbar>
    <ContextualToolbar
      v-if="editable && selectedField"
      class="print-exclude"
      testid="field-toolbar"
      label="Field actions"
      context="Field"
      placement="above"
      :left="fieldToolbarAnchor().left"
      :top="fieldToolbarAnchor().top"
      :anchor-bottom="fieldToolbarAnchor().bottom"
    >
      <FieldToolbar
        :locked="isFieldLocked(selectedField)"
        :read-only="isReadOnlyField(selectedField)"
        @field-action="handleFieldAction"
      />
    </ContextualToolbar>
    <ContextualToolbar
      v-if="imageSelection && imageHandles"
      class="print-exclude"
      testid="image-toolbar"
      label="Image formatting"
      context="Image"
      :left="imageToolbarStyle().left"
      :top="imageToolbarStyle().top"
      :lift="imageToolbarStyle().placement === 'above'"
    >
      <ImageToolbar
        :wrap="imageAt(doc, imageSelection)?.wrap || 'inline'"
        @wrap="handleImageWrap"
        @options="openImageOptions"
        @delete="deleteImageWithCaptionPrompt"
      />
    </ContextualToolbar>
    <ShapeToolbarIsland
      v-if="shapeSelection && shapeHandles && doc && shapeAt(doc, shapeSelection) && !movingShape"
      :box="shapeToolbarStyle()"
      :state="shapeToolbarState()"
      :format-dialog-open="shapeFormatOpen"
      :theme="documentThemeOf(doc)"
      :can-align="!drawingArrange.alignDisabledReason.value"
      :can-distribute="!drawingArrange.distributeDisabledReason.value"
      v-on="shapeToolbarOn"
      @align="(edge = '') => editorCommands.get(`drawing.align-${edge}`)?.()"
      @distribute="(axis = '') => editorCommands.get(`drawing.distribute-${axis}`)?.()"
      @edit-points="pointEdit.enter"
      @format-dismiss="shapeFormatOpen = false"
    />
    <ContextualToolbar
      v-if="chartSelection && chartHandles && doc && chartAt(doc, chartSelection)"
      class="print-exclude"
      testid="chart-toolbar"
      label="Chart formatting"
      context="Chart"
      :left="chartToolbarStyle().left"
      :top="chartToolbarStyle().top"
      :lift="chartToolbarStyle().placement === 'above'"
    >
      <ChartToolbar
        :unsupported="chartAt(doc, chartSelection)?.unsupportedChartXml !== undefined"
        @edit-data="handleChartEditData(chartSelection.section, chartSelection.block)"
        @options="openChartOptions"
        @delete="deleteChartWithCaptionPrompt"
        @duplicate="handleChartDuplicate"
      />
    </ContextualToolbar>
    <RulerUnitMenu
      v-if="rulerUnitMenu.open"
      class="print-exclude"
      :left="rulerUnitMenu.left"
      :top="rulerUnitMenu.top"
      :unit="doc?.measurementUnit || 'cm'"
      @select="selectRulerUnit"
      @dismiss="rulerUnitMenu = { ...rulerUnitMenu, open: false }"
    />
    <EditorContextMenu
      v-if="contextMenu.visible"
      class="print-exclude"
      :left="contextMenu.left"
      :top="contextMenu.top"
      :table-actions="!!contextMenuTableCell"
      :row-is-header="tableTools.contextRowIsHeader()"
      :can-merge-cells="canMergeCells"
      :list-actions="!!contextMenuListTarget"
      :field-actions="contextMenuOnField"
      :field-locked="contextMenuFieldLocked"
      :caption-action="captionTargetSelected"
      :drawing-actions="!!(shapeSelection || shapeMultiSelection.length)"
      :can-distribute-drawings="!drawingArrange.distributeDisabledReason.value"
      @cut="contextMenuCut"
      @copy="contextMenuCopy"
      @paste="contextMenuPaste"
      @table-action="(id = '') => id === 'repeat-header-row' ? tableTools.toggleContextRepeatHeaderRow() : tables.handleContextMenuTableAction(id)"
      @list-action="handleContextMenuListAction"
      @field-action="handleFieldAction"
      @caption-action="() => { closeContextMenu(); openCaptionDialog(); }"
      @drawing-action="(id) => { closeContextMenu(); editorCommands.get(`drawing.${id}`)?.(); }"
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
      @apply="tables.handleSplitCellApply"
      @cancel="splitCellDialogOpen = false"
    />
    <TablePropertiesDialog
      v-if="tablePropertiesOpen"
      :theme="documentThemeOf(doc)"
      :initial="tablePropertiesInitial"
      :unit="doc?.measurementUnit || 'cm'"
      @apply="tables.handleTablePropertiesApply"
      @cancel="tablePropertiesOpen = false"
    />
    <PageSetupDialog
      v-if="pageSetupOpen"
      :theme="documentThemeOf(doc)"
      :initial="pageSetupInitial"
      @apply="handlePageSetupApply"
      @cancel="pageSetupOpen = false"
    />
    <ColumnsDialog
      v-if="columnsDialogOpen"
      :initial="columnsDialogInitial"
      :unit="doc?.measurementUnit || 'cm'"
      @apply="handleColumnsApply"
      @cancel="columnsDialogOpen = false"
    />
    <BordersShadingDialog
      v-if="bordersDialogOpen"
      :theme="documentThemeOf(doc)"
      :initial="bordersDialogInitial"
      :unit="doc?.measurementUnit || 'cm'"
      @apply="applyBordersShading"
      @cancel="closeBordersDialog"
    />
    <TableToolsDialogs :tools="tableTools" />
    <ThemeGalleryDialog
      v-if="themeGallery.open.value"
      :current-theme="themeGallery.currentTheme.value"
      :preview-paragraphs="themeGallery.previewParagraphs.value"
      @apply="themeGallery.applyTheme"
      @cancel="themeGallery.closeThemeGallery"
    />
    <InsertTableDialog
      v-if="insertTableDialogOpen"
      @apply="tables.handleInsertTable"
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
    <CommandPalette
      v-if="commandPaletteOpen"
      :commands="paletteCommands"
      :extra-rows="paletteStyleRows"
      :availability="paletteAvailability"
      :recent-ids="paletteRecentIds"
      :common-ids="paletteCommonIds"
      @run="runCommandFromPalette"
      @dismiss="closeCommandPalette"
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
      :unit="doc?.measurementUnit || 'cm'"
      @apply="handleImageOptionsApplyAndRescan"
      @cancel="imageOptionsOpen = false"
    />
    <ShapePositionDialog
      v-if="shapePositionOpen"
      :initial="shapePositionInitial"
      :unit="doc?.measurementUnit || 'cm'"
      @apply="handleShapePositionApply"
      @dismiss="shapePositionOpen = false"
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
      :theme="documentThemeOf(doc)"
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
      :print-background="printBackground"
      @update:range-mode="printRangeMode = $event"
      @update:range-text="printRangeText = $event"
      @update:print-background="printBackground = $event"
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
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-4);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

/* The format painter button's armed state (SPEC-020 AC-11, AC-14). */
.toolbar-button[aria-pressed="true"] {
  background: var(--surface-selected);
  border-color: var(--accent);
}

/* Sticky so the conflict/too-large notice stays visible while the canvas
   scrolls (RecoveryBar carries the same look); they are mutually exclusive,
   so a single top slot suffices. */
.notice-bar {
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  align-items: center;
  gap: var(--space-4);
  padding: var(--space-3) var(--space-5);
  background: var(--surface-selected);
  border-bottom: var(--border-width) solid var(--accent-subtle);
  color: var(--text-accent);
  font-size: var(--text-md);
}

/* Since SPEC-065 the seven auxiliary panes share one docked, tabbed region
   (SidePanel.vue) instead of seven absolutely-positioned overlay docks. The
   editor is a flex row: the document column and, at ≥900px, the side panel,
   which displaces the document (below 900px the panel overlays it — SidePanel
   owns that responsive behaviour). page-wrap stays first in the DOM so focus
   runs document -> panel (AC-21); a left dock only reorders the panel
   visually. */
.editor-layout {
  position: relative;
  min-height: 100%;
  display: flex;
  align-items: flex-start;
}

/* The document column flexes to fill the space the panel leaves (AC-8d); its
   page stays centred within it (the SPEC-002 invariant now holds within the
   narrowed column rather than the whole canvas). */
.page-wrap {
  position: relative;
  flex: 1;
  min-width: 0;
  box-sizing: border-box;
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

/* The paginated canvas (SPEC-028): one relative wrapper hosting every page
   box plus the page-relative overlays (find highlights, table/image
   handles), so overlay coordinates keep working across pages. The page
   surfaces themselves live in DocPage.vue. */
.pages {
  position: relative;
}

/* SPEC-063 zoom: the viewport is sized to the SCALED footprint so scrollbars
   reflect the zoom (AC-17); `.pages` is centred inside it (left:50% offset by
   half its own unscaled width) so the top-center scale pivot (AC-16) lands on
   the viewport centre — the fix for the prior off-centre pivot. */
.pages-viewport {
  position: relative;
}

.pages-viewport--scaled {
  margin: 0 auto;
}

.pages--scaled {
  position: absolute;
  top: 0;
  left: 50%;
  translate: -50% 0; /* centre via own width (%), not a stale JS measure lagging the grid relayout (AC-36); separate `translate` keeps `transform` the bare scale matrix (AC-16) */
  transform: scale(var(--zoom, 1));
  transform-origin: top center;
}

/* SPEC-063 AC-53/AC-54: the ruler shares `.pages`' scaled footprint so it
   centres (and left-aligns/scrolls when the footprint overflows the column)
   exactly as the page does, keeping a ruler centimetre over the document
   centimetre it measures. When scaled it takes the footprint width and centres
   the ruler — which lays itself out at the zoomed page width, no transform —
   on the footprint centre (its scoped `margin: 0 auto` neutralised); at 100 %
   it is a plain passthrough. */
.ruler-sizer--scaled {
  margin: 0 auto;
  display: flex;
  justify-content: center;
}

.ruler-sizer--scaled :deep(.ruler) {
  margin-left: 0;
  margin-right: 0;
  /* The ruler's layout width is its zoomed page width — never let the flex
     box shrink it. */
  flex-shrink: 0;
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
  border: 1px solid var(--accent);
  border-radius: 2px;
  background: var(--surface-base);
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

/* SPEC-068: object handles enlarge for a coarse pointer so a fingertip can grab
   them (AC-42's touch-equivalent for handle drag; handles are already shown
   while selected rather than on hover — AC-20). The position/size numeric
   dialogs remain the fallback where a drag is impractical. */
@media (pointer: coarse) {
  .image-resize-handle {
    width: 20px;
    height: 20px;
  }
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
  padding: var(--space-3) var(--space-6);
  border-radius: var(--radius-pill);
  background: var(--surface-brand);
  color: var(--text-on-brand);
  font-size: var(--text-md);
  box-shadow: var(--shadow-2);
  pointer-events: none;
}

.not-found {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--space-3);
  color: var(--text-secondary);
}

.not-found-message {
  margin: 0;
  font-size: var(--text-xl);
}
</style>
