<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import ConversionReportDialog from "../components/ConversionReportDialog.vue";
import CustomSpacingDialog from "../components/CustomSpacingDialog.vue";
import DocumentRuler from "../components/DocumentRuler.vue";
import DownloadMenu from "../components/DownloadMenu.vue";
import EditorContextMenu from "../components/EditorContextMenu.vue";
import FindBar from "../components/FindBar.vue";
import FontFamilySelect from "../components/FontFamilySelect.vue";
import FontSizeControl from "../components/FontSizeControl.vue";
import FormatToolbar from "../components/FormatToolbar.vue";
import HighlightPicker from "../components/HighlightPicker.vue";
import ParagraphToolbar from "../components/ParagraphToolbar.vue";
import PasteOptionsChip from "../components/PasteOptionsChip.vue";
import TextColorPicker from "../components/TextColorPicker.vue";
import UndoRedoControls from "../components/UndoRedoControls.vue";
import { useAutosave } from "../composables/useAutosave";
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
import { asCharacterFormatKey, asParagraphAlignment, DEFAULT_TITLE } from "../lib/documentModel";
import { documentFontFamilies } from "../lib/fonts";
import { loadRecentColors, noHighlight, pushRecentColor } from "../lib/colors";
import { acquireEditLock, noEditLock } from "../lib/editLock";
import {
  BLOCK_SELECTOR,
  caretFromDomPoint,
  domPointForCaret,
  extendDomSelection,
  modelRangeFromSelection,
  noElement,
  setDomCaret,
  setDomSelection,
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
} from "../lib/caretGeometry";
import { blockText, compareCaret, noCaret } from "../lib/textEditing";
import {
  asPasteMode,
  clipboardFlavours,
  DOC_MODEL_MIME,
  executePaste,
  extractFragment,
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
  applyParagraphFormat,
  blockCount,
  blockOrdinal,
  caretSelection,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentEnd,
  documentStart,
  emptyParagraphFormatting,
  emptyToolbarFormatValues,
  firstTouchedFormatting,
  formatState,
  hasSpaceBefore,
  INDENT_STEP_CM,
  insertLineBreak,
  insertText,
  isCollapsed,
  noCommandResult,
  noParagraphCommand,
  noPendingFormat,
  noSelection,
  selectAll,
  setAlignment,
  setPageMargins,
  splitBlock,
  toggleFormat,
  toggleSpaceBefore,
  toggleVerticalAlign,
  toolbarFormatValues,
  verticalAlignState,
} from "../lib/commands";
import { headingTag, paragraphStyle, runKey, runStyle } from "../lib/renderFormatting";
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
// `plainPasteNext` is armed by the Ctrl/Cmd+Shift+V keydown (AC-27): the
// browser still fires a `paste` event for that combination, and the flag
// routes it down the plain-text path. `pastePayload`/`pastePendingFormat`
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
// The custom-spacing dialog (AC-12..14): open flag plus the prefill taken
// from the first touched paragraph at the moment it opens.
const customSpacingOpen = ref(false);
const customSpacingInitial = ref({ mode: "multiple", value: 1, beforePt: 0, afterPt: 0 });
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
  dismissPasteChip();
  closeContextMenu();
  resetFindState();
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
  }
}

onMounted(() => {
  document.addEventListener("selectionchange", handleSelectionChange);
  document.addEventListener("keydown", handleFindShortcut);
  attachScrollListener();
  if (doc.value) {
    openSession(documentId.value);
  }
});

onBeforeUnmount(() => {
  closeSession();
  cancelFindScheduled();
  document.removeEventListener("selectionchange", handleSelectionChange);
  document.removeEventListener("keydown", handleFindShortcut);
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
  autosave.markDirty();
  updateCaretStatus({ start: snapshot.selection.anchor, end: snapshot.selection.focus });
  refreshUndoRedoState();
  root.focus();
  void restoreCaretAfterRender();
  scheduleFindRefresh();
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
    }
  } else if (type === "insertParagraph") {
    result = splitBlock(model, selection);
    label = "new paragraph";
  } else if (type === "insertLineBreak") {
    // Shift+Enter: a forced line break, stored as a literal "\n" that
    // pre-wrap rendering displays without a new paragraph.
    result = insertLineBreak(model, selection);
    label = "line break";
  } else if (type === "deleteContentBackward") {
    result = deleteBackward(model, selection);
    label = "delete";
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
  }
}

async function restoreCaretAfterRender() {
  await nextTick();
  const root = contentRoot.value;
  if (!root || !caret) {
    return;
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
  const expected = model.sections
    .map((section) => section.blocks.map(blockText).join(""))
    .join("");
  if (root.textContent !== expected) {
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

// --- the editor's context menu (SPEC-016 AC-4..6) --------------------------

/** Right-click on the canvas: the app's own menu replaces the native one so
 * its entries can drive the clipboard pipeline (and tests can reach them). */
function handleContextMenu(event = new MouseEvent("contextmenu")) {
  if (!doc.value || !editable.value) {
    return;
  }
  event.preventDefault();
  contextMenu.value = { visible: true, left: event.clientX, top: event.clientY };
}

function closeContextMenu() {
  if (contextMenu.value.visible) {
    contextMenu.value = { visible: false, left: 0, top: 0 };
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
  if (extend) {
    extendDomSelection(root, target);
  } else {
    setDomCaret(root, target);
  }
  caret = target;
  const range = modelRangeFromSelection(root, document.getSelection());
  updateCaretStatus(range ?? { start: target, end: target });
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
function moveCaretByPage(direction = 1, extend = false) {
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

// --- character formatting toolbar/shortcuts (SPEC-012, SPEC-013) -----------

/** Bold/italic/underline/strikethrough (SPEC-012 AC-3..6, AC-16, AC-17)
 * and subscript/superscript (SPEC-013): a real selection is formatted
 * through the command API (which keeps the selection live, see applyEdit);
 * a collapsed caret instead sets/flips the pending format the next typed
 * text will carry (AC-9, AC-10). */
function toggleCharacterFormat(key = "") {
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
function handleRulerUpdate(update = { kind: "", valueCm: 0 }) {
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
  }
}

function handleKeyDown(event = new KeyboardEvent("keydown")) {
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value || event.altKey) {
    return;
  }
  const mod = event.ctrlKey || event.metaKey;
  const key = event.key.toLowerCase();
  if (mod && (event.key === "Home" || event.key === "End")) {
    // Ctrl/Cmd+Home/End: start/end of the document, Shift extends.
    event.preventDefault();
    applyCaretMove(
      event.key === "Home" ? documentStart() : documentEnd(model),
      event.shiftKey,
    );
    if (scrollContainer) {
      scrollContainer.scrollTop =
        event.key === "Home" ? 0 : scrollContainer.scrollHeight;
    }
  } else if (mod && !event.shiftKey && key === "z") {
    // Ctrl/Cmd+Z: undo (SPEC-015 AC-1).
    event.preventDefault();
    performUndo();
  } else if (mod && event.shiftKey && key === "z") {
    // Ctrl/Cmd+Shift+Z: redo (SPEC-015 AC-2).
    event.preventDefault();
    performRedo();
  } else if (mod && !event.shiftKey && key === "y") {
    // Ctrl+Y: redo, the Windows convention (SPEC-015 AC-3).
    event.preventDefault();
    performRedo();
  } else if (mod && event.shiftKey && key === "v") {
    // Ctrl/Cmd+Shift+V (SPEC-016 AC-27): arm the plain-text flag and let
    // the browser fire the paste event it still produces for this combo —
    // only that event can hand us the clipboard's contents.
    plainPasteNext = true;
  } else if (mod && !event.shiftKey && key === "v") {
    // A plain Ctrl/Cmd+V clears a stale plain-paste flag (armed but never
    // consumed because an empty clipboard fired no paste event).
    plainPasteNext = false;
  } else if (event.key === "Escape") {
    // Escape dismisses the paste options chip and the context menu.
    dismissPasteChip();
    closeContextMenu();
  } else if (mod && !event.shiftKey && key === "a") {
    // Ctrl/Cmd+A: select the whole document, exactly to the model's extent.
    event.preventDefault();
    pendingFormat = noPendingFormat();
    const all = selectAll(model);
    setDomSelection(root, all.anchor, all.focus);
    caret = all.focus;
    updateCaretStatus({ start: all.anchor, end: all.focus });
  } else if (mod && !event.shiftKey && (key === "b" || key === "i" || key === "u")) {
    // Ctrl/Cmd+B/I/U (AC-3, AC-4, AC-5).
    event.preventDefault();
    toggleCharacterFormat(key === "b" ? "bold" : key === "i" ? "italic" : "underline");
  } else if (mod && event.shiftKey && key === "x") {
    // Ctrl/Cmd+Shift+X: strikethrough (AC-6).
    event.preventDefault();
    toggleCharacterFormat("strikethrough");
  } else if (mod && !event.shiftKey && (key === "," || key === ".")) {
    // Ctrl/Cmd+, and Ctrl/Cmd+. : subscript / superscript (SPEC-013).
    event.preventDefault();
    toggleCharacterFormat(key === "," ? "subscript" : "superscript");
  } else if (mod && !event.shiftKey && (key === "l" || key === "e" || key === "r" || key === "j")) {
    // Ctrl/Cmd+L/E/R/J: alignment (SPEC-014 AC-3).
    event.preventDefault();
    handleAlign(key === "l" ? "left" : key === "e" ? "center" : key === "r" ? "right" : "justify");
  } else if (!mod && event.key === "Tab") {
    // SPEC-014 AC-8..10. A caret collapsed at offset 0 — an empty paragraph
    // or the start of its first line — dispatches the toolbar's own indent
    // command; a wrapped continuation line's visual start sits at a non-zero
    // offset, so it falls through to the literal tab. Always prevented:
    // the browser's default would move focus out of the editor.
    event.preventDefault();
    const selection = currentSelection(root);
    if (!selection) {
      return;
    }
    if (isCollapsed(selection) && selection.focus.offset === 0) {
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
  } else if (event.key === "PageDown" || event.key === "PageUp") {
    event.preventDefault();
    moveCaretByPage(event.key === "PageDown" ? 1 : -1, event.shiftKey);
  }
}

/** Clicking anywhere on the page canvas outside the text (margins, empty
 * space) places the caret at the nearest text position. Clicks inside the
 * editable content keep the browser's native precise placement. */
function handlePageMouseDown(event = new MouseEvent("mousedown")) {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  if (event.target instanceof Node && root.contains(event.target)) {
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

const pageStyle = computed(() => {
  if (!doc.value) {
    return {};
  }
  const { widthCm, heightCm } = doc.value.defaultStyles.pageSize;
  return { width: `${widthCm}cm`, height: `${heightCm}cm` };
});

const marginStyle = computed(() => {
  if (!doc.value) {
    return {};
  }
  const { topCm, rightCm, bottomCm, leftCm } = doc.value.defaultStyles.margins;
  return {
    top: `${topCm}cm`,
    right: `${rightCm}cm`,
    bottom: `${bottomCm}cm`,
    left: `${leftCm}cm`,
  };
});

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
  return doc.value.sections.flatMap((section, sectionIndex) =>
    section.blocks.map((block, blockIndex) => {
      const address = `${sectionIndex}:${blockIndex}`;
      const suffix = bump.key === address ? `:${bump.count}` : "";
      return {
        key: `${bump.epoch}:${address}${suffix}`,
        block,
        section: sectionIndex,
        blockIndex,
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
let findRafId = 0;
let findDebounceTimer = 0;
let findHintTimer = 0;

function currentFindOptions() {
  return {
    matchCase: findMatchCase.value,
    wholeWords: findWholeWords.value,
    regex: findUseRegex.value,
  };
}

/** Ctrl/Cmd+F (and Ctrl/Cmd+H, which also reveals the replace row) at the
 * document level, so the shortcut works whether focus is in the canvas or
 * already in the bar; always prevented so the browser's own find never
 * opens. */
function handleFindShortcut(event = new KeyboardEvent("keydown")) {
  if (!doc.value || !editable.value || event.altKey) {
    return;
  }
  const mod = event.ctrlKey || event.metaKey;
  if (!mod || event.shiftKey) {
    return;
  }
  const key = event.key.toLowerCase();
  if (key === "f") {
    event.preventDefault();
    openFind(false);
  } else if (key === "h") {
    event.preventDefault();
    openFind(true);
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
  if (findRafId !== 0) {
    return;
  }
  findRafId = window.requestAnimationFrame(() => {
    findRafId = 0;
    measureFindHighlights();
  });
}

function scheduleScrollCurrentIntoView() {
  window.requestAnimationFrame(scrollCurrentMatchIntoView);
}

function cancelFindScheduled() {
  if (findRafId !== 0) {
    window.cancelAnimationFrame(findRafId);
    findRafId = 0;
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
 * screen — the only ones whose matches are worth measuring. */
function visibleFindBlockKeys(root = noElement(), bounds = { top: 0, bottom: 0 }) {
  const keys = new Set();
  if (!root) {
    return keys;
  }
  const margin = 200;
  for (const element of root.querySelectorAll(BLOCK_SELECTOR)) {
    const rect = element.getBoundingClientRect();
    if (rect.bottom < bounds.top - margin || rect.top > bounds.bottom + margin) {
      continue;
    }
    keys.add(`${element.getAttribute("data-section")}:${element.getAttribute("data-block")}`);
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
        @align="handleAlign"
        @indent="handleIndent"
        @spacing-preset="handleSpacingPreset"
        @custom-spacing="openCustomSpacing"
        @toggle-space-before="handleToggleSpaceBefore"
        @toggle-direction="handleToggleDirection"
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
        :page-width-cm="doc.defaultStyles.pageSize.widthCm"
        :margin-left-cm="doc.defaultStyles.margins.leftCm"
        :margin-right-cm="doc.defaultStyles.margins.rightCm"
        :formatting="paragraphFormatting"
        :disabled="!editable"
        @update="handleRulerUpdate"
      />
      <div
        ref="pageRef"
        class="page"
        data-testid="document-page"
        :style="pageStyle"
        @mousedown="handlePageMouseDown"
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
          class="page-margin"
          data-testid="page-margin"
          :style="marginStyle"
        >
          <div
            ref="contentRoot"
            class="document-content"
            data-testid="document-content"
            role="textbox"
            aria-multiline="true"
            aria-label="Document content"
            :aria-readonly="editable ? undefined : 'true'"
            :contenteditable="editable"
            spellcheck="false"
            @beforeinput="handleBeforeInput"
            @keydown="handleKeyDown"
            @compositionstart="handleCompositionStart"
            @compositionend="handleCompositionEnd"
            @copy="handleCopy"
            @cut="handleCut"
            @paste="handlePaste"
            @contextmenu="handleContextMenu"
          >
            <component
              :is="headingTag(entry.block.headingLevel)"
              v-for="entry in flatBlocks"
              :key="entry.key"
              v-memo="[entry.block, entry.section, entry.blockIndex]"
              class="doc-paragraph"
              :style="paragraphStyle(entry.block)"
              :data-section="entry.section"
              :data-block="entry.blockIndex"
            >
              <template v-if="entry.block.runs.length > 0">
                <span
                  v-for="(run, runIndex) in entry.block.runs"
                  :key="runKey(run, runIndex)"
                  :style="runStyle(run.formatting)"
                >{{ run.text }}</span>
              </template>
              <br v-else>
            </component>
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
    </div>
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
    <PasteOptionsChip
      v-if="pasteChip.visible"
      :left="pasteChip.left"
      :top="pasteChip.top"
      :applied="pasteChip.applied"
      @select="handlePasteOption"
      @dismiss="dismissPasteChip"
    />
    <EditorContextMenu
      v-if="contextMenu.visible"
      :left="contextMenu.left"
      :top="contextMenu.top"
      @cut="contextMenuCut"
      @copy="contextMenuCopy"
      @paste="contextMenuPaste"
      @dismiss="closeContextMenu"
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

/* A page-shaped surface sized to the document's own defaultStyles (in real
   centimetres via CSS's absolute cm unit) — later specs render content onto
   it. .page-margin marks where the page margins end and content begins. */
.page {
  position: relative;
  margin: 0 auto;
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3);
  box-sizing: border-box;
}

.page-margin {
  position: absolute;
  border: 1px dashed #c8c8c8;
}

.document-content {
  height: 100%;
  overflow: hidden;
  color: #202124;
  outline: none;
}

.doc-paragraph {
  margin: 0;
  font-weight: normal;
  font-size: inherit;
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
