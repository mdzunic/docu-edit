// SPEC-021's keyboard map: the one place a keydown in the canvas is turned
// into an editor action.
//
// Extracted from DocumentView (ARCH-1). It was the view's largest single
// function by a wide margin — 253 lines against a 4,582-line budget the file
// had one line left of — and it is feature logic, not wiring, which is what
// the convention in CLAUDE.md says belongs in a composable here.
//
// The dependency list below is long, and honestly so: a keyboard map is the
// one surface that genuinely reaches every feature, because every feature has
// a key. Splitting it by feature would scatter the precedence order this file
// encodes — the sequence of early returns IS the specification of which
// surface wins a key when several want it (a selected shape's Delete before
// the caret's, a selected TOC's typing guard before ordinary input), and that
// order is only reviewable when it is readable top to bottom in one place.
//
// Because the deps are read at setup, DocumentView must call this AFTER every
// composable it names — in practice at the end of setup, once
// `editorCommands` (the shortcut registry) exists.
import type { Ref } from "vue";
import { adjustIndent, changeListLevel, documentEnd, documentStart, insertText } from "../../lib/commands";
import { appendTableRowForTab } from "../../lib/tables";
import {
  isContentKey,
  isDeleteKey,
  isDuplicateCombo,
  isPlainPasteCombo,
  isUnmodifiedArrowKey,
  nudgeDirectionForKey,
} from "../../lib/editorKeys";
import { resolveTabAction } from "../../lib/tabKey";
import { findShortcut } from "../../lib/shortcuts";
import type { NudgeDirection } from "../../lib/core/shapes";
import type { PendingFormat } from "../../lib/documentModel";
import type { ParagraphCommand } from "../../lib/commands";
import type { EditorContext } from "./context";

/** A selection or editing-mode ref the map reads only for PRESENCE — every
 * use is a bare truthiness check on `.value`, deciding whether that surface
 * gets the key before the caret does. The map never looks inside the value,
 * so naming each surface's own type here would add a dependency it does not
 * have; `unknown` is the accurate width, not a shortcut around one. */
type ActiveFlag = Readonly<Ref<unknown>>;

export interface UseKeyboardMapDeps extends EditorContext {
  // --- selections the map must clear, or act on before the caret does ------
  shapeSelection: ActiveFlag;
  imageSelection: ActiveFlag;
  chartSelection: ActiveFlag;
  tocSelection: ActiveFlag;
  tofSelection: ActiveFlag;
  editingShapeText: ActiveFlag;
  editingChartData: ActiveFlag;
  editingChartOptions: ActiveFlag;
  clearShapeSelection: () => void;
  clearImageSelection: () => void;
  clearChartSelection: () => void;
  clearTocSelection: () => void;
  clearTofSelection: () => void;

  // --- popups and modes Escape dismisses -----------------------------------
  dismissPasteChip: () => void;
  dismissLinkChip: () => void;
  closeContextMenu: () => void;
  disarmPainter: () => void;
  disarmShapeDraw: () => void;
  exitHeaderFooterEditing: () => void;
  exitNoteEditing: () => void;

  // --- feature actions a key dispatches to ---------------------------------
  handleShapeDelete: () => void;
  handleShapeDuplicate: () => void;
  handleShapeNudge: (direction: NudgeDirection, fine: boolean) => void;
  handleChartDuplicate: () => void;
  deleteChartWithCaptionPrompt: () => void;
  deleteImageWithCaptionPrompt: () => void;
  handleTocDelete: () => void;
  handleTofDelete: () => void;
  handleTableArrowKey: (event?: KeyboardEvent) => boolean;
  handleColumnArrowKey: (event?: KeyboardEvent) => boolean;
  /** Both fallbacks take the key NAME, not the event (usePageNavigation). */
  scheduleTableArrowFallback: (key?: string) => void;
  scheduleColumnArrowFallback: (key?: string) => void;
  handleLineEdgeKey: (event?: KeyboardEvent) => boolean;
  moveCaretByPage: (direction?: number, extend?: boolean) => Promise<void>;
  applyParagraphCommand: (command?: ParagraphCommand | null, label?: string, refocus?: boolean) => void;

  // --- paste bookkeeping (SPEC-016) ----------------------------------------
  armPlainPaste: () => void;
  clearPlainPaste: () => void;

  /** The pending character format (SPEC-012) a typed character should take.
   * A getter, not a ref: it is a plain `let` in the view, and must be read
   * at keystroke time rather than captured when the deps are built. */
  getPendingFormat: () => PendingFormat | null;
  /** The SPEC-021 shortcut registry: command id -> handler. */
  editorCommands: Map<string, () => void>;
}

export function useKeyboardMap(deps: UseKeyboardMapDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    applyEdit,
    applyCaretMove,
    clearCellSelection,
    getScrollContainer,
    addToast,
    shapeSelection,
    imageSelection,
    chartSelection,
    tocSelection,
    tofSelection,
    editingShapeText,
    editingChartData,
    editingChartOptions,
    clearShapeSelection,
    clearImageSelection,
    clearChartSelection,
    clearTocSelection,
    clearTofSelection,
    dismissPasteChip,
    dismissLinkChip,
    closeContextMenu,
    disarmPainter,
    disarmShapeDraw,
    exitHeaderFooterEditing,
    exitNoteEditing,
    handleShapeDelete,
    handleShapeDuplicate,
    handleShapeNudge,
    handleChartDuplicate,
    deleteChartWithCaptionPrompt,
    deleteImageWithCaptionPrompt,
    handleTocDelete,
    handleTofDelete,
    handleTableArrowKey,
    handleColumnArrowKey,
    scheduleTableArrowFallback,
    scheduleColumnArrowFallback,
    handleLineEdgeKey,
    moveCaretByPage,
    applyParagraphCommand,
    armPlainPaste,
    clearPlainPaste,
    getPendingFormat,
    editorCommands,
  } = deps;

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
      clearChartSelection();
      clearTocSelection();
      clearTofSelection();
      exitHeaderFooterEditing();
      exitNoteEditing();
      return;
    }
    if (shapeSelection.value && !editingShapeText.value) {
      // A selected shape's keyboard interactions (SPEC-051 AC-40..43): a
      // dedicated keydown path, the same reason images.ts's own image-delete
      // path is one — a shape click never places a DOM caret/selection inside
      // the contenteditable region.
      if (isDeleteKey(event)) {
        event.preventDefault();
        handleShapeDelete();
        return;
      }
      if (isDuplicateCombo(event)) {
        event.preventDefault();
        handleShapeDuplicate();
        return;
      }
      const direction = nudgeDirectionForKey(event.key);
      if (direction) {
        event.preventDefault();
        handleShapeNudge(direction, event.ctrlKey || event.metaKey);
        return;
      }
    }
    if (chartSelection.value && !editingChartData.value && !editingChartOptions.value) {
      // A selected chart's keyboard interactions (SPEC-052) — the same
      // dedicated keydown path shapes/images use: a chart click never places
      // a DOM caret/selection inside a contenteditable region.
      if (isDeleteKey(event)) {
        event.preventDefault();
        deleteChartWithCaptionPrompt();
        return;
      }
      if (isDuplicateCombo(event)) {
        event.preventDefault();
        handleChartDuplicate();
        return;
      }
    }
    if (tocSelection.value && isDeleteKey(event)) {
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
    if (imageSelection.value && isDeleteKey(event)) {
      // Deleting a selected image (SPEC-024 AC-30) is a dedicated keydown
      // path, not the native beforeinput deletion flow: an image click never
      // places a DOM caret/selection inside the contenteditable region for
      // that flow to act on.
      event.preventDefault();
      deleteImageWithCaptionPrompt();
      return;
    }
    if (tofSelection.value && isDeleteKey(event)) {
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
    if (isUnmodifiedArrowKey(event)) {
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
    if (isPlainPasteCombo(event)) {
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
        const scrollContainer = getScrollContainer();
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
        const selection = currentSelection(root as HTMLElement);
        if (!selection) {
          break;
        }
        // Which of SPEC-054's four rules applies is model-derived, and lives
        // in lib/tabKey.ts where it can be unit-tested; this only applies it.
        const action = resolveTabAction(model, selection, event.shiftKey);
        if (action.kind === "moveCaret") {
          applyCaretMove(action.caret);
        } else if (action.kind === "appendRow") {
          applyEdit(
            appendTableRowForTab(model, selection, action.location),
            selection,
            "add table row",
          );
        } else if (action.kind === "listLevel") {
          applyParagraphCommand(
            (current, target) => changeListLevel(current, target, action.delta),
            "list level",
          );
        } else if (action.kind === "indent") {
          applyParagraphCommand(
            (current, target) => adjustIndent(current, target, action.deltaCm),
            "indent",
          );
        } else if (action.kind === "insertTab") {
          applyEdit(insertText(model, selection, "\t", getPendingFormat()), selection, "typing", { groupable: true });
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

  return { handleKeyDown };
}
