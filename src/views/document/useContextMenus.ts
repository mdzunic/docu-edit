// The editor's context menus: the right-click menu whose entries drive the
// clipboard pipeline (SPEC-016 AC-4..6), its table and numbering-control
// sections (SPEC-022 AC-17..19, SPEC-032 AC-14..16), the spell-check
// suggestions menu the right-click hit test hands misspelled words to
// (SPEC-035 AC-2, AC-3), and the Edit menu's Select all / Paste without
// formatting commands, which share the menu's async-clipboard read.
//
// The two right-click target variables — the table cell and the list
// paragraph a menu was opened against — stay in DocumentView as plain
// non-reactive lets: the template reads them fresh on each menu render, and
// useTables' wiring closes over the cell. This composable reaches them
// through get/set deps, the same arrangement useTables uses for
// tableFormattingLocation.
import { nextTick, ref } from "vue";
import {
  caretSelection,
  continueListNumbering,
  noPendingFormat,
  restartListNumbering,
  selectAll,
  setListNumberingValue,
} from "../../lib/commands";
import { serializeFragment } from "../../lib/clipboard";
import {
  paragraphElementForCaret,
  paragraphFromEventTarget,
  setDomSelection,
} from "../../lib/domCaret";
import { replaceMatch } from "../../lib/findReplace";
import { applyCasePattern, detectCasePattern } from "../../lib/spellcheck/suggestions";
import { paragraphAt } from "../../lib/textEditing";
import { noSpellContextMenu } from "./useSpellCheck";
import type { Ref } from "vue";
import type { ClipboardSnapshot } from "../../lib/clipboard";
import type { ModelSelection } from "../../lib/commands";
import type { ParagraphTarget, TableCellTarget } from "../../lib/domCaret";
import type { PendingFormat } from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";
import type { PasteFlavours } from "./useClipboard";
import type { SpellCheckIssue, SpellUnderlineRect } from "./useSpellCheck";

/** Reads text/html, text/plain and image presence via the async clipboard
 * API — all a button click can reach. Null when the API is missing or the
 * read was denied. A module-level pure function (captures nothing from
 * useContextMenus's closure). */
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

export interface ContextMenusDeps extends EditorContext {
  /** useCanvasGestures' tableCellAtTarget helper — it lives with the canvas
   * gesture hub because the mousedown/mousemove cell-drag handlers there use
   * it too; the view passes the destructured const through. */
  tableCellAtTarget: (eventTarget?: EventTarget | null) => TableCellTarget | null;
  /** Setter for the view's plain `contextMenuTableCell` variable (see the
   * file header): set before the menu's visible ref flips, nulled on close.
   * (Reading it back stays with useTables, the actions' target consumer.) */
  setContextMenuTableCell: (cell: TableCellTarget | null) => void;
  /** get/set over the view's plain `contextMenuListTarget` variable — the
   * numbering-control actions' target (SPEC-032 AC-14..16). */
  getContextMenuListTarget: () => ParagraphTarget | null;
  setContextMenuListTarget: (target: ParagraphTarget | null) => void;
  // --- from useSpellCheck (SPEC-035) ----------------------------------------
  spellUnderlineRects: Ref<SpellUnderlineRect[]>;
  suggestionsFor: (word: string, lang: string) => string[];
  ignoreOccurrence: (issue: SpellCheckIssue) => void;
  ignoreAllOccurrences: (word: string) => void;
  addWordToDictionary: (word: string, lang: string) => void;
  // --- from useClipboard (SPEC-016) -----------------------------------------
  getLastCopy: () => ClipboardSnapshot | null;
  performPaste: (
    flavours: PasteFlavours,
    selection: ModelSelection | null,
    plainOnly?: boolean,
  ) => void;
  /** Live write of the view's raw caret — menuSelectAll places the whole-
   * document selection itself (same pattern as useFindReplace's setCaret). */
  setCaret: (next: CaretPosition | null) => void;
  updateCaretStatus: (range?: { start: CaretPosition; end: CaretPosition }) => void;
  /** Live write of the view's pending collapsed-caret format (SPEC-012) —
   * menuSelectAll discards it, like any real caret move. */
  setPendingFormat: (next: PendingFormat | null) => void;
}

export function useContextMenus(deps: ContextMenusDeps) {
  const {
    doc,
    contentRoot,
    pageRef,
    editable,
    currentSelection,
    applyEdit,
    addToast,
    onSessionClose,
    tableCellAtTarget,
    setContextMenuTableCell,
    getContextMenuListTarget,
    setContextMenuListTarget,
    spellUnderlineRects,
    suggestionsFor,
    ignoreOccurrence,
    ignoreAllOccurrences,
    addWordToDictionary,
    getLastCopy,
    performPaste,
    setCaret,
    updateCaretStatus,
    setPendingFormat,
  } = deps;

  const contextMenu = ref({ visible: false, left: 0, top: 0 });
  const spellContextMenu = ref(noSpellContextMenu());

  // --- the editor's context menu (SPEC-016 AC-4..6) --------------------------

  /** Right-click on the canvas: the app's own menu replaces the native one so
   * its entries can drive the clipboard pipeline (and tests can reach them). */
  function handleContextMenu(event: MouseEvent): void {
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
    setContextMenuTableCell(tableCellAtTarget(event.target));
    // A right-click on a list paragraph adds the numbering-control section,
    // targeted at that paragraph (SPEC-032 AC-14..16).
    setContextMenuListTarget(listParagraphAtTarget(event.target));
    contextMenu.value = { visible: true, left: event.clientX, top: event.clientY };
  }

  /** The list paragraph under a right-click, or null when the click landed
   * outside a list item (SPEC-032). */
  function listParagraphAtTarget(eventTarget: EventTarget | null): ParagraphTarget | null {
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

  function closeContextMenu(): void {
    if (contextMenu.value.visible) {
      contextMenu.value = { visible: false, left: 0, top: 0 };
    }
    setContextMenuTableCell(null);
    setContextMenuListTarget(null);
  }

  // --- the spell-check context menu (SPEC-035 AC-2, AC-3) ---------------------

  /** The underline rect under a point, in client coordinates, or null when the
   * point is not over a flagged word. The underline layer cannot do this hit
   * test itself: its rects cover whole glyph boxes and paint above the text, so
   * making them pointer-interactive would swallow ordinary clicks, double-clicks
   * and drags (see the .spell-underline rule). The rects are stored relative to
   * the .pages box — the same space useSpellCheck measured them in. */
  function spellRectAtPoint(clientX = 0, clientY = 0): SpellUnderlineRect | null {
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
  function handleSpellContextMenu(event: MouseEvent, rect: SpellUnderlineRect): void {
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

  function closeSpellContextMenu(): void {
    spellContextMenu.value = noSpellContextMenu();
  }

  /** Choosing a suggestion (SPEC-035 AC-4, AC-4.1): replaceMatch preserves
   * the word's character formatting; applyCasePattern reshapes the
   * suggestion to the misspelling's own capitalisation pattern first. */
  function applySpellSuggestion(suggestion = ""): void {
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
  function spellIgnore(): void {
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
  function spellIgnoreAll(): void {
    const state = spellContextMenu.value;
    closeSpellContextMenu();
    ignoreAllOccurrences(state.word);
  }

  /** "Add to dictionary" (SPEC-035 AC-7): persists across sessions. */
  function spellAddToDictionary(): void {
    const state = spellContextMenu.value;
    closeSpellContextMenu();
    addWordToDictionary(state.word, state.lang);
  }

  /** A numbering-control entry picked from the context menu (SPEC-032
   * AC-14..16), applied at the right-clicked paragraph. "Set numbering
   * value…" prompts for the ordinal; cancelling the prompt changes nothing. */
  function handleContextMenuListAction(action = ""): void {
    const target = getContextMenuListTarget();
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
   * MIME type). The events land in useClipboard's handleCut/handleCopy. */
  function contextMenuCut(): void {
    closeContextMenu();
    const root = contentRoot.value as HTMLElement | null;
    if (!root) {
      return;
    }
    root.focus();
    if (typeof document.execCommand === "function") {
      document.execCommand("cut");
    }
  }

  function contextMenuCopy(): void {
    closeContextMenu();
    const root = contentRoot.value as HTMLElement | null;
    if (!root) {
      return;
    }
    root.focus();
    if (typeof document.execCommand === "function") {
      document.execCommand("copy");
    }
  }

  /** The menu's Paste (AC-6). The async API cannot carry the private MIME
   * type, so a clipboard whose text/html matches our last copy is recognised
   * as our own and pasted from the local fragment snapshot with exact
   * fidelity; anything else takes the same import paths as a keyboard
   * paste. */
  async function contextMenuPaste(): Promise<void> {
    closeContextMenu();
    const root = contentRoot.value as HTMLElement | null;
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

  // Edit-menu Select all / Paste without formatting: menu-invocable forms of
  // the Ctrl+A / Ctrl+Shift+V commands (their keyboard paths call these too).
  function menuSelectAll(): void {
    const root = contentRoot.value;
    const model = doc.value;
    if (!root || !model) {
      return;
    }
    setPendingFormat(noPendingFormat());
    const all = selectAll(model);
    setCaret(all.focus);
    updateCaretStatus({ start: all.anchor, end: all.focus });
    if (paragraphElementForCaret(root, all.anchor) && paragraphElementForCaret(root, all.focus)) {
      setDomSelection(root, all.anchor, all.focus);
    } else {
      void nextTick().then(() => setDomSelection(root, all.anchor, all.focus));
    }
  }

  async function menuPastePlain(): Promise<void> {
    const root = contentRoot.value as HTMLElement | null;
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

  // The session-close reset the view's closeSession used to call inline.
  onSessionClose(closeContextMenu);

  return {
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
  };
}
