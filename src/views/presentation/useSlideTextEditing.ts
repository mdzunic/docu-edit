// The slide text-editing session (SPEC-081): double-click (or Enter) starts an
// in-place contenteditable session; typing/Enter/Backspace/Delete/caret moves
// are the browser's, and after every input `normalizeTextFrameDom` walks the
// host back into the frame's `paragraphs` (normalise-after-input, so the DOM
// never diverges). Formatting NEVER round-trips through execCommand — a toolbar
// or keyboard command edits the model (src/lib/presentation/text.ts) and
// re-renders the host. Typing coalesces into one undo entry per burst (a pause
// ≥ 1s, ending the session, or any non-typing command closes it); formatting is
// one entry each. Every edit marks the slide dirty and persists.
//
// SPEC-091 extends the SAME machinery to table cells: a `TableFrame`'s cell is
// a `ParagraphOwner` (it has `paragraphs`), so `startEditingCell` opens a
// session against a cell's `SlideParagraph[]` and every formatting command,
// toolbar control and the DOM render/normalise path work unchanged. Tab/
// Shift+Tab move between non-covered cells (Tab in the last cell appends a row);
// arrow keys are left to the browser so the caret can leave the table at its
// edges.
//
// SPEC-097: Tab/Shift+Tab demote/promote; Cmd/Ctrl+K opens the link dialog;
// every committed text edit tells the after-edit hook (the autofit refit).
//
// SPEC-098: a find hit opens a session on the frame it is in — a shape's text
// too (a `ShapeFrame` with `paragraphs` is a `ParagraphOwner` like a cell), so
// the match is selected in the same host. Only `startEditing` reaches a shape:
// double-click and Enter still start sessions on text frames alone.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, ShapeFrame, Slide, SlideFrame, TableFrame, TextFrame } from "../../lib/presentation/model";
import { deckThemeOf } from "../../lib/presentation/theme";
import type { ModelColor } from "../../lib/core/theme";
import {
  applyParagraphFormat,
  applyRunFormat,
  demote,
  markSlideEdited,
  promote,
  selectionFormatState,
  toggleBullet,
  wholeFrameRange,
  type ParagraphFormatPatch,
  type ParagraphOwner,
  type RunFormatPatch,
  type SelectionFormatState,
  type TextRange,
} from "../../lib/presentation/text";
import { insertTableRow, nextCellAddress, ownerCellAt, prevCellAddress } from "../../lib/presentation/tables";
import {
  normalizeTextFrameDom,
  placeCaretAtPoint,
  renderEditableParagraphs,
  selectAllIn,
  selectionRangeOf,
  setSelectionRange,
} from "../../components/presentation/slideTextDom";
import { defaultParagraphAlign } from "../../components/presentation/slideRender";
import type { FrameSelection } from "./useFrameSelection";
import {
  copyPlainText,
  createTypingBurst,
  insertPlainText,
  isCollapsed,
  paragraphSpan,
  placeCaretAtEnd,
  reidentifyEditedFrame,
  type CaretPoint,
  type CellAddress,
  type ParagraphCommand,
  type TextCommand,
  type TextCommandTarget,
} from "./slideTextSession";

export type { CaretPoint, CellAddress } from "./slideTextSession";

/** A shape frame holding text — what a session edits in a shape (SPEC-098). */
type ShapeTextFrame = ShapeFrame & ParagraphOwner;
function isShapeText(frame: SlideFrame): frame is ShapeTextFrame {
  return frame.kind === "shape" && Array.isArray(frame.paragraphs);
}

/** How long a pause closes a typing burst (SPEC-081: ≥ 1s). */
export const TYPING_BURST_MS = 1000;

export interface SlideTextEditingDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  selection: FrameSelection;
  /** Snapshot pre-command state for undo (usePresentationUndo.record). */
  record: (label?: string) => void;
  /** Re-render the shallowRef + persist write-through (PresentationView). */
  commit: () => void;
}

export interface SlideTextEditing {
  editingFrameId: Ref<string | null>;
  /** The cell of the frame being edited (SPEC-091), or null for a text frame /
   * no session. Drives the stage's in-cell contenteditable host. */
  editingCell: Ref<CellAddress | null>;
  /** The table cell the panel's commands anchor to — the last cell clicked or
   * edited (SPEC-091). Persists after a session ends so the Table panel keeps
   * acting on it; PresentationView mirrors it into useTables. */
  activeCell: Ref<CellAddress | null>;
  /** Whether the toolbar's formatting controls act on something (an editing
   * session, or a selected text frame). */
  toolbarEnabled: ComputedRef<boolean>;
  /** The current selection's formatting for the toolbar (null when nothing is
   * targeted). */
  toolbarState: Ref<SelectionFormatState | null>;
  /** `range` (SPEC-098's find navigation) opens the session with it selected. */
  startEditing: (frameId: string, selectAll?: boolean, point?: CaretPoint, range?: TextRange) => void;
  /** Starts a session inside a table cell (AC-6a). */
  startEditingCell: (frameId: string, row: number, col: number, point?: CaretPoint, range?: TextRange) => void;
  /** Single-click a table cell: anchor the panel there and move an active
   * session onto it, without starting one otherwise (AC-9c). */
  focusCell: (frameId: string, row: number, col: number) => void;
  stopEditing: () => void;
  registerHost: (el: HTMLElement | null) => void;
  onInput: () => void;
  onKeydown: (event: KeyboardEvent) => void;
  onPaste: (event: ClipboardEvent) => void;
  // Formatting commands (toolbar + keyboard).
  toggleBold: () => void;
  toggleItalic: () => void;
  toggleUnderline: () => void;
  setFontFamily: (family: string) => void;
  setFontSize: (sizePt: number) => void;
  setColor: (color: ModelColor) => void;
  setAlign: (align: NonNullable<TextFrame["paragraphs"][number]["align"]>) => void;
  toggleBulletCmd: () => void;
  indent: (delta: 1 | -1) => void;
  // SPEC-097 hooks for useSlideTextDepth (types: slideTextSession.ts).
  /** What a command would act on now (the live/last DOM selection in a
   * session, else the whole selected text frame), or null. */
  commandTarget: () => TextCommandTarget | null;
  /** ONE undo entry: record → apply → markSlideEdited → re-render → commit;
   * `range` overrides the command range. False when nothing is targeted. */
  runTextCommand: (label: string, apply: TextCommand, range?: TextRange) => boolean;
  runParagraphCommand: (label: string, apply: ParagraphCommand) => boolean;
  /** Registers what Cmd/Ctrl+K does inside a session (the link dialog). */
  setLinkShortcut: (handler: (() => void) | null) => void;
  /** Registers what runs, with the frame's id, after each text edit committed
   * here (typing, paste, cut, a command, the end of a session that changed
   * text) — already marked edited, so autofit may write back (AC-38/41). */
  setAfterEdit: (handler: ((frameId: string) => void) | null) => void;
}

export function useSlideTextEditing(deps: SlideTextEditingDeps): SlideTextEditing {
  const editingFrameId = ref<string | null>(null);
  const editingCell = ref<CellAddress | null>(null);
  const activeCell = ref<CellAddress | null>(null);
  const host = ref<HTMLElement | null>(null);
  // Bumped when the DOM caret/selection moves (no model change) so the toolbar
  // state recomputes; model edits recompute it through the reactive frame.
  const selectionTick = ref(0);
  let suppressInput = false;

  // --- undo-burst state -----------------------------------------------------
  const { note: noteTyping, close: closeBurst } = createTypingBurst((label) => deps.record(label), "Typing", TYPING_BURST_MS);

  // Fresh frame + slide identities so the read-only views re-render (AC-43/44).
  const bumpFrame = (frame: SlideFrame): void => reidentifyEditedFrame(deps.model.value, deps.currentSlide.value, frame.id);

  // --- the after-edit hook (SPEC-097) ---------------------------------------
  let afterEdit: ((frameId: string) => void) | null = null;
  let sessionEdited = false; // so a session that changed nothing never refits
  function edited(frame: SlideFrame): void {
    if (editingFrameId.value) sessionEdited = true;
    afterEdit?.(frame.id);
  }

  // --- frame / owner lookup -------------------------------------------------
  function frameById(frameId: string | null): SlideFrame | null {
    const slide = deps.currentSlide.value;
    if (!slide || frameId === null) return null;
    return slide.frames.find((f) => f.id === frameId) ?? null;
  }
  /** A frame a session can open on: a text frame, or a shape holding text. */
  function sessionFrameById(frameId: string | null): TextFrame | ShapeTextFrame | null {
    const frame = frameById(frameId);
    return frame && (frame.kind === "text" || isShapeText(frame)) ? frame : null;
  }
  function tableFrameById(frameId: string | null): TableFrame | null {
    const frame = frameById(frameId);
    return frame && frame.kind === "table" ? frame : null;
  }

  /** The frame that currently holds an editing session (text or table). */
  const editingFrame = computed<SlideFrame | null>(() => frameById(editingFrameId.value));

  /** The paragraph list the session edits: a text frame's own, or the addressed
   * table cell's. Null when no session (or the cell no longer exists). */
  const editingOwner = computed<ParagraphOwner | null>(() => {
    const frame = editingFrame.value;
    if (!frame) return null;
    if (frame.kind === "text" || isShapeText(frame)) return frame;
    if (frame.kind === "table" && editingCell.value) {
      const cell = frame.rows[editingCell.value.row]?.cells[editingCell.value.col];
      return cell && !cell.covered ? cell : null;
    }
    return null;
  });

  /** The frame a command targets: the one being edited, else a selected text
   * frame (AC-25). Used to re-clone for reactivity. */
  function targetFrame(): SlideFrame | null {
    if (editingFrame.value) return editingFrame.value;
    const selected = deps.selection.selectedFrame.value;
    return selected && selected.kind === "text" ? selected : null;
  }
  /** The paragraph owner a command formats: the edited cell/frame, else a
   * selected text frame. */
  function targetOwner(): ParagraphOwner | null {
    if (editingOwner.value) return editingOwner.value;
    const selected = deps.selection.selectedFrame.value;
    return selected && selected.kind === "text" ? selected : null;
  }

  const toolbarEnabled = computed(() => targetOwner() !== null);

  // Reactive over the target owner's model (via commit's triggerRef) AND the
  // DOM selection (via selectionTick), so it survives undo/redo and caret moves.
  const toolbarState = computed<SelectionFormatState | null>(() => {
    const owner = targetOwner();
    if (!owner) return null;
    selectionTick.value; // eslint-disable-line no-unused-expressions -- reactive touch
    let range: TextRange;
    if (editingOwner.value && host.value) {
      const dom = selectionRangeOf(host.value) ?? lastRange;
      range = dom ? { start: dom.start, end: dom.end } : { start: { para: 0, offset: 0 }, end: { para: 0, offset: 0 } };
    } else {
      range = wholeFrameRange(owner);
    }
    return selectionFormatState(owner, range, deckThemeOf(deps.model.value));
  });

  // --- session lifecycle ----------------------------------------------------
  function startEditing(frameId: string, selectAll = false, point?: CaretPoint, range?: TextRange): void {
    if (!sessionFrameById(frameId)) return;
    pendingRange = range ?? null;
    deps.selection.select(frameId);
    editingCell.value = null;
    editingFrameId.value = frameId;
    // registerHost (from the mounted contenteditable) finishes the setup; a
    // re-entry into the same frame just refreshes.
    if (host.value) {
      renderAndFocus(selectAll, point);
    } else {
      pendingSelectAll = selectAll;
      pendingPoint = point ?? null;
    }
  }

  function startEditingCell(frameId: string, row: number, col: number, point?: CaretPoint, range?: TextRange): void {
    const table = tableFrameById(frameId);
    const owner = table ? ownerCellAt(table, row, col) : null;
    if (!owner) return;
    pendingRange = range ?? null;
    deps.selection.select(frameId);
    editingCell.value = { row: owner.row, col: owner.col };
    activeCell.value = { row: owner.row, col: owner.col };
    // Re-register against the new cell's host: a running session in another cell
    // must drop its listeners before the host element is replaced.
    if (host.value) detachAndClearHost();
    editingFrameId.value = frameId;
    if (host.value) {
      renderAndFocus(false, point);
    } else {
      pendingSelectAll = false;
      pendingPoint = point ?? null;
    }
  }

  function focusCell(frameId: string, row: number, col: number): void {
    const table = tableFrameById(frameId);
    if (!table) return;
    const owner = ownerCellAt(table, row, col);
    if (!owner) return;
    activeCell.value = { row: owner.row, col: owner.col };
    // Move a live session onto the clicked cell; otherwise just anchor + select
    // (a single click never starts a session — AC-6a).
    if (editingFrameId.value === frameId && editingCell.value) {
      startEditingCell(frameId, owner.row, owner.col);
    } else {
      deps.selection.select(frameId);
    }
  }

  let pendingSelectAll = false;
  let pendingPoint: CaretPoint | null = null;
  let pendingRange: TextRange | null = null; // consumed by the next render

  let detachHost: (() => void) | null = null;
  function detachAndClearHost(): void {
    if (detachHost) {
      detachHost();
      detachHost = null;
    }
    host.value = null;
  }
  function registerHost(el: HTMLElement | null): void {
    if (detachHost) {
      detachHost();
      detachHost = null;
    }
    host.value = el;
    if (el && editingOwner.value) {
      const onSelectionChange = (): void => {
        // Only while the host is focused — a toolbar click moves focus off the
        // contenteditable and collapses its selection; ignoring that keeps the
        // pre-click selection available for the command (lastRange).
        if (el.ownerDocument.activeElement !== el) return;
        const range = selectionRangeOf(el);
        if (range) {
          lastRange = range;
          selectionTick.value++;
        }
      };
      el.addEventListener("input", onInput);
      el.addEventListener("keydown", onKeydown);
      el.addEventListener("paste", onPaste);
      el.addEventListener("copy", onCopy);
      el.addEventListener("cut", onCut);
      el.ownerDocument.addEventListener("selectionchange", onSelectionChange);
      detachHost = () => {
        el.removeEventListener("input", onInput);
        el.removeEventListener("keydown", onKeydown);
        el.removeEventListener("paste", onPaste);
        el.removeEventListener("copy", onCopy);
        el.removeEventListener("cut", onCut);
        el.ownerDocument.removeEventListener("selectionchange", onSelectionChange);
      };
      renderAndFocus(pendingSelectAll, pendingPoint ?? undefined);
      pendingSelectAll = false;
      pendingPoint = null;
    }
  }

  // The last selection observed inside the host — used when a toolbar click has
  // moved focus off the contenteditable (its live selection would be empty).
  let lastRange: { start: { para: number; offset: number }; end: { para: number; offset: number } } | null = null;

  function renderAndFocus(selectAll: boolean, point?: CaretPoint): void {
    const el = host.value;
    const owner = editingOwner.value;
    if (!el || !owner) return;
    renderHost(el, owner);
    el.focus();
    // A find hit's range is selected (SPEC-098); a double-click drops the
    // caret at the click point (AC-1); a select-all start (Enter-to-edit
    // selects nothing, but a keyboard entry may) or a point that lands
    // outside the host falls back to the end.
    const range = pendingRange;
    pendingRange = null;
    if (range) setSelectionRange(el, range.start, range.end);
    else if (selectAll) selectAllIn(el);
    else if (point && placeCaretAtPoint(el, point.x, point.y)) {
      /* caret placed at the click point */
    } else placeCaretAtEnd(el, owner);
    selectionTick.value++;
  }

  /** Renders the owner into the host in the deck's theme; a shape's
   * unaligned paragraphs centre there as they do on the stage. */
  function renderHost(el: HTMLElement, owner: ParagraphOwner): void {
    const frame = editingFrame.value;
    renderEditableParagraphs(el, owner.paragraphs, deckThemeOf(deps.model.value), frame ? defaultParagraphAlign(frame) : "left");
  }

  function stopEditing(): void {
    closeBurst();
    // A session that changed nothing has nothing to read back — its DOM is
    // the rendered model — so it ends without marking the slide (SPEC-098:
    // stepping through find hits is no mutation).
    if (host.value && editingOwner.value && sessionEdited) syncModelFromDom();
    sessionEdited = false;
    detachAndClearHost();
    editingFrameId.value = null;
    editingCell.value = null;
    lastRange = null;
    selectionTick.value++;
  }

  // --- normalise-after-input ------------------------------------------------
  function onInput(): void {
    if (suppressInput) return;
    noteTyping();
    syncModelFromDom();
  }

  function syncModelFromDom(notify = true): void {
    const el = host.value;
    const owner = editingOwner.value;
    const frame = editingFrame.value;
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!el || !owner || !frame || !slide || !model) return;
    // normalizeTextFrameDom recovers paragraph-level props (align/bullet/level)
    // per DOM block, so bullet/indent stay glued to their paragraph across a
    // keystroke that changes the paragraph count (Enter, Backspace-merge).
    owner.paragraphs = normalizeTextFrameDom(el);
    markSlideEdited(slide, model.format); // before bumpFrame so the clone carries dirty
    bumpFrame(frame);
    deps.commit();
    selectionTick.value++;
    if (notify) edited(frame);
  }

  function onPaste(event: ClipboardEvent): void {
    if (!host.value || !editingOwner.value) return;
    event.preventDefault();
    const text = event.clipboardData?.getData("text/plain") ?? "";
    if (text.length === 0) return;
    closeBurst();
    deps.record("Paste");
    suppressInput = true;
    insertPlainText(host.value, text);
    suppressInput = false;
    syncModelFromDom();
  }

  // Cut/copy put PLAIN text on the clipboard (AC-14).
  const onCopy = (event: ClipboardEvent): void => copyPlainText(event, host.value);
  function onCut(event: ClipboardEvent): void {
    if (!host.value || !editingOwner.value) return;
    onCopy(event);
    closeBurst();
    deps.record("Cut");
    suppressInput = true;
    host.value.ownerDocument.getSelection()?.deleteFromDocument();
    suppressInput = false;
    syncModelFromDom();
  }

  function onKeydown(event: KeyboardEvent): void {
    if (!editingOwner.value) return;
    if (event.key === "Escape") {
      event.preventDefault();
      stopEditing();
      return;
    }
    // Tab/Shift+Tab move between non-covered cells; Tab in the last cell appends
    // a row (AC-6b). Only a table-cell session traps Tab.
    if (event.key === "Tab" && editingCell.value) {
      event.preventDefault();
      tabCell(event.shiftKey ? -1 : 1);
      return;
    }
    // SPEC-097 AC-23: in a text session Tab demotes / Shift+Tab promotes (one
    // "Indent" entry; onGlobalKeydown's Tab frame-cycling skips while editing).
    if (event.key === "Tab" && !event.metaKey && !event.ctrlKey) {
      event.preventDefault();
      runCommand("Indent", (owner, range) => (event.shiftKey ? promote : demote)(owner, paragraphSpan(range)));
      return;
    }
    const mod = event.metaKey || event.ctrlKey;
    if (mod && !event.shiftKey && !event.altKey) {
      const key = event.key.toLowerCase();
      if (key === "b") {
        event.preventDefault();
        toggleBold();
      } else if (key === "i") {
        event.preventDefault();
        toggleItalic();
      } else if (key === "u") {
        event.preventDefault();
        toggleUnderline();
      } else if (key === "a") {
        event.preventDefault();
        if (host.value) selectAllIn(host.value);
        selectionTick.value++;
      } else if (key === "k" && linkShortcut) {
        event.preventDefault(); // SPEC-097 AC-46: the slide link dialog
        linkShortcut();
      }
    }
    // Arrow keys are deliberately NOT trapped: the browser moves the caret, so
    // it can leave the table at its edges (AC-6c).
  }

  /** Moves the cell session forward/back one non-covered cell, committing the
   * cell just left. Tab past the last cell appends a row (AC-6b). */
  function tabCell(dir: 1 | -1): void {
    const frame = tableFrameById(editingFrameId.value);
    const cell = editingCell.value;
    if (!frame || !cell) return;
    if (host.value && editingOwner.value) syncModelFromDom();
    const next = dir > 0 ? nextCellAddress(frame, cell.row, cell.col) : prevCellAddress(frame, cell.row, cell.col);
    if (next) {
      startEditingCell(frame.id, next.row, next.col);
      return;
    }
    if (dir < 0) return; // Shift+Tab in the first cell is a no-op.
    // Last cell: append a row and land in its first cell.
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model) return;
    closeBurst();
    deps.record("Insert row");
    insertTableRow(frame, frame.rows.length - 1, "below");
    markSlideEdited(slide, model.format);
    bumpFrame(frame);
    deps.commit();
    startEditingCell(frame.id, frame.rows.length - 1, 0);
  }

  // --- formatting commands --------------------------------------------------
  /** The model TextRange a command applies to: the DOM selection when editing,
   * else the whole selected frame (AC-25). */
  function commandRange(owner: ParagraphOwner): TextRange {
    if (editingOwner.value && host.value) {
      const live = selectionRangeOf(host.value);
      // Prefer a real (non-collapsed) selection; a toolbar click collapses the
      // live one, so fall back to the last selection captured while focused.
      const chosen = live && !isCollapsed(live) ? live : lastRange && !isCollapsed(lastRange) ? lastRange : live ?? lastRange;
      if (chosen) return { start: chosen.start, end: chosen.end };
    }
    return wholeFrameRange(owner);
  }

  const runCommand = (label: string, apply: TextCommand): void => void runTextCommand(label, apply);
  function runTextCommand(label: string, apply: TextCommand, override?: TextRange): boolean {
    const owner = targetOwner();
    const frame = targetFrame();
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!owner || !frame || !slide || !model) return false;
    closeBurst();
    const range = override ?? commandRange(owner);
    deps.record(label);
    apply(owner, range, frame);
    markSlideEdited(slide, model.format);
    if (editingOwner.value && host.value) {
      renderHost(host.value, owner);
      setSelectionRange(host.value, range.start, range.end);
    }
    bumpFrame(frame);
    deps.commit();
    selectionTick.value++;
    edited(frame);
    return true;
  }

  const runParagraphCommand = (label: string, apply: ParagraphCommand): boolean =>
    runTextCommand(label, (owner, range, frame) => apply(owner, paragraphSpan(range), frame));
  function commandTarget(): TextCommandTarget | null {
    const owner = targetOwner();
    const frame = targetFrame();
    return owner && frame ? { owner, frame, range: commandRange(owner) } : null;
  }
  let linkShortcut: (() => void) | null = null;
  const setLinkShortcut = (handler: (() => void) | null): void => void (linkShortcut = handler);
  const setAfterEdit = (handler: ((frameId: string) => void) | null): void => void (afterEdit = handler);

  function paragraphCommand(label: string, patch: ParagraphFormatPatch): void {
    runCommand(label, (owner, range) =>
      applyParagraphFormat(owner, { startPara: range.start.para, endPara: range.end.para }, patch),
    );
  }

  function toggleRunMark(label: string, key: "bold" | "italic" | "underline"): void {
    const state = toolbarState.value;
    const on = state ? state[key] !== true : true; // mixed/false → turn on
    const patch: RunFormatPatch = { [key]: on };
    runCommand(label, (owner, range) => applyRunFormat(owner, range, patch));
  }

  const toggleBold = (): void => toggleRunMark("Bold", "bold");
  const toggleItalic = (): void => toggleRunMark("Italic", "italic");
  const toggleUnderline = (): void => toggleRunMark("Underline", "underline");
  const setFontFamily = (family: string): void =>
    runCommand("Font", (owner, range) => applyRunFormat(owner, range, { fontFamily: family }));
  const setFontSize = (sizePt: number): void =>
    runCommand("Font size", (owner, range) => applyRunFormat(owner, range, { sizePt }));
  const setColor = (color: ModelColor): void =>
    runCommand("Text colour", (owner, range) => applyRunFormat(owner, range, { color }));
  const setAlign = (align: NonNullable<TextFrame["paragraphs"][number]["align"]>): void =>
    paragraphCommand("Align", { align });
  const toggleBulletCmd = (): void =>
    runCommand("Bullet", (owner, range) => toggleBullet(owner, { startPara: range.start.para, endPara: range.end.para }));
  // Same demote/promote as Tab (SPEC-097): a default marker follows its level.
  const indent = (delta: 1 | -1): void =>
    runCommand(delta > 0 ? "Indent" : "Outdent", (owner, range) => (delta > 0 ? demote : promote)(owner, paragraphSpan(range)));

  return {
    editingFrameId,
    editingCell,
    activeCell,
    toolbarEnabled,
    toolbarState,
    startEditing,
    startEditingCell,
    focusCell,
    stopEditing,
    registerHost,
    onInput,
    onKeydown,
    onPaste,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    setFontFamily,
    setFontSize,
    setColor,
    setAlign,
    toggleBulletCmd,
    indent,
    commandTarget,
    runTextCommand,
    runParagraphCommand,
    setLinkShortcut,
    setAfterEdit,
  };
}
