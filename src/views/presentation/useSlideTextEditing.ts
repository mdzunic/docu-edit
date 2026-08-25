// The slide text-editing session (SPEC-081): double-click (or Enter) starts an
// in-place contenteditable session; typing/Enter/Backspace/Delete/caret moves
// are the browser's, and after every input `normalizeTextFrameDom` walks the
// host back into the frame's `paragraphs` (normalise-after-input, so the DOM
// never diverges). Formatting NEVER round-trips through execCommand — a toolbar
// or keyboard command edits the model (src/lib/presentation/text.ts) and
// re-renders the host. Typing coalesces into one undo entry per burst (a pause
// ≥ 1s, ending the session, or any non-typing command closes it); formatting is
// one entry each. Every edit marks the slide dirty and persists.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, TextFrame } from "../../lib/presentation/model";
import {
  applyParagraphFormat,
  applyRunFormat,
  markSlideEdited,
  selectionFormatState,
  setIndentLevel,
  toggleBullet,
  wholeFrameRange,
  type ParagraphFormatPatch,
  type RunFormatPatch,
  type SelectionFormatState,
  type TextRange,
} from "../../lib/presentation/text";
import {
  normalizeTextFrameDom,
  placeCaret,
  renderEditableParagraphs,
  selectAllIn,
  selectionRangeOf,
  setSelectionRange,
} from "../../components/presentation/slideTextDom";
import type { FrameSelection } from "./useFrameSelection";

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
  /** Whether the toolbar's formatting controls act on something (an editing
   * session, or a selected text frame). */
  toolbarEnabled: ComputedRef<boolean>;
  /** The current selection's formatting for the toolbar (null when nothing is
   * targeted). */
  toolbarState: Ref<SelectionFormatState | null>;
  startEditing: (frameId: string, selectAll?: boolean) => void;
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
  setColor: (color: string) => void;
  setAlign: (align: NonNullable<TextFrame["paragraphs"][number]["align"]>) => void;
  toggleBulletCmd: () => void;
  indent: (delta: 1 | -1) => void;
}

export function useSlideTextEditing(deps: SlideTextEditingDeps): SlideTextEditing {
  const editingFrameId = ref<string | null>(null);
  const host = ref<HTMLElement | null>(null);
  // Bumped when the DOM caret/selection moves (no model change) so the toolbar
  // state recomputes; model edits recompute it through the reactive frame.
  const selectionTick = ref(0);
  let suppressInput = false;

  // --- undo-burst state -----------------------------------------------------
  let burstActive = false;
  let burstTimer = 0;
  function noteTyping(): void {
    if (!burstActive) {
      deps.record("Typing");
      burstActive = true;
    }
    if (burstTimer) clearTimeout(burstTimer);
    burstTimer = window.setTimeout(closeBurst, TYPING_BURST_MS);
  }
  function closeBurst(): void {
    burstActive = false;
    if (burstTimer) {
      clearTimeout(burstTimer);
      burstTimer = 0;
    }
  }

  // The slides model is a shallowRef and child views' computeds key off object
  // identity, so a nested `frame.paragraphs = …` is not seen by the read-only
  // renderers (the stage's other frames, the filmstrip thumbnail — AC-43/44).
  // Replace BOTH the edited frame and its slide with fresh objects (same ids)
  // so every `:slide`/`:frame`-bound view re-renders. The editing frame's
  // contenteditable host is imperatively managed and keyed by id, so it is not
  // disturbed; the slide id is unchanged, so selection is not dropped.
  function bumpFrame(frame: TextFrame): void {
    const model = deps.model.value;
    const current = deps.currentSlide.value;
    if (!model || !current) return;
    const frames = current.frames.map((f) => (f.id === frame.id ? { ...f } : f));
    const index = model.slides.indexOf(current);
    if (index >= 0) {
      model.slides[index] = { ...current, frames };
    }
  }

  // --- frame lookup ---------------------------------------------------------
  function textFrameById(frameId: string | null): TextFrame | null {
    const slide = deps.currentSlide.value;
    if (!slide || frameId === null) return null;
    const frame = slide.frames.find((f) => f.id === frameId);
    return frame && frame.kind === "text" ? frame : null;
  }
  const editingFrame = computed(() => textFrameById(editingFrameId.value));
  /** The frame a command targets: the one being edited, else a selected text
   * frame (AC-25). */
  function targetFrame(): TextFrame | null {
    if (editingFrame.value) return editingFrame.value;
    const selected = deps.selection.selectedFrame.value;
    return selected && selected.kind === "text" ? selected : null;
  }

  const toolbarEnabled = computed(() => targetFrame() !== null);

  // Reactive over the target frame's model (via commit's triggerRef) AND the
  // DOM selection (via selectionTick), so it survives undo/redo and caret moves.
  const toolbarState = computed<SelectionFormatState | null>(() => {
    const frame = targetFrame();
    if (!frame) return null;
    selectionTick.value; // eslint-disable-line no-unused-expressions -- reactive touch
    let range: TextRange;
    if (editingFrame.value && host.value) {
      const dom = selectionRangeOf(host.value) ?? lastRange;
      range = dom ? { start: dom.start, end: dom.end } : { start: { para: 0, offset: 0 }, end: { para: 0, offset: 0 } };
    } else {
      range = wholeFrameRange(frame);
    }
    return selectionFormatState(frame, range);
  });

  // --- session lifecycle ----------------------------------------------------
  function startEditing(frameId: string, selectAll = false): void {
    if (!textFrameById(frameId)) return;
    deps.selection.select(frameId);
    editingFrameId.value = frameId;
    // registerHost (from the mounted contenteditable) finishes the setup; a
    // re-entry into the same frame just refreshes.
    if (host.value) {
      renderAndFocus(selectAll);
    } else {
      pendingSelectAll = selectAll;
    }
  }

  let pendingSelectAll = false;

  let detachHost: (() => void) | null = null;
  function registerHost(el: HTMLElement | null): void {
    if (detachHost) {
      detachHost();
      detachHost = null;
    }
    host.value = el;
    if (el && editingFrame.value) {
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
      renderAndFocus(pendingSelectAll);
      pendingSelectAll = false;
    }
  }

  // The last selection observed inside the host — used when a toolbar click has
  // moved focus off the contenteditable (its live selection would be empty).
  let lastRange: { start: { para: number; offset: number }; end: { para: number; offset: number } } | null = null;

  function renderAndFocus(selectAll: boolean): void {
    const el = host.value;
    const frame = editingFrame.value;
    if (!el || !frame) return;
    renderEditableParagraphs(el, frame.paragraphs);
    el.focus();
    if (selectAll) selectAllIn(el);
    else placeCaretAtEnd(el, frame);
    selectionTick.value++;
  }

  function placeCaretAtEnd(el: HTMLElement, frame: TextFrame): void {
    const last = Math.max(0, frame.paragraphs.length - 1);
    const lastLen = frame.paragraphs[last]?.runs.reduce((n, r) => n + r.text.length, 0) ?? 0;
    placeCaret(el, { para: last, offset: lastLen });
  }

  function stopEditing(): void {
    closeBurst();
    if (host.value && editingFrame.value) {
      syncModelFromDom();
    }
    if (detachHost) {
      detachHost();
      detachHost = null;
    }
    editingFrameId.value = null;
    host.value = null;
    lastRange = null;
    selectionTick.value++;
  }

  // --- normalise-after-input ------------------------------------------------
  function onInput(): void {
    if (suppressInput) return;
    noteTyping();
    syncModelFromDom();
  }

  function syncModelFromDom(): void {
    const el = host.value;
    const frame = editingFrame.value;
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!el || !frame || !slide || !model) return;
    const next = normalizeTextFrameDom(el);
    mergeParagraphProps(next, frame.paragraphs);
    frame.paragraphs = next;
    markSlideEdited(slide, model.format); // before bumpFrame so the clone carries dirty
    bumpFrame(frame);
    deps.commit();
    selectionTick.value++;
  }

  function onPaste(event: ClipboardEvent): void {
    if (!host.value || !editingFrame.value) return;
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

  /** Cut/copy put PLAIN text on the clipboard (AC-14) — the styled DOM is never
   * written, so a paste elsewhere carries no slide formatting; the empty-
   * placeholder prompt lives outside the host, so it can never be copied. */
  function onCopy(event: ClipboardEvent): void {
    const text = host.value?.ownerDocument.getSelection()?.toString() ?? "";
    if (!event.clipboardData) return;
    event.preventDefault();
    event.clipboardData.setData("text/plain", text);
  }
  function onCut(event: ClipboardEvent): void {
    if (!host.value || !editingFrame.value) return;
    onCopy(event);
    closeBurst();
    deps.record("Cut");
    suppressInput = true;
    host.value.ownerDocument.getSelection()?.deleteFromDocument();
    suppressInput = false;
    syncModelFromDom();
  }

  function onKeydown(event: KeyboardEvent): void {
    if (!editingFrame.value) return;
    if (event.key === "Escape") {
      event.preventDefault();
      stopEditing();
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
      }
    }
  }

  // --- formatting commands --------------------------------------------------
  /** The model TextRange a command applies to: the DOM selection when editing,
   * else the whole selected frame (AC-25). */
  function commandRange(frame: TextFrame): TextRange {
    if (editingFrame.value && host.value) {
      const live = selectionRangeOf(host.value);
      // Prefer a real (non-collapsed) selection; a toolbar click collapses the
      // live one, so fall back to the last selection captured while focused.
      const chosen = live && !isCollapsed(live) ? live : lastRange && !isCollapsed(lastRange) ? lastRange : live ?? lastRange;
      if (chosen) return { start: chosen.start, end: chosen.end };
    }
    return wholeFrameRange(frame);
  }

  function runCommand(label: string, apply: (frame: TextFrame, range: TextRange) => void): void {
    const frame = targetFrame();
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!frame || !slide || !model) return;
    closeBurst();
    const range = commandRange(frame);
    deps.record(label);
    apply(frame, range);
    markSlideEdited(slide, model.format);
    if (editingFrame.value && host.value) {
      renderEditableParagraphs(host.value, frame.paragraphs);
      setSelectionRange(host.value, range.start, range.end);
    }
    bumpFrame(frame);
    deps.commit();
    selectionTick.value++;
  }

  function paragraphCommand(label: string, patch: ParagraphFormatPatch): void {
    runCommand(label, (frame, range) =>
      applyParagraphFormat(frame, { startPara: range.start.para, endPara: range.end.para }, patch),
    );
  }

  function toggleRunMark(label: string, key: "bold" | "italic" | "underline"): void {
    const state = toolbarState.value;
    const on = state ? state[key] !== true : true; // mixed/false → turn on
    const patch: RunFormatPatch = { [key]: on };
    runCommand(label, (frame, range) => applyRunFormat(frame, range, patch));
  }

  const toggleBold = (): void => toggleRunMark("Bold", "bold");
  const toggleItalic = (): void => toggleRunMark("Italic", "italic");
  const toggleUnderline = (): void => toggleRunMark("Underline", "underline");
  const setFontFamily = (family: string): void =>
    runCommand("Font", (frame, range) => applyRunFormat(frame, range, { fontFamily: family }));
  const setFontSize = (sizePt: number): void =>
    runCommand("Font size", (frame, range) => applyRunFormat(frame, range, { sizePt }));
  const setColor = (color: string): void =>
    runCommand("Text colour", (frame, range) => applyRunFormat(frame, range, { color }));
  const setAlign = (align: NonNullable<TextFrame["paragraphs"][number]["align"]>): void =>
    paragraphCommand("Align", { align });
  const toggleBulletCmd = (): void =>
    runCommand("Bullet", (frame, range) => toggleBullet(frame, { startPara: range.start.para, endPara: range.end.para }));
  const indent = (delta: 1 | -1): void =>
    runCommand(delta > 0 ? "Indent" : "Outdent", (frame, range) =>
      setIndentLevel(frame, { startPara: range.start.para, endPara: range.end.para }, delta),
    );

  return {
    editingFrameId,
    toolbarEnabled,
    toolbarState,
    startEditing,
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
  };
}

function isCollapsed(range: { start: { para: number; offset: number }; end: { para: number; offset: number } }): boolean {
  return range.start.para === range.end.para && range.start.offset === range.end.offset;
}

type SlideParagraphList = TextFrame["paragraphs"];

/** Carries a frame's paragraph-level properties (align/bullet/level) — which
 * typing does not touch — across a DOM normalise, matching by index. */
function mergeParagraphProps(next: SlideParagraphList, prev: SlideParagraphList): void {
  for (let i = 0; i < next.length; i++) {
    const before = prev[i];
    if (!before) continue;
    if (before.align) next[i].align = before.align;
    if (before.bullet) next[i].bullet = before.bullet;
    if (before.level) next[i].level = before.level;
  }
}

/** Inserts plain text at the caret, newlines becoming paragraph breaks
 * (AC-13). Deletes any selected content first. */
function insertPlainText(host: HTMLElement, text: string): void {
  const doc = host.ownerDocument;
  const sel = doc.getSelection();
  if (!sel || sel.rangeCount === 0) return;
  let range = sel.getRangeAt(0);
  range.deleteContents();
  const parts = text.replace(/\r\n?/g, "\n").split("\n");
  for (let i = 0; i < parts.length; i++) {
    if (i > 0) {
      const br = doc.createElement("br");
      range.insertNode(br);
      range.setStartAfter(br);
    }
    if (parts[i].length > 0) {
      const node = doc.createTextNode(parts[i]);
      range.insertNode(node);
      range.setStartAfter(node);
    }
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    range = sel.getRangeAt(0);
  }
}
