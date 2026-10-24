// The slide text session's plain helpers and shared types (SPEC-081, widened
// by SPEC-097), kept beside useSlideTextEditing.ts so the composable stays
// within its composableSize.spec.ts budget: the caret/cell address shapes, the
// command signatures useSlideTextDepth runs through `runTextCommand`, the
// frame re-identification every edit ends with, and the DOM helpers the
// session uses (caret-at-end, plain-text paste and copy).
import type { PresentationModel, ShapeFrame, Slide, SlideFrame, SlideParagraph } from "../../lib/presentation/model";
import type { ParagraphOwner, ParagraphRange, TextRange } from "../../lib/presentation/text";
import { placeCaret } from "../../components/presentation/slideTextDom";
import { defaultParagraphAlign } from "../../components/presentation/slideRender";

/** The viewport coordinates of a double-click, so the session drops the caret
 * at the click point rather than the end of the frame (AC-1). */
export interface CaretPoint {
  x: number;
  y: number;
}

/** A table cell address (SPEC-091). */
export interface CellAddress {
  row: number;
  col: number;
}

/** What a text command acts on (SPEC-097): the paragraph owner (a text frame
 * or a table cell), its frame and the character range. */
export interface TextCommandTarget {
  owner: ParagraphOwner;
  frame: SlideFrame;
  range: TextRange;
}

/** A command over a character range of a paragraph owner. */
export type TextCommand = (owner: ParagraphOwner, range: TextRange, frame: SlideFrame) => void;

/** A command over the paragraphs a range touches. */
export type ParagraphCommand = (owner: ParagraphOwner, range: ParagraphRange, frame: SlideFrame) => void;

/** The paragraphs a frame-level session edits: a text frame's, or (SPEC-098)
 * a shape's text — a find hit there opens a session on it, its match
 * selected in the same host. Null for any other frame. Only `startEditing`
 * reaches a shape: double-click and Enter still open text frames alone. */
export function frameTextOwner(frame: SlideFrame | null): (SlideFrame & ParagraphOwner) | null {
  return frame && (frame.kind === "text" || isShapeText(frame)) ? frame : null;
}
function isShapeText(frame: SlideFrame): frame is ShapeFrame & ParagraphOwner {
  return frame.kind === "shape" && Array.isArray(frame.paragraphs);
}

/** How a session's frame lays out a paragraph with no alignment of its own —
 * a shape's text centres (SPEC-079) — for the host and the toolbar state. */
export function unalignedAlign(frame: SlideFrame | null): NonNullable<SlideParagraph["align"]> {
  return frame ? defaultParagraphAlign(frame) : "left";
}

/** The inclusive paragraph span of a character range. */
export function paragraphSpan(range: TextRange): ParagraphRange {
  return {
    startPara: Math.min(range.start.para, range.end.para),
    endPara: Math.max(range.start.para, range.end.para),
  };
}

export function isCollapsed(range: TextRange): boolean {
  return range.start.para === range.end.para && range.start.offset === range.end.offset;
}

/** The slides model is a shallowRef and child views' computeds key off object
 * identity, so a nested `frame.paragraphs = …` is not seen by the read-only
 * renderers (the stage's other frames, the filmstrip thumbnail — SPEC-081
 * AC-43/44). Replaces BOTH the edited frame and its slide with fresh objects
 * (same ids) so every `:slide`/`:frame`-bound view re-renders. The editing
 * frame's contenteditable host is imperatively managed and keyed by id, so it
 * is not disturbed; the slide id is unchanged, so selection is not dropped. */
export function reidentifyEditedFrame(model: PresentationModel | null, current: Slide | null, frameId: string): void {
  if (!model || !current) return;
  const index = model.slides.indexOf(current);
  if (index >= 0) {
    model.slides[index] = { ...current, frames: current.frames.map((f) => (f.id === frameId ? { ...f } : f)) };
  }
}

/** A typing burst (SPEC-081): the first keystroke records ONE undo entry under
 * `label`, and a pause of `ms` — or `close()`, called by any non-typing
 * command or the end of the session — ends it. */
export function createTypingBurst(record: (label: string) => void, label: string, ms: number): { note: () => void; close: () => void } {
  let active = false;
  let timer = 0;
  function close(): void {
    active = false;
    if (timer) {
      clearTimeout(timer);
      timer = 0;
    }
  }
  function note(): void {
    if (!active) {
      record(label);
      active = true;
    }
    if (timer) clearTimeout(timer);
    timer = window.setTimeout(close, ms);
  }
  return { note, close };
}

/** Cut/copy put PLAIN text on the clipboard (SPEC-081 AC-14) — the styled DOM
 * is never written, so a paste elsewhere carries no slide formatting; the
 * empty-placeholder prompt lives outside the host, so it can never be copied. */
export function copyPlainText(event: ClipboardEvent, host: HTMLElement | null): void {
  const text = host?.ownerDocument.getSelection()?.toString() ?? "";
  if (!event.clipboardData) return;
  event.preventDefault();
  event.clipboardData.setData("text/plain", text);
}

export function placeCaretAtEnd(el: HTMLElement, owner: ParagraphOwner): void {
  const last = Math.max(0, owner.paragraphs.length - 1);
  const lastLen = owner.paragraphs[last]?.runs.reduce((n, r) => n + r.text.length, 0) ?? 0;
  placeCaret(el, { para: last, offset: lastLen });
}

/** Inserts plain text at the caret, newlines becoming paragraph breaks
 * (AC-13). Deletes any selected content first. */
export function insertPlainText(host: HTMLElement, text: string): void {
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
