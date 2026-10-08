// The slide text session's plain helpers and shared types (SPEC-081, widened
// by SPEC-097), kept beside useSlideTextEditing.ts so the composable stays
// within its composableSize.spec.ts budget: the caret/cell address shapes, the
// command signatures useSlideTextDepth runs through `runTextCommand`, and the
// DOM helpers the session uses (caret-at-end, plain-text paste).
import type { SlideFrame } from "../../lib/presentation/model";
import type { ParagraphOwner, ParagraphRange, TextRange } from "../../lib/presentation/text";
import { placeCaret } from "../../components/presentation/slideTextDom";

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
