// Plain helpers, types and messages for SPEC-097's text-depth composables
// (useSlideTextDepth, useSlideTextBody, useSlideLinks) — framework-free, so
// they stay out of the composables' size budgets and test without Vue.
import type { SlideParagraph } from "../../lib/presentation/model";
import type { FieldState, ParagraphOwner, ParagraphRange, TextRange } from "../../lib/presentation/text";
import type { TextBodyFrame } from "../../lib/presentation/textBody";
import { linkOf, type SlideLinkTarget } from "../../lib/presentation/slideLinks";
import { frameTextHost, measureFrameFit, type FitScale } from "../../components/presentation/slideTextFit";
import { isCollapsed } from "./slideTextSession";

/** What a rendered text body measures to (slideTextFit.ts#measureFrameFit). */
export interface FrameFit {
  scale: FitScale;
  contentHeightCm: number;
}

/** A frame measurer: the stage DOM in the app, a fake in unit tests. */
export type FrameMeasurer = (frame: TextBodyFrame) => FrameFit | null;

/** The open link dialog: what it links (a character range of the edited
 * owner, or a whole frame) and the link it starts from. */
export interface SlideLinkDialogState {
  kind: "run" | "frame";
  frameId: string;
  /** The character range a run link covers (captured before the dialog took
   * focus); null for a frame link. */
  range: TextRange | null;
  /** The existing link under the range / on the frame, for prefill. */
  initial: SlideLinkTarget | null;
}

/** One `link-slide` option: a 1-based slide number labelled by its title. */
export interface SlideLinkChoice {
  number: number;
  label: string;
}

/** The custom-bullet message when a glyph is not exactly one grapheme. */
export const BULLET_CHAR_ERROR = "Choose a single character for the bullet.";
/** The picture-bullet message for a non-image file. */
export const BULLET_IMAGE_TYPE_ERROR = "That file isn't a supported image (PNG, JPEG, GIF or WebP).";
/** The picture-bullet message for an unreadable file. */
export const BULLET_IMAGE_READ_ERROR = "That image couldn't be read.";
/** The link dialog's message for a dangerous or empty address. */
export const LINK_URL_ERROR = "This address type isn't allowed in links";
/** The link dialog's message for a slide outside the deck. */
export const LINK_SLIDE_ERROR = "Choose a slide in this deck";

/** The bubbling event the stage dispatches on Cmd/Ctrl+click of an in-deck
 * slide link (`detail.slide` is 1-based). */
export const SLIDE_LINK_FOLLOW_EVENT = "slide-link-follow";

/** The DOM measurer: the frame's rendered body on the editor stage, or null
 * when it is not rendered (jsdom, another slide). Only the interactive
 * stage's top-level frames carry `data-frame-id` (filmstrip thumbnails and the
 * build preview do not — each thumbnail is its own `slide-stage`, so that test
 * id must not scope this), and the stage workspace narrows it further. */
export function stageFit(frame: TextBodyFrame): FrameFit | null {
  // jsdom has neither layout nor `CSS.escape` (frameTextHost's selector).
  if (typeof document === "undefined" || typeof CSS === "undefined") return null;
  const workspace = document.querySelector('[data-testid="stage-workspace"]') ?? document;
  return measureFrameFit(frame, frameTextHost(frame.id, workspace));
}

/** A selection field as a control value: the number, or "" when mixed/absent. */
export function numberOrBlank(value: FieldState<number>): number | "" {
  return typeof value === "number" ? value : "";
}

/** One value folded over a paragraph span: shared (r3), or "" when mixed. */
export function foldParagraphs(owner: ParagraphOwner, range: ParagraphRange, read: (para: SlideParagraph) => number): number | "" {
  let value: number | undefined;
  for (let p = range.startPara; p <= range.endPara; p++) {
    const para = owner.paragraphs[p];
    if (!para) continue;
    const next = read(para);
    if (value === undefined) value = next;
    else if (Math.abs(value - next) > 1e-9) return "";
  }
  return value === undefined ? "" : Math.round(value * 1000) / 1000;
}

/** The run link at a range's start, for the dialog's prefill. */
export function runLinkAt(owner: ParagraphOwner, range: TextRange): SlideLinkTarget | null {
  const para = owner.paragraphs[range.start.para];
  if (!para) return null;
  let offset = 0;
  for (const run of para.runs) {
    const end = offset + run.text.length;
    if (range.start.offset < end || (range.start.offset === end && run === para.runs[para.runs.length - 1])) {
      return linkOf(run);
    }
    offset = end;
  }
  return null;
}

function paragraphText(owner: ParagraphOwner, para: number): string {
  return owner.paragraphs[para]?.runs.map((run) => run.text).join("") ?? "";
}

function isWord(ch: string | undefined): boolean {
  return ch !== undefined && /\S/.test(ch);
}

/** A collapsed caret widens to the word it sits in (so Cmd/Ctrl+K on a caret
 * links that word); a real selection is kept as is. */
export function wordRange(owner: ParagraphOwner, range: TextRange): TextRange {
  if (!isCollapsed(range)) return range;
  const text = paragraphText(owner, range.start.para);
  let start = range.start.offset;
  let end = range.start.offset;
  while (start > 0 && isWord(text[start - 1])) start--;
  while (end < text.length && isWord(text[end])) end++;
  return { start: { para: range.start.para, offset: start }, end: { para: range.start.para, offset: end } };
}

/** A collapsed caret widens to its whole paragraph (letter spacing on a caret
 * spaces the paragraph rather than doing nothing). */
export function paragraphsOf(owner: ParagraphOwner, range: TextRange): TextRange {
  if (!isCollapsed(range)) return range;
  const length = paragraphText(owner, range.start.para).length;
  return { start: { para: range.start.para, offset: 0 }, end: { para: range.start.para, offset: length } };
}
