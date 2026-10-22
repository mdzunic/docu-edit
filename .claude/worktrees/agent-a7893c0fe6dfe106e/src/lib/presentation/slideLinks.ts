// Slide hyperlinks (SPEC-097): which frames may carry one, and the commands
// that set/clear a run link or a whole-frame link. A link is two scalars —
// `linkHref` (an external URL, always `sanitizeHref`-clean) or `linkSlide`
// (a 1-based in-deck slide number), at most one set — plus `linkTooltip`, on
// a `SlideRun` and on a text/image/shape frame. A run link inside a linked
// frame wins for that run.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core — the URL safety helpers arrive through
// src/lib/core/urls.ts (SPEC-097's two-phase move out of src/lib/links.ts).
import { sanitizeHref } from "../core/urls";
import type { SlideFrame, SlideRun } from "./model";
import { applyRunFormat, type ParagraphOwner, type TextRange } from "./text";

/** A link target: exactly one of `href`/`slide` is set. */
export interface SlideLinkTarget {
  href?: string;
  slide?: number;
  tooltip?: string;
}

/** The frame kinds that serialise a frame link — text, image and shape (the
 * same shape as SPEC-086's isAnimatableKind). Group and foreign frames
 * re-emit verbatim, so they never carry one. */
export function isLinkableKind(kind: SlideFrame["kind"]): boolean {
  return kind === "text" || kind === "image" || kind === "shape";
}

/** A link target cleaned for the model: the href sanitised (a dangerous or
 * empty URL yields null), a slide number clamped to the deck, an empty
 * tooltip dropped. Null when nothing linkable remains. */
export function cleanLinkTarget(target: SlideLinkTarget, slideCount: number): SlideLinkTarget | null {
  const tooltip = target.tooltip?.trim() ? target.tooltip.trim() : undefined;
  if (target.slide !== undefined) {
    const slide = Math.round(target.slide);
    if (!Number.isFinite(slide) || slide < 1 || slide > Math.max(1, slideCount)) return null;
    return tooltip ? { slide, tooltip } : { slide };
  }
  const href = target.href !== undefined ? sanitizeHref(target.href) : null;
  if (!href) return null;
  return tooltip ? { href, tooltip } : { href };
}

/** Links (or, with `null`, unlinks) the characters of a range. Returns
 * whether anything changed. */
export function setRunLink(owner: ParagraphOwner, range: TextRange, target: SlideLinkTarget | null): boolean {
  if (!target) {
    return applyRunFormat(owner, range, { linkHref: null, linkSlide: null, linkTooltip: null });
  }
  return applyRunFormat(owner, range, {
    linkHref: target.href ?? null,
    linkSlide: target.slide ?? null,
    linkTooltip: target.tooltip ?? null,
  });
}

/** Links (or unlinks) a whole frame. A non-linkable kind is left untouched
 * (returns false). */
export function setFrameLink(frame: SlideFrame, target: SlideLinkTarget | null): boolean {
  if (!isLinkableKind(frame.kind)) return false;
  delete frame.linkHref;
  delete frame.linkSlide;
  delete frame.linkTooltip;
  if (target?.href) frame.linkHref = target.href;
  else if (target?.slide) frame.linkSlide = target.slide;
  if (target && (target.href || target.slide) && target.tooltip) frame.linkTooltip = target.tooltip;
  return true;
}

/** The link a run renders under: its own, else its frame's (a run link wins
 * for that run). */
export function effectiveRunLink(run: SlideRun, frame?: Pick<SlideFrame, "linkHref" | "linkSlide" | "linkTooltip">): SlideLinkTarget | null {
  if (run.linkHref) return { href: run.linkHref, tooltip: run.linkTooltip };
  if (run.linkSlide) return { slide: run.linkSlide, tooltip: run.linkTooltip };
  if (frame?.linkHref) return { href: frame.linkHref, tooltip: frame.linkTooltip };
  if (frame?.linkSlide) return { slide: frame.linkSlide, tooltip: frame.linkTooltip };
  return null;
}

/** The link a run or frame carries, as a target (for the dialog's prefill). */
export function linkOf(item: Pick<SlideRun, "linkHref" | "linkSlide" | "linkTooltip">): SlideLinkTarget | null {
  if (item.linkHref) return { href: item.linkHref, tooltip: item.linkTooltip };
  if (item.linkSlide) return { slide: item.linkSlide, tooltip: item.linkTooltip };
  return null;
}
