// Pure, framework-free derivations over the SPEC-079 slide model. The
// preparation milestone surfaced a per-slide text outline (title + lines);
// SPEC-079 keeps that observable — the library's slide count and the round
// trip's outline assertions — as a DERIVED view of the real frame model
// rather than a stored shape, so there is one source of truth. Everything
// here is a pure function of a `Slide`: no Vue, no store, no I/O.
import type { Slide, SlideFrame, SlideParagraph } from "./model";

/** The visible text of one paragraph: its runs concatenated in order. */
export function paragraphText(paragraph: SlideParagraph): string {
  return paragraph.runs.map((run) => run.text).join("");
}

/** Every paragraph line a frame contributes to the outline, in order —
 * descending into groups (grouped text is still slide text). Non-text frames
 * (image, foreign) contribute nothing. Empty lines are dropped. */
function frameLines(frame: SlideFrame): string[] {
  if (frame.kind === "text") {
    return frame.paragraphs.map(paragraphText).filter((line) => line.trim().length > 0);
  }
  if (frame.kind === "shape") {
    return (frame.paragraphs ?? []).map(paragraphText).filter((line) => line.trim().length > 0);
  }
  if (frame.kind === "group") {
    return frame.children.flatMap(frameLines);
  }
  return [];
}

/** A slide's read-only outline: the title (the first `role:"title"` frame's
 * joined text) and every other frame's paragraph lines in order. This is the
 * preparation milestone's `SlidePreview`, re-expressed as a derivation so the
 * library and the round-trip suite keep one source of truth (AC-4). */
export function slideOutline(slide: Slide): { title?: string; lines: string[] } {
  const titleFrame = slide.frames.find((frame) => frame.kind === "text" && frame.role === "title");
  const title = titleFrame ? frameLines(titleFrame).join(" ") : "";
  const lines: string[] = [];
  for (const frame of slide.frames) {
    if (frame === titleFrame) {
      continue;
    }
    lines.push(...frameLines(frame));
  }
  const outline: { title?: string; lines: string[] } = { lines };
  if (title.length > 0) {
    outline.title = title;
  }
  return outline;
}

/** The frame kinds of a slide in z-order — the cross-format structural
 * assertion (AC-6): both readers must produce the same kinds in the same
 * order for the twin fixtures. */
export function slideFrameKinds(slide: Slide): SlideFrame["kind"][] {
  return slide.frames.map((frame) => frame.kind);
}
