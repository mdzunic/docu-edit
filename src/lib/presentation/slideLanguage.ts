// Language on slide runs (SPEC-098 — the normative source for slide-run
// language). A run carries its source tag verbatim in `SlideRun.lang`; the deck
// default `PresentationModel.language` stands in for the text model's
// `DefaultStyles.language`, so a run resolves its checking language through
// core's one `effectiveLanguage` chain. The deck default is derived at read
// and never serialised — see the spec's wire section for why.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import { effectiveLanguage } from "../core/language";
import type { PresentationModel, Slide, SlideFrame, SlideParagraph, SlideRun } from "./model";

function frameParagraphs(frame: SlideFrame): SlideParagraph[][] {
  switch (frame.kind) {
    case "text":
      return [frame.paragraphs];
    case "shape":
      return frame.paragraphs ? [frame.paragraphs] : [];
    case "table":
      return frame.rows.flatMap((row) => row.cells.map((cell) => cell.paragraphs));
    case "group":
      return frame.children.flatMap(frameParagraphs);
    default:
      return [];
  }
}

/** The deck's checking default as read: the most frequent non-empty
 * `SlideRun.lang` over every frame's runs (ties go to the tag met first), or
 * undefined when no run declares one. */
export function deckLanguageOf(slides: Slide[]): string | undefined {
  const counts = new Map<string, number>();
  for (const slide of slides) {
    for (const paragraphs of slide.frames.flatMap(frameParagraphs)) {
      for (const para of paragraphs) {
        for (const run of para.runs) {
          if (run.lang) counts.set(run.lang, (counts.get(run.lang) ?? 0) + 1);
        }
      }
    }
  }
  let best: string | undefined;
  let bestCount = 0;
  for (const [tag, count] of counts) {
    if (count > bestCount) {
      best = tag;
      bestCount = count;
    }
  }
  return best;
}

/** The language a slide run is checked in: its own tag, else the deck's,
 * else `DEFAULT_LANGUAGE` (core's chain). */
export function runLanguage(run: SlideRun, model: Pick<PresentationModel, "language">): string {
  return effectiveLanguage(run.lang, model.language);
}
