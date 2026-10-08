// Slide list markers and indents (SPEC-097), DERIVED, never stored: every
// rendering surface — the stage (slideRender.ts#markedParagraphs), filmstrip
// thumbnails, the play view, the build-preview layer and SPEC-088's HTML
// export — reads a paragraph list's markers through `paragraphMarkers` here,
// so they can never disagree. Built over src/lib/core/lists.ts (markerText,
// formatOrdinal, clampListLevel), which was written product-neutral for
// exactly this ("a Slides product numbers the list items inside a text box
// with the same marker math"). Slide level L is core level L + 1.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import {
  clampListLevel,
  formatOrdinal,
  LIST_INDENT_STEP_CM,
  LIST_MARKER_WIDTH_CM,
  markerText,
  type ListMarkerStyle,
} from "../core/lists";
import type { SlideParagraph } from "./model";
import { paragraphMarkerStyle } from "./text";

/** One paragraph's rendered marker. `text` is the glyph/ordinal string (""
 * for a non-list paragraph or a picture bullet); `imageId` names the picture
 * bullet's `model.media` part. */
export interface ParagraphMarker {
  text: string;
  imageId?: string;
}

/** The core list level (1..9) of a slide paragraph level (0..8). */
export function coreListLevel(level: number | undefined): number {
  return clampListLevel((level ?? 0) + 1);
}

/** The marker of every paragraph in a text body, counting numbered items per
 * level by adjacency: a non-list paragraph restarts every level, a shallower
 * item restarts each deeper level, a different item kind or number style at
 * the same level restarts that level, and `startAt` seeds a run's first
 * number (later items in the run continue it). */
export function paragraphMarkers(paragraphs: SlideParagraph[]): ParagraphMarker[] {
  // Per core level: the running style and the last number issued.
  const runs = new Map<number, { style: ListMarkerStyle; value: number }>();
  return paragraphs.map((para) => {
    const style = paragraphMarkerStyle(para);
    if (!style) {
      runs.clear();
      return { text: "" };
    }
    const level = coreListLevel(para.level);
    for (const key of Array.from(runs.keys())) {
      if (key > level) runs.delete(key);
    }
    if (para.bulletImageId) {
      runs.delete(level);
      return { text: "", imageId: para.bulletImageId };
    }
    if (para.bullet !== "number") {
      runs.delete(level);
      return { text: para.bulletChar ?? markerText(style, 1) };
    }
    const current = runs.get(level);
    const value = current && current.style === style ? current.value + 1 : (para.startAt ?? 1);
    runs.set(level, { style, value });
    return { text: markerText(style, value) };
  });
}

/** The ordinal a numbered marker would show, formatted without its "."
 * suffix — the list popover previews styles through `markerText(style, 1)`;
 * this is its suffix-free sibling for compact labels. */
export function ordinalLabel(style: ListMarkerStyle, ordinal: number): string {
  return formatOrdinal(style, ordinal);
}

/** Whether a paragraph is a list item. */
export function isListParagraph(para: SlideParagraph): boolean {
  return para.bullet !== undefined;
}

/** A paragraph's text edge in cm: its own `indentCm`, else
 * `(level + 1) × LIST_INDENT_STEP_CM` for a list item and
 * `level × LIST_INDENT_STEP_CM` for plain text (so an unindented title stays
 * flush left). */
export function effectiveIndentCm(para: SlideParagraph): number {
  if (para.indentCm !== undefined) return para.indentCm;
  const level = para.level ?? 0;
  return (isListParagraph(para) ? level + 1 : level) * LIST_INDENT_STEP_CM;
}

/** A paragraph's hanging marker box in cm: its own `hangingCm`, else
 * `LIST_MARKER_WIDTH_CM` for a list item and 0 for plain text. */
export function effectiveHangingCm(para: SlideParagraph): number {
  if (para.hangingCm !== undefined) return para.hangingCm;
  return isListParagraph(para) ? LIST_MARKER_WIDTH_CM : 0;
}
