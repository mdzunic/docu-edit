// Formatting-marks segmentation (SPEC-028 AC-13..15): splits run text into
// plain-text, space and tab segments so the renderer can wrap the
// whitespace characters in marked spans while keeping the DOM text
// byte-identical to the model (the marks themselves are CSS
// pseudo-elements, never text — the canvas's model/DOM equality check and
// the caret's offset arithmetic both depend on that).

import type { TabLeader } from "../documentModel";

export interface MarkSegment {
  text: string;
  /** "text" renders bare; "space"/"tab" get a marked span. */
  kind: "text" | "space" | "tab";
  /** SPEC-054: a tab segment's resolved box width in px (the advance it fills
   * to reach its stop) — set only on "tab" segments the renderer sizes. */
  widthPx?: number;
  /** SPEC-054: the leader the tab's box is filled with (dots/hyphens/an
   * underscore), if its target stop carried one. */
  leader?: TabLeader;
}

/** Splits text into mark segments. Consecutive spaces split one segment per
 * space so each renders its own middle-dot mark; other characters
 * (including "\n" soft breaks) stay in plain segments. */
export function markSegments(text: string): MarkSegment[] {
  const segments: MarkSegment[] = [];
  let plain = "";
  for (const char of text) {
    if (char === " " || char === " " || char === "\t") {
      if (plain) {
        segments.push({ text: plain, kind: "text" });
        plain = "";
      }
      segments.push({ text: char, kind: char === "\t" ? "tab" : "space" });
    } else {
      plain += char;
    }
  }
  if (plain) {
    segments.push({ text: plain, kind: "text" });
  }
  return segments;
}

/** Splits text at tabs only (SPEC-054), leaving other characters — including
 * spaces — in plain segments. Used when formatting marks are OFF but tabs
 * still need their own sized boxes; the tab keeps its real "\t" text so the
 * caret arithmetic counts it exactly as before. */
export function tabOnlySegments(text: string): MarkSegment[] {
  const segments: MarkSegment[] = [];
  let plain = "";
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === "\t") {
      if (plain) {
        segments.push({ text: plain, kind: "text" });
        plain = "";
      }
      segments.push({ text: "\t", kind: "tab" });
    } else {
      plain += text[i];
    }
  }
  if (plain) {
    segments.push({ text: plain, kind: "text" });
  }
  return segments;
}
