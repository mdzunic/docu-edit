// Formatting-marks segmentation (SPEC-028 AC-13..15): splits run text into
// plain-text, space and tab segments so the renderer can wrap the
// whitespace characters in marked spans while keeping the DOM text
// byte-identical to the model (the marks themselves are CSS
// pseudo-elements, never text — the canvas's model/DOM equality check and
// the caret's offset arithmetic both depend on that).

export interface MarkSegment {
  text: string;
  /** "text" renders bare; "space"/"tab" get a marked span. */
  kind: "text" | "space" | "tab";
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
