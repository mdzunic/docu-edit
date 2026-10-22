// Tab-stop layout (SPEC-054): resolves where each U+0009 in a paragraph
// advances to and how wide its box is, given the paragraph's effective tab
// stops and the document default interval. Framework-free and measurer-
// agnostic — the caller supplies a `measure(from, to)` that returns the pixel
// width of a text slice in the paragraph's own fonts, so the exact same
// resolution drives both the canvas wrap predictor (textMetrics.ts) and the
// DOM renderer (usePaginationView.ts). Keeping them on one function is what
// makes a tabbed line measure and render identically.
//
// All positions are pixels measured from the content box's start edge (the
// margin edge after any start indent). A tab stop's `positionCm` is already
// start-relative, so no indent term enters here; the browser's margin/RTL
// mirroring places the whole box, and the wrap predictor subtracts the same
// indent from its available width.
import { PX_PER_CM } from "./core/units";
import type { TabLeader, TabStop, TabStopAlignment } from "./documentModel";

/** A resolved tab stop in pixels — what layoutParagraphTabs consumes. */
export interface TabStopPx {
  positionPx: number;
  alignment: TabStopAlignment;
  leader?: TabLeader;
}

/** One U+0009's resolved advance: the width of its own inline box, plus the
 * leader (if the target stop carried one) to fill that box with. */
export interface TabAdvance {
  widthPx: number;
  leader?: TabLeader;
}

/** A paragraph's resolved tab geometry. */
export interface ParagraphTabLayout {
  /** Keyed by the character index of each U+0009 in the paragraph text. */
  advances: Map<number, TabAdvance>;
  /** Bar-stop positions in px from the content-box start edge — each draws a
   * vertical rule through the line height and consumes no advance (AC-22/23). */
  bars: number[];
}

/** Sub-pixel tolerance, matching textMetrics' FIT_EPS. */
const EPS = 0.02;

/** Characters a decimal stop may align on when the document-language
 * separator is absent (AC-20): the ASCII dot and comma plus the Unicode
 * decimal separators. */
const SEPARATOR_CHARS = new Set([".", ",", "·", "٫", "⎖"]);

/** The default tab interval in cm when a document sets none (kept local so
 * this module has no dependency on the model factory). */
export const DEFAULT_TAB_INTERVAL_CM = 1.27;

/** Converts effective stops (cm, start-relative) to pixels, sorted by
 * position — the form layoutParagraphTabs and the bar-rule renderer expect. */
export function stopsToPx(stops: TabStop[]): TabStopPx[] {
  return stops
    .map((stop) => {
      const px: TabStopPx = { positionPx: stop.positionCm * PX_PER_CM, alignment: stop.alignment };
      if (stop.leader !== undefined) {
        px.leader = stop.leader;
      }
      return px;
    })
    .toSorted((a, b) => a.positionPx - b.positionPx);
}

/** The next tab stop past `penX` (px): the nearest explicit non-bar stop
 * after it, else the next multiple of the default interval anchored at the
 * content-box start (AC-15, AC-18). Bar stops are skipped — a tab never
 * targets one (AC-23). */
function nextStop(penX: number, stopsPx: TabStopPx[], defaultPx: number): TabStopPx {
  for (const stop of stopsPx) {
    if (stop.alignment !== "bar" && stop.positionPx > penX + EPS) {
      return stop;
    }
  }
  const multiple = Math.floor(penX / defaultPx + EPS) + 1;
  return { positionPx: multiple * defaultPx, alignment: "left" };
}

/** The index a decimal stop aligns on within `[from, to)`: the document-
 * language separator first (AC-19), then the first of any separator character
 * (AC-20), else -1 for the right-alignment fallback (AC-21). */
function separatorIndex(text: string, from: number, to: number, decimalSeparator: string): number {
  if (decimalSeparator.length > 0) {
    const primary = text.indexOf(decimalSeparator, from);
    if (primary >= from && primary < to) {
      return primary;
    }
  }
  for (let i = from; i < to; i += 1) {
    if (SEPARATOR_CHARS.has(text[i])) {
      return i;
    }
  }
  return -1;
}

/** Resolves every U+0009 in `text` to a box width and collects bar-rule
 * positions. `measure(from, to)` returns the pixel width of `text[from, to)`
 * in the paragraph's fonts. The pen resets to the content-box start on each
 * hard line break (`\n`); soft wraps are not modelled (a disclosed
 * limitation — a tabbed line that soft-wraps is rare). */
export function layoutParagraphTabs(
  text: string,
  stopsPx: TabStopPx[],
  defaultPx: number,
  decimalSeparator: string,
  measure: (from: number, to: number) => number,
): ParagraphTabLayout {
  const advances = new Map<number, TabAdvance>();
  const bars = stopsPx.filter((stop) => stop.alignment === "bar").map((stop) => stop.positionPx);
  const grid = defaultPx > 0 ? defaultPx : DEFAULT_TAB_INTERVAL_CM * PX_PER_CM;
  const length = text.length;
  let penX = 0;
  let i = 0;
  while (i < length) {
    const ch = text[i];
    if (ch === "\n") {
      penX = 0;
      i += 1;
      continue;
    }
    if (ch === "\t") {
      // The text that follows this tab, up to the next tab or hard break —
      // what a right/centre/decimal stop measures its alignment against.
      let end = i + 1;
      while (end < length && text[end] !== "\t" && text[end] !== "\n") {
        end += 1;
      }
      const stop = nextStop(penX, stopsPx, grid);
      const followWidth = end > i + 1 ? measure(i + 1, end) : 0;
      let width: number;
      if (stop.alignment === "right") {
        width = stop.positionPx - penX - followWidth;
      } else if (stop.alignment === "center") {
        width = stop.positionPx - penX - followWidth / 2;
      } else if (stop.alignment === "decimal") {
        const sep = separatorIndex(text, i + 1, end, decimalSeparator);
        const prefixWidth = sep >= 0 ? measure(i + 1, sep) : followWidth;
        width = stop.positionPx - penX - prefixWidth;
      } else {
        // left, and the default grid.
        width = stop.positionPx - penX;
      }
      if (!(width > 0)) {
        width = 0;
      }
      const advance: TabAdvance = { widthPx: width };
      if (stop.leader !== undefined) {
        advance.leader = stop.leader;
      }
      advances.set(i, advance);
      penX += width;
      i += 1;
      continue;
    }
    // A plain run up to the next control character, measured in one call.
    let end = i + 1;
    while (end < length && text[end] !== "\t" && text[end] !== "\n") {
      end += 1;
    }
    penX += measure(i, end);
    i = end;
  }
  return { advances, bars };
}
