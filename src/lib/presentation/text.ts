// The framework-free text command layer for slide editing (SPEC-081). The
// editing session (src/views/presentation/useSlideTextEditing.ts) drives the
// DOM through contenteditable and reads it back with `normalizeTextFrameDom`
// (a rendering concern, kept beside the renderer); FORMATTING, by contrast,
// never round-trips through the browser — a toolbar/keyboard command edits the
// model here (splitting runs at the selection edges) and the stage re-renders.
// So this module is the normative source for run/paragraph formatting and the
// selection-address maths, all pure and unit-tested without a DOM.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type { PresentationFormat, SlideParagraph, SlideRun, TextFrame } from "./model";
import { paragraphText } from "./slideModel";

/** A caret/anchor inside a text frame: a paragraph index and a character
 * offset within that paragraph's concatenated run text. */
export interface TextPoint {
  para: number;
  offset: number;
}

/** A (possibly collapsed) selection over a frame's paragraphs. `start` is
 * always the earlier point after `normalizeRange`. */
export interface TextRange {
  start: TextPoint;
  end: TextPoint;
}

/** The run-level formatting a command sets. A boolean `false` REMOVES the mark
 * (model stores only `true`); a string/number sets it. Absent fields are left
 * untouched. */
export interface RunFormatPatch {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  sizePt?: number;
  color?: string;
}

/** The paragraph-level formatting a command sets. `align:"left"`, `bullet:null`
 * and `level:0` clear back to the model default (the field is dropped). */
export interface ParagraphFormatPatch {
  align?: SlideParagraph["align"];
  bullet?: SlideParagraph["bullet"] | null;
  level?: number;
}

/** A paragraph range (inclusive) for paragraph-level commands. */
export interface ParagraphRange {
  startPara: number;
  endPara: number;
}

/** The maximum indent level (SPEC-081: level 0–8). */
export const MAX_INDENT_LEVEL = 8;

/** A run's formatting only (its text stripped). */
type RunMarks = Omit<SlideRun, "text">;

function marksOf(run: SlideRun): RunMarks {
  const { text: _text, ...marks } = run;
  return marks;
}

function marksEqual(a: RunMarks, b: RunMarks): boolean {
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    a.fontFamily === b.fontFamily &&
    a.sizePt === b.sizePt &&
    a.color === b.color
  );
}

/** Expands a paragraph into a per-character marks array (one entry per code
 * unit), the neutral form every run-level edit works on. */
function expand(para: SlideParagraph): RunMarks[] {
  const chars: RunMarks[] = [];
  for (const run of para.runs) {
    const marks = marksOf(run);
    for (let i = 0; i < run.text.length; i++) {
      chars.push({ ...marks });
    }
  }
  return chars;
}

/** Coalesces a per-character marks array back into runs, merging adjacent
 * characters with identical marks. The plain text is supplied alongside. */
function coalesce(text: string, chars: RunMarks[]): SlideRun[] {
  const runs: SlideRun[] = [];
  let start = 0;
  for (let i = 1; i <= chars.length; i++) {
    if (i === chars.length || !marksEqual(chars[i], chars[start])) {
      runs.push(buildRun(text.slice(start, i), chars[start]));
      start = i;
    }
  }
  return runs;
}

function buildRun(text: string, marks: RunMarks): SlideRun {
  const run: SlideRun = { text };
  if (marks.bold) run.bold = true;
  if (marks.italic) run.italic = true;
  if (marks.underline) run.underline = true;
  if (marks.fontFamily) run.fontFamily = marks.fontFamily;
  if (marks.sizePt !== undefined) run.sizePt = marks.sizePt;
  if (marks.color) run.color = marks.color;
  return run;
}

function applyPatchToMarks(marks: RunMarks, patch: RunFormatPatch): void {
  if (patch.bold !== undefined) {
    if (patch.bold) marks.bold = true;
    else delete marks.bold;
  }
  if (patch.italic !== undefined) {
    if (patch.italic) marks.italic = true;
    else delete marks.italic;
  }
  if (patch.underline !== undefined) {
    if (patch.underline) marks.underline = true;
    else delete marks.underline;
  }
  if (patch.fontFamily !== undefined) {
    if (patch.fontFamily) marks.fontFamily = patch.fontFamily;
    else delete marks.fontFamily;
  }
  if (patch.sizePt !== undefined) {
    marks.sizePt = patch.sizePt;
  }
  if (patch.color !== undefined) {
    if (patch.color) marks.color = patch.color;
    else delete marks.color;
  }
}

function clamp(value: number, lo: number, hi: number): number {
  return Math.min(Math.max(value, lo), hi);
}

/** Orders a range's endpoints so `start` is not after `end`, clamped to the
 * frame's paragraphs and each paragraph's length. */
export function normalizeRange(frame: TextFrame, range: TextRange): TextRange {
  const clampPoint = (point: TextPoint): TextPoint => {
    const para = clamp(point.para, 0, Math.max(0, frame.paragraphs.length - 1));
    const len = frame.paragraphs[para] ? paragraphText(frame.paragraphs[para]).length : 0;
    return { para, offset: clamp(point.offset, 0, len) };
  };
  const start = clampPoint(range.start);
  const end = clampPoint(range.end);
  const ordered =
    start.para < end.para || (start.para === end.para && start.offset <= end.offset)
      ? { start, end }
      : { start: end, end: start };
  return ordered;
}

/** A range covering the whole frame — every paragraph, every character. Used
 * when a command applies to a selected-but-not-edited frame (AC-25). */
export function wholeFrameRange(frame: TextFrame): TextRange {
  const last = Math.max(0, frame.paragraphs.length - 1);
  const lastLen = frame.paragraphs[last] ? paragraphText(frame.paragraphs[last]).length : 0;
  return { start: { para: 0, offset: 0 }, end: { para: last, offset: lastLen } };
}

/** Applies run-level formatting to a range, splitting runs at the selection
 * edges so the mark lands on exactly the covered characters. A collapsed range
 * changes nothing (the session tracks a pending format for the next keystroke,
 * out of this pure layer). Returns whether anything changed. */
export function applyRunFormat(frame: TextFrame, range: TextRange, patch: RunFormatPatch): boolean {
  const norm = normalizeRange(frame, range);
  let changed = false;
  for (let p = norm.start.para; p <= norm.end.para; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    const text = paragraphText(para);
    const from = p === norm.start.para ? norm.start.offset : 0;
    const to = p === norm.end.para ? norm.end.offset : text.length;
    if (from >= to) continue;
    const chars = expand(para);
    for (let i = from; i < to; i++) {
      applyPatchToMarks(chars[i], patch);
    }
    para.runs = coalesce(text, chars);
    changed = true;
  }
  return changed;
}

/** Applies paragraph-level formatting (align/bullet/level) to every paragraph
 * a range touches. */
export function applyParagraphFormat(
  frame: TextFrame,
  range: ParagraphRange,
  patch: ParagraphFormatPatch,
): void {
  const start = clamp(range.startPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const end = clamp(range.endPara, 0, Math.max(0, frame.paragraphs.length - 1));
  for (let p = start; p <= end; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    if (patch.align !== undefined) {
      if (patch.align === "left") delete para.align;
      else para.align = patch.align;
    }
    if (patch.bullet !== undefined) {
      if (patch.bullet === null) delete para.bullet;
      else para.bullet = patch.bullet;
    }
    if (patch.level !== undefined) {
      const lvl = clamp(patch.level, 0, MAX_INDENT_LEVEL);
      if (lvl === 0) delete para.level;
      else para.level = lvl;
    }
  }
}

/** Toggles the disc bullet over a paragraph range: turns it OFF when every
 * paragraph in range already has a disc, otherwise turns it ON for all
 * (SPEC-081 writes "disc" only; numbered authoring is out of scope). */
export function toggleBullet(frame: TextFrame, range: ParagraphRange): void {
  const start = clamp(range.startPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const end = clamp(range.endPara, 0, Math.max(0, frame.paragraphs.length - 1));
  let allDisc = true;
  for (let p = start; p <= end; p++) {
    if (frame.paragraphs[p] && frame.paragraphs[p].bullet !== "disc") {
      allDisc = false;
      break;
    }
  }
  applyParagraphFormat(frame, { startPara: start, endPara: end }, { bullet: allDisc ? null : "disc" });
}

/** Shifts every paragraph in a range by `delta` indent levels, clamped 0–8. */
export function setIndentLevel(frame: TextFrame, range: ParagraphRange, delta: number): void {
  const start = clamp(range.startPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const end = clamp(range.endPara, 0, Math.max(0, frame.paragraphs.length - 1));
  for (let p = start; p <= end; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    const next = clamp((para.level ?? 0) + delta, 0, MAX_INDENT_LEVEL);
    if (next === 0) delete para.level;
    else para.level = next;
  }
}

/** A toolbar-facing value that may be uniform, mixed across the selection, or
 * absent. `"mixed"` renders indeterminate (AC-24). */
export type FieldState<T> = T | "mixed" | undefined;

export interface SelectionFormatState {
  bold: FieldState<boolean>;
  italic: FieldState<boolean>;
  underline: FieldState<boolean>;
  fontFamily: FieldState<string>;
  sizePt: FieldState<number>;
  color: FieldState<string>;
  align: FieldState<NonNullable<SlideParagraph["align"]>>;
  bullet: FieldState<boolean>;
  level: FieldState<number>;
}

function fold<T>(current: FieldState<T> | "unset", value: T): FieldState<T> | "unset" {
  if (current === "unset") return value;
  if (current === "mixed") return "mixed";
  return current === value ? current : "mixed";
}

/** The formatting state a toolbar reflects for a selection: each run field
 * folded over every covered character, each paragraph field over every touched
 * paragraph (AC-24). A collapsed caret reads the character to its left (or the
 * run at the caret when at a paragraph start). */
export function selectionFormatState(frame: TextFrame, range: TextRange): SelectionFormatState {
  const norm = normalizeRange(frame, range);
  let bold: FieldState<boolean> | "unset" = "unset";
  let italic: FieldState<boolean> | "unset" = "unset";
  let underline: FieldState<boolean> | "unset" = "unset";
  let fontFamily: FieldState<string> | "unset" = "unset";
  let sizePt: FieldState<number> | "unset" = "unset";
  let color: FieldState<string> | "unset" = "unset";
  let align: FieldState<NonNullable<SlideParagraph["align"]>> | "unset" = "unset";
  let bullet: FieldState<boolean> | "unset" = "unset";
  let level: FieldState<number> | "unset" = "unset";

  for (let p = norm.start.para; p <= norm.end.para; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    align = fold(align, para.align ?? "left");
    bullet = fold(bullet, para.bullet === "disc");
    level = fold(level, para.level ?? 0);

    const text = paragraphText(para);
    const chars = expand(para);
    let from = p === norm.start.para ? norm.start.offset : 0;
    let to = p === norm.end.para ? norm.end.offset : text.length;
    if (from === to) {
      // Collapsed within this paragraph: read the character left of the caret,
      // or the first character when at the very start.
      if (chars.length === 0) continue;
      from = from > 0 ? from - 1 : 0;
      to = from + 1;
    }
    for (let i = from; i < to && i < chars.length; i++) {
      const m = chars[i];
      bold = fold(bold, !!m.bold);
      italic = fold(italic, !!m.italic);
      underline = fold(underline, !!m.underline);
      fontFamily = fold(fontFamily, m.fontFamily ?? "");
      sizePt = fold(sizePt, m.sizePt ?? DEFAULT_SIZE_PT);
      color = fold(color, m.color ?? "");
    }
  }

  const clean = <T>(state: FieldState<T> | "unset"): FieldState<T> => (state === "unset" ? undefined : state);
  return {
    bold: clean(bold),
    italic: clean(italic),
    underline: clean(underline),
    fontFamily: clean(fontFamily),
    sizePt: clean(sizePt),
    color: clean(color),
    align: clean(align),
    bullet: clean(bullet),
    level: clean(level),
  };
}

/** The renderer/reader default point size — a run without `sizePt` reads as
 * this in the toolbar. Kept in sync with the readers' `DEFAULT_FONT_SIZE_PT`. */
export const DEFAULT_SIZE_PT = 18;

/** Marks a slide edited by any SPEC-081 text command: the writers rebuild it
 * on save, and an odp page drops its verbatim `sourceXml` so the page
 * serialiser regenerates it (AC-30/31). `preserved`/`partName` stay — the
 * rebuild consumes them. */
export function markSlideEdited(slide: { dirty?: true; sourceXml?: string }, format: PresentationFormat): void {
  slide.dirty = true;
  if (format === "odp") {
    delete slide.sourceXml;
  }
}
