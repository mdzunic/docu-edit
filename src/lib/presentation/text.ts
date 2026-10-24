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
import type { PresentationFormat, SlideParagraph, SlideRun } from "./model";
import { paragraphText } from "./slideModel";
import { defaultMarkerStyle, markerType, type ListMarkerStyle } from "../core/lists";
import {
  DEFAULT_THEME,
  resolveColor,
  resolveFont,
  sameModelColor,
  sameModelFont,
  type DocumentTheme,
  type ModelColor,
  type ModelFont,
} from "../core/theme";

/** The minimum a formatting command needs: a paragraph list to edit. A
 * `TextFrame` satisfies it, and so does a `SlideTableCell` (SPEC-091 AC-6a) —
 * so the exact same run/paragraph command layer edits a table cell's
 * `paragraphs` unchanged. */
export interface ParagraphOwner {
  paragraphs: SlideParagraph[];
}

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
  fontFamily?: ModelFont;
  sizePt?: number;
  color?: ModelColor;
  /** SPEC-097: letter spacing in points; `null` or 0 clears. */
  letterSpacingPt?: number | null;
  /** SPEC-097 run link: `null` clears. Setting one target clears the other
   *  (at most one of `linkHref`/`linkSlide` is ever set). */
  linkHref?: string | null;
  linkSlide?: number | null;
  linkTooltip?: string | null;
  /** SPEC-098: sets the run language tag (stored verbatim); `clearLang`
   *  drops it back to the deck default. */
  lang?: string;
  clearLang?: true;
  /** SPEC-098 "Do not check spelling": `false` removes the mark. */
  noProof?: boolean;
}

/** The paragraph-level formatting a command sets. `align:"left"`, `bullet:null`
 * and `level:0` clear back to the model default (the field is dropped); every
 * SPEC-097 field clears with `null` the same way `bullet: null` does. */
export interface ParagraphFormatPatch {
  align?: SlideParagraph["align"];
  bullet?: SlideParagraph["bullet"] | null;
  level?: number;
  bulletStyle?: ListMarkerStyle | null;
  bulletChar?: string | null;
  bulletImageId?: string | null;
  startAt?: number | null;
  indentCm?: number | null;
  hangingCm?: number | null;
  lineSpacing?: number | null;
  lineSpacingPt?: number | null;
  lineSpacingRule?: "exact" | null;
  spaceBeforePt?: number | null;
  spaceAfterPt?: number | null;
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
    sameModelFont(a.fontFamily, b.fontFamily) &&
    a.sizePt === b.sizePt &&
    sameModelColor(a.color, b.color) &&
    a.letterSpacingPt === b.letterSpacingPt &&
    a.linkHref === b.linkHref &&
    a.linkSlide === b.linkSlide &&
    a.linkTooltip === b.linkTooltip &&
    a.lang === b.lang &&
    !!a.noProof === !!b.noProof
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
  if (marks.letterSpacingPt !== undefined) run.letterSpacingPt = marks.letterSpacingPt;
  if (marks.linkHref !== undefined) run.linkHref = marks.linkHref;
  if (marks.linkSlide !== undefined) run.linkSlide = marks.linkSlide;
  if (marks.linkTooltip !== undefined) run.linkTooltip = marks.linkTooltip;
  if (marks.lang !== undefined) run.lang = marks.lang;
  if (marks.noProof) run.noProof = true;
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
  if (patch.letterSpacingPt !== undefined) {
    const pt = patch.letterSpacingPt === null ? 0 : clampLetterSpacing(patch.letterSpacingPt);
    if (pt === 0) delete marks.letterSpacingPt;
    else marks.letterSpacingPt = pt;
  }
  if (patch.linkHref !== undefined) {
    if (patch.linkHref) {
      marks.linkHref = patch.linkHref;
      delete marks.linkSlide;
    } else delete marks.linkHref;
  }
  if (patch.linkSlide !== undefined) {
    if (patch.linkSlide) {
      marks.linkSlide = patch.linkSlide;
      delete marks.linkHref;
    } else delete marks.linkSlide;
  }
  if (patch.linkTooltip !== undefined) {
    if (patch.linkTooltip) marks.linkTooltip = patch.linkTooltip;
    else delete marks.linkTooltip;
  }
  if (marks.linkHref === undefined && marks.linkSlide === undefined) {
    delete marks.linkTooltip; // a tooltip is meaningless without a link
  }
  if (patch.clearLang) delete marks.lang;
  if (patch.lang) marks.lang = patch.lang;
  if (patch.noProof !== undefined) {
    if (patch.noProof) marks.noProof = true;
    else delete marks.noProof;
  }
}

function clamp(value: number, lo: number, hi: number): number {
  return Math.min(Math.max(value, lo), hi);
}

/** Orders a range's endpoints so `start` is not after `end`, clamped to the
 * frame's paragraphs and each paragraph's length. */
export function normalizeRange(frame: ParagraphOwner, range: TextRange): TextRange {
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
export function wholeFrameRange(frame: ParagraphOwner): TextRange {
  const last = Math.max(0, frame.paragraphs.length - 1);
  const lastLen = frame.paragraphs[last] ? paragraphText(frame.paragraphs[last]).length : 0;
  return { start: { para: 0, offset: 0 }, end: { para: last, offset: lastLen } };
}

/** Applies run-level formatting to a range, splitting runs at the selection
 * edges so the mark lands on exactly the covered characters. A collapsed range
 * changes nothing (the session tracks a pending format for the next keystroke,
 * out of this pure layer). Returns whether anything changed. */
export function applyRunFormat(frame: ParagraphOwner, range: TextRange, patch: RunFormatPatch): boolean {
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

/** Replaces the half-open character range [start, end) of a paragraph's text
 * with `replacement` (SPEC-098's find-and-replace and spelling correction),
 * splitting runs at the edges. The replacement takes the marks of the run
 * containing `start` (SPEC-017's rule) and neighbouring identical runs
 * re-coalesce; paragraph-level fields are untouched. */
export function replaceParagraphText(paragraph: SlideParagraph, start: number, end: number, replacement: string): void {
  const text = paragraphText(paragraph);
  const from = clamp(start, 0, text.length);
  const to = clamp(end, from, text.length);
  const chars = expand(paragraph);
  const source = chars[from] ?? chars[from - 1] ?? {};
  const inserted = Array.from({ length: replacement.length }, () => ({ ...source }));
  paragraph.runs = coalesce(text.slice(0, from) + replacement + text.slice(to), [
    ...chars.slice(0, from),
    ...inserted,
    ...chars.slice(to),
  ]);
}

/** The SPEC-097 value ranges, clamped by every command (and by the panel). */
export const LINE_SPACING_RANGE = { min: 0.25, max: 5 } as const;
export const LINE_SPACING_PT_RANGE = { min: 4, max: 1000 } as const;
export const PARAGRAPH_SPACE_PT_RANGE = { min: 0, max: 500 } as const;
export const LETTER_SPACING_PT_RANGE = { min: -10, max: 10 } as const;
export const START_AT_RANGE = { min: 1, max: 32767 } as const;

const round = (value: number, step: number): number => Math.round(value / step) * step;
const r3 = (value: number): number => Math.round(value * 1000) / 1000;

/** Letter spacing clamped to −10…10 pt in 0.1 steps. */
export function clampLetterSpacing(pt: number): number {
  return Math.round(clamp(pt, LETTER_SPACING_PT_RANGE.min, LETTER_SPACING_PT_RANGE.max) * 10) / 10;
}

/** The patch keys whose edit clears a paragraph's preserved source properties
 * (`pptxPPr`/`odpParaProps`): the preserved `a:buFont`/`a:buClr`/`a:buSzPct`
 * and any unrecognised bullet or spacing form shadow the marker and the
 * spacing, so a command that edits those must not let stale markup be
 * re-emitted over it — exactly Table.docxTblPr's documented rule. Alignment,
 * level and the (fully modelled) indents shadow nothing preserved. */
const SHADOWING_KEYS: (keyof ParagraphFormatPatch)[] = [
  "bullet",
  "bulletStyle",
  "bulletChar",
  "bulletImageId",
  "startAt",
  "lineSpacing",
  "lineSpacingPt",
  "lineSpacingRule",
  "spaceBeforePt",
  "spaceAfterPt",
];

/** Drops a paragraph's preserved source properties (both formats). */
export function clearParagraphPreservation(para: SlideParagraph): void {
  delete para.pptxPPr;
  delete para.odpParaProps;
}

function setOrDrop<K extends keyof SlideParagraph>(para: SlideParagraph, key: K, value: SlideParagraph[K] | undefined): void {
  if (value === undefined) delete para[key];
  else para[key] = value;
}

/** Applies paragraph-level formatting to every paragraph a range touches. */
export function applyParagraphFormat(
  frame: ParagraphOwner,
  range: ParagraphRange,
  patch: ParagraphFormatPatch,
): void {
  const start = clamp(range.startPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const end = clamp(range.endPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const shadows = SHADOWING_KEYS.some((key) => patch[key] !== undefined);
  for (let p = start; p <= end; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    if (shadows) clearParagraphPreservation(para);
    if (patch.align !== undefined) {
      if (patch.align === "left") delete para.align;
      else para.align = patch.align;
    }
    if (patch.bullet !== undefined) {
      if (patch.bullet === null) {
        // Leaving the list drops every marker field with it, so a stale
        // style can never disagree with an absent `bullet`.
        delete para.bullet;
        delete para.bulletStyle;
        delete para.bulletChar;
        delete para.bulletImageId;
        delete para.startAt;
      } else para.bullet = patch.bullet;
    }
    if (patch.level !== undefined) {
      const lvl = clamp(patch.level, 0, MAX_INDENT_LEVEL);
      if (lvl === 0) delete para.level;
      else para.level = lvl;
    }
    if (patch.bulletStyle !== undefined) {
      const style = patch.bulletStyle;
      if (style === null || !para.bullet) delete para.bulletStyle;
      else {
        para.bullet = markerType(style) === "bullet" ? "disc" : "number";
        if (style === defaultStyleFor(para.bullet)) delete para.bulletStyle;
        else para.bulletStyle = style;
      }
    }
    if (patch.bulletChar !== undefined) {
      setOrDrop(para, "bulletChar", patch.bulletChar && para.bullet === "disc" ? patch.bulletChar : undefined);
    }
    if (patch.bulletImageId !== undefined) {
      setOrDrop(para, "bulletImageId", patch.bulletImageId && para.bullet ? patch.bulletImageId : undefined);
    }
    if (patch.startAt !== undefined) {
      const at = patch.startAt === null ? undefined : Math.round(clamp(patch.startAt, START_AT_RANGE.min, START_AT_RANGE.max));
      setOrDrop(para, "startAt", at === undefined || at === 1 ? undefined : at);
    }
    if (patch.indentCm !== undefined) {
      setOrDrop(para, "indentCm", patch.indentCm === null ? undefined : r3(clamp(patch.indentCm, 0, 50)));
    }
    if (patch.hangingCm !== undefined) {
      setOrDrop(para, "hangingCm", patch.hangingCm === null ? undefined : r3(clamp(patch.hangingCm, 0, 50)));
    }
    if (patch.lineSpacing !== undefined) {
      const v = patch.lineSpacing === null ? undefined : r3(clamp(patch.lineSpacing, LINE_SPACING_RANGE.min, LINE_SPACING_RANGE.max));
      setOrDrop(para, "lineSpacing", v === 1 ? undefined : v);
    }
    if (patch.lineSpacingPt !== undefined) {
      const v = patch.lineSpacingPt === null ? undefined : round(clamp(patch.lineSpacingPt, LINE_SPACING_PT_RANGE.min, LINE_SPACING_PT_RANGE.max), 0.01);
      setOrDrop(para, "lineSpacingPt", v === undefined ? undefined : r3(v));
    }
    if (patch.lineSpacingRule !== undefined) {
      setOrDrop(para, "lineSpacingRule", patch.lineSpacingRule ?? undefined);
    }
    if (patch.spaceBeforePt !== undefined) {
      const v = patch.spaceBeforePt === null ? 0 : r3(clamp(patch.spaceBeforePt, PARAGRAPH_SPACE_PT_RANGE.min, PARAGRAPH_SPACE_PT_RANGE.max));
      setOrDrop(para, "spaceBeforePt", v === 0 ? undefined : v);
    }
    if (patch.spaceAfterPt !== undefined) {
      const v = patch.spaceAfterPt === null ? 0 : r3(clamp(patch.spaceAfterPt, PARAGRAPH_SPACE_PT_RANGE.min, PARAGRAPH_SPACE_PT_RANGE.max));
      setOrDrop(para, "spaceAfterPt", v === 0 ? undefined : v);
    }
    if (para.lineSpacingPt === undefined) delete para.lineSpacingRule;
    if (para.bullet !== "disc") delete para.bulletChar;
  }
}

/** The model default marker style of a list type (AC-1's absent rule). */
function defaultStyleFor(bullet: NonNullable<SlideParagraph["bullet"]>): ListMarkerStyle {
  return bullet === "number" ? "decimal" : "disc";
}

/** A list paragraph's EFFECTIVE marker style: `bulletStyle` when it agrees
 * with `bullet` (markerType), else the type's absent default; undefined for a
 * paragraph that is not a list item. */
export function paragraphMarkerStyle(para: SlideParagraph): ListMarkerStyle | undefined {
  if (!para.bullet) return undefined;
  const type = para.bullet === "number" ? "number" : "bullet";
  if (para.bulletStyle && markerType(para.bulletStyle) === type) return para.bulletStyle;
  return defaultStyleFor(para.bullet);
}

/** Toggles a list over a paragraph range (SPEC-081, widened by SPEC-097):
 * turns it OFF when every paragraph in range already is a list item of
 * `style`'s type (bulleted for a bullet style, numbered for a number style),
 * otherwise turns it ON for all with `style`. The default third argument
 * keeps every SPEC-081 call site's "disc" behaviour unchanged. */
export function toggleBullet(frame: ParagraphOwner, range: ParagraphRange, style: ListMarkerStyle = "disc"): void {
  const start = clamp(range.startPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const end = clamp(range.endPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const bullet = markerType(style) === "bullet" ? "disc" : "number";
  let allOn = true;
  for (let p = start; p <= end; p++) {
    if (frame.paragraphs[p] && frame.paragraphs[p].bullet !== bullet) {
      allOn = false;
      break;
    }
  }
  const pr = { startPara: start, endPara: end };
  if (allOn) {
    applyParagraphFormat(frame, pr, { bullet: null });
    return;
  }
  applyParagraphFormat(frame, pr, { bullet, bulletChar: null, bulletImageId: null });
  applyParagraphFormat(frame, pr, { bulletStyle: style });
}

/** What `setListStyle` applies: a marker style (bullet or number — the
 * paragraphs' list type follows it), a custom bullet glyph, or a picture
 * bullet (a `model.media` part name) — both of the latter make the items
 * bulleted. */
export type ListStyleChoice = ListMarkerStyle | { bulletChar: string } | { bulletImageId: string };

/** Sets the list marker of every paragraph in range (making each a list item
 * when it was not). A custom glyph/picture keeps (or starts) a bulleted list. */
export function setListStyle(frame: ParagraphOwner, range: ParagraphRange, choice: ListStyleChoice): void {
  if (typeof choice === "string") {
    applyParagraphFormat(frame, range, {
      bullet: markerType(choice) === "bullet" ? "disc" : "number",
      bulletChar: null,
      bulletImageId: null,
    });
    applyParagraphFormat(frame, range, { bulletStyle: choice });
    return;
  }
  if ("bulletChar" in choice) {
    applyParagraphFormat(frame, range, { bullet: "disc", bulletImageId: null, bulletStyle: null });
    applyParagraphFormat(frame, range, { bulletChar: choice.bulletChar });
    return;
  }
  // A picture bullet is a bullet (pptx `a:buBlip` replaces any `a:buAutoNum`),
  // so the items become bulleted and drop their numbering fields.
  applyParagraphFormat(frame, range, { bullet: "disc", bulletChar: null, bulletStyle: null, startAt: null });
  applyParagraphFormat(frame, range, { bulletImageId: choice.bulletImageId });
}

/** Paragraph spacing before/after (pt, 0–500; `null`/0 clears). */
export function setParagraphSpacing(
  frame: ParagraphOwner,
  range: ParagraphRange,
  spacing: { spaceBeforePt?: number | null; spaceAfterPt?: number | null },
): void {
  applyParagraphFormat(frame, range, spacing);
}

/** A line-spacing choice in the panel's vocabulary (CustomSpacingDialog's
 * multiple / exactly / at least). */
export type LineSpacingChoice =
  | { mode: "multiple"; value: number }
  | { mode: "exact"; pt: number }
  | { mode: "atLeast"; pt: number };

/** Sets line spacing: a multiple drops any point value; a point value wins
 * over the multiple (AC-7) with the rule `"exact"` or absent ("at least"). */
export function setLineSpacing(frame: ParagraphOwner, range: ParagraphRange, choice: LineSpacingChoice): void {
  if (choice.mode === "multiple") {
    applyParagraphFormat(frame, range, { lineSpacing: choice.value, lineSpacingPt: null, lineSpacingRule: null });
  } else {
    applyParagraphFormat(frame, range, {
      lineSpacing: null,
      lineSpacingPt: choice.pt,
      lineSpacingRule: choice.mode === "exact" ? "exact" : null,
    });
  }
}

/** Letter spacing over a character range (pt, −10…10 in 0.1 steps; 0
 * clears). Returns whether anything changed. */
export function setLetterSpacing(frame: ParagraphOwner, range: TextRange, pt: number): boolean {
  return applyRunFormat(frame, range, { letterSpacingPt: pt });
}

/** Shifts every paragraph in a range by `delta` indent levels, clamped 0–8. */
export function setIndentLevel(frame: ParagraphOwner, range: ParagraphRange, delta: number): void {
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

/** Demote (+1) / promote (−1) wrapping `setIndentLevel` (SPEC-097): a list
 * item whose marker is its old level's default-cycle style (core/lists.ts —
 * disc/circle/square, decimal/a./i.) moves to the new level's default, so a
 * demoted "3." reads "a."; an explicit indent shifts by one
 * LIST_INDENT_STEP_CM with it. Clamped 0–8. */
function shiftLevel(frame: ParagraphOwner, range: ParagraphRange, delta: 1 | -1): void {
  const start = clamp(range.startPara, 0, Math.max(0, frame.paragraphs.length - 1));
  const end = clamp(range.endPara, 0, Math.max(0, frame.paragraphs.length - 1));
  for (let p = start; p <= end; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    const from = para.level ?? 0;
    const to = clamp(from + delta, 0, MAX_INDENT_LEVEL);
    if (to === from) continue;
    const style = paragraphMarkerStyle(para);
    if (style && para.bullet && !para.bulletChar && !para.bulletImageId) {
      const type = para.bullet === "number" ? "number" : "bullet";
      if (style === defaultMarkerStyle(type, from + 1)) {
        const next = defaultMarkerStyle(type, to + 1);
        if (next === defaultStyleFor(para.bullet)) delete para.bulletStyle;
        else para.bulletStyle = next;
      }
    }
    if (para.indentCm !== undefined) para.indentCm = r3(Math.max(0, para.indentCm + (to - from) * 1.27));
    setIndentLevel(frame, { startPara: p, endPara: p }, to - from);
  }
}

export function demote(frame: ParagraphOwner, range: ParagraphRange): void {
  shiftLevel(frame, range, 1);
}

export function promote(frame: ParagraphOwner, range: ParagraphRange): void {
  shiftLevel(frame, range, -1);
}

/** A paragraph's line spacing as a single comparable token, so the selection
 * state folds it like any scalar: `multiple:1.5`, `exact:24`, `atLeast:18`. */
export function lineSpacingKey(para: SlideParagraph): string {
  if (para.lineSpacingPt !== undefined) {
    return `${para.lineSpacingRule === "exact" ? "exact" : "atLeast"}:${para.lineSpacingPt}`;
  }
  return `multiple:${para.lineSpacing ?? 1}`;
}

/** The inverse of `lineSpacingKey`. */
export function parseLineSpacingKey(key: string): LineSpacingChoice {
  const [mode, raw] = key.split(":");
  const value = Number(raw);
  if (mode === "exact") return { mode: "exact", pt: value };
  if (mode === "atLeast") return { mode: "atLeast", pt: value };
  return { mode: "multiple", value: Number.isFinite(value) ? value : 1 };
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
  /** SPEC-097: the effective marker style, `"none"` for a non-list paragraph. */
  bulletStyle: FieldState<ListMarkerStyle | "none">;
  /** Whether the paragraphs are numbered (the slides-numbering toggle). */
  numbered: FieldState<boolean>;
  startAt: FieldState<number>;
  /** `lineSpacingKey` tokens (`multiple:1.5`, `exact:24`, `atLeast:18`). */
  lineSpacing: FieldState<string>;
  spaceBeforePt: FieldState<number>;
  spaceAfterPt: FieldState<number>;
  letterSpacingPt: FieldState<number>;
  /** SPEC-098: the runs' own language tag (`""` where a run has none and
   *  inherits the deck default) and their "Do not check spelling" mark. */
  lang: FieldState<string>;
  noProof: FieldState<boolean>;
}

function fold<T>(current: FieldState<T> | "unset", value: T): FieldState<T> | "unset" {
  if (current === "unset") return value;
  if (current === "mixed") return "mixed";
  return current === value ? current : "mixed";
}

const clean = <T>(state: FieldState<T> | "unset"): FieldState<T> => (state === "unset" ? undefined : state);

/** The formatting state a toolbar reflects for a selection: each run field
 * folded over every covered character, each paragraph field over every touched
 * paragraph (AC-24). A collapsed caret reads the character to its left (or the
 * run at the caret when at a paragraph start). */
export function selectionFormatState(
  frame: ParagraphOwner,
  range: TextRange,
  theme: DocumentTheme = DEFAULT_THEME,
): SelectionFormatState {
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
  let bulletStyle: FieldState<ListMarkerStyle | "none"> | "unset" = "unset";
  let numbered: FieldState<boolean> | "unset" = "unset";
  let startAt: FieldState<number> | "unset" = "unset";
  let lineSpacing: FieldState<string> | "unset" = "unset";
  let spaceBeforePt: FieldState<number> | "unset" = "unset";
  let spaceAfterPt: FieldState<number> | "unset" = "unset";
  let letterSpacingPt: FieldState<number> | "unset" = "unset";
  let lang: FieldState<string> | "unset" = "unset";
  let noProof: FieldState<boolean> | "unset" = "unset";

  for (let p = norm.start.para; p <= norm.end.para; p++) {
    const para = frame.paragraphs[p];
    if (!para) continue;
    align = fold(align, para.align ?? "left");
    bullet = fold(bullet, para.bullet === "disc");
    level = fold(level, para.level ?? 0);
    bulletStyle = fold(bulletStyle, paragraphMarkerStyle(para) ?? "none");
    numbered = fold(numbered, para.bullet === "number");
    startAt = fold(startAt, para.startAt ?? 1);
    lineSpacing = fold(lineSpacing, lineSpacingKey(para));
    spaceBeforePt = fold(spaceBeforePt, para.spaceBeforePt ?? 0);
    spaceAfterPt = fold(spaceAfterPt, para.spaceAfterPt ?? 0);

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
      // The toolbar reflects EFFECTIVE values, so refs fold resolved (the
      // single-resolution-point rule, SPEC-094 AC-13).
      fontFamily = fold(fontFamily, m.fontFamily ? resolveFont(m.fontFamily, theme) : "");
      sizePt = fold(sizePt, m.sizePt ?? DEFAULT_SIZE_PT);
      color = fold(color, m.color ? resolveColor(m.color, theme) : "");
      letterSpacingPt = fold(letterSpacingPt, m.letterSpacingPt ?? 0);
      lang = fold(lang, m.lang ?? "");
      noProof = fold(noProof, !!m.noProof);
    }
  }

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
    bulletStyle: clean(bulletStyle),
    numbered: clean(numbered),
    startAt: clean(startAt),
    lineSpacing: clean(lineSpacing),
    spaceBeforePt: clean(spaceBeforePt),
    spaceAfterPt: clean(spaceAfterPt),
    letterSpacingPt: clean(letterSpacingPt),
    lang: clean(lang),
    noProof: clean(noProof),
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
