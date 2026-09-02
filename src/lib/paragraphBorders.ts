// Paragraph borders and shading (SPEC-074): the merge rule for consecutive
// identically-bordered paragraphs, the per-fragment edge arithmetic a page
// break needs, and the CSS a bordered paragraph paints with. Pure and
// framework-free — the layout code, the renderer, the print path and the
// tests all read the same answers from here.
//
// It lives a layer above core/borders.ts because it speaks the *document
// model*: what a Block is, which of them are paragraphs, what their indents
// are. core/borders.ts stays the product-neutral vocabulary.
import {
  borderSideCss,
  borderSideVisible,
  borderSidesEqual,
  type BorderLineStyle,
  type BorderSide,
} from "./core/borders";
import {
  DEFAULT_PARAGRAPH_BORDER_PADDING_PT,
  isParagraph,
  type Block,
  type Paragraph,
  type ParagraphBorders,
  type ParagraphFormatting,
} from "./documentModel";

/** The four physical padding edges, in the order the CSS shorthand wants. */
export type BorderEdge = "top" | "bottom" | "left" | "right";

/** One edge's padding in points: the paragraph's own value, else the 4pt
 * default (SPEC-074 AC-9). */
export function paragraphPaddingPt(borders: ParagraphBorders | undefined, edge: BorderEdge): number {
  const value = borders?.paddingPt?.[edge];
  return value === undefined ? DEFAULT_PARAGRAPH_BORDER_PADDING_PT : value;
}

/** Whether a paragraph paints anything at all — any visible edge or a
 * shading fill. An all-`"none"` `borders` object paints nothing but is still
 * meaningful as an explicit override of a style's borders, so it survives in
 * the model; it just contributes no CSS. */
export function hasBorderPaint(formatting: ParagraphFormatting | undefined): boolean {
  if (!formatting) {
    return false;
  }
  if (formatting.shadingFill !== undefined) {
    return true;
  }
  const borders = formatting.borders;
  return (
    borders !== undefined &&
    (borderSideVisible(borders.top) ||
      borderSideVisible(borders.bottom) ||
      borderSideVisible(borders.left) ||
      borderSideVisible(borders.right) ||
      borderSideVisible(borders.between))
  );
}

/** Whether two paragraphs' padding is the same, treating an absent edge as
 * the 4pt default (SPEC-074 AC-9) — a paragraph that spells its padding out
 * as 4pt on every edge IS identical to one that leaves it absent, so the two
 * still merge (AC-12). */
function paddingEqual(a: ParagraphBorders | undefined, b: ParagraphBorders | undefined): boolean {
  const edges: BorderEdge[] = ["top", "bottom", "left", "right"];
  return edges.every((edge) => paragraphPaddingPt(a, edge) === paragraphPaddingPt(b, edge));
}

/** Structural equality for two paragraphs' borders (SPEC-074 AC-12's
 * "identical"). Absent equals absent; otherwise every edge and the
 * default-normalised padding must match. */
export function paragraphBordersEqual(
  a: ParagraphBorders | undefined,
  b: ParagraphBorders | undefined,
): boolean {
  if (a === undefined || b === undefined) {
    // An absent `borders` is only identical to another absent one — but two
    // paragraphs with no borders at all never form a run anyway (see
    // borderRuns), so this arm only matters when one side has shading.
    return a === b;
  }
  return (
    borderSidesEqual(a.top, b.top) &&
    borderSidesEqual(a.bottom, b.bottom) &&
    borderSidesEqual(a.left, b.left) &&
    borderSidesEqual(a.right, b.right) &&
    borderSidesEqual(a.between, b.between) &&
    paddingEqual(a, b)
  );
}

/** Which edges one paragraph of a border run paints (SPEC-074 AC-12).
 * `between` is "paint the between edge *below* me", i.e. this paragraph is
 * an internal member of a run whose `between` edge is set. */
export interface BorderRunFlags {
  top: boolean;
  bottom: boolean;
  sides: boolean;
  between: boolean;
  /** Another member of the same run follows this paragraph: the box carries
   * on below it. Independent of `between` — a "Box" run has no between edge
   * yet still must not let a margin open at its internal boundaries, or the
   * side rules break there (AC-12). */
  continuesBelow: boolean;
}

const NO_PAINT: BorderRunFlags = {
  top: false,
  bottom: false,
  sides: false,
  between: false,
  continuesBelow: false,
};

/** Whether two consecutive paragraphs merge into one bordered block: both
 * paint something, their borders are structurally equal, their shading is
 * equal, and their indents agree (SPEC-074 "Merging" — a different indent
 * breaks the run because the two boxes would not line up). */
function mergesWith(previous: Paragraph, next: Paragraph): boolean {
  const a = previous.formatting;
  const b = next.formatting;
  if (!hasBorderPaint(a) || !hasBorderPaint(b)) {
    return false;
  }
  if (a?.shadingFill !== b?.shadingFill) {
    return false;
  }
  if (!paragraphBordersEqual(a?.borders, b?.borders)) {
    return false;
  }
  return (
    (a?.indentStartCm ?? 0) === (b?.indentStartCm ?? 0) &&
    (a?.indentEndCm ?? 0) === (b?.indentEndCm ?? 0)
  );
}

/** Groups consecutive blocks into border runs (SPEC-074 AC-12). Returns, per
 * block index, which edges that paragraph should paint: the outer edges draw
 * once around the run (top above the first, bottom below the last), the
 * sides draw down every member, and the `between` edge draws at each
 * internal boundary. A non-paragraph block, or a paragraph that paints
 * nothing, gets all-false and breaks any run in progress. */
export function borderRuns(blocks: readonly Block[]): BorderRunFlags[] {
  const flags: BorderRunFlags[] = blocks.map(() => NO_PAINT);
  let index = 0;
  while (index < blocks.length) {
    const block = blocks[index];
    if (!isParagraph(block) || !hasBorderPaint(block.formatting)) {
      index += 1;
      continue;
    }
    // Extend the run as far as consecutive identical paragraphs allow.
    let end = index;
    while (end + 1 < blocks.length) {
      const next = blocks[end + 1];
      if (!isParagraph(next) || !mergesWith(blocks[end] as Paragraph, next)) {
        break;
      }
      end += 1;
    }
    const between = borderSideVisible(block.formatting?.borders?.between);
    for (let i = index; i <= end; i++) {
      flags[i] = {
        top: i === index && borderSideVisible(block.formatting?.borders?.top),
        bottom: i === end && borderSideVisible(block.formatting?.borders?.bottom),
        sides: true,
        between: i < end && between,
        continuesBelow: i < end,
      };
    }
    index = end + 1;
  }
  return flags;
}

// --- painting (AC-13..16) ----------------------------------------------------

/** Which edges one *fragment* paints. A paragraph that fits on one page has
 * exactly one fragment carrying its run flags; a paragraph split across a
 * page break (AC-14..16) has several — the sides paint on all of them, the
 * top only on the first and the bottom only on the last. */
export interface BorderPaintEdges extends BorderRunFlags {
  /** This fragment sits directly under a run's internal `between` rule: it
   * takes the top padding (the rule is right above it) but no top border of
   * its own. */
  belowBetween: boolean;
  /** The box carries on above this fragment — it is a run member with another
   * member before it. Collapses the top margin so no gap opens in the side
   * rules (AC-12). Implied by `belowBetween`, but true for a plain "Box" run
   * as well, which has no `between` edge at all. */
  continuesAbove: boolean;
}

/** A standalone paragraph's edges — every side it sets, painted once. Used
 * by the render paths that have no pagination context (shape text frames,
 * the outline pane's previews, a table cell's paragraphs). */
export const ALL_EDGES: BorderPaintEdges = {
  top: true,
  bottom: true,
  sides: true,
  between: false,
  continuesBelow: false,
  belowBetween: false,
  continuesAbove: false,
};

/** The fragment edges for paragraph `index` of `runFlags`, narrowed by where
 * this fragment sits within the paragraph's own page-split (SPEC-074
 * AC-14..16): the top border only on the fragment holding the first line,
 * the bottom only on the fragment holding the last, the sides on all. */
export function fragmentEdges(
  runFlags: readonly BorderRunFlags[],
  index: number,
  firstFragment: boolean,
  lastFragment: boolean,
): BorderPaintEdges {
  const flags = runFlags[index] ?? NO_PAINT;
  // The run's internal boundaries are the paragraph's own first/last edges:
  // the fragments in between are page-split seams, which the page break —
  // not a margin — already separates.
  const previous = runFlags[index - 1] ?? NO_PAINT;
  return {
    top: flags.top && firstFragment,
    bottom: flags.bottom && lastFragment,
    sides: flags.sides,
    between: flags.between && lastFragment,
    continuesBelow: flags.continuesBelow && lastFragment,
    belowBetween: firstFragment && index > 0 && previous.between,
    continuesAbove: firstFragment && index > 0 && previous.continuesBelow,
  };
}

function pt(value: number): string {
  return `${Math.round(value * 1000) / 1000}pt`;
}

/** The CSS a bordered/shaded paragraph fragment paints with (AC-13..16).
 * Every border and padding key is emitted explicitly — including the
 * `"none"`/`"0"` ones — so a caller can merge this over a base style object
 * and be sure no edge from a previous computation survives.
 *
 * The shading fills the padding box, so it spans the paragraph's full
 * content width including its padding (AC-13); the paragraph's indents stay
 * on its margins, so the whole box follows them. */
export function paragraphBorderStyle(
  formatting: ParagraphFormatting | undefined,
  edges: BorderPaintEdges = ALL_EDGES,
): Record<string, string> {
  if (!hasBorderPaint(formatting)) {
    return {};
  }
  const borders = formatting?.borders;
  const style: Record<string, string> = {};

  const top: BorderSide | undefined = edges.top ? borders?.top : undefined;
  // The run's internal rule is the `between` edge drawn as this fragment's
  // bottom border — which is exactly what OOXML's w:between means.
  const bottom: BorderSide | undefined = edges.bottom
    ? borders?.bottom
    : edges.between
      ? borders?.between
      : undefined;
  const left: BorderSide | undefined = edges.sides ? borders?.left : undefined;
  const right: BorderSide | undefined = edges.sides ? borders?.right : undefined;

  style.borderTop = borderSideCss(top);
  style.borderBottom = borderSideCss(bottom);
  style.borderLeft = borderSideCss(left);
  style.borderRight = borderSideCss(right);

  // Padding sits on an edge only where that edge's rule is actually drawn —
  // a bottom-rule heading (a `---` autoformat rule) must not gain a 4pt
  // indent on three sides it has no border on. A fragment continuing onto
  // the next page draws no top/bottom rule and so takes no vertical padding
  // there either, which is what keeps a split box flush with the page edge.
  style.paddingTop = borderSideVisible(top) || edges.belowBetween ? pt(paragraphPaddingPt(borders, "top")) : "0";
  style.paddingBottom = borderSideVisible(bottom) ? pt(paragraphPaddingPt(borders, "bottom")) : "0";
  style.paddingLeft = borderSideVisible(left) ? pt(paragraphPaddingPt(borders, "left")) : "0";
  style.paddingRight = borderSideVisible(right) ? pt(paragraphPaddingPt(borders, "right")) : "0";

  if (formatting?.shadingFill !== undefined) {
    style.backgroundColor = formatting.shadingFill;
    // Formatting marks are drawn on the shading, so the shading is the
    // surface they contrast against (AC-26).
    Object.assign(style, paragraphMarkColorStyle(formatting));
  }
  // Merged paragraphs must read as one block: no gap can open between two
  // members of a run, or the page would show through the box where the side
  // rules break (AC-12). This is run membership, not the `between` edge —
  // a "Box" run draws no internal rule and still must not come apart at a
  // member's space-before/space-after.
  if (edges.continuesAbove) {
    style.marginTop = "0";
  }
  if (edges.continuesBelow) {
    style.marginBottom = "0";
  }
  return style;
}

// --- formatting marks on a shaded paragraph (AC-26) --------------------------

/** The canvas's default formatting-mark grey (SPEC-028) — what a mark paints
 * as on unshaded paper. */
export const DEFAULT_MARK_COLOR = "#80868b";

/** The mark colour a shaded paragraph needs (SPEC-074 AC-26): formatting
 * marks are drawn *on the shading*, so the shading — not the page — is the
 * surface they must contrast against. The default grey survives on light
 * fills; on a dark one the marks flip to a light grey, which is the only way
 * a pilcrow stays visible in a dark callout.
 *
 * Returned as the `--mark-color` custom property the mark rules read, so the
 * decision is made once here rather than duplicated across four CSS rules. */
export function paragraphMarkColorStyle(
  formatting: ParagraphFormatting | undefined,
): Record<string, string> {
  const fill = formatting?.shadingFill;
  if (fill === undefined) {
    return {};
  }
  return { "--mark-color": isDarkFill(fill) ? "#d0d3d6" : DEFAULT_MARK_COLOR };
}

/** Whether a `"#rrggbb"` fill is dark enough that the default mark grey
 * would disappear into it. Relative luminance under the midpoint, the same
 * sRGB weighting the accessibility checker's contrast maths uses. */
export function isDarkFill(hex: string): boolean {
  const clean = hex.replace("#", "");
  if (clean.length !== 6) {
    return false;
  }
  const channel = (offset: number): number => {
    const value = Number.parseInt(clean.slice(offset, offset + 2), 16);
    return Number.isNaN(value) ? 0 : value / 255;
  };
  const linear = (c: number): number => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const luminance = 0.2126 * linear(channel(0)) + 0.7152 * linear(channel(2)) + 0.0722 * linear(channel(4));
  return luminance < 0.4;
}

// --- presets (AC-8) ----------------------------------------------------------

/** The dialog's preset row (SPEC-074 AC-8). `"custom"` is never applied —
 * it is what `detectBorderPreset` reports for an edge set none of the named
 * presets describes. */
export type ParagraphBorderPreset = "none" | "box" | "all" | "topAndBottom" | "custom";

/** Which preset an edge set matches, or `"custom"` (AC-8). */
export function detectBorderPreset(borders: ParagraphBorders | undefined): ParagraphBorderPreset {
  const top = borderSideVisible(borders?.top);
  const bottom = borderSideVisible(borders?.bottom);
  const left = borderSideVisible(borders?.left);
  const right = borderSideVisible(borders?.right);
  const between = borderSideVisible(borders?.between);
  if (!top && !bottom && !left && !right && !between) {
    return "none";
  }
  const outer = top && bottom && left && right;
  // Every named preset draws one identical line on each of its edges; a run
  // of differing widths or colours is a custom set even when the *shape*
  // matches.
  if (outer && !sidesUniform(borders, between)) {
    return "custom";
  }
  if (outer && between) {
    return "all";
  }
  if (outer && !between) {
    return "box";
  }
  if (top && bottom && !left && !right && !between && borderSidesEqual(borders?.top, borders?.bottom)) {
    return "topAndBottom";
  }
  return "custom";
}

function sidesUniform(borders: ParagraphBorders | undefined, includeBetween: boolean): boolean {
  const reference = borders?.top;
  const sides = [borders?.bottom, borders?.left, borders?.right];
  if (includeBetween) {
    sides.push(borders?.between);
  }
  return sides.every((side) => borderSidesEqual(reference, side));
}

/** Applies a preset (AC-8), drawing `line` on each of the preset's edges and
 * keeping whatever padding the paragraph already had. `"none"` returns
 * undefined — no borders at all — and `"custom"` leaves `existing`
 * untouched, since it describes an edge set the user built by hand. */
export function applyBorderPreset(
  preset: ParagraphBorderPreset,
  line: BorderSide,
  existing?: ParagraphBorders,
): ParagraphBorders | undefined {
  if (preset === "none") {
    return undefined;
  }
  if (preset === "custom") {
    return existing;
  }
  const borders: ParagraphBorders = {};
  if (existing?.paddingPt !== undefined) {
    borders.paddingPt = { ...existing.paddingPt };
  }
  borders.top = { ...line };
  borders.bottom = { ...line };
  if (preset === "box" || preset === "all") {
    borders.left = { ...line };
    borders.right = { ...line };
  }
  if (preset === "all") {
    borders.between = { ...line };
  }
  return borders;
}

// --- the dialog's editable draft (AC-6..AC-9) --------------------------------
//
// The Borders and shading dialog edits five independent edges, each with an
// on/off, a style, a width and a colour, plus four paddings. That state lives
// here rather than in the `.vue` because espree-parsed `.vue` scripts cannot
// carry the type annotations it needs (see AppShell.vue) — and because a
// pure draft is testable without mounting a component.

/** One edge's editable state. `on` is separate from `style: "none"` so
 * turning an edge off keeps the line the user picked for when it comes back. */
export interface BorderEdgeDraft {
  on: boolean;
  style: BorderLineStyle;
  widthPt: number;
  color: string;
}

/** Every edge and padding the dialog edits. Padding is in points — the
 * model's unit and OOXML's `@w:space`; the dialog converts for display. */
export interface BorderDraft {
  top: BorderEdgeDraft;
  bottom: BorderEdgeDraft;
  left: BorderEdgeDraft;
  right: BorderEdgeDraft;
  between: BorderEdgeDraft;
  paddingPt: { top: number; bottom: number; left: number; right: number };
}

/** Every edge the draft holds, in the order the dialog lists them. */
export const BORDER_DRAFT_EDGES = ["top", "bottom", "left", "right", "between"] as const;

export type BorderDraftEdge = (typeof BORDER_DRAFT_EDGES)[number];

/** The line an edge seeds from when it has none of its own. */
const DEFAULT_LINE: BorderSide = { style: "single", widthPt: 1, color: "#000000" };

function edgeDraftFrom(side: BorderSide | undefined): BorderEdgeDraft {
  const visible = borderSideVisible(side);
  return {
    on: visible,
    style: visible && side ? side.style : DEFAULT_LINE.style,
    widthPt: visible && side ? side.widthPt : DEFAULT_LINE.widthPt,
    // "auto" is a file-format value with no colour-input equivalent; it
    // shows (and, if untouched, re-saves) as the black it always painted as.
    color: visible && side && side.color !== "auto" ? side.color : DEFAULT_LINE.color,
  };
}

/** The dialog's opening state for a paragraph's (or a style's) borders. */
export function borderDraftFrom(borders: ParagraphBorders | undefined): BorderDraft {
  return {
    top: edgeDraftFrom(borders?.top),
    bottom: edgeDraftFrom(borders?.bottom),
    left: edgeDraftFrom(borders?.left),
    right: edgeDraftFrom(borders?.right),
    between: edgeDraftFrom(borders?.between),
    paddingPt: {
      top: paragraphPaddingPt(borders, "top"),
      bottom: paragraphPaddingPt(borders, "bottom"),
      left: paragraphPaddingPt(borders, "left"),
      right: paragraphPaddingPt(borders, "right"),
    },
  };
}

/** The draft as a `ParagraphBorders`, or undefined when no edge paints.
 * Padding is only carried when it differs from the 4pt default (AC-9), so an
 * untouched paragraph stays byte-identical through a save. */
export function bordersFromDraft(draft: BorderDraft): ParagraphBorders | undefined {
  const borders: ParagraphBorders = {};
  for (const edge of BORDER_DRAFT_EDGES) {
    const state = draft[edge];
    if (state.on) {
      borders[edge] = { style: state.style, widthPt: state.widthPt, color: state.color };
    }
  }
  const padding = draft.paddingPt;
  if (
    padding.top !== DEFAULT_PARAGRAPH_BORDER_PADDING_PT ||
    padding.bottom !== DEFAULT_PARAGRAPH_BORDER_PADDING_PT ||
    padding.left !== DEFAULT_PARAGRAPH_BORDER_PADDING_PT ||
    padding.right !== DEFAULT_PARAGRAPH_BORDER_PADDING_PT
  ) {
    borders.paddingPt = { ...padding };
  }
  return normaliseParagraphBorders(borders);
}

/** The draft a preset produces (AC-8), drawing on whichever line the user
 * already has: switching Box → All keeps their style, width and colour
 * rather than resetting to a plain hairline. Padding is untouched. */
export function applyPresetToDraft(preset: ParagraphBorderPreset, draft: BorderDraft): BorderDraft {
  const line = currentDraftLine(draft);
  const applied = applyBorderPreset(preset, line, bordersFromDraft(draft));
  const next: BorderDraft = { ...draft, paddingPt: { ...draft.paddingPt } };
  for (const edge of BORDER_DRAFT_EDGES) {
    const side = applied?.[edge];
    next[edge] = side
      ? { on: true, style: side.style, widthPt: side.widthPt, color: side.color }
      : { ...draft[edge], on: false };
  }
  return next;
}

/** The line a preset draws with: the first edge that is on, else the
 * default hairline. */
function currentDraftLine(draft: BorderDraft): BorderSide {
  for (const edge of BORDER_DRAFT_EDGES) {
    const state = draft[edge];
    if (state.on) {
      return { style: state.style, widthPt: state.widthPt, color: state.color };
    }
  }
  return { ...DEFAULT_LINE };
}

/** Sets one edge's line style from a picker's plain string (AC-7). Choosing
 * `"none"` is how an edge is switched off from the keyboard; every other
 * value implies the edge is on. */
export function setDraftEdgeStyle(draft: BorderDraft, edge: BorderDraftEdge, style: string): BorderDraft {
  const next = { ...draft, paddingPt: { ...draft.paddingPt } };
  if (style === "none") {
    next[edge] = { ...draft[edge], on: false };
  } else {
    next[edge] = { ...draft[edge], on: true, style: style as BorderLineStyle };
  }
  return next;
}

/** Sets one edge's width, snapped to the eighth-of-a-point grid `@w:sz`
 * carries exactly. A non-positive or unparseable entry leaves it alone. */
export function setDraftEdgeWidth(draft: BorderDraft, edge: BorderDraftEdge, widthPt: number): BorderDraft {
  if (!Number.isFinite(widthPt) || widthPt <= 0) {
    return draft;
  }
  const next = { ...draft, paddingPt: { ...draft.paddingPt } };
  next[edge] = { ...draft[edge], widthPt: Math.round(widthPt * 8) / 8 };
  return next;
}

/** Typed pass-through for a preset key coming from a `.vue` template's plain
 * string (espree bars the annotation — see documentModel.asTabStopAlignment). */
export function asParagraphBorderPreset(value: string): ParagraphBorderPreset {
  return value as ParagraphBorderPreset;
}

/** Typed pass-through for an edge key coming from a `.vue` template's plain
 * string. */
export function asBorderDraftEdge(value: string): BorderDraftEdge {
  return value as BorderDraftEdge;
}

/** Typed empty seed for a `.vue` ref holding a draft (espree bars the
 * annotation, so the ref takes its type from this factory). */
export function emptyBorderDraft(): BorderDraft {
  return borderDraftFrom(undefined);
}

/** The padding edges as plain strings, for a `.vue` template's `v-for`
 * (which cannot narrow them back to `BorderEdge` on its own). */
export const BORDER_PADDING_EDGES: readonly BorderEdge[] = ["top", "bottom", "left", "right"];

/** The dialog's padding fields, held as display text per edge — a typed
 * record so the `.vue` can index it with a `BorderEdge` from the `v-for`. */
export type BorderPaddingText = Record<BorderEdge, string>;

/** Typed empty seed for the padding-text ref. */
export function emptyBorderPaddingText(): BorderPaddingText {
  return { top: "", bottom: "", left: "", right: "" };
}

/** Typed pass-through for a padding edge coming from a `.vue` template's
 * plain string. */
export function asBorderEdge(value: string): BorderEdge {
  return value as BorderEdge;
}

/** Whether an event target is a checked checkbox — the `.vue` templates
 * cannot narrow `$event.target` themselves (espree, see AppShell.vue). */
export function isEventTargetChecked(target: EventTarget | null): boolean {
  return target instanceof HTMLInputElement && target.checked;
}

/** Drops a `borders` object that no longer paints anything, so the model
 * never carries an empty husk (which would break the merge comparison
 * against a genuinely unbordered neighbour). Padding alone is not worth
 * keeping — it is meaningless without an edge to space. */
export function normaliseParagraphBorders(borders: ParagraphBorders | undefined): ParagraphBorders | undefined {
  if (!borders) {
    return undefined;
  }
  const painted =
    borderSideVisible(borders.top) ||
    borderSideVisible(borders.bottom) ||
    borderSideVisible(borders.left) ||
    borderSideVisible(borders.right) ||
    borderSideVisible(borders.between);
  return painted ? borders : undefined;
}

/** Typed `undefined` seed for a `ParagraphBorders` field in an espree-parsed
 * `.vue` script's default parameter (see commands.noSelection). */
export function noParagraphBorders(): ParagraphBorders | null {
  return null;
}

/** Typed `null` seed for the shading fill, likewise. */
export function noShadingFill(): string | null {
  return null;
}
