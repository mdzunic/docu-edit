// Cell rendering (SPEC-117 "Rendering", "Border CSS and weights" and
// "Overflow"): everything the sheet's cell and border layers draw, decided
// here as plain data so it is unit-testable without mounting a component.
// SheetCellLayer.vue applies it — it wraps the resolved family in SPEC-048's
// fontStack() (app layer; this directory may not import it) and positions
// the elements. Nothing here measures a font or touches the DOM.
//
// - cellCss: a style → the declarations of a cell's box (`.sheet-cell`, a
//   column flexbox: the fill, horizontal alignment as `text-align`, vertical
//   alignment as `justify-content`) and of its text element
//   (`.sheet-cell-text`: font, colour, decoration, padding with the indent,
//   white-space, rotation). Every colour resolves through core's
//   resolveColor — the single resolution point — and automatic is black.
// - resolveEdges: the border layer's segments, one per visible edge,
//   resolved once between the two cells sharing it (the heavier style wins
//   by the pinned rank; a tie goes to the right/bottom cell), merged
//   interiors suppressed, hidden rows and columns skipped.
// - overflowExtent / overflowAlignment / offscreenSpillCandidate: where a
//   string may spill into empty neighbours, found by binary search over the
//   row's sorted cells and never more than OVERFLOW_SCAN_COLS columns away.
// - estimateTextWidth: the per-character estimate shrink-to-fit and fill use
//   (and SPEC-118's #### decisions) — an estimate, never a measurement.
import { findRow } from "./address";
import type { GridRange } from "./address";
import { CELL_PADDING_PX, INDENT_PX_PER_LEVEL } from "./geometry";
import type { GridMetrics } from "./geometry";
import { generalText } from "./general";
import { mergeAt, mergesIn } from "./merges";
import type { MergeIndex } from "./merges";
import { MAX_COLS, MAX_ROWS } from "./model";
import type {
  Cell,
  CellAlignment,
  CellBorderEdge,
  CellFill,
  CellFont,
  CellStyle,
  CellValue,
  Sheet,
  SheetBorderStyle,
} from "./model";
import { hexToRgb, resolveColor, resolveFont, rgbToHex } from "../core/theme";
import type { DocumentTheme, ModelColor } from "../core/theme";

/** CSS declarations: kebab-case property → value. */
export type CssDeclarations = Record<string, string>;

/** The automatic colour — of a font, a border edge and a pattern's foreground. */
export const AUTOMATIC_COLOR = "#000000";
/** A pattern fill's background when the file names none. */
export const PATTERN_BACKGROUND = "#ffffff";
/** The base size a superscript/subscript cell scales when its style names
 * none (Calibri 11, Excel's Normal style). */
export const DEFAULT_FONT_SIZE_PT = 11;
/** Superscript and subscript text is drawn at this fraction of its base size. */
export const SCRIPT_SCALE = 0.7;

/** A number as CSS text: at most three decimals, no float noise (7.7, not
 * 7.699999999999999). */
function num(value: number): string {
  return String(Math.round(value * 1000) / 1000);
}

// --- values and alignment ----------------------------------------------------

/** What a cell holds, for General alignment and overflow. */
export type ValueKind = "number" | "string" | "boolean" | "error" | "empty";

/** The kind of a cell value; `""` is a string (it blocks overflow like any
 * value), only `null` is empty. */
export function valueKind(value: CellValue): ValueKind {
  if (value === null) {
    return "empty";
  }
  if (typeof value === "number") {
    return "number";
  }
  if (typeof value === "string") {
    return "string";
  }
  if (typeof value === "boolean") {
    return "boolean";
  }
  return "error";
}

/** A horizontal alignment with General resolved. */
export type ResolvedHorizontal = Exclude<NonNullable<CellAlignment["horizontal"]>, "general">;

/** The horizontal alignment a cell renders with: the style's own, or for
 * General (and absent) numbers right, booleans and errors centred, strings
 * (and empty cells) left. */
export function resolvedHorizontal(style: CellStyle, kind: ValueKind): ResolvedHorizontal {
  const horizontal = style.alignment?.horizontal;
  if (horizontal !== undefined && horizontal !== "general") {
    return horizontal;
  }
  if (kind === "number") {
    return "right";
  }
  return kind === "boolean" || kind === "error" ? "center" : "left";
}

const TEXT_ALIGN: Record<ResolvedHorizontal, string> = {
  left: "left",
  center: "center",
  right: "right",
  fill: "left",
  justify: "justify",
  centerContinuous: "center",
  distributed: "justify",
};

/** Vertical alignment as the column flexbox's `justify-content`; absent =
 * bottom. Justify and distributed wrap from the top. */
const JUSTIFY_CONTENT: Record<NonNullable<CellAlignment["vertical"]>, string> = {
  top: "flex-start",
  center: "center",
  bottom: "flex-end",
  justify: "flex-start",
  distributed: "flex-start",
};

/** Whether an alignment wraps its text inside the cell: `wrap`, and the
 * justified/distributed alignments (which Excel always wraps). */
function wraps(alignment: CellAlignment | undefined): boolean {
  if (!alignment) {
    return false;
  }
  return (
    alignment.wrap === true ||
    alignment.horizontal === "justify" ||
    alignment.horizontal === "distributed" ||
    alignment.vertical === "justify" ||
    alignment.vertical === "distributed"
  );
}

// --- fonts -------------------------------------------------------------------

/** `text-decoration` for a font: underline (double and doubleAccounting →
 * `underline double`) and strike; undefined when it has neither. */
function textDecoration(font: CellFont): string | undefined {
  const lines: string[] = [];
  if (font.underline) {
    lines.push("underline");
  }
  if (font.strike) {
    lines.push("line-through");
  }
  if (lines.length === 0) {
    return undefined;
  }
  if (font.underline === "double" || font.underline === "doubleAccounting") {
    lines.push("double");
  }
  return lines.join(" ");
}

/** A font's declarations. For a cell (`cell` true) the colour is always set
 * (automatic → black) and a script's base size defaults to 11 pt; for a
 * rich-text run only the properties present are emitted, a script without a
 * size scaling the cell's size (`0.7em`). */
function fontCss(font: CellFont, theme: DocumentTheme, cell: boolean): CssDeclarations {
  const css: CssDeclarations = {};
  if (font.family !== undefined) {
    const family = resolveFont(font.family, theme);
    if (family.trim() !== "") {
      css["font-family"] = family;
    }
  }
  if (font.vertAlign !== undefined) {
    css["vertical-align"] = font.vertAlign === "superscript" ? "super" : "sub";
    const base = font.sizePt ?? (cell ? DEFAULT_FONT_SIZE_PT : undefined);
    css["font-size"] = base === undefined ? `${SCRIPT_SCALE}em` : `${num(base * SCRIPT_SCALE)}pt`;
  } else if (font.sizePt !== undefined) {
    css["font-size"] = `${num(font.sizePt)}pt`;
  }
  if (font.bold) {
    css["font-weight"] = "bold";
  }
  if (font.italic) {
    css["font-style"] = "italic";
  }
  const decoration = textDecoration(font);
  if (decoration !== undefined) {
    css["text-decoration"] = decoration;
  }
  if (font.color !== undefined) {
    css.color = resolveColor(font.color, theme);
  } else if (cell) {
    css.color = AUTOMATIC_COLOR;
  }
  return css;
}

/** A rich-text run span's declarations — only the properties its font
 * carries, each overriding the cell font (SPEC-117 "Rendering"). */
export function runCss(font: CellFont, theme: DocumentTheme): CssDeclarations {
  return fontCss(font, theme, false);
}

// --- fills -------------------------------------------------------------------

/** The solid patterns: fg mixed over bg at this fraction. */
const PATTERN_MIX: Record<string, number> = {
  gray0625: 0.0625,
  gray125: 0.125,
  lightGray: 0.25,
  mediumGray: 0.5,
  darkGray: 0.75,
};

/** The line patterns' stripe angles; a grid or trellis layers two. */
const STRIPE_ANGLES: Record<string, readonly number[]> = {
  Horizontal: [0],
  Vertical: [90],
  Down: [45],
  Up: [135],
  Grid: [0, 90],
  Trellis: [45, 135],
};

const STRIPE_PATTERN = /^(light|dark)(Horizontal|Vertical|Down|Up|Grid|Trellis)$/;
const HEX_COLOR = /^#[0-9a-f]{6}$/i;

/** fg over bg at `amount` (0..1), channel by channel in sRGB. */
function mixColors(fg: string, bg: string, amount: number): string {
  if (!HEX_COLOR.test(fg) || !HEX_COLOR.test(bg)) {
    return fg;
  }
  const [fr, fgreen, fb] = hexToRgb(fg);
  const [br, bgreen, bb] = hexToRgb(bg);
  return rgbToHex([
    fr * amount + br * (1 - amount),
    fgreen * amount + bgreen * (1 - amount),
    fb * amount + bb * (1 - amount),
  ]);
}

/** One stripe layer: an L px line of fg every 4 px over `gap`. */
function stripes(angle: number, line: number, fg: string, gap: string): string {
  return `repeating-linear-gradient(${angle}deg, ${fg} 0 ${line}px, ${gap} ${line}px 4px)`;
}

function patternCss(
  pattern: string,
  foreground: ModelColor | undefined,
  background: ModelColor | undefined,
  theme: DocumentTheme,
): CssDeclarations {
  const fg = foreground === undefined ? AUTOMATIC_COLOR : resolveColor(foreground, theme);
  const bg = background === undefined ? PATTERN_BACKGROUND : resolveColor(background, theme);
  if (pattern === "solid") {
    return { "background-color": fg };
  }
  const mix = PATTERN_MIX[pattern];
  if (mix !== undefined) {
    return { "background-color": mixColors(fg, bg, mix) };
  }
  const match = STRIPE_PATTERN.exec(pattern);
  if (!match) {
    return {};
  }
  const line = match[1] === "light" ? 1 : 2;
  const angles = STRIPE_ANGLES[match[2]!]!;
  // Layered patterns: every layer but the bottom one keeps its gaps
  // transparent so the layers beneath show through.
  const layers = angles.map((angle, i) => stripes(angle, line, fg, i === angles.length - 1 ? bg : "transparent"));
  return { "background-color": bg, "background-image": layers.join(", ") };
}

function gradientCss(gradient: Extract<CellFill, { kind: "gradient" }>, theme: DocumentTheme): CssDeclarations {
  const stops = gradient.stops;
  if (stops.length === 0) {
    return {};
  }
  if (stops.length === 1) {
    // CSS needs two stops; one is a solid colour.
    return { "background-color": resolveColor(stops[0]!.color, theme) };
  }
  const list = stops.map((stop) => `${resolveColor(stop.color, theme)} ${num(stop.position * 100)}%`).join(", ");
  if (gradient.type === "path") {
    const x = (((gradient.left ?? 0) + (gradient.right ?? 0)) / 2) * 100;
    const y = (((gradient.top ?? 0) + (gradient.bottom ?? 0)) / 2) * 100;
    return { "background-image": `radial-gradient(at ${num(x)}% ${num(y)}%, ${list})` };
  }
  return { "background-image": `linear-gradient(${num((gradient.degree ?? 0) + 90)}deg, ${list})` };
}

/** A cell fill's declarations (`background-color` / `background-image`):
 * solid fills as the colour; patterns per the pinned table — the grays as a
 * solid sRGB mix of fg over bg, the line patterns as repeating gradients
 * (grid and trellis layered) over the bg colour, fg defaulting to #000000 and
 * bg to #ffffff; gradients as CSS gradients (linear at degree + 90, path as a
 * radial gradient at the focus). Empty for no fill — the cell stays
 * transparent and the gridlines show. */
export function fillCss(fill: CellFill | undefined, theme: DocumentTheme): CssDeclarations {
  if (fill === undefined) {
    return {};
  }
  switch (fill.kind) {
    case "solid":
      return { "background-color": resolveColor(fill.color, theme) };
    case "pattern":
      return patternCss(fill.pattern, fill.fg, fill.bg, theme);
    case "gradient":
      return gradientCss(fill, theme);
    default:
      return {};
  }
}

// --- the cell ----------------------------------------------------------------

/** The text element's padding: CELL_PADDING_PX each side, plus the indent on
 * the aligned side (left for left/General-left/fill, right for right, both
 * for distributed — Excel's "Distributed (Indent)"). */
function paddingCss(horizontal: ResolvedHorizontal, indent: number | undefined): string {
  const extra = indent !== undefined && indent > 0 ? indent * INDENT_PX_PER_LEVEL : 0;
  const onLeft = horizontal === "left" || horizontal === "fill" || horizontal === "distributed";
  const onRight = horizontal === "right" || horizontal === "distributed";
  const left = CELL_PADDING_PX + (onLeft ? extra : 0);
  const right = CELL_PADDING_PX + (onRight ? extra : 0);
  return left === right ? `0 ${num(left)}px` : `0 ${num(right)}px 0 ${num(left)}px`;
}

/** A cell's declarations for a value of `kind` (SPEC-117 "Rendering"):
 *
 * - `box` (`.sheet-cell`, a column flexbox): the fill (fillCss), `text-align`
 *   (left/center/right/justify — fill → left, centerContinuous → center,
 *   distributed → justify with `text-align-last: justify`) and the vertical
 *   alignment as `justify-content` (top → flex-start, center → center,
 *   bottom/absent → flex-end, justify/distributed → flex-start).
 * - `text` (`.sheet-cell-text`): `font-family` (the resolved family name,
 *   unquoted, omitted when the style names none — the component wraps it in
 *   fontStack()), `font-size` in pt (omitted when absent), `font-weight`,
 *   `font-style`, `text-decoration`, `color` (automatic → #000000),
 *   superscript/subscript as `vertical-align` with the size scaled by 0.7,
 *   `padding` (`0 2px` plus the indent on the aligned side), `white-space`
 *   (`pre`; `pre-wrap` for wrapped, justified and distributed text, clipped —
 *   rows never grow) and rotation (1–90 → `rotate(−θdeg)`, 91–180 →
 *   `rotate((θ − 90)deg)` about the centre; 255 → vertical stacked text).
 *
 * Shrink-to-fit and fill need the text and the width, so the component
 * applies shrinkScale and fillRepeat itself. */
export function cellCss(
  style: CellStyle,
  theme: DocumentTheme,
  kind: ValueKind = "string",
): { box: CssDeclarations; text: CssDeclarations } {
  const alignment = style.alignment;
  const horizontal = resolvedHorizontal(style, kind);
  const box: CssDeclarations = fillCss(style.fill, theme);
  box["text-align"] = TEXT_ALIGN[horizontal] ?? "left";
  if (horizontal === "distributed") {
    box["text-align-last"] = "justify";
  }
  box["justify-content"] = JUSTIFY_CONTENT[alignment?.vertical ?? "bottom"] ?? "flex-end";

  const text = fontCss(style.font ?? {}, theme, true);
  text.padding = paddingCss(horizontal, alignment?.indent);
  text["white-space"] = wraps(alignment) ? "pre-wrap" : "pre";
  const rotation = alignment?.textRotation;
  if (rotation === 255) {
    text["writing-mode"] = "vertical-lr";
    text["text-orientation"] = "upright";
  } else if (rotation !== undefined && rotation >= 1 && rotation <= 180) {
    text.transform = `rotate(${num(rotation <= 90 ? -rotation : rotation - 90)}deg)`;
    text["transform-origin"] = "center";
  }
  return { box, text };
}

// --- borders -----------------------------------------------------------------

/** A border style's rendering and its weight in the shared-edge rule (rank 1
 * is the heaviest). */
export interface BorderCss {
  width: number; // px
  lineStyle: "solid" | "double" | "dashed" | "dotted";
  opacity?: number;
  rank: number;
}

/** The pinned "Border CSS and weights" table, heaviest first. */
export const BORDER_STYLES: Readonly<Record<SheetBorderStyle, BorderCss>> = {
  thick: { rank: 1, width: 3, lineStyle: "solid" },
  double: { rank: 2, width: 3, lineStyle: "double" },
  medium: { rank: 3, width: 2, lineStyle: "solid" },
  mediumDashed: { rank: 4, width: 2, lineStyle: "dashed" },
  mediumDashDot: { rank: 5, width: 2, lineStyle: "dashed" },
  slantDashDot: { rank: 6, width: 2, lineStyle: "dashed" },
  mediumDashDotDot: { rank: 7, width: 2, lineStyle: "dotted" },
  thin: { rank: 8, width: 1, lineStyle: "solid" },
  dashed: { rank: 9, width: 1, lineStyle: "dashed" },
  dashDot: { rank: 10, width: 1, lineStyle: "dashed" },
  dashDotDot: { rank: 11, width: 1, lineStyle: "dotted" },
  dotted: { rank: 12, width: 1, lineStyle: "dotted" },
  hair: { rank: 13, width: 1, lineStyle: "dotted", opacity: 0.6 },
};

/** A border style's CSS and rank (an unknown style draws as thin). */
export function borderCss(style: SheetBorderStyle): BorderCss {
  return BORDER_STYLES[style] ?? BORDER_STYLES.thin;
}

/** A border style's rank in the shared-edge rule: 1 (thick) … 13 (hair). */
export function borderRank(style: SheetBorderStyle): number {
  return borderCss(style).rank;
}

/** A CSS `border` shorthand for a style and a resolved colour
 * ("3px double #000000"); the hair style's 60% opacity is the element's. */
export function borderShorthand(style: SheetBorderStyle, color: string): string {
  const css = borderCss(style);
  return `${css.width}px ${css.lineStyle} ${color}`;
}

/** A border edge's resolved colour; automatic is black. */
function edgeColor(edge: CellBorderEdge, theme: DocumentTheme): string {
  return edge.color === undefined ? AUTOMATIC_COLOR : resolveColor(edge.color, theme);
}

/** One resolved border edge to draw, in unscaled sheet px: a horizontal
 * segment runs from (x, y) to (x + length, y), a vertical one from (x, y) to
 * (x, y + length). */
export interface EdgeSegment {
  x: number;
  y: number;
  length: number;
  orientation: "horizontal" | "vertical";
  style: SheetBorderStyle;
  color: string; // resolved #rrggbb; automatic = #000000
}

/** The queries of one axis of GridMetrics. */
interface AxisQueries {
  start(i: number): number;
  size(i: number): number;
  at(pos: number): number;
  total: number;
}

function rowQueries(metrics: GridMetrics): AxisQueries {
  return {
    start: (r) => metrics.rowTop(r),
    size: (r) => metrics.rowHeight(r),
    at: (y) => metrics.rowAt(y),
    total: metrics.totalHeight(MAX_ROWS),
  };
}

function colQueries(metrics: GridMetrics): AxisQueries {
  return {
    start: (c) => metrics.colLeft(c),
    size: (c) => metrics.colWidth(c),
    at: (x) => metrics.colAt(x),
    total: metrics.totalWidth(MAX_COLS),
  };
}

/** The visible (non-zero-size) items in [first, last], in order — a hidden
 * run is skipped in one query, however long. */
function visibleItems(axis: AxisQueries, first: number, last: number): number[] {
  const out: number[] = [];
  if (first > last) {
    return out;
  }
  let pos = axis.start(Math.max(0, first));
  while (pos < axis.total) {
    const i = axis.at(pos);
    if (i > last) {
      break;
    }
    out.push(i);
    const next = axis.start(i) + axis.size(i);
    if (next <= pos) {
      break;
    }
    pos = next;
  }
  return out;
}

/** The nearest visible item before `i`, or undefined at the grid's edge. */
function visibleBefore(axis: AxisQueries, i: number): number | undefined {
  const start = axis.start(i);
  return start > 0 ? axis.at(start - 1) : undefined;
}

/** The nearest visible item after `i`, or undefined at the grid's edge. */
function visibleAfter(axis: AxisQueries, i: number): number | undefined {
  const end = axis.start(i) + axis.size(i);
  return end < axis.total ? axis.at(end) : undefined;
}

/** The elements of an ascending list within [lo, hi]. */
function within(list: readonly number[], lo: number, hi: number): number[] {
  let a = 0;
  let b = list.length;
  while (a < b) {
    const mid = (a + b) >>> 1;
    if (list[mid]! < lo) {
      a = mid + 1;
    } else {
      b = mid;
    }
  }
  const out: number[] = [];
  for (let i = a; i < list.length && list[i]! <= hi; i += 1) {
    out.push(list[i]!);
  }
  return out;
}

/** Consecutive pairs of a list framed by an optional item on each side:
 * (before, l0), (l0, l1), …, (ln, after). */
function boundaries(
  list: readonly number[],
  before: number | undefined,
  after: number | undefined,
): Array<[number | undefined, number | undefined]> {
  const out: Array<[number | undefined, number | undefined]> = [];
  let previous = before;
  for (const item of list) {
    out.push([previous, item]);
    previous = item;
  }
  out.push([previous, after]);
  return out;
}

type Side = "top" | "right" | "bottom" | "left";

/** A grid position as one number (row-major; exact below 2^53). */
function key(row: number, col: number): number {
  return row * MAX_COLS + col;
}

/** The border layer's segments for the visible cells of `view` (SPEC-117
 * "Border CSS and weights"): one per visible resolved edge, including the
 * edges on the view's outer boundary, which resolve against the nearest
 * visible neighbour just outside it.
 *
 * Shared-edge rule: the two cells' facing edges compete by rank — the
 * heavier (lower rank) wins; a tie (the same style) goes to the right or
 * bottom cell's edge, i.e. the neighbour's `left`/`top`, colour included;
 * one side absent → the other. Merges: interior edges draw nothing; an
 * outer edge comes from the member cell on that edge of the range. Hidden
 * rows and columns produce no segment (the visible cells either side of a
 * hidden run face each other). Adjacent equal segments stay separate.
 *
 * `styleAt` is the cell's effective style (cellStyle(model,
 * effectiveStyleId(sheet, row, col))); it is called at most once per
 * position. `sheet` is the sheet `view` belongs to — every per-cell fact the
 * pass needs arrives through `styleAt`, `metrics` and `merges`. */
export function resolveEdges(
  view: GridRange,
  sheet: Sheet,
  styleAt: (row: number, col: number) => CellStyle,
  metrics: GridMetrics,
  merges: MergeIndex,
  theme: DocumentTheme,
): EdgeSegment[] {
  const rowAxis = rowQueries(metrics);
  const colAxis = colQueries(metrics);
  const rows = visibleItems(rowAxis, view.top, Math.min(view.bottom, MAX_ROWS - 1));
  const cols = visibleItems(colAxis, view.left, Math.min(view.right, MAX_COLS - 1));
  if (rows.length === 0 || cols.length === 0) {
    return [];
  }
  const above = visibleBefore(rowAxis, rows[0]!);
  const below = visibleAfter(rowAxis, rows.at(-1)!);
  const before = visibleBefore(colAxis, cols[0]!);
  const after = visibleAfter(colAxis, cols.at(-1)!);
  const allRows = [...(above === undefined ? [] : [above]), ...rows, ...(below === undefined ? [] : [below])];
  const allCols = [...(before === undefined ? [] : [before]), ...cols, ...(after === undefined ? [] : [after])];

  // The merge of every position the edges look at, from one range query.
  const mergeOf = new Map<number, GridRange>();
  const extent: GridRange = { top: allRows[0]!, bottom: allRows.at(-1)!, left: allCols[0]!, right: allCols.at(-1)! };
  for (const merge of mergesIn(merges, extent)) {
    const mergeRows = within(allRows, merge.top, merge.bottom);
    const mergeCols = within(allCols, merge.left, merge.right);
    for (const row of mergeRows) {
      for (const col of mergeCols) {
        mergeOf.set(key(row, col), merge);
      }
    }
  }
  const styles = new Map<number, CellStyle>();
  const styleOf = (row: number, col: number): CellStyle => {
    const k = key(row, col);
    let style = styles.get(k);
    if (style === undefined) {
      style = styleAt(row, col);
      styles.set(k, style);
    }
    return style;
  };
  /** A cell's edge on `side` — for a merge member, the member on that edge
   * of the range in the same row/column. */
  const facing = (row: number, col: number, side: Side): CellBorderEdge | undefined => {
    const merge = mergeOf.get(key(row, col));
    if (merge !== undefined) {
      if (side === "top") {
        row = merge.top;
      } else if (side === "bottom") {
        row = merge.bottom;
      } else if (side === "left") {
        col = merge.left;
      } else {
        col = merge.right;
      }
    }
    return styleOf(row, col).borders?.[side];
  };
  const sameMerge = (r1: number, c1: number, r2: number, c2: number): boolean => {
    const merge = mergeOf.get(key(r1, c1));
    return merge !== undefined && merge === mergeOf.get(key(r2, c2));
  };
  const winner = (first: CellBorderEdge | undefined, second: CellBorderEdge | undefined): CellBorderEdge | undefined => {
    if (first === undefined) {
      return second;
    }
    if (second === undefined) {
      return first;
    }
    return borderRank(first.style) < borderRank(second.style) ? first : second;
  };

  const segments: EdgeSegment[] = [];
  for (const [upper, lower] of boundaries(rows, above, below)) {
    const y = lower !== undefined ? metrics.rowTop(lower) : metrics.rowTop(upper!) + metrics.rowHeight(upper!);
    for (const col of cols) {
      if (upper !== undefined && lower !== undefined && sameMerge(upper, col, lower, col)) {
        continue;
      }
      const edge = winner(
        upper === undefined ? undefined : facing(upper, col, "bottom"),
        lower === undefined ? undefined : facing(lower, col, "top"),
      );
      if (edge !== undefined) {
        segments.push({
          x: metrics.colLeft(col),
          y,
          length: metrics.colWidth(col),
          orientation: "horizontal",
          style: edge.style,
          color: edgeColor(edge, theme),
        });
      }
    }
  }
  for (const [leftCol, rightCol] of boundaries(cols, before, after)) {
    const x = rightCol !== undefined ? metrics.colLeft(rightCol) : metrics.colLeft(leftCol!) + metrics.colWidth(leftCol!);
    for (const row of rows) {
      if (leftCol !== undefined && rightCol !== undefined && sameMerge(row, leftCol, row, rightCol)) {
        continue;
      }
      const edge = winner(
        leftCol === undefined ? undefined : facing(row, leftCol, "right"),
        rightCol === undefined ? undefined : facing(row, rightCol, "left"),
      );
      if (edge !== undefined) {
        segments.push({
          x,
          y: metrics.rowTop(row),
          length: metrics.rowHeight(row),
          orientation: "vertical",
          style: edge.style,
          color: edgeColor(edge, theme),
        });
      }
    }
  }
  return segments;
}

/** A cell's diagonal border(s), drawn as inline SVG lines across the cell
 * (or merged range): `up` runs bottom-left → top-right, `down` top-left →
 * bottom-right. `width`, `lineStyle` and `opacity` come from the border
 * table; `dasharray` is the SVG `stroke-dasharray` of a dashed or dotted
 * style (a double style is drawn by the component as it sees fit — a solid
 * line of `width` at its simplest). */
export interface DiagonalLines {
  up: boolean;
  down: boolean;
  width: number;
  lineStyle: BorderCss["lineStyle"];
  dasharray?: string;
  opacity?: number;
  color: string; // resolved #rrggbb; automatic = #000000
}

/** The diagonal(s) a style draws, or undefined when it draws none (no
 * diagonal edge, or one with neither `up` nor `down`). */
export function diagonalLines(style: CellStyle, theme: DocumentTheme): DiagonalLines | undefined {
  const diagonal = style.borders?.diagonal;
  if (diagonal === undefined || (diagonal.up !== true && diagonal.down !== true)) {
    return undefined;
  }
  const css = borderCss(diagonal.style);
  const lines: DiagonalLines = {
    up: diagonal.up === true,
    down: diagonal.down === true,
    width: css.width,
    lineStyle: css.lineStyle,
    color: edgeColor(diagonal, theme),
  };
  if (css.lineStyle === "dashed") {
    lines.dasharray = `${3 * css.width} ${css.width}`;
  } else if (css.lineStyle === "dotted") {
    lines.dasharray = `${css.width} ${css.width}`;
  }
  if (css.opacity !== undefined) {
    lines.opacity = css.opacity;
  }
  return lines;
}

// --- overflow ----------------------------------------------------------------

/** How far the overflow pass scans for a blocking cell, in columns. */
export const OVERFLOW_SCAN_COLS = 256;

/** The index of the first cell in an ascending row whose column is ≥ `col`. */
function firstCellAtOrAfter(cells: readonly Cell[], col: number): number {
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (cells[mid]!.col < col) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** The first blocking column in (col, limit], or limit + 1: a stored value
 * (`""` included) or a merged cell. */
function spillStopRight(cells: readonly Cell[], row: number, col: number, limit: number, merges: MergeIndex): number {
  let stop = limit + 1;
  for (let i = firstCellAtOrAfter(cells, col + 1); i < cells.length && cells[i]!.col < stop; i += 1) {
    if (cells[i]!.value !== null) {
      stop = cells[i]!.col;
      break;
    }
  }
  if (stop > col + 1) {
    for (const merge of mergesIn(merges, { top: row, bottom: row, left: col + 1, right: stop - 1 })) {
      stop = Math.min(stop, Math.max(merge.left, col + 1));
    }
  }
  return stop;
}

/** The last blocking column in [limit, col), or limit − 1. */
function spillStopLeft(cells: readonly Cell[], row: number, col: number, limit: number, merges: MergeIndex): number {
  let stop = limit - 1;
  for (let i = firstCellAtOrAfter(cells, col) - 1; i >= 0 && cells[i]!.col > stop; i -= 1) {
    if (cells[i]!.value !== null) {
      stop = cells[i]!.col;
      break;
    }
  }
  if (stop < col - 1) {
    for (const merge of mergesIn(merges, { top: row, bottom: row, left: stop + 1, right: col - 1 })) {
      stop = Math.max(stop, Math.min(merge.right, col - 1));
    }
  }
  return stop;
}

/** The room, in px, a string at (row, col) may spill into on each side
 * (SPEC-117 "Overflow"): the empty neighbours up to the nearest blocking cell
 * — a stored cell with a value (`""` included), a merged cell, the frozen
 * column boundary (frozen text never spills into the unfrozen columns, nor
 * the reverse) or the grid's edge — scanning at most OVERFLOW_SCAN_COLS
 * columns. Hidden columns add 0 px and block only if they hold a value.
 * Left-aligned text spills right only, right-aligned left only, centred
 * text min(leftRoom, rightRoom) both ways. A merged or hidden cell never
 * spills ({ 0, 0 }). */
export function overflowExtent(
  sheet: Sheet,
  row: number,
  col: number,
  align: "left" | "right" | "center",
  metrics: GridMetrics,
  merges: MergeIndex,
  frozenCols = 0,
): { left: number; right: number } {
  if (metrics.colWidth(col) === 0 || mergeAt(merges, row, col) !== undefined) {
    return { left: 0, right: 0 };
  }
  const cells = findRow(sheet, row)?.cells ?? [];
  let left = 0;
  let right = 0;
  if (align !== "right") {
    let limit = Math.min(MAX_COLS - 1, col + OVERFLOW_SCAN_COLS);
    if (col < frozenCols) {
      limit = Math.min(limit, frozenCols - 1);
    }
    right = metrics.colLeft(spillStopRight(cells, row, col, limit, merges)) - metrics.colLeft(col + 1);
  }
  if (align !== "left") {
    let limit = Math.max(0, col - OVERFLOW_SCAN_COLS);
    if (frozenCols > 0 && col >= frozenCols) {
      limit = Math.max(limit, frozenCols);
    }
    left = metrics.colLeft(col) - metrics.colLeft(spillStopLeft(cells, row, col, limit, merges) + 1);
  }
  if (align === "center") {
    const room = Math.min(left, right);
    return { left: room, right: room };
  }
  return { left, right };
}

/** The direction a cell's text may spill, or null when it may not:
 * non-strings (and `""`) never spill, nor do wrapped, filled, justified or
 * distributed, rotated (any textRotation but 0, 255 included) and
 * shrink-to-fit text. General and left → "left" (spills right), right →
 * "right" (spills left), center and centerContinuous → "center". */
export function overflowAlignment(style: CellStyle, value: CellValue): "left" | "right" | "center" | null {
  if (typeof value !== "string" || value === "") {
    return null;
  }
  const alignment = style.alignment;
  if (alignment !== undefined) {
    if (wraps(alignment) || alignment.shrinkToFit === true) {
      return null;
    }
    if (alignment.textRotation !== undefined && alignment.textRotation !== 0) {
      return null;
    }
  }
  switch (alignment?.horizontal) {
    case undefined:
    case "general":
    case "left":
      return "left";
    case "right":
      return "right";
    case "center":
    case "centerContinuous":
      return "center";
    default:
      return null; // fill, justify, distributed
  }
}

/** Whether a cell's text could spill: a non-empty string. */
function spillsText(cell: Cell): boolean {
  return typeof cell.value === "string" && cell.value !== "";
}

/** The stored cells outside a visible row's column window [firstCol,
 * lastCol] whose text could paint into it ("a string anchored left of the
 * visible window still paints into it"): on each side the nearest stored
 * cell holding a value within OVERFLOW_SCAN_COLS columns, when that value is
 * a non-empty string — a nearer value would block anything beyond it.
 * `left` is the candidate for left-aligned and centred text, `right` for
 * right-aligned and centred text; the component checks their alignment and
 * overflowExtent. */
export function offscreenSpillCandidate(
  sheet: Sheet,
  row: number,
  firstCol: number,
  lastCol: number,
): { left?: Cell; right?: Cell } {
  const out: { left?: Cell; right?: Cell } = {};
  const cells = findRow(sheet, row)?.cells;
  if (cells === undefined || cells.length === 0) {
    return out;
  }
  for (let i = firstCellAtOrAfter(cells, firstCol) - 1; i >= 0 && cells[i]!.col >= firstCol - OVERFLOW_SCAN_COLS; i -= 1) {
    const cell = cells[i]!;
    if (cell.value !== null) {
      if (spillsText(cell)) {
        out.left = cell;
      }
      break;
    }
  }
  for (
    let i = firstCellAtOrAfter(cells, lastCol + 1);
    i < cells.length && cells[i]!.col <= lastCol + OVERFLOW_SCAN_COLS;
    i += 1
  ) {
    const cell = cells[i]!;
    if (cell.value !== null) {
      if (spillsText(cell)) {
        out.right = cell;
      }
      break;
    }
  }
  return out;
}

// --- text width, shrink and fill ---------------------------------------------

const NARROW_CHARACTERS: ReadonlySet<string> = new Set([".", ",", ":", " "]);

/** A text's estimated width in px: Σ per-character weights × MDW, weight 0.5
 * for `.` `,` `:` and space and 1 for everything else — an estimate, never a
 * measurement (SPEC-117 "Alignment"; SPEC-118 reuses it). */
export function estimateTextWidth(text: string, mdw: number): number {
  let units = 0;
  for (const character of text) {
    units += NARROW_CHARACTERS.has(character) ? 0.5 : 1;
  }
  return units * mdw;
}

/** The shrink-to-fit font scale: min(1, available / estimated width) — 1 for
 * text with no width. */
export function shrinkScale(text: string, availablePx: number, mdw: number): number {
  const width = estimateTextWidth(text, mdw);
  if (width <= 0) {
    return 1;
  }
  return Math.min(1, Math.max(0, availablePx) / width);
}

/** A fill-aligned cell's text: repeated ceil(width / text width) times (at
 * least once) so it fills the cell, the excess clipped by the component. */
export function fillRepeat(text: string, widthPx: number, mdw: number): string {
  const unit = estimateTextWidth(text, mdw);
  if (unit <= 0 || !Number.isFinite(widthPx)) {
    return text;
  }
  return text.repeat(Math.max(1, Math.ceil(widthPx / unit)));
}

/** The text a cell shows: generalText(value) (SPEC-118 replaces it), except
 * that a numeric zero shows "" on a sheet with `showZeros: false`. */
export function displayText(value: CellValue, sheet: Pick<Sheet, "showZeros">): string {
  if (value === 0 && sheet.showZeros === false) {
    return "";
  }
  return generalText(value);
}
