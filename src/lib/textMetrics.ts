// Canvas-predicted paragraph metrics (SPEC-028): line wrapping computed with
// 2D-canvas text measurement instead of mounting the paragraph in the DOM.
// The bulk pass over a large document cannot afford to render every block
// just to measure it (mounting ~10k paragraphs costs seconds; measureText
// costs microseconds), so unrendered paragraphs are measured here and only
// rendered blocks use lib/pageMetrics' DOM path. Both must agree: the canvas
// shares the browser's text shaper, so word widths match the inline layout
// engine, and the greedy wrap below reproduces CSS `white-space: pre-wrap`
// breaking (validated against DOM line boxes across mixed formatting,
// punctuation, repeated spaces and accented text). Vertical geometry
// (`line-height: normal` is font-metric dependent and not exposed by the
// canvas API) comes from a one-per-font hidden DOM probe the measurer
// caches.
//
// Framework-free like the rest of src/lib; the canvas measurer degrades to
// null where the environment cannot measure text (jsdom), in which case
// callers keep their estimates and the DOM path.
import type { Paragraph, RunFormatting, TabStop } from "./documentModel";
import { listIndentCm } from "./core/lists";
import { type BlockMetrics, type LineBox } from "./core/pageMetrics";
import { PX_PER_CM } from "./core/units";
import { layoutParagraphTabs, stopsToPx, type ParagraphTabLayout } from "./tabLayout";

const PX_PER_PT = 96 / 72;

/** Matches lib/pagination's EPS spirit: sub-pixel noise between the canvas
 * and the layout engine must not flip a fit decision. */
const FIT_EPS = 0.02;

/** CSS `tab-size` default: tab stops every 8 space advances. */
const TAB_SIZE_SPACES = 8;

/** A run's font resolved against the canvas's inherited base font — the
 * key both measurement backends (canvas width, DOM geometry probe) share. */
export interface ResolvedFont {
  family: string;
  sizePx: number;
  bold: boolean;
  italic: boolean;
}

/** One font's vertical layout geometry: the layout advance between line
 * tops (CSS line box height for `line-height: normal`) and the ink height
 * `Range.getClientRects` reports for a line — the same two numbers
 * lib/pageMetrics reads off rendered lines. */
export interface LineGeometry {
  advancePx: number;
  inkHeightPx: number;
}

/** The measurement backend the wrap core runs against — canvas + DOM probe
 * in the app, a deterministic fake in unit tests. */
export interface TextMeasurer {
  widthOf(text: string, font: ResolvedFont): number;
  /** `line-height: normal` geometry for a font, or null where the
   * environment cannot measure (jsdom) — the wrap then reports failure. */
  normalGeometryOf(font: ResolvedFont): LineGeometry | null;
}

/** The canvas's inherited base font (the app reads it off the content root's
 * computed style once): what a run without explicit font formatting
 * renders in. */
export interface BaseFont {
  family: string;
  sizePx: number;
}

export interface MeasuredParagraph {
  metrics: BlockMetrics;
  /** Character offset each line starts at — the same offsets the DOM path
   * resolves with Range hit-testing, but known for free here because the
   * wrap decided them. */
  lineOffsets: number[];
}

/** The tab context a paragraph is measured/rendered against (SPEC-054): its
 * effective stops (styles.tabContextFor), the document default interval and
 * the decimal separator for its language. Absent leaves tabs on the browser's
 * default 8-space grid (the pre-SPEC-054 behaviour), used by callers that do
 * not resolve tab stops (unit tests, note measurement). */
export interface TabLayoutConfig {
  stops: TabStop[];
  defaultTabStopCm: number;
  decimalSeparator: string;
}

/** Builds the paragraph's font-boundary spans without vertical geometry —
 * enough for measureRange, which only reads font/start/end. Shared by the
 * wrap predictor and the standalone tab-layout helper. */
function fontSpansOf(paragraph: Paragraph, base: BaseFont): { text: string; spans: FontSpan[] } {
  const text = paragraph.runs.map((run) => run.text).join("");
  const spans: FontSpan[] = [];
  let offset = 0;
  for (const run of paragraph.runs) {
    if (run.text.length === 0) {
      continue;
    }
    spans.push({
      font: resolveRunFont(run.formatting, base),
      start: offset,
      end: offset + run.text.length,
      advancePx: 0,
      inkHeightPx: 0,
    });
    offset += run.text.length;
  }
  return { text, spans };
}

/** The tab geometry for a paragraph (SPEC-054): the same resolution the wrap
 * predictor uses, exposed for the DOM renderer so measured and rendered tab
 * advances are identical. */
export function paragraphTabLayout(
  paragraph: Paragraph,
  config: TabLayoutConfig,
  base: BaseFont,
  measurer: TextMeasurer,
): ParagraphTabLayout {
  const { text, spans } = fontSpansOf(paragraph, base);
  return layoutParagraphTabs(
    text,
    stopsToPx(config.stops),
    config.defaultTabStopCm * PX_PER_CM,
    config.decimalSeparator,
    (from, to) => measureRange(text, spans, from, to, measurer),
  );
}

function resolveRunFont(formatting: RunFormatting, base: BaseFont): ResolvedFont {
  let sizePx = formatting.fontSizePt ? formatting.fontSizePt * PX_PER_PT : base.sizePx;
  if (formatting.vertAlign) {
    // runStyle renders sub/superscript at 0.65 of the run's size (or 0.65em
    // of the inherited size).
    sizePx = formatting.fontSizePt ? formatting.fontSizePt * 0.65 * PX_PER_PT : base.sizePx * 0.65;
  }
  return {
    family: formatting.fontFamily ?? base.family,
    sizePx,
    bold: formatting.bold === true,
    italic: formatting.italic === true,
  };
}

interface FontSpan {
  font: ResolvedFont;
  start: number;
  end: number;
  /** Precomputed per-span vertical rollup — resolved once so the wrap's
   * per-token bookkeeping is pure arithmetic. */
  advancePx: number;
  inkHeightPx: number;
}

/** Measures `[from, to)` of the paragraph text, split at font boundaries so
 * each chunk is shaped in its own font — kerning inside a word is kept,
 * matching how the layout engine shapes per-styled-segment. */
function measureRange(
  text: string,
  spans: FontSpan[],
  from: number,
  to: number,
  measurer: TextMeasurer,
): number {
  let width = 0;
  for (const span of spans) {
    if (span.end <= from) {
      continue;
    }
    if (span.start >= to) {
      break;
    }
    const start = Math.max(span.start, from);
    const end = Math.min(span.end, to);
    if (end > start) {
      width += measurer.widthOf(text.slice(start, end), span.font);
    }
  }
  return width;
}

function fontAt(spans: FontSpan[], offset: number): ResolvedFont | null {
  for (const span of spans) {
    if (offset >= span.start && offset < span.end) {
      return span.font;
    }
  }
  return spans.length > 0 ? spans[spans.length - 1].font : null;
}

/** The per-line layout advance under the paragraph's line-height rule:
 * `lineSpacingPt` fixes it outright (spans inherit the absolute value), a
 * unitless `lineSpacing` scales each font's own size, and `normal` uses the
 * font's own metric-derived advance. */
function advanceFor(
  font: ResolvedFont,
  geometry: LineGeometry,
  lineSpacingPt: number | undefined,
  lineSpacing: number | undefined,
): number {
  if (lineSpacingPt !== undefined) {
    return lineSpacingPt * PX_PER_PT;
  }
  if (lineSpacing !== undefined) {
    return font.sizePx * lineSpacing;
  }
  return geometry.advancePx;
}

/** Predicts a paragraph's rendered metrics (line boxes, height, margins)
 * and line-start offsets at a section content width, without touching the
 * DOM. Returns null when the measurer cannot supply geometry — the caller
 * keeps its estimate and waits for a DOM measurement instead. */
export function measureParagraphText(
  paragraph: Paragraph,
  contentWidthPx: number,
  base: BaseFont,
  marksOn: boolean,
  measurer: TextMeasurer,
  tabConfig?: TabLayoutConfig,
): MeasuredParagraph | null {
  const formatting = paragraph.formatting;
  const baseFont: ResolvedFont = { family: base.family, sizePx: base.sizePx, bold: false, italic: false };
  const strutOrNull = measurer.normalGeometryOf(baseFont);
  if (!strutOrNull) {
    return null;
  }
  const strut: LineGeometry = strutOrNull;

  // Horizontal geometry, mirroring paragraphStyle: logical start/end
  // indents (a list level's indent when no direct one is set), first-line
  // indent on line 0 only. The hanging list marker paints outside the
  // content box (negative margin) and is not subtracted.
  const indentStart =
    paragraph.list && formatting?.indentStartCm === undefined
      ? listIndentCm(paragraph.list) * PX_PER_CM
      : (formatting?.indentStartCm ?? 0) * PX_PER_CM;
  const indentEnd = (formatting?.indentEndCm ?? 0) * PX_PER_CM;
  const firstIndent = (formatting?.indentFirstLineCm ?? 0) * PX_PER_CM;
  const availWidth = Math.max(1, contentWidthPx - indentStart - indentEnd);
  const firstWidth = Math.max(1, availWidth - firstIndent);

  const lineSpacingPt = formatting?.lineSpacingPt;
  const lineSpacing = formatting?.lineSpacing;

  const text = paragraph.runs.map((run) => run.text).join("");
  const spans: FontSpan[] = [];
  let offset = 0;
  for (const run of paragraph.runs) {
    if (run.text.length === 0) {
      continue;
    }
    const font = resolveRunFont(run.formatting, base);
    const geometry = measurer.normalGeometryOf(font);
    if (!geometry) {
      return null;
    }
    spans.push({
      font,
      start: offset,
      end: offset + run.text.length,
      advancePx: advanceFor(font, geometry, lineSpacingPt, lineSpacing),
      inkHeightPx: geometry.inkHeightPx,
    });
    offset += run.text.length;
  }

  // Per-line vertical rollup: every line contains the paragraph's strut,
  // plus whatever run fonts place glyphs on it. Span geometry is resolved
  // above, so this is pure arithmetic on the wrap's hot path.
  const strutAdvance = advanceFor(baseFont, strut, lineSpacingPt, lineSpacing);
  let lineAdvance = strutAdvance;
  let lineInk = strut.inkHeightPx;
  const advances: number[] = [];
  const inks: number[] = [];
  const offsets: number[] = [0];

  function resetLineRollup(): void {
    lineAdvance = strutAdvance;
    lineInk = strut.inkHeightPx;
  }

  function registerFonts(from: number, to: number): void {
    for (const span of spans) {
      if (span.end <= from || span.start >= to) {
        continue;
      }
      if (span.advancePx > lineAdvance) {
        lineAdvance = span.advancePx;
      }
      if (span.inkHeightPx > lineInk) {
        lineInk = span.inkHeightPx;
      }
    }
  }

  function closeLine(nextStart: number | null): void {
    advances.push(lineAdvance);
    inks.push(lineInk);
    if (nextStart !== null) {
      offsets.push(nextStart);
      resetLineRollup();
    }
  }

  const lineWidth = (): number => (advances.length === 0 ? firstWidth : availWidth);

  // Resolved tab advances (SPEC-054): when the caller supplies the paragraph's
  // tab context, each U+0009 advances to its stop with the right alignment;
  // otherwise tabs fall back to the browser's default 8-space grid below.
  const tabAdvances =
    tabConfig && text.includes("\t")
      ? layoutParagraphTabs(
          text,
          stopsToPx(tabConfig.stops),
          tabConfig.defaultTabStopCm * PX_PER_CM,
          tabConfig.decimalSeparator,
          (from, to) => measureRange(text, spans, from, to, measurer),
        ).advances
      : null;

  // Fast path: most paragraphs of a bulk document fit on one line, and one
  // whole-text measurement decides that without the per-token walk (the
  // whole string is also the closest match to how the layout engine shapes
  // an unbroken line). Tabs and newlines need the full walk.
  const length = text.length;
  const pilcrowWidth = marksOn ? measurer.widthOf("¶", baseFont) : 0;
  if (length > 0 && !text.includes("\n") && !text.includes("\t")) {
    const total = measureRange(text, spans, 0, length, measurer) + pilcrowWidth;
    if (total <= firstWidth + FIT_EPS) {
      registerFonts(0, length);
      closeLine(null);
      return assembleParagraph(contentWidthPx, formatting, advances, inks, offsets);
    }
  }

  let x = 0;
  // The offset a break would start the next line at — after the whitespace
  // run preceding the current word (trailing whitespace hangs past the edge
  // in pre-wrap and never wraps by itself).
  let breakAt = -1;
  let i = 0;
  while (i < length) {
    const ch = text[i];
    if (ch === "\n") {
      // Forced break: pre-wrap renders the newline as a hard line end.
      registerFonts(i, i + 1);
      closeLine(i + 1);
      x = 0;
      breakAt = -1;
      i += 1;
      continue;
    }
    if (ch === "\t") {
      const resolved = tabAdvances?.get(i);
      if (resolved) {
        x += resolved.widthPx;
      } else {
        const font = fontAt(spans, i) ?? baseFont;
        const stop = TAB_SIZE_SPACES * measurer.widthOf(" ", font);
        x = stop > 0 ? (Math.floor(x / stop + FIT_EPS) + 1) * stop : x;
      }
      registerFonts(i, i + 1);
      breakAt = i + 1;
      i += 1;
      continue;
    }
    if (ch === " ") {
      let j = i + 1;
      while (j < length && text[j] === " ") {
        j += 1;
      }
      x += measureRange(text, spans, i, j, measurer);
      registerFonts(i, j);
      breakAt = j;
      i = j;
      continue;
    }
    let j = i + 1;
    while (j < length && text[j] !== " " && text[j] !== "\t" && text[j] !== "\n") {
      j += 1;
    }
    const wordWidth = measureRange(text, spans, i, j, measurer);
    if (x + wordWidth > lineWidth() + FIT_EPS && breakAt > offsets[offsets.length - 1]) {
      closeLine(breakAt);
      // The dropped trailing whitespace hangs on the previous line; the new
      // line starts at the word (breakAt normally equals i).
      x = breakAt < i ? measureRange(text, spans, breakAt, i, measurer) : 0;
      breakAt = -1;
    }
    x += wordWidth;
    registerFonts(i, j);
    i = j;
  }
  if (marksOn) {
    // The formatting-marks pilcrow (`::after` on the paragraph, so the base
    // font) adds real inline width to the last line and can wrap to its own
    // line when a trailing space leaves a break opportunity.
    if (x + pilcrowWidth > lineWidth() + FIT_EPS && breakAt === length && length > 0) {
      closeLine(length);
      x = 0;
    }
    x += pilcrowWidth;
  }
  closeLine(null);
  return assembleParagraph(contentWidthPx, formatting, advances, inks, offsets);
}

function assembleParagraph(
  contentWidthPx: number,
  formatting: Paragraph["formatting"],
  advances: number[],
  inks: number[],
  offsets: number[],
): MeasuredParagraph {
  const lines: LineBox[] = [];
  let top = 0;
  let heightPx = 0;
  for (let line = 0; line < advances.length; line += 1) {
    // Half-leading: the ink band centres inside the layout advance, which
    // is where DOM line rects sit relative to the block top.
    lines.push({ top: top + (advances[line] - inks[line]) / 2, height: inks[line] });
    top += advances[line];
    heightPx += advances[line];
  }
  return {
    metrics: {
      contentWidthPx,
      heightPx,
      marginTopPx: (formatting?.spaceBeforePt ?? 0) * PX_PER_PT,
      marginBottomPx: (formatting?.spaceAfterPt ?? 0) * PX_PER_PT,
      lines,
    },
    lineOffsets: offsets,
  };
}

/** Quotes font family names for the canvas `font` shorthand: an unquoted
 * name containing spaces or other non-identifier characters would make the
 * whole declaration invalid (silently keeping the previous font). Generic
 * keywords and already-quoted names pass through. */
function cssFamilies(family: string): string {
  return family
    .split(",")
    .map((part) => {
      const name = part.trim();
      if (name.length === 0 || name.startsWith('"') || name.startsWith("'")) {
        return name;
      }
      return /^[a-zA-Z][a-zA-Z0-9-]*$/.test(name) ? name : `"${name.replace(/"/g, '\\"')}"`;
    })
    .filter((name) => name.length > 0)
    .join(", ");
}

function cssFontOf(font: ResolvedFont): string {
  const style = font.italic ? "italic " : "";
  const weight = font.bold ? "bold " : "";
  return `${style}${weight}${font.sizePx}px ${cssFamilies(font.family)}`;
}

export interface CanvasTextMeasurer extends TextMeasurer {
  /** Drops cached font geometry — call when loaded web fonts change what a
   * family resolves to (the same moment block metrics are invalidated). */
  reset(): void;
}

/** The production measurer: canvas measureText for widths (the canvas
 * shares the layout engine's text shaper) plus a hidden two-line DOM probe
 * per font for `line-height: normal` geometry, mounted under `probeHost`
 * so it inherits the app's font context. Returns null where the
 * environment cannot measure text (jsdom). */
export function createCanvasTextMeasurer(
  probeHost: () => HTMLElement | null,
): CanvasTextMeasurer | null {
  if (typeof document === "undefined") {
    return null;
  }
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return null;
  }
  ctx.font = "16px sans-serif";
  if (!(ctx.measureText("M").width > 0)) {
    return null;
  }
  let activeFont = "";
  const geometry = new Map<string, LineGeometry | null>();
  // The canvas font shorthand per resolved-font OBJECT: widthOf runs once
  // per token on the wrap's hot path, and rebuilding (and re-quoting) the
  // string each time dominated bulk measurement.
  const fontCss = new WeakMap<ResolvedFont, string>();

  function cssOf(font: ResolvedFont): string {
    let css = fontCss.get(font);
    if (css === undefined) {
      css = cssFontOf(font);
      fontCss.set(font, css);
    }
    return css;
  }

  function probeGeometry(font: ResolvedFont): LineGeometry | null {
    const host = probeHost();
    if (!host) {
      return null;
    }
    const probe = document.createElement("div");
    probe.style.position = "absolute";
    probe.style.left = "-9999px";
    probe.style.top = "0";
    probe.style.visibility = "hidden";
    probe.style.whiteSpace = "pre";
    probe.style.margin = "0";
    probe.style.padding = "0";
    probe.style.border = "0";
    probe.style.lineHeight = "normal";
    probe.style.fontFamily = font.family;
    probe.style.fontSize = `${font.sizePx}px`;
    probe.style.fontWeight = font.bold ? "bold" : "normal";
    probe.style.fontStyle = font.italic ? "italic" : "normal";
    probe.setAttribute("aria-hidden", "true");
    probe.textContent = "Mg\nMg";
    host.appendChild(probe);
    const height = probe.getBoundingClientRect().height;
    let inkHeightPx = 0;
    const range = document.createRange();
    range.selectNodeContents(probe);
    try {
      for (const rect of Array.from(range.getClientRects())) {
        if (rect.height > 0) {
          inkHeightPx = rect.height;
          break;
        }
      }
    } catch {
      inkHeightPx = 0;
    }
    probe.remove();
    if (!(height > 0)) {
      return null;
    }
    const advancePx = height / 2;
    return { advancePx, inkHeightPx: inkHeightPx > 0 ? inkHeightPx : advancePx };
  }

  return {
    widthOf(textValue: string, font: ResolvedFont): number {
      const css = cssOf(font);
      if (css !== activeFont) {
        ctx!.font = css;
        activeFont = css;
      }
      return ctx!.measureText(textValue).width;
    },
    normalGeometryOf(font: ResolvedFont): LineGeometry | null {
      const key = cssOf(font);
      if (!geometry.has(key)) {
        geometry.set(key, probeGeometry(font));
      }
      return geometry.get(key) ?? null;
    },
    reset(): void {
      geometry.clear();
    },
  };
}
