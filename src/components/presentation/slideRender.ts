// Pure render helpers for the slide components (SPEC-079). The `.vue` scripts
// here follow this repo's ESLint-espree convention — runtime `defineProps`,
// no TS type syntax (generics, annotations, casts, type imports) — so every
// typed transformation (prop coercion, per-frame style computation, frame
// narrowing) lives in this `.ts` module and the templates call into it. Fully
// framework-agnostic: plain data in, inline-style records / booleans out.
import { cmToPx } from "../../lib/core/units";
import { mediaDataUri, type MediaPart } from "../../lib/core/media";
import { computeOwnerGrid, type BorderSide } from "../../lib/core/tables";
import {
  DEFAULT_TABLE_CELL_PADDING_CM,
  DEFAULT_TABLE_ROW_HEIGHT_CM,
  type Slide,
  type SlideFrame,
  type SlideParagraph,
  type SlideRun,
  type SlideTableRow,
  type TableFrame,
} from "../../lib/presentation/model";
import { slideOutline } from "../../lib/presentation/slideModel";
import { fieldKindToken, visibleFields } from "../../lib/presentation/headerFooter";
import type {
  PlaceholderRole,
  PresentationModel,
} from "../../lib/presentation/model";
import { DEFAULT_THEME, resolveColor, resolveFont, type DocumentTheme } from "../../lib/core/theme";
import { sanitizeHref } from "../../lib/core/urls";
import { effectiveHangingCm, effectiveIndentCm, paragraphMarkers } from "../../lib/presentation/slideLists";
import { effectiveRunLink, isLinkableKind, linkOf, type SlideLinkTarget } from "../../lib/presentation/slideLinks";
import { storedScale, type TextBodyFrame } from "../../lib/presentation/textBody";
import type { ShapeVisual } from "../../lib/core/shapes";
import { shapeDefinition } from "../../lib/core/shapeCatalogue";
import type { ChartVisual } from "../../lib/core/charts";

/** pt → px at 96/72, the ratio the text editor uses. */
const PT_TO_PX = 96 / 72;
const DEFAULT_FONT_SIZE_PT = 18;
/** CSS's single line pitch (line-height 1.2), the SPEC-079 renderer's own. */
const SINGLE_LINE_HEIGHT = 1.2;

/** A centimetre length as a px string (the model's own cm → CSS px). */
export function cmPx(cm: number): string {
  return `${cmToPx(cm)}px`;
}

/** The default field look (SPEC-093 `DEFAULT_FIELD_BAND`). */
const DEFAULT_FIELD_SIZE_PT = 12;
const DEFAULT_FIELD_COLOR = "#595959";
const JUSTIFY_FOR_ALIGN: Record<"left" | "center" | "right", string> = {
  left: "flex-start",
  center: "center",
  right: "flex-end",
};

/** One deck field positioned on the stage (SPEC-093): its DOM token, resolved
 * text and inline style. */
export interface StageField {
  kindToken: string;
  text: string;
  style: Record<string, string>;
}

/** The visible deck fields a slide renders at `instant` (SPEC-093 AC-27/28):
 * every surface (stage, filmstrip, play view, build preview) resolves through
 * `fieldText`, so a suppressed title slide or an off field yields no element. */
export function stageFields(model: PresentationModel, slideIndex: number, instant: Date): StageField[] {
  return visibleFields(model, slideIndex, instant)
    .filter(({ text }) => text !== "")
    .map(({ field, text }) => ({
      kindToken: fieldKindToken(field.kind),
      text,
      style: {
        left: cmPx(field.xCm),
        top: cmPx(field.yCm),
        width: cmPx(field.widthCm),
        height: cmPx(field.heightCm),
        "justify-content": JUSTIFY_FOR_ALIGN[field.align ?? "left"],
        "font-size": `${(field.sizePt ?? DEFAULT_FIELD_SIZE_PT) * PT_TO_PX}px`,
        color: field.color ?? DEFAULT_FIELD_COLOR,
        ...(field.fontFamily ? { "font-family": field.fontFamily } : {}),
      },
    }));
}

/** Coerces a `.vue` template's runtime pre-resolved field list to `StageField[]`. */
export function asStageFields(value: unknown): StageField[] {
  return (value ?? []) as StageField[];
}

/** Coerces a runtime array-of-field-lists (one per slide) for the filmstrip. */
export function asStageFieldsBySlide(value: unknown): StageField[][] {
  return (value ?? []) as StageField[][];
}

// --- prop coercion (the `as*` pattern, like documentModel's asShape) --------
export function asSlide(value: unknown): Slide {
  return value as Slide;
}
export function asSlideFrame(value: unknown): SlideFrame {
  return value as SlideFrame;
}
export function asSlideFrames(value: unknown): SlideFrame[] {
  return (value ?? []) as SlideFrame[];
}
export function asStrings(value: unknown): string[] {
  return (value ?? []) as string[];
}
export function asMediaParts(value: unknown): MediaPart[] {
  return (value ?? []) as MediaPart[];
}
export function asSlides(value: unknown): Slide[] {
  return (value ?? []) as Slide[];
}
export function asResolver(value: unknown): (mediaId: string) => string {
  return value as (mediaId: string) => string;
}
export function asNumber(value: unknown): number {
  return Number(value) || 0;
}
/** Coerces a template's runtime theme prop; absent = Office (SPEC-094 AC-7). */
export function asTheme(value: unknown): DocumentTheme {
  return (value ?? DEFAULT_THEME) as DocumentTheme;
}
export function asString(value: unknown): string {
  return value === undefined || value === null ? "" : String(value);
}

// --- media resolution -------------------------------------------------------
/** A resolver from `mediaId` to a `data:` URI over a media-part list. */
export function mediaResolver(media: MediaPart[]): (mediaId: string) => string {
  const map = new Map(media.map((part) => [part.mediaId, part]));
  return (mediaId) => mediaDataUri(map.get(mediaId));
}

// --- frame narrowing (templates cannot narrow a discriminated union) --------
export function frameKind(frame: SlideFrame): SlideFrame["kind"] {
  return frame.kind;
}
export function isForeignVisible(frame: SlideFrame): boolean {
  return frame.kind === "foreign" && frame.noBounds !== true;
}
export function imageSrc(frame: SlideFrame, resolveMedia: (mediaId: string) => string): string {
  return frame.kind === "image" ? resolveMedia(frame.mediaId) : "";
}
export function foreignLabel(frame: SlideFrame): string {
  return frame.kind === "foreign" ? frame.label : "";
}
export function groupChildren(frame: SlideFrame): SlideFrame[] {
  return frame.kind === "group" ? frame.children : [];
}

/** The names a placeholder is announced under, by role. Every placeholder used
 * to be "Placeholder", so a slide's title and its body were indistinguishable
 * to anything reading the accessible name — two controls, one name, while the
 * screen plainly said "Click to add title" and "Click to add text". */
const PLACEHOLDER_NAMES: Record<PlaceholderRole, string> = {
  title: "Title placeholder",
  subtitle: "Subtitle placeholder",
  body: "Body placeholder",
  other: "Placeholder",
};

/** The accessible name a top-level frame is exposed under (AC-52): "Image",
 * "Text box", the catalogue shape name, "Group", or a foreign label. */
export function frameAccessibleName(frame: SlideFrame): string {
  switch (frame.kind) {
    case "image":
      return "Image";
    case "group":
      return "Group";
    case "foreign":
      return frame.label || "Object";
    case "shape":
      return shapeDefinition(frame.geometry)?.label ?? "Shape";
    case "table":
      return "Table";
    case "chart":
      return "Chart";
    case "media":
      return frame.media === "audio" ? "Audio" : "Video";
    default:
      return frame.role ? PLACEHOLDER_NAMES[frame.role] : "Text box";
  }
}

/** A frame's ARIA role: `group` for shapes/images/groups/foreign boxes; text
 * frames keep SPEC-081's text semantics (no group role, AC-51). */
export function frameRole(frame: SlideFrame): string | undefined {
  return frame.kind === "text" ? undefined : "group";
}

// --- positioning & fills ----------------------------------------------------
export function framePositionStyle(frame: SlideFrame): Record<string, string> {
  const style: Record<string, string> = {
    position: "absolute",
    left: cmPx(frame.xCm),
    top: cmPx(frame.yCm),
    width: cmPx(frame.widthCm),
    height: cmPx(frame.heightCm),
  };
  if (frame.rotationDeg) {
    style.transform = `rotate(${frame.rotationDeg}deg)`;
    style.transformOrigin = "center";
  }
  return style;
}

/** The origin shim inside a group's box: children carry absolute SLIDE-space
 * coordinates, so the shim shifts the positioning context back to the stage's
 * origin. The group box itself then occupies its real geometry, which makes a
 * `rotate()` transform on it spin about the group's own centre (SPEC-096
 * AC-45) instead of whatever corner a zero-sized wrapper would pin. */
export function groupOriginStyle(frame: SlideFrame): Record<string, string> {
  return { position: "absolute", left: cmPx(-frame.xCm), top: cmPx(-frame.yCm) };
}

/** A text frame's own background colour, or "transparent". Scheme references
 * resolve through `resolveColor` (SPEC-094 AC-13). */
export function textFrameFill(frame: SlideFrame, theme: DocumentTheme = DEFAULT_THEME): string {
  if (frame.kind === "text" && frame.fill && frame.fill.type === "solid") {
    return resolveColor(frame.fill.color, theme);
  }
  return "transparent";
}

/** The slide surface background — always light (white unless the deck says
 * otherwise). */
export function surfaceBackground(slide: Slide, theme: DocumentTheme = DEFAULT_THEME): string {
  const bg = slide.background;
  return bg && bg.type === "solid" ? resolveColor(bg.color, theme) : "#ffffff";
}

// --- text ------------------------------------------------------------------
/** One paragraph annotated with its derived list marker (SPEC-097: numbering
 * is DERIVED by adjacency in lib/presentation/slideLists.ts, never stored).
 * `marker` is the glyph/ordinal text ("" for a non-list paragraph or a
 * picture bullet); `markerImageId` names a picture bullet's media part. The
 * `{ para, marker }` shape is SPEC-079's, kept so every consumer (the stage,
 * the library's SlideMiniature) reads one list. */
export interface MarkedParagraph {
  para: SlideParagraph;
  marker: string;
  markerImageId?: string;
}

/** A paragraph list (a text body or a table cell) annotated with its markers,
 * through `paragraphMarkers` — the one derivation every surface shares (the
 * stage, filmstrip thumbnails, the play view and the build preview render
 * through SlideFrameView; SPEC-088's export calls `paragraphMarkers` itself). */
export function markParagraphList(paragraphs: SlideParagraph[]): MarkedParagraph[] {
  const markers = paragraphMarkers(paragraphs);
  return paragraphs.map((para, index) => {
    const marker = markers[index] ?? { text: "" };
    return marker.imageId ? { para, marker: "", markerImageId: marker.imageId } : { para, marker: marker.text };
  });
}

export function markedParagraphs(frame: SlideFrame): MarkedParagraph[] {
  const paragraphs = frame.kind === "text" ? frame.paragraphs : frame.kind === "shape" ? frame.paragraphs ?? [] : [];
  return markParagraphList(paragraphs);
}

/** Whether a text frame renders its empty-placeholder prompt (SPEC-081): a
 * placeholder (has a role) whose paragraphs are all empty. */
export function isEmptyTextPlaceholder(frame: SlideFrame): boolean {
  if (frame.kind !== "text" || frame.role === undefined) {
    return false;
  }
  return frame.paragraphs.every((para) => para.runs.every((run) => run.text.length === 0));
}

/** The prompt an empty placeholder shows outside an editing session — never
 * enters the model, the files, or the clipboard (AC-15/16). */
export function placeholderPrompt(frame: SlideFrame): string {
  if (frame.kind !== "text") {
    return "";
  }
  return frame.role === "title" ? "Click to add title" : "Click to add text";
}

/** A pt length as a px string that scales with the text body's autofit font
 * scale (`--fit-scale`, set on the `.text-body` by slideTextFit.ts#bodyStyle;
 * 1 anywhere else, e.g. a table cell). */
function scaledPtPx(pt: number): string {
  return `calc(${round3(pt * PT_TO_PX)}px * var(--fit-scale, 1))`;
}

const round3 = (value: number): number => Math.round(value * 1000) / 1000;

/** A paragraph's line pitch as a CSS `line-height` (SPEC-097 AC-6..8), scaled
 * by the body's `--line-scale`:
 * - a multiple `m` (absent = 1) → unitless `1.2 × m`, so every run's line
 *   grows with its own size, exactly SPEC-079's `line-height: 1.2` at 1×;
 * - `lineSpacingPt` with rule "exact" → that pitch in px (scaled with the
 *   font too, so a shrink step actually shrinks it);
 * - `lineSpacingPt` "at least" (rule absent) → the pitch in px on the
 *   PARAGRAPH only, which CSS treats as the line box's strut — and
 *   `runLineHeight` gives the runs their own natural 1.2 pitch, so a line is
 *   at least the pitch and grows for bigger type. That is the CSS spelling
 *   of "at least" (CSS has no `line-height: max(...)` per line). */
export function paragraphLineHeight(para: SlideParagraph): string {
  if (para.lineSpacingPt !== undefined) {
    return `calc(${round3(para.lineSpacingPt * PT_TO_PX)}px * var(--fit-scale, 1) * var(--line-scale, 1))`;
  }
  const multiple = para.lineSpacing ?? 1;
  return `calc(${round3(SINGLE_LINE_HEIGHT * multiple)} * var(--line-scale, 1))`;
}

/** The run-level line height of an "at least" paragraph (see
 * `paragraphLineHeight`); undefined when the runs simply inherit. */
export function runLineHeight(para: SlideParagraph): string | undefined {
  if (para.lineSpacingPt !== undefined && para.lineSpacingRule !== "exact") {
    return `calc(${SINGLE_LINE_HEIGHT} * var(--line-scale, 1))`;
  }
  return undefined;
}

/** A paragraph's block style (SPEC-079, widened by SPEC-097):
 * - indent from the model (AC-32/33): the text edge sits `indentCm` in
 *   (`margin-left`) and the first line starts `hangingCm` earlier
 *   (`text-indent: -hanging`), where the marker box (`markerStyle`) hangs —
 *   both through slideLists' `effectiveIndentCm`/`effectiveHangingCm`, so a
 *   list item's default is `(level + 1) × 1.27cm` (the writers' value; the old
 *   private 0.95cm/level is gone — a deliberate correction);
 * - line pitch (`paragraphLineHeight`) and space before/after in pt → px,
 *   scaled with the body's fit scale.
 * `defaultAlign` is the frame's own: "left" for a text frame/cell, "center"
 * for shape text (SPEC-079's centred shape text, now only when the paragraph
 * states no alignment of its own). */
export function paragraphStyle(para: SlideParagraph, defaultAlign: NonNullable<SlideParagraph["align"]> = "left"): Record<string, string> {
  const style: Record<string, string> = {
    textAlign: para.align ?? defaultAlign,
    marginLeft: cmPx(effectiveIndentCm(para)),
    lineHeight: paragraphLineHeight(para),
  };
  const hanging = effectiveHangingCm(para);
  if (hanging !== 0) style.textIndent = cmPx(-hanging);
  if (para.spaceBeforePt) style.marginTop = scaledPtPx(para.spaceBeforePt);
  if (para.spaceAfterPt) style.marginBottom = scaledPtPx(para.spaceAfterPt);
  return style;
}

/** The marker box of a list paragraph: an inline-block `hangingCm` wide (at
 * least — a wider ordinal pushes its text on rather than overprinting it),
 * sized and coloured like the paragraph's first run (PowerPoint's 100%
 * `buSzPct`/follow-text colour; bullet colour/size are out of scope). */
export function markerStyle(entry: MarkedParagraph, theme: DocumentTheme = DEFAULT_THEME): Record<string, string> {
  const first = entry.para.runs[0];
  const style: Record<string, string> = {
    display: "inline-block",
    minWidth: cmPx(effectiveHangingCm(entry.para)),
    textIndent: "0",
    fontSize: scaledPtPx(first?.sizePt ?? DEFAULT_FONT_SIZE_PT),
    color: first?.color ? resolveColor(first.color, theme) : "#000000",
  };
  const lineHeight = runLineHeight(entry.para);
  if (lineHeight) style.lineHeight = lineHeight;
  return style;
}

/** Whether a marked paragraph shows a marker box at all. */
export function hasMarker(entry: MarkedParagraph): boolean {
  return entry.marker !== "" || entry.markerImageId !== undefined;
}

/** A picture bullet's image source ("" when the paragraph has none or its
 * media does not resolve — the marker then falls back to "•"). */
export function markerImageSrc(entry: MarkedParagraph, resolveMedia: (mediaId: string) => string): string {
  return entry.markerImageId ? resolveMedia(entry.markerImageId) : "";
}

/** The text a marker box shows when it has no (resolvable) picture. */
export function markerFallbackText(entry: MarkedParagraph): string {
  return entry.marker || "•";
}

export function runStyle(run: SlideRun, theme: DocumentTheme = DEFAULT_THEME): Record<string, string> {
  const style: Record<string, string> = {};
  if (run.bold) style.fontWeight = "700";
  if (run.italic) style.fontStyle = "italic";
  if (run.underline) style.textDecoration = "underline";
  if (run.fontFamily) style.fontFamily = `'${resolveFont(run.fontFamily, theme)}', sans-serif`;
  // Through the body's `--fit-scale` (SPEC-097 autofit): one rendered body
  // re-scales by changing a custom property, without re-rendering its runs.
  style.fontSize = scaledPtPx(run.sizePt ?? DEFAULT_FONT_SIZE_PT);
  style.color = run.color ? resolveColor(run.color, theme) : "#000000";
  if (run.letterSpacingPt) style.letterSpacing = scaledPtPx(run.letterSpacingPt);
  return style;
}

// --- links (SPEC-097 AC-47..50) ----------------------------------------------
/** The test id (and `closest()` hook) of every rendered link element — a
 * linked run's `<a>` and a linked frame's overlay. */
export const SLIDE_LINK_TESTID = "slide-link";
/** The bubbling event the editor stage dispatches from a Cmd/Ctrl+clicked
 * in-deck link (`detail.slide` is 1-based); the editor's link composable
 * listens for it on `document` and navigates. */
export const SLIDE_LINK_FOLLOW_EVENT = "slide-link-follow";

/** A link as the renderer emits it. */
export interface LinkRender {
  /** The anchor's `href`: the external URL, or `#slide-N` for an in-deck jump. */
  href: string;
  /** The sanitised external URL ("" for an in-deck link). */
  external: string;
  /** The 1-based in-deck slide (0 for an external link). */
  slide: number;
  /** The hover tooltip (undefined = no `title`). */
  tooltip: string | undefined;
}

/** A link target as rendered, or null when there is nothing safe to link:
 * an external href is re-checked through `sanitizeHref` at render (the model
 * only ever stores clean hrefs; the renderer does not trust that). */
export function linkRender(target: SlideLinkTarget | null): LinkRender | null {
  if (!target) return null;
  const tooltip = target.tooltip ? target.tooltip : undefined;
  if (target.slide !== undefined && Number.isInteger(target.slide) && target.slide >= 1) {
    return { href: `#slide-${target.slide}`, external: "", slide: target.slide, tooltip };
  }
  const href = target.href !== undefined ? sanitizeHref(target.href) : null;
  return href ? { href, external: href, slide: 0, tooltip } : null;
}

/** One run as rendered: text, inline style, and its link (own, else its
 * frame's — `effectiveRunLink`). A linked run is underlined in its OWN colour
 * (theme `hlink` belongs to SPEC-094). */
export interface RunRender {
  text: string;
  style: Record<string, string>;
  link: LinkRender | null;
}

export function runRenders(para: SlideParagraph, frame: SlideFrame | null, theme: DocumentTheme = DEFAULT_THEME): RunRender[] {
  const lineHeight = runLineHeight(para);
  return para.runs.map((run) => {
    const style = runStyle(run, theme);
    if (lineHeight) style.lineHeight = lineHeight;
    const link = linkRender(effectiveRunLink(run, frame ?? undefined));
    if (link) style.textDecoration = "underline";
    return { text: run.text, style, link };
  });
}

/** A whole-frame link (text/image/shape frames only), as rendered. */
export function frameLinkRender(frame: SlideFrame): LinkRender | null {
  return isLinkableKind(frame.kind) ? linkRender(linkOf(frame)) : null;
}

/** The rendered link element an event landed in (`closest()` over the
 * `slide-link` test id), or null. */
export function slideLinkElement(target: EventTarget | null): HTMLElement | null {
  if (!(target instanceof Element)) return null;
  const el = target.closest(`[data-testid="${SLIDE_LINK_TESTID}"]`);
  return el instanceof HTMLElement ? el : null;
}

/** The target a rendered link element carries (its `data-link-slide` /
 * `data-link-href`), re-sanitised; null when it carries nothing usable. */
export function slideLinkTargetOf(el: Element): { slide?: number; href?: string } | null {
  const slide = Number(el.getAttribute("data-link-slide"));
  if (Number.isInteger(slide) && slide >= 1) return { slide };
  const raw = el.getAttribute("data-link-href");
  const href = raw ? sanitizeHref(raw) : null;
  return href ? { href } : null;
}

/** Opens an external link the way both playback surfaces do: a new tab with
 * no opener and no referrer. */
export function openExternalLink(href: string): void {
  window.open(href, "_blank", "noopener,noreferrer");
}

/** The editor stage's Cmd/Ctrl+click on a link element: an external link
 * opens in a new tab; an in-deck link dispatches the bubbling
 * `slide-link-follow` event (detail `{ slide }`, 1-based) from the element.
 * Returns whether anything was followed. */
export function followEditorLink(el: Element): boolean {
  const target = slideLinkTargetOf(el);
  if (!target) return false;
  if (target.slide !== undefined) {
    el.dispatchEvent(new CustomEvent(SLIDE_LINK_FOLLOW_EVENT, { bubbles: true, detail: { slide: target.slide } }));
  } else if (target.href) {
    openExternalLink(target.href);
  }
  return true;
}

// --- text body (SPEC-097 AC-11..15, 37..40) ----------------------------------
/** A frame narrowed to a text-body frame (text or shape), or null. */
export function textBodyFrameOf(frame: SlideFrame): TextBodyFrame | null {
  return frame.kind === "text" || frame.kind === "shape" ? frame : null;
}

/** Whether a text-body frame's text does not wrap (`body.wrap: false`). */
export function bodyNoWrap(frame: SlideFrame): boolean {
  const body = textBodyFrameOf(frame)?.body;
  return body?.wrap === false;
}

/** The scale a text body renders at before the stage measures it: the
 * frame's STORED shrink scale (`textBody.storedScale`), 100/100 otherwise. */
export function initialFitScale(frame: SlideFrame): { fontScalePct: number; lineScalePct: number } {
  const body = textBodyFrameOf(frame);
  return body ? storedScale(body) : { fontScalePct: 100, lineScalePct: 100 };
}

/** The "nothing measured yet" scale a renderer's local ref starts from
 * (typed here — the `.vue` scripts carry no type syntax). */
export function noMeasuredScale(): { fontScalePct: number; lineScalePct: number } | null {
  return null;
}

/** Whether two fit scales differ enough to re-render (measurement noise
 * below a hundredth of a percent is "the same"). */
export function fitScaleDiffers(
  a: { fontScalePct: number; lineScalePct: number },
  b: { fontScalePct: number; lineScalePct: number },
): boolean {
  return Math.abs(a.fontScalePct - b.fontScalePct) > 0.01 || Math.abs(a.lineScalePct - b.lineScalePct) > 0.01;
}

/** The default paragraph alignment of a frame's text: shape text centres
 * (SPEC-079's renderer assumption) unless a paragraph says otherwise. */
export function defaultParagraphAlign(frame: SlideFrame): "left" | "center" {
  return frame.kind === "shape" ? "center" : "left";
}

// --- shapes -----------------------------------------------------------------
/** A shape frame as a ShapeVisual for `paintShapeIntoCanvas`, or null for a
 * non-shape. Rotation is left to the frame's CSS transform (AC-36), so the
 * canvas paints unrotated. */
export function shapeVisualOf(frame: SlideFrame, theme: DocumentTheme = DEFAULT_THEME): ShapeVisual | null {
  if (frame.kind !== "shape") {
    return null;
  }
  const visual: ShapeVisual = { geometry: frame.geometry };
  if (frame.fill && frame.fill.type === "solid") {
    visual.fill = { kind: "solid", color: resolveColor(frame.fill.color, theme) };
  }
  if (frame.stroke) {
    visual.outline = { color: resolveColor(frame.stroke.color, theme), widthPt: frame.stroke.widthPt, dash: "solid" };
  }
  if (frame.adjustments && frame.adjustments.length > 0) {
    visual.adjustments = frame.adjustments;
  }
  return visual;
}

// --- charts (SPEC-092) ------------------------------------------------------
/** A chart frame as a `ChartVisual` for `paintChart`, or null for a non-chart.
 * The frame IS a `ChartVisual` (it extends it), so this narrows and hands the
 * painter the chart-type/categories/series/options (and `unsupportedChartXml`,
 * which makes `paintChart` draw its dashed "Chart not supported" placeholder).
 * Rotation is left to the frame's CSS transform, so the canvas paints
 * unrotated — the same rule `shapeVisualOf` follows. */
export function chartVisualOf(frame: SlideFrame): ChartVisual | null {
  if (frame.kind !== "chart") {
    return null;
  }
  const visual: ChartVisual = {
    chartType: frame.chartType,
    categories: frame.categories,
    series: frame.series,
    options: frame.options,
  };
  if (frame.unsupportedChartXml !== undefined) {
    visual.unsupportedChartXml = frame.unsupportedChartXml;
  }
  return visual;
}

// --- tables (SPEC-091) ------------------------------------------------------
// The effective look resolves per "Table styles" (AC-4b): an explicit per-cell
// value wins; else the header/banding look driven by the frame's firstRow/
// bandRow flags; else the neutral built-in fallback (a thin grid). Deck table
// styles (styleId) are not themselves parsed — they resolve to the fallback.

const NEUTRAL_BORDER_COLOR = "#cfcfcf";
const HEADER_FILL = "#d9d9d9";
const BAND_FILL = "#f3f3f3";

/** A frame narrowed to a table, or null. */
export function tableOf(frame: SlideFrame): TableFrame | null {
  return frame.kind === "table" ? frame : null;
}

/** The `<colgroup>` widths (px) for a table frame. */
export function tableColumnWidths(frame: SlideFrame): string[] {
  return frame.kind === "table" ? frame.columnWidthsCm.map((cm) => cmPx(cm)) : [];
}

/** One internal grid boundary: the index of the row/column left/above it and
 * its position as a percentage of the frame's width/height, so the drag-resize
 * handles (AC-8a/b) can sit on the grid lines regardless of stage scale. */
export interface TableBoundary {
  index: number;
  offsetPct: number;
}

function boundariesFrom(sizesCm: number[]): TableBoundary[] {
  const total = sizesCm.reduce((sum, s) => sum + s, 0);
  if (total <= 0) {
    return [];
  }
  const out: TableBoundary[] = [];
  let acc = 0;
  for (let i = 0; i < sizesCm.length - 1; i++) {
    acc += sizesCm[i];
    out.push({ index: i, offsetPct: (acc / total) * 100 });
  }
  return out;
}

/** Per-row heights (px) for a table frame's `<tr>` elements, so a row resized
 * via AC-8b renders at its `heightCm` (a minimum — content may push it taller).
 * Without this the browser distributes the table's height equally and a single
 * row drag would grow the whole grid rather than the targeted row. */
export function tableRowHeights(frame: SlideFrame): string[] {
  const table = tableOf(frame);
  return table ? table.rows.map((row) => cmPx(row.heightCm ?? DEFAULT_TABLE_ROW_HEIGHT_CM)) : [];
}

/** The internal column boundaries a selected table exposes for drag-resize. */
export function tableColumnBoundaries(frame: SlideFrame): TableBoundary[] {
  const table = tableOf(frame);
  return table ? boundariesFrom(table.columnWidthsCm) : [];
}

/** The internal row boundaries a selected table exposes for drag-resize. */
export function tableRowBoundaries(frame: SlideFrame): TableBoundary[] {
  const table = tableOf(frame);
  return table ? boundariesFrom(table.rows.map((row) => row.heightCm ?? DEFAULT_TABLE_ROW_HEIGHT_CM)) : [];
}

function isHeaderRow(frame: TableFrame, rowIndex: number, row: SlideTableRow): boolean {
  return row.isHeader === true || (frame.firstRow === true && rowIndex === 0);
}

function bandFillFor(frame: TableFrame, rowIndex: number): string | undefined {
  if (!frame.bandRow) {
    return undefined;
  }
  const bodyStart = frame.firstRow ? 1 : 0;
  if (rowIndex < bodyStart) {
    return undefined;
  }
  return (rowIndex - bodyStart) % 2 === 1 ? BAND_FILL : undefined;
}

function effectiveCellFill(
  frame: TableFrame,
  rowIndex: number,
  row: SlideTableRow,
  colIndex: number,
  theme: DocumentTheme,
): string {
  const cell = row.cells[colIndex];
  if (cell.fill) {
    return cell.fill.type === "solid" ? resolveColor(cell.fill.color, theme) : "transparent";
  }
  if (isHeaderRow(frame, rowIndex, row)) {
    return HEADER_FILL;
  }
  return bandFillFor(frame, rowIndex) ?? "transparent";
}

function borderCss(side: BorderSide | undefined): string {
  if (!side) {
    return `1px solid ${NEUTRAL_BORDER_COLOR}`; // neutral fallback grid
  }
  if (side.style === "none") {
    return "none";
  }
  const line =
    side.style === "double" ? "double" : side.style === "dashed" ? "dashed" : side.style === "dotted" ? "dotted" : "solid";
  return `${Math.max(1, side.widthPt)}px ${line} ${side.color}`;
}

/** One rendered (owner) cell of a table: its span, computed style and marked
 * paragraphs. Covered slots emit nothing (AC-4a). */
export interface TableCellRender {
  rowIndex: number;
  colIndex: number;
  colSpan: number;
  rowSpan: number;
  style: Record<string, string>;
  paragraphs: MarkedParagraph[];
}

/** The table's rows as arrays of visible owner cells (covered slots dropped),
 * each carrying the colspan/rowspan its owner claims and its effective style. */
export function tableRenderRows(frame: SlideFrame, theme: DocumentTheme = DEFAULT_THEME): TableCellRender[][] {
  const table = tableOf(frame);
  if (!table) {
    return [];
  }
  const grid = computeOwnerGrid(table);
  return table.rows.map((row, r) => {
    const out: TableCellRender[] = [];
    for (let c = 0; c < table.columnWidthsCm.length; c++) {
      const owner = grid[r][c];
      if (owner.originRow !== r || owner.originCol !== c) {
        continue; // covered slot — its owner's colspan/rowspan covers it
      }
      const cell = row.cells[c];
      out.push({
        rowIndex: r,
        colIndex: c,
        colSpan: cell.colSpan ?? 1,
        rowSpan: cell.rowSpan ?? 1,
        style: {
          borderTop: borderCss(cell.borders?.top),
          borderBottom: borderCss(cell.borders?.bottom),
          borderLeft: borderCss(cell.borders?.left),
          borderRight: borderCss(cell.borders?.right),
          padding: cmPx(cell.paddingCm ?? DEFAULT_TABLE_CELL_PADDING_CM),
          verticalAlign: cell.verticalAlign ?? "top",
          background: effectiveCellFill(table, r, row, c, theme),
        },
        paragraphs: markParagraphList(cell.paragraphs),
      });
    }
    return out;
  });
}

// --- filmstrip --------------------------------------------------------------
export function filmstripLabel(slide: Slide, index: number, total: number): string {
  const outline = slideOutline(slide);
  const name = outline.title && outline.title.trim().length > 0 ? outline.title : "untitled";
  return `Slide ${index + 1} of ${total} — ${name}`;
}

export function filmstripThumbScale(widthCm: number, thumbWidthPx: number): number {
  return thumbWidthPx / Math.max(1, cmToPx(widthCm));
}

// --- filmstrip menu / drag DOM helpers (SPEC-080) --------------------------
// The .vue script is espree-linted (no TS casts/generics), so the typed DOM
// queries the context menu needs live here.

/** Focuses the first element matching `selector` inside `container`. */
export function focusSelector(container: Element | null | undefined, selector: string): void {
  const el = container?.querySelector(selector);
  if (el instanceof HTMLElement) {
    el.focus();
  }
}

/** The focusable menu-item buttons inside a rendered menu, in DOM order. */
export function menuButtons(container: Element | null | undefined): HTMLElement[] {
  const nodes = container ? Array.from(container.querySelectorAll('[role="menuitem"]')) : [];
  return nodes.filter((node): node is HTMLElement => node instanceof HTMLElement);
}

/** A viewport anchor point just inside an element (for a keyboard-opened
 * menu), or the origin when it can't be found. */
export function anchorPoint(container: Element | null | undefined, selector: string): { x: number; y: number } {
  const el = container?.querySelector(selector);
  if (el instanceof HTMLElement) {
    const rect = el.getBoundingClientRect();
    return { x: rect.left + 24, y: rect.bottom };
  }
  return { x: 0, y: 0 };
}
