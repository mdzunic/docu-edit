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
import type { PresentationModel } from "../../lib/presentation/model";
import { DEFAULT_THEME, resolveColor, resolveFont, type DocumentTheme } from "../../lib/core/theme";
import type { ShapeVisual } from "../../lib/core/shapes";
import { shapeDefinition } from "../../lib/core/shapeCatalogue";
import type { ChartVisual } from "../../lib/core/charts";

/** pt → px at 96/72, the ratio the text editor uses. */
const PT_TO_PX = 96 / 72;
const LEVEL_INDENT_CM = 0.95;
const DEFAULT_FONT_SIZE_PT = 18;

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
    default:
      return frame.role ? "Placeholder" : "Text box";
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
/** A text/shape frame's paragraphs annotated with their bullet marker
 * (numbered bullets restart per contiguous same-level run). */
export function markParagraphList(paragraphs: SlideParagraph[]): { para: SlideParagraph; marker: string }[] {
  const counts: Record<number, number> = {};
  let prev: SlideParagraph | undefined;
  return paragraphs.map((para) => {
    let marker = "";
    if (para.bullet === "number") {
      const level = para.level ?? 0;
      if (prev && prev.bullet === "number" && (prev.level ?? 0) === level) {
        counts[level] = (counts[level] ?? 0) + 1;
      } else {
        counts[level] = 1;
      }
      marker = `${counts[level]}.`;
    } else if (para.bullet === "disc") {
      marker = "•";
    }
    prev = para;
    return { para, marker };
  });
}

export function markedParagraphs(frame: SlideFrame): { para: SlideParagraph; marker: string }[] {
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

export function paragraphStyle(para: SlideParagraph): Record<string, string> {
  return {
    textAlign: para.align ?? "left",
    marginLeft: para.level ? cmPx(para.level * LEVEL_INDENT_CM) : "0",
  };
}

export function runStyle(run: SlideRun, theme: DocumentTheme = DEFAULT_THEME): Record<string, string> {
  const style: Record<string, string> = {};
  if (run.bold) style.fontWeight = "700";
  if (run.italic) style.fontStyle = "italic";
  if (run.underline) style.textDecoration = "underline";
  if (run.fontFamily) style.fontFamily = `'${resolveFont(run.fontFamily, theme)}', sans-serif`;
  style.fontSize = `${(run.sizePt ?? DEFAULT_FONT_SIZE_PT) * PT_TO_PX}px`;
  style.color = run.color ? resolveColor(run.color, theme) : "#000000";
  return style;
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
    visual.fill = resolveColor(frame.fill.color, theme);
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
  paragraphs: { para: SlideParagraph; marker: string }[];
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
