// Library thumbnails (SPEC-067): a lightweight rendered preview of a
// document's first page, produced from the model alone. The product's real
// page renderer (DocPage.vue + usePagination) is DOM-measurement driven and
// would have to mount the whole document off-screen — far too heavy to run
// per row of a 200-document library. Instead we extract the first page's
// leading blocks as a small list of styled lines — paragraphs as text,
// images as their resolved data: URI scaled the way the page scales them —
// and let the card render them as a scaled-down paper miniature. It is a
// genuine preview of the first page's content, cheap enough to schedule
// through the idle scheduler.
//
// This module is the pure, framework-free part: it builds and caches thumbnail
// data. The IntersectionObserver gating and idle scheduling live in the
// LibraryThumbnail component, so only visible cards ever call in here.
import type { Block, Chart, DocumentModel, Image, Shape, Table, TableCell } from "./documentModel";
import {
  isChart,
  isImage,
  isShape,
  isTable,
  sectionContentWidthCm,
  sectionPageSize,
} from "./documentModel";
import type { ShapeVisual } from "./core/shapes";
import { mediaDataUri } from "./core/media";
import { cmToPx } from "./core/units";
import { blockText } from "./textEditing";

/** An image on a thumbnail preview line, resolved for rendering. */
export interface ThumbnailImage {
  /** The media part as a `data:` URI (core/media's mediaDataUri — the same
   * resolve-only-for-rendering rule as the page canvas). */
  src: string;
  /** Display width as a fraction (0..1] of the section's content width, so
   * the miniature sizes the image exactly like the page does. */
  widthFraction: number;
  /** Display aspect ratio (widthCm / heightCm). The page stretches the
   * bitmap to its display box, so the miniature must render this ratio, not
   * the bitmap's natural one. */
  aspectRatio: number;
}

/** A table on a thumbnail preview line: its leading cells' text, laid out
 * with the table's own column proportions, so the miniature shows the same
 * grid the page does. */
export interface ThumbnailTable {
  /** Up to THUMBNAIL_TABLE_MAX_ROWS × THUMBNAIL_TABLE_MAX_COLUMNS of cell
   * text, in grid order (covered cells of a merge read as ""). */
  rows: string[][];
  /** How many leading rows are header rows (rendered emphasised). */
  headerRows: number;
  /** Column widths as fractions of the shown columns' total (sums to ~1). */
  columnFractions: number[];
  /** Table width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
}

/** A chart on a thumbnail preview line: the chart block itself (deep-cloned,
 * so the cache never aliases the live model), which the card paints small
 * through the same core/chartCanvas painter the page uses. */
export interface ThumbnailChart {
  chart: Chart;
  /** Display width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
  /** Display aspect ratio (widthCm / heightCm). */
  aspectRatio: number;
}

/** A shape or text box on a thumbnail preview line: the paintable visual
 * subset for the card's canvas (the same core/shapeCanvas painter the page
 * uses) plus the text frame's leading lines, which the page renders as DOM
 * — so a text box previews as its text, not as an empty rectangle. */
export interface ThumbnailShape {
  /** The ShapeVisual to paint — a stripped clone: `unsupportedGeometryXml`
   * is reduced to "" when present (only its presence drives the dashed
   * placeholder), so the cache never holds the preserved XML. */
  shape: ShapeVisual;
  /** The shape's box in model px — the canvas backing-store size. */
  widthPx: number;
  heightPx: number;
  /** Leading non-empty paragraphs of the shape's text frame, capped. */
  textLines: string[];
  /** The text frame's vertical alignment ("top" when there is no frame). */
  vAlign: "top" | "middle" | "bottom";
  /** Display width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
  /** Display aspect ratio (width / height, degenerate axes clamped). */
  aspectRatio: number;
}

/** One line of a thumbnail preview: a text line, an image, a table, a chart,
 * or a shape. Exactly one of the optional members is present on a non-text
 * line; `text` is then "" and `level` 0. All optional members are absent on
 * every text line, so pre-image thumbnails and their tests keep their exact
 * shape. */
export interface ThumbnailLine {
  text: string;
  /** 0 for body text, 1..6 for a heading of that level — the card scales the
   * type so a document's shape (a big title over paragraphs) reads at a
   * glance. */
  level: number;
  image?: ThumbnailImage;
  table?: ThumbnailTable;
  chart?: ThumbnailChart;
  shape?: ThumbnailShape;
}

export interface Thumbnail {
  lines: ThumbnailLine[];
  /** True when the model could not be previewed (corrupt, or nothing
   * renderable). The card then shows a format-badge placeholder, never an
   * error. */
  failed: boolean;
}

/** How many leading lines a first-page preview shows before it is visually
 * "full". More than this never fits the miniature. An image consumes several
 * line slots — its share of the page height — so a picture-heavy first page
 * fills the budget (and stops the walk, and its data-URI decodes) as fast as
 * it fills the miniature. */
export const THUMBNAIL_MAX_LINES = 16;

/** Paper aspect ratio (A4 portrait, width / height) the placeholder and the
 * rendered thumbnail both use, so nothing reflows when the real one arrives. */
export const THUMBNAIL_ASPECT_RATIO = 210 / 297;

/** How much of a table the miniature shows: enough rows and columns to read
 * its shape, never the whole grid. */
export const THUMBNAIL_TABLE_MAX_ROWS = 6;
export const THUMBNAIL_TABLE_MAX_COLUMNS = 6;

function isParagraph(block: Block): block is Extract<Block, { type: "paragraph" }> {
  return block.type === "paragraph";
}

/** Resolves an image block to its thumbnail line, or null when it cannot be
 * previewed (media part missing/empty, degenerate display size) — a bad
 * image is skipped, never fails the whole thumbnail. */
function thumbnailImage(
  model: DocumentModel,
  sectionIndex: number,
  block: Image,
): ThumbnailImage | null {
  const part = (model.media ?? []).find((media) => media.mediaId === block.mediaId);
  if (!part || part.data.length === 0 || !(block.widthCm > 0) || !(block.heightCm > 0)) {
    return null;
  }
  const src = mediaDataUri(part);
  return {
    src,
    widthFraction: drawingWidthFraction(model, sectionIndex, block),
    aspectRatio: block.widthCm / block.heightCm,
  };
}

/** How many of the THUMBNAIL_MAX_LINES slots a drawing (image/chart/shape)
 * consumes: its share of the page height, so the budget still means "about
 * one miniature's worth". */
function drawingLineCost(
  model: DocumentModel,
  sectionIndex: number,
  block: Image | Chart | Shape,
): number {
  const pageHeightCm = sectionPageSize(model, sectionIndex).heightCm;
  const share = pageHeightCm > 0 ? block.heightCm / pageHeightCm : 0;
  const cost = Math.round(share * THUMBNAIL_MAX_LINES);
  return Math.min(THUMBNAIL_MAX_LINES, Math.max(1, Number.isFinite(cost) ? cost : 1));
}

/** A drawing's display width as a fraction of the section's content width,
 * clamped so it stays visible and never overflows the miniature. */
function drawingWidthFraction(
  model: DocumentModel,
  sectionIndex: number,
  block: Image | Chart | Shape,
): number {
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  return Math.min(1, Math.max(0.02, block.widthCm / contentWidthCm));
}

/** A table cell's text for the miniature — its paragraphs joined, trimmed
 * and capped (the card ellipsises anyway; the cap just keeps cached
 * thumbnails small). */
function cellPreviewText(cell: TableCell): string {
  return cell.blocks
    .map((block) => blockText(block))
    .join(" ")
    .trim()
    .slice(0, 60);
}

/** Resolves a table block to its thumbnail line, or null for a table with
 * nothing to show. */
function thumbnailTable(
  model: DocumentModel,
  sectionIndex: number,
  block: Table,
): ThumbnailTable | null {
  const rows = block.rows.slice(0, THUMBNAIL_TABLE_MAX_ROWS);
  if (rows.length === 0) {
    return null;
  }
  const columnCount = Math.min(
    THUMBNAIL_TABLE_MAX_COLUMNS,
    Math.max(1, block.columnWidthsCm.length),
  );
  const widthsCm = block.columnWidthsCm.slice(0, columnCount);
  const totalShownCm = widthsCm.reduce((sum, cm) => sum + Math.max(0, cm), 0);
  const columnFractions =
    totalShownCm > 0
      ? widthsCm.map((cm) => Math.max(0, cm) / totalShownCm)
      : widthsCm.map(() => 1 / columnCount);
  const totalCm = block.columnWidthsCm.reduce((sum, cm) => sum + Math.max(0, cm), 0);
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  const widthFraction =
    block.widthMode === "percent" && typeof block.widthPercent === "number"
      ? Math.min(1, Math.max(0.05, block.widthPercent / 100))
      : Math.min(1, Math.max(0.05, totalCm > 0 ? totalCm / contentWidthCm : 1));
  let headerRows = 0;
  while (headerRows < rows.length && rows[headerRows].isHeader) {
    headerRows += 1;
  }
  return {
    rows: rows.map((row) => row.cells.slice(0, columnCount).map(cellPreviewText)),
    headerRows,
    columnFractions,
    widthFraction,
  };
}

/** How many text-frame lines a shape miniature shows. */
const SHAPE_TEXT_MAX_LINES = 4;

/** Resolves a shape/text-box block to its thumbnail line, or null for a
 * degenerate one. A line/arrow's zero-ish cross axis is clamped so its
 * ratio and cost stay finite. */
function thumbnailShape(
  model: DocumentModel,
  sectionIndex: number,
  block: Shape,
): ThumbnailShape | null {
  const widthCm = Math.max(block.widthCm, 0.05);
  const heightCm = Math.max(block.heightCm, 0.05);
  if (!(widthCm > 0) || !(heightCm > 0)) {
    return null;
  }
  const visual = JSON.parse(
    JSON.stringify({
      geometry: block.geometry,
      fill: block.fill,
      outline: block.outline,
      rotationDeg: block.rotationDeg,
      adjustments: block.adjustments,
      unsupportedGeometryXml: block.unsupportedGeometryXml === undefined ? undefined : "",
    }),
  ) as ShapeVisual;
  const textLines = (block.text?.blocks ?? [])
    .filter(isParagraph)
    .map((paragraph) => blockText(paragraph).trim())
    .filter((text) => text.length > 0)
    .slice(0, SHAPE_TEXT_MAX_LINES)
    .map((text) => text.slice(0, 60));
  return {
    shape: visual,
    widthPx: Math.max(1, Math.round(cmToPx(widthCm))),
    heightPx: Math.max(1, Math.round(cmToPx(heightCm))),
    textLines,
    vAlign: block.text?.vAlign ?? "top",
    widthFraction: drawingWidthFraction(model, sectionIndex, block),
    aspectRatio: widthCm / heightCm,
  };
}

/** Resolves a chart block to its thumbnail line (a deep clone — the cache
 * must never alias the live model), or null for a degenerate one. */
function thumbnailChart(
  model: DocumentModel,
  sectionIndex: number,
  block: Chart,
): ThumbnailChart | null {
  if (!(block.widthCm > 0) || !(block.heightCm > 0)) {
    return null;
  }
  return {
    chart: JSON.parse(JSON.stringify(block)) as Chart,
    widthFraction: drawingWidthFraction(model, sectionIndex, block),
    aspectRatio: block.widthCm / block.heightCm,
  };
}

/** Extracts the leading styled lines of a document's first page. Walks blocks
 * in order — paragraphs as text (with heading level), images as data: URIs,
 * tables as capped text grids, charts and shapes as paintable clones for the
 * card's canvas — until the preview is full; whatever cannot be previewed
 * (a broken image, a degenerate drawing) is skipped rather than failing the
 * whole thumbnail. */
export function buildThumbnail(model: DocumentModel): Thumbnail {
  try {
    const lines: ThumbnailLine[] = [];
    let used = 0;
    const sections = model.sections ?? [];
    for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
      for (const block of sections[sectionIndex].blocks ?? []) {
        if (used >= THUMBNAIL_MAX_LINES) {
          return { lines, failed: false };
        }
        if (isImage(block)) {
          const image = thumbnailImage(model, sectionIndex, block);
          if (image) {
            lines.push({ text: "", level: 0, image });
            used += drawingLineCost(model, sectionIndex, block);
          }
          continue;
        }
        if (isTable(block)) {
          const table = thumbnailTable(model, sectionIndex, block);
          if (table) {
            lines.push({ text: "", level: 0, table });
            used += table.rows.length;
          }
          continue;
        }
        if (isChart(block)) {
          const chart = thumbnailChart(model, sectionIndex, block);
          if (chart) {
            lines.push({ text: "", level: 0, chart });
            used += drawingLineCost(model, sectionIndex, block);
          }
          continue;
        }
        if (isShape(block)) {
          const shape = thumbnailShape(model, sectionIndex, block);
          if (shape) {
            lines.push({ text: "", level: 0, shape });
            used += drawingLineCost(model, sectionIndex, block);
          }
          continue;
        }
        if (!isParagraph(block)) {
          continue;
        }
        const text = blockText(block).trim();
        if (text.length === 0) {
          continue;
        }
        lines.push({ text, level: block.headingLevel ?? 0 });
        used += 1;
      }
      if (used >= THUMBNAIL_MAX_LINES) {
        break;
      }
    }
    return { lines, failed: false };
  } catch {
    return { lines: [], failed: true };
  }
}

// Cache keyed by document id, storing the modified timestamp it was built for
// (SPEC-067): a thumbnail is produced once per edit, not once per view. A
// re-render for the same id overwrites the entry, so a stale timestamp is
// discarded automatically. The cache is purely in-memory and never persisted,
// so it costs no IndexedDB quota — quota pressure "evicts" it for free — and
// clearing it is always safe.
interface CacheEntry {
  lastModified: number;
  thumbnail: Thumbnail;
}

const cache = new Map<string, CacheEntry>();

/** The cached thumbnail for a document at a given modified time, or undefined
 * when none exists or the cached one is for an older edit (and must be
 * re-rendered). */
export function getCachedThumbnail(
  id: string,
  lastModified: number,
): Thumbnail | undefined {
  const entry = cache.get(id);
  if (entry && entry.lastModified === lastModified) {
    return entry.thumbnail;
  }
  return undefined;
}

/** Builds (or returns the cached) thumbnail for a document, keyed by id and
 * modified time. Overwrites any stale entry for the same id. */
export function renderThumbnail(
  id: string,
  lastModified: number,
  model: DocumentModel,
): Thumbnail {
  const cached = getCachedThumbnail(id, lastModified);
  if (cached) {
    return cached;
  }
  const thumbnail = buildThumbnail(model);
  cache.set(id, { lastModified, thumbnail });
  return thumbnail;
}

/** Drops one document's cached thumbnail. */
export function evictThumbnail(id: string): void {
  cache.delete(id);
}

/** Drops every cached thumbnail. Safe at any time — the cache is derived data. */
export function clearThumbnailCache(): void {
  cache.clear();
}
