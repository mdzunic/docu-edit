// The pure half of media library thumbnails (SPEC-107): downscale math, the
// per-record cache (SPEC-067's rules: keyed by id + lastModified, in-memory
// only so quota pressure evicts it for free), and the PDF placeholder's
// figures. The DOM half — decoding the image and painting the downscaled
// raster — lives in the LibraryThumbnail component behind its existing
// IntersectionObserver + idle-scheduler gating.
import { formatLength, ptToCm, type MeasurementUnit } from "../core/units";
import { ensureMediaEncoded } from "../core/media";
import { currentImagePart } from "./editModel";
import type { MediaModel, PdfSummary } from "./model";

/** The longest edge a cached media thumbnail raster may have. */
export const MEDIA_THUMBNAIL_MAX_EDGE_PX = 512;

/** The size a source image downscales to for its thumbnail: capped at
 * MEDIA_THUMBNAIL_MAX_EDGE_PX on the longer edge, never upscaled, aspect
 * ratio preserved, and at least 1×1. */
export function thumbnailRasterSize(
  widthPx: number,
  heightPx: number,
  maxEdgePx: number = MEDIA_THUMBNAIL_MAX_EDGE_PX,
): { widthPx: number; heightPx: number } {
  const width = Math.max(1, Math.floor(widthPx));
  const height = Math.max(1, Math.floor(heightPx));
  const longest = Math.max(width, height);
  if (longest <= maxEdgePx) {
    return { widthPx: width, heightPx: height };
  }
  const scale = maxEdgePx / longest;
  return {
    widthPx: Math.max(1, Math.round(width * scale)),
    heightPx: Math.max(1, Math.round(height * scale)),
  };
}

// Cache keyed by document id, storing the modified timestamp the raster was
// built for — a thumbnail is produced once per record, not once per view
// (the same contract as libraryThumbnail.ts's cache for text documents).
interface CacheEntry {
  lastModified: number;
  dataUrl: string;
}

const cache = new Map<string, CacheEntry>();

/** The cached downscaled raster (as a data: URL) for a media document at a
 * given modified time, or undefined when it must be (re)built. */
export function getCachedMediaThumbnail(id: string, lastModified: number): string | undefined {
  const entry = cache.get(id);
  return entry && entry.lastModified === lastModified ? entry.dataUrl : undefined;
}

/** Stores a freshly built raster, overwriting any stale entry for the id. */
export function cacheMediaThumbnail(id: string, lastModified: number, dataUrl: string): void {
  cache.set(id, { lastModified, dataUrl });
}

/** Drops one media document's cached thumbnail. */
export function evictMediaThumbnail(id: string): void {
  cache.delete(id);
}

/** Drops every cached media thumbnail. Safe at any time — derived data. */
export function clearMediaThumbnailCache(): void {
  cache.clear();
}

/** A page's size label, e.g. "612 × 792 pt". */
export function pageSizeLabel(page: { widthPt: number; heightPt: number }): string {
  return `${Math.round(page.widthPt)} × ${Math.round(page.heightPt)} pt`;
}

/** A page's size in a SPEC-064 measurement unit, e.g. "21 × 29.7 cm". */
export function pageSizeInUnitLabel(page: { widthPt: number; heightPt: number }, unit: MeasurementUnit): string {
  return `${formatLength(ptToCm(page.widthPt), unit)} × ${formatLength(ptToCm(page.heightPt), unit)} ${unit}`;
}

/** The details panel's per-page size line: one size when every page shares
 * it, else "mixed" qualified with the first page's size (AC: per-page size
 * or "mixed" with the first page's size). Given a measurement unit, the
 * size also reads in it — "595 × 842 pt · 21 × 29.7 cm" (SPEC-108). */
export function pdfPageSizesLabel(pdf: PdfSummary, unit?: MeasurementUnit): string {
  const first = pdf.pages[0];
  if (!first) {
    return "";
  }
  const uniform = pdf.pages.every(
    (page) => page.widthPt === first.widthPt && page.heightPt === first.heightPt,
  );
  const size = unit && unit !== "pt" ? `${pageSizeLabel(first)} · ${pageSizeInUnitLabel(first, unit)}` : pageSizeLabel(first);
  return uniform ? size : `mixed (first ${size})`;
}

/** What the PDF card's paper-shaped placeholder shows wherever no first-page
 * raster is available: the REAL aspect ratio of the first page (rotation
 * applied, so a rotated page reads landscape) plus the page count. The card
 * sizes its paper from this ratio, so it is always positive and finite. */
export function pdfCardPlaceholder(
  pdf: PdfSummary,
): { aspectRatio: number; pageCount: number; pageSize: string } {
  const first = pdf.pages[0] ?? { widthPt: 612, heightPt: 792, rotate: 0 as const };
  const swapped = first.rotate === 90 || first.rotate === 270;
  const width = swapped ? first.heightPt : first.widthPt;
  const height = swapped ? first.widthPt : first.heightPt;
  return {
    // A zero-area page has no shape to be true to: it draws square.
    aspectRatio: width > 0 && height > 0 ? width / height : 1,
    pageCount: pdf.pageCount,
    pageSize: pageSizeLabel(first),
  };
}

/** The image an <img> element should thumbnail from: the media document's
 * current image — its stored payload, or for an edited image (SPEC-112) its
 * saved render — as a data: URI ("" when there is none). Kept here (not in
 * the component) so the .vue file needs no base64 knowledge. */
export function mediaImageDataUri(model: MediaModel): string {
  const part = currentImagePart(model);
  ensureMediaEncoded(part);
  return part && part.data ? `data:${part.contentType};base64,${part.data}` : "";
}
