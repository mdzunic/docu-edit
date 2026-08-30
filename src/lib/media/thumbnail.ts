// The pure half of media library thumbnails (SPEC-107): downscale math, the
// per-record cache (SPEC-067's rules: keyed by id + lastModified, in-memory
// only so quota pressure evicts it for free), and the PDF placeholder's
// figures. The DOM half — decoding the image and painting the downscaled
// raster — lives in the LibraryThumbnail component behind its existing
// IntersectionObserver + idle-scheduler gating.
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

/** The details panel's per-page size line: one size when every page shares
 * it, else "mixed" qualified with the first page's size (AC: per-page size
 * or "mixed" with the first page's size). */
export function pdfPageSizesLabel(pdf: PdfSummary): string {
  const first = pdf.pages[0];
  if (!first) {
    return "";
  }
  const uniform = pdf.pages.every(
    (page) => page.widthPt === first.widthPt && page.heightPt === first.heightPt,
  );
  return uniform ? pageSizeLabel(first) : `mixed (first ${pageSizeLabel(first)})`;
}

/** What the PDF card's paper-shaped placeholder shows: the REAL aspect
 * ratio of the first page (rotation applied, so a rotated page reads
 * landscape) plus the page count. No page is ever rasterised — a card never
 * fakes a preview it cannot produce. */
export function pdfCardPlaceholder(
  pdf: PdfSummary,
): { aspectRatio: number; pageCount: number; pageSize: string } {
  const first = pdf.pages[0] ?? { widthPt: 612, heightPt: 792, rotate: 0 as const };
  const swapped = first.rotate === 90 || first.rotate === 270;
  const width = swapped ? first.heightPt : first.widthPt;
  const height = swapped ? first.widthPt : first.heightPt;
  return {
    aspectRatio: height > 0 ? width / height : 1,
    pageCount: pdf.pageCount,
    pageSize: pageSizeLabel(first),
  };
}

/** The image an <img> element should thumbnail from: the media document's
 * stored payload as a data: URI. Kept here (not in the component) so the
 * .vue file needs no base64 knowledge. */
export function mediaImageDataUri(model: MediaModel): string {
  const part = model.media.find((entry) => entry.mediaId === model.sourceMediaId);
  return part ? `data:${part.contentType};base64,${part.data}` : "";
}
