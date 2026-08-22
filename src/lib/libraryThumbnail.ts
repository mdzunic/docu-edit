// Library thumbnails (SPEC-067): a lightweight rendered preview of a
// document's first page, produced from the model alone. The product's real
// page renderer (DocPage.vue + usePagination) is DOM-measurement driven and
// would have to mount the whole document off-screen — far too heavy to run
// per row of a 200-document library. Instead we extract the first page's
// leading blocks as a small list of styled text lines and let the card render
// them as a scaled-down paper miniature. It is a genuine preview of the first
// page's content, cheap enough to schedule through the idle scheduler.
//
// This module is the pure, framework-free part: it builds and caches thumbnail
// data. The IntersectionObserver gating and idle scheduling live in the
// LibraryThumbnail component, so only visible cards ever call in here.
import type { Block, DocumentModel } from "./documentModel";
import { blockText } from "./textEditing";

/** One line of a thumbnail preview. */
export interface ThumbnailLine {
  text: string;
  /** 0 for body text, 1..6 for a heading of that level — the card scales the
   * type so a document's shape (a big title over paragraphs) reads at a
   * glance. */
  level: number;
}

export interface Thumbnail {
  lines: ThumbnailLine[];
  /** True when the model could not be previewed (corrupt, or nothing
   * renderable). The card then shows a format-badge placeholder, never an
   * error. */
  failed: boolean;
}

/** How many leading lines a first-page preview shows before it is visually
 * "full". More than this never fits the miniature. */
export const THUMBNAIL_MAX_LINES = 16;

/** Paper aspect ratio (A4 portrait, width / height) the placeholder and the
 * rendered thumbnail both use, so nothing reflows when the real one arrives. */
export const THUMBNAIL_ASPECT_RATIO = 210 / 297;

function isParagraph(block: Block): block is Extract<Block, { type: "paragraph" }> {
  return block.type === "paragraph";
}

/** Extracts the leading styled lines of a document's first page. Walks blocks
 * in order, taking each paragraph's text (with its heading level) until the
 * preview is full; non-text blocks (images, tables) are skipped rather than
 * failing the whole thumbnail. */
export function buildThumbnail(model: DocumentModel): Thumbnail {
  try {
    const lines: ThumbnailLine[] = [];
    const sections = model.sections ?? [];
    for (const section of sections) {
      for (const block of section.blocks ?? []) {
        if (lines.length >= THUMBNAIL_MAX_LINES) {
          return { lines, failed: false };
        }
        if (!isParagraph(block)) {
          continue;
        }
        const text = blockText(block).trim();
        if (text.length === 0) {
          continue;
        }
        lines.push({ text, level: block.headingLevel ?? 0 });
      }
      if (lines.length >= THUMBNAIL_MAX_LINES) {
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
