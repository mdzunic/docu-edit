// Cheap, model-only statistics the library shows per document (SPEC-067):
// a word count and an estimated page count. Kept in src/lib as framework-free
// logic so it is unit-testable without mounting a component.
//
// Page count here is deliberately an ESTIMATE, not a real layout pass. True
// pagination (usePagination) is DOM-measurement driven and needs the document
// mounted on screen — far too heavy to run for every row of a 200-document
// library (SPEC-067's render budget). Words-per-page is the same approximation
// a word processor's status bar uses when it has not yet paginated.
import type { DocumentModel } from "./documentModel";
import { documentBodyStats } from "./wordCount";

/** Words that fit on a page for the estimate. A full A4 page of body text is
 * roughly 500 words; the figure only needs to be stable and plausible, not
 * exact — the card says "3 pages", not a layout guarantee. */
export const WORDS_PER_PAGE = 500;

export interface LibraryStats {
  /** Body word count (headers, footers and notes excluded). */
  words: number;
  /** Estimated page count, always at least 1 (an empty document is 1 page). */
  pages: number;
}

/** The word count and estimated page count for a document's body. */
export function documentLibraryStats(model: DocumentModel): LibraryStats {
  const words = documentBodyStats(model).words;
  const pages = Math.max(1, Math.ceil(words / WORDS_PER_PAGE));
  return { words, pages };
}

// In-memory cache keyed by id + modified time (SPEC-067), so a card that
// scrolls out and back does not re-walk the model. Derived data only; safe to
// clear at any time and never persisted.
interface CachedStats {
  lastModified: number;
  stats: LibraryStats;
}

const statsCache = new Map<string, CachedStats>();

/** documentLibraryStats with per-document caching, keyed by modified time so a
 * fresh edit recomputes. */
export function getDocumentLibraryStats(
  id: string,
  lastModified: number,
  model: DocumentModel,
): LibraryStats {
  const hit = statsCache.get(id);
  if (hit && hit.lastModified === lastModified) {
    return hit.stats;
  }
  const stats = documentLibraryStats(model);
  statsCache.set(id, { lastModified, stats });
  return stats;
}

/** Drops every cached stat. */
export function clearLibraryStatsCache(): void {
  statsCache.clear();
}
