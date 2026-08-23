// Pure, framework-free derivations for the library view (SPEC-067): the
// subsequence search + highlight, the format/starred filter, the counts the
// filter chips show, and the starred-first sort. Kept out of the composable so
// the search/sort/filter behaviour is unit-testable without any reactivity.
import type { DocumentFormat, LibraryEntry } from "./documentStore";
import type { LibrarySort } from "./libraryPrefs";

/** The document formats that get a filter chip, in display order. */
export const FILTERABLE_FORMATS: readonly DocumentFormat[] = [
  "docx",
  "odt",
  "md",
  "new",
];

/** A run of a title, flagged as matched (a query character) or not, so the
 * row/card can highlight the matched characters (AC-9). */
export interface HighlightSegment {
  text: string;
  match: boolean;
}

/** Case-insensitive subsequence match (AC-8). Returns the indices in `text`
 * that the query's characters matched, in order — or null when `text` does
 * not contain the query as a subsequence. An empty query matches everything
 * (returns no indices). */
export function subsequenceMatch(query: string, text: string): number[] | null {
  if (query.length === 0) {
    return [];
  }
  const q = query.toLowerCase();
  const t = text.toLowerCase();
  const indices: number[] = [];
  let qi = 0;
  for (let ti = 0; ti < t.length && qi < q.length; ti += 1) {
    if (t[ti] === q[qi]) {
      indices.push(ti);
      qi += 1;
    }
  }
  return qi === q.length ? indices : null;
}

/** Splits `text` into matched/unmatched segments for highlighting, given the
 * matched indices from subsequenceMatch. Consecutive matched characters are
 * merged into one segment. */
export function highlightSegments(
  text: string,
  matchedIndices: number[],
): HighlightSegment[] {
  if (matchedIndices.length === 0) {
    return text.length === 0 ? [] : [{ text, match: false }];
  }
  const matched = new Set(matchedIndices);
  const segments: HighlightSegment[] = [];
  let current = "";
  let currentMatch = matched.has(0);
  for (let i = 0; i < text.length; i += 1) {
    const isMatch = matched.has(i);
    if (isMatch === currentMatch) {
      current += text[i];
    } else {
      if (current.length > 0) {
        segments.push({ text: current, match: currentMatch });
      }
      current = text[i]!;
      currentMatch = isMatch;
    }
  }
  if (current.length > 0) {
    segments.push({ text: current, match: currentMatch });
  }
  return segments;
}

export interface LibraryFilter {
  /** The search query; matched by subsequence against the title. */
  search: string;
  /** Active format chips; empty means every format passes. */
  formats: ReadonlySet<DocumentFormat>;
  /** When true, only starred documents pass. */
  starredOnly: boolean;
}

/** Applies only the search filter — the basis for the chip counts, which show
 * how many documents of each facet match the current search (AC-16, AC-17). */
export function searchLibrary(
  entries: readonly LibraryEntry[],
  search: string,
): LibraryEntry[] {
  const query = search.trim();
  if (query.length === 0) {
    return [...entries];
  }
  return entries.filter((entry) => subsequenceMatch(query, entry.title) !== null);
}

/** Applies the search, format and starred filters together (AC-8, AC-16,
 * AC-17). */
export function filterLibrary(
  entries: readonly LibraryEntry[],
  filter: LibraryFilter,
): LibraryEntry[] {
  return searchLibrary(entries, filter.search).filter((entry) => {
    if (filter.starredOnly && !entry.starred) {
      return false;
    }
    if (filter.formats.size > 0 && !filter.formats.has(entry.format)) {
      return false;
    }
    return true;
  });
}

/** Live counts for the filter chips: for each format, and for starred, how
 * many documents match the current search (AC-16, AC-17). Computed over the
 * search-filtered set so the counts move as the user types, but independent of
 * which chips are toggled. */
export function filterCounts(
  entries: readonly LibraryEntry[],
  search: string,
): { formats: Record<DocumentFormat, number>; starred: number } {
  const searched = searchLibrary(entries, search);
  const formats: Record<DocumentFormat, number> = {
    docx: 0,
    odt: 0,
    md: 0,
    new: 0,
  };
  let starred = 0;
  for (const entry of searched) {
    formats[entry.format] += 1;
    if (entry.starred) {
      starred += 1;
    }
  }
  return { formats, starred };
}

/** Sorts entries by the chosen key and direction, with starred documents
 * always first (AC-11..AC-14, AC-21). `sizeOf` supplies the byte size for the
 * "size" sort. Stable and deterministic: ties break by newest-modified, then
 * id. */
export function sortLibrary(
  entries: readonly LibraryEntry[],
  sort: LibrarySort,
  sizeOf: (id: string) => number,
): LibraryEntry[] {
  const dir = sort.direction === "asc" ? 1 : -1;
  return [...entries].sort((a, b) => {
    if (a.starred !== b.starred) {
      return a.starred ? -1 : 1;
    }
    let cmp = 0;
    if (sort.key === "title") {
      cmp = a.title.localeCompare(b.title, undefined, { sensitivity: "base" });
    } else if (sort.key === "size") {
      cmp = sizeOf(a.id) - sizeOf(b.id);
    } else {
      cmp = a.lastModified - b.lastModified;
    }
    if (cmp !== 0) {
      return cmp * dir;
    }
    if (a.lastModified !== b.lastModified) {
      return b.lastModified - a.lastModified;
    }
    return a.id.localeCompare(b.id);
  });
}
