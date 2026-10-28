// Pure, framework-free derivations for the library view (SPEC-067): the
// subsequence search + highlight, the format/starred filter, the counts the
// filter chips show, and the starred-first sort. Kept out of the composable so
// the search/sort/filter behaviour is unit-testable without any reactivity.
import type { DocumentFormat, LibraryEntry } from "./documentStore";
import type { LibrarySort } from "./libraryPrefs";

/** Every format the filter can act on, in display order. The media formats
 * (SPEC-107) are filterable but do NOT each get their own chip — see
 * FORMAT_CHIPS below. The workbook formats (SPEC-116) follow the decks; a
 * workbook's template or macro-enabled variant filters with its format.
 * SPEC-126's CSV and TSV follow the workbook packages. */
export const FILTERABLE_FORMATS: readonly DocumentFormat[] = [
  "docx",
  "odt",
  "md",
  "pptx",
  "odp",
  "xlsx",
  "ods",
  "csv",
  "tsv",
  "pdf",
  "png",
  "jpeg",
  "gif",
  "svg",
  "webp",
  "new",
];

/** The image formats one "Images" chip toggles together. */
export const IMAGE_FORMATS: readonly DocumentFormat[] = ["png", "jpeg", "gif", "svg", "webp"];

/** One filter chip: a label over one or more formats. The document formats
 * keep a chip each (id === format, so their testids are stable); the media
 * formats group by KIND — "Images" and "PDFs" are one click each
 * (SPEC-107). */
export interface FormatChip {
  id: string;
  label: string;
  formats: readonly DocumentFormat[];
}

export const FORMAT_CHIPS: readonly FormatChip[] = [
  ...(["docx", "odt", "md", "pptx", "odp", "xlsx", "ods", "csv", "tsv"] as const).map((format) => ({
    id: format,
    label: format.toUpperCase(),
    formats: [format] as readonly DocumentFormat[],
  })),
  { id: "pdf", label: "PDFs", formats: ["pdf"] },
  { id: "images", label: "Images", formats: IMAGE_FORMATS },
  { id: "new", label: "New", formats: ["new"] },
];

/** Whether a chip reads active: every format it stands for is toggled on. */
export function chipActive(chip: FormatChip, active: ReadonlySet<DocumentFormat>): boolean {
  return chip.formats.every((format) => active.has(format));
}

/** The next active-format set after toggling a chip: all of its formats off
 * when the chip is fully active, all on otherwise. */
export function toggleChip(
  chip: FormatChip,
  active: ReadonlySet<DocumentFormat>,
): Set<DocumentFormat> {
  const next = new Set(active);
  if (chipActive(chip, active)) {
    for (const format of chip.formats) {
      next.delete(format);
    }
  } else {
    for (const format of chip.formats) {
      next.add(format);
    }
  }
  return next;
}

/** A chip's live count: the sum of its formats' counts. */
export function chipCount(chip: FormatChip, counts: Record<DocumentFormat, number>): number {
  return chip.formats.reduce((sum, format) => sum + counts[format], 0);
}

/** Whether a chip should be unavailable: it stands for nothing in the library
 * and is not currently on.
 *
 * A library with one document offered nine filters, eight of them reading "0",
 * every one of them a click that could only ever produce "No documents match
 * your search or filters" — nine advertised routes, eight of them dead ends.
 * An ACTIVE chip stays enabled whatever it counts, or a filter that empties
 * the list could not be switched back off. */
export function chipUnavailable(
  chip: FormatChip,
  active: ReadonlySet<DocumentFormat>,
  counts: Record<DocumentFormat, number>,
): boolean {
  return chipCount(chip, counts) === 0 && !chipActive(chip, active);
}

/** A run of a title, flagged as matched (a query character) or not, so the
 * row/card can highlight the matched characters (AC-9). */
export interface HighlightSegment {
  text: string;
  match: boolean;
}

/** Typed pass-through for highlight segments arriving through a `.vue`
 * component's runtime-typed Array prop (espree bars the `PropType`
 * annotation — the findReplace.asHighlightRects idiom). SPEC-114's snippet
 * segments share the shape. */
export function asHighlightSegments(value: unknown): HighlightSegment[] {
  return value as HighlightSegment[];
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
    pptx: 0,
    odp: 0,
    xlsx: 0,
    ods: 0,
    csv: 0,
    tsv: 0,
    pdf: 0,
    png: 0,
    jpeg: 0,
    gif: 0,
    svg: 0,
    webp: 0,
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
  return [...entries].toSorted((a, b) => {
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
