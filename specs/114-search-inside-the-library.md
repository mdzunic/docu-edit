---
id: SPEC-114
title: Search inside the library (experimental)
status: draft
area: shell
depends_on: [SPEC-017, SPEC-047, SPEC-067, SPEC-098, SPEC-107, SPEC-113]
---
## Why
The library's search field matches titles. That is the right default and a
poor ceiling: nobody remembers what they called the file, they remember the
sentence in it. "The one with the Rotterdam figures" is how people look for
documents, and this product — which holds every byte locally and can already
walk a document model, a slide outline and a PDF's object graph — is
unusually well placed to answer it without sending a word anywhere.

It is filed as experimental for two honest reasons. It needs an index, and an
index is a second copy of the user's text with a storage cost and a staleness
failure mode. And PDF text extraction has a hard ceiling: a scanned page has
no text to find, and a PDF whose fonts carry no `/ToUnicode` map cannot be
decoded to characters without a renderer. A search that silently returns
nothing for those files would be worse than no search at all, so this spec
requires it to say which documents it could not read, and why.

**This file is the normative source for the search index, the tokeniser, the
query grammar, the ranking function and `core/pdfText.ts`.**

## Acceptance criteria
- [ ] The feature is gated by SPEC-113's `library-search` experiment: with the flag off the scope control does not render, no index is built, no index record is read or written, and the library behaves exactly as SPEC-067 specifies — asserted as a test
- [ ] `src/lib/core/searchIndex.ts` is the product-neutral engine: `tokenise(text)`, `buildIndex(docs)`, `mergeIndex`, `parseQuery(input): Query | QueryError`, `search(index, query, options): SearchHit[]` and `snippet(text, hit, options): SnippetSegment[]` — pure, framework-free, no DOM, and knowing nothing about documents, decks or PDFs beyond the `{ id, fields: { title, body } }` it is handed
- [ ] Tokenisation is specified, not incidental: lowercase, Unicode NFKD-normalised with combining marks stripped (so `café` matches `cafe`), split on anything that is not a Unicode letter, digit or `'`/`-` inside a word, numbers kept whole, tokens capped at 64 characters, and any run of CJK/Kana indexed as overlapping **bigrams** so a language without spaces is searchable at all — each rule with its own test, because a tokeniser that changes silently invalidates every stored index
- [ ] The query grammar is small and total: space-separated terms are AND, `"a phrase"` matches adjacent tokens, a leading `-` excludes, `title:`, `kind:` and `format:` filter, the **last** term is a prefix match while the user is typing (so `rotter` finds `Rotterdam`) and an explicit trailing space ends that, and anything unparseable returns a `QueryError` with the offending position rather than throwing or matching everything
- [ ] Ranking is a stated formula, so results are reproducible and testable: BM25 with `k1 = 1.2`, `b = 0.75` over the body field, the title field scored with the same formula and multiplied by `TITLE_BOOST = 3`, summed across terms, with ties broken by `lastModified` newest-first and then by id — no learned weights, no personalisation, no "smart" reordering
- [ ] Extraction is per product and each is honest about its ceiling: `src/lib/documentText.ts` walks a `DocumentModel`'s blocks (paragraphs, tables, notes, comments, headers/footers) in document order; `src/lib/presentation/deckText.ts` takes each slide's title, paragraph text and SPEC-083 notes; `src/lib/core/pdfText.ts` extracts what a PDF's content streams can be decoded to; and each returns `{ text, coverage: "full" | "partial" | "none", reason? }` rather than an empty string that lies
- [ ] `core/pdfText.ts` is specified to its limits per the Reference: it tokenises content streams for the text operators (`BT`/`ET`, `Tf`, `Td`/`TD`/`Tm`/`T*`, `Tj`/`TJ`/`'`/`"`), maps bytes to characters through `/ToUnicode` where the font has one, otherwise through `/Encoding` + `/Differences` over WinAnsi / MacRoman / Standard for simple fonts, treats a Type0/CID font with no `/ToUnicode` as **unextractable for that page**, inserts a space where a `TJ` offset exceeds `WORD_GAP_THOUSANDTHS = 200` and a newline where the text-line matrix moves down, and never rasterises anything
- [ ] Coverage is surfaced, not swallowed: a document whose extraction is `none` appears in a "Not searched" list (`search-unsearched`, expandable, with a per-document reason — "scanned image, no text layer", "fonts carry no character map", "over the size ceiling") and is never silently absent; a PDF this product produced through SPEC-034 or SPEC-108 is always `full`, which is asserted, and a SPEC-110 image assembly is always `none`, which is also asserted
- [ ] Indexing never blocks the app and never blocks the library: it runs in a Web Worker through the existing `src/lib/workers/workerRpc.ts` (the SPEC-035/036 pattern), scheduled off `runWhenIdle`, one document at a time, newest first; the library renders and is fully usable while it runs, `search-index-progress` reports `Indexed {done} of {total}` with `role="status"`, and closing the tab mid-build loses only the unindexed remainder
- [ ] The index is a cache with a stamp, not a source of truth: it is persisted per document under an `INDEX_KEY_PREFIX = "index:"` record carrying `{ docId, indexVersion, lastModified, tokens }`, invalidated whenever `lastModified` moves or `INDEX_VERSION` (bumped by any tokeniser change) does not match, evicted with thumbnails under quota pressure and by `deleteDocument`, excluded for trashed documents, and rebuildable from scratch by "Rebuild index" in Labs — losing it costs time and nothing else
- [ ] Bounds are stated and enforced: `MAX_INDEXED_CHARS = 2_000_000` per document (truncated, coverage `partial`, with the reason), `MAX_INDEX_BYTES = 64 * 1024 * 1024` across the library (the oldest-used records evicted first, and the "Not searched" list saying so), and `MAX_PDF_TEXT_PAGES = 2000`
- [ ] The library search field gains a scope control (`search-scope`, `Titles` default / `Everything`) and, in `Everything`, renders results as a grouped list (`search-results`, `search-result` per document) showing the entry's title, kind badge, match count and up to `MAX_SNIPPETS_PER_DOC = 3` snippets with the matched terms marked (`search-snippet`, reusing SPEC-067's highlight rendering), keyboard-navigable with the same arrow/Enter conventions as the rows they replace
- [ ] A result opens **at** its match: a text document at `/doc/:id?find=<query>` opening SPEC-017's find bar with the query compiled and the nth match selected, a deck at `/slides/:id?find=<query>` through SPEC-098's deck find, and a PDF at `/media/:id?find=<query>` which shows the query above the viewer with a copy button and appends `#search=` to the embed URL — a fragment only some browsers honour, which the UI states rather than implies
- [ ] Search is fast enough to type into: over a warm index of 200 documents totalling 20 MB of text, a query returns in under 50 ms measured in the same harness SPEC-047 uses, input is debounced at `SEARCH_DEBOUNCE_MS = 150`, an in-flight search is abandoned when the query changes, and the result list renders at most 50 documents with "Show more" rather than all of them
- [ ] Nothing leaves the browser: this feature makes no network request of any kind and does not use SPEC-113's model gateway — it depends on SPEC-113 **only** for the flag registry, which SPEC-113's boundary test already enforces
- [ ] Tests: `src/lib/core/searchIndex.spec.ts` — the tokeniser rule by rule (diacritics, CJK bigrams, hyphenation, the 64-character cap), the query grammar including every error case, BM25 scores against hand-computed values, tie-breaking, phrase and exclusion matching, and snippet windows around a match at the start, middle and end of a text; `src/lib/core/pdfText.spec.ts` — `/ToUnicode` decoding, `/Differences` over each base encoding, the CID-without-map refusal, `TJ` word gaps, line breaks, and the round-trip guarantee that a `buildPdf` output re-extracts the strings that went in; `src/lib/libraryIndex.spec.ts` — staleness invalidation, eviction order, trashed exclusion, worker fallback when workers are unavailable; e2e (`e2e/SPEC-114.spec.ts`): enable the flag → open a library holding a `.docx`, a `.pptx`, a text PDF and `spec107-photo.jpg` → switch scope to Everything → wait for the progress line to finish → search a phrase that appears in the docx body → one result with a snippet → Enter opens `/doc/:id` with the find bar carrying the query and the match selected → search a word in a slide → the deck result opens `/slides/:id` at the slide → the image and the scanned fixture appear under "Not searched" with their reasons

## Out of scope
Fuzzy matching, stemming, lemmatisation, synonyms, spelling correction and
"did you mean" — the tokeniser's normalisation is the whole tolerance, and
each of those is a language-data dependency. Semantic or embedding-based
search, and anything model-backed (SPEC-113's gateway is not used here).
Regular-expression search across the library (SPEC-017 owns regex within one
document). Searching inside media that has no text: an image, an animated
GIF, an SVG's text nodes, or a scanned PDF — OCR is not in this product.
Searching document *properties*, comments-only views, or track-changes
history as separate scopes. Replacing across the library. Saved searches,
search history and search-driven smart folders. Indexing an external folder
or anything not in the library. Ranking that learns from what the user opens.

## Reference

### Dependencies restated
- `specs/017-find-and-replace.md` — `src/lib/findReplace.ts`:
  `compileQuery(text, options): CompiledQuery | QueryError`,
  `findMatches(model, compiled): FindMatch[]`, `matchSelection(match)` and
  `INVALID_REGEX_MESSAGE`; the find bar is the surface a text result opens
  into, reached with the query pre-filled — **this spec adds a query
  parameter to that route and changes nothing about the find bar itself**.
- `specs/098-…` — the deck-side find surface a presentation result opens into.
- `specs/067-…` — the library header's search field, `subsequenceMatch` /
  `highlightSegments` (which stay the title-scope matcher — full-text scope
  uses the index instead), and the row/card keyboard conventions results
  inherit.
- `specs/047-large-document-performance.md` — `runWhenIdle`/
  `yieldToMainThread` and the rule that background work never moves the
  interaction baseline; `src/lib/workers/workerRpc.ts` is the existing
  worker transport (spell check and word count already ride it).
- `specs/107-…` — `parsePdf`, `PdfDocument`, `resolve`, `streamData` (which
  already inflates `/FlateDecode` with predictors — the only decoding
  `pdfText.ts` needs, since it reads content streams and not images), and the
  media kind's route.
- `specs/113-…` — `isExperimentEnabled("library-search")` and the Labs
  section this feature adds "Rebuild index" and "Clear index" to.

### The index
```ts
// src/lib/core/searchIndex.ts
export const INDEX_VERSION = 1;    // bumped by ANY tokeniser or scoring change

export interface IndexedDoc { id: string; lastModified: number; title: string; body: string }
export interface SearchHit {
  id: string;
  score: number;
  /** Character offsets into the body, in order, for snippets and "open at match". */
  positions: number[];
  titleMatches: number;
}
export interface SnippetSegment { text: string; match: boolean }

export function tokenise(text: string): Array<{ token: string; start: number }>;
export function buildIndex(docs: IndexedDoc[]): SearchIndex;
export function search(index: SearchIndex, query: Query, options?: { limit?: number }): SearchHit[];
export function snippet(body: string, hit: SearchHit, options?: { radius?: number }): SnippetSegment[];
```
The persisted per-document record holds the document's own posting lists, so
one document re-indexing rewrites one record; the searchable index is the
merge of the loaded records, held in the worker. `SearchIndex` is plain JSON
(the store's standing requirement).

### PDF text, and its ceiling
`core/pdfText.ts` reads content streams — it never renders. What it can do:
decode a simple font's bytes through `/ToUnicode`, or through `/Encoding`'s
base plus `/Differences` against the three standard tables; reconstruct word
and line breaks from `TJ` offsets and the text-line matrix; and concatenate
pages in order.

What it cannot do, stated once so no later spec assumes otherwise: a Type0/CID
font with no `/ToUnicode` map has no byte-to-character relation available
outside the font program itself, so those pages are `none`; a scanned page has
no text operators at all, so it is `none`; and glyphs positioned by a
renderer's own metrics may join or split words differently than a viewer
shows. Coverage is reported per document and the "Not searched" list is where
that honesty is visible to the user. **The guarantee this product can make is
narrower and exact: every PDF it writes itself carries a `/ToUnicode` CMap
(SPEC-034 already does), so its own output is always fully searchable** — and
a PDF assembled from photographs contains no text and says so.

### Tests
- `src/lib/core/searchIndex.spec.ts`, `src/lib/core/pdfText.spec.ts`,
  `src/lib/libraryIndex.spec.ts` as in the acceptance criteria.
- `src/lib/documentText.spec.ts` / `presentation/deckText.spec.ts` — block
  order, table cell order, notes and comments included, and the character
  offsets they report lining up with what `matchSelection` needs.
- `src/composables/useLibrarySearch.spec.ts` — scope switching, debouncing,
  abandoning an in-flight search, the 50-document render cap, the flag-off
  path loading nothing.
- `e2e/SPEC-114.spec.ts` as in the acceptance criteria.

### Specs this changes
- `specs/067-…` — the search field gains a scope control and a second result
  rendering; title-scope behaviour is untouched.
- `specs/017-…` / `specs/098-…` — each editor route accepts a `find` query
  parameter that pre-fills its existing find surface. No find behaviour
  changes.
- `specs/107-…` — `/media/:id` accepts the same parameter and displays it.
- `specs/113-…` — `EXPERIMENTS` gains the `library-search` entry, and Labs
  gains "Rebuild index" / "Clear index".
- `specs/000-roadmap.md` gains this spec's phase row.
