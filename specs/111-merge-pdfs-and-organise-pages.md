---
id: SPEC-111
title: Merge PDFs, and organise their pages
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-107, SPEC-108, SPEC-109, SPEC-110]
---
## Why
Three PDFs that should be one. A twelve-page scan with page 7 upside down and
page 3 scanned twice. A contract that needs the appendix someone sent
separately, and the signature page photographed on a phone. Combining and
tidying PDFs is the most-searched-for document operation there is, and every
answer is a website you upload your contract to.

SPEC-107 gave this product a structural PDF reader — objects, cross-reference
tables and streams, the page tree, page geometry — and SPEC-110 taught it to
write a PDF from images. Merging is what those two make possible: copy the
reachable object graph of each chosen page into one new file, renumber, build
a new page tree, and emit. No page is interpreted, no glyph is decoded, and
**every content stream is copied byte-for-byte** — which is exactly the
round-trip value this repo already holds for `.docx` and `.odt`, applied to a
third format.

**This file is the normative source for `assemblePdf`, for the page-level
operations (reorder, rotate, delete, extract, insert), for what survives a
merge and what is dropped, and for the page organiser.**

## Acceptance criteria
- [ ] `src/lib/core/pdfDocument.ts` gains the writer half per the Reference: `assemblePdf(plan: PdfAssemblyPlan): Promise<PdfAssemblyResult>` copying the reachable object graph of each selected page into one output, renumbering objects from 1, materialising every inherited page attribute onto the copied page dictionary, rebuilding `/Pages`, `/Root` and the trailer, and writing a classic cross-reference table with a PDF 1.7 header
- [ ] Streams are copied, never re-encoded: a copied stream keeps its `/Filter`, `/DecodeParms` and its raw bytes verbatim, so merging one source with no page changes produces an output whose per-page content streams are `equalBytes` with the source's — asserted directly, and the reason the merge is safe on files this product cannot render
- [ ] What survives is enumerated, not hoped for, per "The merge contract": page content, resources, XObjects, fonts, colour spaces, shadings and patterns (by graph reachability); page `/Annots` including `/Link` and `/Widget` with their appearance streams; `/Rotate`, `/MediaBox`, `/CropBox` and the other box entries; optional content groups with their `/OCProperties` merged; and per-source outline trees, re-parented under one top-level entry per source
- [ ] What is dropped is reported, never silently lost: `PdfDropRecord { kind: PdfDropKind; source: string; page?: number; reason: string }` over the closed union `acroForm | structureTree | javascript | openAction | embeddedFile | namedDestination | outlineEntry | annotation | pageLabel | encryptedSource | unreadableSource`, returned from `assemblePdf` and surfaced in the UI as a "What changed" disclosure (`merge-drops`, one row per record, grouped by kind with counts) — a merge that drops nothing shows nothing
- [ ] Interactive forms degrade honestly: `/Widget` annotations are **kept** with their `/AP` appearance streams, so a filled form still looks filled, while the catalog's `/AcroForm` is dropped and one `acroForm` record per source explains that fields are no longer fillable — the alternative (dropping the widgets) silently erases visible content, and merging two `/AcroForm` field trees would need name-collision rules this spec is not the place for
- [ ] Links keep working where they can: an internal `/GoTo` destination is remapped to the copied page's new object; a destination naming a page that is not in the output makes the annotation inert (removed) with a `namedDestination` or `annotation` record; an external `/URI` link is copied unchanged; and `/Dests` name-tree entries are merged with a source-qualified prefix on collision
- [ ] An encrypted source is refused before anything else runs (SPEC-107's `PdfEncryptedError`) with a message naming the file, and a source that fails to parse is refused the same way; neither aborts a plan silently, and a plan whose remaining sources are all unusable produces no file and no library record
- [ ] The page organiser is the SPEC-109 workspace with PDFs as sources: adding a PDF to the tray expands it into **page** items (`assembly-page-card`, showing page number, source badge, page size, rotation and a `page-selected` state), and the tray gains `page-rotate-left`/`page-rotate-right` (adding ±90 to `/Rotate` mod 360), `page-delete`, `page-duplicate`, range selection (`Shift`+click, `Cmd/Ctrl`+click, `Cmd/Ctrl+A`), drag and `Alt+↑`/`Alt+↓` reordering, and "Select all pages from this source"
- [ ] Page cards carry no rasterised preview and say so rather than faking one: each card is a paper-shaped placeholder at the page's real aspect with its number and source, and verification happens in the **result preview** — `merge-preview` renders the currently planned output in the browser's own PDF viewer (SPEC-107's `<object>` + `blob:` URL) behind a `merge-refresh` action, rebuilt at most once every `PREVIEW_DEBOUNCE_MS = 750` and never automatically for a plan over `PREVIEW_MAX_PAGES = 200`
- [ ] Images and PDF pages mix in one plan: an image item in a PDF-output assembly is built into a one-page PDF through SPEC-110's `buildImagePdf` on the current sheet, parsed with `parsePdf`, and merged as an ordinary source — so "insert this photo of the signature page after page 4" is the same operation as inserting a PDF page, with no second code path; a `assembly-blank-page` action inserts an empty page at the current sheet size
- [ ] The three page-level outputs are one flow with three destinations: "Merge to one PDF" (the whole tray, in order), "Extract selected pages" (a new PDF of the selection only), and "Save as separate PDFs" (one file per source-contiguous run, downloaded as a `.zip` built with `core/zipWriter.ts` using `store: true`) — each landing in the library as a SPEC-107 media document or downloading, and none of them touching a source record
- [ ] Nothing about a merge is destructive: sources are opened read-only, `assemblePdf` is given deep-frozen inputs in the unit tests, "Delete page" removes a page from the *plan* and never from the stored source, and the tray's plan is persisted and restorable exactly as SPEC-109 specifies
- [ ] Scale and cancellation: `MAX_MERGE_PAGES = 2000` and `MAX_MERGE_SOURCE_BYTES = 500 * 1024 * 1024` are refused up front with a message naming the limit; copying yields to the main thread every `MERGE_YIELD_PAGES = 8` pages, `assembly-progress` reports `Page {done} of {total}` with `role="status"`, and `assembly-cancel` stops at the next page boundary leaving no file, no record and no mutation
- [ ] Correctness is proved by re-reading, not by eyeballing: every produced file is re-parsed with `parsePdf` before it is stored or downloaded, and a build whose own output does not re-parse — or whose page count and per-page geometry do not match the plan — fails loudly and stores nothing; `assemblePdf` run twice on the same plan is `equalBytes`-identical
- [ ] Tests: `src/lib/core/pdfAssemble.spec.ts` — single-source identity (page count, geometry, `equalBytes` content streams), two-source merge with object renumbering (no collisions, every reference resolvable), inherited-attribute materialisation, rotation arithmetic, link remapping and inert-link removal, outline re-parenting, the drop-record table with one case per `PdfDropKind`, encrypted and unreadable refusals, every bound, and the twice-identical property — over fixtures built by `scripts/build-media-fixtures.ts` plus the checked-in `spec107-xrefstream.pdf` and a new `spec111-forms.pdf`/`spec111-outlines.pdf`; e2e (`e2e/SPEC-111.spec.ts`): open the library with `spec107-pages.pdf` (3 pages) and a second 2-page fixture → select both → "Combine…" → format PDF → the tray shows 5 page cards → rotate page 2 right, delete page 4, drag page 5 to the front → Save to library → the new row's details panel reads 4 pages, page 1 is the former page 5's size, and page 3 (the former page 2) reads `rotate 90` → download and re-open it → the same 4 pages → the "What changed" disclosure listed the dropped structure tree

## Out of scope
Rendering any page (SPEC-107's Out of scope explains why, and it is why page
cards carry no thumbnail). Editing page content: text, images, redaction of
existing PDF content, deleting an object from a page, or changing a page's
size or crop box — a merged page keeps the boxes it arrived with. Merging
`/AcroForm` field trees, flattening form fields, filling forms, and digital
signatures (a merge invalidates any signature on a source; the drop report
says so and the signature bytes are not carried). Merging tagged-PDF
structure trees, so the output is not a tagged PDF. Page labels (`/PageLabels`
is dropped and reported; output pages are numbered 1..n). Compressing or
linearising the output, object streams and cross-reference streams on the
**write** side (the output uses a classic table, which every reader accepts).
Encryption or password-protecting the output. OCR (SPEC-110's exclusion
carries). Comparing two PDFs. Watermarking, stamping, headers and footers on
existing pages.

## Reference

### Dependencies restated
- `specs/107-…` — `src/lib/core/pdfDocument.ts`: `parsePdf(bytes):
  Promise<PdfDocument>`, `PdfValue` (`null`/`bool`/`num`/`str`/`name`/`array`/
  `dict`/`stream`/`ref`), `PdfDocument { version, bytes, objects: Map<number,
  PdfValue>, trailer, pages: PdfPageRef[], info?, linearised, reconstructed }`,
  `PdfPageRef { ref, widthPt, heightPt, rotate, mediaBoxPt, cropBoxPt?,
  inherited }`, `resolve(doc, value)` (depth-capped), `streamData(doc,
  stream)` (FlateDecode + PNG predictors, ASCIIHex/ASCII85; DCT/JPX/CCITT/
  JBIG2/RunLength kept **raw and undecoded**, which is precisely what makes a
  verbatim copy possible), `PdfParseError`/`PdfEncryptedError`, and the
  bounds `MAX_PDF_OBJECTS`/`MAX_XREF_CHAIN`/`MAX_RESOLVE_DEPTH`. Also
  `MediaModel`/`PdfSummary` and the media payload convention.
- `specs/109-…` — the `/assemble` workspace, the tray and its reordering
  vocabulary, the persisted plan, the progress/cancel pattern, and the rule
  that sources are never mutated.
- `specs/110-…` — `buildImagePdf(pages, sources, options)` and the passthrough
  table, reused verbatim for an image item inside a PDF plan.
- `src/lib/core/zipWriter.ts` — `zip(entries)` with `store: true` for the
  already-compressed multi-file output.
- SPEC-008's retention principle, restated: content this layer does not model
  is carried, not discarded, and what genuinely cannot be carried is
  **reported**. `PdfDropRecord` is SPEC-008's `DropRecord` shape over its own
  closed union, exactly as SPEC-106's `DeckDropRecord` is — importing the
  text side's `DropKind` would pull `documentModel.ts` across a product
  boundary.

### The plan and the writer
```ts
// src/lib/core/pdfDocument.ts (writer half)
export interface PdfPageSelection {
  /** A parsed source, by the key the plan uses for reporting. */
  source: string;
  doc: PdfDocument;
  /** 0-based index into doc.pages. */
  pageIndex: number;
  /** Added to the page's own /Rotate, normalised into 0|90|180|270. */
  rotateBy?: 0 | 90 | 180 | 270;
}

export interface PdfAssemblyPlan {
  pages: PdfPageSelection[];        // output order
  title: string;
  author?: string;
  subject?: string;
  /** One top-level outline entry per source, in first-appearance order. */
  outline: boolean;
  modified?: Date;
}

export interface PdfAssemblyResult {
  bytes: Uint8Array;
  dropped: PdfDropRecord[];
}

export function assemblePdf(plan: PdfAssemblyPlan): Promise<PdfAssemblyResult>;
```

The copy, in order:

1. **Reachability.** From each selected page's dictionary, walk every `/ref`
   breadth-first, skipping `/Parent` (rebuilt) and skipping the catalog. The
   visited set is per source, so an object shared by two selected pages of the
   same file is copied once; the same object in two *different* files is
   copied twice — content-identical dedupe across sources is explicitly not
   attempted, because comparing decoded streams would defeat the verbatim
   copy this spec is built on.
2. **Renumbering.** Each copied object gets the next output number; a
   `{t:"ref"}` inside a copied value is rewritten through the per-source map.
3. **Inheritance materialisation.** `/Resources`, `/MediaBox`, `/CropBox` and
   `/Rotate` are written onto the copied page dictionary from
   `PdfPageRef.inherited`, so the page no longer depends on a parent chain
   that is not being copied. `rotateBy` is added here.
4. **Structure.** A new `/Pages` node with `/Kids` in plan order and `/Count`,
   a new `/Root` catalog, `/Info` from the plan (never from a source, so two
   sources cannot fight over the title), and the merged `/OCProperties` and
   `/Dests` where present.
5. **Emission.** Objects in output-number order, then a classic xref table and
   trailer. No object streams, no compression pass, no linearisation: the
   output is larger than a producer-optimised file and every reader opens it.

### The merge contract
| Construct | Result |
| --- | --- |
| page content streams, resources, fonts, XObjects | copied verbatim |
| `/Rotate`, `/MediaBox`, `/CropBox`, `/BleedBox`, `/TrimBox`, `/ArtBox` | copied, materialised onto the page |
| `/Annots` — `/Link` with an internal destination | copied, destination remapped, or removed + reported |
| `/Annots` — `/Link` with `/URI` | copied unchanged |
| `/Annots` — `/Widget` | copied with `/AP`; no longer interactive + reported |
| `/Annots` — markup (text, highlight, stamp, ink) | copied unchanged |
| outlines | one top-level entry per source, children remapped; entries to dropped pages removed + reported |
| `/OCProperties` | `/OCGs` concatenated, `/D` `ON`/`OFF` unioned |
| `/AcroForm`, `/StructTreeRoot`, `/Names/JavaScript`, `/OpenAction`, `/EmbeddedFiles`, `/PageLabels` | dropped + reported |
| encryption | refused |

### Why page cards have no preview
Rasterising a page needs the renderer SPEC-107 explains this product does not
have and cannot add. Two alternatives were considered and rejected: one
scaled `<iframe src="blob:…#page=N">` per card (an entire PDF viewer instance
per page — a 200-page plan is unusable, and the fragment parameters are
Chromium-only), and shipping a renderer (a dependency `package.json` forbids,
and a multi-month build besides). What ships instead is a card that states
what it knows — number, source, size, rotation — and a **result** preview in
the browser's own viewer, which is where a user actually checks a merge. The
spec says this out loud so nobody re-litigates it in review.

### Tests
- `src/lib/core/pdfAssemble.spec.ts` as in the acceptance criteria. The
  single-source identity test is the anchor: parse → assemble with every page
  selected → re-parse → same page count, same geometry, `equalBytes` content
  streams. Any regression in the copier breaks it.
- `src/lib/core/pdfDocument.spec.ts` gains reference-rewriting and
  reachability cases beside its SPEC-107 reader cases.
- `src/views/assemble/usePdfMerge.spec.ts` — page expansion from a source,
  rotation arithmetic through the UI, extract and split destinations, the
  refusals, cancellation, the re-parse self-check, and frozen-source
  assertions.
- `e2e/SPEC-111.spec.ts` as in the acceptance criteria.

### Specs this changes
- `specs/109-…`/`specs/110-…` — the workspace accepts PDF sources and gains
  page-level operations and two extra destinations; existing behaviours and
  tests are unchanged.
- `specs/107-…` — `core/pdfDocument.ts` grows its writer half; the reader's
  contract is untouched.
- `specs/000-roadmap.md` gains this spec's phase row.
