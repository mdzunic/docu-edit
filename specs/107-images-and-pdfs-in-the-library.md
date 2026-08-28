---
id: SPEC-107
title: Images and PDFs in the library
status: draft
area: io
depends_on: [SPEC-003, SPEC-009, SPEC-024, SPEC-058, SPEC-060, SPEC-063, SPEC-067]
---
## Why
The library holds documents and decks. It does not hold the two file types
that arrive alongside them in every real folder: the scan, the screenshot,
the logo, the signed contract someone sent back as a PDF. A user who wants
to put four scans into one PDF, or crop a screenshot before dropping it into
a report, leaves this product to do it — and every one of those errands ends
somewhere else.

The pieces are already here. `src/lib/core/media.ts` types and validates an
image payload; `core/imageDimensions.ts` reads its natural size from the
bytes without a decoder; the store already splits image payloads into their
own IndexedDB record so a 40 MB library does not block the landing page
(SPEC-047 AC-15); `documentKinds.ts` was built as a registry whose own comment
says that adding a future kind is an edit there plus a new product layer; and SPEC-034 already
writes PDF 1.7 by hand, which means this repo already knows the file format
from the inside.

What is missing is a **third kind**. This file adds it: `"media"`, a product
whose model is a byte payload plus what can be learned about it without
rendering it, living at `/media/:id`. It is the normative source for the
`MediaModel`, for `MediaFormat` and content-type detection, and for
`src/lib/core/pdfDocument.ts` — the structural PDF reader every later spec in
this phase merges, reorders and searches with.

**The one thing this spec does not do is rasterise a PDF.** There is no PDF
renderer in this repo and no dependency may be added to build one; a
content-stream interpreter with font programs is a product, not a feature.
The browser ships one, so the viewer uses it, and everywhere a rendered PDF
page would be needed the spec says so and offers something honest instead.

## Acceptance criteria
- [ ] `documentKinds.ts` gains `"media"` per the Reference: `DocumentKind = "text" | "presentation" | "media"`, `MEDIA_FORMATS = ["pdf", "png", "jpeg", "gif", "svg", "webp"]`, `kindOfFormat` returns `"media"` for each, `editorPathFor("media", id)` is `/media/${id}`, and `OPENABLE_EXTENSIONS` gains `pdf`, `png`, `jpg`, `jpeg`, `gif`, `svg`, `webp` — every kind-dispatched call site (router guards, `HomeView`, `OpenFromUrlView`, `useLibrary.downloadDocuments`, the pickers' `accept`) keeps deriving from this module and gains no local knowledge of what a PDF is
- [ ] `src/lib/media/` is the third product layer, importing only its own siblings and `src/lib/core` — never `documentModel.ts`, the commands, `docx/`, `odf/`, `presentation/`, the store or Vue — enforced by a new `src/lib/media/boundary.spec.ts` modelled line-for-line on `presentation/boundary.spec.ts`, and by a new block in `core/boundary.spec.ts` asserting the reverse direction (no core or text or presentation module imports `media/`)
- [ ] `src/lib/media/model.ts` exports `MediaModel` per the Reference — `kind: "media"`, `id`, `title`, `format`, `contentType`, `sourceMediaId`, `media: MediaPart[]`, `byteLength`, optional `image`/`pdf` summaries and `properties` — and it is plain JSON: `JSON.parse(JSON.stringify(model))` `toEqual`s it, the payload living in `media` as the SPEC-024 `MediaPart` the store already splits out, so listing a library of 200 images structured-clones no pixels
- [ ] `src/lib/core/pdfDocument.ts` exports `parsePdf(bytes): Promise<PdfDocument>` per "The PDF structural layer": header version, both cross-reference flavours (classic tables with `/Prev` chains and hybrid `/XRefStm`, and PDF 1.5 cross-reference streams), object streams inflated through `DecompressionStream("deflate")`, the page tree walked with attribute inheritance, and per page a `widthPt`/`heightPt`/`rotate` derived from the effective `/MediaBox`, `/CropBox` and `/Rotate` — no content stream is interpreted and no glyph is decoded
- [ ] Broken and hostile input is refused, never guessed: an encrypted file throws `PdfEncryptedError` ("This PDF is password-protected"), a file whose xref is unusable is retried once by scanning the whole buffer for `N G obj` headers and rebuilding the table (recording `reconstructed: true`), and anything still unreadable throws `PdfParseError` with the byte offset; parsing is bounded — `MAX_PDF_OBJECTS = 500_000`, `MAX_XREF_CHAIN = 64`, a resolution depth cap and a visited set, so a cyclic `/Prev`, a self-referential page tree or a decompression bomb terminates with an error rather than a hang
- [ ] `openAnyDocumentFile` sniffs media **before** it attempts an unzip, from the bytes and never the extension: `%PDF-` within the first 1024 bytes → pdf; the PNG, GIF87a/89a, JPEG SOI and `RIFF….WEBP` signatures → their formats; and an XML document whose root element is `<svg>` (after an optional BOM, XML declaration, doctype and comments, within the first 4 KiB) → svg — `.svg` being the one format whose sniff is structural rather than a magic number, stated here the way SPEC-050 states Markdown's extension exception
- [ ] Every open path reaches media unchanged in shape: the library picker, drag-and-drop and `/open?url=` all funnel through `openAnyDocumentFile`, the drop name-check accepts the new extensions, and the drop rejection message becomes "We can only open .docx, .odt, .md, .pptx, .odp, .pdf and image files" — a **behavioural string change** whose existing assertions in `e2e/SPEC-009.spec.ts`, `e2e/SPEC-050.spec.ts` and `AppShell.spec.ts` are updated with it, and `openDocumentFile`'s text-only contract still rejects a PDF with a specific message rather than misparsing it
- [ ] `MAX_FILE_SIZE_BYTES` stays 50 MB for every kind (one limit, one message), and a media file that passes it is stored verbatim: the bytes that came in are the bytes that go back out — downloading an untouched media document is `equalBytes`-identical to the source file, asserted in `src/lib/media/roundTrip.spec.ts` for all six formats
- [ ] The library treats media as first-class rows and cards: `formatBadgeLabel` covers `PDF`/`PNG`/`JPEG`/`GIF`/`SVG`/`WEBP`, `FILTERABLE_FORMATS` gains them, the format chips gain a `kind` grouping so "Images" and "PDFs" are one click each, the size column reports `formatFileSize(byteLength)`, and search, sort, starring, multi-select, trash, restore and bulk download all work on media entries with no new code path — `LibraryEntry.kind` already carries the routing
- [ ] Thumbnails are honest per kind: an image card renders a downscaled raster of the image itself, produced once per record through `IntersectionObserver` + the idle scheduler and cached under the SPEC-067 rules (id + `lastModified`, evicted before anything else under quota pressure), capped at `MEDIA_THUMBNAIL_MAX_EDGE_PX = 512`; a PDF card renders a paper-shaped placeholder at the **real** aspect ratio of its first page with the page count and page size on it, and the spec states plainly that no page is rasterised — a card never fakes a preview it cannot produce
- [ ] `/media/:id` (`MediaView.vue`) is the viewer, sharing SPEC-058's app bar and SPEC-060's Close: an image opens fit-to-window with the SPEC-063 zoom vocabulary (`fit`, `100%`, a zoom control, Cmd/Ctrl+`+`/`-`/`0`, drag to pan, double-click to toggle fit/100%), an `<img>` with the file name as its `alt`, and a details panel with dimensions, format, byte size and stored date; the route is kind-guarded exactly as `/doc/:id` and `/slides/:id` are, in both directions
- [ ] A PDF at `/media/:id` renders in the **browser's own** viewer — an `<object type="application/pdf">` over a `blob:` object URL revoked on unmount, with a `<embed>` and then a "Your browser cannot display PDFs here" panel offering Download and Open in a new tab as the nested fallbacks — beside the same details panel, which for a PDF reports page count, per-page size (or "mixed" with the first page's size), PDF version, `/Info` title/author/producer/dates and whether the file is linearised or was reconstructed; the panel is built from `parsePdf`, so it is there even when the embed is not
- [ ] Media never enters a path built for documents: `getMedia(id)` narrows like `getPresentation`, `getOrCreateDocument` and `saveDraft` throw on a media id, `useAutosave.beginSession` bails, the version-history and comment surfaces are absent, and the editor-only File menu entries do not render — the failure mode for a wrong kind is a loud throw or a redirect, never a silent overwrite
- [ ] `core/imageDimensions.ts` gains WebP per the Reference (the `RIFF`/`WEBP` container with its `VP8 `, `VP8L` and `VP8X` chunk variants) and keeps returning `null` for anything malformed; a `LIBRARY_IMAGE_MIME_TYPES` constant in `core/media.ts` adds `image/webp` to the four SPEC-024 accepts **without touching** `ACCEPTED_IMAGE_MIME_TYPES`, which stays the insert-into-a-document list SPEC-024 owns
- [ ] Fixtures and tests: `scripts/build-media-fixtures.ts` writes `e2e/fixtures/spec107-photo.jpg`, `-shot.png`, `-mark.svg`, `-anim.gif`, `-shot.webp` and `-pages.pdf` (a 3-page PDF with mixed page sizes and a rotated page, built by the app's own `buildPdf` so it is reproducible), plus `-xrefstream.pdf` and `-encrypted.pdf` checked in beside them with a note recording how each was produced; `core/pdfDocument.spec.ts` covers both xref flavours, object streams, inheritance, rotation, the reconstruction path, encryption refusal and every bound
- [ ] e2e (`e2e/SPEC-107.spec.ts`): drop `spec107-pages.pdf` on the library → a row with a `PDF` badge and a 3-page card → open it → the details panel reads 3 pages and the embed element is present → Close → drop `spec107-shot.png` → its card shows a real thumbnail → open it → zoom to 100% and back to fit → select both rows → bulk download → the PDF's bytes are `equalBytes` with the fixture → trash and restore both → filter by "Images" shows one row, "PDFs" the other

## Out of scope
Rendering a PDF page to pixels — the whole class: PDF page thumbnails, a
canvas PDF viewer, page previews in any later spec of this phase, and
printing a PDF through anything but the browser's own viewer. Editing a PDF's
text or objects. PDF forms, signatures, annotations as an editable surface,
encrypted PDFs (refused with a message; decrypting them is not a feature this
product wants). Extracting a PDF's text (SPEC-114 owns it, and owns its own
honesty about coverage). Converting a PDF into an editable document. Animated
GIF playback controls (an animated GIF renders as the browser plays it, and
is stored and returned verbatim). Camera or scanner capture. HEIC, TIFF, BMP,
AVIF and RAW — each needs a decoder the platform does not universally provide;
the six formats here are the ones every target browser decodes and this repo
can measure without one. EXIF, ICC profiles and colour management: bytes are
preserved, nothing is interpreted, and an EXIF-rotated JPEG displays as the
browser displays it. Editing an image (SPEC-112), assembling anything
(SPEC-108–111), and any cloud or server storage.

## Reference

### Dependencies restated
- `specs/003-local-document-library.md` / `specs/067-…` — the store is
  `src/lib/documentStore.ts`: an in-memory map written through to IndexedDB,
  `listDocuments()` → `LibraryEntry[]` (`id`, `title`, `format`, `kind`,
  `lastModified`, `starred`, `trashedAt?`), `trashDocument`/`restoreDocument`/
  `deleteDocument`/`setStarred`/`libraryStorageUsage`, thumbnails cached by id
  + `lastModified` and evicted first under quota pressure, and the whole
  search/sort/filter/multi-select surface in `src/lib/libraryView.ts` +
  `src/composables/useLibrary.ts` + `src/views/library/*`.
- `specs/024-images-insert-resize-and-wrap.md` — `MediaPart { mediaId;
  contentType; data /* base64 */ }` in `src/lib/core/media.ts`, one payload
  per id, `mediaDataUri(part)`, `MAX_IMAGE_SIZE_BYTES`, and
  `ACCEPTED_IMAGE_MIME_TYPES` = png/jpeg/gif/svg+xml — **the insert-into-a-
  document list, which this spec does not change**. `core/imageDimensions.ts`
  exports `decodeImageDimensions(bytes, contentType): ImageDimensions | null`,
  reading PNG/GIF/JPEG/SVG sizes from the bytes with no image decoder so it
  works identically under Vitest+jsdom.
- `specs/009-drag-drop-and-open-from-url.md` — `openAnyDocumentFile` is the
  one funnel (picker, drop, `/open?url=`); `MAX_FILE_SIZE_BYTES = 50 * 1024 *
  1024` and `FILE_TOO_LARGE_MESSAGE` are shared by all of them;
  `hasSupportedExtension` gates the drop before any byte is read, because a
  drop has no `accept` filter.
- `specs/034-print-and-export-pdf.md` — `buildPdf(input): Uint8Array` in
  `src/lib/pdf/buildPdf.ts` writes PDF 1.7 by hand (objects, xref, trailer),
  which is why this spec's *reader* needs no dependency either. **If SPEC-104
  has landed, that module is `src/lib/core/pdfBuild.ts` and its siblings are
  the flat `core/pdf*.ts` names; this spec's new file is `core/pdfDocument.ts`
  either way, and imports nothing from the writer.**
- `docs/presentations-architecture.md` — the layering this extends: core is
  product-neutral and flat, each product layer imports core and its own
  siblings only, and the two (now three) products meet in the app layer
  (`documentKinds.ts`, `openDocument.ts`, `documentStore.ts`, the router, the
  views). This file adds a third box, not a fourth rule.

### The media model
```ts
// src/lib/media/model.ts
export type MediaFormat = "pdf" | "png" | "jpeg" | "gif" | "svg" | "webp";

export interface MediaModel {
  kind: "media";                    // the store's discriminant, as PresentationModel does
  id: string;
  title: string;                    // file name without the extension (titleFromFileName)
  format: MediaFormat;
  contentType: string;              // "application/pdf", "image/png", …
  /** The MediaPart in `media` holding the source bytes, verbatim. */
  sourceMediaId: string;
  media: MediaPart[];               // exactly one part here; SPEC-112 adds a second
  byteLength: number;               // the source length, so a row renders without the payload
  image?: { widthPx: number; heightPx: number };   // absent for pdf, and for an unreadable image
  pdf?: PdfSummary;                 // pdf only
  properties?: DocumentProperties;  // from /Info, mapped onto core/documentProperties.ts
}

export interface PdfSummary {
  version: string;                  // "1.7"
  pageCount: number;
  pages: Array<{ widthPt: number; heightPt: number; rotate: 0 | 90 | 180 | 270 }>;
  linearised: boolean;
  reconstructed: boolean;           // the xref had to be rebuilt by scanning
}
```
`media` is the array the store's SPEC-047 split already writes to its own
`MEDIA_KEY_PREFIX` record, on the same reference-identity latch — so a media
model never rewrites its payload on a rename or a star, and the landing page
never loads one. `StoredModel` becomes `DocumentModel | PresentationModel |
MediaModel`; `isPresentationModel` is unchanged (it tests `kind ===
"presentation"`), and `isMediaModel` joins it.

### The PDF structural layer
`src/lib/core/pdfDocument.ts` is a **structural** reader: it turns bytes into
an object graph and a page list, and stops there. Async, because inflating an
object or xref stream goes through the platform's `DecompressionStream`, the
same trade `core/zip.ts` and `pdf/imageEncode.ts` already make.

```ts
export type PdfValue =
  | { t: "null" } | { t: "bool"; v: boolean } | { t: "num"; v: number }
  | { t: "str"; v: Uint8Array } | { t: "name"; v: string }
  | { t: "array"; v: PdfValue[] } | { t: "dict"; v: Map<string, PdfValue> }
  | { t: "stream"; dict: Map<string, PdfValue>; raw: Uint8Array }
  | { t: "ref"; num: number; gen: number };

export interface PdfPageRef {
  ref: { num: number; gen: number };
  widthPt: number; heightPt: number;      // the effective CropBox, normalised
  rotate: 0 | 90 | 180 | 270;
  mediaBoxPt: [number, number, number, number];
  cropBoxPt?: [number, number, number, number];
  inherited: Map<string, PdfValue>;       // Resources/MediaBox/CropBox/Rotate as resolved
}

export interface PdfDocument {
  version: string;
  bytes: Uint8Array;                      // the source, retained for byte-preserving copy
  objects: Map<number, PdfValue>;         // by object number (generation 0 assumed, see below)
  trailer: Map<string, PdfValue>;
  pages: PdfPageRef[];
  info?: Map<string, PdfValue>;
  linearised: boolean;
  reconstructed: boolean;
}

export function parsePdf(bytes: Uint8Array): Promise<PdfDocument>;
export function resolve(doc: PdfDocument, value: PdfValue): PdfValue;   // follows refs, depth-capped
export function streamData(doc: PdfDocument, stream): Promise<Uint8Array>; // FlateDecode + predictors
export class PdfParseError extends Error { readonly offset: number }
export class PdfEncryptedError extends PdfParseError {}
```

Decisions worth pinning, because guessing them later would be a rewrite:

- **Generations.** Real-world non-incremental PDFs use generation 0 for every
  live object. The map is keyed by object number; a non-zero generation in an
  xref entry is honoured when reading (the newest xref section wins) but is
  never a second map key. A file that genuinely needs generation-keyed
  lookup is out of this spec's ceiling and is reported, not half-read.
- **Stream filters.** `/FlateDecode` (with `/Predictor` 2 and 10–15 — PNG
  predictors are what xref streams actually use) is decoded; `/DCTDecode`,
  `/JPXDecode`, `/CCITTFaxDecode`, `/JBIG2Decode` and `/RunLengthDecode`
  streams are **kept raw and never decoded** — nothing here needs their
  samples, and a later spec copies them byte-for-byte. `/ASCIIHexDecode` and
  `/ASCII85Decode` are decoded (they are cheap and appear in hand-made files).
  An unknown filter on a stream this layer must read is a `PdfParseError`.
- **`/Length` as an indirect reference** is resolved after the xref is built;
  when it is wrong or unresolvable the stream is bounded by scanning to the
  next `endstream`, which is what every real reader does.
- **Bounds are part of the contract**, not defensive extras: `MAX_PDF_OBJECTS
  = 500_000`, `MAX_XREF_CHAIN = 64`, `MAX_RESOLVE_DEPTH = 64`, a visited-set
  on the page-tree walk, and `MAX_INFLATED_STREAM_BYTES = 256 * 1024 * 1024`.
  Each has a test that trips it.
- **Reconstruction.** When the trailer or xref cannot be read, scan for
  `\d+ \d+ obj` headers, take the last definition of each object number, find
  the `/Type /Catalog`, and set `reconstructed`. This is not a nicety: a
  meaningful share of PDFs in the wild have a broken xref, and the
  alternative is telling the user their file is corrupt when every other
  reader opens it.

### Detection
`core/packageDetect.ts` gains `detectMedia(bytes): MediaFormat | null`, run
by `openAnyDocumentFile` **first** — media files are not ZIPs, so sniffing
before the unzip attempt is both cheaper and unambiguous:

| Format | Signal |
| --- | --- |
| pdf | `%PDF-` within the first 1024 bytes (the offset is recorded; a prefixed file still parses) |
| png | `89 50 4E 47 0D 0A 1A 0A` |
| jpeg | `FF D8 FF` |
| gif | `GIF87a` / `GIF89a` |
| webp | `RIFF` at 0 and `WEBP` at 8 |
| svg | XML whose first element is `svg` (BOM, `<?xml…?>`, `<!DOCTYPE…>` and comments skipped), within 4 KiB |

Order matters only for svg, which is checked last because it is the only
structural one. `.md` keeps its extension exception; nothing else gains one.

### The viewer
`src/views/MediaView.vue` is a wiring hub under the same rule as the other
two views, with `MediaView.size.spec.ts` opening its ratchet ledger at the
line count it lands on. Its parts:
`src/views/media/useMediaDocument.ts` (load, kind guard, blob URL lifecycle —
one `URL.createObjectURL` per open, revoked in `onScopeDispose`),
`useImageViewport.ts` (fit/zoom/pan, reusing `src/lib/zoom.ts`'s clamp and
step vocabulary rather than a second one), and the components
`MediaImageCanvas.vue` (`media-image`), `MediaPdfEmbed.vue` (`media-pdf`,
`media-pdf-fallback`) and `MediaDetailsPanel.vue` (`media-details`, with
`detail-pages`, `detail-size`, `detail-dimensions`, `detail-format`,
`detail-producer`). Testids for the app bar reuse SPEC-060's.

**Why an embed and not a canvas.** Rasterising a PDF page needs a content
stream interpreter, an operator model, embedded font programs (Type1, CFF,
TrueType, CID), shading, transparency groups and colour spaces. That is a
multi-month build and a permanent maintenance surface, and `package.json` is
frozen. Every target browser ships a viewer; using it is a hundred lines and
is correct on day one. What it costs is stated rather than hidden: no page
thumbnails, no in-page selection API, no control over the toolbar, and a
fallback panel where the platform has no viewer. The alternative considered
and rejected was one scaled `<iframe src="blob:…#page=N&view=Fit">` per
thumbnail — it is cross-browser unreliable and costs an entire PDF instance
per card, which a 200-item grid cannot pay.

### Tests
- `src/lib/core/pdfDocument.spec.ts` — the lexer over hand-written byte
  fixtures, classic xref + `/Prev` chains, hybrid `/XRefStm`, xref streams
  with PNG predictors, object streams, page-tree inheritance and rotation,
  the reconstruction path, `PdfEncryptedError`, and one test per bound.
- `src/lib/media/model.spec.ts` (JSON round trip, summaries),
  `roundTrip.spec.ts` (byte-identical download for all six formats),
  `detect.spec.ts` (every signature, plus the SVG false-positive cases: an
  XHTML file, an XML file whose *comment* mentions svg, an empty file).
- `src/lib/core/imageDimensions.spec.ts` gains the three WebP chunk layouts.
- `src/lib/media/boundary.spec.ts` and the new reverse block in
  `core/boundary.spec.ts`.
- `src/views/MediaView.spec.ts`, `src/views/media/useImageViewport.spec.ts`,
  `HomeView.media.spec.ts` (badges, chips, thumbnails, bulk actions over
  mixed kinds), `e2e/SPEC-107.spec.ts`.

### Specs this changes
- `specs/003-…`/`specs/067-…` — the library holds a third kind; every listed
  behaviour applies to it unchanged. No store contract is removed.
- `specs/009-…`/`specs/050-…` — the accepted-extension list and the drop
  rejection string grow; their e2e assertions are updated with them.
- `specs/024-…` — `core/imageDimensions.ts` gains WebP and `core/media.ts`
  gains a library-side MIME list. The insert-image accept list is untouched.
- `specs/000-roadmap.md` gains this spec's phase row.
