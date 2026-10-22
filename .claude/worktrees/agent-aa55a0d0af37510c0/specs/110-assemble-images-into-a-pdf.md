---
id: SPEC-110
title: Assemble images into a PDF
status: draft
area: export
complexity: normal
story_points: 5
depends_on: [SPEC-034, SPEC-107, SPEC-108, SPEC-109]
---
## Why
The most common thing anyone does with a pile of images is turn it into one
PDF. Six photographed pages of a signed contract. Twelve scans of receipts
for an expense claim. A passport and a utility bill, because the form takes
one attachment. It is the single errand that sends people to an advertising-
funded website with their identity documents.

SPEC-109 built the workspace and can already put many images on one sheet.
This spec adds the other axis — **many sheets** — and the output that makes
them a document: a real PDF, written by the writer this repo already owns,
with the source photographs embedded rather than re-photographed.

The quality rule is the point of the spec. A phone photo is a baseline JPEG;
`encodeImage` already passes one straight through as `/DCTDecode`. If the
assembly path re-encodes it — by drawing it to a canvas to crop, rotate or
scale — a 4 MB scan becomes a 12 MB blur. **This file is the normative source
for when an image is passed through and when it is re-encoded**, and for
multi-page assembly, PDF metadata and page numbering.

## Acceptance criteria
- [ ] The SPEC-109 workspace gains `pdf` in `assembly-format`, which switches the inspector from one sheet to **pages**: a `assembly-page-mode` select of `one-per-page` (default), `n-up` (with `assembly-n-up` of 2, 4, 6, 9 or 16 reusing SPEC-109's grid layout inside one page) and `single-page` (every item on one sheet, the SPEC-109 behaviour, producing a one-page PDF), plus `assembly-page-numbers` (off by default) and a `assembly-page-count` readout
- [ ] The preview becomes paged: a page strip (`assembly-pages`, items `assembly-page`) beneath the preview with the current page highlighted, `PageUp`/`PageDown` and the strip both navigating, and the sheet controls applying to **every** page — one sheet setup per assembly, not per page, with per-page overrides explicitly out of scope
- [ ] `src/lib/media/imagesToPdf.ts` exports `planPdfPages(items, plan, sheet): PdfAssemblyPage[]` (pure, reusing `planAssembly` per page) and `buildImagePdf(pages, sources, options): Promise<Uint8Array>`, which composes one `PdfPageInput` per page — a background `PdfRectItem`, one `PdfImageItem` per placed item at its rectangle, and the page number as a `PdfTextItem` when enabled — and calls `buildPdf` once; it adds no PDF serialisation of its own, and a construct the writer cannot emit is fixed in the writer
- [ ] Passthrough is the default and its conditions are exact: a source is embedded **verbatim** as `/DCTDecode` when it is a baseline JPEG (`parseJpegHeader().baseline`, 1 or 3 components), its placement crops nothing (`fit` is `contain` or `stretch`, or `cover`/`actual` resolved to the full source rectangle), it carries no SPEC-112 edit stack, and its `opacity` is 1 with `blend` `normal`; **any** other combination re-encodes through the canvas path, and the reason is shown per item in the tray (`item-encoding`, reading "original JPEG" or "re-encoded: cropped")
- [ ] Re-encoding respects the resolution the user chose: `encodeImage` is called with SPEC-108's `maxDpi` set from the sheet's `dpi`, so a 600 dpi assembly is not silently downsampled to 300; an image whose natural resolution is below the sheet's is embedded at its own resolution and never upsampled, and `assembly-estimate` reports the effective dpi of the lowest-resolution placement so a user learns *before* exporting that page 4 will print soft
- [ ] The same source placed on several pages is embedded once — `PdfImageItem.src` keys `PdfBuildInput.images`, which is the writer's existing dedupe — and a 40-page assembly of the same letterhead carries one copy of it
- [ ] PDF metadata is filled rather than left empty: `/Title` from `assembly-title`, `/Author` and `/Subject` from an `assembly-metadata` disclosure (empty by default, never auto-filled from the user's identity), `/Creator` and `/Producer` naming this application, and one level-1 outline entry per **item** (its source title) so a 30-scan PDF has a navigable bookmark list — reusing `PdfHeading` exactly as SPEC-034 uses it
- [ ] A page number, when enabled, is a 9 pt centred `PdfTextItem` 1 cm above the sheet's bottom edge in `#555555`, numbered from 1 over the produced pages, using the bundled face SPEC-034 already embeds — no new font is added, and it is drawn inside the margin, never over a placed image
- [ ] Limits are stated and enforced before work starts: `MAX_PDF_ASSEMBLY_PAGES = 500` and `MAX_PDF_ASSEMBLY_BYTES = 500 * 1024 * 1024` (the estimated output), each refused with a message naming the limit and what to remove; the estimate is computed from the passthrough decision per item, so it is a real number and not a guess
- [ ] Building never blocks the page and can be abandoned: pages are composed one at a time with `yieldToMainThread()` between them, `assembly-progress` reports `Page {done} of {total}` with `role="status"`, `assembly-cancel` stops at the next page boundary leaving no download, no library record and no mutation, and every intermediate canvas is released after encoding
- [ ] The result lands as a SPEC-107 media document (format `pdf`, title from `assembly-title` under SPEC-003's duplicate numbering, its `PdfSummary` produced by running `parsePdf` over the bytes just written — which also serves as a self-check: a build whose own output will not re-parse fails loudly and stores nothing) or downloads through `core/triggerDownload.ts`; no source item is read-locked, moved, renamed or modified
- [ ] Determinism: `buildImagePdf` called twice on the same plan and sources produces `equalBytes`-identical bytes (`src/lib/core/equality.ts`) when `options.modified` is fixed, and a passthrough item's embedded stream is `equalBytes` with the source file's own JPEG data — asserted directly, because passthrough silently breaking into re-encoding is the failure this spec exists to prevent
- [ ] Tests: `src/lib/media/imagesToPdf.spec.ts` — page planning for the three page modes over 1, 5 and 13 items, the passthrough decision table (every row, both outcomes), image dedupe across pages, the outline tree, page-number geometry, `MAX_*` refusals, and the twice-identical property; e2e (`e2e/SPEC-110.spec.ts`): open the library with three image fixtures → select all → "Combine…" → format PDF, one per page, A4 portrait, page numbers on → the strip shows 3 pages → Save to library → the new `PDF` row's details panel reads 3 pages at 21 × 29.7 cm → download it, and the JPEG fixture's bytes appear verbatim inside the downloaded file → switch to `n-up` 4 → the strip shows 1 page

## Out of scope
OCR and any text layer over a scanned image — a PDF produced here is a picture
of text, its outline entries are the only searchable strings, and SPEC-114
reports it as unextractable rather than pretending otherwise. Per-page sheet
overrides, mixed page sizes within one assembly, and per-page rotation
(rotate the image in SPEC-112 instead). Cropping, deskewing, auto-orientation
and edge detection of a photographed page — image work belongs to SPEC-112.
PDF/A, tagged PDF, embedded colour profiles, encryption and signing.
Compressing an existing PDF. Merging a PDF **into** the assembly (SPEC-111).
Captions or file names printed under each image. Contact-sheet index pages.
Splitting an assembly across several files.

## Reference

### Dependencies restated
- `specs/034-print-and-export-pdf.md` — `buildPdf(input)` over `PdfPageInput
  { widthPt, heightPt, items: PdfTextItem[], links, headings: PdfHeading[],
  images: PdfImageItem[] }` (+ `rects` per SPEC-104/108), coordinates in
  points from the page's **top**-left, `PdfBuildInput { title, pages, fonts,
  images: Map<string, PdfImageData> }` embedding each keyed image once,
  `encodeImage` passing a baseline JPEG through as `/DCTDecode` and drawing
  everything else to a canvas as zlib RGB samples with an `/SMask` when it is
  not opaque, `parseJpegHeader(bytes): JpegHeader | null` with its `baseline`
  and `components` fields, `PDF_MIME_TYPE`, `pdfFileName(title, fallback)`.
  Where SPEC-104 has landed these are the flat `core/pdf*.ts` modules.
- `specs/108-…` — `SheetSetup`/`SheetGeometry`, `resolveSheet`,
  `placeOnSheet`, `SHEET_DPI_CHOICES`, `encodeImage`'s new `maxDpi`, the
  `r3` rounding rule, and `SheetSetupPanel.vue`.
- `specs/109-…` — the `/assemble` workspace, its tray, `AssemblyPlan`,
  `planAssembly(items, plan, sheet): PlacedItem[]` with `sourcePx` per placed
  item, `composeCanvas.ts`, the progress/cancel pattern, the persisted plan
  and the no-source-mutation rule. This spec adds an output format and a page
  axis to that surface; it introduces no second workspace.
- `specs/107-…` — `MediaModel`, `parsePdf(bytes)` and `PdfSummary`, and the
  media payload convention.

### Page planning
```ts
// src/lib/media/imagesToPdf.ts
export type PageMode = "one-per-page" | "n-up" | "single-page";

export interface PdfAssemblyPage {
  /** Indices into the tray, in order, that this page carries. */
  itemIndices: number[];
  placed: PlacedItem[];      // from planAssembly, against this page's content box
}

export interface ImagePdfOptions {
  title: string;
  author?: string;
  subject?: string;
  pageNumbers: boolean;
  modified?: Date;           // fixed by tests; defaults to now
}

export function planPdfPages(items: AssemblyItem[], plan: AssemblyPlan & { pageMode: PageMode; nUp: number }, sheet: SheetGeometry): PdfAssemblyPage[];
export function buildImagePdf(pages: PdfAssemblyPage[], sources: Map<string, MediaPart>, options: ImagePdfOptions): Promise<Uint8Array>;
```
`one-per-page` is `n-up` with `nUp = 1`, and is spelled separately because it
is the default and because its cell **is** the content box (no gap
arithmetic). `single-page` delegates entirely to SPEC-109's plan.

### The passthrough decision
| Condition | Result |
| --- | --- |
| baseline JPEG, 1 or 3 components | eligible |
| progressive / arithmetic / CMYK JPEG | re-encode |
| PNG, GIF, WEBP, SVG | re-encode (SVG rasterised at its placed size) |
| `sourcePx` covers the whole image | eligible |
| `sourcePx` is a sub-rectangle (a `cover`/`actual` crop) | re-encode |
| `opacity < 1` or `blend !== "normal"` | re-encode |
| a SPEC-112 edit stack is present | re-encode from the edited render |

An eligible item's bytes are handed to `PdfBuildInput.images` as
`{ filter: "DCTDecode", data: <the source bytes> }` with the header's own
dimensions and colour space — the file's own scan data, untouched. Every
re-encode goes through `encodeImage` with `maxDpi = sheet.dpi`. The tray
shows which happened per item, because "why did my 3 MB scan become 11 MB"
is the question this table exists to answer before it is asked.

### Tests
- `src/lib/media/imagesToPdf.spec.ts` as in the acceptance criteria, with the
  passthrough table as a data-driven case list so a new source format cannot
  be added without deciding its row.
- `src/views/assemble/usePdfAssembly.spec.ts` — page-mode switching keeps the
  tray order, the estimate's effective-dpi warning, the `MAX_*` refusals,
  cancellation leaving nothing behind, and the parse-back self-check failing
  loudly on a corrupted build (a stubbed `buildPdf` returning junk stores
  nothing and reports).
- `e2e/SPEC-110.spec.ts` as in the acceptance criteria — including the
  verbatim-JPEG assertion, which is what proves passthrough end to end.

### Specs this changes
- `specs/109-…` — the workspace gains a format, a page axis and a page strip;
  its existing behaviours are unchanged and its tests keep passing.
- `specs/034-…` — no behaviour change; this is a second caller of the writer,
  and any gap it finds is fixed in the writer rather than worked around here.
- `specs/000-roadmap.md` gains this spec's phase row.
