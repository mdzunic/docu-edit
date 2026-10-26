---
id: SPEC-144
title: Images from the command line — pages, slides and pictures to PNG, JPEG and WebP, and pictures to PDF
status: draft
area: platform
complexity: normal
story_points: 8
depends_on: [SPEC-143, SPEC-104, SPEC-107, SPEC-108, SPEC-112]
---
## Why
The second thing a shell is asked to do with a document is "give me a picture
of it": a slide for a wiki page, page one of a report for a thumbnail, a
chart sheet for a chat, a scanned PDF as PNGs for an OCR tool, a folder of
phone photos as JPEGs at a sane size, an SVG logo as a PNG. The app can do
some of this — SPEC-104 exports slides as PNG, SPEC-108 turns one image into
a one-page PDF, SPEC-107 renders a PDF's first page for its card — and none
of it from a command.

All of it is already built on the pieces SPEC-143 put behind the command.
SPEC-104 rasterises a slide from the *same* `PdfPageInput` the PDF is built
from (`core/pdfRaster.ts`'s `rasterisePdfPage`), so "a page as an image" is
the page job SPEC-143 produces for every product, painted instead of written
— a document page and a sheet page get images for free, and a PNG and a PDF of
the same page always agree. pdf.js, the one sanctioned rendering dependency,
already renders a PDF page in the browser. Decoding and encoding a picture is
a canvas away. This spec adds image targets to the render tier, PDF-to-image
and image-to-image conversion, and image-to-PDF through SPEC-108's own
`convertImageToPdf`. **This file is the normative source for image output
naming, the pixel-size options and their clamp, JPEG/WebP encoding of page
rasters, multi-page PDF rasterisation, image-to-image conversion and the
`convert` options for image-to-PDF.**

## Acceptance criteria
- [ ] Image targets `png`, `jpeg` (alias `jpg`) and `webp` are browser-tier `CONVERSIONS` rows from text, presentation and workbook (one image per page or slide of the SPEC-143 page job, accepting that product's SPEC-143 options — `--pages`, `--slides`/`--include-hidden`, `--sheet`/`--all-sheets`) and from media (an image to another image format, a PDF to one image per page); `gif` and `svg` are never targets, and asking for one is a usage error (exit 1, from the arguments alone) naming the three that are
- [ ] Decks render slides, never sheets: a presentation's image job is SPEC-104's image plan — layout `"slides"`, one slide per image, each at the slide's own aspect — and `--layout`, `--paper`, `--orientation`, `--note-lines` and `--frame-slides` are usage errors with an image target; text and workbook images are their PDF pages exactly
- [ ] Media inputs need no view: for a file the funnel opens as `media`, the bridge's `open` keeps the opened `MediaModel` in the page and resolves without storing, navigating or waiting for a page producer; `RenderRequest` gains a media variant per "Bridge additions", and every media conversion below runs in the page on that model
- [ ] Page images are painted by `core/pdfRaster.ts` after `registerRasterFaces(job.fonts)` (`src/lib/pdfAssets.ts`) has registered the job's faces, as SPEC-104's image export does: PNG through the existing `rasterisePdfPage(page, widthPx)` (its output unchanged), JPEG and WebP through a new `rasterisePdfPageAs(page, widthPx, { type, quality })` that shares its painting and composites onto opaque white for JPEG (which has no alpha) before encoding; for every page, PNG pixels and the PNG decoded from the deck's own SPEC-104 export agree (the e2e compares SPEC-104's image download with the command's PNG by `equalBytes`)
- [ ] Pixel size per "Pixel size": `--scale <n>` (default 2 — SPEC-104's default), `--dpi <n>` or `--width <px>`/`--height <px>`, mutually exclusive; a result whose longer edge passes `MAX_PNG_EDGE_PX` (8192) is scaled down to it, aspect kept, with the notice `<file>: scaled to <w> × <h> px, the largest image this exports`; `--quality <1–100>` (default 90) applies to JPEG and WebP and is a usage error with PNG
- [ ] A conversion that would write more than `MAX_PNG_EXPORT_SLIDES` (200) images is refused (exit 3) with `imagesRefusal`'s sentence reworded for its noun (`Image export is limited to 200 pages at a time. Choose a smaller range with --pages.`), naming the option that narrows it for that product
- [ ] Output naming per "Image naming": one image is `<stem>.<ext>` (or exactly `-o`); several are `<stem>-<n>.<ext>` where `<n>` is the page or slide number in the source (not the position in the selection), zero-padded to `max(2, digits of the source's page count)`; `-o <dir/name.ext>` with several images uses `<dir/name>-<n>.<ext>`; `--zip` writes one `<stem>.zip` (STORED entries, through `core/zipWriter.ts` as SPEC-104's archive is) holding the images under their numbered names; `-o -` needs a single image or `--zip`; the `--json` record's `files` lists every image (or the archive)
- [ ] A PDF input to images: `src/lib/pdfRaster.ts` gains `rasterisePdfPages(bytes, pages: number[], size: (page: { widthPt; heightPt }) => number, type, quality)` on the same lazily loaded legacy pdf.js build, rendering pages one at a time and releasing each canvas, and `rasterisePdfFirstPage` becomes a caller of it (the library's thumbnails unchanged); `--pages` selects, `--scale`/`--dpi`/`--width`/`--height` size it with `--scale 1` meaning 96 dpi; a password-protected PDF is exit 2 with SPEC-107's message, a page pdf.js cannot render is exit 6 naming the page
- [ ] An image input to another image format: decoded in the page with `createImageBitmap(blob, { imageOrientation: "from-image" })` (so a JPEG's EXIF rotation is applied), drawn at the requested size (default: its own pixel size) and encoded with `canvas.toBlob(type, quality)`; an animated GIF or WebP keeps its first frame with the notice `only the first frame was kept`; SVG goes through `core/svgToPng.ts`'s `convertSvgToPng(bytes, size)` first; a same-format image conversion is offered only with a size option (without one it is refused, exit 3, `<file> is already a .png — give --width, --height, --scale or --dpi to resize it`); the written file carries no metadata, which the docs say
- [ ] An image input to `pdf` is a browser-tier row through SPEC-108's `convertImageToPdf(source, setup, title)` with `BROWSER_CONVERT_STEPS`, run in the page on the `MediaModel` the bridge opened; the setup starts at `SHEET_DEFAULTS` and the options under "Image to PDF" override it; the result is validated by a new pure `validateSheetSetupCm(setup): string | null` in `core/sheetLayout.ts` with the messages under "Image to PDF" (exit 1 for that input) — `useSheetSetup`, whose messages depend on the unit preference, is unchanged; the PDF is titled by `--title` or the input's stem; a PDF input to `pdf` stays unoffered
- [ ] `RenderRequest` (SPEC-143) widens additively with the image targets and their options and the bridge posts each image as its own output (`/0`, `/1`, …, in page order) or, with `--zip`, the archive alone; `AutomationBridge.version` stays 1; a `--verbose` run shows `rendering page <i> of <n>` progress through SPEC-143's phase text
- [ ] `docs/cli.md` gains "Images": targets per product, pixel-size options and the clamp, naming and `--zip`, PDF pages to images, image to image (orientation, first frame, no metadata), image to PDF and its sheet options
- [ ] e2e (`e2e/SPEC-144.spec.ts`, each test given 180 s): `convert mini-deck.pptx --to png` writes one numbered PNG per shown slide, and slide 1's bytes equal the app's own SPEC-104 single-slide PNG download at the same scale; `convert multipage.docx --to jpg --pages 1 --width 800` (SPEC-143's test-built document) writes one `.jpeg` whose decoded size is 800 px wide at the page's aspect; `convert mini-book.xlsx --to webp --zip` writes one archive whose entries decode as WebP; `convert spec107-pages.pdf --to png --dpi 72` writes three PNGs whose pixel sizes equal the pages' displayed point sizes (612 × 792, the `/Rotate 90` A4 page as 842 × 595, and 720 × 540); `convert fixture-image.png --to pdf --page-size A4 --fit contain` writes a one-page A4 PDF (SPEC-107's structural reader) that the library opens; `convert fixture-image.png --to png` is refused with the "already a .png" message

## Out of scope
Combining several images into one image or one PDF (SPEC-109/110's assembly;
a follow-up once SPEC-110 lands), merging or splitting PDFs (SPEC-111), and
any image edit beyond resizing — crop, rotate, adjust and annotate stay
SPEC-112's editor. GIF, SVG, TIFF, AVIF or HEIC output. Preserving EXIF, ICC
or other metadata. Colour management. Vector output for pages. Thumbnails of
the library's own size rules. Animated output.

## Reference

### Dependencies restated
- `specs/143-headless-render-host-and-pdf.md` — the render host, the bridge
  (`AutomationBridge` version 1, `RenderRequest`, `open`, `render`), the
  output route (`POST /__docu-edit-cli/output/<token>/<n>` with
  `x-docu-edit-mime`), `PageJob` (a `PdfBuildInput` plus notices), page
  producers, per-product page options (`--pages`, `--slides`,
  `--include-hidden`, `--sheet`), `--timeout`, phase text, isolation; exit 5
  and 6. `specs/141-…`/`142-…` — grammar, exit codes, report, `--json`
  `files`, output naming (`<stem>`), atomic writes, `openBytes`.
- `specs/104-export-deck-to-pdf-and-images.md` / `presentation/deckExport.ts`
  — `MAX_PNG_EDGE_PX` (8192), `MAX_PNG_EXPORT_SLIDES` (200), `PNG_SCALES`
  (`[1, 2, 4]`; the dialog's default scale is 2), `pngPixelSize(slideWidthCm,
  slideHeightCm, scale) → { widthPx; heightPx; clamped }` (`widthPx =
  round(cmToPx(widthCm) × scale)`, clamped on the longer edge), `imagesRefusal(count)`, `slidePngFileName` (the app's own naming,
  which the command does not use — files are named after files), the STORED
  zip of SPEC-104's archive through `core/zipWriter.ts`; `useDeckExport`'s
  image plan (Slides layout, one slide per page).
- `src/lib/pdfAssets.ts` — `registerRasterFaces(fonts)` registers the PDF's
  embedded faces with `document.fonts` under the names `faceCssFont` uses,
  once per face per session.
- `src/lib/core/pdfRaster.ts` — `rasterisePdfPage(page: PdfPageInput,
  widthPx): Promise<Uint8Array>` (PNG; the height follows the page's aspect;
  fonts loaded through `faceCssFont` and `document.fonts.load`),
  `paintPdfPage(context, page, scale, images)`.
- `specs/107-images-and-pdfs-in-the-library.md` — `MediaModel` (format `png`
  `jpeg` `gif` `svg` `webp` `pdf`; parts with content type and pixel size),
  the password-protected PDF message, `src/lib/pdfRaster.ts`'s
  `rasterisePdfFirstPage(bytes, maxEdgePx)` on the legacy pdf.js build,
  loaded lazily with its worker as a separate chunk.
- `specs/108-the-assembly-sheet-page-size-and-margins.md` —
  `convertImageToPdf(source: MediaModel, setup: SheetSetup, title: string,
  steps?: ConvertSteps): Promise<Uint8Array>` and `BROWSER_CONVERT_STEPS` in
  `src/views/library/useConvertToPdf.ts`; `SheetSetup { paper: PaperSizeName |
  "Fit to content"; widthCm?; heightCm?; orientation; margins; fit: "contain" |
  "cover" | "stretch" | "actual"; align: SheetAlign; background: "#rrggbb" |
  "transparent"; dpi: 72 | 150 | 300 | 600 }`, `SHEET_DEFAULTS`,
  `SHEET_ALIGNS` (nine points, `"top-left"` … `"bottom-right"`),
  `PAPER_SIZE_NAMES` (`A4 Letter Legal A3 A5 Tabloid Custom`),
  `SHEET_DPI_CHOICES` (`72 150 300 600`), `MAX_SHEET_CM` (500).
- `src/lib/core/svgToPng.ts` — `convertSvgToPng(svgBytes, size?) →
  { data; widthPx; heightPx }`.
- `specs/112-edit-an-image.md` — `currentImageBytes(model)`; a file opened by
  the command has no edit stack, so its current image is its source.
- Equality: `equalBytes` (`src/lib/core/equality.ts`); decoded pixel sizes via
  `core/imageDimensions.ts`'s `decodeImageDimensions`.

### Pixel size
| Option | Page or slide (`widthPt` × `heightPt`) | Image input (`w` × `h` px) |
| --- | --- | --- |
| `--scale s` (default 2) | `widthPx = round(widthPt / 72 × 96 × s)` — for a slide, SPEC-104's `pngPixelSize` | `round(w × s)` — default for images is its own size, not 2 |
| `--dpi d` | `round(widthPt / 72 × d)` | usage error (an image has no physical size here) |
| `--width W` / `--height H` | that edge, the other by aspect | same |
| both `--width` and `--height` | usage error for pages; for images, exact size (aspect may change) | |

Heights always follow from the width by the source's aspect except where
both are given for an image. The `MAX_PNG_EDGE_PX` clamp applies last, to
every result.

### Image naming
| Situation | Files |
| --- | --- |
| one page/slide/image | `report.png` |
| `report.docx --to png` (12 pages) | `report-01.png` … `report-12.png` |
| `--pages 3,10` of 12 | `report-03.png`, `report-10.png` |
| `scan.pdf --to jpg` (140 pages) | `scan-001.jpeg` … |
| `-o out/cover.webp`, several | `out/cover-01.webp` … |
| `--zip` | `report.zip` containing `report-01.png` … |

The extension written is the canonical one (`jpeg`), whatever alias was
asked for, except that `-o` is used exactly as given. Numbered names depend
on the page count, which only the opened input knows, so they are checked as
they are written (SPEC-141: colliding with a file the run already wrote is
exit 4).

### Image to PDF
| Option | `SheetSetup` field | Values |
| --- | --- | --- |
| `--page-size <name>` | `paper` | a `PAPER_SIZE_NAMES` entry except `Custom`, or `fit` (= "Fit to content") |
| `--page-width <cm>` `--page-height <cm>` | `paper: "Custom"`, `widthCm`, `heightCm` | both or neither |
| `--orientation portrait\|landscape` | `orientation` | |
| `--margin <cm>` | `margins` (all four) | |
| `--fit contain\|cover\|stretch\|actual` | `fit` | |
| `--align <point>` | `align` | a `SHEET_ALIGNS` entry |
| `--background <#rrggbb\|transparent>` | `background` | |
| `--dpi <72\|150\|300\|600>` | `dpi` | the ceiling for the embedded image |

`validateSheetSetupCm` returns the first failing rule's message, in this
order: `"<name>" is not a page size — use one of A4, Letter, Legal, A3, A5,
Tabloid or fit.`; `A page side must be more than 0 and at most 500 cm.`;
`Margins cannot be negative.`; `The margins leave no room on a <w> × <h> cm
page.` (when left + right ≥ width or top + bottom ≥ height; `fit` pages
skip this check); `"<value>" is not a colour. Use #rrggbb or
transparent.`; `Resolution must be one of 72, 150, 300, 600 dpi.`. Unknown
`--fit`, `--align` and `--orientation` values are usage errors from the
arguments alone.

### Bridge additions
```ts
// AutomationBridge.open, for a media file: resolves { kind: "media", id: "", title }
// without storing or navigating; the model stays in the page for render().
type MediaRenderRequest =
  | { kind: "media"; target: "png" | "jpeg" | "webp"; pages?: string; scale?: number; dpi?: number;
      widthPx?: number; heightPx?: number; quality?: number; zip?: boolean }
  | { kind: "media"; target: "pdf"; setup: SheetSetup; title: string };
```
The text, presentation and spreadsheet variants gain the same `scale`,
`dpi`, `widthPx`, `heightPx`, `quality` and `zip` fields with image targets.

### Tests
- `src/lib/core/pdfRaster.spec.ts` gains `rasterisePdfPageAs` cases with the
  recording stand-in (white composite for JPEG only; type and quality passed
  to the encoder) and asserts `rasterisePdfPage` paints exactly as before.
- `src/lib/core/sheetLayout.spec.ts` gains `validateSheetSetupCm`: every
  message in order, `fit` pages skipping the margin check.
- `src/lib/pdfRaster.spec.ts` — `rasterisePdfPages` with a stubbed pdf.js
  (one canvas alive at a time; page selection; first-page caller unchanged).
- `src/cli/imageOptions.spec.ts` — the pixel-size table, every exclusivity
  and usage error, the clamp notice, quality, naming (padding from the
  source's page count, `-o` with several images, `--zip`, standard output),
  the 200-image refusal per product, the image-to-PDF option table.
- `src/automation/imageRender.spec.ts` (jsdom, stubbed producer and encoder)
  — request widening, one POST per image in order, archive-only with `--zip`.
- `e2e/SPEC-144.spec.ts`; every SPEC-104, SPEC-107 and SPEC-108 test passes
  unchanged.

### Specs this changes
- `specs/143-…` — `RenderRequest` gains image targets (additive, version 1).
- `specs/104-…` / `core/pdfRaster.ts` — a JPEG/WebP sibling of
  `rasterisePdfPage`; PNG output unchanged.
- `specs/107-…` / `src/lib/pdfRaster.ts` — a multi-page entry point the
  first-page preview now uses.
- `specs/108-…` / `core/sheetLayout.ts` — a pure, cm-based
  `validateSheetSetupCm`; `useSheetSetup` unchanged.
- `specs/141-…` — the registry gains the image rows and `docs/cli.md` a
  section.
- `specs/000-roadmap.md` — Phase 19 row.
