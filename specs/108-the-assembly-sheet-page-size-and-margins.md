---
id: SPEC-108
title: The assembly sheet — page size, margins and fit
status: draft
area: layout
depends_on: [SPEC-027, SPEC-034, SPEC-058, SPEC-064, SPEC-066, SPEC-107]
---
## Why
Every output the next three specs produce lands on a rectangle: a merged
image, a PDF built from photos, a page inserted between two PDFs. Each one
needs the same six answers — how big is the sheet, which way round, how much
margin, what happens when the picture's shape does not match the sheet's,
what fills the leftover, and at what resolution. Written three times, they
would be three slightly different answers, and the product would have a
"page size" that means something different in each dialog.

So the sheet comes first, with its own small end-to-end feature to prove it:
turn one image in the library into a PDF. That is a real errand — someone
needs a scan as a PDF because the form only takes PDFs — and it is the whole
geometry exercised once, with a preview a user can see before anything is
written.

**This file is the normative source for `SheetSetup`, for the fit rules, for
the resolution ladder and for the shared setup panel.** SPEC-109, SPEC-110
and SPEC-111 spend it; none of them redefines a millimetre of it.

## Acceptance criteria
- [ ] `src/lib/core/sheetLayout.ts` exports `SheetSetup`, `SHEET_DEFAULTS`, `resolveSheet(setup, itemAspect?): SheetGeometry`, `placeOnSheet(itemPx, box, fit, align): RectCm` and `sheetPixelSize(geometry, dpi): { widthPx; heightPx }` per the Reference — pure maths in centimetres, no DOM, no canvas, framework-free, importing only core siblings
- [ ] Paper is SPEC-027's, not a second table: `SheetSetup.paper` is a `PaperSizeName` resolved through `core/pageSetup.ts`'s `PAPER_SIZES`/`PAPER_SIZE_NAMES`, with `"Custom"` taking `widthCm`/`heightCm`, `orientation` applied through the existing `orientationOf`/swap helpers, and one addition owned here: `"Fit to content"`, a sheet whose size **is** the content's own size at the chosen resolution plus the margins, so converting a 1200×800 screenshot produces a PDF page of exactly that shape rather than a letterboxed A4
- [ ] Margins are SPEC-027's model and SPEC-066's presets: four independent centimetre values, the `MARGIN_PRESETS` (Normal / Narrow / Moderate / Wide) offered as a preset select plus per-edge inputs, a link toggle that keeps all four equal, and validation that the margins leave a content box of at least `MIN_CONTENT_CM = 1` in both axes — a margin set that would not is rejected at input with the reason, never silently clamped
- [ ] Fit has four values and each is specified rather than named: `contain` (scale to fit inside the content box, whole item visible, aspect preserved — the default), `cover` (scale to fill, aspect preserved, overflow cropped equally on the two overflowing edges), `stretch` (fill exactly, aspect broken), `actual` (one image pixel = one `dpi` dot, no scaling, cropped to the content box when larger); `align` places the result within the content box on a nine-point grid, defaulting to `center`
- [ ] The leftover is explicit: `background` is `"#rrggbb"` or `"transparent"`, defaulting to `#ffffff` for PDF output and `transparent` for PNG output, and `transparent` is refused with a stated reason for JPEG and for PDF (PDF pages have no alpha; the value falls back to white and the panel says so before the user commits)
- [ ] Resolution is a ladder, not a free number: `SHEET_DPI_CHOICES = [72, 150, 300, 600]` with 150 the default for raster output, PDF placement independent of it (a PDF page is vector geometry — `dpi` bounds only how many samples an embedded image carries), and an estimated output size shown live so a 600 dpi A3 is a number the user sees before pressing the button
- [ ] `src/components/assembly/SheetSetupPanel.vue` is the one control every assembly surface embeds — a dumb `:controller`-taking component in the `TransitionsPanel.vue` mould — with testids `sheet-paper`, `sheet-portrait`, `sheet-landscape`, `sheet-width`, `sheet-height`, `sheet-margin-preset`, `sheet-margin-top|right|bottom|left`, `sheet-margin-link`, `sheet-fit`, `sheet-align`, `sheet-background`, `sheet-dpi`, `sheet-estimate` and `sheet-error`, every control native and keyboard-operable, and every length rendered in the user's SPEC-064 measurement unit while the model stays centimetres
- [ ] `src/components/assembly/SheetPreview.vue` (`sheet-preview`) draws the resolved sheet to scale — paper outline, margin guides, the placed item's box and the cropped-away region shaded — updating on every control change without touching the source bytes, and states the resulting page size and placed size in the active unit beneath it
- [ ] The end-to-end surface: a library image row's overflow menu and the selection bar gain "Convert to PDF…" (`menu-convert-pdf`, `bulk-convert-pdf`, the bulk form limited to one item in this spec), opening `ConvertToPdfDialog.vue` (`convert-pdf-dialog`) — a `BaseDialog` shell holding `SheetSetupPanel` + `SheetPreview`, a title field defaulting to the image's title, and actions `convert-pdf-save` ("Save to library"), `convert-pdf-download` and `convert-pdf-cancel`
- [ ] Conversion goes through the existing writer and adds no serialisation of its own: one `PdfPageInput` sized `cmToPt(sheet.widthCm) × cmToPt(sheet.heightCm)` carrying one `PdfImageItem` at the placed rectangle plus a background `PdfRectItem`, the image encoded by `encodeImage` and handed to `buildPdf`; SVG is rasterised first through `core/svgToPng.ts` at the placed size and the dialog says so
- [ ] `encodeImage` gains an optional `maxDpi` (default `300`, preserving SPEC-034's cap and every existing assertion) so a 600 dpi conversion is not silently downsampled to 300, and `cover`/`actual` crops are applied **before** encoding — a cropped-away region never reaches the output file
- [ ] Saving to the library creates a **new** media document (SPEC-107's `MediaModel`, format `pdf`, its `pdf` summary filled by `parsePdf` on the bytes just written) titled `<title>.pdf` with SPEC-003's duplicate-title numbering, leaving the source untouched; downloading goes through `core/triggerDownload.ts` with `PDF_MIME_TYPE`; both paths report a failure without mutating anything
- [ ] Choices persist for the session in the composable and, for the three that are really preferences, in `src/lib/preferences.ts`: `assemblyPaper`, `assemblyMargins` and `assemblyDpi`, defaulting to the user's existing page-setup defaults where they have them — a user who works in Letter does not pick Letter every time
- [ ] Determinism and purity: `resolveSheet`/`placeOnSheet` are pure and total (no `Date`, no randomness, no DOM), every returned length is rounded to 3 decimal places by one shared helper so a value never differs in the last bit between the preview and the output, converting the same image twice with the same setup yields `equalBytes`-identical PDFs (the writer stamps no date beyond the caller-supplied one), and nothing in the flow mutates the source model or its `MediaPart`
- [ ] Tests: `src/lib/core/sheetLayout.spec.ts` covers all four fits against a wide and a tall item on a wide and a tall sheet (16 cases, exact numbers), the nine alignments, "Fit to content" in both orientations, the margin-validation boundary, the dpi ladder and the rounding rule; `ConvertToPdfDialog.spec.ts` covers control enablement and the transparent-background refusals; e2e (`e2e/SPEC-108.spec.ts`): open the library with `spec107-shot.png` → row menu → "Convert to PDF…" → paper A4, fit `contain`, margins Narrow → the preview reports the placed size → Save to library → a new `PDF` row appears whose details panel reads 1 page at 21 × 29.7 cm → repeat with "Fit to content" → the new PDF's page is the image's own aspect within 0.01 cm

## Out of scope
Multiple items on one sheet (SPEC-109 for images, SPEC-110 for PDF pages) —
this spec places exactly one. Bleed, crop marks, spreads, facing pages,
gutters and binding offsets. Colour management, ICC profiles and CMYK output.
Headers, footers, page numbers or captions on an assembled sheet (SPEC-110
adds page numbers where they are a print artefact). Named sheet presets a
user can save. Templates. Vector passthrough for SVG (it is rasterised, and
the panel says so). Converting a PDF **to** an image — that needs a
rasteriser this product does not have (SPEC-107's Out of scope explains why).
Any output format beyond PDF here; PNG and JPEG output arrive with SPEC-109,
which is the first spec that needs a raster encoder of its own.

## Reference

### Dependencies restated
- `specs/027-page-setup.md` / `src/lib/core/pageSetup.ts` — `PAPER_SIZES`
  (A4 21 × 29.7, Letter 21.59 × 27.94, Legal, A3, A5, Tabloid),
  `PAPER_SIZE_NAMES` (those six plus `"Custom"`), `matchPaperSize(pageSize)`
  within a 0.05 cm tolerance, `Orientation`/`orientationOf(pageSize)` (wider
  than tall = landscape), and the swap helper the orientation toggle uses.
  `MARGIN_PRESETS` and `MARGIN_PRESET_NAMES` are in `src/lib/preferences.ts`.
- `specs/034-print-and-export-pdf.md` — `buildPdf(input: PdfBuildInput):
  Uint8Array` over `PdfPageInput { widthPt, heightPt, items, links, headings,
  images }`, `PdfImageItem { src, xPt, yPt, widthPt, heightPt, onTop? }`
  positioned from the page's **top**-left, `PdfImageData { widthPx, heightPx,
  filter: "DCTDecode" | "FlateDecode", colorSpace, data, alpha? }`,
  `encodeImage` passing a baseline JPEG through untouched and drawing
  everything else to a canvas, `PDF_MIME_TYPE`, `pdfFileName(title,
  fallback)`. **Where SPEC-104 has landed these are `core/pdfBuild.ts` /
  `core/pdfImages.ts` and `PdfPageInput` already carries `rects?:
  PdfRectItem[]`; where it has not, this spec adds the single rect the
  background needs, with the same field shape SPEC-104 specifies, so the two
  converge rather than collide.**
- `specs/064-…` — the measurement-unit preference and its formatters: every
  displayed length goes through them; the model is centimetres everywhere.
- `specs/107-…` — `MediaModel`, `MediaFormat`, the media kind at
  `/media/:id`, `parsePdf(bytes)` and its `PdfSummary`, and the rule that a
  stored media payload is a `MediaPart` in `model.media`.
- `src/lib/core/units.ts` — `PX_PER_CM = 96 / 2.54`, `cmToPx`/`pxToCm`,
  `MeasurementUnit` (`cm` | `mm` | `in` | `pt`) with `toCm`/`fromCm`,
  `parseLengthToCm`, `formatLength` and `roundForDisplay`, which are what the
  panel's inputs use. Points come from `fromCm(cm, "pt")`; the two named
  wrappers `cmToPt`/`ptToCm` are **added** to that module beside `cmToPx`,
  since the PDF path spells them everywhere and the module has no such pair
  today.

### The setup
```ts
// src/lib/core/sheetLayout.ts
export type SheetFit = "contain" | "cover" | "stretch" | "actual";
export type SheetAlign =
  | "top-left" | "top" | "top-right"
  | "left" | "center" | "right"
  | "bottom-left" | "bottom" | "bottom-right";

export interface SheetSetup {
  /** A PAPER_SIZE_NAMES entry, "Custom", or "Fit to content". */
  paper: PaperSizeName | "Fit to content";
  widthCm?: number;   // Custom only
  heightCm?: number;  // Custom only
  orientation: Orientation;
  margins: { topCm: number; rightCm: number; bottomCm: number; leftCm: number };
  fit: SheetFit;
  align: SheetAlign;
  /** "#rrggbb" or "transparent". */
  background: string;
  /** One of SHEET_DPI_CHOICES. */
  dpi: number;
}

export interface SheetGeometry {
  widthCm: number; heightCm: number;
  content: { xCm: number; yCm: number; widthCm: number; heightCm: number };
}

export const SHEET_DEFAULTS: SheetSetup;   // A4, portrait, Normal margins, contain, center, #ffffff, 150
export const SHEET_DPI_CHOICES: readonly number[] = [72, 150, 300, 600];
export const MIN_CONTENT_CM = 1;
export const MAX_SHEET_CM = 500;           // refuses a sheet no output path can produce
```

`resolveSheet(setup, itemAspect?)` returns the sheet and its content box.
For `"Fit to content"` the sheet is `itemPx / dpi` converted to centimetres
plus the margins, which is why `itemAspect` is optional everywhere else and
**required** there — calling it without one throws rather than inventing a
size.

### The fit rules, stated as arithmetic
With a content box `W × H` and an item of `w × h` (pixels at `dpi`, so
`wCm = w / dpi * 2.54`):

| Fit | Placed size | Cropped |
| --- | --- | --- |
| `contain` | `s = min(W/wCm, H/hCm)`, size `wCm·s × hCm·s` | nothing |
| `cover` | `s = max(W/wCm, H/hCm)`, size clipped to `W × H` | the overflow, split by `align` |
| `stretch` | `W × H` | nothing (aspect broken) |
| `actual` | `wCm × hCm`, clipped to `W × H` | the overflow, split by `align` |

`align` positions the placed box in the content box (and, for the cropping
fits, chooses which part of the item survives): `center` splits the
difference, an edge value takes that edge. Every result is rounded by the
shared `r3` helper — the preview and the writer must agree exactly, and two
independent roundings are how they stop agreeing.

### Why the sheet is its own spec
Three consumers, one geometry. The alternative — each merge spec carrying its
own page-size control — was rejected for the reason the roadmap's
conventions already give: one normative source per shared concept. The
"Convert to PDF" feature here is deliberately small; its job is to make the
geometry shippable and testable end to end before three specs depend on it,
not to be the product's assembly story. That story is SPEC-110's.

### Tests
- `src/lib/core/sheetLayout.spec.ts` — the 16-case fit matrix with exact
  expected rectangles, the nine alignments, "Fit to content" both ways,
  margin validation at the boundary, `MAX_SHEET_CM`, the dpi ladder,
  `sheetPixelSize` rounding, and the `r3` agreement property (preview and
  writer given the same setup produce identical numbers).
- `src/components/assembly/SheetSetupPanel.spec.ts` — unit conversion in and
  out, the margin link toggle, the transparent-background refusals, error
  states; `SheetPreview.spec.ts` — the drawn rectangles for two fits.
- `src/views/library/useConvertToPdf.spec.ts` — the built `PdfPageInput` for
  three setups, `maxDpi` pass-through, no mutation of the source model,
  byte-identical output on a repeat run.
- `e2e/SPEC-108.spec.ts` as in the acceptance criteria.

### Specs this changes
- `specs/034-…` — `encodeImage` gains an optional `maxDpi` (default keeps
  today's 300 dpi cap and every existing assertion); where SPEC-104 has not
  landed, `PdfPageInput` gains `rects?: PdfRectItem[]` in exactly the shape
  SPEC-104 specifies.
- `specs/066-preferences.md` — three new preference keys (`assemblyPaper`,
  `assemblyMargins`, `assemblyDpi`) in the existing store, with defaults
  taken from the user's page-setup defaults.
- `specs/067-…` — the row menu and selection bar gain one entry each, shown
  only for image entries.
- `specs/000-roadmap.md` gains this spec's phase row.
