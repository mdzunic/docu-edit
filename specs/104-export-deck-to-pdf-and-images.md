---
id: SPEC-104
title: Export a deck to PDF, images and handouts
status: draft
area: export
complexity: normal
story_points: 8
depends_on: [SPEC-034, SPEC-079, SPEC-080, SPEC-081, SPEC-083, SPEC-084, SPEC-085, SPEC-086, SPEC-088, SPEC-090, SPEC-093]
---
## Why
A deck's second-most common fate, after being sent, is being *printed* — as
handouts for the room, as notes pages for the speaker, as a PDF for the people
who could not come. This product cannot produce a single page of any of it:
SPEC-034's PDF writer and print pipeline are both wired to the word processor's
paginated canvas, SPEC-088 exports HTML that nobody can staple, and there is no
way to get a PNG of one slide for a wiki or a chat.

Nothing here needs inventing: that writer already turns positioned text, images
and canvas-painted shapes into a PDF with embedded subset fonts, and the
filmstrip already renders any slide off the editor stage through `<SlideStage
:interactive="false">`. This spec moves the PDF layer into `src/lib/core` where
both products can reach it, adds the one drawing primitive slides need
(filled/stroked rectangles) and puts a deck-shaped sheet in front of it. It also
fills the model gap printing exposes — a deck cannot say "skip this slide" — so
**this file is the normative source for `Slide.hidden` and its two wire
mappings**, for the handout/notes sheet geometry and for deck raster export.

## Acceptance criteria
- [ ] `Slide` gains `hidden?: true` (absent = shown; only `true` is ever stored) per the Reference: pptx reads and writes `p:sld/@show="0"` as a single-attribute splice inside `preserved["@root"]` (the SPEC-093 `showSpecialPlsOnTitleSld` precedent — every other root attribute keeps its bytes and position), odp reads and writes `presentation:visibility="hidden"` on the page's `style:drawing-page-properties` through SPEC-085's `StyleRegistry.drawingPageStyle`; readers set no `dirty`/`notesDirty`/`structureDirty` and mutate no `opaque` part
- [ ] The filmstrip's slide menu (`SlideFilmstrip.vue`, `item.testid` rows) gains "Hide slide"/"Show slide" (testid `slide-menu-hide`): a hidden thumbnail renders at 45% opacity with its number struck through and `filmstrip-hidden` on the item, the toggle records one undo entry (`"Hide slide"` / `"Show slide"`) via SPEC-080's `record` → mutate → `commit()` order, calls `markSlideEdited(slide, model.format)` and persists write-through; the editor never skips a hidden slide (filmstrip click and PageDown/PageUp still reach it, it is fully editable)
- [ ] SPEC-084's `usePlayback` gains one optional dep, `isHidden?: (slideIndex: number) => boolean` (absent ⇒ nothing hidden): `advance`, `back` and `goTo` step over hidden slides, `atEnd` is reached when no later visible slide exists, and a direct `/slides/:id/present/:n` deep link to a hidden slide still shows it — `useDeckActions.present()` pushes that very route for the *current* slide, so presenting from a hidden slide opens on it and only advancing steps over one; the rest of the advance pipeline in `usePlayback.ts` is unchanged and stays SPEC-084's
- [ ] SPEC-034's PDF layer moves into core as flat modules per the Reference (two-phase, behaviour-identical, colocated specs moving with their subjects, `core/boundary.spec.ts` green — core is flat, no `../` and no subdirectory) and is generalised without changing its text-side behaviour: `ExtractOptions` gains `imageSelector` (default `"img.doc-image"`), `canvasSelector` (default `"canvas.doc-shape-canvas, canvas.doc-chart-canvas"`) and `rectSelector` (default absent), `PdfPageInput` gains optional `rects?: PdfRectItem[]` which `core/pdfBuild.ts` paints — with no `rectSelector` no `rects` key is emitted and every existing SPEC-034 assertion is untouched
- [ ] `src/lib/presentation/printLayout.ts` (framework-free, importing only `lib/core` and presentation siblings) exports `DeckPrintLayout`, `DECK_PRINT_LAYOUTS`, `sheetPageSize(options)`, `resolveSlideRange(model, options)` and `planSheet(model, options): SheetPage[]` per the Reference: 1/2/3/4/6/9-up handouts, notes pages and full-bleed slides, slot rectangles in cm, note-line rows, page count, hidden-slide filtering, and range parsing through `core/pageRange.ts`'s `parseNumberRange(input, total, "slide")`
- [ ] The Download menu (`PresentationDownloadButton.vue`, alongside `download-copy-as` and SPEC-088's `presentation-export-html`) gains "PDF (.pdf)…" (`presentation-export-pdf`) and "Images (.png)…" (`presentation-export-images`), both opening one `BaseDialog`-shelled `DeckPrintDialog.vue` (`deck-print-dialog`) on the matching tab (Cmd/Ctrl+P inside `/slides/:id` opens it on the PDF tab), whose controls are all native and keyboard-operable: range radios `deck-print-range-all` / `deck-print-range-current` / `deck-print-range-custom` with `deck-print-range-input` and `deck-print-range-error`, layout select `deck-print-layout` (Slides, Notes pages, Handout 1/2/3/4/6/9 per page), paper select `deck-print-paper` (A4 default, from `core/pageSetup.PAPER_SIZES`) with orientation radios `deck-print-portrait`/`deck-print-landscape`, checkboxes `deck-print-note-lines` (handouts 1/2/3/4 only, forced on and disabled for 3-up, disabled for 6/9-up), `deck-print-frame` ("Frame slides", default on except for the Slides layout) and `deck-print-hidden` ("Include hidden slides", default off, disabled when the deck has none), PNG scale select `deck-print-scale` (1×/2×/4×, default 2×), and actions `deck-print-print`, `deck-print-pdf`, `deck-print-images`, `deck-print-cancel`
- [ ] The render for export never touches the editor stage: `DeckPrintSheet.vue` (`deck-print-sheet`) mounts off-screen (`position: fixed; left: -20000px`, `aria-hidden`, `inert`) one `deck-print-page` per planned page, each `deck-print-slot` holding `<SlideStage :interactive="false" :scale="…" :raster-scale="2">`; `SlideStage.vue` gains a `rasterScale` prop (Number, default 1) forwarded to `SlideFrameView.vue`, which passes it as `paintShapeIntoCanvas(canvas, shape, w, h, zoomFraction)`'s existing fifth argument (SPEC-063's zoom hook) so canvas-painted shapes get a backing store of `devicePixelRatio × rasterScale` and are never upscaled by the 300 dpi sampler in `core/pdfImages.ts`
- [ ] Print goes through the browser: `deck-print-print` marks `body.is-deck-print`, shows only the sheet under `@media print`, injects `@page { size: {w}cm {h}cm; margin: 0 }` — the deck's own `slideWidthCm`/`slideHeightCm` (SPEC-090) for the Slides layout, the chosen paper otherwise — waits for the sheet to settle (`SHEET_SETTLE_DEADLINE_MS = 8000`, the SPEC-034 `realizeAllPages` precedent), calls `window.print()`, then restores; selection overlays, resize handles, placeholder prompts (`slide-placeholder-prompt`) and SPEC-086 animation badges carry `.print-exclude`, the marker class `core/pdfExtract.ts` already honours, so they neither print nor export
- [ ] PDF export builds real pages, not screenshots: one `PdfPageInput` per sheet page from `extractPageContent(pageEl, widthPt, heightPt, {imageSelector, canvasSelector, rectSelector})`, fonts loaded and subset exactly as `usePdfExport.ts` does, `slideOutline(slide).title ?? \`Slide ${n}\`` emitted as a level-1 `PdfHeading` per slide so the PDF gets bookmarks, `/Title` set from `model.title`, downloaded as `pdfFileName(model.title, DEFAULT_PRESENTATION_TITLE)` through `core/triggerDownload.ts`'s `triggerBlobDownload(bytes, name, PDF_MIME_TYPE)`; slide text is selectable and searchable in the result
- [ ] PNG export renders the same sheet pages at the Slides layout through `core/pdfRaster.ts`'s `rasterisePdfPage(page, widthPx): Promise<Uint8Array>` — an absolute width, never a multiplier: `useDeckExport` computes `widthPx = Math.round(cmToPx(model.slideWidthCm) * scale)` and owns the `MAX_PNG_EDGE_PX = 8192` clamp, so a 16:9 deck (33.867 × 19.05 cm at `PX_PER_CM`) yields exactly 1280×720 / 2560×1440 / 5120×2880 at 1×/2×/4×, and any scale whose longer edge would pass the cap is clamped down before the call, `deck-print-estimate` saying so; one slide downloads as `{title}-slide-{NN}.png` (zero-padded to the deck's digit count, minimum 2), the whole range as `{title}-slides.zip` built with `core/zipWriter.ts`'s `zip(entries)` and `store: true` (PNG is already compressed), capped at `MAX_PNG_EXPORT_SLIDES = 200`
- [ ] Notes pages and handout note lines match the Reference geometry: a notes page carries the slide fitted into the top half and its `Slide.notes` paragraphs (SPEC-083) below at 11 pt / 1.3 line height, overflowing onto continuation pages that repeat the notes only; note lines are 0.5 pt `#808080` rules `NOTE_LINE_SPACING_CM = 0.85` apart; "Frame slides" draws a 0.75 pt `#000000` border round each slot; handout and notes pages carry a 9 pt centred page number (`deck-print-page-number`), the Slides layout carries none
- [ ] Fidelity is stated, not hoped for: every frame exports fully visible and fully built (SPEC-086's editor-stage rule — an exit effect's end state would silently delete content from a printout), transitions have no print representation at all, a SPEC-095 `MediaFrame` exports as its poster or the neutral plate with the play badge excluded, and text inside a frame with `rotationDeg` prints unrotated at the frame's axis-aligned box (shape *geometry* rotation is exact — `paintShapeIntoCanvas` bakes it into the pixels) with `deck-print-rotation-note` shown when the selected range contains such a frame
- [ ] A 200-slide deck stays usable: pages are mounted, extracted and unmounted in a rolling window of `EXPORT_WINDOW_PAGES = 4` with `yieldToMainThread()` (`core/idleScheduler.ts`) between pages, `deck-print-progress` reports `Page {done} of {total}` with `role="status"`, `deck-print-abort` stops at the next page boundary leaving no download and no model mutation, printing a range over `PRINT_PAGE_LIMIT = 250` pages is refused with a message naming the limit, and each raster canvas is released (`canvas.width = 0`) after encoding
- [ ] Round trip and purity: an untouched deck still downloads byte-identically — per-part `equalBytes` (`src/lib/core/equality.ts`) after `unzip`, with `roundTrip.spec.ts`'s existing `expectUntouched` and odp splice-equality assertions unchanged; hide → save → reopen reproduces `hidden` exactly in both formats and re-saving the reopened model is byte-stable for the slide part; exporting mutates nothing (deep-frozen model input in the unit tests, no dirty flags, no store write); `buildPdf(input)` called twice on the same input is `equalBytes`-identical (the writer stamps no date). **Note the relation available here**: `equalSemantic`/`resolve` in `src/lib/equality.ts` are typed `(model: DocumentModel)` and are **not** usable on a `PresentationModel`; presentation round trips assert per-part `equalBytes` plus `equalStructural` over the `slides` array and explicit field-level `toEqual`
- [ ] Fixture coverage: `src/lib/presentation/fixtures/build.ts` gains `buildPptxHiddenFixtureBytes` / `buildOdpHiddenFixtureBytes` (a 3-slide deck whose middle slide is hidden, one slide carrying notes and a themed shape) plus `EXPECTED_HIDDEN_FLAGS`, appended under a banner comment without touching the four mini-deck slides, written to `e2e/fixtures/spec104-hidden.pptx` / `.odp` by `scripts/build-presentation-fixtures.ts`
- [ ] e2e (`e2e/SPEC-104.spec.ts`): open `e2e/fixtures/spec104-hidden.pptx` → the middle thumbnail shows `filmstrip-hidden` → Download menu → "PDF (.pdf)…" → layout "Handout 4 per page", range All, hidden off → PDF → the captured download starts `%PDF-1.7` and its page tree reads `/Count 1` (two visible slides, one sheet); switch the layout to "Slides" → `/Count 2`, then turn "Include hidden slides" on → `/Count 3` — that 2-vs-3 delta is what pins the option; switch to Images, scale 2×, range "1" → the captured PNG's IHDR (bytes 16–23) reads 2560×1440 → with slide 1 selected, Present → playback opens on slide 1 and one advance skips the hidden slide 2 to land on slide 3 → back in the editor hide slide 1 from the filmstrip menu → undo restores it → download, reopen → slide 2 is still hidden and slide 1 is not

## Out of scope
PDF/A, tagged-PDF accessibility beyond outline bookmarks, password protection,
printing from the play view, a printed presenter view, rehearsal timings on
handouts, a handout master or a notes master (SPEC-089 owns the master chain;
SPEC-093's slide-level header/footer/number/date fields export as the ordinary
frames they render as), editing the notes-page or handout template, a Notes and
Handouts tab, per-handout-page numbering schemes, JPEG/WEBP/SVG raster targets
(PNG only), exporting an animation build sequence as multiple pages, printing a
selection of frames, server-side or headless rendering, section-aware handouts
(SPEC-099's `sections` are not a print axis), and sending the export anywhere —
SPEC-045's generation API stays a follow-up.

## Reference

### Dependencies restated
- `specs/034-print-and-export-pdf.md` — **the mechanism this spec inherits; its
  PDF half is a generated file, not a browser print-to-PDF.** `buildPdf(input:
  PdfBuildInput): Uint8Array`, beside `PDF_MIME_TYPE` and `pdfFileName(title,
  fallback)`, writes PDF 1.7 by hand over `PdfPageInput { widthPt, heightPt,
  items: PdfTextItem[], links, headings, images }`, embedding subsets of the
  bundled Liberation/Carlito/Caladea faces; `extractPageContent(pageEl, widthPt,
  heightPt, options)` fills one off a *realised DOM page*, reading text through
  `Range.getClientRects` and both `img.doc-image` and the shape/chart canvases as
  images — a canvas is already re-used as an image rather than taught to the
  writer as a vector path, and this spec follows that precedent. `encodeImage`
  caps samples at 300 dpi of the printed size, `.print-exclude` is the shared
  screen-only marker both paths honour, and the print half is `window.print()`
  under an injected `@page` rule sized to the document's own geometry. **No
  dependency is added: that writer, `core/zipWriter.ts` and `canvas.toBlob` cover
  everything here.**
- `specs/079-slide-canvas-rendering.md` — `SlideStage.vue` renders a fixed
  `cmToPx(widthCm) × cmToPx(heightCm)` surface (`core/units.ts`, `PX_PER_CM =
  96/2.54`) fitted with `transform: scale()`, `:interactive="false"` being the
  render the filmstrip already uses; `SlideFrameView.vue` paints shapes with
  `paintShapeIntoCanvas(canvas, shape, widthPx, heightPx, zoomFraction = 1)`
  (`core/shapeCanvas.ts`), sizing the backing store at `devicePixelRatio *
  zoomFraction`; `slideOutline(slide)` is in `slideModel.ts`.
- `specs/080-…` writer rebuild rules, the Download menu and the undo stack
  (`record(label)` **before** mutating, then `commit()`); `specs/088-…` the
  `presentation-export-html` entry the two new items sit beside; `specs/083-…`
  `Slide.notes?: string[]`, absent when none — its "Printing notes" exclusion is
  what this fills; `specs/084-…` the normative advance pipeline (`usePlayback`'s
  `advance`/`back`/`goTo`/`atEnd` over `PlaybackStep`), which gains `isHidden?`
  and nothing else; `specs/085-…` the drawing-page style index and
  `StyleRegistry.drawingPageStyle`; `specs/086-…` the rule that the editor stage
  always renders frames fully visible, and the golden-fixture-first process
  re-used below; `specs/090-…` editable `slideWidthCm`/`slideHeightCm`, so a 4:3
  or portrait deck prints as such. SPEC-081's preservation slots:
  `Slide.preserved["@root"]` is the whole `p:sld` attribute string
  (leading-space-prefixed), re-emitted by `serializeSlideXml` as
  `<p:sld{rootAttrs}>`; `preserved["@page"]` is the odp `draw:page` attribute
  string, carrying `draw:style-name`.

### Hidden slides
**pptx** — `<p:sld show="0">`. `@show` lives in `preserved["@root"]`, which is
also the byte source for every other root attribute, so the slot is *mirrored*,
not consumed: the reader sets `slide.hidden = true` when `@root` carries `show="0"`
(or `show="false"`) and leaves the string alone; toggling splices just that
attribute in place (appended after the last attribute when absent, removed with
its leading space when the slide is shown again). Nothing else in the string
moves, so an untouched slide still serialises byte-identically.

**odp** — ODF 1.2 §15.36 puts `presentation:visibility` (`visible` | `hidden`) on
`style:drawing-page-properties`, not on `draw:page`. The reader's `StyleResolver`
(`src/lib/presentation/odp/reader.ts:93`) already indexes drawing-page styles by
name for SPEC-085's transition attributes; it gains `drawingPageVisibility(name)`
beside `drawingPageTransition(name)` (`reader.ts:129`). The writer mints the page's
`de-DP{n}` style through the existing `StyleRegistry.drawingPageStyle(extraProps,
baseProps)`, appending `presentation:visibility="hidden"` to `extraProps` and
stripping any inherited `presentation:visibility` from `baseProps` in
`drawingPageBaseProps` — exactly how the `smil:*` attributes are handled today.
**The exact attribute placement is the one wire detail to pin
rather than trust**: follow SPEC-086's golden-first process — a checked-in
canonical string in `odp/visibilityGolden.ts`, unit tests asserting the serialiser
equals it and the reader accepts it, and cross-app verification in Impress
recorded in the PR, not in CI.

### What moves to `src/lib/core` (two-phase)
`src/lib/pdf/` is already product-neutral in substance (`buildPdf.ts` imports only
its `./fontMap`/`./sfnt` siblings, and `fontMap.ts` only `../fonts/substitution`,
which imports nothing at all) but sits outside core, and core is **flat** —
`core/boundary.spec.ts` rejects any specifier not starting with `"./"` — so this
is a move to flat, prefixed names, not a `core/pdf/` subdirectory: `buildPdf.ts` →
`core/pdfBuild.ts`; `domExtract.ts` → `core/pdfExtract.ts`; `imageEncode.ts` →
`core/pdfImages.ts`; `fontMap.ts` → `core/pdfFonts.ts`; `sfnt.ts` → `core/sfnt.ts`;
`src/lib/fonts/substitution.ts` → `core/fontSubstitution.ts`; `parsePageRange` +
`RANGE_HINT` out of `src/lib/printRange.ts` → `core/pageRange.ts`, which exports
`parseNumberRange(input, total, noun = "page")` and `rangeHint(noun)` while
`printRange.ts` keeps both as thin wrappers, so every existing SPEC-034 error
string is unchanged and the deck dialog passes `"slide"`. Phase 1: move the files
and their colocated specs verbatim, leave the old paths as re-export shims, run
the suite. Phase 2: rewrite **every** importer, then delete the shims —
`src/lib/pdf/` has three outside itself (`src/views/document/usePdfExport.ts`,
`useFonts.ts` and `src/lib/fonts/subset.spec.ts`, which reaches `../pdf/sfnt`),
and `fonts/substitution` has seven, `e2e/SPEC-048.spec.ts` among them, all
repointed in the same phase or the shim deletion reds the suite. No behaviour
changes in either phase. New in core: `pdfRaster.ts` —
`rasterisePdfPage(page: PdfPageInput, widthPx: number): Promise<Uint8Array>`
paints one already-extracted page onto a canvas (text via `ctx.fillText` with
`faceCssFont(face, sizePx)` added to `pdfFonts.ts`, so the raster uses the very
face the PDF embeds; images via `drawImage`; rects as `fillRect`/`strokeRect`) and
returns PNG bytes from `canvas.toBlob` — browser-bound like `core/svgToPng.ts`,
which core already permits. Generalisations to `pdfExtract.ts`, all
default-preserving:
```ts
export interface PdfRectItem {
  xPt: number; yPt: number; widthPt: number; heightPt: number;
  fill?: [number, number, number]; stroke?: [number, number, number];  // rgb 0..1
  strokeWidthPt?: number; onTop?: true;   // 0.75 default; onTop paints after text
}
// PdfPageInput gains `rects?: PdfRectItem[]`, emitted only with a rectSelector;
// ExtractOptions gains imageSelector, canvasSelector, rectSelector.
```
`pdfBuild.ts` paints a rect as `r g b rg x y w h re f` and/or `r g b RG w w re S`
before the text run (after it, for `onTop`); a matched element contributes its
computed `backgroundColor` (skipped when fully transparent) and its uniform
border. That is the whole vector vocabulary added — enough for backgrounds, frame
fills and strokes, foreign-frame boxes, handout frames and note lines.

### The export sheet and its geometry
`DeckPrintSheet.vue` renders `planSheet(model, options)` — pure maths, no DOM —
into off-screen pages. All lengths cm; paper from `core/pageSetup.PAPER_SIZES`
(A4 21 × 29.7 default), `HANDOUT_MARGIN_CM = 1.27`, `HANDOUT_GUTTER_CM = 0.6`.
Grids, columns × rows: `slides` 1 × 1 full bleed on a page that *is* the deck's
slide size; `notes` 1 × 1 with the slide in the top half and its notes below;
`handout-1` 1 × 1, `-2` 1 × 2, `-3` 1 × 3, `-4` 2 × 2, `-6` 2 × 3, `-9` 3 × 3.
Each slot fits the slide at its native aspect, centred. With note lines on the
slot splits 55/45 horizontally — slide left, ruled lines right — which is why
6/9-up cannot carry them and 3-up always does. `resolveSlideRange` applies the
range (`all` | `current` | a parsed custom list of *slide* numbers) and then the
hidden filter unless `includeHidden`; an empty result disables every action and
shows `deck-print-range-error`.

### PDF and PNG output, progress and cancellation
`src/views/presentation/useDeckExport.ts` is the one wiring call
`PresentationView.vue` gains (`useDeckExport({ model, currentIndex, slideTotal,
reportError })`), returning `{ open, tab, layout, paper, orientation, range,
rangeError, noteLines, frameSlides, includeHidden, scale, estimate, pageCount,
rotationWarning, progress, busy, print, exportPdf, exportImages, exportSlidePng,
abort, close }`; `DeckPrintDialog.vue` is a dumb `:controller`-taking component in
the `TransitionsPanel.vue` mould (runtime-declared props, emits nothing), and
dialog choices persist for the session in the composable — `src/lib/preferences.ts`
is untouched. The run, per page: mount → `nextTick` + settle wait →
`extractPageContent` → (PDF) push a `PdfPageInput`, (PNG) `rasterisePdfPage` →
unmount → `yieldToMainThread()`, with at most `EXPORT_WINDOW_PAGES` mounted at
once — printing is the one path that must mount every page, hence
`PRINT_PAGE_LIMIT`. To make room in `PresentationView.vue` the fit-scaling block
(`stageContainer`, `stageScale`, `recomputeScale`, the `resize` listener and the
`ResizeObserver`) extracts to `src/views/presentation/useStageFit.ts` returning
`{ stageContainer, stageScale, recomputeScale }` — SPEC-094 owns that module, so
where it has landed an equivalent extraction of the same size pays instead, and
`PresentationView.size.spec.ts`'s `LINE_BUDGET` ratchets **down** from wherever
it stands, with a new ledger row. It never goes up.

### Tests
- `src/lib/presentation/printLayout.spec.ts` — slot geometry for all nine layouts,
  note-line rows, page counts, range resolution, hidden filtering.
- `src/lib/core/pdfBuild.spec.ts`, `pdfImages.spec.ts`, `pdfFonts.spec.ts`,
  `sfnt.spec.ts`, `fontSubstitution.spec.ts`, `pageRange.spec.ts` — moved
  unchanged (coverage preserved, never weakened) plus cases for `PdfRectItem`,
  `faceCssFont` and `parseNumberRange`'s noun; `pdfFonts.spec.ts` carries today's
  `domExtract` helper coverage (`fontMap.spec.ts` imports `parseCssColor` from it)
  and needs only its import repointed, and `printRange.spec.ts` keeps its strings.
  `pdfExtract.spec.ts` is **new**, not moved (`domExtract.ts` is browser-bound and
  has no colocated spec): `PdfRectItem` extraction and the selector defaults.
- `src/lib/presentation/hiddenRoundTrip.spec.ts` — hide → `writePptx`/`writeOdp` →
  re-read → `hidden` per slide; untouched-deck per-part `equalBytes`;
  byte-stability of a re-saved reopened model; `preserved["@root"]` splice
  add/remove idempotence; odp `visibilityGolden` equality and reader tolerance
  (`roundTrip.spec.ts` keeps every assertion and gains the new fixtures).
- `src/views/presentation/usePlayback.spec.ts` — hidden slides skipped by
  `advance`/`back`/`goTo`, `atEnd` with a hidden tail, deep link still shows one;
  `useDeckExport.spec.ts` — range/hidden resolution, progress ticks, abort leaving
  no download or mutation, `PRINT_PAGE_LIMIT` refusal, `MAX_PNG_EDGE_PX` clamping;
  `DeckPrintDialog.spec.ts` — the control-enablement matrix; `e2e/SPEC-104.spec.ts`.

### Specs this changes
- SPEC-034's PDF layer moves into `src/lib/core` (behaviour-identical) and gains
  the rect primitive and selector options, its own ACs and tests unchanged;
  SPEC-084's `usePlayback` gains the optional `isHidden` dep and the advance
  semantics it owns are otherwise untouched; SPEC-080's Download menu gains two
  items beside SPEC-088's export entry; SPEC-083's "Printing notes" exclusion is
  now filled by this file.
- SPEC-088's HTML export gains the same `Slide.hidden` rule (hidden slides are
  omitted unless the export opts in) — one AC to add there, not here.
- `specs/000-roadmap.md` gains this spec's phase row; nothing else changes.
