# src/lib/core — the product-neutral layer

Everything in this directory is independent of the word-processor's document
model, so a second product built on the same engine (the planned Slides
editor) can import it without dragging in `documentModel.ts`, `commands.ts`
or any `.docx`/`.odt` reader. The contract, enforced by `boundary.spec.ts`:

> **A core module may import only sibling core modules.** No `../` imports,
> no `vue`, no path aliases. DOM and canvas APIs are fine (several modules
> measure or paint); the document model is not.

What lives here, by area:

- **Units and geometry** — `units.ts` (cm/in/pt/px/twips/EMU/half-points,
  display units), `shapeGeometry.ts` (polygons, rotation, hit testing),
  `pageSetup.ts` (PageSize/Margins, the paper catalogue, orientation).
- **Drawing objects** — `shapes.ts`, `shapeCanvas.ts`, `shapeIcons.ts`,
  `charts.ts`, `chartCanvas.ts`, `chartIcons.ts`. The `*Visual` interfaces
  are the painters' contract: the document model's `Shape`/`Chart` blocks
  satisfy them structurally, and a Slides model can too.
- **Tables** — `tables.ts`: border/alignment/width-mode types, grid limits,
  cell-range arithmetic, the merged-cell owner grid (`computeOwnerGrid`,
  structural over any row/cell shape), DOM resize previews. Cell *content*
  stays product-typed (a document cell holds paragraphs) and out of core.
- **Lists** — `lists.ts` (types, marker styles, ordinal formatting, level
  planning), `listDefinitions.ts` (presets, marker composition, instance
  counters). Deriving markers by walking a document stays in `lib/lists.ts`.
- **Media** — `media.ts` (MediaPart, acceptance rules, fit math, the
  content-type ↔ file-extension mapping both package formats share),
  `imageDimensions.ts`, `svgToPng.ts`.
- **Assembly** — `sheetLayout.ts` (SPEC-108's sheet: paper, margins, fit,
  align, dpi) and `assemblyLayout.ts` (SPEC-109's five layouts: where each
  tray item lands on that sheet, in centimetres, from its pixel size alone).
- **Text** — `grapheme.ts`, `marks.ts`, `autocorrect.ts`,
  `specialCharacters.ts`, `language.ts`, `fonts.ts` (the standard catalogue).
- **Layout measurement** — `pageMetrics.ts` (LineBox/BlockMetrics and the
  DOM measurers; the pagination engine itself is document-shaped and stays
  in `lib/pagination.ts`).
- **PDF** (SPEC-034's writer, moved here by SPEC-104 so both products can
  export) — `pdfBuild.ts` (the PDF 1.7 writer), `pdfExtract.ts` (a realised
  DOM page → `PdfPageInput`, with configurable image/canvas/rect selectors),
  `pdfImages.ts` (image XObject encoding), `pdfFonts.ts` (CSS family → bundled
  face), `sfnt.ts`, `svgPath.ts`, `fontSubstitution.ts` (SPEC-048's metric
  twins), `pageRange.ts` (page/slide range parsing) and `pdfRaster.ts` (one
  extracted page painted to PNG bytes).
- **Package plumbing** — `zip.ts`, `zipWriter.ts`, `xml.ts` (parsing, OPC
  relationships, canonicalization, the XML escape family — the canonical
  text/attribute pair plus the writers' full-entity `escapeXml`),
  `equality.ts` (equalBytes). Shared by OOXML and ODF packages of any
  product (`.pptx`/`.odp` included).
- **App plumbing** — `idb.ts` (generic IndexedDB layer), `idleScheduler.ts`,
  `autosaveScheduler.ts`, `versionCadenceScheduler.ts`, `editLock.ts`,
  `platform.ts`, `colors.ts`, `commentIdentity.ts`, `menuBar.ts`,
  `fileReader.ts` (File → ArrayBuffer), `arrayWindow.ts`
  (reference-equality array windowing for diff engines), `motion.ts`
  (`prefersReducedMotion`/`motionBaseMs` — the `--motion-base` token reads),
  `windowChannel.ts` (SPEC-100's payload-agnostic same-origin channel between
  two windows: BroadcastChannel, else a `localStorage` `storage`-event
  fallback, else none — messages, where `editLock.ts` is exclusion).

Known candidates deliberately not (yet) extracted: `textMetrics.ts`'s wrap
engine (typed against `Paragraph`/`RunFormatting`), the pagination engine,
and a generic `Table<TCell>` — each needs a real design pass, not a move.

Comments referencing SPEC-NNN acceptance criteria travel with the code they
document; the spec set remains the behavioural source of truth.
