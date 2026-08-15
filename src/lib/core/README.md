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
  cell-range arithmetic, DOM resize previews. Cell *content* stays
  product-typed (a document cell holds paragraphs) and out of core.
- **Lists** — `lists.ts` (types, marker styles, ordinal formatting, level
  planning), `listDefinitions.ts` (presets, marker composition, instance
  counters). Deriving markers by walking a document stays in `lib/lists.ts`.
- **Media** — `media.ts` (MediaPart, acceptance rules, fit math),
  `imageDimensions.ts`, `svgToPng.ts`.
- **Text** — `grapheme.ts`, `marks.ts`, `autocorrect.ts`,
  `specialCharacters.ts`, `language.ts`, `fonts.ts` (the standard catalogue).
- **Layout measurement** — `pageMetrics.ts` (LineBox/BlockMetrics and the
  DOM measurers; the pagination engine itself is document-shaped and stays
  in `lib/pagination.ts`).
- **Package plumbing** — `zip.ts`, `zipWriter.ts`, `xml.ts` (parsing, OPC
  relationships, canonicalization), `equality.ts` (equalBytes). Shared by
  OOXML and ODF packages of any product (`.pptx`/`.odp` included).
- **App plumbing** — `idb.ts` (generic IndexedDB layer), `idleScheduler.ts`,
  `autosaveScheduler.ts`, `versionCadenceScheduler.ts`, `editLock.ts`,
  `platform.ts`, `colors.ts`, `commentIdentity.ts`, `menuBar.ts`.

Known candidates deliberately not (yet) extracted: `textMetrics.ts`'s wrap
engine (typed against `Paragraph`/`RunFormatting`), the pagination engine,
and a generic `Table<TCell>` — each needs a real design pass, not a move.

Comments referencing SPEC-NNN acceptance criteria travel with the code they
document; the spec set remains the behavioural source of truth.
