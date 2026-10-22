---
id: SPEC-080
title: Create a presentation, and slide management
status: draft
area: editing
depends_on: [SPEC-079]
---
## Why
After SPEC-079 the app can *show* a deck but not start or restructure one:
there is no "New presentation", and the filmstrip is a viewer. This spec
adds the first mutations — create a blank deck, add, duplicate, delete and
reorder slides — and with them the machinery every later editing spec
rides on: from-scratch package writers for both formats, the
part-substitution save path that rebuilds only what changed while
re-emitting everything else verbatim, the presentation undo stack, and
write-through persistence. The word processor walked this exact arc in
SPEC-002/006/007; this is the presentation product's version, sized to
slides. **This file is the normative source for the presentation writers'
rebuild rules and the presentation undo stack.**

## Acceptance criteria
- [ ] The landing page offers "New presentation" beside "New document" (HomeView wiring, product-neutral library components untouched); activating it stores a fresh blank deck via `storePresentation` and navigates to `/slides/:id`
- [ ] A blank deck is 16:9 (33.867×19.05cm), `format: "pptx"`, one slide with a centred title placeholder and a subtitle placeholder (geometry per the Reference), title `DEFAULT_PRESENTATION_TITLE`, no `opaque` — and it renders on the SPEC-079 stage exactly as its saved file re-opens
- [ ] `lib/documentKinds.ts`'s "always a text document" note for the `"new"` format is resolved: blank presentations carry `"pptx"` from birth, `kindOfFormat` needs no new format, and the library badge shows PPTX
- [ ] Downloading a deck that has no retained package produces a valid from-scratch file in its format; the from-scratch `.pptx` re-opens in this app with the same slides, placeholders and geometry (`readPresentationBytes` round trip), and the from-scratch `.odp` does too, `mimetype` first and STOREd
- [ ] The deck header offers "Download" (own format) and "Download a copy as…" the other format: a `.pptx`-born deck saves as `.odp` and back, converting through the model (retained parts do not leak across formats — a cross-format save never emits source-format parts)
- [ ] The filmstrip gains slide management: an "Add slide" button plus a per-slide context menu (right-click and a ⋯ affordance) with New slide, Duplicate slide, Delete slide, Move up, Move down — all also keyboard-reachable (menu key/Shift+F10 on a focused thumbnail opens the menu)
- [ ] New slide inserts after the current slide and becomes current: title + body placeholders per the Reference geometry, styled by the deck's defaults; on a deck opened from a file, the new slide's `.pptx` part references the current slide's layout (falling back to the deck's first layout), and the new `.odp` page uses the current page's master
- [ ] Duplicate slide inserts an exact copy after its source and becomes current: same frames (fresh frame/slide ids), same preserved XML and (pptx) a copied part + copied `.rels` under a fresh part name — duplicating then downloading yields a package whose duplicated slide part bytes equal the source slide's, rels rewritten to the fresh names only where they must be
- [ ] Delete slide removes it (allowed down to zero slides; the stage then shows an empty state with an "Add slide" action); the current slide moves to the nearest remaining neighbour
- [ ] Move up/down swaps positions and keeps the moved slide current; dragging a thumbnail reorders with a visible drop indicator, and dropping outside the filmstrip cancels
- [ ] Every operation above is undoable and redoable: Cmd/Ctrl+Z and Shift+Cmd/Ctrl+Z (plus toolbar buttons) walk a presentation undo stack capped at 100 entries; undo of a delete restores the slide's frames, preserved XML, part linkage and position; the stack is per-open-view (not persisted)
- [ ] Every operation persists write-through: reload after any operation (unit-level: store round trip) shows the mutated deck; `touchDocument` keeps the library's recency in step
- [ ] Saving a deck opened from a file after only structural operations rebuilds only the structural parts — pptx: `ppt/presentation.xml`, its `.rels`, `[Content_Types].xml`; odp: `content.xml` (+ `META-INF/manifest.xml` only when parts were added/removed) — every slide part / `draw:page` body re-emitted byte-verbatim (`equalBytes` per part against the source for untouched slides; splice-equality for odp pages per the Reference)
- [ ] Saving a deck opened from a file with **no** operations performed still re-emits every part byte-identical — the prep milestone's round-trip suite passes unchanged (the writers rebuild nothing unless `structureDirty` or a slide is dirty)
- [ ] Reordering slides in a `.pptx` with unmodelled deck-level parts (notesMasters, custom tags, thumbnails) keeps those parts and their relationships byte-verbatim; only `p:sldIdLst` order, the rebuilt rels and Content-Types change
- [ ] e2e (`e2e/SPEC-080.spec.ts`): create a blank presentation from the landing page → add a slide → duplicate it → reorder by keyboard → delete one → undo the delete → download and re-open the file → the filmstrip shows the expected slide count and order (by slide titles)

## Out of scope
Text or object editing (SPEC-081/082) — a blank deck's placeholders render
their empty state but cannot be typed into yet. Choosing a theme or layout
at creation; a layout picker for new slides (new slides are title+body
only). Slide size setup (16:9 fixed for new decks; opened decks keep
theirs). Copy/paste of slides across decks or apps. Autosave/drafts (the
store's draft machinery stays text-only, per the prep milestone's
`saveDraft` guard). Renaming from the deck view beyond what the library
already offers. Persisting the undo stack.

## Reference

### Dependencies restated
- SPEC-079's slide model: `Slide { id, frames, background?, partName?,
  sourceXml?, preserved? }`, frames with cm geometry, `slideWidthCm`/
  `slideHeightCm`, minted ids never serialised. Renderer:
  `SlideStage.vue` + `SlideFilmstrip.vue`; view wiring in
  `PresentationView.vue`.
- Store: `storePresentation(model)`, `getPresentation(id)`,
  `touchDocument(id)`, `duplicateDocument` conventions, write-through
  persistence and the media split; `LibraryEntry.kind` already routes
  presentation rows. Round-trip equality: part-level `equalBytes`
  (`core/equality.ts`).
- Writers today (`presentation/pptx/writer.ts` `writePptx`,
  `presentation/odp/writer.ts` `writeOdp`): re-emit `model.opaque`
  verbatim; `.odp` `mimetype` first + STOREd via `zipWriter`'s
  `ZipEntryInput.store`. This spec extends both.
- XML plumbing: `core/xml.ts` (`parseRelationships`, `relsPathFor`,
  `resolvePartPath`, serialisation + escaping helpers), `core/zip.ts` /
  `core/zipWriter.ts`, `core/units.ts` (`cmToEmu`).

### Model additions (in `presentation/model.ts`)
```ts
export interface Slide {
  // ...SPEC-079 fields...
  /** Set by any command that changes this slide's own content or
   *  properties (SPEC-081+ set it; this spec's new/duplicated slides are
   *  born without partName/sourceXml instead). Dirty slides are rebuilt
   *  by the writers; clean ones re-emit verbatim. */
  dirty?: true;
}
export interface PresentationModel {
  // ...
  /** Set when slides were added/removed/reordered since open. Absent on
   *  a freshly opened deck; from-scratch decks are always fully built. */
  structureDirty?: true;
  /** odp only: content.xml split at read — everything before the first
   *  draw:page and everything after the last (the office:presentation
   *  close onwards). Absent for pptx and from-scratch decks. */
  odpContentEnvelope?: { prefix: string; suffix: string };
}
```
`parseOdp` records each page's `sourceXml` (SPEC-079) plus this envelope;
splice-equality below means `prefix + pages.join("") + suffix` equals the
source `content.xml` byte-for-byte for an untouched deck — asserted once
in the reader suite so the writer's verbatim path is trustworthy.

### Commands (`src/lib/presentation/slides.ts`, framework-free)
`addSlide(model, afterIndex)`, `duplicateSlide(model, index)`,
`deleteSlide(model, index)`, `moveSlide(model, from, to)` — each mutates
`model.slides`, sets `structureDirty`, and returns the index to select.
New-slide placeholder geometry (fractions of slide size): title x 8%,
y 8%, w 84%, h 16%; body x 8%, y 30%, w 84%, h 60% — for the blank deck's
first slide, title y 32% h 18% and subtitle y 52% h 12%, both centred
text. `createBlankPresentation(): PresentationModel` lives beside them.
Placeholder roles map to `phType`: title→`title`, subtitle→`subTitle`,
body→`body` with `phIdx: 1`.

### The undo stack (`src/views/document/`-style composable, presentation-side)
`src/views/presentation/usePresentationUndo.ts` — the first module of a
new `src/views/presentation/` composable directory mirroring
`src/views/document/` (PresentationView is a wiring hub under the same
philosophy; add `PresentationView.size.spec.ts` ratcheted at its
post-SPEC-080 line count, only ever downward). API:
`{ record(label), undo(), redo(), canUndo, canRedo, reset() }`. `record`
snapshots `{ slides, structureDirty }` via `structuredClone` **before** a
command runs (opaque/media excluded — commands in this spec never touch
them); undo swaps the live model's fields with the snapshot and persists.
Cap 100, oldest dropped. SPEC-081+ push their edits through the same
`record`.

### pptx writer rebuild rules (`writePptx`)
With no `opaque`: build the full minimal package —
`[Content_Types].xml`, `_rels/.rels`, `docProps/core.xml` (title +
timestamps via the SPEC-041 machinery in `core/opcProperties.ts`; reuse
the writer convention that stamps `modified` per save — and mind the
clock-coupled-test rule: tests never hard-code "today"),
`ppt/presentation.xml` (`p:sldSz` from the model's cm via `cmToEmu`),
`ppt/_rels/presentation.xml.rels`, one `ppt/slideMasters/slideMaster1.xml`
+ rels + `ppt/theme/theme1.xml` (fixed minimal Office theme: Calibri
minor/major, standard accent scheme), one
`ppt/slideLayouts/slideLayout1.xml` (+ rels → master), and per slide
`ppt/slides/slideN.xml` (+ rels → layout1) serialised by the frame
serialiser SPEC-081 owns — until SPEC-081 lands, from-scratch slides
contain only empty placeholders, which this writer can already emit
(`p:sp` with `p:ph` + `a:xfrm` + empty `p:txBody`).

With `opaque` (the substitution path): start from `opaque` in order.
- Slides deleted from the model: skip their part + their `.rels`
  (relationship targets they alone referenced — their notesSlide — are
  also skipped; shared targets like layouts/media stay).
- Duplicated slides: copy the source part bytes + `.rels` under the next
  free `ppt/slides/slideN.xml` name (N = max existing + 1…), appended
  after the last slide part.
- New slides: a fresh part serialised as in the from-scratch path but
  referencing the resolved layout part by relationship.
- `ppt/presentation.xml`: rebuilt only when `structureDirty` (or any
  slide dirty, once SPEC-081 exists): the source part's XML with its
  `p:sldIdLst` replaced — slide ids re-minted 256… in model order,
  `r:id`s pointing into the rebuilt presentation `.rels` (existing
  non-slide relationships preserved with their original ids; slide
  relationships renumbered after the highest surviving id). Everything
  else in the part — `p:sldMasterIdLst`, `p:notesMasterIdLst`, sizes,
  `p:defaultTextStyle` — is carried by string splice of the `sldIdLst`
  element, not re-serialised.
- `[Content_Types].xml`: add/remove slide `Override` entries by string
  splice before `</Types>`; every other declaration byte-verbatim.

### odp writer rebuild rules (`writeOdp`)
With no `opaque`: minimal package — `mimetype` (first, STOREd),
`META-INF/manifest.xml`, `meta.xml` (via `core/odfMeta.ts`),
`styles.xml` (one master page `Default` with the drawing-page defaults),
`content.xml` with generated automatic styles (SPEC-081's serialiser; this
spec emits empty placeholder frames the same way as pptx). With `opaque`:
re-emit all parts verbatim except `content.xml`, rebuilt as
`envelope.prefix + pages + envelope.suffix` where each untouched page
contributes `sourceXml` verbatim and new/duplicated pages contribute
generated markup (duplicate = the source page's `sourceXml` with
`draw:name` uniqued; ODF ids inside are left as-is — `draw:page` content
carries no package-wide-unique ids the formats we write depend on).
`META-INF/manifest.xml` is spliced only when parts were added/removed.

### Cross-format "Download a copy as…"
`writeOdp` on a pptx-born model (or the reverse) must see a model with no
usable retained package: the download path passes `stripRetention(model)`
(a pure helper returning a copy without `opaque`/`odpContentEnvelope`/
`partName`/`sourceXml`/`preserved`, with every slide implicitly rebuilt).
Fidelity is whatever the model holds — placeholders, and post-081 the
full frame set; foreign frames cannot cross formats and are dropped from
the copy, counted into a one-line notice on the download ("N objects
couldn't be converted"), matching SPEC-050's conversion-notice tone.
The stored deck itself keeps its format and retention — "a copy as" never
mutates the open model.

### UI
Deck header: the existing Download button becomes a split
"Download ▾" (own format primary; menu: the other format's "Download a
copy as…"; SPEC-088 later appends its Export entry to this menu). Undo/redo buttons + shortcuts live on the
view (shortcut handling ignores events from inputs/contenteditable — the
SPEC-021 convention). Filmstrip context menu uses the design-system menu
primitives (`core/menuBar.ts` types + the shared menu components used by
`DocumentMenuBar`); drag-reorder uses pointer events with
`aria-grabbed`-free simple announcements ("Slide 3 moved to position 1",
`aria-live`), and each thumbnail keeps its SPEC-079 option semantics.
Testids: `presentation-add-slide`, `presentation-slide-menu`,
`presentation-download`, `presentation-download-menu`,
`presentation-undo`, `presentation-redo`.

### Tests
- `presentation/slides.spec.ts`: every command, selection return values,
  dirty/structure flags, blank-deck shape, `stripRetention`.
- `presentation/roundTrip.spec.ts` additions: from-scratch pptx/odp
  build → parse round trip (`equalStructural` is a text-model relation;
  presentations assert deep-equality of the parsed model minus minted
  ids — write a `strippedForCompare` helper mirroring SPEC-078's
  "compare resolved facts, not minted values" rule); substitution-path
  part-level `equalBytes` per AC; no-op save byte-identity; odp splice
  equality.
- `views/PresentationView.spec.ts` additions: undo/redo wiring, empty
  deck state, context menu keyboard path.
- `HomeView` spec: the create action routes and stores (mirroring the
  existing "New document" test).
- e2e per AC. The download-and-reopen step uses the existing e2e
  download-capture helper from the presentations suite.

### Specs this changes
- `docs/presentations-architecture.md`'s "no editing yet" note and
  `documentKinds.ts`'s `"new"`-format comment are both updated by this
  spec's implementation.
- SPEC-008's retention promise now applies at part granularity for
  presentations: unmodelled parts and untouched slides survive verbatim;
  this spec's substitution rules are the presentation-side codification.
