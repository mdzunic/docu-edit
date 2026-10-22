---
id: SPEC-090
title: Slide size, orientation and deck setup
status: draft
area: layout
depends_on: [SPEC-041, SPEC-064, SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-086, SPEC-089]
---
## Why
Every deck this editor creates is hard-wired to 16:9 and every deck it opens keeps
whatever size it arrived with: `slideWidthCm`/`slideHeightCm` exist and
`DEFAULT_SLIDE_WIDTH_CM`/`DEFAULT_SLIDE_HEIGHT_CM` are declared, but nothing edits them
— SPEC-080 deferred it. A deck for a 4:3 projector, an A4 handout or a portrait poster
cannot be authored here at all.

Deck setup is also where a deck's metadata belongs. Both readers already parse
`docProps/core.xml` and `meta.xml` into `PresentationModel.properties` (through
SPEC-041's product-neutral `core/opcProperties.ts` / `core/odfMeta.ts`) and both writers
throw it away. So one dialog carries both tabs.

**This file is the normative source for the deck-setup model, the slide-size catalogue,
the content-rescale maths and the deck-level write-back of `p:sldSz` /
`style:page-layout-properties` and the metadata parts.** SPEC-093 owns the placeholders
that *display* a slide number; this file owns only the number the deck counts from.

## Acceptance criteria
- [ ] `PresentationModel` gains `firstSlideNumber?: number` (1–9999, absent = 1, only a value ≠ 1 stored, set from `deck-first-slide-number`) and `deckDirty?: true` (absent = deck-level setup unchanged since open), and `slideWidthCm`/`slideHeightCm` become editable state per the Reference; the model stays plain JSON (`JSON.parse(JSON.stringify(model))` `toEqual(model)`, as `roundTrip.spec.ts` already asserts)
- [ ] `src/lib/core/pageSetup.ts` (already product-neutral — its header says "a Slides product sizes its slides … with the same types and matching/validation logic, just a different catalogue") gains the landscape-first `SLIDE_SIZES` catalogue, `SlideSizeName`, `asSlideSizeName`, `matchSlideSize`, `slideSizeFor`, `MIN_SLIDE_DIMENSION_CM = 2.54`, `MAX_SLIDE_DIMENSION_CM = 142.24`, reusing the existing `PageSize`, `Orientation`, `orientationOf`, `swapOrientation` and `PAPER_SIZES` rather than restating any of them
- [ ] A "Deck setup…" control in `DeckActions.vue` (testid `presentation-deck-setup`, emit `deck-setup`) opens one `BaseDialog`-shelled `DeckSetupDialog.vue` with two tabs — testids `deck-setup-dialog`, `deck-setup-tab-size`, `deck-setup-tab-properties`, `deck-setup-apply`, `deck-setup-cancel`; both tabs are keyboard-operable and the tabstrip copies `PropertiesDialog.vue`'s roving-tabindex pattern
- [ ] Size tab: preset selector `deck-size-preset` offering 16:9 widescreen / 4:3 standard / 16:10 / A4 paper / Letter paper / Custom (the open deck's size preselects via `matchSlideSize`, ±0.05 cm tolerance), orientation radios `deck-size-portrait` / `deck-size-landscape` (width and height swapped via `swapOrientation`), and width/height fields `deck-size-width` / `deck-size-height`
- [ ] Width and height are read through SPEC-064's `parseLengthToCm(input, preferences.measurementUnit)` and shown through `formatLength`, so `"10in"`, `"254mm"` and a bare number in the current unit all work (unit label `deck-size-unit`); a value outside 2.54–142.24 cm or unparseable shows `deck-size-error` and disables Apply, changing nothing; typing in either field switches the preset to Custom and picking a preset refills both
- [ ] The rescale decision is an explicit radio pair, `deck-size-rescale-scale` (default) and `deck-size-rescale-keep`, with the maths in the Reference: scale is uniform (`min(newW/oldW, newH/oldH)`), centres the scaled content, scales run `sizePt` and materialises SPEC-081's `DEFAULT_SIZE_PT` (18, `src/lib/presentation/text.ts`) when absent (0.5 pt steps, clamped 6–96 — those bounds are restated rather than imported, because `PRESENTATION_FONT_SIZE_MIN`/`MAX` live in `PresentationToolbar.vue`, which `lib/presentation` may not import), and leaves rotation and stroke widths alone; keep leaves all geometry untouched and shows `deck-size-overflow-warning` naming how many frames now fall outside the slide
- [ ] `src/lib/presentation/deckSetup.ts` (framework-free, importing only `lib/core` and presentation siblings) exports `applyDeckSize(model, size, mode)`, `framesOutsideSlide(model)`, `setFirstSlideNumber(model, n)` and `applyDeckProperties(model, properties)` per the Reference; each sets `model.deckDirty = true`, `applyDeckSize` in `"scale"` mode also calls SPEC-081's `markSlideEdited(slide, model.format)` per slide, and geometry rounds through the repo's `r3` convention so build → re-parse stays within ±0.01 cm
- [ ] pptx: `deckDirty` routes `writePptx` to `buildSubstitution`, which splices `p:sldSz` `cx`/`cy` (`cmToEmu`, whole EMU) into the existing `ppt/presentation.xml` and sets or removes `firstSlideNum` on the `p:presentation` start tag; `p:notesSz` is deliberately left byte-verbatim (PowerPoint sizes notes pages independently), and every other byte of `presentation.xml` — and every other part — stays as `buildSubstitution` already emits it
- [ ] odp: `deckDirty` joins `writeOdp`'s `structural` condition and extends the `styles.xml` substitution slot SPEC-089 introduces, rewriting `fo:page-width` / `fo:page-height` / `style:print-orientation` on the `style:page-layout-properties` of every `style:page-layout` a `style:master-page` references, per the Reference — notes page layouts (`presentation:notes/@style:page-layout-name`) are left untouched, and a deck with no page layout at all gains one
- [ ] Properties tab edits `properties.title`, `subject`, `keywords`, `comments`, `author`, `lastModifiedBy` (testids `deck-properties-title`, `-subject`, `-keywords`, `-comments`, `-author`, `-last-modified-by`), shows `created`/`modified` read-only (`deck-properties-created`, `-modified`) and lists preserved `extraFields` read-only (`deck-properties-extra-fields`); saving rebuilds `docProps/core.xml` via `buildCorePropertiesXml` and `meta.xml` via `buildMetaXml`, passing `customProperties` / `unsupportedCustomProperties` / `extraFields` straight back so nothing this dialog does not edit is lost
- [ ] Apply records exactly one undo entry (`"Deck setup"` / `"Deck properties"`) and one `commit()`: SPEC-080's `usePresentationUndo` `Snapshot` extends per the Reference, so undo restores size, `firstSlideNumber`, `properties` and `deckDirty` together; the stage, the filmstrip thumbnails and `stageScale` re-measure on that same commit, and the change persists write-through
- [ ] Round trip: an untouched deck still downloads byte-identically — `equalBytes` (`src/lib/core/equality.ts`) per part after `unzip`, over every part of both formats, with `roundTrip.spec.ts`'s existing `expectUntouched` and odp splice-equality assertions unchanged; a size/properties change → save → reopen reproduces `slideWidthCm`/`slideHeightCm` (±0.01 cm), `firstSlideNumber` and the six metadata fields exactly, and re-saving the reopened model is byte-stable for `presentation.xml` / `styles.xml` / `docProps/core.xml` / `meta.xml`
- [ ] `PresentationView.vue` gains exactly one `useDeckSetup({ model, currentIndex, record, commit, reportError })` call plus the dialog hookup; **this file owns the deck-header extraction** — the `.deck-header` block moves into `src/components/presentation/DeckHeader.vue`, and every sibling spec naming that extraction falls back to an equivalent one where this has already landed — and `PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 10 from whatever it stands at when this lands (550 → 540 today) with a new ledger row appended; the budget never rises, and the dialog is a dumb `:controller`-taking component like `TransitionsPanel.vue`
- [ ] e2e (`e2e/SPEC-090.spec.ts`): open `e2e/fixtures/mini-deck.pptx` → Deck setup → preset "4:3 standard" with "Scale content" → Apply → the stage's aspect ratio and slide 4's frame geometry change proportionally → undo restores both → redo → Properties tab → set author and subject → Apply → download, reopen → the preset reads 4:3 and author/subject persist; the size half repeats against `mini-deck.odp`

## Out of scope
Slide layouts, masters and the layout picker (still SPEC-080's out of scope); per-slide
size overrides (a deck has exactly one); notes-page and handout sizing (`p:notesSz` and
ODF notes page layouts are preserved, never edited); custom properties on decks, "Remove
personal information on save" and the Statistics tab (all SPEC-041, text-side only);
slide-number, date and footer *placeholders* (SPEC-093); reflow or re-autofit after a
rescale (no autofit field exists); theme setup (no theme model exists); rescaling
`GroupFrame`/`ForeignFrame` internals — their verbatim `xml` re-emits unchanged and only
the frame's own bounds move.

## Reference

### Dependencies restated
- `specs/079-slide-canvas-rendering.md` — the slide model: `PresentationModel` = `{
  kind: "presentation"; id; title; format: "pptx"|"odp"; slides: Slide[]; slideWidthCm:
  number; slideHeightCm: number; opaque?: OpaquePart[]; media?: MediaPart[];
  properties?: DocumentProperties; structureDirty?: true; odpContentEnvelope?: { prefix;
  suffix } }` — lengths cm, font sizes pt, `DEFAULT_SLIDE_WIDTH_CM = 33.867` /
  `DEFAULT_SLIDE_HEIGHT_CM = 19.05`. Each `SlideFrame` (text/image/shape/group/foreign)
  has `xCm, yCm, widthCm, heightCm, rotationDeg?`; `SlideRun.sizePt?` is effective,
  absent = SPEC-081's `DEFAULT_SIZE_PT` (18, `src/lib/presentation/text.ts`).
- `specs/080-create-presentation-and-slide-management.md` — writer paths and undo.
  `writePptx`: byte-verbatim when `!structureDirty && !anyDirty && !anyNotes`,
  `buildTextEditSubstitution` when `!structureDirty`, else `buildSubstitution` — the
  only path that rebuilds `ppt/presentation.xml` (via `spliceSldIdLst`). `writeOdp`:
  `structural = structureDirty || slides.some(s => s.dirty)`, substituting only
  `content.xml`, `META-INF/manifest.xml` and media parts — **`styles.xml` has no
  substitution slot in SPEC-080's writer; `specs/089-slide-layouts-and-masters.md`
  introduces one, and this file extends it.** `usePresentationUndo` exports `{ record(label?), undo,
  redo, canUndo, canRedo, reset }` over a private `Snapshot { slides; structureDirty?
  }`; `stripRetention` drops `opaque`/`odpContentEnvelope`/`structureDirty`. SPEC-081's
  `markSlideEdited(slide, format)` sets `slide.dirty = true` and, for odp, deletes
  `slide.sourceXml`; SPEC-082 adds `r3(v) = Math.round(v * 1000) / 1000` rounding.
- `specs/041-document-properties-and-metadata.md` — `DocumentProperties` (`title?,
  subject?, keywords?, comments?, author?, lastModifiedBy?, created?, modified?,
  customProperties?, unsupportedCustomProperties?, extraFields?`) in
  `src/lib/core/documentProperties.ts`, plus `buildCorePropertiesXml(input:
  CorePropertiesInput, modified: Date)` (`core/opcProperties.ts`) and
  `buildMetaXml(input: OdfMetaInput, modified: Date)` (`core/odfMeta.ts`).
  `OdfMetaInput` carries `customProperties`, `unsupportedCustomProperties` and
  `extraFields`; `CorePropertiesInput` carries `extraFields` (pptx custom properties
  live in the untouched `docProps/custom.xml`).
- `specs/064-vertical-ruler-and-measurement-units.md` — `preferences.measurementUnit:
  "cm"|"mm"|"in"|"pt"`, plus `parseLengthToCm( input, unit): number | null`,
  `formatLength(cm, unit)`, `cmToEmu`/`emuToCm` in `src/lib/core/units.ts`.

### Model delta (extends the model SPEC-079 defines)
```ts
export interface PresentationModel {   // …existing fields unchanged…
  /** Slide-number origin (pptx `p:presentation @firstSlideNum`). 1..9999;
   *  absent = 1, and only a value != 1 is stored. */
  firstSlideNumber?: number;
  /** Deck-level setup (size, first slide number, properties) changed since
   *  open -> the writers rebuild the deck-level parts. Absent = unchanged;
   *  only `true` is stored, exactly like `structureDirty`. */
  deckDirty?: true;
}
```
`stripRetention` deletes `deckDirty` beside `structureDirty` and keeps
`firstSlideNumber` (content, not retention). `usePresentationUndo`'s `Snapshot` gains
`deck: { slideWidthCm; slideHeightCm; firstSlideNumber?; deckDirty?: true; properties?
}`, cloned with `slides`; `applySnapshot` deletes the absent keys.

### The size catalogue (core — one place, two products)
`src/lib/core/pageSetup.ts` already owns `PageSize`, `PAPER_SIZES`, `matchPaperSize`,
`Orientation`, `orientationOf`, `swapOrientation` and `pageSizeFor`, and its header
already says a Slides product belongs here — so extend it rather than start a
presentation-side twin; no two-phase move is needed. Only the catalogue is new,
landscape-first where `PAPER_SIZES` is portrait-first:
```ts
export type SlideSizeName = "16:9" | "4:3" | "16:10" | "A4" | "Letter" | "Custom";
export const SLIDE_SIZES: Record<Exclude<SlideSizeName, "Custom">, PageSize> = {
  "16:9":  { widthCm: 33.867, heightCm: 19.05 },  // 12192000 x 6858000 EMU
  "4:3":   { widthCm: 25.4,   heightCm: 19.05 },  //  9144000 x 6858000 EMU
  "16:10": { widthCm: 25.4,   heightCm: 15.875 },
  A4:      swapOrientation(PAPER_SIZES.A4),       // 29.7 x 21
  Letter:  swapOrientation(PAPER_SIZES.Letter),   // 27.94 x 21.59
};
export function matchSlideSize(size: PageSize): SlideSizeName;  // ±0.05 cm, either orientation
export function slideSizeFor(name: SlideSizeName, o: Orientation): PageSize;
export function asSlideSizeName(value: string): SlideSizeName;  // espree pass-through
export const MIN_SLIDE_DIMENSION_CM = 2.54, MAX_SLIDE_DIMENSION_CM = 142.24;  // 1..56 in
```

### Rescaling content — exact maths
`applyDeckSize(model, { widthCm, heightCm }, mode: "scale" | "keep")`, with `oldW/oldH`
the current size and `s = Math.min(widthCm / oldW, heightCm / oldH)`:
- `"scale"` — every frame of every slide, group and foreign frames' own bounds and
  nested `GroupFrame.children` (already slide-space) included: `xCm' = r3(x * s + dx)`,
  `yCm' = r3(y * s + dy)`, `widthCm' = r3(w * s)`, `heightCm' = r3(h * s)`, with `dx =
  (widthCm - oldW * s) / 2`, `dy = (heightCm - oldH * s) / 2`.
  `GroupFrame.xml`/`ForeignFrame.xml` are **not** rewritten — they re-emit verbatim,
  placed by the frame's bounds. Every `SlideRun` gets `sizePt' = clamp(round((run.sizePt
  ?? DEFAULT_SIZE_PT) * s * 2) / 2, 6, 96)`, materialising the default when absent;
  `rotationDeg`, stroke widths, fills and text are untouched. Then
  `markSlideEdited(slide, model.format)`.
- `"keep"` — no frame touched, no slide dirtied; only the deck fields change.
  `framesOutsideSlide(model)` counts frames whose rotated bounds fall entirely outside
  `0..widthCm × 0..heightCm`, for the warning line.
- Both set the size (r3) and `deckDirty = true`; re-applying the same size is a no-op on
  geometry.

### pptx wire mapping
`ppt/presentation.xml` carries `<p:sldSz cx="…" cy="…"/>` in EMU (`cmToEmu`, whole
integers) after `<p:sldIdLst>`, then `<p:notesSz cx cy/>`. `buildSubstitution` gains
`spliceSldSz(presXml, model)` beside `spliceSldIdLst`: replace the existing element's
`cx`/`cy` in place, keeping its other attributes and the surrounding bytes; when absent,
insert one before `<p:notesSz` (else `</p:presentation>`). `firstSlideNum` is an
attribute of the `p:presentation` start tag — set when `firstSlideNumber` is present,
removed when absent, and `p:notesSz` is never recomputed. The `docProps/core.xml` part
named by the root `…/metadata/core-properties` relationship is substituted with
`buildCorePropertiesXml({ title: properties.title ?? model.title, subject, author,
lastModifiedBy, keywords, comments, created, extraFields }, modified)`;
`docProps/custom.xml` and `app.xml` stay verbatim, and `buildFromScratch` takes the same
input instead of today's `{ title: model.title }`.

### odp wire mapping
The size lives in `styles.xml`: `<office:automatic-styles>` holds `<style:page-layout
style:name="PM1"><style:page-layout-properties fo:page-width="33.867cm"
fo:page-height="19.05cm" style:print-orientation="landscape"/></style:page-layout>`, and
`<office:master-styles>` holds `<style:master-page style:page-layout-name="PM1" …/>`.
`writeOdp`'s `styles.xml` slot (SPEC-089's) is applied on `deckDirty` too: collect every
`style:page-layout-name` on a `style:master-page` (ignoring the one on its
`presentation:notes` child) and rewrite each referenced layout's `fo:page-width`,
`fo:page-height` and `style:print-orientation` (`orientationOf(size)`) — inserting an
attribute when absent, leaving every other byte alone. When `styles.xml` names no layout
the same rewrite runs over the (already rebuilt) `content.xml` automatic styles, the
fallback the reader's own `pageSize()` search mirrors; when neither has one, splice a
`<style:page-layout style:name="PMde1">` into `styles.xml` and add
`style:page-layout-name="PMde1"` to each bare `style:master-page`. `meta.xml` is
substituted with `buildMetaXml({ …the same six fields, created, extraFields,
customProperties, unsupportedCustomProperties }, modified)`, and `buildFromScratch`
takes the same input, still emitting `mimetype` first with `store: true`.

### Wire details this spec does not pin (golden-fixture-first, per SPEC-086)
Two details are **not** asserted from memory; confirm each with SPEC-086's process
before writing bytes — hand-author the candidate markup into the fixture builders,
round-trip it through the readers in a unit test, record manual PowerPoint/Impress
verification in the PR.
1. **odp slide-number origin.** `presentation:settings` carries show settings
   (`presentation:start-page` = the slide the show starts *from*), not a numbering
   origin; the plausible carrier is `style:first-page-number` on
   `style:page-layout-properties`. Until a fixture confirms it, `firstSlideNumber` is
   **pptx-only** — on odp it lives in the model, is not serialised, and the field shows
   a "not saved in .odp" hint.
2. **`p:notesSz` on a size change.** The never-recomputed rule assumes PowerPoint sizes
   notes pages independently; if a fixture shows otherwise, revisit that rule alone —
   preserving source bytes is safe either way.

### UI and view wiring
`src/views/presentation/useDeckSetup.ts` (`DeckSetupDeps { model, currentIndex, record,
commit, reportError }`) returns `{ open, isOpen, close, activeTab, setTab, preset,
setPreset, orientation, setOrientation, widthText, heightText, setWidth, setHeight,
unit, sizeError, rescaleMode, setRescaleMode, overflowCount, firstSlideNumber,
setFirstSlideNumber, propertiesDraft, setProperty, canApply, apply }` — `record(…)`,
mutate, `commit()`; drafts are local, so Cancel changes nothing.
`src/components/presentation/DeckSetupDialog.vue` takes exactly `{ controller: { type:
Object, required: true } }`, emits nothing, reads `c.preset.value` etc. and uses runtime
`defineProps` with no TS annotations in the SFC (espree) — typed transforms go in a
sibling `deckSetupFields.ts`.

### Tests
- `src/lib/core/pageSetup.spec.ts` — `matchSlideSize` tolerance in both orientations,
  `slideSizeFor`, the clamps; the paper-size expectations stay untouched.
- `src/lib/presentation/deckSetup.spec.ts` (new) — scale maths (centring, both scale
  directions, group children, foreign bounds, font clamp at 6 and 96, absent-`sizePt`),
  keep mode, `framesOutsideSlide`, `deckDirty`/`markSlideEdited`, first slide number.
- `pptx/writer.spec.ts` / `odp/writer.spec.ts` (both new) — the `p:sldSz` and
  `firstSlideNum` splices, `p:notesSz` untouched, the page-layout rewrite with its
  notes-layout exclusion and no-layout fallback, both metadata parts.
- `roundTrip.spec.ts` — a new `describe` per the round-trip AC: `equalBytes` per part
  for the untouched deck, field-level `toEqual` for the changed one; its existing
  `expectUntouched`, `STRUCTURAL_PPTX_PARTS` and splice-equality assertions re-run
  unchanged (preserved, never weakened).
- `useDeckSetup.spec.ts` — preset/orientation/custom interplay, unit parsing and
  rejection, one undo entry per Apply, Cancel; `PresentationView.size.spec.ts` for the
  ratcheted-down budget; `e2e/SPEC-090.spec.ts` for the AC scenario over the checked-in
  `mini-deck.pptx`/`mini-deck.odp` (no new fixture, so
  `scripts/build-presentation-fixtures.ts` is unchanged).

### Specs this changes
- SPEC-080's out-of-scope "slide size setup" is superseded here; its writer path
  selection gains `deckDirty` as a third dirty axis (a second deck-level part in
  SPEC-089's `styles.xml` slot) and its `Snapshot` gains `deck`.
- SPEC-079's `slideWidthCm`/`slideHeightCm` become editable rather than read-at-open;
  `DEFAULT_SLIDE_*` stay the from-scratch values and now also name
  `SLIDE_SIZES["16:9"]`. SPEC-041's `DocumentProperties` gains a second consumer; the
  text-side dialog and its `properties-*` testids are untouched.
- `specs/000-roadmap.md` gains this spec's Phase-13 row and drops "Slide size setup
  dialog" from the backlog; nothing else changes.
