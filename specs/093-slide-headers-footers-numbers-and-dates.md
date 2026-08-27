---
id: SPEC-093
title: Slide footers, slide numbers and dates
status: draft
area: formatting
depends_on: [SPEC-053, SPEC-079, SPEC-080, SPEC-081, SPEC-083, SPEC-089, SPEC-090]
---
## Why
Every deck that leaves a building carries the same three strings along its bottom
edge: a footer ("Confidential", the client's name), the date, and the slide number.
Today they are frozen text — SPEC-079 reads a slide-number or date placeholder as an
ordinary `TextFrame` holding "3" or "14 May 2026", and its out of scope says exactly
that: *fields beyond slide-number/date placeholders rendering their literal cached
text*. Nothing can turn a footer on, off, or edit it; and because SPEC-080 can
insert, delete, duplicate and reorder slides, the frozen numbers are wrong the moment
a deck is edited — on screen and in the file, since SPEC-081's `runXml` emits `<a:r>`
and would rewrite a live `a:fld` into a dead literal run.

This spec models the three fields properly: deck-level values, per-slide visibility, a
Header & Footer dialog with Apply / Apply to all, and a value recomputed at every
render and every save. It reuses the word processor's date machinery (SPEC-053's
closed catalogue and renderer) rather than growing a second one, and takes the
placeholder geometry from SPEC-089 instead of guessing where a footer sits. **This
file is the normative source for the slide field model (`SlideFieldPlaceholder`,
`DeckHeaderFooter`), the `a:fld`/declaration wire mappings, and the rule that a cached
field value is rewritten on save.**

## Acceptance criteria
- [ ] `PresentationModel` gains `headerFooter?: DeckHeaderFooter` and `headerFooterDirty?: true` (the slide-number origin is SPEC-090's `firstSlideNumber?: number`, read here and never written); `Slide` gains `fields?: SlideFieldPlaceholder[]` (at most one entry per kind `footer`/`slideNumber`/`date`; presence **is** visibility; absent/`[]` = the slide shows none) — full shapes, defaults and absent-vs-present rules per the Reference, all in `src/lib/presentation/model.ts`
- [ ] Both readers populate them and **stop** emitting these placeholders as ordinary frames (supersedes SPEC-079): pptx from each slide's `p:sp` whose `p:ph@type` is `ftr`/`sldNum`/`dt` (its source `p:sp` kept in `Slide.preserved["p:ph:ftr"|"p:ph:sldNum"|"p:ph:dt"]`), plus `p:presentation@showSpecialPlsOnTitleSld` and the master/layout `p:hf`; odp from `presentation:footer-decl`/`presentation:date-time-decl` in `office:presentation`, each page's `presentation:use-footer-name`/`use-date-time-name`, and the drawing-page style's `presentation:display-footer`/`display-page-number`/`display-date-time`
- [ ] Geometry and look resolve once at read into the entry (slide's own `a:xfrm`/`svg:x…` → the readers' existing placeholder maps: pptx `SlideContext.layoutPh` then `masterPh` at `phKey("ftr"|"sldNum"|"dt", idx)`, odp `OdpContext.masterPh` at `"footer"`/`"page-number"`/`"date-time"` → the `DEFAULT_FIELD_BAND` fallback in the Reference), so rendering never re-walks the layout chain and no geometry here waits on SPEC-089
- [ ] `src/lib/presentation/headerFooter.ts` (framework-free, imports only presentation siblings and `lib/core`) owns the resolver `fieldText(model, slideIndex, kind, instant): string`, `slideNumberFor(model, slideIndex): number` = `(model.firstSlideNumber ?? 1) + slideIndex` (SPEC-090's field), `defaultFieldRect`, `setFieldVisibility`, `applyHeaderFooter(model, values, scope)` and `SLIDE_FIELD_KINDS`; it returns `""` for every kind on a title slide when `showOnTitleSlide === false` (title = SPEC-089's layout type `title`, or index 0 alone when that model is absent), and that same suppression applies at save; every surface (stage, filmstrip, `PresentationPlayView`, build preview, SPEC-088 export) renders through it, and the stage element carries `data-testid="slide-field"` + `data-field-kind="footer|slide-number|date"`
- [ ] Slide numbers are live: `addSlide`/`deleteSlide`/`moveSlide`/`duplicateSlide` (SPEC-080) change no field data, yet every visible number re-renders immediately from the slide's new index, and the filmstrip thumbnails agree
- [ ] A "Header & footer" toolbar button (`slides-header-footer`) opens a modal `BaseDialog` (`src/components/BaseDialog.vue`, testid `header-footer-dialog`, title "Header and footer") with: footer checkbox + text input, date checkbox with "Update automatically" / "Fixed" radios, a format `<select>` of the six catalogue date patterns previewed against today's instant, a fixed-text input, a slide-number checkbox, "Don't show on title slide", and Cancel / Apply / Apply to all — testids `footer-show`, `footer-text`, `footer-date-show`, `footer-date-auto`, `footer-date-fixed`, `footer-date-fixed-text`, `footer-date-format`, `footer-slide-number`, `footer-hide-on-title`, `footer-cancel`, `footer-apply`, `footer-apply-all`; every control native and keyboard-operable
- [ ] Apply writes the deck values and this slide's visibility; Apply to all writes the deck values and applies the same visibility to every slide, dropping per-slide text overrides; both `record()` before mutating and `commit()` once (undo labels "Header and footer" / "Apply header and footer to all slides"), mark every touched slide via `markSlideEdited(slide, model.format)`, set `model.headerFooterDirty = true`, and persist write-through — SPEC-080's `usePresentationUndo` `Snapshot` gains `headerFooter?` and `headerFooterDirty?` per the Reference, so that one undo reverts footer text, date mode and format together with the per-slide visibility rather than stranding the deck values; Cancel changes nothing
- [ ] Dates: the six supported patterns are `d MMMM yyyy`, `d MMM yyyy`, `dd/MM/yyyy`, `MM/dd/yyyy`, `yyyy-MM-dd`, `EEEE, d MMMM yyyy` (default `d MMMM yyyy`), rendered by `renderDateTime(format, instant, DEFAULT_LANGUAGE)` after the two-phase move of SPEC-053's catalogue into `src/lib/core/dateTimeFormats.ts`; the save-time instant is a new `now?: Date` option on `WritePptxOptions`/`WriteOdpOptions` (default core `now()`), added beside their existing `modified?: Date` so a round-trip test pins it; a source date field whose wire format is outside the mapping table keeps `wireType`/`wireStyleName` verbatim, renders its cached text (pptx) or the default pattern (odp — ODF caches nothing), never recomputes, and the dialog disables the format select behind a note (`footer-date-format-preserved`)
- [ ] pptx save rewrites the cached value: each rebuilt slide re-emits its preserved `p:sp` with only the `a:t` inside its `a:fld`/first run replaced by the freshly computed string (`<a:fld id="{guid}" type="slidenum">` for numbers, the mapped `datetime*` type for an auto date, a plain `a:r` for footer text and for a fixed date); a slide enabling a field for the first time gets the canonical `p:sp` from the Reference; `showSpecialPlsOnTitleSld` is a single-attribute splice on the `p:presentation` start tag (the SPEC-083 `p:notesMasterIdLst` precedent — the rest of `presentation.xml` stays byte-identical) applied only when `headerFooterDirty`
- [ ] odp save regenerates the `presentation:footer-decl`/`presentation:date-time-decl` block inside `odpContentEnvelope.prefix` **only** when `headerFooterDirty` (otherwise the prefix is emitted verbatim and the `prefix + pages + suffix` byte-recombination invariant is untouched), minting `ftr1…`/`dtd1…` in first-use order, reusing the source declaration name when a value is unchanged, referencing them from each rebuilt page's `presentation:use-*-name`, and folding `presentation:display-*` into the drawing-page style through `StyleRegistry.drawingPageStyle` — field placeholders are never serialised into `draw:page`
- [ ] Round trip, both formats: a deck with all three fields on, a custom footer, a fixed date on one slide and an auto date elsewhere → save → reopen → `equalStructural` (`src/lib/equality.ts`) over the reopened vs expected model holds for `headerFooter`, `firstSlideNumber` and every `Slide.fields` entry; every part not carrying a field placeholder stays `equalBytes`-identical (`src/lib/core/equality.ts`, part-level after `unzip`), a no-edit download is byte-identical, and the **one sanctioned delta** is pptx's locale-dependent `datetime*` type (Reference: "Date formats") which the test pins explicitly (`equalSemantic` is `DocumentModel`-typed and is not used here)
- [ ] `stripRetention` also clears `headerFooterDirty`; a cross-format "Download a copy as…" and a from-scratch deck emit the fields from the model — pptx: the placeholder `p:sp` per slide plus `<p:hf sldNum="1" hdr="0" ftr="1" dt="1"/>` on the generated master; odp: the declarations, the master-page placeholders and each page's display flags — and reopening the copy reproduces the same `Slide.fields`
- [ ] Notes pages keep theirs (SPEC-083 delta): the notes-slide rebuild no longer drops the source part's non-body placeholders — `p:ph` of type `hdr`/`ftr`/`dt`/`sldNum` are re-emitted verbatim with their cached `a:t` refreshed the same way; the dialog does not edit notes/handout fields and `presentation:header-decl` is preserved untouched in both formats
- [ ] Fixtures: `src/lib/presentation/fixtures/build.ts` gains, under a new banner and without touching the four mini-deck slides, `buildPptxHeaderFooterFixtureBytes()`/`buildOdpHeaderFooterFixtureBytes()` + `EXPECTED_HEADER_FOOTER`, a title slide with fields suppressed, a slide with a per-slide footer override, and a slide with a deliberately out-of-subset date (`type="datetime13"` / an Impress `D1` data style) pinning the preserve-don't-parse path; `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec093-headers-footers.pptx` and `.odp`
- [ ] All feature logic lives in `src/views/presentation/useHeaderFooter.ts` (deps `{ model, currentSlide, currentIndex, record, commit }`) and `src/components/presentation/HeaderFooterDialog.vue` (props `{ controller, slideNumber }`, emits nothing, reads `c.open.value` like `TransitionsPanel.vue`); `src/views/PresentationView.vue` gains only `const headerFooter = useHeaderFooter({ model: presentation, currentSlide, currentIndex, record: slideUndo.record, commit });`, the toolbar prop/handler, and one `<HeaderFooterDialog v-if="headerFooter.open.value" :controller="headerFooter" />`; the wiring is paid for by extracting the `.deck-header` block into `src/components/presentation/DeckHeader.vue` — SPEC-090 owns that extraction, so where it has already landed an equivalent extraction of the same size pays instead — and `src/views/PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 10 from whatever it stands at when this lands (550 → 540 today) with a ledger row appended (the budget never rises)
- [ ] e2e (`e2e/SPEC-093.spec.ts`): open `spec093-headers-footers.pptx` → footer, date and "2" visible on slide 2 → Header & footer → set footer text "Confidential", tick slide number and auto date, Apply to all → slide 1 still shows none (the fixture deck has "Don't show on title slide" set) → duplicate slide 2 → the copy reads "3" and the following slide reads "4" without any further input → download, reopen → the same three values render, and the date equals `renderDateTime("d MMMM yyyy", new Date(), DEFAULT_LANGUAGE)` computed in the test (SPEC-053's e2e precedent, so no clock is pinned)

## Out of scope
Headers on slides (PresentationML has none; `p:ph type="hdr"` and
`presentation:header-decl` exist only for notes/handouts and are preserved, not
edited), the Notes and Handouts tab, handout masters, per-slide date formats
(deck-level; only the *text* may diverge per slide), "Start numbering at" and
slide-size setup (SPEC-090 owns deck setup and `firstSlideNumber`; this file only
reads it), moving/resizing a field placeholder on the stage or editing its
font or colour, arbitrary fields in ordinary text boxes (a `text:date`/`a:fld` outside a
field placeholder keeps rendering its literal cached text per SPEC-079), hidden-slide
numbering, time-of-day formats, non-Gregorian calendars, and animating a field
placeholder (SPEC-086's `isAnimatableKind` is unchanged; these are not frames).

## Reference

### Dependencies restated
- `specs/079-slide-canvas-rendering.md` — the slide model: `Slide { id, frames,
  background?, notes?, dirty?, notesDirty?, transition?, animations?, partName?,
  sourceXml?, preserved? }`, `PresentationModel { kind: "presentation", id, title,
  format: "pptx"|"odp", slides, slideWidthCm, slideHeightCm, opaque?, media?,
  properties?, structureDirty?, odpContentEnvelope? }`; lengths cm, sizes pt, colours
  lowercase `#rrggbb`; `preserved` is `Record<string,string>`, deleted when it empties;
  `TextFrame` keeps `role?`, `phType?`, `phIdx?`.
- `specs/080-create-presentation-and-slide-management.md` — writer path selection
  (`!structureDirty && !anyDirty && !anyNotes` → every part verbatim; `!structureDirty`
  → text-edit substitution; else full substitution), `usePresentationUndo`
  (`record(label)` **before** mutating, then `commit()`), `stripRetention`,
  `createBlankPresentation`, and the add/delete/move/duplicate commands.
- `specs/081-slide-text-editing.md` — `serializeSlideXml(slide, rels, slidePartPath)`
  and its fixed slot order `<p:sld{rootAttrs}><p:cSld>{bg}<p:spTree>…</p:spTree>
  </p:cSld>{clrMapOvr}{transition}{timing}{extLst}</p:sld>` with `modelled ??
  preserved ?? ""` at every slot; odp `serializePage(slide, registry, usedNames,
  master, styleSource)` → `<draw:page{attrs}>{frames}{notes}{anim}</draw:page>` and
  `StyleRegistry.drawingPageStyle(smilProps, baseProps)` minting `de-DP{n}`;
  `markSlideEdited(slide, format)` (`slide.dirty = true`; odp also `delete
  slide.sourceXml`) is the single dirtying primitive. `specs/083-speaker-notes.md`
  adds `Slide.notes?: string[]` and the notes-slide rebuild this spec extends.
- `specs/089-slide-layouts-and-masters.md` — read here for one thing only: the layout's
  *type*, to know a title slide (absent that model, "title slide" means index 0); no layout
  data is copied here. Placeholder geometry does **not** come from it — SPEC-089's own out
  of scope hands these placeholders to this file — but from the maps both readers already
  build (`buildPlaceholderMap` → `SlideContext.layoutPh`/`masterPh`, keyed `phKey(type,
  idx)` with `parseXfrm` geometry; `buildMasterPlaceholders` → `OdpContext.masterPh`, keyed
  by `presentation:class`), else `DEFAULT_FIELD_BAND`.
- `specs/090-slide-size-and-deck-setup.md` owns deck setup and the slide-number origin
  `PresentationModel.firstSlideNumber?: number` (1–9999, absent = 1), which it reads
  from and writes back to pptx `p:presentation@firstSlideNum`. Read here, never written.
- `specs/053-fields-date-time-and-cross-references.md` — the normative date/time
  catalogue, today in `src/lib/fields.ts`, which `lib/presentation` may not import:
  `DATE_TIME_CATALOGUE: readonly DateTimeFormatEntry[]` (`{ format, docxPicture,
  odtStyleTag: "date"|"time", odtStyleBody, isTime }`), `isSupportedDateTimeFormat`,
  `dateTimeFormatEntry(format)`, `renderDateTime(format: string, instant: Date,
  language: string): string | null`, `odtDateStyleName(entry)` → `FieldDT{index}`,
  `formatFromOdtDateStyleName(name)`.
- Equality: `equalBytes(a, b)` from `src/lib/core/equality.ts` (part-level after
  `unzip`, never container bytes); `equalStructural`/`firstStructuralDifference` from
  `src/lib/equality.ts` (`IGNORED_KEYS` drops `id`, so `Slide.id` is ignored and every
  field below is compared).

### The model (extends SPEC-079's, delta only)
```ts
export type SlideFieldKind = "footer" | "slideNumber" | "date";
export interface DeckHeaderFooter {
  footerText?: string;        // deck default; absent = none
  dateSource?: "fixed";       // absent = auto-updating
  dateFixedText?: string;     // shown when dateSource === "fixed"
  dateFormat?: string;        // a core catalogue pattern; absent = "d MMMM yyyy"
  showOnTitleSlide?: false;   // absent = true; only false is stored
}
export interface SlideFieldPlaceholder {
  kind: SlideFieldKind;
  xCm: number; yCm: number; widthCm: number; heightCm: number;  // resolved
  align?: "left" | "center" | "right";                          // absent = left
  sizePt?: number; color?: string; fontFamily?: string;         // effective look
  text?: string;          // per-slide override; absent = the deck value
  wireType?: string;      // pptx a:fld @type / odp declaration name, verbatim
  wireStyleName?: string; // odp style:data-style-name, verbatim
}
```
`Slide.fields?: SlideFieldPlaceholder[]`; `PresentationModel.headerFooter?` and
`.headerFooterDirty?: true`. All plain JSON — `roundTrip.spec.ts`'s
`JSON.parse(JSON.stringify(model))` assertion must still hold; `slideNumber` entries
never carry `text`. SPEC-080's `usePresentationUndo` `Snapshot` gains `headerFooter?` and
`headerFooterDirty?`, cloned with `slides` and restored by `applySnapshot`, which deletes
them when absent as it already does `structureDirty`.
`DEFAULT_FIELD_BAND` (cm, from `slideWidthCm`/`slideHeightCm`):
height 0.85, `y = slideHeightCm - 1.27`, inset 1.27 each side, three equal boxes with
0.42 gaps, order date/footer/number, align left/center/right, 12pt `#595959`.

### Date formats and the core move (two-phase)
Phase one: move `DateTimeFormatEntry`, `DATE_TIME_CATALOGUE`, `isSupportedDateTimeFormat`,
`dateTimeFormatEntry`, `renderDateTime`, `odtDateStyleName`, `formatFromOdtDateStyleName`
and `canonicalDateTimeFormatFromDocx` **unchanged, names included** into
`src/lib/core/dateTimeFormats.ts`, `src/lib/fields.ts` re-exporting them so
`fields.spec.ts`, `commands.fields.spec.ts` and the docx/odt suites pin the behaviour
untouched; phase two, `presentation/headerFooter.ts` imports core — the shape of
SPEC-079's `core/drawingMl.ts` move. `src/lib/clock.ts` (`now()`, `setClock`) moves the
same way to `src/lib/core/clock.ts`, and `WritePptxOptions`/`WriteOdpOptions` gain a
`now?: Date` beside their existing `modified?: Date`, defaulting to core `now()`, so a
save's recomputed field values are pinnable (the clock-coupled-test rule).
Mapping (pptx `a:fld@type` ⇄ pattern), **provisional — see the golden note**:
`datetime1` ⇄ `dd/MM/yyyy`, `datetime2` ⇄ `EEEE, d MMMM yyyy`, `datetime3` ⇄ `d MMMM
yyyy`, `datetime4` ⇄ `d MMM yyyy`. `MM/dd/yyyy` and `yyyy-MM-dd` have no faithful
`datetime*` equivalent: offered on odp only; on pptx they write a plain `a:r` (this
editor still recomputes it; the dialog says so). Any other `datetime*` → unsupported:
`wireType` verbatim, cached `a:t` rendered, never recomputed. odp is exact both ways
through `odtDateStyleName`/`formatFromOdtDateStyleName`; a foreign data style
(Impress's `D1`) is unsupported: `wireStyleName` verbatim, default pattern rendered.

### pptx wire mapping
Read, per slide: `p:sp` with `p:nvSpPr/p:nvPr/p:ph@type` ∈ `ftr|sldNum|dt` → a
`SlideFieldPlaceholder` (`ftr`→footer, `sldNum`→slideNumber, `dt`→date), the whole `p:sp`
into `preserved["p:ph:{type}"]`; its first `a:fld@type` gives `wireType` (`slidenum`,
`datetime*`), a body with no `a:fld` is fixed text → `text`. Deck:
`p:presentation@showSpecialPlsOnTitleSld="0"` → `showOnTitleSlide: false`
(`@firstSlideNum` is SPEC-090's read); `p:sldMaster`/`p:sldLayout` `p:hf`
(`sldNum`/`hdr`/`ftr`/`dt`, default true, after `p:timing` and before `p:txStyles`) is
read only to know what the master offers, else opaque. Write, per rebuilt slide: start
from `preserved["p:ph:{type}"]` and replace only the `<a:t>…</a:t>` inside its
field/first run; when absent, emit
```
<p:sp><p:nvSpPr><p:cNvPr id="{next}" name="{Footer|Slide Number|Date} Placeholder"/>
<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="{ftr|sldNum|dt}"
sz="quarter" idx="{10|12|11}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm>…rect…</a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="{l|ctr|r}"/><a:fld id="{uuid braces}"
type="{slidenum|datetime*}"><a:t>{computed}</a:t></a:fld></a:p></p:txBody></p:sp>
```
(footer and fixed date use `<a:r><a:rPr lang="en-US"/><a:t>…</a:t></a:r>`.) A suppressed
placeholder is simply not emitted. The `idx` values and `p:hf`'s schema position are two
of the details the golden fixture must confirm.

### odp wire mapping
Read: `office:presentation`'s `<presentation:footer-decl
presentation:name="ftr1">text</…>` and `<presentation:date-time-decl
presentation:name="dtd1" presentation:source="fixed|current-date"
style:data-style-name="D1">fixed text</…>`; each `draw:page`'s
`presentation:use-footer-name`/`use-date-time-name` (already inside
`preserved["@page"]`) picks the declaration; the drawing-page style supplies
`presentation:display-footer`/`display-page-number`/`display-date-time` (**absent is
read as `true`; confirm against a real Impress fixture and let the fixture win**).
Geometry/look come from `OdpContext.masterPh` — `buildMasterPlaceholders` already keys the
master page's `draw:frame presentation:class="footer"|"page-number"|"date-time"` (bodies
`<presentation:footer/>`, `<text:page-number>`, `<presentation:date-time/>`) — else
`DEFAULT_FIELD_BAND`.
Write: declarations regenerated into a copy of `odpContentEnvelope.prefix` (they must
precede the first `draw:page`) only when `headerFooterDirty`; distinct values mint
distinct names in first-use order; an auto date's `style:data-style-name` is
`odtDateStyleName(entry)` (`FieldDT{n}`) with its `<number:date-style
style:name="FieldDT{n}">{odtStyleBody}</…>` spliced through the existing
`spliceAutomaticStyles` path — or the preserved `wireStyleName` when the format never
changed. Per page the `presentation:use-*-name` attributes are set/replaced in the page
attribute string and the display flags are appended to `baseProps` before
`drawingPageStyle(smilProps, baseProps)`, after stripping any existing
`presentation:display-*` (SPEC-085's `smil:*` strip discipline). Field placeholders are
never emitted as `draw:frame`s.

### Golden-fixture-first note (normative process)
Three details above are stated but unconfirmed: PowerPoint's `datetime1..13` render
*locale-dependently*, so the mapping is a proposal; the `p:ph@idx` conventions; and
ODF's default for an absent `presentation:display-*`. Per SPEC-086's process, implementation
checks in `src/lib/presentation/pptx/fieldGolden.ts` and `odp/fieldGolden.ts` holding
hand-authored canonical bytes plus the model they encode, asserts serialiser output equals
them exactly and that the readers accept them plus one attribute-reordered
PowerPoint-flavoured and one Impress-flavoured variant. Where a real file contradicts this
Reference the file wins and the mapping is corrected in the same PR; cross-app verification
(round-tripped deck opened in PowerPoint and Impress, number advancing) is manual QA in the PR.

### Tests
- `src/lib/presentation/headerFooter.spec.ts` — resolver output per kind,
  `slideNumberFor` with/without `firstSlideNumber`, title-slide suppression, per-slide
  `text` precedence, `DEFAULT_FIELD_BAND` maths, `applyHeaderFooter` scope rules and
  dirtying; `src/lib/core/dateTimeFormats.spec.ts` — the moved catalogue's new home,
  with `src/lib/fields.spec.ts` passing unchanged.
- `src/lib/presentation/pptx/serialize.spec.ts`, `odp/serialize.spec.ts` — golden
  placeholder/declaration output, cached-`a:t` rewrite, the `p:presentation` attribute
  splice, the display-flag fold; `notes.spec.ts` — the non-body placeholder delta.
- `src/lib/presentation/headerFooterRoundTrip.spec.ts` — the AC round trip in both formats
  with `equalStructural`, untouched parts with `equalBytes`, the no-edit byte-identical
  download, the pptx `datetime*` delta pinned.
- `src/views/presentation/useHeaderFooter.spec.ts`, `PresentationView.size.spec.ts`
  (lowered budget), `e2e/SPEC-093.spec.ts`.

### Specs this changes
- SPEC-079's out of scope loses "fields beyond slide-number/date placeholders
  rendering their literal cached text", its reader rules gain the placeholder
  diversion above, and its `preserved` key list gains `p:ph:ftr`, `p:ph:sldNum`,
  `p:ph:dt`. Existing outline/frame-kind expectations are re-expressed where a fixture
  slide carried a field placeholder — values unchanged, coverage never weakened.
- SPEC-081's `runXml` gains its one exception (a field placeholder is not a plain run);
  SPEC-080's undo `Snapshot` widens by the two deck-level keys above; SPEC-083's
  notes-slide rebuild preserves non-body placeholders; SPEC-088 renders fields through
  `fieldText`, not their cached strings; SPEC-090 owns `firstSlideNumber` and this file
  only reads it. `specs/000-roadmap.md` gains this spec's phase row; nothing else changes.
