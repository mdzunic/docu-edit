---
id: SPEC-099
title: The slide sorter, deck sections and slide reuse
status: draft
area: structure
complexity: normal
story_points: 8
depends_on: [SPEC-016, SPEC-079, SPEC-080, SPEC-082, SPEC-086, SPEC-089, SPEC-090, SPEC-094]
---
## Why
A deck longer than a dozen slides cannot be reorganised here. `SlideFilmstrip.vue` is one
scrolling column with one selected slide and a one-at-a-time drag: no grid, no
multi-select, no way to group slides into named parts, and no way to move a slide — let
alone five — from one deck into another. SPEC-080 put all of it out of scope ("Copy/paste
of slides across decks or apps"); the roadmap backlog has carried "Slide sorter / grid
view, copy-paste of slides across decks" ever since.

Sections are also a live fidelity bug. PowerPoint stores them in
`p:presentation/p:extLst`, which survives verbatim inside `opaque`, while
`buildSubstitution` (`pptx/writer.ts:522`) re-mints every `<p:sldId id="{256 + index}">`
and splices `p:sldIdLst` alone. Reorder one slide in a sectioned deck and the preserved
`p14:sectionLst` still points at the old ids: the sections silently reshuffle.

The shape is deliberately additive. The sorter is a second *view* over the same
`currentIndex` and the same commands the filmstrip already emits — both render one
extracted `SlideThumb.vue` and read one shared selection composable, so nothing forks.
Sections are a thin ordered list of names plus one key per slide: what `p14:sectionLst`
encodes, and what ODF has no counterpart for, so this file states an honest odp mapping
rather than inventing a native feature. Cross-deck movement is an app-internal clipboard
plus a "Reuse slides" picker; a rich cross-application paste is not authorable from a
browser. **This file is the normative source for the deck-section model, the
`p14:sectionLst` wire mapping, the slide clipboard payload, and the import-repair rules
every cross-deck slide copy obeys.** SPEC-089 owns layouts and SPEC-094 the theme; this
file only consumes them at import time.

## Acceptance criteria
- [ ] `PresentationModel` gains `sections?: DeckSection[]` and `Slide` gains `sectionKey?: string` per the Reference (`DeckSection { key; name; collapsed?: true }`, `key` deterministic and deliberately not named `id` because `equalStructural`'s `IGNORED_KEYS` drops that key); absent `sections` means a deck with none and `[]` is never stored; `collapsed` is editor view state serialised to neither format; the model stays plain JSON, so `roundTrip.spec.ts`'s `JSON.parse(JSON.stringify(model))` `toEqual(model)` still holds
- [ ] `src/lib/presentation/sections.ts` (framework-free; presentation siblings + `lib/core` only) exports `sectionsOf(model): SectionRun[]`, `addSection(model, atIndex, name): string`, `renameSection(model, key, name)`, `moveSection(model, key, delta): number`, `removeSection(model, key, mode: "sectionOnly" | "withSlides"): number`, `setSlideSection(model, slideIndex, key)`, `normaliseSections(model)` and `MAX_SECTION_NAME_LENGTH = 120`; every command sets `model.structureDirty = true`, `addSection` at an index > 0 with nothing covering `[0, atIndex)` first mints a leading "Default Section" so the run invariant cannot be violated, and `normaliseSections` — idempotent — runs as the tail of every structural command in `src/lib/presentation/slides.ts` (`addSlide`, `duplicateSlide`, `deleteSlide`, `moveSlide`) plus paste and import: a slide whose `sectionKey` names no entry adopts its predecessor's (or the first entry when it is `slides[0]`), `sections` re-orders to first-appearance order with an emptied section keeping its slot, and a key held by no slide and no entry is dropped
- [ ] The toolbar gains a "Sorter" toggle (`slides-sorter-toggle`, `aria-pressed`) in `PresentationToolbar.vue`'s right-aligned `div.tb-group.tb-slide-actions` beside `slides-transition-toggle`, backed by a new `presentationSorterOpen` preference added to all three places in `src/lib/preferences.ts` (interface, defaults `false`, `isBoolean` validator); it swaps `.deck-body` for `src/components/presentation/SlideSorter.vue` — root `slides-sorter`, a `ul[role="listbox"][aria-multiselectable="true"]` of `sorter-slide` cards, each a `<SlideStage :interactive="false">` thumbnail at `filmstripThumbScale(widthCm, 208)` with its 1-based number and `aria-selected`, plus `sorter-live` (`role="status"`) announcing selection size and move results
- [ ] Multi-selection lives in `src/views/presentation/useSlideSelection.ts` and is shared verbatim by the sorter **and** `SlideFilmstrip.vue`, which gains a `selectedIndexes` prop and renders the same extracted `src/components/presentation/SlideThumb.vue`: click selects one and sets `currentIndex`, Cmd/Ctrl+click toggles, Shift+click extends from the anchor, Cmd/Ctrl+A selects all, Escape collapses back to the current slide, arrows/Home/End move the roving selection and Shift+arrow extends — the filmstrip's existing keyboard, context menu and `reorder` emit are reused, never forked
- [ ] Sorter direct manipulation acts on the whole selection: drag-to-reorder shows one `sorter-drop-indicator` and moves every selected slide as a block to the drop index (one `moveSlide` per slide, one undo entry "Move slides"), `sorter-duplicate` / Cmd/Ctrl+D duplicates them in place, `sorter-delete` / Delete removes them, dropping outside cancels, and each card's context menu reuses `SlideFilmstrip.vue`'s existing `slide-menu-*` items
- [ ] Sections render as run headers in the sorter: `sorter-section` with `sorter-section-name` (double-click or `section-rename` opens `section-rename-input`; Enter commits a trimmed 1–120 character name, Escape cancels, empty is rejected), `sorter-section-toggle` (`aria-expanded`; collapsing hides the run's cards and shows its slide count), and `sorter-section-menu` holding `section-add`, `section-rename`, `section-move-up`, `section-move-down`, `section-remove` and `section-remove-with-slides`; moving a section moves its whole run, and a slide dragged past a header adopts the section it lands in
- [ ] pptx sections read and write per "pptx wire mapping": the reader models any `p14:sectionLst` it finds (names in document order, membership by `p14:sldId/@id` resolved through the `p:sldIdLst` walk at `pptx/reader.ts:882`, unknown ids ignored, a slide left unassigned joining the preceding run) so a section list is never left preserved and the model is the single source; the writer recomputes membership from the model's slide order against the freshly minted `256 + index` ids inside `buildSubstitution`, and `spliceSectionLst` replaces only the existing `sectionLst` element — reusing the prefix it found and leaving the enclosing `p:ext` and every sibling `p:ext` byte-verbatim — else splices a whole `p:ext`/`p:extLst` per the mapping, adding `xmlns:p14` + `mc:Ignorable="p14"` to the root only when absent
- [ ] odp sections are stated, not faked: ODF 1.3 has no deck-section construct, so they ride in `meta.xml` as one `meta:user-defined meta:name="docu-edit:sections" meta:value-type="string"` custom property holding the compact JSON of "The odp mapping" — read through `parseOdfCustomProperties` (`src/lib/core/odfMeta.ts`), which `odp/reader.ts:675` already calls, and written through `buildMetaXml`, so `writeOdp` gains `meta.xml` to its substituted parts (a deck whose sections did not change re-emits it byte-verbatim, and unrelated `meta:user-defined` elements survive through `customProperties`/`unsupportedCustomProperties`); `sorter-odp-note` says once per odp deck that Impress will not display them
- [ ] Copy/cut/paste of slides works within a deck and across decks **in this app**: `src/lib/presentation/slideClipboard.ts` exports `SLIDES_MIME = "application/x-docu-edit-slides+json"`, `MAX_SLIDE_CLIPBOARD_BYTES = 8 * 1024 * 1024`, `slideClipboardFlavours(model, indexes)`, `parseSlideClipboardPayload(raw): SlideClipboardPayload | null` (a whitelist rebuild, never a spliced `JSON.parse`) and a module-level in-tab store keyed by `nonce`; `sorter-copy`/`sorter-cut`/`sorter-paste` and Cmd/Ctrl+C/X/V write `text/plain` (the slides' `slideOutline` text) plus `SLIDES_MIME` through `ClipboardEvent.clipboardData` exactly as SPEC-016's `DOC_MODEL_MIME` does, omitting the private flavour above the cap while the in-tab store still pastes those slides intact; cross-application behaviour is specified rather than hoped for — pasting these slides into PowerPoint, Keynote or a text editor yields the outline **text** and never slides (no `PowerPoint 14.0 Slides Package`/`Art::GVML` flavour is authorable from a browser), and a paste carrying no recognised `SLIDES_MIME` payload falls through to SPEC-081's text-session `onPaste` or, outside a text session, is a no-op announcing "The clipboard doesn't hold any slides" on `sorter-live`, so the sorter never silently produces an empty slide
- [ ] Every slide entering a deck from elsewhere goes through one `importSlides(target, source, indexes, options): SlideImportReport` in `src/lib/presentation/slideImport.ts` obeying "Import repair": fresh `Slide.id` and fresh `SlideFrame.id` for every frame and group child via `slides.ts`'s `remintFrameIds` (exported for this), animations retargeted through that same id map with SPEC-086's `retargetAnimations`, per-slide retention stripped by the newly factored-out `stripSlideParts(slide)` — `partName`, `sourceXml` and `preserved` deleted, and nothing else — after which **this path** sets `dirty = true` and `notesDirty = true` so the imported slide rebuilds from the model in both formats, while `stripRetention` keeps *deleting* `dirty` and keeps leaving `notesDirty` alone exactly as it does today, then `target.structureDirty = true`; `SlideImportReport { slides: number; droppedFrames: number; droppedFrameNames: string[]; lostTransitions: number; lostAnimations: number; mediaReused: number; rescaled: boolean }` renders as `sorter-import-report` with `sorter-import-report-dismiss` and is honest about what stripping costs — a transition or timing that existed only as `preserved["p:transition"]`/`["p:timing"]`/`["draw:transition"]`/`["anim"]` arrives without it and is counted, and a `GroupFrame`/`ForeignFrame` whose verbatim `xml` carries a package reference this editor cannot re-home (`r:embed`, `r:id`, `r:link`, an odp `xlink:href` into the source package) other than a plain image is dropped and named through `frameAccessibleName`, never emitted as a dangling reference
- [ ] Media travels de-duplicated: each imported `ImageFrame`'s `MediaPart` is matched against `target.media` by `contentType` + base64 `data`, reusing that `mediaId` when identical (counted in `mediaReused`) and otherwise minting a fresh one through `frames.ts`'s `mintMediaId` (exported for this) so the part lands in the **target's** directory per `mediaDirectory(model)` (`ppt/media/` or `Pictures/`), with the frame's `mediaId` remapped; `target.media` is replaced with a new array rather than mutated in place, so `documentStore.ts`'s `persistMedia` reference latch fires
- [ ] "Reuse slides" imports from another deck in the library: `slides-reuse-toggle` opens `slides-reuse-dialog` over `src/components/BaseDialog.vue` with `reuse-deck-select` (from `listDocuments()` filtered to `kind === "presentation"`, excluding this deck and any `isLegacyPresentationModel` record, read with `getPresentation(id)`), a listbox of `reuse-slide` thumbnails plus `reuse-select-all`, the radio pair `reuse-use-destination` (default) / `reuse-keep-source`, and `reuse-insert` / `reuse-cancel`; **use destination** leaves every SPEC-094 `ThemeColorRef`/`ModelFont` a reference so it resolves against this deck's theme and maps `layoutKey` to the destination layout of the same `SlideLayoutKind` (SPEC-089) or drops it, **keep source** resolves every reference to a literal through the *source* deck's theme with `resolveColor`/`resolveFont` (`src/lib/core/theme.ts`) and drops `layoutKey`; either way a slide-size mismatch rescales per SPEC-090 and sets `report.rescaled`
- [ ] `PresentationView.vue` gains exactly one `useSlideSorter({ model: presentation, currentIndex, record: slideUndo.record, commit, reportError })` call plus the `<SlideSorter>` / `<ReuseSlidesDialog>` mounts and one toolbar prop/emit pair; the deck-navigation block (`slideTotal`, `currentSlide`, `deckMedia`, `liveLabel`, `slideCountLabel`, `selectSlide`, `addSlideAction`, `onFilmstripCommand`, `onFilmstripReorder`) is extracted to `src/views/presentation/useDeckNavigation.ts` and `src/views/PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 15 from whatever it stands at when this lands (550 → 535 today) with a ledger row appended — never up; every mutation is `record(label)` → mutate → `commit()` with the labels "New section", "Rename section", "Move section", "Remove section", "Move slides", "Duplicate slides", "Delete slides", "Paste slides", "Reuse slides"
- [ ] Round trip: a deck with three sections, a pasted slide and two reused slides → save → reopen reproduces `sections`, every `sectionKey`, slide order, retargeted animations and remapped `mediaId`s under `equalStructural` (`src/lib/equality.ts`) over the modelled subset — no semantic relation is claimed, because `equalSemantic` is typed `(model: DocumentModel)` and is unusable on a `PresentationModel` — while a section-only edit re-emits `ppt/presentation.xml` (+ its `.rels` and `[Content_Types].xml`, per SPEC-080's structural path) / `meta.xml` and every other part is `equalBytes` (`src/lib/core/equality.ts`) against the source after `unzip`; a no-edit download stays byte-identical part-for-part, and `stripRetention` keeps `sections`/`sectionKey` so a cross-format "Download a copy as…" carries them into the other format's mapping
- [ ] Fixture coverage: `fixtures/build.ts` gains a `spec099-sections` builder pair appended under a banner comment (never mutating the four mini-deck slides) — six slides in two named sections, one carrying an image, one an animation, and one deliberately out-of-subset deck-level extra (a second unrelated `p:ext` in `p:extLst` for pptx, an unrelated `meta:user-defined` for odp) to pin the preserve path — with its expectation table beside it, and `scripts/build-presentation-fixtures.ts` writes both new files into `e2e/fixtures/`
- [ ] e2e (`e2e/SPEC-099.spec.ts`): open `mini-deck.pptx`, return to the library, open `spec099-sections.pptx` → Sorter → six cards under two section headers → Shift+click selects slides 2–3 → Copy → click slide 5 → Paste → two copies land after it inside that section → collapse the first section, then drag a slide into the second → Reuse slides → pick "mini-deck" → select two → "Use destination theme" → Insert → the report names what it kept → download, reopen → sections, names and order unchanged and the imported slide's image renders

## Out of scope
Section zoom and summary-zoom slides; custom slide shows (`p:custShowLst`,
`presentation:show`), per-slide hide/skip and exporting one section; presenter view
(SPEC-100); a rich cross-application slide flavour in either direction — pasting a
PowerPoint slide *in* stays unsupported, and dropping a `.pptx` onto the sorter is
SPEC-009's open path; importing a source deck's layouts, masters, themes or media
library (SPEC-089 owns layouts, SPEC-094 the theme; only what a copied slide carries is
re-homed); merging two decks as one command; a shared undo stack across two decks.

## Reference

### Dependencies restated
- `specs/079-slide-canvas-rendering.md` owns the slide model: `Slide` is `{ id, frames,
  background?, notes?, dirty?, notesDirty?, transition?, animations?, partName?,
  sourceXml?, preserved? }`, `SlideFrame = TextFrame | ImageFrame | ShapeFrame |
  GroupFrame | ForeignFrame`, and `PresentationModel` carries `kind: "presentation"`,
  `id`, `title`, `format`, `slides`, `slideWidthCm`/`slideHeightCm`, `opaque?`, `media?`,
  `properties?`, `structureDirty?` and `odpContentEnvelope?`; lengths cm. **`Slide.id`
  and `SlideFrame.id` are `crypto.randomUUID()` minted at read, never serialised.**
- `specs/080-create-presentation-and-slide-management.md` owns the rebuild rules and the
  undo stack: `writePptx` re-zips every part verbatim when nothing is dirty and
  `structureDirty` is unset, else substitutes per slide (`buildTextEditSubstitution`) or
  rebuilds `presentation.xml`, its `.rels` and `[Content_Types].xml`
  (`buildSubstitution`, minting `<p:sldId id="{256 + index}" r:id="…"/>` through
  `spliceSldIdLst` at `writer.ts:719`); `writeOdp` rebuilds `content.xml` from
  `odpContentEnvelope` when `structureDirty || any dirty`. `slides.ts` exports
  `addSlide`, `duplicateSlide`, `deleteSlide` and `moveSlide(model, from, to)` — each
  returning the index to select and setting `structureDirty` — plus `stripRetention`,
  which drops `opaque`, `odpContentEnvelope`, `structureDirty` and per-slide
  `partName`/`sourceXml`/`preserved`/`dirty`; `remintFrameIds(frames, idMap)` is private
  there. Undo is `record(label)` before the mutation, then `commit()`.
- `specs/082-slide-objects-images-shapes-text-boxes.md` owns frames and media:
  `MediaPart { mediaId; contentType; data }` (`src/lib/core/media.ts`),
  `mediaDirectory(model)` → `"Pictures/" | "ppt/media/"`, `frames.ts`'s private
  `mintMediaId(model, contentType)` → `{dir}image{n}.{ext}`, and
  `frameAccessibleName(frame)`.
  `specs/086-slide-animations.md` owns `SlideAnimation { frameId; effect; direction?;
  trigger; durationMs?; delayMs? }` and `retargetAnimations(animations, idMap:
  Map<string, string>)`, which drops any animation whose target the map lacks —
  `duplicateSlide` is the precedent, reused unchanged.
- `specs/089-slide-layouts-and-masters.md` owns `SlideLayout { key; kind; name; slots;
  masterName? }`, `SlideLayoutKind` and `Slide.layoutKey` (mapped by kind or dropped here,
  never authored); `specs/094-deck-themes.md` owns `DocumentTheme`, `ThemeColorRef`,
  `ModelColor`/`ModelFont` and `resolveColor(c, theme: DocumentTheme)`/`resolveFont` in
  `src/lib/core/theme.ts`, plus `DeckTheme extends DocumentTheme` (adding `colorMap`,
  `partName`, `dirty`) in `presentation/model.ts` — resolved or preserved here, never
  edited; `specs/090-slide-size-and-deck-setup.md` owns the rescale maths in
  `deckSetup.ts`, whose per-slide half `scaleSlideGeometry(slide, factor, offsetXCm,
  offsetYCm)` is created there by whichever spec lands first.
- `specs/016-cut-copy-paste.md` owns the clipboard discipline: `DOC_MODEL_MIME =
  "application/x-docu-edit-docmodel+json"` in `src/lib/clipboard.ts`, its `{ mimeType,
  version, … }` envelope, the rule that only `ClipboardEvent.clipboardData` carries a
  private type, and a whitelist parser because the clipboard is untrusted.
- Equality: `equalBytes(a: Uint8Array, b: Uint8Array): boolean`
  (`src/lib/core/equality.ts`), part-level after `unzip` since the container is rebuilt
  on save; `equalStructural(a: unknown, b: unknown)` (`src/lib/equality.ts`), whose
  `IGNORED_KEYS` drops `"id"` — hence `DeckSection.key`. `equalSemantic` is typed
  `(model: DocumentModel)` and **cannot** be used on a `PresentationModel`; no such
  relation exists here and this spec does not invent one.
  `presentation/boundary.spec.ts` restricts `lib/presentation` to presentation siblings
  and `lib/core` (a bare `"vue"` is itself a violation).

### The section model (normative)
```ts
export interface DeckSection {
  /** Deck-stable — pptx: the source `p14:section@id` (a brace GUID), else a minted
   *  `de-sec{n}`. NOT named `id`. */
  key: string;
  name: string;      // trimmed, 1..MAX_SECTION_NAME_LENGTH
  collapsed?: true;  // editor view state only — serialised to neither format
}
export interface SectionRun { section: DeckSection; startIndex: number; endIndex: number }
```
`PresentationModel.sections?: DeckSection[]` in deck order; `Slide.sectionKey?: string`.
Invariant: `sections` is absent, **or** `slides[0].sectionKey` names one of its entries —
a leading unsectioned run is impossible.

### pptx wire mapping
```xml
<p:presentation … xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p14">
  …<p:sldIdLst><p:sldId id="256" r:id="rId2"/>…</p:sldIdLst>…
  <p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst>
    <p14:section name="Intro" id="{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}">
      <p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst>
    </p14:section></p14:sectionLst></p:ext></p:extLst>
</p:presentation>
```
The `uri` GUID, the `p14` namespace URI and the element names above are the wire detail this
spec is least certain of, so they are **pinned by a checked-in golden, not by memory**,
SPEC-086's process applying unchanged: `pptx/sectionsGolden.ts` holds `SECTIONS_GOLDEN` and
the `GOLDEN_SECTIONS` model it encodes, with PowerPoint verification recorded in the PR, not
CI. `spliceSectionLst(presXml, xml)` mirrors `spliceSldIdLst`: replace a match of
`/<([A-Za-z0-9]+):sectionLst[\s\S]*?<\/\1:sectionLst>/` in place, reusing the prefix found
and touching neither the enclosing nor any sibling `p:ext`; else splice a `p:ext` before
`</p:extLst>`; else insert a whole `p:extLst` last in `p:presentation`.

### The odp mapping (a stated loss, not an invention)
ODF 1.3 has no deck-section construct (`presentation:show` is a custom slide show, a
different feature), so sections ride in `meta.xml` as one custom property:
```
meta:name="docu-edit:sections"   meta:value-type="string"
value = {"v":1,"sections":[{"key":"de-sec1","name":"Intro","n":2},…]}
```
`n` is the run length, so membership needs no per-slide attribute; a total overrunning the
deck is clamped at read and a shortfall extends the last run. This is valid ODF that
LibreOffice keeps as a custom property and displays as sections nowhere, which is why
`sorter-odp-note` says so once — a silent no-op would be the fidelity lie this repo's first
law forbids. Read is free (`odp/reader.ts:675` fills `properties.customProperties`); write
adds `meta.xml` to `writeOdp`'s substituted parts through `buildMetaXml`.

### Import repair (normative — every cross-deck path funnels here)
`importSlides(target, source, indexes, options)` in `slideImport.ts`, per slide:
1. `structuredClone` the source slide; fresh `Slide.id`; `remintFrameIds(frames, idMap)`
   recursing into `GroupFrame.children` for fresh `SlideFrame.id`s.
2. Drop frames that cannot be re-homed, recording their `frameAccessibleName`.
3. `animations = retargetAnimations(animations, idMap)` — a dropped frame takes its
   animations with it; a modelled `Slide.transition` survives.
4. `stripSlideParts(slide)` — only the `partName`/`sourceXml`/`preserved` deletion,
   factored out of `stripRetention` (`slides.ts:195`) and called by both. The two
   callers keep their *own* dirty handling: `stripRetention` still deletes `dirty`
   and still never touches `notesDirty`, so `slides.spec.ts` is preserved and never
   weakened, while this path sets both flags to `true`.
5. Re-home media, rescale when the slide sizes differ, apply the theme/layout option.
6. Splice into `target.slides`, take `sectionKey` from the insertion point, then run
   `normaliseSections(target)` and set `target.structureDirty = true`.

### The clipboard payload, and the core hoist (two-phase)
```ts
export interface SlideClipboardPayload {
  mimeType: typeof SLIDES_MIME; version: 1; nonce: string;
  format: PresentationFormat; slideWidthCm: number; slideHeightCm: number;
  slides: Slide[];      // ids present but ignored — always re-minted
  media: MediaPart[];   // only the parts the copied slides reference
  theme?: DeckTheme;    // SPEC-094, for "keep source formatting"
}
```
The envelope and its guards are the product-neutral half of `src/lib/clipboard.ts`, and all
that moves — but the envelope pair must be *extracted*, not relocated: no
`parseClipboardEnvelope` exists today, the check being inlined in `parseFragmentPayload`
(`clipboard.ts:365`, `payload.mimeType !== DOC_MODEL_MIME || payload.version !== 1` at line
376) and the builder inside `serializeFragment` (line 215). **Phase one** lifts
`parseClipboardEnvelope(raw, mimeType, version)` out of those two and hoists it with the
private `isFiniteNumber`/`isSafeColor` guards into a new `src/lib/core/clipboardPayload.ts`
that both callers then use, so `src/lib/clipboard.spec.ts` passes untouched — coverage
preserved, never weakened. **Phase two** builds `slideClipboard.ts` on them, its sanitiser
also accepting a SPEC-094 `ThemeColorRef`.

### View wiring
`src/views/presentation/useSlideSorter.ts` — `SlideSorterDeps { model:
ShallowRef<PresentationModel | null>; currentIndex: Ref<number>; record: (label?:
string) => void; commit: () => void; reportError: (message: string) => void }` →
`{ open, toggle, rows, selection, selectAt, selectAll, clear, addSection, renameSection,
moveSection, removeSection, toggleCollapsed, copy, cut, paste, canPaste, duplicate,
remove, reorder, reuse, report, dismissReport, onGlobalKeydown }`, building
`useSlideSelection` and `useReuseSlides` internally as `usePresentationEditing` builds
`useFrameSelection`, `open = computed(() => preferences.presentationSorterOpen)`, `rows` a
fresh-array computed per the `useAnimations` rule.

### Tests
- `sections.spec.ts` — every command, the "Default Section" rule, `normaliseSections`
  idempotence after each structural command, paste and import.
- `slideImport.spec.ts` — id re-minting including group children, animation retargeting
  and drop-with-frame, `stripSlideParts` leaving each caller's dirty handling intact,
  media de-duplication, rescale, both options.
- `slideClipboard.spec.ts` — flavours, the cap, the whitelist parser against hostile
  payloads, nonce precedence; plus `src/lib/core/clipboardPayload.spec.ts` with
  `src/lib/clipboard.spec.ts` unchanged, proving the hoist behaviour-preserving.
- `pptx/sections.spec.ts` and `odp/sections.spec.ts` — golden equality both ways, the
  in-place splice, the sibling-`p:ext` and sibling-`meta:user-defined` preserve paths, and
  membership recomputed after a reorder; `sectionsRoundTrip.spec.ts` runs the round-trip
  AC's deck, `roundTrip.spec.ts` staying untouched.
- `useSlideSorter.spec.ts`, `useSlideSelection.spec.ts`, `SlideSorter.spec.ts`,
  `SlideFilmstrip.spec.ts` (extended, never weakened), `e2e/SPEC-099.spec.ts`.

### Specs this changes
- `specs/080-create-presentation-and-slide-management.md` — its "Copy/paste of slides
  across decks or apps" exclusion is resolved here; its structural commands gain the
  `normaliseSections` tail, `remintFrameIds` and `stripSlideParts` become exported, and
  `buildSubstitution` gains the `spliceSectionLst` call — existing expectations
  re-expressed against the same values, never weakened.
- `specs/079-slide-canvas-rendering.md` — the filmstrip gains multi-selection and renders
  the extracted `SlideThumb.vue`, its listbox semantics otherwise unchanged;
  `specs/016-cut-copy-paste.md` — its envelope and guards move to
  `src/lib/core/clipboardPayload.ts`, contract unchanged.
- `specs/089-slide-layouts-and-masters.md` and `specs/094-deck-themes.md` gain a consumer
  at import time; `specs/090-slide-size-and-deck-setup.md` shares `scaleSlideGeometry`.
  `specs/000-roadmap.md` gains this spec's phase row and loses the backlog line "Slide
  sorter / grid view, copy-paste of slides across decks"; nothing else changes.
