---
id: SPEC-089
title: Slide layouts — the layout catalogue, picker and Reset slide
status: draft
area: layout
depends_on: [SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-086]
---
## Why
SPEC-079 resolves placeholder inheritance through layout → master at read and
then throws the chain away — `SlideContext.layoutPh`/`masterPh` bake effective
geometry into the frames and die with the parse, so nothing downstream knows a
slide *has* a layout. SPEC-080's out of scope pins every new slide to title +
body with no picker. A deck's own "Section header", "Two content" and "Blank"
layouts are therefore invisible, every New slide looks identical, and a
placeholder dragged out of place can never be put back.

The shape is deliberately conservative: model the deck's **own** layouts as a
read-derived catalogue keyed deterministically (pptx: the layout part path; odp:
master + page-layout style name), pick from it with real thumbnails rendered by
the existing `SlideStage`, and apply by re-mapping placeholders by role and
`idx`, keeping their text. Writing reuses an existing layout part **by
relationship**; a part is authored only when the deck carries no counterpart at
all, and SPEC-080's `buildFromScratch` grows from its single `type="blank"`
layout to a minimal standard set. Layout and master *editing* is a separate
future spec: this file picks and applies, it never authors a master.

**This file is the normative source for the layout model (`SlideLayout`,
`LayoutSlot`, `SlideLayoutKind`), layout keys, and apply/reset semantics.**
SPEC-093 (headers and footers), SPEC-094 (themes) and SPEC-102 (accessibility)
build on it.

## Acceptance criteria
- [ ] `PresentationModel` gains `layouts?: SlideLayout[]` and `Slide` gains `layoutKey?: string` per the Reference: a deterministic `key` (never a UUID, and never named `id` — see the Equality note), `kind` from the closed union title/titleAndBody/sectionHeader/twoContent/titleOnly/blank/other, the deck's own `name`, and `slots: LayoutSlot[]` carrying role + `phType`/`phIdx` + cm geometry; absent `layouts` means a deck stored before this spec (re-derived on the next open), `[]` means a package with no layout parts; the three types are declared in `src/lib/presentation/model.ts` beside `SlideTransition`/`SlideAnimation`, and their rules (`STANDARD_LAYOUTS`, `classifyLayout`, `layoutFor(model, slide)`, `applySlideLayout`, `resetSlideLayout`, `LayoutApplyReport`) in the new framework-free `src/lib/presentation/layouts.ts` (presentation siblings + `lib/core` only)
- [ ] Both readers populate the catalogue without dirtying anything: pptx walks `p:sldLayout` parts reached from each slide's `.rels` `slideLayout` relationship (and the master's `p:sldLayoutIdLst`), keying by absolute part path, naming by `p:cSld/@name` and classifying by `p:sldLayout/@type`; odp reads `presentation:presentation-page-layout-name` off `draw:page` and the matching `style:presentation-page-layout` in `styles.xml`, keying `` `${master}::${pageLayout}` ``; readers never set `dirty`/`structureDirty`, never mutate `opaque`, and never rewrite `preserved["@page"]`
- [ ] A layout that matches no standard kind classifies `"other"`, keeps its own name and slots, and is offered and applied unchanged; a deck with no layout parts (or an odp with no page-layout style) yields `layouts: []` and the gallery falls back to the synthesised standard set marked "will be added to this deck" (`layout-card-new-badge`)
- [ ] `applySlideLayout(model, slideIndex, layoutKey): LayoutApplyReport` re-maps placeholders per the Reference — exact `phType`+`phIdx` match first, then same role in document order, then role order title→subtitle→body→other — moving each matched frame to its slot's geometry and adopting the slot's `phType`/`phIdx`, keeping every paragraph; unmatched placeholders and every non-placeholder frame (no `role`) stay exactly where they are and are counted in `report.keptInPlace`; it sets `slide.layoutKey` and ends with `markSlideEdited(slide, model.format)`
- [ ] `resetSlideLayout(model, slideIndex): LayoutApplyReport` snaps every placeholder frame back to its slot's geometry (the model stores effective geometry only, so "reset" means re-applying the layout's slots) leaving text, fills, strokes, non-placeholder frames, notes, transitions and animations untouched; it is disabled (`layout-reset-slide` `aria-disabled`) when the slide has no resolvable layout
- [ ] `addSlide(model, afterIndex, layoutKey?)` extends the command SPEC-080 defines: with a key, the new slide's frames are one empty `TextFrame` per slot of that layout (role/`phType`/`phIdx`/geometry from the slot) and `layoutKey` is set; without one, SPEC-080's title + body behaviour is unchanged
- [ ] The layout gallery is a modal over the shared `src/components/BaseDialog.vue` (the shell SPEC-046's `TemplateGalleryDialog.vue` uses), root `slides-layout-gallery`, `ul[role="listbox"]` of `button[role="option"][aria-selected]` cards `layout-card` / `layout-card-${index}`, each card a **real thumbnail** — a synthesised preview `Slide` of the layout's slots rendered by `<SlideStage :interactive="false">` at `filmstripThumbScale(widthCm, 176)` — plus the layout's name; `layout-apply`, `layout-gallery-close`, `layout-reset-slide`, arrow-key/Home/End roving focus with Enter to apply and Escape to close, and — after an apply that could not place everything — a dismissible report (`layout-report`, `layout-report-dismiss`) naming the count and the frame names from `useFrameManipulation.ts#frameAccessibleName` ("2 items kept in place"); nothing is ever deleted
- [ ] Entry points, all keyboard-operable: a `Layout` toolbar toggle (`slides-layout-toggle`, in `PresentationToolbar.vue`'s right-aligned `div.tb-group.tb-slide-actions` beside `slides-transition-toggle`/`slides-animate-toggle`, prop `layoutOpen`, emit `layout-toggle`) opening it in change mode for the current slide; filmstrip context-menu items "Change layout…" (`slide-menu-layout`) and "New slide with layout…" (`slide-menu-new-with-layout`); the gallery is a dialog, not a panel, so `src/lib/preferences.ts` gains **no** key
- [ ] Every change dirties the slide, records exactly one undo entry per gesture with labels "Change slide layout" / "Reset slide" (SPEC-080's `record(label)` before mutate, `commit()` after) and persists write-through; undo restores geometry, `phType`/`phIdx` and `layoutKey` together
- [ ] pptx write: a slide whose `layoutKey` differs from its source `.rels` re-emits **only** that `.rels` with the `slideLayout` relationship's `Target` repointed (`relativeTarget(slidePart, layoutPart)`), every other relationship byte-identical — a delta to both `buildTextEditSubstitution` (writer.ts:100) and `buildSubstitution`'s `rebuiltOwn` block; `resolveLayoutPath` (writer.ts:692) gains a first branch honouring `slide.layoutKey` ahead of its nearest-preceding-slide heuristic; a synthesised layout with no counterpart is authored once as `ppt/slideLayouts/slideLayout{n}.xml` past the highest existing, with its `.rels` to the master, a `[Content_Types].xml` override and a `p:sldLayoutId` spliced into the master's `p:sldLayoutIdLst`
- [ ] odp write: a layout change rewrites `presentation:presentation-page-layout-name` inside `preserved["@page"]` (adding the attribute when absent) exactly as SPEC-085's transition path rewrites `draw:style-name` via `setDrawStyleName` in `odp/serialize.ts`; `draw:master-page-name` is never changed (layout and master are separate ODF axes); a synthesised page layout splices one `<style:presentation-page-layout style:name="de-AL{n}">` before `</office:styles>` in `styles.xml`, for which **this file introduces `writeOdp`'s `styles.xml` substitution slot**, beside the `content.xml`/`META-INF/manifest.xml`/media parts it already substitutes — a deck that changed no layout re-emits `styles.xml` byte-verbatim
- [ ] From-scratch decks carry the minimal standard set: pptx `buildFromScratch` emits `slideLayout1..6` (title/obj/secHead/twoObj/titleOnly/blank per the Reference table) with six `p:sldLayoutId`s and six Content-Types overrides instead of today's single `type="blank"` part; odp `buildFromScratch` emits the six `style:presentation-page-layout` styles in `styles.xml` and `serializePage` stamps the chosen `presentation:presentation-page-layout-name` on each page
- [ ] Round trip: open → change one slide's layout → save → reopen reproduces `layoutKey`, the same catalogue entry and placeholder geometry within ±0.01 cm in both formats, compared with `equalStructural` (`src/lib/equality.ts`) over the modelled subset; every part the change did not touch is `equalBytes` (`src/lib/core/equality.ts`) against the source after `unzip` — slide `.rels` only for the changed slide, `styles.xml` only when a page layout was authored — and a no-edit download stays byte-identical part-for-part
- [ ] Fixture coverage: `fixtures/build.ts` gains a `spec089-layouts` builder pair (never mutating the four mini-deck slides) whose decks carry the six standard layouts, one deliberately **exotic** layout (`p:sldLayout type="cust"` / an Impress page layout with an unrecognised placeholder object) to pin the `"other"` path, and one slide with no layout reference at all; `scripts/build-presentation-fixtures.ts` writes both new files into `e2e/fixtures/`
- [ ] e2e (`e2e/SPEC-089.spec.ts`): open `spec089-layouts.pptx` → filmstrip menu → "New slide with layout…" → gallery shows the deck's layouts as thumbnails → choose "Section header" → new slide has exactly that layout's placeholders → type a title → toolbar Layout → choose "Two content" → title text survives and two body placeholders appear → drag the title away → Reset slide → it snaps back → download, reopen → the slide still reports "Two content" as its layout

## Out of scope
Creating, editing, renaming, reordering or deleting layouts and masters (the
separate future spec this one leaves room for), including the slide master's
shapes, background and `p:txStyles`, the notes master and the handout master;
per-layout headers, footers, slide-number and date placeholders (SPEC-093) and
theme colour/font schemes (SPEC-094); thumbnails of real slide content (they
show slots); placeholder prompt-text authoring; multiple masters as a
user-visible concept (their layouts simply flatten into one catalogue); new
roles for `p:ph` types outside SPEC-079's `roleOf` mapping; per-slide colour-map
overrides (`p:clrMapOvr` stays preserved verbatim per SPEC-081); vertical-text
and picture-caption kinds, which classify `"other"` and still work.

## Reference

### Dependencies restated
- `specs/079-slide-canvas-rendering.md` owns the slide model: `Slide` is `{ id,
  frames, background?, notes?, dirty?, notesDirty?, transition?, animations?,
  partName?, sourceXml?, preserved? }` and `TextFrame` is `{ id, xCm, yCm,
  widthCm, heightCm, rotationDeg?, kind: "text", role?: PlaceholderRole,
  phType?, phIdx?, paragraphs, fill?, stroke? }`, `PlaceholderRole = "title" |
  "subtitle" | "body" | "other"`, lengths cm. Its reader keys inheritance by the
  private `phKey(type, idx)` (`ctrTitle`→`title`, absent type→`body`) — kept
  here as the match rule and exported from `pptx/reader.ts` for `layouts.ts`.
- `specs/080-create-presentation-and-slide-management.md` owns the rebuild
  rules and the undo stack: `writePptx` picks verbatim /
  `buildTextEditSubstitution` / `buildSubstitution` off `structureDirty` +
  per-slide `dirty`; `writeOdp` rebuilds `content.xml` from
  `odpContentEnvelope` when `structureDirty || any dirty`; `usePresentationUndo`
  is `record(label) → mutate → commit()` over `structuredClone(model.slides)`.
  `addSlide(model, afterIndex): number`, `createBlankPresentation()` and
  `stripRetention(model)` live in `src/lib/presentation/slides.ts`.
- `specs/081-slide-text-editing.md` owns slide-part serialisation:
  `serializeSlideXml(slide, rels, slidePartPath)` with `modelled ?? preserved ??
  ""` at every slot; `serializePage(slide, registry, usedNames, master,
  styleSource)` emits `<draw:page{attrs}>{frames}{notes}{anim}</draw:page>` with
  `attrs` = `preserved["@page"]`, `draw:name` uniqued. `markSlideEdited(slide,
  format)` sets `dirty` and, for odp, deletes `sourceXml`.
- `specs/082-slide-objects-images-shapes-text-boxes.md` owns frame commands and
  selection; `frameAccessibleName(frame)` (`useFrameManipulation.ts`) names
  frames for the report, and `filmstripThumbScale(widthCm, thumbWidthPx)`
  (`src/components/presentation/slideRender.ts`) is the thumbnail scale the
  filmstrip already uses — reused here with `SlideStage.vue`, not re-invented.
- Equality: `equalBytes(a: Uint8Array, b: Uint8Array): boolean` from
  `src/lib/core/equality.ts` (part-level after `unzip`, never container bytes);
  `equalStructural(a: unknown, b: unknown)` from `src/lib/equality.ts`, whose
  `IGNORED_KEYS` drops `"id"` — hence `SlideLayout.key`, not `.id`.
  `equalSemantic` is typed for `DocumentModel` and unusable here.

### The layout model (normative — types in `model.ts`, rules in `layouts.ts`)
```ts
export type SlideLayoutKind =
  | "title" | "titleAndBody" | "sectionHeader"
  | "twoContent" | "titleOnly" | "blank" | "other";
export interface LayoutSlot {
  role: PlaceholderRole;
  phType?: string; phIdx?: number;  // pptx p:ph @type/@idx, written back verbatim
  xCm: number; yCm: number; widthCm: number; heightCm: number;
}
export interface SlideLayout {
  /** Deterministic, deck-stable, survives save → reopen. pptx: the layout part;
   *  odp: `${master}::${pageLayout}`; synthesised: `standard:{kind}`. Not `id`. */
  key: string;
  kind: SlideLayoutKind;
  name: string;              // the deck's own name, shown on the card
  slots: LayoutSlot[];       // [] for "blank"
  masterName?: string;       // odp draw:master-page-name / pptx master part
}
```
`PresentationModel.layouts?: SlideLayout[]` in package order (every master's
layouts flattened, de-duplicated by `key`); `Slide.layoutKey?: string`. Both are
plain serialisable data per SPEC-079's rule, so `JSON.parse(JSON.stringify(
model))` still `toEqual`s the model. A `layoutKey` naming no catalogue entry
resolves to `undefined` and is left alone on save — never repaired, never dropped.
`STANDARD_LAYOUTS` fractions `(x, y, w, h)`, equal to `slides.ts`'s
existing placeholder fractions where the kinds overlap: **title** — `ctrTitle`
(0.08, 0.32, 0.84, 0.18), `subTitle` (0.08, 0.52, 0.84, 0.12); **titleAndBody** —
title (0.08, 0.08, 0.84, 0.16), body `idx 1` (0.08, 0.30, 0.84, 0.60);
**sectionHeader** — title (0.08, 0.42, 0.84, 0.20), body `idx 1` (0.08, 0.64,
0.84, 0.12); **twoContent** — title (0.08, 0.08, 0.84, 0.16), body `idx 1` (0.08,
0.30, 0.41, 0.60), body `idx 2` (0.51, 0.30, 0.41, 0.60); **titleOnly** — title
(0.08, 0.08, 0.84, 0.16); **blank** — none. Rounded by the private `r3`.

### Reading the source layout (both formats)
**pptx — confirmed in-repo.** `resolveChain(slidePath, slideRels, deps)`
(`pptx/reader.ts:1074`) already walks slide `.rels` `…/relationships/slideLayout`
→ layout part → its `.rels` `…/slideMaster`; this spec keeps the parsed
`layoutRoot` instead of discarding it. `name` = `p:cSld/@name`; `kind` =
`p:sldLayout/@type` (`title`→title, `obj`/`tx`→titleAndBody,
`secHead`→sectionHeader, `twoObj`/`twoTxTwoObj`→twoContent,
`titleOnly`→titleOnly, `blank`→blank, anything else→other; no `type` classifies
from the slot roles); `slots` = the existing
`buildPlaceholderMap(layoutSpTree(layoutRoot))` walk with `parseXfrm` geometry
through `core/units.ts#emuToCm`. `writer.ts:758` (`<p:sldLayout … type="blank"
preserve="1"><p:cSld name="Blank">`) proves both attribute positions in-repo.
**odp — partly unconfirmed, golden-first.** `draw:page` carries
`presentation:presentation-page-layout-name`, naming a
`<style:presentation-page-layout style:name="…">` in `styles.xml`'s
`<office:styles>` whose children are *believed* to be
`<presentation:placeholder presentation:object="title|outline|subtitle|text|
graphic|object|notes" svg:x svg:y svg:width svg:height/>`, mapped by the
existing `roleOf(cls)` (`odp/reader.ts:286`). **That child element name, the
`presentation:object` vocabulary and whether Impress always emits the reference
are NOT confirmed from a real file in this repo.** Use SPEC-086's golden-first
process: capture the real markup once into
`src/lib/presentation/odp/layoutGolden.ts` as `LAYOUT_GOLDEN` + the
`SlideLayout[]` it encodes, assert reader and splice against it, treat those
goldens as normative bytes, and record cross-app Impress verification in the PR,
not CI. A page with no reference (common) classifies its own `draw:frame
presentation:class` set into a synthetic entry keyed `` `${master}::` ``.

### Applying a layout, and Reset slide
Matching is three-pass, first hit per slot: exact `phKey(phType, phIdx)`; then
the first unconsumed frame of the same `role`; then the first unconsumed frame
in role order title→subtitle→body→other. A matched frame keeps `paragraphs`,
`fill`, `stroke` and `rotationDeg`, taking the slot's geometry, `phType` and
`phIdx`. Slots with no frame mint an empty `TextFrame` in slot order, spliced
**behind** existing non-placeholder frames so a picture is never covered.
Unmatched placeholders and every frame without a `role` keep their geometry and
their place in `frames`. `LayoutApplyReport = { applied: boolean; placed: number;
created: number; keptInPlace: string[] }` (frame ids). `resetSlideLayout` is the
same walk with only the exact-`phKey` pass and nothing created — the word
processor's `resetParagraphStyle` (`src/lib/commands.styles.ts:163`) one level
up. Neither touches `notes`, `transition`, `animations`, `background` or
`preserved` beyond the odp rewrite below.

### Writer rules (deltas to SPEC-080's rebuild)
No new writer path: both formats keep their verbatim/substitution selection, and
a deck whose layouts were untouched takes exactly the bytes it takes today. pptx
— the only newly substituted part is a changed slide's own `.rels` (one
relationship `Target` rewritten, every other relationship re-emitted with its id
and target unchanged), plus, when a missing layout is authored, the new layout
part, its `.rels`, the `[Content_Types].xml` override and the master with one
`p:sldLayoutId` spliced into `p:sldLayoutIdLst` (copy `linkImageRels`'s
append-only discipline). odp — the only new substituted part is `styles.xml`, and
only when a page layout is authored; the page's own change lives inside
`preserved["@page"]`, leaving `odpContentEnvelope`'s splice-equality invariant
(`prefix + pages.join("") + suffix`) intact for every other page. That slot is
this file's to introduce; SPEC-090 and SPEC-094 extend it, never re-declare it.
`stripRetention` (cross-format "Download a copy as…") drops `preserved` but must
**not** strip `layouts`/`layoutKey`.

### UI
`src/views/presentation/useSlideLayouts.ts`, deps `{ model, currentSlide,
currentIndex, record, commit, addSlide }` (the last being `useSlideCommands.add`),
returning `{ open, mode, openForChange, openForNew, close, cards, currentKey,
choose, apply, resetSlide, canReset, report, dismissReport }`, `cards` a
fresh-array computed (`(model.layouts ?? []).slice()` plus the synthesised
fallbacks) per the `useAnimations` reactivity rule.
`PresentationView.vue` gains **one** `useSlideLayouts({ … })` call, one
`<LayoutGalleryDialog>`, one toolbar prop/emit pair and the two filmstrip menu
entries, paid for by extracting the `.stage-workspace` block into
`components/presentation/SlideStageArea.vue`, or an equivalent extraction where a
sibling spec has already landed that one; `PresentationView.size.spec.ts`'s
`LINE_BUDGET` drops by at least 10 from whatever it stands at when this lands
(550 → 540 today), with a ledger row appended — the budget never rises.
`LayoutGalleryDialog.vue` declares exactly `{ controller: { type: Object,
required: true }, slideNumber: { type: Number, default: 1 } }`, emits nothing,
aliases `const c = props.controller;`, and keeps typed transforms in a sibling
`layoutCards.ts` (espree lints `.vue` — runtime `defineProps` only).

### Tests
- `src/lib/presentation/layouts.spec.ts` — every pptx `type` plus the exotic
  fallback, slot maths, the three-pass match (an exact-`idx` collision and a
  surplus placeholder included), `resetSlideLayout`, `addSlide` with and without
  a key. Reader specs cover catalogue derivation, `layoutKey` assignment and the
  no-reference fallback; `odp/layoutGolden.ts` is asserted both ways.
- `roundTrip.spec.ts` gains layout round trips — `equalBytes` per untouched part,
  the changed `.rels` diff confined to one `Target`, `styles.xml` verbatim unless
  authored, byte stability on a second write; existing assertions neither
  re-expressed nor weakened.
- `layoutsRoundTrip.spec.ts`, `layoutsFixtures.spec.ts`, `useSlideLayouts.spec.ts`,
  `LayoutGalleryDialog.spec.ts` and `e2e/SPEC-089.spec.ts`.

### Specs this changes
- SPEC-080's out of scope loses "new slides are title + body only, no layout
  picker": `addSlide` gains its optional third argument and `buildFromScratch`
  five more layout parts per format; its rebuild rules are otherwise unchanged.
- SPEC-079's "inheritance is resolved and discarded" note gains its exception:
  the chain is retained as a catalogue, frames' effective geometry unchanged.
  SPEC-081's serialiser slots are unchanged — only a slide's `.rels` and the odp
  `preserved["@page"]` attribute string move.
- `specs/000-roadmap.md` gains this spec's phase row; nothing else changes.
