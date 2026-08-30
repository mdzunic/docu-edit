---
id: SPEC-094
title: Deck themes — colour and font schemes
status: draft
area: formatting
complexity: complex
story_points: 13
depends_on: [SPEC-077, SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-089]
---
## Why
SPEC-079 reads the theme and throws the link away. Its resolver
(`presentation/pptx/color.ts`) walks a slide's `a:schemeClr` through the
master's `p:clrMap` into the theme's `a:clrScheme`, applies
`lumMod`/`lumOff`/`shade`/`tint`, and stores the answer — a frozen
`#rrggbb`; `themeMinorLatin` does the same to fonts; the theme part
survives only as an `OpaquePart`. So a deck built on accent1 reopens as a
dozen unrelated blues, editing one run rewrites it `a:srgbClr`, and there
is no way to change a deck's look at all.

`specs/077-document-themes.md` owns the shared theme concept; this file is
its presentation counterpart and reuses that model rather than forking it
— twelve colours, two fonts, the same `ModelColor` widening and single
resolution point — gathered in `src/lib/core/theme.ts`, the only place
`lib/presentation`'s boundary rule can reach `DocumentTheme`. **This file is the
normative source for the deck theme model, the theme-reference wire
mappings in both formats, and the picker's storage rule.** That last rule
is the point: a theme swatch must serialise as a scheme reference, not the
hex it resolves to, or the theme is decoration.

## Acceptance criteria
- [ ] `src/lib/core/theme.ts` (new, product-neutral) holds the shared model per "The shared theme model": `DocumentTheme`, `ThemeColorRef`, `ModelColor`, `ModelFont`, `resolveColor`, `resolveFont`, `isThemeColor`, `DEFAULT_THEME` (Office — dk1 `#000000`, lt1 `#ffffff`, dk2 `#44546a`, lt2 `#e7e6e6`, accent1–6 `#4472c4`/`#ed7d31`/`#a5a5a5`/`#ffc000`/`#5b9bd5`/`#70ad47`, hlink `#0563c1`, folHlink `#954f72`, Calibri Light/Calibri) and `BUILT_IN_THEMES` (6, Office first), each pinned by unit test
- [ ] The `shade`/`tint`/`lumMod`/`lumOff` maths moves out of `presentation/pptx/color.ts` into `core/theme.ts` two-phase per "Enabling move": phase one hoists `applyTransform` plus the sRGB↔HSL helpers and `pptx/color.ts` re-imports them so `pptx/color.spec.ts` passes unchanged; phase two builds `resolveColor` on them — coverage preserved, never weakened
- [ ] `PresentationModel` gains `theme?: DeckTheme` per the Reference (`DocumentTheme` plus `colorMap`, `partName`, `dirty`); absent means `DEFAULT_THEME` and every colour on that deck stays literal — nothing is retroactively converted to a reference
- [ ] `SlideRun.color`, `SlideRun.fontFamily`, `SlideFill.color` and `FrameStroke.color` **widen** to `ModelColor`/`ModelFont`; the string form stays valid, so no stored model migrates and no shipped test asserting a literal breaks; every renderer, serialiser, filmstrip thumbnail, playback style and contrast check reaches a literal only through `resolveColor`/`resolveFont`
- [ ] The pptx reader keeps SPEC-079's resolution as the normative *effective value* and additionally records the reference: `a:schemeClr` becomes `{ theme: slot, lumMod?, lumOff?, shade?, tint? }` with transforms as 0..1 fractions in source order (any other transform → the value stays a literal); `a:latin@typeface` `+mj-lt`/`+mn-lt` and the inherited-default case become `{ theme: "major" | "minor" }` instead of the baked `themeMinorLatin` string
- [ ] The pptx writer serialises a `ThemeColorRef` as `<a:schemeClr val="{slide-space name}">` plus `<a:lumMod val="60000"/>`-style children in the canonical order lumMod, lumOff, shade, tint — the slide-space name being the `p:clrMap` attribute whose value is the ref's slot (`dk1`→`tx1` under the default map), falling back to the slot name when nothing maps to it; a literal still writes `a:srgbClr`; a font ref writes `<a:latin typeface="+mj-lt"/>`; asserted on written XML in `pptx/serialize.spec.ts`
- [ ] A slide carrying `preserved["p:clrMapOvr"]` writes its refs through that override's mapping; when the override cannot be parsed that slide's refs write as `a:srgbClr` literals so the rendered colour never lies, and the panel says the slide is opted out
- [ ] Editing the theme rewrites `ppt/theme/themeN.xml` **in place** — only `a:clrScheme` and `a:fontScheme` are spliced, while `a:fmtScheme`, `a:effectStyleLst`, `a:bgFillStyleLst`, `a:objectDefaults` and `a:extraClrSchemeLst` survive byte-verbatim, per SPEC-008's preserve-don't-regenerate rule
- [ ] Writer path selection extends SPEC-080: `writePptx`'s verbatim branch additionally requires `model.theme?.dirty !== true`; applying a theme to a pptx deck dirties **no slide** (the slides' refs do not change), so every part except the theme part compares `equalBytes` (`src/lib/core/equality.ts`, part-level after `unzip`) with the source
- [ ] odp maps per "odp wire mapping": ODF has no first-class theme, so the scheme is written as named `de-theme-*` styles under `office:styles` in `styles.xml` and a ref is carried by `style:parent-style-name` on the automatic style beside the resolved literal; a ref carrying transforms writes the literal only and **loses its reference on re-open** — stated in the panel, not silent; a theme change re-serialises only pages whose styles are theme-parented, every other part staying `equalBytes`
- [ ] Applying a theme is one undo entry ("Apply theme"), and undoing it puts the old theme back: SPEC-080's `usePresentationUndo` `Snapshot` gains `theme?: DeckTheme`, cloned with `slides` and restored by `applySnapshot` (absent key deleted), so the deck cannot be left recoloured by an undo. Every scheme-bound colour and theme-font-bound run on every slide re-renders at once — stage, filmstrip, transition and animation preview layers — while every run, fill or stroke holding an explicit literal keeps exactly the value it had, asserted by a fixture slide mixing both
- [ ] The toolbar gains a "Design" toggle (`slides-theme-toggle`, `aria-pressed`; preference `presentationThemeOpen` added to all three places in `src/lib/preferences.ts`, default `false`) opening `slides-theme-panel`: a `theme-gallery` of `theme-card` buttons — the deck's current theme first and marked `theme-current`, then the six built-ins, each drawing its own accents as swatches under its `theme-name` — plus `theme-apply`, and a Customise disclosure with twelve `theme-color-{slot}` inputs, `theme-font-major`, `theme-font-minor`, `theme-reset`; all keyboard-operable
- [ ] The pickers become theme-aware without disturbing what exists: `src/components/TextColorPicker.vue` gains an optional `theme` prop rendering a theme row **above** the standard palette — ten swatches (dk1/lt1/dk2/lt2 + accent1–6), each with five lumMod/lumOff variants, accessible names in the "Accent 1, lighter 40%" form, testid `text-color-theme-swatch` — with recents and custom unchanged below; `FrameStylePopover.vue` gains the same row (`style-theme-swatch`) above its `style-fill-color`/`style-stroke-color` inputs, which keep their testids and behaviour; a theme swatch stores the `ThemeColorRef`, anything else stores a literal
- [ ] Round trip: a deck whose runs, fills and strokes mix plain refs, transformed refs, theme fonts and explicit literals → save → reopen reproduces `model.theme` and every ref exactly in pptx (`equalStructural`, `src/lib/equality.ts`, over the theme and the frames), and in odp reproduces the untransformed refs and the theme while transformed refs arrive as their literals — the one sanctioned odp delta, pinned by the test; a no-edit download of either stays byte-identical part-for-part
- [ ] Fixture coverage: `fixtures/build.ts` gains a `spec094-themes` builder pair appended under a banner (never mutating the four mini-deck slides) with a slide mixing an accent1 fill, an accent1+lumMod 60%/lumOff 40% shape, a `+mn-lt` run, a `+mj-lt` title and an explicitly-coloured run, plus a slide carrying a `p:clrMapOvr` (pptx) / an unrecognised extension theme block (odp) to pin the preserve path; both writes added to `scripts/build-presentation-fixtures.ts`
- [ ] e2e (`e2e/SPEC-094.spec.ts`): open `spec094-themes.pptx` → Design → the current theme card is marked → apply a built-in → the accent-filled shape and the themed title change colour while the explicitly-coloured run does not → open the style popover, pick the accent1 theme swatch for a shape fill → download, reopen → the shape still tracks accent1 (apply a second theme and it moves again), the literal run is unmoved, and the panel names the applied theme

## Out of scope
Theme *effects* (`a:fmtScheme`, `a:effectStyleLst`, `a:lnStyleLst`,
`a:bgFillStyleLst`) — preserved, never modelled; background styles and
`p:bgRef` (see below); slide-master and layout authoring, layout
backgrounds and placeholder inheritance, owned by
`specs/089-slide-layouts-and-masters.md`; per-slide theme variants beyond
honouring an existing `p:clrMapOvr`; multiple theme parts per deck (all
kept, only the first master's editable); importing a theme from another
deck or a `.thmx`; retroactively converting literals into references; the
*application*'s theme (SPEC-058/059).

## Reference

### Dependencies restated
- `specs/077-document-themes.md` owns the shared concept — `DocumentTheme`
  (twelve colours + `fonts.major`/`minor`), the `ModelColor` widening and
  the single-resolution-point discipline, reused verbatim. Its
  `ModelColor`/`resolveColor`/`isThemeColor` already sit in the importable
  `core/colors.ts`; `DocumentTheme` sits in `documentModel.ts`, out of
  `lib/presentation`'s reach (`presentation/boundary.spec.ts` resolves only
  under `src/lib/core`), so it moves to `src/lib/core/theme.ts`, the colour
  helpers joining it for cohesion — whichever spec lands first creates it.
- `specs/079-slide-canvas-rendering.md` — **normative** colour resolution:
  `presentation/pptx/color.ts` exports `ColorContext { scheme; map }`,
  `parseClrScheme(themeRoot)`, `parseClrMap(masterRoot)`,
  `resolveColorElement(el, ctx)`, `resolveFirstColor(container, ctx)`;
  `a:schemeClr@val` → `ctx.map[val] ?? val` → `ctx.scheme[…]`, then transforms
  in order with `val = attribute/100000` (`shade` c·val, `tint` c·val+(1−val),
  `lumMod`/`lumOff` via HSL) — the literal this spec keeps the ref beside.
- `specs/080-create-presentation-and-slide-management.md` — writer path
  selection (`writePptx`: no `structureDirty`/`dirty`/`notesDirty` → every
  part re-zipped verbatim, else a substitution build; `writeOdp` rebuilds
  `content.xml` from `odpContentEnvelope` when structural) and undo
  (`record(label)`, mutate, `commit()`).
- `specs/081-slide-text-editing.md` — `runXml` (private, one per format, in
  `pptx/serialize.ts` and `odp/serialize.ts`) writes a run's colour as
  `<a:solidFill><a:srgbClr val="{HEX}"/></a:solidFill>` and its font as
  `<a:latin typeface="…"/>`; `serializeSlideXml`'s child order is fixed,
  every slot `modelled ?? preserved ?? ""`, never both; `markSlideEdited` sets
  `dirty` and for odp deletes `sourceXml`; odp automatic styles come from
  `StyleRegistry`, writing `fo:color`/`style:font-name`/`draw:fill-color`.
- `specs/082-slide-objects-images-shapes-text-boxes.md` —
  `FrameStylePopover.vue` (`frame-style-popover`, `style-fill-color`,
  `style-no-fill`, `style-stroke-toggle`, `style-stroke-color`,
  `style-stroke-width`), `setFrameFill`/`setFrameStroke` in
  `presentation/frames.ts`, deck defaults `#4472c4` / `#2f528f` 1pt — plain
  SPEC-051 literals, no theme derivation, literal until a theme swatch is picked.
- `specs/089-slide-layouts-and-masters.md` owns the slide → layout → master
  chain; this file reads only the theme part at its end (until it lands, via
  SPEC-079's private `resolveChain(...)`, which gains the theme capture).

### The shared theme model
```ts
// src/lib/core/theme.ts — product-neutral, imported by both products.
export type ThemeColorSlot = "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2"
  | "accent3" | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink";
export interface DocumentTheme { name: string;
  colors: Record<ThemeColorSlot, string>;      // lowercase #rrggbb
  fonts: { major: string; minor: string } }    // headings, body
/** A scheme reference plus transforms, each a 0..1 fraction (val/100000). */
export interface ThemeColorRef { theme: ThemeColorSlot;
  lumMod?: number; lumOff?: number; shade?: number; tint?: number }
export type ModelColor = string | ThemeColorRef;   // widens every colour field
export type ModelFont = string | { theme: "major" | "minor" };
/** The single resolution point; nothing else interprets a ModelColor. */
export function resolveColor(c: ModelColor, theme: DocumentTheme): string;
export function resolveFont(f: ModelFont, theme: DocumentTheme): string;
export const DEFAULT_THEME: DocumentTheme;              // "Office"
export const BUILT_IN_THEMES: DocumentTheme[];          // 6, Office first
// presentation/model.ts — PresentationModel gains `theme?: DeckTheme`;
// absent = unthemed (DEFAULT_THEME, colours literal).
export interface DeckTheme extends DocumentTheme {
  /** pptx `p:clrMap` from the first master: "bg1"|"tx1"|"bg2"|"tx2"|
   *  "accent1".."folHlink" → scheme slot. Absent = identity. */
  colorMap?: Record<string, string>;
  partName?: string;   // "ppt/theme/theme1.xml" | "styles.xml"
  dirty?: true }       // edited → the writers rewrite that part
```
Plain serialisable data, untouched by `equalStructural`'s `IGNORED_KEYS`.

**Enabling move (two-phase).** `presentation/pptx/color.ts` privately owns
`applyTransform`, `hexToRgb`, `rgbToHex`, `rgbToHsl`, `hslToRgb`, `clamp01`
— pure, DOM-free, exactly what `resolveColor` needs. Phase one hoists them
into `core/theme.ts` with `pptx/color.ts` re-importing, leaving
`pptx/color.spec.ts` untouched; phase two defines `resolveColor` over them.

### pptx wire mapping
Read: the theme part reached through master → `THEME_RELATIONSHIP` fills
`DeckTheme` from `a:themeElements/a:clrScheme` (each child's inner
`a:srgbClr@val` or `a:sysClr@lastClr` — `dk1`/`lt1` are usually `a:sysClr`,
and the source form is remembered) and
`a:fontScheme/a:majorFont|a:minorFont/a:latin@typeface`; `colorMap` from
the master's `p:clrMap`; `partName` from the resolved path. No scheme →
`model.theme` stays absent. Write: the part is an `OpaquePart` re-emitted
verbatim unless `theme.dirty`, in which case those two elements are spliced
by string substitution — SPEC-082's Content-Types/manifest discipline — so
the untouched remainder stays bytes. `dk1`/`lt1` keep the source's form
(`a:sysClr` keeps `val`, `lastClr` updated); ours writes `a:srgbClr`.

**The `p:bgRef` question.** `Slide.background`'s `color` widens like any
other, so a background resolved from `p:bg/p:bgPr/a:solidFill/a:schemeClr` now
tracks the theme; a `p:bg/p:bgRef` into `a:bgFillStyleLst` stays unmodelled in
`preserved["p:bg"]`, re-emitted verbatim and moving only with the theme part
it points into. Background styles as an authoring surface belong to SPEC-089.

### odp wire mapping
ODF has no theme part, and pretending otherwise would be the fidelity lie
this repo's first law forbids. The honest mapping is:
- **The scheme** goes into `styles.xml`'s `office:styles` as a named block —
  `<style:style style:name="de-theme-accent1" style:family="graphic"
  style:display-name="Accent 1"><style:graphic-properties draw:fill="solid"
  draw:fill-color="#4472c4"/></style:style>` per slot, plus
  `de-theme-major`/`de-theme-minor` (`style:family="text"`): valid ODF,
  ignorable by Impress, read back by name.
- **A reference on content** rides on
  `style:parent-style-name="de-theme-{slot}"` on the automatic style
  `StyleRegistry` mints, beside the resolved literal so other readers see
  the right colour; `StyleResolver` already walks parents.
- **The stated loss**: a transformed ref has no faithful parent, so it writes
  the literal alone and re-opens as one, and a font ref resolves to a family
  name — the panel says so once for an odp deck.
- **The emission route**: `styles.xml` reaches the package through the
  substitution slot SPEC-089 introduces in `writeOdp` (SPEC-090 extends it
  for page size); a theme edit routes the rewritten part through that same
  slot, and an untouched deck re-emits `styles.xml` byte-verbatim.
- **Preserve**: LibreOffice writes an extension theme block whose exact
  shape (namespace, element and attribute names) **this spec does not claim
  to know** and never authors — it rides along inside the verbatim
  `styles.xml`, dropped with a stated loss on edit.

### Golden-first implementation note (normative process)
The discipline SPEC-086 uses for `p:timing`: the theme splice and the odp
`de-theme-*` block are pinned by checked-in canonical strings —
`presentation/pptx/themeGolden.ts` (`THEME_GOLDEN`, `GOLDEN_THEME`) and
`presentation/odp/themeGolden.ts` (`ODP_THEME_GOLDEN`, re-exporting that
`GOLDEN_THEME`) — hand-authored against ECMA-376 §20.1.6 and ODF 1.3 §16,
not screen-scraped. Tests assert `writeTheme(GOLDEN_THEME) === THEME_GOLDEN`;
cross-app checking is manual QA recorded in the PR.

### View wiring
`src/views/presentation/useDeckTheme.ts` — `DeckThemeDeps { model:
ShallowRef<PresentationModel | null>; record: (label?: string) => void;
commit: () => void }` → `{ open, toggle, theme, themeName, isCustom, gallery, currentThemeId,
previewId, hoverPreview, clearPreview, apply, setSchemeColor, setFont,
reset, lossNote }`, `open` being `computed(() =>
preferences.presentationThemeOpen)` per SPEC-085's panel pattern.
`src/components/presentation/ThemesPanel.vue` takes exactly `{ controller: {
type: Object, required: true } }`, emits nothing, and uses runtime `defineProps`
(espree bars the generic form). `apply()` is `record("Apply theme")` → set
`model.theme` (+ `dirty`; + `markSlideEdited` per affected slide, odp only) →
`commit()`, its `Snapshot` carrying `theme`. `PresentationView.vue` gains one
`useDeckTheme({ model: presentation, record: slideUndo.record, commit })`
call, one prop pair on `<PresentationToolbar>` and one
`<ThemesPanel :controller="theme" />` below the workspace, paid for by
**this file's** extraction of `stageScale`/`recomputeScale`/the
`ResizeObserver`/the `resize` listener into
`src/views/presentation/useStageFit.ts` — the one owner of that module, so a
sibling naming it falls back to an equivalent extraction — after which
`PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 10 from
whatever it stands at when this lands (550 → 540 today) with a ledger row
— never up. The application theme (SPEC-058/059) stays
independent: CSS properties in `docu-edit:preferences` that never reach a
saved file, where the deck theme always does.

### Tests
- `src/lib/core/theme.spec.ts` — the twelve slots, `resolveColor` per
  transform against known PowerPoint values, `resolveFont`, each built-in's
  contrast; `pptx/color.spec.ts` unchanged, proving the hoist
  behaviour-preserving; `presentation/theme.spec.ts` — apply/reset, which
  slides dirty per format, refs untouched by apply.
- `pptx/theme.spec.ts`, `odp/theme.spec.ts` — golden equality both ways,
  clrMap reverse mapping incl. the `p:clrMapOvr` opt-out, the splice leaving
  `a:fmtScheme` byte-identical; `pptx`/`odp` `serialize.spec.ts` —
  `a:schemeClr`/`+mn-lt` emission and the `de-theme-*` parent link.
- `themeRoundTrip.spec.ts` — the mixed deck in both formats:
  `equalStructural` over the modelled subset, `equalBytes`
  (`core/equality.ts`) per untouched part after `unzip`.
- `useDeckTheme.spec.ts`, `TextColorPicker.spec.ts` (theme row, swatch
  names, ref-not-hex storage), `PresentationView.size.spec.ts`, `e2e/SPEC-094.spec.ts`.

### Specs this changes
- `specs/079-slide-canvas-rendering.md` — the reader's baked `themeMinorLatin`
  font and scheme colours become references; the maths is unchanged and its
  tests re-expressed against the same values, never weakened.
- `specs/080-create-presentation-and-slide-management.md` — the verbatim
  writer branch gains the `theme?.dirty !== true` term and its undo `Snapshot`
  gains `theme`; `specs/081-slide-text-editing.md` — run serialisation emits
  `a:schemeClr`/`+mn-lt` for refs, and `StyleRegistry` gains the
  `style:parent-style-name` link.
- `specs/082-slide-objects-images-shapes-text-boxes.md` — its "theme-aware
  fill palettes" exclusion is resolved here; `specs/077-document-themes.md`
  — its model moves to `src/lib/core/theme.ts`, contract otherwise unchanged;
  `specs/000-roadmap.md` gains this spec's phase row.
