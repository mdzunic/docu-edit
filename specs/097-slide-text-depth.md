---
id: SPEC-097
title: Slide text depth — lists, spacing, autofit and links
status: draft
area: formatting
depends_on: [SPEC-014, SPEC-018, SPEC-025, SPEC-032, SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-084]
---
## Why
SPEC-081's out of scope is a list of everything a real slide's text does that this
editor cannot: "Numbered-bullet authoring (the 'number' bullet renders read-only; the
toggle writes 'disc' only)… Text autofit; overflowing text keeps overflowing exactly as
SPEC-079 renders it. Line spacing, letter spacing, highlight, links." SPEC-079 adds
columns in a text body and vertical text. So a deck's bullets cannot be numbered, its
title spills off the slide with no way to shrink it, and its "click here" is dead ink.

Worse, the gap is destructive. `pPrXml` in `src/lib/presentation/pptx/serialize.ts`
rebuilds `a:pPr` from `algn`/`lvl`/one bullet element and `txBodyXml` emits a bare
`<a:bodyPr/>`, so the first edit to a slide silently discards its source `a:lnSpc`,
`a:spcBef`, `a:buFont`, `marL`/`indent`, insets, anchor and autofit — a first-law
violation hiding behind an out-of-scope line. This spec models that vocabulary and adds
preservation slots for the rest of it. The reuse is nearly total: `src/lib/core/lists.ts`
was written product-neutral for exactly this ("a Slides product numbers the list items
inside a text box with the same marker math") and SPEC-025's URL-safety helpers are
neutral logic in a text-model module. **This file is the normative source for the slide
list, paragraph spacing, text-body and hyperlink models, their
`a:pPr`/`a:bodyPr`/`a:hlinkClick` and ODF counterparts, and the rule that autofit
measurement never dirties a slide.**

## Acceptance criteria
- [ ] `SlideParagraph` gains the list delta: `bulletStyle?: ListMarkerStyle` (from `src/lib/core/lists.ts`; absent = "disc" for `bullet: "disc"`, "decimal" for `bullet: "number"`, and `markerType(bulletStyle)` must agree with `bullet`), `bulletChar?: string` (one grapheme, `bullet: "disc"` only, wins over `bulletStyle`), `bulletImageId?: string` (a `model.media` `MediaPart`, wins over both — deliberately not named `id`, per `equalStructural`'s `IGNORED_KEYS`), `startAt?: number` (1–32767, absent = 1/continue), `indentCm?: number` (absent = `(level + 1) × LIST_INDENT_STEP_CM`, 1.27) and `hangingCm?: number` (absent = `LIST_MARKER_WIDTH_CM`, 0.635); `bullet`/`level` keep their SPEC-079 meaning and range 0–8
- [ ] `SlideParagraph` also gains `lineSpacing?: number` (multiple, 0.25–5, absent = 1), `lineSpacingPt?: number` (4–1000, wins over `lineSpacing`), `lineSpacingRule?: "exact"` (absent = "at least" — the same three fields and the same absent-value meaning `ParagraphFormatting` carries in `src/lib/documentModel.ts:444-457`), `spaceBeforePt?: number` and `spaceAfterPt?: number` (0–500, absent = 0); `SlideRun` gains `letterSpacingPt?: number` (−10…10, 0.1 steps, absent = 0) and both `marksEqual` implementations — `src/lib/presentation/text.ts:66` and `src/components/presentation/slideTextDom.ts` — compare it so run coalescing still works
- [ ] `TextFrame` and `ShapeFrame` gain `body?: TextBodyProps` — declared in `src/lib/presentation/model.ts` beside `SlideTransition`/`SlideAnimation`, never in `textBody.ts` (the Reference says why): `autofit?: "none" | "shrink" | "resizeShape"` (absent = "none"), `fontScalePct?: number` (10–100) and `lineScalePct?: number` (80–100) valid only with "shrink", `anchor?: "top" | "middle" | "bottom"` (absent = "top" on a `TextFrame`, "middle" on a `ShapeFrame` — today's renderer assumption becomes a stored default), `insetsCm?: { left; top; right; bottom }` (absent = 0.254/0.127/0.254/0.127, the pptx EMU defaults), `columns?: { count: number; gapCm: number }` (count 1–16, absent = 1) and `wrap?: false`
- [ ] Nothing unmodelled is destroyed: `SlideParagraph` gains `pptxPPr?: string` / `odpParaProps?: string` and the two body-carrying frames gain `pptxBodyPr?: string` / `odpBodyProps?: string` holding the source's unmodelled `a:pPr`/`a:bodyPr` children and attributes (odp: the automatic style's property string), each **cleared by the command that edits the modelled property it shadows** — exactly `Table.docxTblPr`'s documented rule — and re-emitted in the schema order the Reference fixes; a paragraph or body this spec models fully re-serialises from the model
- [ ] `src/lib/presentation/text.ts` grows the commands: `ParagraphFormatPatch` widens with every new paragraph field (`null` clears, as `bullet: null` already does), `toggleBullet(frame, range, style: ListMarkerStyle = "disc")` takes an optional third argument, so the one production call site (`useSlideTextEditing.ts:400`) and the three in `text.spec.ts` (`:64`, `:66`, `:72`) are unchanged by the default, plus `setListStyle`, `setParagraphSpacing`, `setLineSpacing`, `setLetterSpacing` and `demote`/`promote` wrapping `setIndentLevel`; `SelectionFormatState` gains `bulletStyle`, `startAt`, `lineSpacing`, `spaceBeforePt`, `spaceAfterPt` and `letterSpacingPt` as `FieldState<…>`s (so the panel shows "mixed"); inside an editing session **Tab demotes and Shift+Tab promotes** every paragraph the selection touches (clamped 0–8, `event.preventDefault()`, never a literal tab character and never a focus move), handled in `useSlideTextEditing.ts#onKeydown` beside its Escape/Cmd-B branches — safe because `stageEdit.onGlobalKeydown`'s Tab frame-cycling already skips while editing — as one undo entry labelled "Indent"
- [ ] List authoring: the toolbar gains `slides-numbering` (numbered toggle, beside today's `slides-bullet`) and `slides-list-menu` opening `list-style-popover` — the eight `core/lists.ts` marker styles as `list-style-{style}` buttons previewed through `markerText(style, 1)`, "Custom bullet…" (`list-bullet-char`, choosing from `src/lib/core/specialCharacters.ts`), "Picture bullet…" (`list-bullet-image` → `list-bullet-image-input`, the `PRESENTATION_IMAGE_MIME_TYPES` set, minted into `model.media` by SPEC-082's naming rule), `list-start-at`, `list-indent` and `list-hanging`; all keyboard-operable and fitting the 720px e2e viewport through SPEC-081's overflow rule
- [ ] Rendered numbering is **derived, never stored**: `markedParagraphs` in `src/components/presentation/slideRender.ts` is rewritten over `core/lists.ts` (`markerText`, `formatOrdinal`, `clampListLevel`, slide level `L` ↔ core level `L + 1`), counting per level by adjacency — a non-list paragraph or a shallower level between two items restarts the deeper level, `startAt` seeds the run — and every surface agrees: stage, `SlideFilmstrip` thumbnails, `PresentationPlayView`, the build-preview layer and SPEC-088's export
- [ ] Rendered *indent* is derived from the same model, so the stage stops disagreeing with what both writers emit: `paragraphStyle` in that same file reads `indentCm`/`hangingCm` under AC-1's `(level + 1) × LIST_INDENT_STEP_CM` default (marker box `hangingCm`, text edge `indentCm`), and the private `LEVEL_INDENT_CM = 0.95` (`slideRender.ts:16`, used at `:171`) is deleted — an existing deck's rendered per-level indent therefore widens from 0.95 cm to the 1.27 cm `odp/serialize.ts` and `a:pPr@marL` already write, a deliberate correction pinned by `slideRender.spec.ts`, not a regression
- [ ] A new Text panel (`slides-text-toggle` → `slides-text-panel`, `<section class="text-panel">`, `aria-label="Text for slide {n}"`, max-height 45% with `overflow-y: auto`, open state persisted as `presentationTextOpen` added to all three places in `src/lib/preferences.ts` beside the other three keys, default `false`) edits the selected frame/paragraphs: `text-line-spacing-mode` (multiple/exactly/at least — `CustomSpacingDialog.vue`'s vocabulary, rebuilt as native panel controls because that dialog is bound to the text model), `text-line-spacing`, `text-space-before`, `text-space-after`, `text-letter-spacing`, `text-autofit`, `text-anchor`, `text-inset-left|-top|-right|-bottom`, `text-columns`, `text-column-gap`
- [ ] Autofit is real, and measured outside `src/lib`: `"shrink"` walks the deterministic ladder in `src/lib/presentation/textBody.ts#fitScale` (pure — it takes a `measureHeightCm(fontScalePct, lineScalePct)` callback and is unit-tested with a fake measurer) against the frame's inset-reduced height, the callback coming from the new DOM-side `src/components/presentation/slideTextFit.ts` beside `slideTextDom.ts`; `"resizeShape"` sets `heightCm` to the measured content height (floor `MIN_FRAME_SIZE_CM` 0.5cm, `r3` rounding); `"none"` overflows exactly as SPEC-079 renders it today
- [ ] Measuring **never dirties a slide**: on open the renderer honours the stored `fontScalePct`/`lineScalePct` and, if its own measurement disagrees, re-renders at the measured scale while leaving the model untouched — the value is written back only inside an edit that has already called `markSlideEdited`, or when the user picks a mode in `text-autofit`; a deck opened, rendered and downloaded with no edit is still byte-identical part-for-part (§"the six ways to break the byte-identical download", rule 1)
- [ ] Hyperlinks: the non-exported `SlideFrameBase` and `SlideRun` each gain the two scalars `linkHref?: string` and `linkSlide?: number` (1-based; at most one set — two scalars rather than one object for exactly the reason `documentModel.ts:51-58` gives, so run coalescing and `marksEqual` stay `===`-per-key), plus `linkTooltip?: string`; only text/image/shape frames serialise one (`isLinkableKind`, matching SPEC-086's `isAnimatableKind`) since group/foreign frames re-emit verbatim; every href passes `sanitizeHref`, which reaches `lib/presentation` only through the two-phase move of `urlScheme`/`isDangerousUrl`/`sanitizeHref`/`DANGEROUS_SCHEMES`/`detectBareUrl`/`AutolinkMatch` from `src/lib/links.ts` into `src/lib/core/urls.ts` per the Reference (phase one re-exports, so `LinkDialog.vue`, `useLinks.ts` and `links.spec.ts` pass unchanged and both `boundary.spec.ts` files stay green)
- [ ] Link authoring and playback: `slides-link` (Cmd/Ctrl+K in a text session, and on the insert bar for a whole selected frame) opens `slide-link-dialog` over `src/components/BaseDialog.vue` with `link-url`, `link-slide` (a `<select>` of the deck's slides by `slideOutline` title), `link-tooltip`, `link-apply`, `link-remove`, `link-cancel`; a linked run renders as `<a data-testid="slide-link">` underlined in the run's own colour (theme `hlink` belongs to SPEC-094 and is not invented here); on the editor stage a plain click still starts a text session and Cmd/Ctrl+click follows the link, while in `PresentationPlayView.vue` a click on a link element navigates — external through `window.open(href, "_blank", "noopener,noreferrer")`, in-deck through `jumpTo(slideIndex - 1)` — and `onStageClick` must **not** also advance (the handler returns early when `event.target.closest("[data-testid='slide-link']")` matches)
- [ ] Round trip, both formats: a deck given numbered and picture bullets with a start-at, 1.5× line spacing, 6pt space-before, 0.5pt letter spacing, shrink autofit, middle anchor, custom insets, two columns and both link kinds → save → reopen reproduces every modelled field, compared field-level with `toEqual` plus `equalStructural` (`src/lib/equality.ts` — whose `XML_STRING_KEYS` does not list `pptxPPr`/`pptxBodyPr`/`odpParaProps`/`odpBodyProps`, so those compare raw, not canonically; `equalSemantic` is `DocumentModel`-typed and unusable here); every part the edit did not touch is `equalBytes` (`src/lib/core/equality.ts`) after `unzip`, a no-edit download is byte-identical, writing the reopened model twice yields identical slide-part bytes (SPEC-081's byte-stability rule), and pptx's `fontScale`/`lnSpcReduction` quantisation to 1000ths of a percent is the one sanctioned delta, pinned by the test
- [ ] Fixtures: `src/lib/presentation/fixtures/build.ts` gains `buildPptxTextDepthFixtureBytes`/`buildOdpTextDepthFixtureBytes` and `EXPECTED_TEXT_DEPTH` under a new banner (the four mini-deck slides are never mutated), each deck carrying one deliberately **out-of-subset** body (a `a:bodyPr vert="vert270"` / an ODF vertical-text frame) plus an unmodelled `a:pPr` child to pin the preserve-don't-parse path; `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec097-text-depth.pptx` and `.odp` and its log line names them
- [ ] `PresentationView.vue` gains exactly one `useSlideTextDepth({ model, currentSlide, currentIndex, record, commit, selection: stageEdit.selection, editing: stageEdit.editing })` wiring call plus the `<TextPanel :controller="…" :slide-number="currentIndex + 1" />` mount and the toolbar prop pair, paid for by extracting the `<PresentationToolbar>` prop/emit block into `src/components/presentation/SlideToolbarHost.vue`; `PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 10 from whatever it stands at when this lands (550 → 540 today) with a ledger row appended — the budget never rises
- [ ] e2e (`e2e/SPEC-097.spec.ts`): open `spec097-text-depth.pptx` → the numbered body renders "1." … "3." with the fixture's start-at honoured → select the body, Tab demotes line 3 and its marker becomes "a." → Text panel → line spacing "exactly 24pt", space-after 6pt → the title overflows, set `text-autofit` to "Shrink text on overflow" and it fits inside the frame → select a word, Cmd+K, link it to slide 3 → Present → clicking the link jumps to slide 3 without also advancing a build → exit, download, reopen → markers, spacing, autofit scale and the link all come back, and the deck's vertical-text frame is still preserved verbatim

## Out of scope
Vertical and RTL text, tab stops inside slide text, `spcFirstLastPara`/`compatLnSpc`,
text warps and WordArt, 3-D text (all preserved through the new slots, never authored).
Bullet colour, size percentage and bullet font (`a:buClr`/`a:buSzPct`/`a:buFont` are
preserved in `pptxPPr`, not edited). Highlight and strikethrough on slide runs.
Per-paragraph text builds — SPEC-086 owns the animation model and still targets whole
frames. Autofit on table cells (SPEC-091 excludes it) and on notes. Theme-driven link
colours (SPEC-094's `hlink`/`folHlink`). Autolink-as-you-type on slides (`detectBareUrl`
moves to core for reuse; nothing calls it here). Link targets beyond this deck's slides
and an absolute URL — no custom shows, OLE action verbs, macro or media actions. Column
rules and balancing (count and gap only).

## Reference

### Dependencies restated
- `specs/079-slide-canvas-rendering.md` owns the slide model: `SlideParagraph { runs;
  align?: "left"|"center"|"right"|"justify"; bullet?: "disc"|"number"; level? /* 0..8 */
  }`, `SlideRun { text; bold?; italic?; underline?; fontFamily?; sizePt?; color? }`
  (effective values; absent = 18pt/`#000000`), `TextFrame { kind:"text"; role?; phType?;
  phIdx?; paragraphs; fill?; stroke? }`, `ShapeFrame`'s optional `paragraphs`,
  cm/pt/`#rrggbb`, `preserved`. Its out-of-scope line claims `normAutofit` font scaling
  "is applied when declared with an explicit `fontScale`" — **no code reads `normAutofit`
  or `fontScale` anywhere under `src/lib/presentation` or `src/components/presentation`
  today** (verified by grep), so this spec implements the read as well as the recompute.
- `specs/081-slide-text-editing.md` owns slide-part serialisation and the editing session:
  `serializeSlideXml(slide, rels, slidePartPath)`, `serializePage(slide, registry,
  usedNames, master, styleSource)`, the `StyleRegistry` minting `de-P{n}`/`de-T{n}`/
  `de-G{n}`, `markSlideEdited(slide, format)` (sets `dirty`, deletes `sourceXml` on odp)
  as the single dirtying primitive, `TextPoint`/`TextRange`/`ParagraphRange`,
  `MAX_INDENT_LEVEL = 8`, `DEFAULT_SIZE_PT = 18`, and the contenteditable round trip
  `renderEditableParagraphs`/`normalizeTextFrameDom` (`slideTextDom.ts`, stamping
  `data-align`/`data-bullet`/`data-level` per block — new fields stamped the same).
- `specs/084-…` navigation (`usePlayback` → `{ current, atEnd, advance, back, goTo }`);
  `PresentationPlayView.vue`'s `advanceStep()` order and its `@click="onStageClick"`
  handler are the integration point for link clicks. `specs/080-…` undo: `record(label)`
  **before** mutating, `commit()` after, one entry per gesture. `specs/082-…` frame
  identity, `MIN_FRAME_SIZE_CM = 0.5`, `PRESENTATION_IMAGE_MIME_TYPES` and the
  media-minting rule (`ppt/media/image{n}.{ext}` / `Pictures/image{n}.{ext}`, `n` = 1 +
  the highest suffix).
- `specs/018-…`/`specs/032-…` own lists, and `src/lib/core/lists.ts` is already
  product-neutral ("a Slides product numbers the list items inside a text box with the
  same marker math"), exporting `ListMarkerStyle = "disc"|"circle"|"square"|"decimal"|
  "lowerAlpha"|"upperAlpha"|"lowerRoman"|"upperRoman"`, `markerText(style, ordinal)`,
  `formatOrdinal`, `isBulletStyle`, `markerType`, `defaultMarkerStyle`, `BULLET_CHARS`,
  `bulletStyleForChar`, `clampListLevel`, `MAX_LIST_LEVEL = 9`, `LIST_INDENT_STEP_CM =
  1.27`, `LIST_MARKER_WIDTH_CM = 0.635`. `specs/014-…` owns paragraph spacing
  (`ParagraphFormatting.spaceBeforePt`, `spaceAfterPt`, `lineSpacing`, `lineSpacingPt`,
  `lineSpacingRule?: "exact"`, absent = "at least"); `specs/025-…` owns hyperlinks
  (`sanitizeHref(url): string | null`, `isDangerousUrl`, `urlScheme`, and the two-scalar
  `linkHref`/`linkAnchor` precedent).

### Reuse, and the two-phase moves into `src/lib/core`
`lib/presentation` may import only presentation siblings and `lib/core`, so shared code
moves down in the pattern `documentModel.ts` used for `MediaPart`: **phase 1** adds it to
core and re-exports from the old module so every import path and colocated spec keeps
working; **phase 2** re-points call sites and deletes the copy. Two small moves:
1. `src/lib/links.ts` → new `src/lib/core/urls.ts`: `DANGEROUS_SCHEMES`, `urlScheme`,
   `isDangerousUrl`, `sanitizeHref`, `AutolinkMatch`, `detectBareUrl` (none touch
   `DocumentModel`; the rest of `links.ts` — bookmarks, `applyLink`, `inlineWriteLayout` —
   stays text-side).
2. `src/lib/core/xml.ts`'s `Relationship` gains `targetMode?: "External"`, parsed by
   `parseRelationships`, emitted by `appendRelationships` in `pptx/writer.ts` — additive
   and optional, so no existing caller changes.
Already neutral, consumed directly: `core/lists.ts` (**both** private per-level indent
constants are deleted in its favour — `odp/serialize.ts:39`'s `LEVEL_INDENT_CM = 1.27` and
`slideRender.ts:16`'s divergent `0.95`), `core/specialCharacters.ts`, `core/units.ts`,
`core/grapheme.ts` (validating `bulletChar`), `core/media.ts`. Not reused:
`DefineListDialog.vue` and `CustomSpacingDialog.vue` are bound to the text model — a slim
`SlideListPopover.vue` and the Text panel render the same vocabulary over `core/lists.ts`.

### Autofit — where the type lives, and measuring outside `src/lib`
`TextBodyProps` and its `autofit`/`anchor` unions are declared in
`src/lib/presentation/model.ts`, exactly where SPEC-085/086 put `SlideTransition` and
`SlideAnimation`: model.ts imports type-only from `../core/*`, feature modules re-export
**from** it and never the reverse (`animations.ts:13`), and `textBody.ts`'s
`effectiveInsetsCm(frame)`/`effectiveAnchor(frame)` take `TextFrame`/`ShapeFrame` — so
declaring the type downstream would invert the layering into a cycle. `src/lib/` must also
stay unit-testable without mounting a component:
- `src/lib/presentation/textBody.ts` (framework-free; derivations only, importing
  `TextBodyProps` from `./model` and re-exporting it): `effectiveInsetsCm(frame)`,
  `effectiveAnchor(frame)`, `AUTOFIT_STEPS = [100, 92.5, 85, 77.5, 70, 65, 60, 55, 50, 45,
  40, 35, 30, 25]`, and `fitScale(measureHeightCm: (fontScalePct: number, lineScalePct:
  number) => number, availableCm: number): { fontScalePct: number; lineScalePct: number }`
  — the first step that fits, allowing `lineScalePct` 90 at font scales ≤ 70 and 80 at
  ≤ 50, else the last step. No DOM.
- `src/components/presentation/slideTextFit.ts` (DOM, beside `slideTextDom.ts` under its
  header rule — "a rendering concern, NOT under src/lib/presentation"):
  `measureBodyHeightCm(host, widthCm, fontScalePct, lineScalePct)`, measuring an off-flow
  clone at the stage's own font stack, plus `applyFitStyle(el, body)` setting
  `--fit-scale`, `line-height`, `padding`, `justify-content` and `column-count`/`-gap`.
The renderer measures; only a command writes — `useSlideTextDepth.ts` stores a scale only
inside a mutation that already called `markSlideEdited`.

### pptx wire mapping (write; read tolerantly)
`a:pPr` attributes: `marL` = `indentCm` in EMU, `indent` = `−hangingCm` in EMU (negative =
hanging), `lvl`, `algn`. Children, in the ECMA-376 `CT_TextParagraphProperties` order the
serialiser must keep: `a:lnSpc` (`a:spcPct val` in 1000ths % for `lineSpacing`,
`a:spcPts val` in 100ths pt for `lineSpacingPt`) → `a:spcBef`/`a:spcAft` (`a:spcPts val`)
→ preserved `a:buClr`/`a:buSzPct`/`a:buFont` → the bullet (`a:buNone` | `a:buChar
char="…"` | `a:buAutoNum type="…" startAt="…"` | `a:buBlip><a:blip r:embed="…"`) →
preserved `a:tabLst`/`a:defRPr`/`a:extLst`. Marker styles map to `buAutoNum@type`
`arabicPeriod`/`alphaLcPeriod`/`alphaUcPeriod`/`romanLcPeriod`/`romanUcPeriod`, and to
`buChar@char` `•`/`○`/`▪` for disc/circle/square; `a:rPr@spc` is `letterSpacingPt` in
100ths of a point. `a:bodyPr` attributes: `lIns`/`tIns`/`rIns`/`bIns` (EMU; defaults
91440/45720/91440/45720), `anchor` `t|ctr|b`, `numCol`, `spcCol` (EMU), `wrap="none"`;
its one autofit child is `a:noAutofit`,
`<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>` (1000ths of a percent) or
`a:spAutoFit`, emitted before any preserved `a:scene3d`/`a:sp3d`/`a:extLst`. Links:
`<a:hlinkClick xmlns:r="…" r:id="rIdN" tooltip="…"/>` as the first child of `a:rPr`
(runs) or of `p:cNvPr` (frames); external targets get a `…/relationships/hyperlink`
relationship with `TargetMode="External"`, in-deck targets a `…/relationships/slide` one
plus `action="ppaction://hlinksldjump"`. `linkFrameRels` generalises `writer.ts:239`'s
`linkImageRels` — same "append after the highest existing rId" rule, same
`appendRelationships` splice; an unreferenced hyperlink relationship is never deleted, so
untouched `.rels` bytes stay identical. **Both pptx media walks are image-frame-only
today**, so `referencedMedia` (`pptx/writer.ts:178`) and `linkImageRels` (`:239`) must
each collect `SlideParagraph.bulletImageId` beside `ImageFrame.mediaId`, or `planMedia`
(`:210`) writes neither the part nor its `<Default Extension=…>` row, no `IMAGE_REL` is
minted, and `<a:buBlip><a:blip r:embed>` dangles. Reading is tolerant of unit form
(`spcPct` vs `spcPts`) and attribute order; a body carrying `vert`, `prstTxWarp` or an
unrecognised autofit child models what it can and keeps the rest in `pptxBodyPr`.

### odp wire mapping (write; read tolerantly)
Paragraph properties join `StyleRegistry.paragraphStyle`'s property string:
`fo:line-height="150%"` (multiple), `fo:line-height="18pt"` (exact),
`style:line-height-at-least="18pt"` (the absent-rule default), `fo:margin-top`/
`fo:margin-bottom` in cm (pt ÷ 72 × 2.54, `r3`), `fo:margin-left` for `indentCm` replacing
today's `level × 1.27`, `fo:text-indent` (negative) for `hangingCm`; `textStyle` gains
`fo:letter-spacing` in cm. Lists become real markup: a run of list paragraphs wraps in
`<text:list text:style-name="de-L{n}">`, one `<text:list-item>` each, nested by level, and
the registry mints `de-L{n}` `text:list-style` elements whose levels are
`text:list-level-style-bullet` (`text:bullet-char`), `-number` (`style:num-format`
`1|a|A|i|I`, `style:num-suffix="."`, `text:start-value`) or `-image` (`xlink:href` into
`Pictures/`). Body properties go on the frame's graphic style:
`fo:padding-left|-top|-right|-bottom`, `draw:textarea-vertical-align="top|middle|bottom"`,
`draw:auto-grow-height="true"` for `"resizeShape"`, and `<style:columns fo:column-count
fo:column-gap/>` inside `style:graphic-properties`. Links: `<text:a xlink:href="…"
xlink:type="simple" office:title="…">` wrapping a run's spans (the reader already flattens
`text:a` — `reader.ts:347` — and now keeps the href), and `<draw:a
xlink:href="…">` wrapping a whole shape element. `referencedMedia(model)` (`odp/writer.ts:31`)
carries the identical trap and collects `bulletImageId` too, feeding `spliceManifest`
(`:50`). **Three ODF details are not settled from memory and go through SPEC-086's
golden-fixture-first process** (`odp/textDepthGolden.ts`, hand-authored against ODF 1.2
§15.5/§16.31, asserted byte-exact by unit test, Impress verification recorded in the PR,
not CI): the shrink-on-overflow attribute (`draw:fit-to-size` vs `style:shrink-to-fit`),
the numbering-restart attribute (`text:continue-numbering="false"` on `text:list` vs
`text:start-value` on the item), and the slide-jump link form (`xlink:href="#{draw:name}"`
vs a
`presentation:event-listener presentation:action="jump"` child). Until a golden exists the
odp writer omits that one property — the field still round-trips on the pptx side — and
the reader preserves the source markup. Never invent bytes.

### Links, editing and playback
`isLinkableKind(frame)` = `kind` is `"text" | "image" | "shape"`. A run link and a frame
link are independent; a frame link covers the whole box and a run link inside it wins for
that run. On the editor stage the anchor is inert (`@click.prevent` unless `metaKey ||
ctrlKey`) so a click still starts a text session per SPEC-081; in playback it is live and
`onStageClick` returns early for it, so a link click never also advances a build, and
`backStep`/`jumpTo` stay SPEC-084's. Undo labels, in the existing voice: "Bullets",
"Numbering", "List style", "Indent", "Line spacing", "Paragraph spacing", "Letter
spacing", "Text autofit", "Text anchor", "Text insets", "Text columns", "Add link",
"Remove link".

### Tests
- `src/lib/presentation/text.spec.ts` (widened patch, `toggleBullet`'s third argument,
  spacing/letter-spacing commands, `SelectionFormatState` mixing, promote/demote
  clamping) and new `src/lib/presentation/textBody.spec.ts` (`fitScale` ladder against a
  synthetic measurer, inset/anchor defaults, `resizeShape` floor).
- `pptx/serialize.spec.ts` and `odp/serialize.spec.ts`: `a:pPr`/`a:bodyPr` child order,
  preserved-slot re-emission, list-style minting, hyperlink rels (including
  `TargetMode="External"`), the odp goldens once authored, and — in both `writer.spec.ts`
  halves — a picture-bullet deck saving with the media part, its content-type row and its
  relationship present. New `src/lib/core/urls.spec.ts`; `links.spec.ts` unchanged.
- New `src/lib/presentation/textDepthRoundTrip.spec.ts` per the round-trip AC;
  `roundTrip.spec.ts`'s byte-identity assertions are untouched.
- `slideTextDom.spec.ts` (new fields survive the contenteditable round trip), new
  `slideTextFit.spec.ts`, new `slideRender.spec.ts` (the rewritten `markedParagraphs` and
  the 1.27 cm indent), `useSlideTextEditing.spec.ts` (Tab/Shift+Tab), new
  `useSlideTextDepth.spec.ts`, `PresentationPlayView.spec.ts` (a link click does not
  advance), `src/views/PresentationView.size.spec.ts`, `e2e/SPEC-097.spec.ts`.

### Specs this changes
- SPEC-079's out-of-scope line on `normAutofit`, columns and "hyperlinks render as plain
  styled text" is superseded; its fixture expectations are unchanged — coverage preserved.
- SPEC-081's out-of-scope list loses numbered authoring, autofit, line spacing, letter
  spacing and links; its serialisers gain the `a:pPr`/`a:bodyPr` and `text:list` rules
  above, and its `pPrXml`/`txBodyXml` stop discarding source properties.
- SPEC-082's media walks (`referencedMedia`/`linkImageRels` in both writers) now also
  collect `bulletImageId`; its image-insert behaviour is unchanged. SPEC-088 exports
  markers, spacing, autofit scale and links through the same helpers; SPEC-094 binds link
  colour to the theme's `hlink`/`folHlink` on top of this, and `specs/000-roadmap.md`
  gains this spec's phase row. Nothing else changes.
