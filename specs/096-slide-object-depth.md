---
id: SPEC-096
title: Slide object depth — rotation, handles, multi-select, groups, align and distribute
status: draft
area: editing
complexity: complex
story_points: 13
depends_on: [SPEC-069, SPEC-072, SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-086]
---
## Why
SPEC-082's out of scope names the hole in one breath: "Rotation *editing*
(rendering only, per SPEC-079) … adjustment-handle dragging on the stage,
multi-select and marquee, group *creation* … snapping to other frames (slide
centrelines/edges only)". A deck author moves exactly one box at a time: three
shapes cannot be lined up, a rounded rectangle's corner cannot be tuned, a
rotated logo cannot be straightened, and ungroup is the only group verb that
exists, because SPEC-079 models a `GroupFrame` as read-only verbatim `xml`.

Two of those gaps are already *silent data loss*, which makes this a
preservation fix as much as a feature. `pptx/reader.ts:471` reads
`ShapeFrame.adjustments` through `parseAvLstAdjustments` while
`pptx/serialize.ts:245` emits a bare `<a:avLst/>` — every adjustment on a
dirtied slide is discarded on save. `odp/reader.ts` never parses rotation while
`odp/serialize.ts:336` writes `draw:transform="rotate(…)"`, so odp rotation
survives one write and dies on the next open. A handle is worthless if its value
cannot come back, so both close here. Everything else is reuse:
`core/shapeGeometry.ts` already owns rotation, `core/shapeCatalogue.ts` the
adjustment specs, `core/units.ts` the units, so the one new neutral module is the
rect algebra SPEC-072 will also want. **This file is the normative source for
slide multi-selection, slide rotation editing, authored `GroupFrame`s (creation,
release and their canonical `p:grpSp` / `draw:g` forms) and slide
align/distribute.**

## Acceptance criteria
- [ ] `src/views/presentation/useFrameSelection.ts` gains multi-selection per the Reference: `selectedFrameIds: Ref<string[]>` (z-order stable, top-level frames only), `toggle(id)`, `selectMany(ids)`, `selectAll()`, `isSelected(id)`, `anchorFrame: ComputedRef<SlideFrame | null>`; `selectedFrameId` stays a `Ref<string | null>` holding `selectedFrameIds[0]` — the first-selected anchor — so `useAnimations`' `selectedFrameId`/`selectFrame` deps and `useSlideTextEditing` are untouched, and the existing `watch(() => currentSlide.value?.id)` clears the whole set
- [ ] Shift+click toggles a frame in or out of the selection, Cmd/Ctrl+A selects every frame on the slide, a click on empty stage without a drag still deselects (SPEC-082 unchanged), and two or more selected frames draw `selection-bounds` — the union box, carrying **no** resize or rotation handles — in place of `frame-selection`
- [ ] Marquee selection: a pointer-down on empty stage that travels past `DRAG_ENGAGE_CM` (hoisted to `stageMath.ts` and exported per the Reference) draws a `marquee-box` and, on release, selects every top-level frame whose `rotatedBounds` box intersects it — Shift adds to the existing selection instead of replacing it, Escape mid-drag abandons the marquee and changes nothing, and a release still under the threshold is the plain deselecting click above
- [ ] Geometry commands apply to the whole selection in ONE undo entry: a drag or an arrow nudge moves every selected frame by the same delta, clamped on the **union** box against `FRAME_MIN_ON_SLIDE_CM = 1` so relative positions never shift; Cmd/Ctrl+D duplicates them all by `DUPLICATE_OFFSET_CM = 0.5` and leaves the copies selected; Delete/Backspace removes them all, each target's animations dropped through `removeAnimationsForFrame`
- [ ] Arrange moves a multi-selection as one block: the new `reorderFrames(model, slideIndex, frameIds, target)` (`src/lib/presentation/frames.ts`, `target` the same `"front" | "forward" | "backward" | "back"` `reorderFrame` takes) lifts the selected frames out of z-order and re-inserts them contiguously, preserving their relative order, in one undo entry, and is a no-op leaving the order untouched when the block already sits at that end
- [ ] Style, ungroup and collapse-to-edit over a multi-selection: fill and stroke apply to every selected frame that supports them (the popover seeded from the anchor, unsupported kinds skipped rather than erroring), Ungroup releases every selected group in one entry, rotate/resize/adjust stay unavailable while `selection-bounds` is showing, and Enter or double-click collapses the selection to that one frame before starting its text session
- [ ] A single selected frame shows a ninth handle, `rotate-handle`, above the `n` handle: dragging sets `rotationDeg` to the pointer's angle about the frame centre through `normalizeRotationDeg` (`src/lib/core/shapes.ts`, whole degrees 0–359), Shift snaps to the nearest multiple of 15°, the live angle reads out in an `aria-live` `rotation-readout`, Escape mid-drag restores the starting angle, and all eight resize handles show `rotatedResizeCursor(handle, rotationDeg)` (`src/lib/core/shapeGeometry.ts`) once rotated
- [ ] Rotation composes with SPEC-082's resize maths per "Rotation maths": `stageMath.resizeRect` keeps its signature and runs in the frame's unrotated local space, the pointer mapped in by `unrotatePoint` and the result re-anchored so the corner opposite the dragged handle stays fixed **in slide space**; hit-testing, marquee intersection, the on-slide clamp and guide activation use `rotatedBounds`; `FrameSelectionOverlay.vue`'s `sel-box` carries `transform: rotate(${rotationDeg}deg)` about the box centre so handles sit on the shape's own corners
- [ ] Adjustment handles land on the stage and their values now survive a save: a selected `ShapeFrame` whose `shapeDefinition(geometry)?.adjustments` is non-empty shows one `adjust-handle` per `ShapeAdjustmentSpec` (diamond, distinct from the square resize handles), constrained along its declared `axis` (`"x" | "y" | "radial"`) between `min` and `max`, live during the drag, `aria-label` from the spec's `label`, arrows nudging 1% of the range and Shift 5%, stored as the 0..1-style fractions `resolvedShapeAdjustments` resolves; `shapeFrameXml` (`pptx/serialize.ts:245`) emits real `<a:gd name="${adjustmentGuideName(i, n)}" fmla="val ${adjustmentToGuideValue(v)}"/>` children inside `<a:avLst>` — the only one of that file's five `<a:avLst/>` sites that ever carries adjustments, per the Reference — and `odp/serialize.ts`'s `<draw:enhanced-geometry>` gains `draw:modifiers`, with an odp reader parsing it back
- [ ] `GroupFrame.xml` becomes optional (`xml?: string`) with exactly `Slide.sourceXml`'s discipline — present = the verbatim source `p:grpSp`/`draw:g`, absent = an authored group the serialisers rebuild; `groupFrames(model, slideIndex, frameIds)` (new, `src/lib/presentation/frames.ts`) replaces two or more selected top-level frames with one `GroupFrame` at their union box carrying them as `children` in slide space and **no** `xml`, and any command that changes an existing group's membership, child geometry or rotation `delete`s its `xml` first — the only way that field ever changes
- [ ] Releasing a *verbatim* group's `xml` while it carries a non-zero `rotationDeg` is lossy (SPEC-079 flattens children without the group's rotation), so it parks in a `pendingRelease` confirm exactly as SPEC-085/086's `pendingReplace` does — `role="alertdialog"` `aria-label="Replace the original grouping"`, `group-release-confirm` / `group-release-cancel`, cancel changing nothing; a group with `rotationDeg` 0 releases silently, and rotating an authored group composes onto its children per "Rotation maths" so `GroupFrame.rotationDeg` is deleted rather than written
- [ ] Authored groups serialise per "Groups and adjustments on the wire": pptx `p:grpSp` whose `a:chOff`/`a:chExt` are set **identical** to `a:off`/`a:ext` (an identity child mapping, so `parseGroupChildMapping`/`mapGroupChild` return the children's slide-space EMU unchanged); odp `<draw:g>` whose children keep their own `svg:x`/`svg:y`, matching `odp/reader.ts:531` ("a draw:g has no box of its own"); and `serializeSlideXml`'s **two-pass** `p:cNvPr` id mint collapses into the single pass the Reference's "Wire ids" makes normative, so `spidFor` and the emitted ids cannot drift once an authored group claims a scaffold id plus one per descendant — `pptx/serialize.spec.ts` pins it with an authored group followed by an animated shape, asserting the emitted `p:spTgt spid` equals that shape's own `p:cNvPr id`. `odp/serialize.ts`'s `wireIdFor` is a single pass already and is unchanged; a group still carrying `xml` re-emits verbatim and claims no id in either format
- [ ] `src/lib/core/arrange.ts` (new, product-neutral, core siblings only) holds the pure cm-rect algebra listed in the Reference and drives an Align menu in `SlideInsertBar.vue`'s ⋯-collapsing secondary group: `align-menu` → `align-popover` with `align-left`, `align-center-x`, `align-right`, `align-top`, `align-middle-y`, `align-bottom`, `distribute-h`, `distribute-v` and an `align-reference` select offering Selection (default at ≥2 frames), Slide (the only option at exactly 1) and First selected; distribute needs ≥3, is `aria-disabled` with a reason otherwise, holds the outermost two and equalises gaps; nothing is two-phase-moved because SPEC-072 is still `draft` and unimplemented (no `alignRects`/`distributeRects` exists in `src/`), so this file exists for SPEC-072 to consume rather than duplicate
- [ ] Smart guides extend SPEC-082's slide-only set: `stageMath.applyGuides` keeps its four-parameter signature and behaviour (so `stageMath.spec.ts` passes with no expectation changed) and gains an optional fifth `others: Rect[]`; while dragging or resizing, a guide fires within `GUIDE_SNAP_CM` (0.15 cm) of any other top-level frame's left/right/top/bottom edge or its horizontal/vertical centre as well as the slide's centrelines and edges, at most one snap per axis with slide centre first, rendered through the existing `align-guide` testid plus a new `data-guide-kind="slide" | "frame"` so existing selectors keep matching
- [ ] A "Position and size…" dialog (`slides-position-toggle` → `slides-position-dialog` over `src/components/BaseDialog.vue`) edits the anchor frame's `position-x`, `position-y`, `position-width`, `position-height` in the `measurementUnit` preference (`formatLength` / `parseLengthToCm` / `roundForDisplay` from `src/lib/core/units.ts`) plus `position-rotation` in whole degrees and a `position-aspect-lock` toggle; each field commits on blur or Enter as one undo entry, clamps to `MIN_FRAME_SIZE_CM = 0.5` and `FRAME_MIN_ON_SLIDE_CM = 1`, and is the keyboard/touch path SPEC-068 requires for every drag gesture this spec adds
- [ ] Undo granularity is one entry per gesture, never per frame: `record(label)` fires once before the first mutation of a drag/rotate/adjust/align/group gesture and `commit()` once on release (SPEC-080's `usePresentationUndo`, whose `label` is still discarded), with labels in the existing voice — "Rotate shape", "Adjust shape", "Group", "Ungroup", "Align left", "Distribute horizontally", "Position and size", and `Move ${n} objects` for a multi-selection move; undo restores geometry, `rotationDeg`, `adjustments`, `children` and `GroupFrame.xml` together
- [ ] Round trip, naming its relations: a no-edit download stays byte-identical part-for-part under `equalBytes` from **`src/lib/core/equality.ts`** (compared after `unzip`, never container bytes), with `roundTrip.spec.ts`'s `expectUntouched` and odp splice-equality assertions unchanged; rotate + adjust + group + align → save → reopen reproduces `rotationDeg` exactly, `adjustments` within 1e-5, geometry within ±0.01 cm and the group's `children` order, compared with `equalStructural` from `src/lib/equality.ts` over the modelled subset — its `IGNORED_KEYS` drops frame `id`s and its `XML_STRING_KEYS` contains `xml`, so a verbatim `GroupFrame.xml` compares canonically via `canonicalizeXml`; `equalSemantic` is typed `(model: DocumentModel)` and is **not** usable on a `PresentationModel`, so no semantic relation is claimed; writing the reopened model twice yields identical slide-part bytes
- [ ] Fixture coverage: `fixtures/build.ts` gains a `spec096-objects` builder pair at the end under its own banner (never mutating the four mini-deck slides) — three deliberately unaligned shapes, a 37°-rotated image, a `roundRect` with a non-default adjustment, a verbatim group whose `a:chOff`/`a:chExt` differ from `a:off`/`a:ext` (pinning the mapping), and a **rotated** verbatim group (pinning the confirm-to-release path) — and `scripts/build-presentation-fixtures.ts` writes `spec096-objects.pptx` and `spec096-objects.odp` into `e2e/fixtures/`
- [ ] e2e (`e2e/SPEC-096.spec.ts`): open `spec096-objects.pptx` → marquee-drag over the three shapes → `selection-bounds` shows, `slide-frame` count unchanged → Align → "Align left" → their positions match → "Distribute vertically" → gaps equal → Group → one frame remains, Ungroup restores three → select the image, drag `rotate-handle` with Shift → `rotation-readout` reads a multiple of 15 → "Position and size…" → set rotation 90 → download, reopen → rotation, alignment and the tuned adjustment all survive

## Out of scope
Connectors between slide frames and arrowheads (SPEC-070's model is still text-side
and unbuilt; a slide connector stays a `line` `ShapeFrame`). Proportional resize or
rotation of a multi-selection as a unit — group it first. Nested group *creation* (an
imported group may be a child, but grouping never re-parents into an existing group).
Copy/paste of frames across slides or decks, "make same size", a visible grid,
user-placed ruler guides, snap-to-grid and its preference (SPEC-072 owns those
text-side and this spec ships neither), alignment to a layout's placeholder slots
(SPEC-089), freeform geometry and edit points (SPEC-071), shape effects (SPEC-073),
image cropping, aligning table cells or chart series (SPEC-091/092). Animating a
group or its children: grouping never widens `isAnimatableKind`, a group and its
descendants stay unanimatable whatever else that predicate later admits, and grouping
frames that carry animations drops those animations with a one-line notice.

## Reference

### Dependencies restated
- **SPEC-079** (`specs/079-slide-canvas-rendering.md`) owns the slide model:
  `SlideFrameBase { id: string; xCm; yCm; widthCm; heightCm: number; rotationDeg?:
  number }` (`id` minted `crypto.randomUUID()` at read, never serialised; lengths
  cm, rotation whole degrees clockwise), `GroupFrame extends SlideFrameBase { kind:
  "group"; children: SlideFrame[]; xml: string }` with children already flattened
  into slide space, `ShapeFrame` adding `geometry: ShapeGeometry`, `adjustments?:
  number[]`, `fill?`, `stroke?`, `paragraphs?`, `ForeignFrame` adding `xml`,
  `label`, `noBounds?`.
- **SPEC-082** (`specs/082-slide-objects-images-shapes-text-boxes.md`) owns
  `src/lib/presentation/frames.ts`: `FrameRect { xCm; yCm; widthCm; heightCm }`,
  `moveFrame`, `setFramePosition`, `resizeFrame`, `deleteFrame`, `duplicateFrame`,
  `reorderFrame(model, slideIndex, frameId, "front" | "forward" | "backward" |
  "back")`, `ungroupFrame`, `setFrameFill`, `setFrameStroke`, plus
  `MIN_FRAME_SIZE_CM = 0.5`, `FRAME_MIN_ON_SLIDE_CM = 1`, `DUPLICATE_OFFSET_CM =
  0.5`, `NUDGE_STEP_CM = 0.1`, `NUDGE_LARGE_STEP_CM = 1`, `GUIDE_SNAP_CM = 0.15`.
  Each command mutates one slide in place, ends with `markSlideEdited(slide,
  model.format)` (`presentation/text.ts` — `slide.dirty = true` and, for odp,
  `delete slide.sourceXml`) and returns the id(s) to select. View side:
  `useFrameManipulation.ts` (exporting `frameNoun`, `frameAccessibleName`),
  `useFrameInsertion.ts`, and `stageMath.ts`, whose entire export list is
  `ResizeHandle`, `RESIZE_HANDLES`, `isCornerHandle`, `clientToCm`, `cmToStagePx`,
  `resizeRect`, `StageGuides`, `applyGuides`. **The drag threshold is not among
  them.** `DRAG_ENGAGE_CM = (3 * 2.54) / 96` is a `const` inside
  `useFrameManipulation.ts:104`'s drag closure, read at `:122` and exported
  nowhere, so a marquee in another composable cannot see it; this spec **hoists it
  to module scope in `stageMath.ts` and exports it**, repointing that one use —
  same value, no behaviour change. `src/lib/core/shapes.ts:221` separately exports
  the numerically identical `DRAW_DRAG_THRESHOLD_CM` for draw-to-insert
  (`useFrameInsertion.ts:110`); **it survives untouched**, and neither is expressed
  in terms of the other — "a drag has begun" and "a drawn box is worth keeping" may
  diverge.
- **SPEC-080** owns the writers' rebuild rules and the undo stack —
  `usePresentationUndo` → `{ record(label?), undo, redo, canUndo, canRedo, reset }`,
  `MAX_UNDO_STEPS = 100`, snapshot `structuredClone(model.slides)` +
  `structureDirty`, `record` **before** the mutation and `commit()` after (one entry
  per commit); the writers re-emit verbatim only when `!structureDirty && !anyDirty
  && !anyNotes`. **SPEC-081** owns slide-part serialisation:
  `serializeSlideXml(slide, rels, slidePartPath)` with its fixed order and
  `modelled ?? preserved ?? ""` precedence, and `serializePage(slide, registry,
  usedNames, master, styleSource)`. **SPEC-086** owns the wire-id passes this
  spec perturbs (see "Wire ids") and `removeAnimationsForFrame(animations,
  frameId)`, which returns the **same reference** when nothing matched.
- **SPEC-069** (`specs/069-shape-catalogue-expansion.md`) owns the adjustment
  vocabulary, already neutral in `src/lib/core/shapeCatalogue.ts`:
  `ShapeAdjustmentSpec { label: string; default: number; min: number; max:
  number; axis: "x" | "y" | "radial" }`, `shapeDefinition(geometry)`,
  `resolvedShapeAdjustments(geometry, stored?)` (defaults filled, clamped, never
  mutating the stored array), `adjustmentToGuideValue(fraction)` /
  `guideValueToAdjustment(value)` (scale 100000) and `adjustmentGuideName(index,
  count)` (`"adj"` alone, else `"adj1"`…). **SPEC-072**
  (`specs/072-align-distribute-snap-and-position.md`) is the text-side
  counterpart and is **still `draft` and unimplemented** — nothing in `src/`
  aligns or distributes — so its conventions are restated here (align to a
  bounding box; distribute leaves the outermost two in place; guides never touch
  the model; a numeric dialog is the keyboard/touch path) and its maths lands in
  `src/lib/core/arrange.ts` for both products.
- Neutral core consumed rather than rebuilt: `core/shapeGeometry.ts` (`Point`,
  `rotatePoint`, `unrotatePoint`, `unrotateDelta`, `rotatedBounds(width, height,
  rotationDeg) → { width; height; padX; padY }`, `rotatedResizeCursor`,
  `pointHitsShape`, `HIT_TOLERANCE_CM`), `core/shapes.ts` (`normalizeRotationDeg` →
  0..359), `core/drawingMl.ts` (`parseAvLstAdjustments`, `parseXfrm`,
  `GroupChildMapping`, `parseGroupChildMapping`, `mapGroupChild`, `GEOMETRY_TO_PRST`,
  `GEOMETRY_TO_DRAW_TYPE`, `firstA`), `core/units.ts` (`MeasurementUnit`,
  `asMeasurementUnit`, `formatLength`, `parseLengthToCm`, `roundForDisplay`,
  `cmToEmu`, `emuToCm`), `core/odfXml.ts` namespaces. `lib/presentation` imports only
  presentation siblings and `lib/core` (`presentation/boundary.spec.ts`; a bare
  `"vue"` is a violation on its own); a core module imports only core siblings.
  Equality: `equalBytes(a, b)` comes from `src/lib/core/equality.ts` (what
  `presentation/roundTrip.spec.ts` imports); `equalStructural` /
  `firstStructuralDifference` from `src/lib/equality.ts` take `unknown` and so work on
  a `PresentationModel`; `equalSemantic` there is typed `(model: DocumentModel)` and
  does not.

### The model delta
```ts
// src/lib/presentation/model.ts — the ONLY model change this spec makes.
export interface GroupFrame extends SlideFrameBase {
  kind: "group";
  children: SlideFrame[];          // coordinates already in slide space
  /** The whole source `p:grpSp` / `draw:g`, verbatim. ABSENT means an authored
   *  group the serialisers rebuild — exactly `Slide.sourceXml`'s discipline. Any
   *  command changing membership, child geometry or rotation deletes it first. */
  xml?: string;
}
```
`rotationDeg`, `adjustments` and `children` already exist; this spec makes them
editable and, for the first two, durable. `GroupFrame.rotationDeg` is never authored.

### Rotation maths (normative)
All of it applies `core/shapeGeometry.ts`; none of it is new maths.
- **Rotate gesture.** `deg = normalizeRotationDeg(atan2(py − cy, px − cx) ×
  180/π + 90)` about the frame centre; Shift → `Math.round(deg / 15) * 15`.
- **Resize of a rotated frame.** The centre `c` and the anchor corner `a`
  (opposite the dragged handle) are taken in slide space *before* the drag; the
  pointer maps into local space with `unrotatePoint(p, widthCm, heightCm,
  rotationDeg)` about `c`, `stageMath.resizeRect(handle, orig, localPointer,
  aspect)` runs unchanged, and the local rect is translated so rotating its
  anchor corner about its own new centre lands on `a`: `newCentre = a +
  R(θ)·(localCentre − localAnchor)`. At θ = 0 this is bit-identical to today,
  which is what keeps `stageMath.spec.ts` green.
- **Bounds.** Hit-testing, marquee intersection, guide candidates and the
  `FRAME_MIN_ON_SLIDE_CM` clamp use `rotatedBounds(widthCm, heightCm,
  rotationDeg)`; `pointHitsShape` with `HIT_TOLERANCE_CM` decides frame hits.
- **Group rotation.** Rotating a group by θ is rigid, so it is expressed
  entirely on the children: each child's centre rotates about the group centre by
  θ and each child's own `rotationDeg` gains θ (both through
  `normalizeRotationDeg`); `GroupFrame.rotationDeg` is then `delete`d. Nothing
  group-level reaches either format, so no group-rotation bytes are invented.
- **odp read gap.** `odp/reader.ts` gains the inverse of `odp/serialize.ts:336`
  — `draw:transform="rotate(r)"` is counter-clockwise radians, so `rotationDeg =
  normalizeRotationDeg(-r × 180/π)`. A `draw:transform` carrying anything but a
  lone `rotate(…)` (a `translate`, `skew` or matrix) is **not** modelled:
  rotation stays 0 and the frame is preserved as it is today.

### Groups and adjustments on the wire
An authored group, pptx, inside `p:spTree` in z-order:
```
<p:grpSp>
 <p:nvGrpSpPr><p:cNvPr id="{n}" name="Group {n}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
 <p:grpSpPr><a:xfrm>
   <a:off x="{X}" y="{Y}"/><a:ext cx="{CX}" cy="{CY}"/>
   <a:chOff x="{X}" y="{Y}"/><a:chExt cx="{CX}" cy="{CY}"/>
 </a:xfrm></p:grpSpPr>
 {children — the same p:sp / p:pic / p:grpSp the serialiser already emits}
</p:grpSp>
```
`chOff`/`chExt` are identical to `off`/`ext`, so `mapGroupChild`'s
`sx = cx/chCx = 1` and the children's slide-space EMU survive re-parse
unchanged — the round trip is provable, not approximate. odp emits
`<draw:g xml:id="de-f{n}" draw:id="de-f{n}">{children}</draw:g>`, each child
keeping its own `svg:x`/`svg:y`/`svg:width`/`svg:height`.

**Wire ids (normative — the trap this spec springs).** `serializeSlideXml`
mints `p:cNvPr id` in TWO passes over one counter (`serialize.ts:61-75`): a
pre-pass fills `spidFor` for the text/shape/image frames, then `nextId` resets to
2 at `:72` and the emission pass hands `() => nextId++` to `frameXml`. The passes
agree today *only* because a group consumes no id. An authored group emits a
scaffold id **plus** one per descendant, so the counters would drift and every
animation after that group would target the wrong shape through
`writePptxTiming(slide.animations, (id) => spidFor.get(id))` (`:86`). Fix it by
construction, not by keeping two walks in step: collapse to a **single pass**
that fills `spidFor` as it emits, `frameXml` returning its markup and the ids it
claimed together. A group still carrying `xml` re-emits verbatim and claims
none; an authored group claims the scaffold id and each descendant's, and its
descendants land in `spidFor` too — addressable even though grouping never makes
them animatable. `odp/serialize.ts`'s `wireIdFor` is already a single pass and is
unchanged.

**Adjustments** replace the empty `<a:avLst/>` in exactly the three *frame*
serialisers — `textFrameXml` (`:194`), `shapeFrameXml` (`:245`) and
`imageFrameXml` (`:263`) — with `<a:avLst>` plus one `<a:gd name="…" fmla="val
N"/>` per resolved adjustment. Only `shapeFrameXml` can ever carry adjustments,
so the other two keep emitting the empty element, and `serializeNotesSlideXml`'s
two `<a:avLst/>` sites (`:119`, `:129`, both `prst="rect"` placeholders with no
adjustments anywhere in the model) are **untouched** — the notes part's bytes are
pinned by round-trip tests. odp's `<draw:enhanced-geometry>` gains
`draw:modifiers="{space-separated guide values}"`.

**Golden-first (SPEC-086's process, normative here too).** The `p:grpSp` scaffold,
the `a:gd` naming and above all ODF's `draw:modifiers` scale are easy to get subtly
wrong from memory, so hand-authored goldens go in `pptx/groupGolden.ts`
(`GROUP_GOLDEN`, `ADJ_GOLDEN`, `GOLDEN_GROUP_FRAME`) and `odp/groupGolden.ts`
(re-exporting the same model), written against ECMA-376 §19.3.1.22 / §20.1.9.11 and
ODF 1.3 §10.3.6 / §19.145 — not screen-scraped. Serialisers assert exact-string
equality against them; readers must accept them plus a reordered-attribute variant.
The `draw:modifiers` scale is **our own convention** (`adjustmentToGuideValue`'s
100000ths, as `shapeCatalogue.ts` already documents for OOXML), so what is promised
is a faithful round trip through our own reader/writer pair and clamped rendering,
not byte-parity with Impress; cross-app verification is manual QA in the PR, not CI.

### `src/lib/core/arrange.ts`, guides and multi-selection semantics
```ts
// Product-neutral: cm rectangles in, cm rectangles out. No Vue, no product
// types, core siblings only. SPEC-072 consumes this for the text side.
export interface Rect { xCm: number; yCm: number; widthCm: number; heightCm: number }
export type AlignEdge = "left" | "centerX" | "right" | "top" | "middleY" | "bottom";
export type AlignReference = "selection" | "first" | "container";
export function boundingRect(rects: readonly Rect[]): Rect | null;
export function alignRects(rects: readonly Rect[], edge: AlignEdge,
                           reference: AlignReference, container: Rect): Rect[];
export function distributeRects(rects: readonly Rect[], axis: "x" | "y"): Rect[];
export interface GuideCandidate { axis: "x" | "y"; line: number; kind: "edge" | "centre" }
export function guideCandidates(others: readonly Rect[], container: Rect): GuideCandidate[];
export function snapRect(rect: Rect, candidates: readonly GuideCandidate[],
                         thresholdCm: number): { rect: Rect; fired: GuideCandidate[] };
```
`alignRects` returns rects in the input order, moving only the aligned axis.
`distributeRects` needs ≥3, sorts by the axis' centre, holds the outermost two and
equalises the *gaps* (not the centres). `snapRect` fires at most one candidate per
axis — slide centre, then slide edge, then nearest frame line.
`stageMath.applyGuides(rect, slideWidthCm, slideHeightCm, thresholdCm, others?)` is
re-expressed over these with its four-argument behaviour preserved exactly.

Two behaviours the criteria state as absences rather than actions: rotate, resize
and adjust are simply **unavailable** while `selection-bounds` shows (it carries no
handles — group first, or collapse the selection), and the Animations panel's Add
still targets the anchor frame alone, SPEC-086 unchanged.

### UI, testids and view wiring
New testids, none colliding with the 112 in use: `selection-bounds`,
`marquee-box`, `rotate-handle`, `rotation-readout`, `adjust-handle`, `align-menu`,
`align-popover`, `align-left`, `align-center-x`, `align-right`, `align-top`,
`align-middle-y`, `align-bottom`, `distribute-h`, `distribute-v`, `align-reference`,
`group-button`, `group-release-confirm`, `group-release-cancel`,
`slides-position-toggle`, `slides-position-dialog`, `position-x`, `position-y`,
`position-width`, `position-height`, `position-rotation`, `position-aspect-lock`.
`ungroup-button`, `align-guide`, `resize-handle`, `frame-selection`, `slide-frame`
and `arrange-menu` keep their meaning.
Logic lands in `src/views/presentation/useFrameArrange.ts`
(`FrameArrangeDeps { model, currentSlide, currentIndex, selection, surface,
record, refresh, commit }`), constructed **inside** `usePresentationEditing.ts`
and re-exported as `stageEdit.arrange`, so `PresentationView.vue` gains **no**
new `use*({ … })` call — only template hookups. UI lands in `SlideInsertBar.vue`
(Align/Group/Position in the existing ⋯-collapsing secondary group),
`FrameSelectionOverlay.vue` (rotation and adjustment handles, the rotated
`sel-box`), and new `PositionSizeDialog.vue` (over `BaseDialog.vue`; props
`frame`, `unit`; emits `apply`, `dismiss`) and `SelectionBoundsOverlay.vue`, all
runtime `defineProps({ … })` / `defineEmits([ … ])` — espree parses `.vue`
scripts and rejects the generic form. To pay for the hookups the `.stage-column`
block (`PresentationView.vue` lines 351–431) extracts into
`SlideStageColumn.vue` and `PresentationView.size.spec.ts`'s `LINE_BUDGET` drops
by at least 30 from whatever it stands at when this lands (550 → 520 today), with
a ledger row appended; the budget is never raised. The dialog is a dialog, not a
panel, so `src/lib/preferences.ts` gains no key.

### Tests
- `src/lib/core/arrange.spec.ts` — six edges against all three references, distribute
  with 2/3/many and ragged rects, snap priority and threshold boundaries,
  `boundingRect` of an empty list. `src/lib/presentation/frames.spec.ts` —
  `groupFrames` union box and child order, `xml` deletion on every mutating path,
  `reorderFrames` block moves at each of the four targets, rotation composition,
  animations dropped on grouping.
- `pptx/serialize.spec.ts` — the single-pass id mint above all: an authored
  group followed by an animated shape, asserting the emitted `p:spTgt spid`
  equals that shape's `p:cNvPr id`, and the notes part's bytes unchanged. Plus
  `odp/serialize.spec.ts`, the two `groupGolden.ts` modules asserted both ways,
  and reader specs for `draw:transform` rotation, `draw:modifiers` and the
  non-identity `chOff`/`chExt` fixture group.
- `src/lib/presentation/roundTrip.spec.ts` gains this spec's round trips
  (`equalBytes` per untouched part, second-write byte stability) with existing
  assertions neither re-expressed nor weakened; new `objectsRoundTrip.spec.ts`
  and `objectsFixtures.spec.ts` beside the transitions/animations pairs.
- `src/views/presentation/stageMath.spec.ts` (unchanged expectations, new
  `others` cases), `useFrameManipulation.spec.ts`, new `useFrameArrange.spec.ts`
  and `useFrameSelection.spec.ts` (that file's first spec),
  `PositionSizeDialog.spec.ts`, `e2e/SPEC-096.spec.ts`.

### Specs this changes
- SPEC-082's out of scope loses rotation editing, adjustment-handle dragging,
  multi-select and marquee, group creation and "snapping to other frames"; its
  `frames.ts` command list gains `groupFrames` and `reorderFrames`, its
  `<a:avLst/>` shortcut becomes a real adjustment list, and its `DRAG_ENGAGE_CM`
  moves from a function-local `const` to an export of `stageMath.ts`. Its guide,
  nudge, clamp and duplicate constants keep their values.
- SPEC-079's "`GroupFrame.xml` is always present, children render only" rule
  becomes "present = verbatim, absent = authored"; SPEC-082's ungroup remains its
  other exit.
- SPEC-086's pptx id minting gains the scaffold and every descendant of an
  authored group, so `spidFor` is rebuilt as a single pass (see "Wire ids") — the
  map's contents change, its `p:spTgt spid` contract does not. `wireIdFor` and the
  animatable-kind predicate are untouched by this spec.
- SPEC-072 is narrowed by adoption, not contradicted: its align/distribute/snap
  maths now has one product-neutral home in `src/lib/core/arrange.ts`; snap-to-grid
  and user-placed guides stay its own. `specs/000-roadmap.md` gains this spec's
  phase row and strikes its backlog entry "Rotation editing, adjustment handles,
  multi-select on slides"; nothing else changes.
