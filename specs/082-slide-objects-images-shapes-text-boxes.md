---
id: SPEC-082
title: Slide objects — insert, arrange and style images, shapes and text boxes
status: draft
area: editing
depends_on: [SPEC-081]
---
## Why
SPEC-081 edits what a slide already contains; a presentation author's
first instinct — put a picture on the slide, drop in a text box, draw an
arrow — still has no path. This spec completes the direct-manipulation
loop on the stage: insert images, text boxes and catalogue shapes; select,
move, resize and nudge them; control z-order; ungroup imported groups;
and style fills and strokes. It deliberately re-uses the word processor's
drawing vocabulary (SPEC-051/069/072 conventions, the core catalogue and
gesture helpers) so the two products feel like one app, and it is the
spec that makes SPEC-087's Keynote import and SPEC-086's animations
meaningful — both need frames users can own and target.

## Acceptance criteria
- [ ] The presentation toolbar gains Insert controls: "Text box", "Image…", and a shape gallery (the same `core/shapeCatalogue` set and preview icons the text product's gallery shows, minus connector-only entries); all three are keyboard-operable and fit the 720px e2e viewport via the SPEC-081 overflow rule
- [ ] Insert → Image opens a file picker (png/jpeg/gif/webp); the image lands centred at its natural size scaled down to fit within 50% of the slide's width/height, becomes a `MediaPart` + `ImageFrame`, is selected, and renders immediately; a broken/unreadable file shows the standard error dialog and changes nothing
- [ ] Insert → Text box arms a draw gesture: drag on the stage draws the box (`drawGestureBox` semantics — Shift for square, minimum size rule) and a click without drag places a default 10×2.5cm box; the new frame opens straight into a text-editing session
- [ ] Choosing a gallery shape arms the same gesture (`constrainedDragSizeCm`/`defaultClickSizeCm` from `core/shapes.ts`); the shape lands with the deck-neutral default style (fill `#4472c4`, stroke `#2f528f` 1pt — the SPEC-051 defaults) and is selected
- [ ] A selected frame moves by pointer drag (pointer capture, cursor feedback) and by arrow-key nudge (0.1cm; Shift+arrow 1cm), clamped so at least 1cm of the frame stays on the slide; while dragging, centre and edge alignment guides appear when within 0.15cm of the slide's centrelines or margins-of-zero (edges), and the frame snaps to them
- [ ] Eight resize handles resize (corners keep aspect with Shift; images keep aspect by default, Shift frees them); minimum size 0.5×0.5cm; text reflows live in text frames and shape text stays centred
- [ ] Delete/Backspace removes the selected frame (not while a text session is active); Cmd/Ctrl+D duplicates it offset by +0.5cm/+0.5cm and selects the copy
- [ ] Arrange controls (toolbar menu + context menu): Bring to front / Bring forward / Send backward / Send to back reorder `frames` and the rendered z-order accordingly
- [ ] A selected `GroupFrame` offers Ungroup (menu + Cmd/Ctrl+Shift+G): children are promoted to top-level frames in place (their slide-space geometry is already resolved per SPEC-079), the group's verbatim `xml` is dropped, the slide dirties, and each child becomes individually selectable/editable; foreign children of a group survive as foreign frames
- [ ] A style popover for the selected frame edits: fill colour (with "No fill"), stroke colour, stroke width (0.25–8pt) for shapes and text boxes; for images only stroke controls; changes preview live and are one undo entry per commit
- [ ] Every operation dirties the slide, records one undo entry with a human label ("Insert image", "Move shape", …), persists write-through, and round-trips: save → reopen shows identical geometry (±0.01cm), z-order, styles and media in both formats
- [ ] Newly inserted images serialise correctly in both formats: pptx — media part added under `ppt/media/` with a Content-Types extension default if missing, slide `.rels` gains the `image` relationship; odp — part under `Pictures/`, `META-INF/manifest.xml` gains its entry; a deck that gained an image still re-emits every untouched part byte-verbatim
- [ ] Selection is visible to tests and assistive tech: the stage exposes each frame as a focusable element with `role="group"`(shapes/images) or the text semantics from SPEC-081, an accessible name ("Image", "Text box", shape name from the catalogue, "Group", foreign label), and `data-testid="slide-frame"`; Tab order follows z-order (SPEC-081's cycle now includes new frames)
- [ ] e2e (`e2e/SPEC-082.spec.ts`): blank deck → insert a text box, type in it → insert a shape by drag → nudge it with arrows → send it to back → insert an image (fixture PNG via the picker) → download, reopen → all three frames present with expected stacking and positions

## Out of scope
Connectors between slide shapes, image cropping/replace,
alt-text authoring, copy/paste of frames across slides/decks, gradients
and shape effects (SPEC-073 depth stays text-side for now), theme-aware
fill palettes. (Rotation editing, adjustment-handle dragging,
multi-select and marquee, group creation and snapping to other frames
were out of scope here originally; SPEC-096 ships them all.)

## Reference

### Dependencies restated
- SPEC-079 model/frame kinds and renderer; SPEC-080 undo
  (`record(label)`) + dirty/persist conventions; SPEC-081 serialisers
  (this spec extends them only where noted), editing-session rules,
  toolbar and overflow behaviour.
- Core drawing vocabulary: `core/shapeCatalogue.ts`
  (`SHAPE_DEFINITIONS`, `shapeDefinition`, `resolvedShapeAdjustments`),
  `core/shapes.ts` (`defaultClickSizeCm`, `constrainedDragSizeCm`,
  `drawGestureBox`, `asNudgeDirection`), `core/shapeGeometry.ts`
  (`pointHitsShape`, `rotatedBounds` for hit/handle placement on rotated
  frames), `core/shapeCanvas.ts` painting, `core/shapeIcons.ts` gallery
  icons, `core/media.ts` (`MediaPart`, data-URI resolution),
  `core/imageDimensions.ts` (natural size probing — the SPEC-024 open
  path uses it; same here).
- SPEC-051's shape style defaults (`#4472c4`/`#2f528f`/1pt) and gesture
  UX; SPEC-072's guide/snap distance conventions (0.15cm) — restated
  here so this file stands alone; implementations must not import
  text-product modules for them (they are constants; duplicate them in
  `presentation/` or hoist to core if already shared).

### Commands (`src/lib/presentation/frames.ts`, framework-free)
`insertTextBoxFrame(model, slideIndex, rect)`,
`insertShapeFrame(model, slideIndex, geometry, rect)`,
`insertImageFrame(model, slideIndex, media, rect)` (adds the `MediaPart`
to `model.media`, mints `mediaId` per the naming rule below),
`moveFrame(model, slideIndex, frameId, dxCm, dyCm)` (clamping),
`resizeFrame(...)` (min sizes, aspect flags),
`deleteFrame`, `duplicateFrame`, `reorderFrame(model, slideIndex,
frameId, "front"|"forward"|"backward"|"back")`,
`ungroupFrame(model, slideIndex, frameId)`,
`setFrameFill`, `setFrameStroke`, and — since SPEC-096 —
`groupFrames(model, slideIndex, frameIds)` and
`reorderFrames(model, slideIndex, frameIds, target)` (a multi-selection
moved as one contiguous block). Each dirties the slide (and for odp
drops `sourceXml`), returns what the view selects. Geometry stays cm;
hit-testing on rotated frames goes through `unrotatePoint`.

### Media naming
Minted `mediaId`s must not collide with parts already in the package:
pptx `ppt/media/image{n}.{ext}` and odp `Pictures/image{n}.{ext}` with
`n` = 1 + the highest existing numeric suffix across `opaque` + `media`.
The writers treat `model.media` as the source of truth for these parts:
a media part present in `media` is emitted from there (base64 →
bytes), taking precedence over an `opaque` twin of the same name — which
keeps SPEC-079's duplication harmless and makes inserted images
first-class. Content-Types (pptx): ensure `<Default Extension="png"
ContentType="image/png"/>` etc. exists, splicing before `</Types>` once
per missing extension. Manifest (odp): splice a
`<manifest:file-entry manifest:full-path="Pictures/…" manifest:media-type="…"/>`
before `</manifest:manifest>`.

### Serialiser deltas (extending SPEC-081's normative rules)
- `p:pic` for *new* images: `p:nvPicPr` scaffold, `a:blip r:embed="rIdN"`
  where `rIdN` is appended after the highest existing id in the rebuilt
  slide `.rels`; `a:stretch/a:fillRect` body; geometry as usual.
- `draw:image` for odp: `xlink:href="Pictures/…"` +
  `xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"`.
- Ungrouped slides simply serialise the promoted frames — no group
  element remains.
- The shape serialiser's bare `<a:avLst/>` shortcut became a REAL
  adjustment list with SPEC-096: stored adjustment fractions emit one
  `a:gd name="adj…" fmla="val N"` guide each (and odp's
  `draw:enhanced-geometry` gains `draw:modifiers`).

### View wiring
`src/views/presentation/useFrameManipulation.ts` (drag/resize/nudge/
guides), `useFrameInsertion.ts` (gesture arming + picker), extending the
SPEC-081 selection composable; overlay components
`components/presentation/FrameSelectionOverlay.vue` (handles + guides)
and `ShapeGalleryPopover` reusing `ShapeGallery.vue` if it is
product-neutral (it lives in `src/components/` and renders from
`core/shapeIcons` — verify its imports; if it touches text-product
modules, build a slim presentation gallery over `core` and file the
duplication note). Pointer interactions must have touch equivalents per
SPEC-068's conventions (pointer events, no hover-only affordances).
Style popover: `FrameStylePopover.vue` using the design-system colour
picker primitives. The drag-engage threshold (`DRAG_ENGAGE_CM`, 3px at
96dpi) started as a function-local const in `useFrameManipulation.ts`;
SPEC-096 hoisted it to a module-scope export of `stageMath.ts` (same
value, no behaviour change) so the marquee gesture shares it.

### Tests
- `presentation/frames.spec.ts`: every command — clamping, aspect rules,
  z-order maths, ungroup promotion (incl. foreign children), media
  minting against colliding names, fill/stroke edits.
- Serialiser specs: new-image pptx/odp golden output; Content-Types and
  manifest splices; rels id allocation.
- Round trip: insert (text box + shape + image) → save → reopen →
  geometry/z-order/style fidelity in both formats; untouched-part byte
  identity; from-scratch deck with an image re-opens.
- View specs: gesture placement maths, guide activation, keyboard
  paths, popover state.
- e2e per AC (fixture PNG checked in under `e2e/fixtures/`; the picker
  is driven via Playwright's `setInputFiles`).

### Specs this changes
- SPEC-081's Tab-cycle and selection-box notes now cover all frame
  kinds; its toolbar overflow section hosts the new Insert controls.
- SPEC-079's `GroupFrame.xml`-is-immutable rule gains its one sanctioned
  exit (ungroup drops it); the read-only "children exist for rendering
  only" comment is updated by implementation.
