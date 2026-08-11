---
id: SPEC-051
title: Drawing shapes and text boxes
status: draft
area: structure
depends_on: [SPEC-024]
---
## Why
SPEC-024 put the first floating object on the page and built the media
pipeline under it. Shapes are the rest of what people draw: a box around a
warning, an arrow pointing at a screenshot, a pull quote in a text box, a
callout on a diagram. They are also among the most common things in a real
`.docx` that a browser editor shows as nothing at all.

This spec has a second job. SPEC-024 anchored exactly one kind of object and
had no reason to generalise. Charts (SPEC-052) and everything else that floats
need the same position, wrap, z-order and grouping behaviour, and if each
object type invents its own the layout engine ends up with three of everything.
So this spec is the **normative definition of the anchored drawing object**,
lifting the fields SPEC-024's `Image` already carries into one interface that
`Image`, `Shape` and later `Chart` all extend — additively, without renaming a
single field that ships today.

## Acceptance criteria
- [ ] "Insert → Shape" opens a gallery of exactly the geometries in "Shape catalogue" below, grouped as Lines, Basic shapes, Arrows and Callouts
- [ ] Choosing a shape puts the canvas in draw mode: dragging draws it at that size, a single click places it at a default 3cm × 2cm, and Shift constrains to a square/circle or, for lines, to 45° increments
- [ ] "Insert → Text box" inserts a rectangle with no fill and no outline, already in text-editing mode
- [ ] Clicking selects a shape and shows eight resize handles, a rotation handle and a floating shape toolbar; which shape a click selects is decided by "Selection and hit testing" below, not by bounding boxes
- [ ] The shape toolbar sets fill (a solid colour or none), outline (colour, width, and solid/dashed/dotted or none), and rotation in whole degrees, each rendering live
- [ ] Double-clicking a shape puts the caret inside it; its text accepts every character and paragraph feature body text does, wraps at the shape's inner width, and honours a top/middle/bottom vertical alignment
- [ ] Text that does not fit is clipped rather than shrunk, and the shape shows an overflow indicator while it is selected
- [ ] A selected shape offers the wrap modes and anchor references in "Anchoring and wrap" below, shows an anchor marker on the paragraph it belongs to, and re-anchors to the nearest paragraph when dragged
- [ ] Bring to front, bring forward, send backward and send to back reorder overlapping objects, and the painted result matches that order
- [ ] Selecting two or more objects offers Group and Ungroup; a group moves, resizes and rotates as one unit and preserves its members' relative positions
- [ ] Arrow keys nudge a selected shape by 1pt and Ctrl/Cmd+arrow by 0.1pt; Delete removes it; Ctrl/Cmd+D duplicates it offset by 0.5cm
- [ ] Every shape carries alt text, set from the shape toolbar and round-tripped in both formats, and is picked up by SPEC-043's "image without alt text" detection with no change to that spec
- [ ] Shapes are painted on the page's canvas layer and each one gets the focusable ARIA proxy element described in "Rendering and test surface" below, so keyboard users and e2e tests can reach a shape without the canvas
- [ ] Shapes round-trip through `.docx` (`wp:inline`/`wp:anchor` → `wps:wsp` with `a:prstGeom`, `a:solidFill`, `a:ln`, `wps:txbx`) and `.odt` (`draw:frame` with `draw:custom-shape`/`draw:enhanced-geometry` and `draw:text-box`), keeping geometry, size, rotation, fill, outline, wrap, z-order, grouping and text
- [ ] A shape whose geometry is outside the catalogue is preserved per SPEC-008, painted as a dashed placeholder at its true bounds with its text, selectable and movable but not re-geometried, and written back unchanged
- [ ] Shapes appear in print and PDF export (SPEC-034) at their on-screen positions
- [ ] Opening the fixtures below renders every catalogue geometry, a grouped pair, a text box and one unsupported geometry as the placeholder

## Out of scope
SmartArt and diagram layouts. Shadows, glows, reflections, bevels and anything
3-D. Gradient, picture and pattern fills. Freeform curves and edit-points.
WordArt and text effects. Connectors that stay attached to the shapes they
join. Aligning and distributing multiple objects. "Set as default shape".
Cropping a shape to a picture. Nested groups and Word's drawing-canvas
container (`wpc:wpc`) — both are preserved per SPEC-008 and rendered as a
placeholder group that moves but does not ungroup, never created by us.

## Reference

### Document model — the implemented one
`specs/002-create-blank-document.md` defines the model normatively and now
describes the shipped `src/lib/documentModel.ts` directly. **Read it before
starting.** The facts this spec rests on:

- `DocumentModel` holds `sections: Section[]`, whose blocks are
  `Block = Paragraph | Table | Image | BreakBlock | TableOfContents`,
  discriminated by a `type` string literal. **An image is a block, not an
  inline** — `Image.type === "image"`.
- There is no `Inline` union and no `LineBreak` node. A `Paragraph` holds
  `runs: Run[]`; a `Run` is `{ text, formatting, field?, note?, styleId?,
  changes? }`, and tabs and forced line breaks are literal `"\t"` and `"\n"`
  inside `Run.text`.
- Lengths are **centimetres** and type sizes **points** (`Image.widthCm`,
  `ParagraphFormatting.indentStartCm`, `spaceBeforePt`).
- `src/lib/equality.ts`'s `IGNORED_KEYS` drops every key named `id` from
  structural comparison, so a reference that must survive a round-trip test is
  never called `id` — the convention `NoteRef.noteId` and `BookmarkAnchor`
  already follow.

### Anchored drawing object — normative
SPEC-024 shipped `Image` as a block carrying its own size, wrap and alt text.
Shapes need all three plus offset, z-order, rotation and grouping, and so will
charts (SPEC-052). This spec lifts the shared set into one interface **that
`Image` extends**, so every existing field keeps its exact path — `Image.widthCm`,
`Image.wrap` and `Image.altText` are untouched and no shipped call site changes:

```ts
// src/lib/documentModel.ts
export type Block =
  | Paragraph | Table | Image | BreakBlock | TableOfContents
  | Shape                                    // added by this spec

/** Everything a block needs in order to sit somewhere other than in the text
 *  flow. Extended by Image (SPEC-024), Shape (here) and Chart (SPEC-052). */
export interface DrawingPlacement {
  widthCm: number
  heightCm: number
  wrap: ImageWrapMode                        // widened below
  /** `wp:docPr/@descr` (.docx) / `svg:desc` (.odt). */
  altText?: string
  /** Offset from the reference edges; absent means the wrap mode's default. */
  offsetXCm?: number
  offsetYCm?: number
  relativeToH?: "page" | "margin" | "column"
  relativeToV?: "page" | "margin" | "paragraph"
  /** Ascending = nearer the reader; negative renders behind body text. */
  zOrder?: number
  /** Whole degrees clockwise, 0–359. */
  rotationDeg?: number
  /** Group membership as an ordinal — deliberately not named `id`, per the
   *  IGNORED_KEYS rule above. */
  groupOrdinal?: number
}

export interface Image extends DrawingPlacement {
  type: "image"
  mediaId: string                            // → DocumentModel.media
  naturalWidthPx: number
  naturalHeightPx: number
  changes?: BlockTrackedChanges
}

export interface Shape extends DrawingPlacement {
  type: "shape"
  geometry: ShapeGeometry                    // the closed set in "Shape catalogue"
  /** "#rrggbb"; absent means no fill. */
  fill?: string
  outline?: { color: string; widthPt: number; dash: "solid" | "dash" | "dot" }
  /** Text inside the shape. Absent for a shape that holds none, e.g. a line.
   *  `blocks` so a shape can hold lists and tables, not only paragraphs. */
  text?: { blocks: Block[]; vAlign: "top" | "middle" | "bottom" }
  changes?: BlockTrackedChanges
}

export function isShape(block: Block | undefined): block is Shape {
  return block?.type === "shape";
}
```

Three things this is deliberate about:

**`groupOrdinal` is an ordinal, not an id.** `IGNORED_KEYS` drops every key
named `id`, so a group expressed as a shared `id` would vanish from
`equalStructural` and a broken group would round-trip green — the same trap
`NoteRef.noteId` is named around.

**Rotation is whole degrees, not OOXML's 60000ths.** The model already stores
human units everywhere else (cm, pt), and both writers already convert on the
way out; storing one field in a format-native unit to save a multiplication
would be the only exception in the model.

**A floating block still lives in the block list**, in document order, at the
position it is anchored to. `relativeToH`/`relativeToV` only say what the
offsets are measured from. That is where `Image` already sits, and a separate
anchor index would be a second copy of a fact the block list already holds.

### Shape catalogue
The closed set. Anything else is preserved, not approximated.

| Group | Editor shape | `.docx` `a:prstGeom/@prst` | `.odt` |
| --- | --- | --- | --- |
| Lines | Line | `line` | `draw:line` |
| Lines | Arrow | `straightConnector1` + `a:tailEnd` | `draw:line` with `draw:marker-end` |
| Basic | Rectangle | `rect` | `draw:custom-shape`, `draw:type="rectangle"` |
| Basic | Rounded rectangle | `roundRect` | `round-rectangle` |
| Basic | Ellipse | `ellipse` | `ellipse` |
| Basic | Triangle | `triangle` | `isosceles-triangle` |
| Basic | Diamond | `diamond` | `diamond` |
| Basic | Star (5-point) | `star5` | `star5` |
| Arrows | Right arrow | `rightArrow` | `right-arrow` |
| Arrows | Left arrow | `leftArrow` | `left-arrow` |
| Arrows | Up/down arrow | `upArrow` / `downArrow` | `up-arrow` / `down-arrow` |
| Callouts | Rectangular callout | `wedgeRectCallout` | `rectangular-callout` |
| Callouts | Rounded callout | `wedgeRoundRectCallout` | `round-rectangular-callout` |

A text box is not a fourteenth geometry: it is `rect` with `fill: 'none'`,
`outline: 'none'` and a `textFrame`. It writes as `wps:txbx` in `.docx` and as
`draw:text-box` in `.odt`, and reads back as a rectangle with those properties,
so the two paths cannot drift.

`.odt`'s `draw:enhanced-geometry` carries a formula-based path. We write the
`draw:type` values above and their standard enhanced-geometry bodies; on read,
a `draw:type` in the table maps to the catalogue geometry and anything else is
an unsupported geometry.

### Anchoring and wrap
The shipped union is
`ImageWrapMode = "inline" | "square" | "behind" | "infront"` (SPEC-024). This
spec adds **`"topAndBottom"`** to it, and because `Image` and `Shape` share the
type, it applies to images too (see "Specs this changes").

Wrap is computed against the object's **bounding box**, not its geometry. Word
and Writer both offer tight/through wrapping against the outline; we do not,
and a rotated ellipse wrapping as a rectangle is the honest, predictable
behaviour rather than a half-implemented approximation.

`"behind"` and `"infront"` take the object out of the text flow entirely and
are ordered against body text by `zOrder`: negative renders behind, zero or
above in front, and absent means zero. Writers map it to
`wp:anchor/@relativeHeight` and to `draw:frame/@draw:z-index`, which are the
same concept in both formats.

### Selection and hit testing
A click selects the **topmost object whose painted region contains the point**,
where painted region means:

- a shape with a fill: the filled geometry, plus its outline;
- a shape with `fill: 'none'`: only its outline, widened to a 3pt hit
  tolerance so a hairline rectangle is still clickable;
- a shape with a `textFrame` holding text: also the text block's bounds, so an
  unfilled text box can be selected by clicking its text.

Bounding-box hit testing is explicitly wrong here: overlapping shapes are the
normal case in a diagram, and a bounding box makes the shape underneath
unreachable.

### Rendering and test surface
`CLAUDE.md` puts complex shapes on HTML Canvas, and geometry, rotation and
z-order are exactly that case. But a canvas has no DOM, and `CLAUDE.md` also
requires e2e tests to select by role and test-id, and SPEC-043 requires the
editing screen to pass an automated accessibility scan. Canvas alone fails
both.

So each drawing paints on the page's canvas layer **and** owns one absolutely
positioned, transparent DOM element over its bounds carrying
`data-testid="drawing-<index>"`, `role="img"` (or `role="group"` when it has a
`textFrame`), `aria-label` set from `altText`, and `tabindex="0"`. Focus,
keyboard nudging, selection and every e2e assertion go through that element;
the canvas only paints. A shape's text is real DOM inside the proxy — it has to
be, because it is editable — and only the geometry, fill and outline are
painted.

### Preservation of what we do not model
Defined normatively in `specs/008-round-trip-preservation.md` and implemented in
`src/lib/opaque.ts`. An unsupported geometry is preserved as unmodelled XML
compared with `canonicalizeXml` (`src/lib/equality.ts`), and a cross-format save
that cannot carry it produces a `DropRecord` in the writer's return value:

```ts
// src/lib/opaque.ts
export type DropKind = "chart" | "contentControl" | "customXml" | "embeddedObject"
  | "macro" | "opaqueBlock" | "opaqueInline" | "opaquePart"
  | "runShading" | "svgImage" | "trackedPropertyChange";
export interface DropRecord { kind: DropKind; location: string; reason: string }
export interface WriteResult { bytes: Uint8Array; dropped: DropRecord[] }
```
`DropKind` has already been widened by SPEC-013 (`runShading`), SPEC-024
(`svgImage`) and SPEC-039 (`trackedPropertyChange`); this spec widens it once
more, with `"shapeGeometry"`, for a geometry the target format cannot express.
It adds no preservation mechanism of its own.

### Specs this changes
- `specs/024-images-insert-resize-and-wrap.md` — `Image` gains
  `extends DrawingPlacement`. Every field it has today keeps its name and path,
  so this is additive: no shipped call site, test or fixture assertion changes.
  Its wrap set gains `topAndBottom`, widening the shipped
  `ImageWrapMode = "inline" | "square" | "behind" | "infront"`, and SPEC-024's
  user-visible behaviour is otherwise untouched.
- `specs/043-accessibility-checker.md` — no change. Shapes carry `altText` in
  the same field images do, so the existing detection covers them; this is
  recorded here so nobody adds a second detection for it.

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalBytes`,
`canonicalizeXml` and `firstStructuralDifference`. SPEC-002 also defines
`equalSemantic` (`equalStructural(resolve(a), resolve(b))`), **which is not
implemented** — a spec that needs it builds it. This one does not:
- **`equalStructural`** — for the same-format round trip of the shape fixtures,
  and for asserting that a drag, a nudge or a z-order change altered nothing
  else in the document.
- For the cross-format round trip, assert `equalStructural` between the source
  model with the writer's `dropped` records removed and the re-opened model.
  Masking is needed because the catalogues differ in expressive power: `.odt`
  has no per-geometry tail marker, and an unsupported `.docx` geometry cannot
  be carried into `.odt` at all. Apply the mask as a transform over `dropped`,
  never as a hand-rolled subset comparison.

`equalBytes` applies only to preserved parts, through SPEC-008's existing rule;
this spec adds no new use of it.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). Adds
`fixtures/docx/word-shapes.docx` and `fixtures/odt/lo-shapes.odt`, each
containing every catalogue geometry, one grouped pair, one text box with two
paragraphs and a bulleted list, one rotated shape, one `behind` and one
`square` wrap, and one geometry outside the catalogue. Manifest entries name
the geometry, size, rotation and wrap that must survive.
