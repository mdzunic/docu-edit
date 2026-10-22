---
id: SPEC-071
title: Freeform shapes, curves and edit points
status: draft
area: structure
depends_on: [SPEC-051, SPEC-069]
---
## Why
SPEC-069 widens the preset catalogue; it does not make the catalogue
*escapable*. Every shape this product can create comes from a fixed list,
and a document containing anything drawn by hand — a highlighted region
around a paragraph, an org-chart bracket, a hand-annotated diagram, a
signature — arrives as `a:custGeom` / a path-bearing
`draw:enhanced-geometry`, lands on SPEC-051's preservation path, and paints
as a **dashed placeholder rectangle**. Custom geometry is not an exotic
case: it is what every drawing tool produces the moment a user reaches for
the pen, and it is also what many "preset" shapes become after another
application flattens them.

So this spec closes the catalogue's escape hatch in both directions:
render the custom geometry that documents already contain, and let a user
draw one. Those are the same feature — a path model with an editor on top —
and separating them would mean shipping a renderer for paths nobody can
create, or an editor writing paths the reader cannot read back.

Edit points is the third face of the same thing. Once a shape's outline is
a path, "convert this preset to a freeform and move its vertices" is a
model operation rather than a new subsystem, and it is the operation that
makes the catalogue's remaining gaps survivable: whatever shape is missing,
a user can draw it.

## Acceptance criteria
- [ ] A `ShapePath` model exists per "The path model" — a list of subpaths of move/line/cubic/close steps in unit-square coordinates — and `Shape` gains `path?: ShapePath`, which takes precedence over `geometry` when present
- [ ] A shape read from `.docx` `a:custGeom/a:pathLst` or from `.odt` `draw:enhanced-geometry/@draw:enhanced-path` is parsed into `ShapePath` and **rendered as its real outline**, replacing the dashed placeholder for the path cases per "What stays preserved"
- [ ] The Insert gallery gains a Freeform group with three tools: **Freeform** (click to place vertices, double-click or Enter to finish, click the origin to close), **Curve** (the same, with smooth cubic segments through the placed points) and **Scribble** (press and drag, sampled and simplified per "Scribble simplification")
- [ ] Escape during a draw cancels it entirely, leaving no block and no undo entry; Backspace removes the last placed vertex
- [ ] An unclosed freeform is a line-like shape — outline only, no fill — and a closed one fills per SPEC-051's fill rules
- [ ] `Format → Edit points` (also on the shape context menu and the shape toolbar) enters point-edit mode for the selected shape; a preset shape entering it is converted to an equivalent `path` first, with one undo entry labelled "convert to freeform"
- [ ] In point-edit mode every vertex shows a handle: dragging moves it, double-clicking toggles it between corner and smooth, Delete removes it, and clicking a segment adds a vertex there
- [ ] A smooth vertex shows two control-point handles; dragging one adjusts the curve, and on a smooth vertex the opposite handle mirrors, while on a corner vertex the two move independently
- [ ] Point-edit mode is left by Escape or by clicking outside, commits **one** undo entry for the whole session's edits, and restores the previous selection
- [ ] Point edits renormalise the shape's bounding box: after any edit the path's extent is re-fitted to the shape's `widthCm`/`heightCm` and offsets with no visual jump, so a dragged vertex outside the old box enlarges the box rather than being clipped
- [ ] A path shape hit-tests against its real outline — inside the filled region, or within `HIT_TOLERANCE_CM` of the stroke for an unfilled one — using the existing `pointHitsShape` contract widened to paths
- [ ] Rotation, resize, wrap, z-order, grouping, alt text and text frames behave for a path shape exactly as for a preset one; resizing scales the path, it does not re-fit it
- [ ] Path shapes round-trip through `.docx` as `a:custGeom` with `a:pathLst/a:path` (`a:moveTo`, `a:lnTo`, `a:cubicBezTo`, `a:close`) in the path's own coordinate space, asserted on the written XML
- [ ] Path shapes round-trip through `.odt` as `draw:enhanced-geometry` with an `@draw:enhanced-path` command string and a matching `@svg:viewBox`, asserted on the written XML
- [ ] A cross-format save carries the path in both directions with no `DropRecord`, since both formats express the same primitives; a path using a primitive we do not model (arcs, quadratic segments) is converted per "Primitive reduction" and reported
- [ ] Fixtures containing a Word freeform, a Writer enhanced-path shape, a closed curve, an open scribble and a shape with smooth and corner vertices open, render as their true outlines, save and reopen `equalStructural`

## Out of scope
Boolean operations on shapes (union, subtract, intersect). Converting a
shape to a picture, or tracing a picture. Text on a path. Filling a
self-intersecting path with a chosen fill rule beyond the default —
non-zero winding is used and stated. Pressure or velocity-sensitive
scribble strokes (SPEC-068 puts stylus input out of scope). Ink annotation
as a review feature. Editing the points of a shape inside a sealed group
(SPEC-057). Arcs and quadratic Bézier segments as *model* primitives —
read, converted, and stated, per "Primitive reduction".

## Reference

### What ships today
`Shape.unsupportedGeometryXml` (`src/lib/documentModel.ts`) holds the
verbatim `a:custGeom` / unrecognised `a:prstGeom` / unrecognised
`draw:enhanced-geometry`, `geometry` becomes a `"rect"` bounding-box
stand-in, and the shape paints as a dashed placeholder that cannot be
edited. `src/lib/core/shapeGeometry.ts` provides `shapePolygon`,
`scalePolygon`, `pointInPolygon`, `pointInEllipse`, `pointHitsShape`,
`shapeClipPath` and `HIT_TOLERANCE_CM = (3 * 2.54) / 72`. SPEC-069
generalises the outline of a preset to a step list (`ShapeOutlinePath`) —
this spec's `ShapePath` is that step list, stored on the model instead of
derived from a definition.

### The path model
```ts
// src/lib/core/shapePath.ts — product-neutral, framework-free
/** Coordinates are fractions of the shape's box: 0..1 in each axis
 *  (values outside the range are legal mid-edit and renormalised on
 *  commit). One shared space for every path, so scaling, rotating and
 *  hit-testing a path is the same arithmetic as for a preset outline. */
export type PathStep =
  | { op: "move"; x: number; y: number }
  | { op: "line"; x: number; y: number }
  | { op: "cubic"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: "close" };

export interface ShapeSubpath {
  steps: PathStep[];
  /** True when the subpath ends with a close; fill applies only to closed
   *  subpaths. */
  closed: boolean;
}

export interface ShapePath {
  subpaths: ShapeSubpath[];
  /** Per-vertex smoothness, parallel to the vertex order, so entering
   *  point-edit mode on a re-opened document restores the right handles
   *  instead of inferring them from the control points. */
  smooth?: boolean[];
}

export function pathBounds(path: ShapePath): { x: number; y: number; w: number; h: number };
export function normalisePath(path: ShapePath): { path: ShapePath; offset: { x: number; y: number }; scale: { x: number; y: number } };
export function pointInPath(point: Point, path: ShapePath, w: number, h: number): boolean;
export function distanceToPath(point: Point, path: ShapePath, w: number, h: number): number;
export function pathToSvgD(path: ShapePath, w: number, h: number): string;
export function simplifyPolyline(points: Point[], toleranceCm: number): Point[];
```

```ts
// src/lib/documentModel.ts
export interface Shape extends DrawingPlacement {
  …
  /** A custom outline. When present it is the shape's geometry, and
   *  `geometry` is only a bounding-box stand-in ("rect"), the same
   *  convention `unsupportedGeometryXml` already uses. Round-trips as
   *  a:custGeom (.docx) and draw:enhanced-path (.odt). */
  path?: ShapePath;
}
```

`normalisePath` is what the renormalisation AC calls: it re-fits a path's
extent to 0..1 and returns the offset and scale the caller folds into the
shape's own box, so nothing moves on screen. It is the direct analogue of
SPEC-057's container renormalisation, and it is pure and unit-tested for
the same reason.

### Primitive reduction
Both formats have primitives this model does not carry: OOXML's `a:arcTo`
and `a:quadBezTo`, ODF's `A`/`Q`/`T` enhanced-path commands, and ODF's
formula-driven paths (`draw:equation` referencing modifiers).

- **Arcs and quadratics are converted on read** to cubics, exactly —
  a quadratic has an exact cubic equivalent, and an arc is subdivided to
  within a stated tolerance of one screen pixel at 100% zoom. The
  conversion is pure, unit-tested against known control points, and
  applied at the reader boundary so the model carries only the three
  primitives above.
- **Formula-driven enhanced paths stay preserved.** A path whose commands
  reference `draw:equation` results cannot be flattened without losing the
  parametric behaviour a modifier drag would change, so it keeps the
  SPEC-051 preservation path unchanged.
- A shape whose path was converted reports nothing on a same-format save
  when the preserved XML is re-emitted, and reports a `"shapeGeometry"`
  `DropRecord` when a converted path is written back as cubics — which is
  a real, if small, fidelity loss and is stated rather than hidden.

### What stays preserved
The placeholder path is narrowed, not removed. After this spec a shape
paints as the dashed placeholder only when its geometry is genuinely
unmodellable: a formula-driven enhanced path, a `prstGeom` outside
SPEC-069's catalogue, or geometry carrying attributes the model does not
express (SPEC-057 names flips as one). Everything path-shaped renders.

`unsupportedGeometryXml` and `path` are mutually exclusive on a shape, and
a unit test asserts it: a shape whose path we modelled has nothing to
preserve, and a preserved shape has no path to edit.

Documents already stored with a preserved custom geometry are **not**
reinterpreted on load, for the reason SPEC-069 states: preserved bytes are
what SPEC-008 promises. Reopening the source file picks up the new reader.

### Drawing gestures
- **Freeform** — each click appends a `line` step; a click within 0.2cm of
  the origin closes the subpath and ends the draw; double-click or Enter
  ends it open. A single click and immediate finish creates nothing.
- **Curve** — the same clicks, but the resulting vertices are joined with
  cubics whose control points are derived to pass smoothly through each
  point (a Catmull–Rom to Bézier conversion), and every vertex is marked
  smooth.
- **Scribble** — pointer-down starts, pointer-move samples, pointer-up
  ends. Samples are simplified with `simplifyPolyline` at a tolerance of
  0.03cm and then smoothed into cubics, so a stroke becomes a few dozen
  vertices rather than a few hundred. The tolerance is a stated constant,
  not a magic number in the gesture code, because the size of the saved
  file depends on it.

All three draw on the same canvas overlay SPEC-051's shape draw mode uses,
with the same Escape-cancels contract, and commit one undo entry labelled
`"draw freeform"`.

### Point-edit mode
A modal state on one selected shape, mutually exclusive with text editing:

- Vertices render as small squares (corner) or circles (smooth); control
  points as small filled dots on thin leader lines, shown only for the
  selected vertex and its neighbours.
- Keyboard: Tab and Shift+Tab move between vertices, arrows nudge by
  0.01cm (Shift by 0.1cm), Enter toggles smooth/corner, Delete removes,
  Escape exits. Each vertex announces its index and position.
- The whole session is one undo entry, captured on exit — the same
  "preview live, commit once" rule SPEC-057 applies to a group resize.
- Entering point-edit mode on a preset shape converts it: SPEC-069's
  `ShapeDefinition.outline` produces the path, `geometry` becomes `"rect"`,
  and the conversion is its own undo entry so it can be undone
  independently of the edits that follow.

### Format mapping
**`.docx`** — `a:custGeom` holding `a:pathLst/a:path` with `@w`/`@h` set to
the path's coordinate extent (the model's 0..1 space scaled to a whole-
number grid, conventionally 21600 as Word uses), containing `a:moveTo`,
`a:lnTo`, `a:cubicBezTo` and `a:close`. `a:gdLst`/`a:ahLst`/`a:cxnLst` are
not written; a source file's are preserved only via the untouched-preserved
path, not alongside a modelled one.

**`.odt`** — `draw:enhanced-geometry` with
`@draw:enhanced-path="M … L … C … Z N"` and
`@svg:viewBox="0 0 <w> <h>"` on the same grid, `@draw:type="non-primitive"`.
The trailing `N` terminates the sub-path list, as ODF requires.

Both writers emit the same grid from the same normalised model, so a
cross-format round trip is a coordinate scale and nothing else — which is
what makes the no-`DropRecord` AC achievable.

### Specs this changes
- `specs/051-drawing-shapes-and-text-boxes.md` — its dashed-placeholder
  rule narrows to the cases in "What stays preserved"; its drawing,
  selection, rotation, wrap and text rules apply unchanged to path shapes.
- `specs/069-shape-catalogue-expansion.md` — `ShapeDefinition.outline`
  becomes the source for preset→freeform conversion; its adjustment
  handles are unavailable on a converted shape (it is no longer a preset),
  which is the same behaviour every other editor has and is worth stating.
- `specs/008-round-trip-preservation.md` — no mechanism change; the set of
  content reaching opaque preservation shrinks.
