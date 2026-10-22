---
id: SPEC-069
title: Shape catalogue expansion and adjustment handles
status: draft
area: structure
depends_on: [SPEC-008, SPEC-051, SPEC-057]
---
## Why
`SHAPE_CATALOGUE` in `src/lib/core/shapes.ts` holds **fourteen** geometries:
line, arrow, rect, roundRect, ellipse, triangle, diamond, star5, four block
arrows and two callouts. ECMA-376's `ST_ShapeType` enumerates the preset
geometries a conforming application is expected to render, and ODF's
`draw:enhanced-geometry` carries a comparable set; both run to well over a
hundred. Fourteen is not a catalogue, it is a sample.

The cost is not "fewer shapes to draw". It is that a real document opens
wrong. Any diagram built from flowchart boxes, any callout that is not one
of our two, any hexagon, cylinder, chevron or banner arrives as a
SPEC-008-preserved `unsupportedGeometryXml` and paints as a **dashed
placeholder rectangle** — the document round-trips faithfully and looks
broken. A flowchart is the single most common drawing in business
documents, and not one of its shapes is in the catalogue.

There is a second, subtler gap that the shipped model cannot express at
all: **adjustment values**. A rounded rectangle's corner radius, a block
arrow's head width, a callout's tail position and a star's point depth are
all parameters in both formats — `a:avLst`/`a:gd` in OOXML,
`draw:modifiers` in ODF — and the model has no field for any of them. So
even the fourteen shapes we do support lose their tuning: a Word rounded
rectangle with a large radius reopens with our fixed radius, and the yellow
adjustment handle every other editor offers has nowhere to write to. That
makes adjustment values part of this spec rather than a follow-up — adding
forty geometries without them would multiply the fidelity loss by forty.

## Acceptance criteria
- [ ] The catalogue grows to the geometries listed in "The expanded catalogue", organised into the groups named there, each with a `prstGeom` name, a `draw:type` name and a label
- [ ] `ShapeGallery.vue` renders the expanded catalogue grouped and scrollable, with a search field filtering by label and by alias, and stays within the Insert-menu size constraint the gallery already respects
- [ ] Every catalogue geometry has a painted outline, a fill region, a hit region and a gallery icon, all derived from one definition per geometry — no geometry is drawable but not hit-testable, or vice versa
- [ ] `DrawingPlacement`-carrying `Shape` gains `adjustments?: number[]`, per "Adjustment values", normalised to the format-neutral 0..1-style fractions defined there
- [ ] A shape whose geometry has adjustments shows one draggable adjustment handle per adjustment, visually distinct from resize handles, constrained to its declared range, live during the drag, committing one undo entry on release
- [ ] Adjustments round-trip: `.docx` `a:avLst/a:gd@fmla` values are read and written for every catalogue geometry that declares them, and `.odt` `draw:modifiers` likewise, asserted on the written XML for at least one geometry per group
- [ ] A geometry read with adjustment values outside our declared range is clamped for rendering and **preserved verbatim** on a same-format save, so tuning we do not model is never silently rewritten
- [ ] Opening a `.docx` or `.odt` containing any expanded-catalogue geometry renders the real shape, not the SPEC-051 dashed placeholder, and the placeholder path continues to cover everything still outside the catalogue
- [ ] Saving an expanded-catalogue shape to the other format writes the corresponding preset name, with a `DropRecord` only where the target format has no equivalent, per "Cross-format mapping"
- [ ] A shape carrying `unsupportedGeometryXml` from an earlier version of the app whose geometry is now in the catalogue is **not** re-interpreted: preserved XML stays preserved, and the catalogue applies to newly-read documents, per "No retroactive reinterpretation"
- [ ] Text inside an expanded-catalogue shape lays out within that geometry's text rectangle, declared per geometry, rather than in its bounding box
- [ ] Every geometry is covered by a unit test asserting its polygon or path is closed, within its bounds, and non-degenerate at the minimum shape size (0.2cm, `MIN_SHAPE_SIZE_CM`)
- [ ] Hit testing on the expanded geometries uses the existing rules unchanged: filled shapes hit inside the painted region, unfilled ones within `HIT_TOLERANCE_CM` of the outline
- [ ] Fixtures containing a flowchart, a banner, a chevron sequence, a cylinder and an adjusted callout open, render, save and reopen `equalStructural` in both formats

## Out of scope
Freeform, curve and scribble drawing, and editing a shape's points
(SPEC-071). Connectors (SPEC-070). Gradient and picture fills, shadows and
3-D effects (SPEC-073). SmartArt and Word's drawing canvas (`wpc:wpc`) —
still preserved, never created, per SPEC-051. Multiple adjustment handles
whose ranges depend on each other beyond the per-adjustment clamp declared
here. Any change to anchoring, wrap, z-order, grouping or rotation —
SPEC-051 and SPEC-057 own those and are untouched. Every `ST_ShapeType`
value: this spec expands the catalogue substantially and deliberately stops
short of exhaustive, keeping SPEC-008 preservation as the answer for the
rest.

## Reference

### What ships today
```ts
// src/lib/core/shapes.ts
export const SHAPE_CATALOGUE: {
  geometry: ShapeGeometry;
  group: "Lines" | "Basic shapes" | "Arrows" | "Callouts";
  label: string;
}[] = [ /* 14 entries */ ];

export const LINE_GEOMETRIES: ReadonlySet<ShapeGeometry>;   // line, arrow
export const MIN_SHAPE_SIZE_CM = 0.2;
export const DEFAULT_AREA_SIZE_CM = { widthCm: 3, heightCm: 2 };
export const DEFAULT_NEW_SHAPE_FILL = "#4472c4";
export interface ShapeVisual { geometry; fill?; outline?; rotationDeg?; unsupportedGeometryXml?; }
```

```ts
// src/lib/core/shapeGeometry.ts
export function shapePolygon(geometry: ShapeGeometry): Point[] | null;
export function roundRectRadius(width: number, height: number): number;
export function roundedCalloutOutline(…): CalloutOutlineStep[];
export function lineSegment(width: number, height: number): [Point, Point];
export function scalePolygon(polygon: Point[], width: number, height: number): Point[];
export function pointInPolygon(point: Point, polygon: Point[]): boolean;
export function pointInEllipse(point: Point, width: number, height: number): boolean;
export function pointHitsShape(point: Point, opts: ShapeHitOptions): boolean;
export function shapeClipPath(geometry: ShapeGeometry, filled: boolean): string | null;
export const HIT_TOLERANCE_CM = (3 * 2.54) / 72;   // 3pt
```

Format mapping is two flat records plus a fallback:

```ts
// src/lib/docx/shape.ts
const PRST_TO_GEOMETRY: Record<string, ShapeGeometry>;   // 14 + straightConnector1
const GEOMETRY_TO_PRST: Record<ShapeGeometry, string>;
// src/lib/odf/shape.ts
const GEOMETRY_TO_DRAW_TYPE: Record<ShapeGeometry, string>;
//   rect -> "rectangle", roundRect -> "round-rectangle",
//   triangle -> "isosceles-triangle", star5 -> "star5",
//   wedgeRectCallout -> "rectangular-callout", …
```

Anything not in those maps takes SPEC-051's preservation path:
`Shape.unsupportedGeometryXml` holds the verbatim `a:prstGeom`/`a:custGeom`
or `draw:enhanced-geometry`, `geometry` becomes a `"rect"` bounding-box
stand-in, and the shape paints as a dashed placeholder.

### One definition per geometry
Four parallel structures — catalogue entry, polygon function, two name maps
— is what makes fourteen geometries feel like a lot. Expanding to forty
demands collapsing them into one table, in the product-neutral core layer
(memory: `src/lib/core` is shared with a future Slides host and its
boundary is enforced by `boundary.spec.ts`):

```ts
// src/lib/core/shapeCatalogue.ts
export interface ShapeDefinition {
  geometry: ShapeGeometry;
  group: ShapeGroup;
  label: string;
  /** Extra search terms for the gallery's filter. */
  aliases?: string[];
  /** OOXML preset name (a:prstGeom/@prst). */
  prst: string;
  /** ODF enhanced-geometry type (draw:type). */
  drawType: string;
  /** Outline in a unit square, given the shape's adjustments. Points are
   *  fractions of width/height, so scaling is scalePolygon's job. A
   *  geometry needing curves returns path steps instead. */
  outline(adjustments: readonly number[]): ShapeOutlinePath;
  /** Declared adjustments: default and clamp range, in the same order as
   *  Shape.adjustments and as the format's own adjustment list. */
  adjustments?: ShapeAdjustmentSpec[];
  /** Where text lays out inside the shape, as unit-square fractions;
   *  defaults to the full box. */
  textRect?(adjustments: readonly number[]): { x: number; y: number; w: number; h: number };
  /** True for zero-area geometries drawn as a segment. */
  isLine?: boolean;
}

export interface ShapeAdjustmentSpec {
  /** Handle name, for the accessible label: "corner radius". */
  label: string;
  default: number;   // 0..1
  min: number;
  max: number;
  /** Which axis the handle moves along, for the drag mapping. */
  axis: "x" | "y" | "radial";
}

export const SHAPE_DEFINITIONS: ShapeDefinition[];
export function shapeDefinition(geometry: ShapeGeometry): ShapeDefinition | undefined;
```

`SHAPE_CATALOGUE`, `shapePolygon`, `PRST_TO_GEOMETRY`,
`GEOMETRY_TO_PRST` and `GEOMETRY_TO_DRAW_TYPE` are all **derived** from
`SHAPE_DEFINITIONS` rather than maintained beside it, so a new geometry is
one entry. The existing exports keep their names and signatures where other
modules import them — this is a refactor behind a stable surface, not an
API change, and the shipped shape tests must pass unmodified.

`ShapeOutlinePath` generalises today's `Point[]` to a small step list
(`moveTo`/`lineTo`/`arcTo`/`quadTo`/`close`) so cylinders, arcs and rounded
forms stop needing bespoke functions like `roundedCalloutOutline`. The
existing polygon geometries return an all-`lineTo` path, and
`shapePolygon()` keeps working by flattening one.

### The expanded catalogue
Groups and geometries, with their `prst` / `draw:type` names. Existing
entries are unchanged; everything else is new.

**Lines** — `line`, `arrow` *(unchanged)*

**Basic shapes** — `rect`, `roundRect`, `ellipse`, `triangle`, `diamond`,
`star5` *(unchanged)*, plus:
`rtTriangle` (right triangle / `right-triangle`), `parallelogram`,
`trapezoid`, `pentagon` (`regular-pentagon`), `hexagon`, `octagon`,
`plus` (`cross`), `can` (cylinder / `can`), `cube`, `heart`, `cloud`,
`donut` (`ring`), `pie`, `arc`, `chord`, `bracketPair` (`bracket-pair`),
`bracePair` (`brace-pair`), `plaque`, `frame`

**Block arrows** — `rightArrow`, `leftArrow`, `upArrow`, `downArrow`
*(unchanged)*, plus: `leftRightArrow`, `upDownArrow`, `quadArrow`,
`bentArrow`, `uturnArrow`, `curvedRightArrow`, `homePlate` (pentagon
arrow / `pentagon-right`), `chevron`, `notchedRightArrow`,
`stripedRightArrow`, `circularArrow`

**Flowchart** — `flowChartProcess`, `flowChartAlternateProcess`,
`flowChartDecision`, `flowChartInputOutput`, `flowChartPredefinedProcess`,
`flowChartInternalStorage`, `flowChartDocument`, `flowChartMultidocument`,
`flowChartTerminator`, `flowChartPreparation`, `flowChartManualInput`,
`flowChartManualOperation`, `flowChartConnector`, `flowChartOffpageConnector`,
`flowChartPunchedCard`, `flowChartMagneticDisk`, `flowChartDelay`,
`flowChartExtract`, `flowChartMerge`, `flowChartSort`

**Stars and banners** — `star4`, `star6`, `star8`, `star10`, `star12`,
`star16`, `star24`, `star32`, `irregularSeal1`, `irregularSeal2`,
`ribbon`, `ribbon2`, `verticalScroll`, `horizontalScroll`, `wave`,
`doubleWave`

**Callouts** — `wedgeRectCallout`, `wedgeRoundRectCallout` *(unchanged)*,
plus: `wedgeEllipseCallout`, `cloudCallout`, `borderCallout1`,
`borderCallout2`, `borderCallout3`, `accentCallout1`

Each entry's `draw:type` uses ODF's own name for the same shape
(`flowchart-process`, `flowchart-decision`, `star6`, `left-right-arrow`,
`round-callout`, …) rather than a transliteration of the OOXML one; the
mapping is per-entry data in the definition table precisely so the two
vocabularies never have to agree.

Implementation may land the groups incrementally — flowchart first, since
it is the highest-value gap — as long as every shipped geometry satisfies
every AC. A geometry that is in the gallery but not round-tripping is not
shipped.

### Adjustment values
```ts
// src/lib/documentModel.ts
export interface Shape extends DrawingPlacement {
  …
  /** Geometry adjustment values, in the order the geometry's
   *  ShapeDefinition declares them, each a fraction in that adjustment's
   *  declared range. Absent means every adjustment is at its default.
   *  Round-trips as a:avLst/a:gd (.docx) and draw:modifiers (.odt). */
  adjustments?: number[];
}
```

Fractions, not format-native units, because the two formats disagree:
OOXML expresses adjustments as guide formulas in shape-relative
hundred-thousandths or as absolute EMU depending on the geometry, and ODF's
`draw:modifiers` is a space-separated list in the shape's own coordinate
space. Normalising to a fraction at the reader boundary keeps the model
format-neutral — the same choice the model already makes for lengths
(centimetres) and rotation (degrees).

Each `ShapeDefinition` owns the conversion for its own geometry, so the
denormalisation lives beside the outline that uses it, and is unit-tested
per geometry against a fixture value from each format.

A shape whose stored adjustment count does not match its definition's is
rendered with the defaults for the missing entries and preserved verbatim
on a same-format save.

### Adjustment handles
One handle per declared adjustment, drawn as a small diamond in the
selection colour (distinct from the square resize handles), positioned by
the geometry's own definition, constrained to its declared axis and range.
Drag previews live; release commits one undo entry labelled
`"adjust shape"`. Keyboard: with the shape selected, Tab cycles focus
through its adjustment handles, arrows nudge by 1% of the range, Shift by
5%, and each handle announces its declared label and current value.

Adjustment handles are hidden for a preserved shape (nothing to write to)
and for members of a sealed container (SPEC-057).

### Cross-format mapping
Every catalogue entry declares both names, so most geometries cross over
directly. Where one format has no equivalent — a handful of the flowchart
and seal shapes differ between the vocabularies — the writer emits the
closest available preset **and** reports a `DropRecord` of the existing
`"shapeGeometry"` kind, so the conversion report (SPEC-008) tells the user
what changed. No new `DropKind` is added.

### No retroactive reinterpretation
A document already in the library may hold a shape with
`unsupportedGeometryXml` set for a geometry this spec now supports. That
XML is preserved verbatim and stays preserved: re-parsing stored models is
outside this spec, the preserved bytes are exactly what SPEC-008 promises,
and a same-format save of that document must still be byte-identical for
the preserved part. Reopening the *source file* is what picks up the new
catalogue. Stating this prevents a plausible-looking "upgrade on load" that
would break SPEC-008's guarantee.

### Specs this changes
- `specs/051-drawing-shapes-and-text-boxes.md` — its fourteen-geometry
  catalogue is superseded by `SHAPE_DEFINITIONS`; its preservation,
  anchoring, wrap, selection, hit-testing, text and rendering rules are
  unchanged and re-asserted over the wider set. Its "Shape catalogue"
  section points here.
- `specs/057-group-and-ungroup-drawings.md` — unchanged. A container's
  members may now be any catalogue geometry; sealing still applies to
  anything outside it.
- `specs/043-accessibility-checker.md` — unchanged; a shape is a drawing
  needing alt text whatever its geometry.
