---
id: SPEC-073
title: Shape styling depth — gradients, effects, arrowheads and presets
status: draft
area: structure
complexity: complex
story_points: 13
depends_on: [SPEC-051, SPEC-069, SPEC-070]
---
## Why
A shape in this product can be one flat colour with one flat outline. The
model says so:

```ts
fill?: string;                                        // "#rrggbb", or nothing
outline?: { color: string; widthPt: number; dash: "solid" | "dash" | "dot" };
```

Everything else a document can say about a shape's appearance — a gradient,
a picture or pattern fill, any transparency at all, a shadow, a glow, a
rounded or mitred line join, an arrowhead on anything other than our one
`arrow` geometry — has nowhere to live. Both formats express all of it
(`a:gradFill`, `a:blipFill`, `a:pattFill`, `a:alpha`, `a:effectLst`,
`a:headEnd`/`a:tailEnd`, `a:round`/`a:bevel`/`a:miter` in OOXML;
`draw:fill="gradient"|"bitmap"|"hatch"`, `draw:opacity`,
`draw:shadow-*`, `draw:marker-*`, `svg:stroke-linejoin` in ODF), so a real
document loses it on open: a gradient-filled banner arrives flat blue, a
shadowed callout arrives shadowless, and a save writes the flattened
version back.

That last part is the sharp edge. SPEC-008's rule is "never discard
document content you did not parse", and for shape *geometry* the product
honours it with `unsupportedGeometryXml`. For shape *styling* it does not:
the fill and outline are parsed into a model that cannot hold what was
there, and the writer emits the reduced version — a silent, permanent
downgrade on every save of every gradient-filled shape in the document.
So this spec is a fidelity fix first and a feature second.

## Acceptance criteria
- [ ] `Shape.fill` widens from a colour string to a `ShapeFill` union per "The fill model" — solid, gradient, picture, pattern or none — with the existing `"#rrggbb"` string accepted and migrated per "Migration"
- [ ] `ShapeOutline` gains transparency, line join, line cap, compound type and arrowheads per "The outline model"
- [ ] A gradient fill supports linear (with an angle) and radial types with two or more stops, each stop carrying a colour, a position and a transparency
- [ ] A picture fill references a media part through the existing image pipeline (SPEC-024), with stretch and tile modes, and round-trips without duplicating the media part
- [ ] Fill and outline transparency are supported as a percentage on solid fills, on every gradient stop, and on outlines, and render correctly over text and over other drawings
- [ ] A shadow effect — outer shadow with colour, transparency, blur, distance and angle — is modelled, rendered and round-tripped; presets for the common shadows are offered
- [ ] Arrowheads apply to **any** line-like shape and to connectors, at either end, with a kind and a size, per SPEC-070's `ArrowheadKind`
- [ ] The shape toolbar's fill and outline controls become pickers that expose the full model: fill kind, gradient editor, transparency slider, line width, dash, join, cap and arrowheads — reachable from the contextual toolbar and from a `Format → Drawing → Shape formatting…` dialog
- [ ] A gallery of **style presets** applies a coordinated fill + outline + effect combination in one click, recorded as one undo entry, per "Style presets"
- [ ] Every styling change records exactly one labelled undo entry; a slider drag previews live and commits once on release
- [ ] Styling applies to a multi-selection and to group members, and applies to a container only by fanning out to its members — which is the point SPEC-057 deliberately deferred, and is resolved here per "Applying to a group"
- [ ] Every property round-trips through `.docx` (`a:gradFill`/`a:blipFill`/`a:pattFill`, `a:alpha`, `a:effectLst/a:outerShdw`, `a:headEnd`/`a:tailEnd`, `a:ln@cap`/`a:round|bevel|miter`) and through `.odt` (graphic style properties: `draw:fill`, `draw:fill-gradient-name` with a `draw:gradient` definition, `draw:opacity`, `draw:shadow-*`, `draw:marker-*`, `svg:stroke-linejoin`), asserted on the written XML
- [ ] A property one format cannot express drops with a `DropRecord` naming it, rather than being silently flattened, and the conversion report (SPEC-008) lists it
- [ ] Styling the product does not model — a three-colour ODF hatch angle we ignore, an OOXML effect beyond shadow — is **preserved verbatim** on a same-format save, per "Preserving what we do not model", closing the current silent-downgrade hole
- [ ] Charts (SPEC-052) and images (SPEC-024) are unaffected: their own fill and border handling is untouched by this spec
- [ ] Fixtures with a gradient-filled shape, a picture-filled shape, a shadowed callout, a semi-transparent shape over text, and a double-headed arrow open, render, save and reopen `equalStructural`

## Out of scope
Glow, soft edges, reflection, bevel and any 3-D effect — modelled as
preserved-not-rendered, per the preservation rule, and left for a later
spec. Text effects inside a shape (SPEC-013 owns run formatting). Theme
colours as fill values — SPEC-077 introduces the theme, and until then
fills are literal colours. Gradient fills on table cells, paragraphs or
page backgrounds (SPEC-074, SPEC-075). An eyedropper. Editing a picture
fill's crop. Applying a style preset to a whole document ("format all
shapes"). WordArt.

## Reference

### The fill model
```ts
// src/lib/core/shapes.ts — ShapeVisual and the model both use these
export type ShapeFill =
  | { kind: "none" }
  | { kind: "solid"; color: string; alpha?: number }
  | { kind: "gradient"; gradient: ShapeGradient }
  | { kind: "picture"; mediaId: string; mode: "stretch" | "tile"; alpha?: number }
  | { kind: "pattern"; preset: string; foreground: string; background: string };

export interface ShapeGradient {
  type: "linear" | "radial";
  /** Whole degrees clockwise from "left to right"; linear only. */
  angleDeg?: number;
  /** Two or more, ordered by position, first at 0 and last at 1. */
  stops: Array<{ position: number; color: string; alpha?: number }>;
}
```

`alpha` is a 0..1 opacity — 1 (or absent) is fully opaque. One convention
everywhere rather than OOXML's "alpha" and ODF's "opacity percentage"
disagreeing in the model; each writer converts.

`mediaId` references the same media store `Image` blocks use (SPEC-024), so
a picture fill and an image of the same file share one part and one
`MediaRecord`, and neither writer duplicates it.

### The outline model
```ts
export interface ShapeOutline {
  color: string;
  widthPt: number;
  dash: "solid" | "dash" | "dot" | "dashDot" | "longDash" | "longDashDot";
  alpha?: number;
  join?: "round" | "bevel" | "miter";        // absent = miter, both formats' default
  cap?: "flat" | "round" | "square";         // absent = flat
  compound?: "single" | "double" | "thickThin" | "thinThick";
  startArrow?: { kind: ArrowheadKind; size: "small" | "medium" | "large" };
  endArrow?: { kind: ArrowheadKind; size: "small" | "medium" | "large" };
}
```

The three existing `dash` values keep their meaning and their spelling, so
no stored model and no shipped test changes; the new values extend the
union.

```ts
export interface ShapeShadow {
  color: string;
  alpha: number;
  blurPt: number;
  distancePt: number;
  angleDeg: number;
}
// on Shape / Connector / DrawingGroup:
shadow?: ShapeShadow;
```

### Migration
`Shape.fill` is `string | ShapeFill` in the type for exactly one release's
worth of code — no. Simpler and testable: `fill?: ShapeFill`, and the
readers, the commands and the model factory produce the union. A
`"#rrggbb"` string appearing in a stored model (documents saved by earlier
versions live in `localStorage`) is normalised on load by one pure
function:

```ts
/** Accepts the legacy string form and the union; returns the union. */
export function normaliseShapeFill(value: unknown): ShapeFill | undefined;
```

called by the document-store read path, not scattered through the render
code. Its unit test covers a legacy string, the union, `undefined`, and
garbage. `equalStructural` then compares like with like, which is what
keeps SPEC-051's round-trip tests meaningful across the change.

`DEFAULT_NEW_SHAPE_FILL` (`"#4472c4"`) and
`DEFAULT_NEW_SHAPE_OUTLINE` in `src/lib/core/shapes.ts` become the union
forms of the same values; the colours do not change.

### Preserving what we do not model
This is the part that fixes the existing hole, and it follows SPEC-008's
established mechanism rather than inventing one:

```ts
// on Shape, Connector and DrawingGroup
/** Verbatim source styling this model could not express — the
 *  a:spPr/a:effectLst subtree (.docx) or the graphic style properties
 *  (.odt) — re-emitted unchanged on a SAME-format save for exactly the
 *  properties the model did not parse. Never re-emitted cross-format;
 *  each unexpressible property reports a DropRecord instead. */
unsupportedStyleXml?: string;
```

The rule, stated so it cannot be implemented as "preserve everything and
also write our own", which would produce invalid output: the reader records
only the *unparsed* property elements, and the writer merges — our modelled
properties are written from the model, and the preserved ones are appended
in schema order. A property that this spec models is never in the preserved
set, so there is exactly one source for each.

`DropKind` (`src/lib/opaque.ts`) gains `"shapeStyle"` for the cross-format
case.

### Applying to a group
SPEC-057 put "applying fill or outline to a container and fanning it out to
members" out of scope, because there was no styling model worth fanning.
Now there is, and the rule is: a container has no fill or outline of its
own — neither format gives `wpg:wgp` or `draw:g` one — so applying a fill
to a selected container **applies it to every member recursively**, in one
undo entry labelled `"format drawings"`, skipping members that cannot take
it (connectors take outline only; nested containers recurse). The
container's toolbar therefore gains fill and outline controls, which
SPEC-057 explicitly did not give it.

A sealed container (SPEC-057) is excluded: rewriting styling inside
preserved XML is exactly what sealing prevents.

### Style presets
A small gallery of coordinated looks, each a `(fill, outline, shadow)`
triple over the current accent — filled, subtle-filled, outlined, outlined
+ shadow, and a "no fill" — applied in one click and one undo entry. The
presets are data in `src/lib/core/shapes.ts`, not a UI concern, so a future
Slides host gets the same set. They are literal colours today; when
SPEC-077 lands theme colours, presets become the natural first consumer.

### Rendering
The canvas painter gains gradient, picture and pattern fills, alpha
compositing, and an outer-shadow pass drawn before the shape. Two
constraints carry over:

- **Zoom**: canvases size their backing store by `devicePixelRatio * zoom`
  (SPEC-063), and a gradient or shadow must be computed in the scaled space
  rather than scaled as a bitmap, or it blurs.
- **Performance**: shadows and gradients are the expensive part of any
  canvas render. SPEC-047's baseline must not move; a page of styled shapes
  paints through the same cached-render path, and the shadow pass is
  skipped entirely for shapes with no shadow rather than drawn
  transparently.

### Specs this changes
- `specs/051-drawing-shapes-and-text-boxes.md` — its flat fill/outline
  model is superseded by the union above; its geometry, anchoring, wrap,
  selection, text and preservation rules are unchanged. Its shipped
  `dash` values and defaults keep their meanings.
- `specs/057-group-and-ungroup-drawings.md` — its "fill and outline on a
  container is out of scope" note is resolved by the fan-out rule above.
  Sealed containers stay excluded.
- `specs/070-connectors.md` — connectors take the widened outline and the
  arrowhead model defined here; `ArrowheadKind` is declared there and used
  here.
- `specs/008-round-trip-preservation.md` — `DropKind` gains
  `"shapeStyle"`; the preservation mechanism is unchanged, and this spec
  closes a case where the shipped code violated its rule.
