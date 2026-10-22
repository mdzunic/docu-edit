---
id: SPEC-070
title: Connectors
status: draft
area: structure
depends_on: [SPEC-051, SPEC-057, SPEC-069]
---
## Why
SPEC-069 gives the product twenty flowchart shapes. Without connectors it
gives the product twenty boxes that cannot be joined. A connector is not a
line — it is a line that *remembers what it is attached to*, so moving a box
drags its arrows along and re-routes them. That memory is the entire
feature, and it is the reason drawing a plain `line` between two boxes is
not a workaround: the moment either box moves, a hand-drawn line is wrong
and stays wrong.

Both formats model connectors as a first-class object with endpoint
attachments — `wps:cxnSp` with `a:stCxn`/`a:endCxn` naming a shape and a
connection-site index in `.docx`, `draw:connector` with
`draw:start-shape`/`draw:end-shape` and glue-point indices in `.odt` — so
a real document's flowchart arrives with attachments this product cannot
represent. Today they land on SPEC-008 opaque preservation and render as
nothing at all: the boxes appear, the arrows between them do not.

There is a specific wrinkle to clear up while doing this, and it is
already shipped: `src/lib/docx/shape.ts` maps the OOXML preset
`straightConnector1` to our `arrow` geometry in both directions. So today
an arrow *writes* as a connector preset with no attachments, and a real
connector *reads* as a plain arrow, losing its endpoints silently. Whatever
this spec does, it must untangle that without breaking the round trip of
documents already saved by the current writer.

## Acceptance criteria
- [ ] A `Connector` drawing block exists per "The model", with a kind (straight, elbow, curved), two endpoints, and an optional attachment per endpoint
- [ ] The Insert gallery offers the three connector kinds in a Connectors group; choosing one enters a two-click draw mode — click or hover a source, click or hover a target
- [ ] While drawing, hovering a shape reveals its **connection sites** (per "Connection sites"), and releasing on one attaches that endpoint to it; releasing on empty canvas leaves that endpoint free at that point
- [ ] Moving, resizing or rotating a shape re-routes every connector attached to it, live during the drag, with no separate commit — the connector's geometry is derived, not stored as an independent position
- [ ] Deleting a shape leaves its connectors in place with the affected endpoints detached at their last position, rather than deleting them
- [ ] An elbow connector routes with orthogonal segments chosen per "Routing", avoiding a path that runs back through either attached shape; a curved connector uses a cubic curve with the same endpoint directions
- [ ] A selected connector shows endpoint handles that can be dragged to re-attach, detach, or attach to a different site, and a midpoint adjustment handle for elbow and curved kinds that shifts the routing, stored as an adjustment
- [ ] Connectors carry line formatting — colour, width, dash — and arrowheads at either end, per "Arrowheads", defaulting to an arrow at the end
- [ ] A connector is an anchored drawing like any other: it has wrap, z-order, alt text, and it can be a member of a group (SPEC-057), where its attachments resolve within the group
- [ ] Connectors round-trip through `.docx` as `wps:cxnSp` — `a:prstGeom` `straightConnector1` / `bentConnector3` / `curvedConnector3`, with `a:stCxn`/`a:endCxn` `@id`/`@idx` when attached — asserted on the written XML
- [ ] Connectors round-trip through `.odt` as `draw:connector` with `draw:type`, `draw:start-shape`/`draw:end-shape` and `draw:start-glue-point`/`draw:end-glue-point` when attached, and `svg:x1/y1/x2/y2` always
- [ ] Attachment survives the round trip through both formats' identifier schemes without any identifier entering the model, per "Identity without ids"
- [ ] The existing `straightConnector1` ↔ `arrow` mapping is resolved per "Untangling the arrow preset", and a `.docx` written by the current shipped writer still opens with its arrows as arrows
- [ ] Cross-format saves preserve attachment; a connector whose attachment cannot be expressed drops to a free-floating connector with a `DropRecord`
- [ ] Fixtures containing an attached straight, elbow and curved connector, a connector with one free end, and a connector inside a group open, render, re-route on a move, save and reopen `equalStructural`

## Out of scope
Automatic routing that avoids *unrelated* shapes — routing considers only
the two attached shapes, as both formats' own renderers do. Connection
sites beyond the standard set in "Connection sites" (per-geometry custom
sites from `a:cxnLst`). Connector *labels* — text on a connector; neither
this spec nor SPEC-051 gives a line a text frame. Auto-connect on shape
creation ("drop a box on a connector to insert it"). Reflow of a whole
diagram. SmartArt. Word's drawing canvas — still preserved, never created.

## Reference

### What ships today, and the trap in it
`src/lib/docx/shape.ts` contains:

```ts
const PRST_TO_GEOMETRY: Record<string, ShapeGeometry> = {
  line: "line",
  straightConnector1: "arrow",     // ← the collision
  …
};
const GEOMETRY_TO_PRST: Record<ShapeGeometry, string> = {
  line: "line",
  arrow: "straightConnector1",     // ← and back
  …
};
```

`.odt`'s arrow is different again: `src/lib/odf/shape.ts` writes a
`draw:custom-shape`/`draw:line` with `draw:marker-end="Arrow"`, which is a
plain line with an arrowhead — the correct representation. So the model's
`arrow` is really "a line with an end marker", and only the `.docx` writer
misexpresses it as a connector preset.

### Untangling the arrow preset
The model's `arrow` geometry stays exactly what it is — a line with an end
arrowhead, unattached — and stops being written as a connector:

- **Write**: `arrow` emits `<a:prstGeom prst="line"/>` inside `wps:wsp`
  with a `<a:tailEnd type="triangle"/>` on its `a:ln`, matching what
  `.odt` already does and what Word produces for a line with an arrowhead.
- **Read**: `straightConnector1` inside a `wps:wsp` (not a `wps:cxnSp`),
  with no `a:stCxn`/`a:endCxn`, reads back as `arrow` — so every `.docx`
  this app has already written opens unchanged. That entry stays in
  `PRST_TO_GEOMETRY` as a **read-only compatibility** mapping, commented
  as such, and `GEOMETRY_TO_PRST` loses its `arrow` row.
- A `wps:cxnSp` — attached or not — reads as a `Connector`, never as a
  `Shape`.

The existing `.docx` shape round-trip tests that assert
`prst="straightConnector1"` for an arrow are the bounded, sanctioned
replacement this change requires: they become assertions on
`prst="line"` plus `a:tailEnd`, with a new case covering the legacy read
path. No other shape test changes.

### The model
```ts
// src/lib/documentModel.ts
export type Block = … | Connector;          // added to the union
export type DrawingMember = Shape | Image | Chart | DrawingGroup | Connector;

export type ConnectorKind = "straight" | "elbow" | "curved";

export interface Connector extends DrawingPlacement {
  type: "connector";
  kind: ConnectorKind;
  /** Endpoints in cm, relative to the connector's own placement box —
   *  always present, and always the truth when an end is unattached.
   *  When an end IS attached, they are the last routed position, kept so
   *  a detached-by-deletion endpoint has somewhere to be and so a reader
   *  that ignores attachment still draws the right line. */
  start: { xCm: number; yCm: number };
  end: { xCm: number; yCm: number };
  startAttachment?: ConnectorAttachment;
  endAttachment?: ConnectorAttachment;
  /** Routing adjustment for elbow/curved kinds: the fraction along the
   *  connector at which the middle segment sits. Absent means 0.5. */
  adjustments?: number[];
  outline?: ShapeOutline;
  startArrow?: ArrowheadKind;
  endArrow?: ArrowheadKind;                 // defaults to "triangle"
  changes?: BlockTrackedChanges;
}

export interface ConnectorAttachment {
  /** Index path to the attached drawing, in the same form SPEC-057 uses
   *  for group members: block index within the section, then member path.
   *  Not an identifier — see "Identity without ids". */
  target: { block: number; path?: number[] };
  /** Which connection site, per "Connection sites". */
  site: number;
}

export function isConnector(block: Block | undefined): block is Connector {
  return block?.type === "connector";
}
```

The geometry a connector *paints* is always derived — from its attachments
where it has them, from `start`/`end` where it does not — so there is no
second source of truth to keep in sync, which is what makes "moving a box
re-routes its arrows" fall out of the existing render path rather than
needing an update pass.

### Identity without ids
`src/lib/equality.ts`'s `IGNORED_KEYS` drops every key named `id` from
structural comparison, and SPEC-057 chose containment over identifiers for
exactly this reason. A connector cannot hold a shape id, so it holds a
**position**: the block index of its target within the section, plus
SPEC-057's member path when the target is inside a group.

That makes attachment a structural fact that `equalStructural` compares
correctly, and it makes every edit that reorders blocks responsible for
re-indexing attachments — the same responsibility those edits already have
for comment anchors and cross-references. The re-indexing lives in one pure
function:

```ts
// src/lib/connectors.ts
/** Rewrites every connector attachment in a section after blocks move,
 *  and detaches those whose target is gone. Called by any command that
 *  inserts, deletes or reorders blocks. */
export function reindexAttachments(section: Section, mapping: BlockIndexMapping): Section;
```

Both file formats *do* use identifiers (`a:stCxn/@id` referencing a
`wp:docPr/@id`; `draw:start-shape` referencing a `draw:id`). The readers
resolve those to indices on the way in, and the writers mint fresh
document-unique ids on the way out — exactly the discipline SPEC-057
already applies to `cNvPr` ids, which are "never modelled".

### Connection sites
Nine sites per shape, indexed 0–8: the four edge midpoints (top, right,
bottom, left), the four corners, and the centre. They are derived from the
target's bounding box, so every geometry has them and no per-geometry data
is needed.

Index order is fixed and stated because it is written to disk: 0 top,
1 right, 2 bottom, 3 left, 4 top-left, 5 top-right, 6 bottom-right,
7 bottom-left, 8 centre. `.docx`'s `@idx` and `.odt`'s glue-point index
use their own numbering per shape; the readers map into this order and the
writers map out, per geometry, with an unmappable index falling back to the
nearest site by position.

Sites are shown while a connector is being drawn or an endpoint dragged —
small circles on the hovered shape — and never otherwise.

### Routing
- **Straight** — a segment between the two resolved points.
- **Elbow** — three segments: leave each endpoint perpendicular to its
  site's edge for a fixed stub, then join with one crossing segment placed
  at the adjustment fraction. Where the two stubs already align on an
  axis, it degenerates to a single segment.
- **Curved** — a cubic Bézier whose control points extend from each
  endpoint along its site's outward normal, scaled by the distance between
  the points.

An endpoint attached to the **centre** site (8) resolves to the point where
the routed line crosses the target's outline, so a connector to a centre
never draws inside the shape. This uses the existing per-geometry outline
from SPEC-069's `ShapeDefinition`.

All routing is pure and lives in `src/lib/core/connectors.ts` — geometry
in, points out — so it is unit-testable without a document and reusable by
a future Slides host, per the core layer's boundary rule.

### Arrowheads
```ts
export type ArrowheadKind = "none" | "triangle" | "open" | "stealth" | "diamond" | "oval";
```

`.docx` writes `a:headEnd`/`a:tailEnd` with `@type`; `.odt` writes
`draw:marker-start`/`draw:marker-end` referencing the standard marker
names, which is the mechanism the shipped `.odt` writer already uses for
`arrow`. SPEC-073 extends arrowheads to plain lines and shape outlines;
this spec introduces the type and uses it for connectors.

### Specs this changes
- `specs/051-drawing-shapes-and-text-boxes.md` — "Connectors" leaves its
  out-of-scope list. Its `arrow` geometry keeps its meaning; only the
  `.docx` preset it writes changes, per "Untangling the arrow preset".
- `specs/057-group-and-ungroup-drawings.md` — `DrawingMember` gains
  `Connector`; a connector inside a container resolves attachments by
  member path within that container. Grouping a connector with the shapes
  it attaches to keeps the attachment; grouping only one of them detaches
  nothing — the attachment path is rewritten by `reindexAttachments`.
- `specs/008-round-trip-preservation.md` — `wps:cxnSp` and
  `draw:connector` leave opaque preservation for the modelled path. No new
  `DropKind`: an unexpressible attachment reports the existing
  `"shapeGeometry"` kind.
- `specs/069-shape-catalogue-expansion.md` — connectors use its
  `ShapeDefinition.outline` for centre-site clipping; nothing in it
  changes.
