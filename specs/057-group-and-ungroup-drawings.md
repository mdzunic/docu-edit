---
id: SPEC-057
title: Group and ungroup drawings — native containers, nesting and editing in place
status: draft
area: structure
depends_on: [SPEC-008, SPEC-024, SPEC-051, SPEC-052]
---
## Why
SPEC-051 shipped Group and Ungroup, and its tests are green — but the group
itself is a fiction private to this editor. Membership is a flat
`groupOrdinal` tag on each member, round-tripped through a private extension
attribute (`de:group`, namespace `urn:docu-edit:shape`) that no other
application honours. Word and Writer open our "group" as loose shapes, and
the first save from either quietly discards the attribute, dissolving the
group. In the other direction a real group — `wpg:wgp` in `.docx`, `draw:g`
in `.odt`, among the most common drawing constructs in real documents —
never reaches the model: the readers preserve it as opaque XML that nothing
renders, so the document opens with its diagram simply missing.

The flat tag has a second cost: it cannot make "one unit" structurally true.
Every operation on a drawing must remember to fan out over the ordinal's
members, and the shipped code shows what happens when one forgets — dragging
a group leaves image and chart members behind, arrow-nudge moves a single
member out of its group, Delete and Ctrl/Cmd+D act on one member, a group
drag skips re-anchoring, and a grouped image's membership is never written
to either format. Each is the same forgotten fan-out.

Both file formats agree on the fix, because both store a group as a
container element holding its members. This spec makes the model match: a
`DrawingGroup` block that contains its members. Moving, nudging, anchoring,
wrapping, z-ordering, deleting and duplicating a group then *are* the
existing single-drawing code paths, with nothing to remember — and the
written file carries the container Word and Writer expect. On top of the
container it adds what people expect of groups: nesting, entering a group to
edit one member, and keyboard access.

## Acceptance criteria
- [ ] The flat `groupOrdinal` field and the `de:group` attribute are retired in favour of the `DrawingGroup` container in "The group container" below; no writer emits `de:group` again, and both readers still accept it per "Legacy migration"
- [ ] With two or more drawings selected — any mix of shapes, images and charts, per "Selection across drawing kinds" — Group (the existing toolbar button, or Ctrl/Cmd+G) replaces them with one container at the first member's block position, and no member moves on screen
- [ ] Ungroup (the existing toolbar button, or Ctrl/Cmd+Shift+G) dissolves the selected container one level; members re-emerge as sibling drawings with unchanged on-screen geometry, per "Ungrouping"
- [ ] Grouping a selection that includes a group nests that group as one member; Ungroup peels one level at a time; nesting reads, renders and writes at arbitrary depth
- [ ] A selected group moves, arrow-nudges, deletes, duplicates, re-anchors (anchor marker following, per SPEC-051's re-anchoring rules), wraps and z-orders as one drawing — each through the same path a single shape uses, and each asserted by a test that also covers a group holding an image or chart member
- [ ] Dragging a group's resize handles scales every member's size and offset per axis, per "Interior scaling"; dragging its rotation handle sets the container's `rotationDeg`; both render live during the drag — a mid-drag frame already shows the members scaled, or the container turned
- [ ] A click selects the outermost container under the pointer by SPEC-051's painted-region rule applied to members; a second click on an already-selected group selects the member under the pointer (one level per click); Escape steps back out one level; double-clicking an already-selected text-bearing member enters its text (SPEC-051's double-click), so text inside a grouped text box is editable in place with every feature SPEC-051 gives shape text
- [ ] A selected member can be dragged, arrow-nudged, resized and rotated inside its group; the container renormalizes per "Normalized form" with no visual jump; Delete removes the member, and a container reduced to one member dissolves into it
- [ ] Bring to front, bring forward, send backward and send to back on a selected member reorder it among its siblings — inside a container, paint order is member order
- [ ] Every top-level drawing — now including images — carries the `drawing-<index>` proxy; a container's proxy has `role="group"` and nests its members' proxies as `drawing-<i>-member-<j>` (recursively); Enter descends into a focused group, Escape ascends, Tab traverses siblings, per "Rendering and test surface"
- [ ] Group, Ungroup and every group or member gesture record exactly one labelled undo entry each — a resize or rotate drag commits once, on release — and undo restores both model and selection
- [ ] A container carries `altText`, editable from the toolbar while it is selected, written as `wp:docPr/@descr` and as `svg:title`/`svg:desc` on `draw:g`, and SPEC-043 treats a container as one drawing
- [ ] Opening a `.docx` with `wpg:wgp` — including `pic:pic` members, `wpg:graphicFrame` chart members, nested `wpg:grpSp`, member and container rotation, and a non-identity child coordinate space — yields containers per "Format mapping — .docx", and saving writes one `wp:anchor`/`wp:inline` holding `wpg:wgp` with the members as children, asserted on the written XML
- [ ] Opening an `.odt` with `draw:g` — including nested groups and `draw:frame` members — yields the same containers per "Format mapping — .odt", and saving writes `draw:g`; a rotated container bakes its rotation on `.odt` save via the exported transform in "Rotation and the bake transform"
- [ ] The shipped SPEC-051 fixtures, which carry `de:group`, open with their grouped pair as one container, and saving then reopening them keeps the container — the migration test in "Legacy migration"
- [ ] A native group containing a member outside SPEC-051's shape catalogue models that member as SPEC-051's preserved placeholder and is **sealed**: it still moves, re-anchors and ungroups, but does not enter, offers no resize or rotation handles, and shows read-only alt text while it contains one at any depth, per "Preserved members and preserved containers"
- [ ] A cross-format save that must drop every member of a container drops the container itself with a `DropRecord` of the new kind `"drawingGroup"`, and cross-format round trips hold under the mask and transform named in "Equality"
- [ ] Opening the fixtures below renders the plain pair, the nested group, the mixed shape-image-text-box group, the rotated group and the chart-member group, each selectable and enterable, and the preserved-member group per the AC above

## Out of scope
Aligning and distributing members or a multi-selection — the natural next
spec, which this one unblocks. A marquee / rubber-band selection gesture —
Shift-click (SPEC-051) remains the only multi-select. Word's "Regroup"
(re-forming the last ungrouped set from memory). Applying fill or outline to
a container and fanning it out to members. Flips: a container or member
carrying `flipH`/`flipV` is preserved per SPEC-008 rather than modelled.
Word's drawing canvas `wpc:wpc` — unchanged from SPEC-051: preserved, never
created, never ungrouped. Editing the content of a preserved member.
Connectors, SmartArt and tight wrap — unchanged from SPEC-051. Group-level
text (neither format gives a container its own text body). Snapping and
alignment guides while dragging members.

## Reference

### Document model — the implemented one
`specs/002-create-blank-document.md` defines the model normatively and
describes the shipped `src/lib/documentModel.ts` directly. **Read it before
starting.** The facts this spec rests on:

- The shipped block union is
  `Block = Paragraph | Table | Image | BreakBlock | TableOfContents | Shape | Chart`
  (`src/lib/documentModel.ts`), discriminated by a `type` string literal.
  Every drawing is a block in the block list, in document order, at the
  position it is anchored to.
- `DrawingPlacement` (SPEC-051, restated in SPEC-052) is the shared base
  `Image`, `Shape` and `Chart` extend: `widthCm`, `heightCm`, `wrap`,
  `altText?`, `offsetXCm?`, `offsetYCm?`, `relativeToH?`, `relativeToV?`,
  `zOrder?`, `rotationDeg?`, and — until this spec — `groupOrdinal?`.
- Lengths are **centimetres**, type sizes **points**, rotation **whole
  degrees clockwise**. The `.docx` writer converts to EMU (360000 per cm —
  computed in `src/lib/core/units.ts`, not a named constant) and to
  60000ths of a degree on the way out; `.odt` stores centimetres and
  degrees directly.
- `src/lib/equality.ts`'s `IGNORED_KEYS` drops every key named `id` from
  structural comparison, so nothing in this spec's model may be called `id`.
  The container below needs no identifier at all — membership is
  containment.

### What SPEC-051 actually shipped for grouping
This is the mechanism being replaced; a reader of this file alone must know
it precisely.

- Membership is `groupOrdinal?: number` on `DrawingPlacement`. All grouping
  operations live in `src/lib/shapes.ts` (`groupShapes`, `ungroupShapes`,
  `groupMembers`, `moveGroup`, `resizeGroup`, `rotateGroup`,
  `groupBoundingBoxCm`) and fan out over blocks sharing the ordinal; the
  group's box is derived, never stored.
- Both writers emit **one drawing per member** — its own `wp:anchor` /
  `wp:inline` paragraph in `.docx` (`src/lib/docx/shape.ts`), its own
  `draw:custom-shape`/`draw:line` in `.odt` (`src/lib/odf/shape.ts`).
  Membership rides as `de:group="<ordinal>"` in the private namespace
  `urn:docu-edit:shape` — written only by the two shape writers and the
  `.docx` chart writer (`src/lib/docx/chart.ts`); the `.odt` chart writer
  and both image writers never tag, so those memberships silently drop on
  save. (`de:rotation` in the same namespace is SPEC-051's
  ODT rotation mechanism; it is **not** touched by this spec.) The readers
  read the attribute back. No `wpg:wgp` or `draw:g` is ever written, and on
  read both fall through to SPEC-008 opaque preservation — verbatim,
  unrendered, unselectable.
- Multi-select is Shift-click and **shape-only**: `useShapes.ts` keeps
  `shapeMultiSelection` (a `ShapeLocation[]`), and its own doc comment flags
  that images are excluded from the gesture. Group/Ungroup appear only as
  the floating shape-toolbar buttons `shape-group-button` /
  `shape-ungroup-button`; there is no menu item, no context-menu entry, no
  keyboard shortcut, and no entry in the `editorCommands` registry
  (`src/views/DocumentView.vue`).
- Clicking a grouped member selects **only that member**; group move,
  resize and rotate run through the member's handles and fan out over the
  ordinal's blocks — move and resize once on release, rotate on every
  mousemove. Undo goes through `applyEdit` → `recordHistoryEdit` with full
  before/after model snapshots and labels (`"group shapes"`,
  `"move group"`, …).
- Images carry no `drawing-<index>` proxy at all — the counter covers
  shapes and charts only (`usePaginationView.ts`, `pageRender.ts`).

Everything in that list is superseded by the sections below except
`de:rotation` and the toolbar buttons, which stay.

### The group container — normative

```ts
// src/lib/documentModel.ts
export type Block =
  | Paragraph | Table | Image | BreakBlock | TableOfContents | Shape | Chart
  | DrawingGroup                             // added by this spec

/** What a group may hold. A nested DrawingGroup is a member like any
 *  other. */
export type DrawingMember = Shape | Image | Chart | DrawingGroup

export interface DrawingGroup extends DrawingPlacement {
  type: "drawingGroup"
  /** Paint order: earlier = further from the reader. Never empty. */
  members: DrawingMember[]
}

export function isDrawingGroup(block: Block | undefined): block is DrawingGroup {
  return block?.type === "drawingGroup";
}
```

(`isDrawingGroup` mirrors the shipped `isShape`/`isToc` helpers, which exist
because espree cannot parse a type predicate in a `.vue` script.)

`DrawingPlacement.groupOrdinal` is **deleted**, along with every code path
that reads or writes it and the `de:group` attribute — see "Legacy
migration" for the read-side compatibility that replaces it, and "Specs this
changes" for the tests this retires.

**Member field semantics inside a container.** A member's `offsetXCm` /
`offsetYCm` are measured from the **container's top-left corner**, in cm,
and are **always present** on a member — normalization writes `0`
explicitly rather than omitting it, so structural equality never hinges on
absent-versus-zero. A member's `relativeToH`, `relativeToV` and `zOrder`
are meaningless inside a container and are **absent** — readers normalize
them away, the writer ignores them, and ungrouping re-derives them from the
container. `wrap` is **required** by the shipped `DrawingPlacement`, so
members carry the fixed placeholder `wrap: "square"` — never read, always
overwritten by Ungroup — keeping a shipped required field required instead
of loosening the whole model's type for an interior state.
`widthCm`/`heightCm`/`rotationDeg` keep their SPEC-051 meanings (a member
rotates about its own centre). The container itself is an ordinary anchored
drawing: its own `wrap`, offsets, anchor references, `zOrder`, `rotationDeg`
and `altText` behave exactly as SPEC-051 defines them for a shape, through
the same code paths.

**Addressing a member.** Interior operations address a member by index path
from the container block: `{ section, block, path: number[] }`, where each
path element indexes `members`. The shape text-frame seam
(`withShapeModel`/`commitShapeModel` in `src/lib/commands.shapes.ts`, which
today address `{ section, block }`) gains the optional `path` so text
editing inside a grouped text box reuses the shipped mechanism unchanged.

**No identifiers.** Membership is containment, so the `IGNORED_KEYS` trap
SPEC-051 named `groupOrdinal` around disappears rather than needing a
convention.

**Never fewer than two members.** Group requires two drawings, an interior
delete dissolves a container at one member, readers dissolve a native
single-child group into that child, and a childless `wpg:wgp`/`draw:g` is
preserved opaquely per SPEC-008 — so `members.length >= 2` holds
everywhere; the code block's "Never empty" is the weaker consequence.

### Normalized form
Whenever a container's `rotationDeg` is 0 (or absent), these invariants
hold:

- `min(member.offsetXCm) === 0` and `min(member.offsetYCm) === 0`;
- the container's `widthCm`/`heightCm` equal the union of the members'
  **unrotated** boxes (`offset + size`; member rotation is ignored for
  bounds, matching how both formats' group extents treat rotated children).

After any interior edit (member move, nudge, resize, rotate, duplicate,
delete), **renormalize**: recompute the union, fold its origin into the container's
own offsets, and re-express member offsets — so nothing moves on screen and
the invariants are restored. A container whose `rotationDeg` is non-zero
keeps its stored box while members are edited inside it: the box is the
rotation frame, and re-deriving it mid-rotation would move the pivot.
Ungrouping and `.odt` saving bake rotation first (below), which restores the
invariants. Readers normalize unrotated containers on open; a rotated
container keeps the file's stored box verbatim.

### Creating a group
Group requires every selected drawing to sit in the **same section** — the
container needs one block position — and the Group control disables
otherwise. It replaces the selected drawings with one container **at the
first member's block position** (document order). Grouping never moves a
*member*; surrounding text may reflow, because replacing individually
wrapped drawings with one `square` container changes the text-exclusion
region. Precisely:

- Each member's current **laid-out page rectangle** (the same geometry the
  painter just used; the gesture layer passes these rectangles into the
  pure model operation, keeping `src/lib` layout-free; a rotated member
  contributes its unrotated frame rectangle, per "Normalized form") is
  taken as its truth — this, not the raw offsets, is what survives, so
  members anchored to different paragraphs and `inline` members group
  correctly. An `inline`
  member leaves the text flow by joining a group, exactly as SPEC-051's
  grouping already converts `inline` to `square`; body text reflows to
  close the gap, the member itself does not move.
- The container anchors to the first member's paragraph with
  `relativeToH: "column"`, `relativeToV: "paragraph"` (SPEC-051's canonical
  default), `wrap` taken from the first member (`"square"` if the first
  member was `inline`), and offsets chosen so the container's box lands on
  the union of the member rectangles.
- Member offsets become their rectangle's position within that box. Member
  order in `members` is the members' current paint order — ascending
  `zOrder`, absent counting as 0, ties broken by block order — and the
  container's `zOrder` is the highest member value, so the group stacks
  where its topmost member painted.

Members' sizes, rotations, fills and text are untouched, and Group leaves
the **new container selected**. Grouping a selection that includes a
container simply makes that container a member — that is all "nested
groups" are.

### Ungrouping
Ungroup dissolves the selected container one level, splicing its members
into the block list at the container's position, in member order. If the
container's `rotationDeg` is non-zero, the bake transform below runs first;
an `inline` container first converts to `square` at its laid-out position —
the same conversion Group applies to an inline member — because members
re-entering the text flow could not hold their offsets.
Each member then inherits the container's `wrap`, `relativeToH`,
`relativeToV` and `zOrder` verbatim, anchors to the container's paragraph,
and gets offsets `container offset + member offset` — so nothing moves on
screen, and the members' mutual paint order is preserved by block order
(SPEC-051's painter breaks `zOrder` ties by block order). A nested member
container comes out as a container — Ungroup peels one level. Ungroup
leaves the emerged members **multi-selected**, so Ctrl/Cmd+G and
Ctrl/Cmd+Shift+G invert each other immediately.

The same splice, for a single member, is how a container reduced to one
member dissolves.

### Interior scaling and container rotation
Resizing a selected container works like resizing a single shape — the
handles move the container's own box, with the `MIN_SHAPE_SIZE_CM` floor
(0.2cm, `src/lib/core/shapes.ts`) applied to **both** container axes. The
interior scales with it: with
`s_x = newWidthCm / oldWidthCm` and `s_y = newHeightCm / oldHeightCm`, every
member (recursively) gets `offsetXCm *= s_x`, `offsetYCm *= s_y`,
`widthCm *= s_x`, `heightCm *= s_y`. Members need no floor of their own —
the factors are never zero. Member `rotationDeg` is unchanged; scaling
applies to a member's unrotated box, which is exactly what both formats do
to a rotated child's frame under a group scale (aspect changes, no shear).

Rotating a selected container sets the **container's** `rotationDeg` — it
does not touch members. This supersedes SPEC-051's flat-group rule of adding
the delta to every member, which existed only because there was no container
to store the angle on. Painting, hit testing and the handles transform by
the container's rotation about its box centre, per SPEC-051's existing
rotation rules for a single shape.

### Rotation and the bake transform
`.docx` stores a group's rotation natively (`@rot` on the group's
`a:xfrm`), so containers round-trip it. ODF's `draw:g` has **no** rotation
(or position/size) attributes of its own — LibreOffice itself rotates a
group by rewriting its children. We do the same, through one exported,
unit-tested transform used by the `.odt` writer, by Ungroup, and by the
cross-format equality mask:

`bakeGroupRotation(group)` — for a container with rotation `θ` and box
centre `C` (in its own frame): each member's centre moves to
`C + R(θ)·(centre − C)` (R clockwise, matching `rotationDeg`), each member's
`rotationDeg` becomes `(rotationDeg + θ) mod 360`, the container's
`rotationDeg` clears, and the container renormalizes. Nested member
containers are baked by rotating their centres and adding `θ` to their own
`rotationDeg` — not by recursing into their members. The `.odt` writer
applies the bake to a serialization-time copy — **saving never mutates the
open model** — level by level, outermost first, until no container in the
subtree carries rotation; the cross-format equality mask applies the
identical procedure.

Cross-format consequence, stated rather than hidden: a rotated container
saved to `.odt` and reopened has its rotation distributed onto members —
visually identical, structurally different. The Equality section masks this
with the same transform, never with a hand-rolled comparison.

### Selection across drawing kinds
- SPEC-051's Shift-click multi-select widens from shapes to **every**
  drawing kind: shapes, images, charts and containers join one
  multi-selection (retiring the shape-only `shapeMultiSelection`
  limitation its own doc comment flags). Group is offered when the
  multi-selection holds two or more drawings in one section. Shift-click
  always resolves to **top-level** drawings — members never join a
  multi-selection, so Group never reparents across containers — and a
  preserved drawing (SPEC-051's placeholder) never joins Group, so sealed
  containers (below) only arise from native files.
- A plain click selects the **outermost container**: SPEC-051's
  painted-region hit rule runs against members (a container has no paint of
  its own), and the hit resolves to the top-level ancestor. An unfilled
  member still needs its 3pt outline tolerance; an empty region inside the
  container's box selects nothing — the box is not a hit surface.
- A second click on an already-selected container selects the **member
  under the pointer**, one level per click (a nested container is selected
  before its members). While a member of container C is selected, a plain
  click resolves within C: a sibling under the pointer is selected
  directly, a click on the already-selected member keeps it, and a click
  outside C falls back to the outermost rule. Escape reverses one level;
  Escape with a top-level drawing selected is unchanged by this spec;
  clicking outside every drawing clears as today. Double-click composes
  from single clicks with no special case:
  on an unselected group it selects the member; a further double-click on a
  text-bearing member enters its text, which is SPEC-051's existing
  double-click.
- A selected member shows its own handles and shape toolbar — fill,
  outline, rotation, alt text apply to it; its wrap and z-order controls
  are replaced by the four reorder verbs acting within `members`. Dragging,
  nudging, resizing and rotating it edit it inside the container, then
  renormalize. Delete removes it (the one-member rule above applies);
  Ctrl/Cmd+D duplicates it as a sibling inserted immediately after its
  source in `members`, offset by 0.5cm.
- A selected container shows the container toolbar: wrap, z-order,
  rotation, alt text, Duplicate, Delete and Ungroup (a sealed container
  hides rotation and shows alt text read-only, per "Preserved members and
  preserved containers"). Fill and outline are absent — fanning them to
  members is out of scope.
- Group/Ungroup register in the `editorCommands` registry as
  `drawing.group` / `drawing.ungroup`, bound to **Ctrl/Cmd+G** and
  **Ctrl/Cmd+Shift+G** (`preventDefault` while the editor has focus), and
  appear in the Format menu; the shape-toolbar buttons stay.

### Format mapping — .docx
Read: inside `wp:anchor`/`wp:inline`, `a:graphic/a:graphicData` with
`uri=".../wordprocessingGroup"` holds `wpg:wgp`. Its `wpg:grpSpPr/a:xfrm`
carries `@rot` (container rotation), `a:off`/`a:ext` (box) and
`a:chOff`/`a:chExt` (the child coordinate space). With per-axis scale
`s = ext / chExt`, a child whose xfrm is `off/ext` in child space has frame
rectangle `(off − chOff) · s` at size `ext · s`; convert to cm, then
normalize — unrotated containers only, per "Normalized form"; a `@rot`
container keeps its stored box (any residual union origin folds into the
container's offsets when normalizing).
Members parse with the shipped element readers: `wps:wsp` → `Shape`
(SPEC-051), `pic:pic` → `Image` (SPEC-024's part, resolved through the
existing media pipeline), `wpg:graphicFrame` holding a chart → `Chart`
(SPEC-052), `wpg:grpSp` → recurse (its own xfrm defines the nested child
space). A member's `altText` lives in its element's own non-visual
properties — `cNvPr/@descr` in whichever prefix its family uses — read and
written there. Anchor position, wrap, z-order and `wp:docPr/@descr` read
from the anchor exactly as for a single shape.

Write: one `wp:anchor` (or `wp:inline` for `wrap: "inline"`) per container,
positioned exactly as SPEC-051 writes a shape's anchor, holding `wpg:wgp`
with an **identity child space** — `chOff` 0,0 and `chExt = ext` = the box
in EMU — members as children in member order at their offsets, `@rot` when
the container is rotated. Every `cNvPr` in the tree gets a fresh
document-unique id at write time; ids are never modelled. The identity
space applies to **unsealed** containers only — a sealed container re-emits
its preserved subtree instead ("Preserved members and preserved
containers").

### Format mapping — .odt
`draw:g` holds its members as child elements and — a schema fact this
design leans on — defines **no geometry of its own**: no size, no
transform, no usable position (readers ignore any stray position attribute
on the `draw:g` itself). Children are positioned in the anchor's
coordinate space directly.
Anchor type, style (wrap), `draw:z-index` and `svg:title`/`svg:desc`
(altText) sit on the `draw:g` and read/write by the same rules SPEC-051
applies to a shape element.

Read: parse children with the shipped element readers (`draw:custom-shape`
/ `draw:line` → `Shape`, `draw:frame` → `Image` or `Chart` per its content,
nested `draw:g` → recurse). The container's box is the union of the
children's unrotated rectangles; its offsets are the union's origin; member
offsets are child position minus that origin. Child `draw:z-index` values
are ignored — paint order inside a container is child document order — and
are not written back. Members keep their own `svg:title`/`svg:desc`
children for member alt text, exactly as standalone elements do.

Write: bake rotation first (above), then emit children at
`container offset + member offset` in the anchor frame, in member order,
with no per-child z-index. Sealed containers re-emit their preserved
subtree instead, position-patched as described below.

### Legacy migration
Both readers keep accepting `de:group`. After a section's blocks are
assembled, drawings sharing an ordinal collapse into one container at the
first member's block index (an ordinal whose members span sections
coalesces per section): the container adopts the first member's `wrap`,
anchor references and `zOrder`; the box is the union of the members' raw
`offsetXCm/offsetYCm` rectangles **treated as same-frame values**; member
offsets re-express against it, their anchor fields clear, and their `wrap`
becomes the member placeholder (`"square"`). An ordinal with a single
member yields no container.

Same-frame is an approximation, and it is the honest one: the flat tag
recorded no layout, members anchored to different paragraphs cannot be
placed relative to each other without laying the document out, and the only
files that carry `de:group` are our own. The migration test: opening a
shipped SPEC-051 fixture yields the pair as one container, and
`equalStructural` holds between that model and the model reopened from its
own save — the save that rewrites the group as a native container.

### Preserved members and preserved containers
A native group member we cannot model as a catalogue shape, image or chart
— including a catalogue geometry carrying an attribute we do not model,
such as a flip — makes its whole container **sealed**; transitively: a
preserved member at any depth seals every container above it. A sealed container is still
modelled: its modellable members parse, paint and hit-test normally, the
unmodellable one paints as SPEC-051's dashed placeholder at its bounds, and
the container additionally carries its source subtree verbatim, the way a
preserved shape carries `unsupportedGeometryXml` (a field named along those
lines — never `id`). What a sealed container offers is exactly what can be
honoured without rewriting preserved geometry:

- it **moves, re-anchors, wraps, z-orders, duplicates and deletes** as one
  drawing — in `.docx` these touch only the anchor around the untouched
  `wpg:wgp` subtree; in `.odt` the writer patches only its direct
  children's position attributes (by the move delta) and the group
  element's own anchor-level attributes (anchor type, style reference,
  `draw:z-index`) — the attributes the same edits touch on an unsealed
  drawing, extending the position-only licence SPEC-051's "movable but not
  re-geometried" already uses;
- it **ungroups**: modelled members exit through their normal standalone
  paths, preserved ones through SPEC-051's preserved-shape path at their
  placeholder bounds;
- it does **not enter** (a second click keeps the container selected),
  shows **no resize or rotation handles**, and its alt text is
  **read-only** — resize and rotation would rewrite coordinates inside
  preserved XML, and the `.odt` home for alt text (`svg:title`/`svg:desc`)
  sits inside the preserved subtree.

On a same-format save, a sealed container's subtree is written back
verbatim apart from those position patches, keeping its original child
coordinate space — reopening re-derives the same model from the same
bytes, which is what keeps the round trip `equalStructural`. Cross-format,
the subtree cannot travel: a sealed container writes as an ordinary
container of its modelled members, and each preserved member drops with
its existing kind — the same trade SPEC-051 makes for a standalone
unsupported geometry.

A container-level construct we do not model — `wpc:wpc`, a `draw:g` or
`wpg:wgp` whose group carries flips, a childless group element — stays on
SPEC-008 opaque preservation exactly as today. This spec widens `DropKind`
(`src/lib/opaque.ts`) once,
with `"drawingGroup"`, for a cross-format save that dropped every member of
a container and so drops the container; per-member drops keep their
existing kinds (`"shapeGeometry"`, `"chart"`, …). SPEC-050's Markdown
writer reports a container through the kinds it already reports drawings
with; no Markdown-side change.

### Rendering and test surface
SPEC-051's split stands: canvas paints, a DOM proxy per drawing carries
focus, ARIA and testids. This spec extends it:

- The painter draws a container by entering its frame (translate; rotate by
  `rotationDeg` about the box centre) and painting members recursively at
  their offsets — images and charts included, so a rotated container
  rotates them too.
- The `drawing-<index>` counter now covers **every** top-level drawing —
  shapes, charts, containers and (new) images, in paint order. A
  container's proxy is `role="group"` with `aria-label` from `altText`,
  falling back to "Group of N drawings", `tabindex="0"`. An image's proxy —
  new in this spec — is `role="img"` with `aria-label` from `altText`,
  falling back to "Image".
- Member proxies nest inside their container's proxy as
  `drawing-<i>-member-<j>` (nested: `…-member-<j>-member-<k>`), with
  SPEC-051's per-kind roles and `tabindex="-1"` until their container is
  entered. Enter on a focused container proxy descends (focus its first
  member), Escape ascends, Tab traverses siblings — the keyboard mirror of
  click-in/click-out.
- Selection outline, handles and the anchor marker (SPEC-051's
  `shape-anchor-marker`, showing which paragraph a floating drawing is
  anchored to) attach to the container
  exactly as to a single shape; an entered member shows its own outline and
  handles inside.

### Undo
Every gesture in this spec commits through the shipped `applyEdit` →
`recordHistoryEdit` snapshot mechanism as **one** labelled entry: `"group
drawings"`, `"ungroup drawings"`, `"move group"`, `"resize group"`,
`"rotate group"`, `"edit group member"` (member move, nudge, resize and
rotate), `"delete group member"`, `"duplicate group member"`, `"reorder
group member"`. Container toolbar edits (wrap, z-order, alt text) reuse
their existing SPEC-051 labels. The shipped entry already snapshots
`{ model, selection }` on both sides, which is what lets undo restore the
pre-gesture selection. Resize and rotate drags — container or member —
preview live and commit once on release; SPEC-051's per-mousemove commits
for group rotation do not survive into the container implementation.

### Specs this changes
- `specs/051-drawing-shapes-and-text-boxes.md` — its "Grouping — move,
  resize, rotate" section, the `groupOrdinal` field and its ordinal-not-id
  rationale are superseded by the container. Its out-of-scope "Nested
  groups" is resolved here; `wpc:wpc` stays out of scope. User-visible
  promises (a group moves/resizes/rotates as one unit and ungroups without
  a visual jump) are unchanged and re-asserted at container level. Two
  shipped behaviours change deliberately: clicking a grouped member now
  selects the container (member selection is the second click), and
  rotating a group stores the angle on the container instead of adding it
  to each member.
- **Superseded tests** — retiring the mechanism retires its tests; this is
  the sanctioned replacement CLAUDE.md's "never weaken a test" rule
  anticipates, and it is bounded to exactly: the grouping blocks of
  `src/lib/shapes.spec.ts` (`groupShapes` … `rotateGroup`, including the
  zero-height-floor case, which the container's symmetric
  `MIN_SHAPE_SIZE_CM` floor replaces), the `de:group` round-trip cases in
  `src/lib/docx/shape.spec.ts` and `src/lib/odf/shape.spec.ts` (replaced by
  native-container round-trip cases plus the legacy-migration cases), and
  the group section of `e2e/SPEC-051.spec.ts`, rewritten to container
  behaviour. The SPEC-051 fixtures are **not** regenerated — they become
  the legacy fixtures — so their grouped pair collapses to one drawing and
  the fixture index map at the top of `e2e/SPEC-051.spec.ts` shifts down by
  one after index 14. Likewise, when images join the `drawing-<index>`
  counter, any e2e spec whose fixture places an image before an indexed
  drawing renumbers. Updating these indices is mechanical, not a weakening.
- `specs/024-images-insert-resize-and-wrap.md` — images join the
  multi-selection gesture, gain the `drawing-<index>` proxy, and can be
  container members (their membership now actually persists — the shipped
  image writers never wrote `de:group`). Image behaviour outside groups is
  untouched.
- `specs/052-charts.md` — charts join the multi-selection gesture and are
  container members in both formats (`wpg:graphicFrame` / `draw:frame`),
  closing the shipped gap where a grouped chart lost membership on `.odt`.
- `specs/021-keyboard-shortcuts-and-reference.md` — the reference gains
  Ctrl/Cmd+G "Group drawings" and Ctrl/Cmd+Shift+G "Ungroup".
- `specs/049-menu-bar-and-toolbar-consolidation.md` — the Format menu gains
  Group and Ungroup, enabled per the selection rules above.
- `specs/043-accessibility-checker.md` — one small change: the shipped
  checker enumerates block types (`isShape`, `isChart` arms in
  `src/lib/accessibilityCheck.ts`), so it gains an `isDrawingGroup` arm. A
  container is one drawing carrying `altText` in the same field, flagged
  when it is empty; members are never separately flagged — recorded so
  nobody adds a member-level detection.
- `specs/008-round-trip-preservation.md` — no mechanism change; `DropKind`
  gains `"drawingGroup"` as above.

### Equality
`src/lib/equality.ts` exports `equalStructural`, `equalSemantic`, `resolve`
and `firstStructuralDifference`; `equalBytes` lives in
`src/lib/core/equality.ts` and `canonicalizeXml` in `src/lib/core/xml.ts`.
(SPEC-051's Equality section predates this split and `equalSemantic`'s
implementation; trust this inventory.) This spec needs only
`equalStructural`.
- **`equalStructural`** — for the same-format round trip of the fixtures
  (native read → write → reopen; readers normalize, the writer emits
  normalized form, so the reopened model equals the opened one), for the
  legacy migration assertion above, and for asserting that grouping,
  entering and ungrouping alter nothing else in the document.
- Cross-format: assert `equalStructural` between the source model — masked
  by the writer's `dropped` records and, for `.docx`→`.odt`, transformed by
  `bakeGroupRotation` — and the reopened model. Both adjustments are named
  transforms over the model, never a hand-rolled subset comparison.

`equalBytes` applies only to preserved parts through SPEC-008's existing
rule; this spec adds no new use of it.

### Fixtures
Fixtures live where SPEC-051's do: built by a script
(`scripts/build-spec057-fixtures.mjs`, sibling of
`build-spec051-fixtures.mjs`), emitted to `e2e/fixtures/` and mirrored
into `src/lib/docx/fixtures/` / `src/lib/odf/fixtures/` for unit tests; the
drawing index map is documented at the top of `e2e/SPEC-057.spec.ts` the way
`e2e/SPEC-051.spec.ts` documents SPEC-051's.

Adds `word-groups.docx` (native `wpg:wgp`, non-identity child space on at
least one group) and `lo-groups.odt` (native `draw:g`), each containing: a
plain two-shape group; a nested group (that pair inside a container with a
third shape); a mixed group of a shape, an image and a text box holding two
paragraphs; a rotated group (container `@rot` in `.docx`; in `.odt` the
baked form this spec's own writer produces, exercising that reading one
needs no container rotation); a group with one
member outside the shape catalogue; and a group containing a chart. The
shipped `word-shapes.docx` / `lo-shapes.odt` are reused unchanged as the
legacy `de:group` fixtures.
