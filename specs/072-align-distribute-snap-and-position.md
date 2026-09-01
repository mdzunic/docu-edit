---
id: SPEC-072
title: Align, distribute, snap and the position dialog
status: draft
area: structure
depends_on: [SPEC-051, SPEC-057, SPEC-064]
---
## Why
SPEC-057 names this spec in its own out-of-scope list — "aligning and
distributing members or a multi-selection — the natural next spec, which
this one unblocks" — and it is the right next one, because SPEC-057 built
everything it needs: a real multi-selection across every drawing kind, a
container whose members have a shared coordinate space, and pure model
operations the gesture layer feeds laid-out rectangles into.

Without it, placing drawings is a pixel-hunting exercise. Three boxes in a
row can only be lined up by dragging each until it looks right, at whatever
zoom happens to be active, with no guide, no snap and no numeric entry.
Every drawing tool solves this with the same three mechanisms, and they
reinforce each other: **align/distribute** for a selection you already
have, **snapping and guides** while you drag, and a **position dialog** for
the cases where the number is what you actually know ("2cm from the left
margin").

The third one carries weight beyond convenience. SPEC-068 disables
drag-based gestures on small screens and promises "a numeric-entry path"
for each; SPEC-064 makes measurement units app-wide and needs a dialog to
show them in. A position dialog is where drawings become accessible to
keyboard and touch users rather than pointer-only.

Narrowed by adoption (SPEC-096): the pure align/distribute/snap rect
algebra now has one product-neutral home in `src/lib/core/arrange.ts`
(`Rect`, `alignRects`, `distributeRects`, `guideCandidates`, `snapRect`)
— this spec consumes it for the text side rather than duplicating it.
Snap-to-grid, its preference and user-placed ruler guides stay this
spec's own.

## Acceptance criteria
- [ ] With two or more drawings selected, six align commands — left, centre, right, top, middle, bottom — align them to the selection's bounding box, per "What alignment aligns to"
- [ ] With a single drawing selected, the same six commands align it to the page's **content area** (margins), and a modifier-free menu affordance switches the reference between margin, page and column, per that section
- [ ] With three or more selected, distribute horizontally and distribute vertically equalise the gaps between them, leaving the outermost two in place
- [ ] Align and distribute work identically on group members (SPEC-057's interior selection) and on top-level drawings, using the container's coordinate space in the former case, and renormalise the container afterwards with no visual jump
- [ ] Each command records exactly one labelled undo entry, and undo restores both model and selection
- [ ] The commands appear in `Format → Drawing → Align` and `→ Distribute`, on the drawing context menu, and on the contextual drawing toolbar; each registers in the `editorCommands` Map and is disabled with a reason when the selection does not support it
- [ ] While dragging or resizing a drawing, alignment guides appear when an edge or centre lines up with another drawing's edge or centre, with the page margins, or with the page centre — drawn as thin lines through the aligned features, per "Guides and snapping"
- [ ] Snapping pulls the dragged drawing onto a candidate alignment within a stated threshold measured in **screen** pixels (so it feels the same at every zoom) and is suppressed while a modifier key is held, so precise placement is always possible
- [ ] Snapping applies to move, resize and to the creation drag of a new drawing, and never to a keyboard nudge — an arrow key always moves by exactly its increment
- [ ] `Snap to guides` and `Snap to grid` are independently toggleable from `Format → Drawing`, persisted in preferences (SPEC-066), with snap-to-grid using a settable grid spacing shown in the current measurement unit
- [ ] A position dialog (`Format → Drawing → Position and size…`) sets a selected drawing's exact width, height, horizontal position, vertical position and rotation, each in the current measurement unit, with the horizontal and vertical position each expressed relative to a chosen reference
- [ ] The dialog offers a lock-aspect-ratio toggle for size, applies live as fields change, validates against the minimum shape size, and commits one undo entry per field change
- [ ] The dialog is the numeric path SPEC-068 promises: every geometric property adjustable by drag is adjustable there, reachable by keyboard, and it reflects the selection live
- [ ] Alignment, distribution and snapping are pure model operations in `src/lib/` taking laid-out rectangles from the gesture layer, keeping `src/lib` layout-free per SPEC-057's established pattern
- [ ] Guides never affect the model: they are transient render-layer artefacts, disappear on release, and are excluded from print and PDF output
- [ ] An e2e test places three shapes at arbitrary positions, aligns them left, distributes them vertically, and asserts their model offsets are exactly equal / exactly evenly spaced — not approximately

## Out of scope
Aligning text or table content (SPEC-014 and SPEC-023 own those). A visible
document grid, dot grid or ruler-dragged persistent guides — snapping
targets are computed, not user-placed. Auto-layout or diagram reflow.
Aligning across sections or pages. Equal-size ("make same size") commands.
Snapping a drawing to text baselines or to paragraph boundaries. Changing
anchoring, wrap or z-order semantics — SPEC-051 owns those.

## Reference

### What SPEC-057 provides
- A multi-selection spanning shapes, images, charts and containers, always
  resolving to **top-level** drawings, or to siblings within one container
  when a member is selected.
- Containers whose members share one coordinate space, with offsets
  measured from the container's top-left and always present.
- The established seam: the gesture layer passes **laid-out page
  rectangles** into pure model operations, so `src/lib` never computes
  layout. Alignment consumes exactly that.
- Renormalisation after any interior edit — recompute the members' union,
  fold its origin into the container's offsets, re-express member offsets —
  which every interior alignment must run.
- `MIN_SHAPE_SIZE_CM` (0.2cm) as the size floor, and rotation stored on the
  container rather than fanned out to members.

### What alignment aligns to
The reference rectangle, in priority order:

1. **Two or more drawings selected** → the union of their laid-out
   rectangles. Align left moves each so its left edge matches the union's
   left edge; the outermost drawing does not move. This is the behaviour
   every drawing tool has and the one users expect.
2. **One drawing selected** → the page's **content area** (inside the
   margins) by default, with `Format → Drawing → Align to` offering Margin
   (default), Page, and Column when the section has more than one
   (SPEC-030). The choice is remembered for the session.
3. **A member of a container selected, with siblings** → the container's
   box; alignment happens inside the container, then it renormalises.

A rotated drawing contributes its **unrotated frame rectangle**, matching
SPEC-057's rule for group bounds — so aligning a rotated shape aligns its
frame, which is both what the formats store and what users can predict.

```ts
// src/lib/drawingAlign.ts — pure, unit-tested
export type AlignEdge = "left" | "centerH" | "right" | "top" | "middleV" | "bottom";
export interface AlignRect { xCm: number; yCm: number; wCm: number; hCm: number }

/** Returns the new position for each input rectangle, in input order. */
export function alignRects(rects: AlignRect[], edge: AlignEdge, reference: AlignRect): Array<{ xCm: number; yCm: number }>;

/** Equalises gaps along one axis; the outermost two do not move.
 *  Fewer than three rectangles returns them unchanged. */
export function distributeRects(rects: AlignRect[], axis: "h" | "v"): Array<{ xCm: number; yCm: number }>;
```

Distribution equalises the **gaps between** the rectangles, not the
distance between their centres — the two differ whenever the drawings are
not the same size, and gap-equalisation is what looks right. The
distinction is stated because it is the likely disagreement, and it is
asserted with differently-sized rectangles in the unit test.

### Guides and snapping
Candidates, computed each drag frame from the drawings on the same page
plus the page itself:

- every other drawing's left, horizontal centre, right, top, vertical
  centre and bottom;
- the content area's four margins and its two centre lines;
- the column boundaries when the section has columns;
- with snap-to-grid on, the nearest grid line.

The dragged drawing's own six features are tested against each candidate.
Within the threshold, the position snaps and a guide line draws through the
matched feature, spanning from the dragged drawing to the matched one so it
is obvious *what* it aligned with. Multiple simultaneous matches draw
multiple guides.

The threshold is **8 screen pixels**, converted to centimetres through the
live zoom (SPEC-063), so it feels identical at 50% and 200% — a snap
threshold stated in centimetres would be unusable at one end or the other.
Holding Alt (Option) suppresses snapping for the duration of the drag.

Snapping runs in the gesture layer over laid-out rectangles, which is where
screen-pixel thresholds belong; the model operation it calls remains pure.
Guides render on the existing overlay layer and carry the `print-exclude`
convention the view already uses for on-screen-only chrome.

### The position dialog
Built on `BaseDialog.vue` with SPEC-058's `UiField`, showing every value in
the current measurement unit (SPEC-064) and accepting a typed unit suffix:

| Group | Fields |
| --- | --- |
| Size | Width, Height, Lock aspect ratio |
| Position | Horizontal: value + relative to (Margin / Page / Column / Character), Vertical: value + relative to (Margin / Page / Paragraph / Line) |
| Rotation | Degrees, 0–359 |

The "relative to" options map to the model's existing `relativeToH` /
`relativeToV` (SPEC-051), so the dialog exposes fields that already exist
rather than introducing placement concepts. Size fields respect
`MIN_SHAPE_SIZE_CM`, and a container's size fields scale its interior per
SPEC-057's interior-scaling rule.

For a selected **container**, the dialog edits the container. For a
selected **member**, it edits the member within its container, showing
positions relative to the container and renormalising on commit.

### Preferences added
Per SPEC-066's schema, three fields:

```ts
snapToGuides: boolean;      // default true
snapToGrid: boolean;        // default false
gridSpacingCm: number;      // default 0.25
```

### Specs this changes
- `specs/057-group-and-ungroup-drawings.md` — the follow-up it names is
  this spec. Nothing in its container model changes; interior alignment
  uses its coordinate space and its renormalisation rule.
- `specs/051-drawing-shapes-and-text-boxes.md` — drag and resize gain
  snapping; the gestures, their undo entries and their commit points are
  otherwise unchanged, and Alt always restores the current free-drag
  behaviour exactly.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — the Format menu
  gains a Drawing submenu holding Align, Distribute, the snap toggles and
  Position and size…, per its rule that a new command names its menu.
- `specs/068-responsive-touch-and-small-screens.md` — the position dialog
  is the numeric fallback it requires for drag-based drawing gestures.
