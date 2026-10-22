---
id: SPEC-078
title: Explicit anchors for floating drawings
status: draft
area: structure
depends_on: [SPEC-051, SPEC-052]
---
## Why
SPEC-051 anchored floating drawings by *convention*: a Shape (and, per
SPEC-052, a Chart) block sits in the block list "right after (or, when
inserted at a paragraph's start, right before) its anchor paragraph", and
everything that needs the anchor — the anchor marker, the re-anchoring math,
the offset origin — infers it from that adjacency
(`anchorParagraphBlockIndex` in `src/views/document/useShapes.ts`: the block
immediately before, if a paragraph, else the block immediately after).

That inference is ambiguous the moment two floating drawings anchor to the
same paragraph. In `[para, d1, d2]`, d2's inferred anchor flips to the
*following* paragraph — its anchor marker points at the wrong paragraph and a
drag re-anchors it against the wrong origin. `reanchorShape` splicing a
drawing in between `para` and `d1` silently changes d1's inferred anchor
without re-expressing its `offsetXCm`/`offsetYCm`, so d1's rendered position
jumps by the height of a paragraph it never moved relative to. `duplicateShape`
— which splices the copy directly after its source — manufactures the
`[para, d1, d2]` arrangement on every use, so "duplicate, then drag" is all
it takes to hit both bugs.

The same convention leaks into rendering: a floating drawing's offsets are
*measured* from its anchor paragraph's top at insertion time
(`finishShapeDraw` subtracts the caret paragraph's top), but *rendered* from
the drawing block's own slot in the page flow — a zero-size anchor element
that sits after (sometimes before) the anchor paragraph. The two disagree by
the anchor paragraph's height or a line box, which is the known
"lands ~14px below where it was dragged" insertion bug pinned in
`e2e/SPEC-051.spec.ts`.

This spec makes the anchor **explicit**: a stable reference from each
floating drawing to its anchor paragraph, maintained across edits, rendered
against, and carried through both file formats by writing the drawing
*inside* its anchor paragraph — which is how `wp:anchor` and
`draw:*[text:anchor-type="paragraph"]` express anchoring natively, so Word
and LibreOffice finally agree with us about what a drawing is anchored to.

## Acceptance criteria
- [ ] `Paragraph` and `DrawingPlacement` each gain an optional `anchorOrdinal`; a floating Shape/Chart and its anchor paragraph share one value, unique per paragraph document-wide
- [ ] A shared resolver (`src/lib/drawingAnchors.ts`) answers "which block index is this drawing's anchor?" — by ordinal when the drawing carries one, else by the adjacency fallback below — and every consumer (anchor marker, re-anchor math, renderer hoisting, writers) uses it; no consumer re-implements adjacency inference
- [ ] Inserting a floating shape, text box or chart pairs it with its target paragraph; duplicating a shape pairs the copy with the *same* paragraph; deleting a drawing (or making it `inline`) releases a pairing no other drawing still uses
- [ ] With two drawings anchored to one paragraph, each one's anchor marker points at that paragraph, and dragging either to another paragraph re-anchors only the dragged one — the other's pairing, offsets and rendered position are untouched
- [ ] Re-anchoring a drawing between another drawing and its anchor paragraph does not change the other drawing's resolved anchor or rendered position
- [ ] Splitting a paragraph keeps `anchorOrdinal` on the first half only; merging paragraphs keeps the surviving paragraph's own; a dangling reference (anchor paragraph deleted) degrades to the adjacency fallback, never crashes
- [ ] The renderer hoists each floating Shape/Chart entry to render immediately before its anchor paragraph's entry (`hoistFloatingDrawingEntries` in `src/lib/pageRender.ts`), and the floating anchor element takes no space in flow, so `offsetXCm`/`offsetYCm` are measured from the anchor paragraph's own top-left — a drawn shape lands where the pointer released it, including vertically
- [ ] `.docx` writes a floating Shape/Chart's `w:drawing` run inside its anchor paragraph's `w:p` (before the runs when the drawing precedes its anchor in the block list, after them otherwise); `.odt` does the same with the `draw:*` element inside the anchor's `text:p`
- [ ] Reading a paragraph that contains floating Shape/Chart drawings mints one fresh ordinal for that paragraph (its first split-off piece) and pairs those drawings with it; a drawing in a paragraph of its own — every file this editor saved before this spec — reads exactly as today, unpaired, and resolves via the fallback
- [ ] A document with two drawings anchored to one paragraph survives save → reopen in both formats with both pairings and both rendered positions intact
- [ ] Existing fixtures and ordinal-less models round-trip byte-shape-identically to today: the writer only changes output for drawings that carry a resolved `anchorOrdinal`

## Out of scope
`Image` blocks: SPEC-024 images have no `offsetXCm`/`offsetYCm` and render by
CSS float, not by anchor-relative offsets, so they gain nothing from an
explicit anchor; they keep their own-paragraph wire form and are never
paired. Anchor-*character* fidelity (Word anchors to a run position inside
the paragraph; we model the paragraph only). Rendering a drawing on its
anchor's page when pagination separates them — hoisting applies within a
page's own entries; a drawing whose anchor paragraph has no entry on the
drawing's page keeps its flow slot. `relativeToH`/`relativeToV` values other
than the shipped `column`/`paragraph` pair. ODT drawings anchored to *list*
paragraphs keep the own-paragraph wire form (a foreign element inside a
`text:list` run would split the list, a worse trade than re-deriving the
pairing on open). Comments/track-changes interactions beyond what the
paragraph split/merge rules above already give.

## Reference

### The model — one pairing, two fields
```ts
// documentModel.ts
export interface Paragraph {
  // ...
  /** Pairs this paragraph with the floating drawings anchored to it
   *  (SPEC-078): present only while at least one Shape/Chart references it.
   *  Unique per paragraph document-wide. Deliberately not named `id` — see
   *  IGNORED_KEYS in lib/equality.ts; a pairing hidden from equalStructural
   *  would let a broken anchor round-trip green, the same trap
   *  `groupOrdinal`'s name guards against. */
  anchorOrdinal?: number;
}

export interface DrawingPlacement {
  // ...
  /** The `anchorOrdinal` of this drawing's anchor paragraph (SPEC-078).
   *  Only ever set on floating (non-`inline`) Shape/Chart blocks. Absent
   *  means "resolve by adjacency" — every pre-SPEC-078 document. */
  anchorOrdinal?: number;
}
```

Minting is `max + 1` over every `anchorOrdinal` in the document (paragraphs
and drawings both, so a dangling reference can never collide with a fresh
pairing). The reader mints canonically — 1..n in document order — which is
what keeps read → write → read structurally stable.

**The block list still holds the drawing, in document order, at its anchor's
position** (SPEC-051's normative statement stands). The ordinal is not a
second copy of that fact; it is the disambiguation of it: *which* of the
blocks near a drawing is its anchor, stable while other drawings splice in
and out beside it.

### Resolution — ordinal first, adjacency fallback
`src/lib/drawingAnchors.ts` (framework-free, like every lib module):

- A drawing with an `anchorOrdinal` resolves to the **first paragraph in its
  own section** carrying the same value. First-in-section is deliberate: a
  paragraph split briefly leaves the value on both halves in the pure
  splitter's caller window, and first-match picks the top half — the piece
  that keeps the paragraph's top edge, which is what offsets are measured
  from.
- A drawing without one — or whose reference dangles — falls back to
  adjacency: walk **backward** over consecutive drawing blocks to the
  nearest paragraph; failing that, walk **forward** the same way; failing
  both, `null`. This subsumes SPEC-051's one-step inference and already
  fixes the `[para, d1, d2]` case for legacy documents, since the walk skips
  d1 instead of giving up on it.

`anchorParagraphBlockIndex` in `useShapes.ts` becomes a call to the
resolver; so do the re-anchor drag's old-anchor lookup and the renderer's
hoisting. Nothing else may re-derive an anchor.

### Command surface
All in `lib/shapes.ts` / `lib/charts.ts`, mirroring each other:

- **insert** (`insertShape`, `insertTextBox`, `insertChart`,
  `insertChartFromTable`): the anchor is the paragraph whose top the caller
  measured offsets from — the caret paragraph itself for an offset-0 or
  end-of-paragraph insert, the *first* half when the insert splits it, the
  table-adjacent paragraph resolved by fallback for `insertChartFromTable`
  (which may have none; then no pairing is minted). Reuse the paragraph's
  existing ordinal when it already anchors something; mint otherwise.
- **duplicate** (`duplicateShape`): the copy carries its source's
  `anchorOrdinal` — two drawings, one paragraph, by construction.
- **re-anchor** (`reanchorShape`): pair with the nearest-paragraph winner
  (mint/reuse on it), release the old pairing if the dragged drawing was its
  last referrer, and splice the block to sit immediately after the new
  anchor exactly as today. The `offsetYCm` adjustment is unchanged — with
  hoisting, "delta between the old and new anchor's tops" is now *exactly*
  the render-origin delta.
- **delete** (`deleteShape`, `deleteChart`) and **wrap changes**
  (`setShapeWrap`, `setChartWrap`): going to `inline` (or away entirely)
  releases the pairing when last-referrer; going from `inline` to floating
  pairs with the fallback-resolved paragraph.
- **paragraph split** (`textEditing.splitParagraph`,
  `documentModel.splitParagraphAt`): the first half keeps `anchorOrdinal`,
  the second half never carries it — same distribution family as bookmarks,
  which already live in both functions.
- **paragraph merge** (`deleteRange`'s cross-paragraph join): the surviving
  paragraph keeps its own fields, so a swallowed paragraph's referrers
  dangle and re-resolve by adjacency to the merged paragraph — the position
  consequence (offsets now measured from the merged paragraph's top) is
  exactly Word's behaviour for the same edit, so no special casing.

Release means: if no other drawing in the document references the ordinal,
delete the field from the paragraph — a paragraph carries `anchorOrdinal`
only while it anchors something, so idle documents stay clean and
`equalStructural` diffs stay quiet.

### Rendering — hoisting and the zero-footprint anchor
`hoistFloatingDrawingEntries(entries, model)` in `src/lib/pageRender.ts`,
called by `usePaginationView` on each realized page's assembled entries:

- For each entry of kind `"shape"`/`"chart"` whose block is floating, resolve
  the anchor block index and find the anchor paragraph's **first** entry on
  the page (`dataSection`/`dataBlock`; a split paragraph's earliest slice
  present). Move the drawing entry to immediately before it and adopt its
  `column`, preserving the relative order of drawings that share an anchor.
- No anchor entry on this page → the drawing entry stays where it is (out of
  scope note above).

The floating anchor element (`.doc-shape-anchor-floating`, and the chart
equivalent) becomes a zero-height in-flow **block** (`display: block;
height: 0; position: relative`): hoisted before its anchor paragraph it
contributes no line box and no height, so its top-left is the anchor
paragraph's own top-left and `left: offsetXCm / top: offsetYCm` on the
drawing measure from the same origin every command measures. A block — not
an inline-block (whose line box added the baseline's worth of drift) and
not `position: absolute` — because the wrap-spacer float the anchor hosts
for `square`/`topAndBottom` text exclusion must escape into the page's text
flow, and a block box is itself immune to floats, so the origin stays
pinned. This closes the SPEC-051 e2e's documented "lands ~14px below"
insertion bug, and that test's weakened assertion is strengthened to pin x
*and* y.

### Wire formats — anchoring by containment
The native encodings both express "anchored to paragraph P" by *containment*,
and that is what we now write. Ordinal **values** are never serialized; the
pairing is reconstructed from containment on read, so there is nothing to
leak into either format.

**`.docx` write**: a floating Shape/Chart whose anchor resolves is not
emitted at its own block slot; its `<w:r><w:drawing>…</w:r>` is emitted
inside the anchor paragraph's `<w:p>` — before the paragraph's runs when the
drawing precedes its anchor in the block list, after them otherwise, in
block order. `paragraphXml` gains optional leading/trailing run-XML
parameters. The section-close rule ("does the section end on a paragraph?")
counts absorbed drawings as absorbed — a section whose blocks end
`[…, para, shape]` now genuinely ends on a paragraph. Drawings with no
resolvable anchor (legacy models, dangling references, a section with no
paragraphs) keep today's own-paragraph form byte-for-byte.

**`.docx` read**: `parseBodyParagraphBlocks` already splits a paragraph
around its drawing runs. When the split yields at least one floating
Shape/Chart *and* at least one paragraph piece, mint the next canonical
ordinal onto the first piece and every such drawing. A drawing-only source
paragraph — the entire pre-SPEC-078 corpus — yields no paragraph piece and
therefore no pairing, exactly as today.

**`.odt`**: the same, symmetrically — `shapeElementXml`/the chart frame
emitted inside the anchor's `text:p` via `paragraphXml`'s existing
leading-XML seam plus a trailing one, and `parseTextBlocks` minting on the
containment it already splits apart. List-paragraph anchors are the one
carve-out (see Out of scope).

Round-trip stability of the *forms*: `[para(1), d(→1)]` writes
`<w:p>runs + drawing</w:p>` and reads back `[para(1), d(→1)]`; a foreign
file's drawing-before-text paragraph reads `[d(→1), para(1)]`, writes the
drawing back at the paragraph's start, and reads the same again. Word and
LibreOffice, opening either, anchor the drawing to the paragraph the editor
means — which today's own-empty-paragraph form gets wrong by one paragraph.

### Equality
`equalStructural` compares `anchorOrdinal` like any other field — the
pairing must be visible to it, per the IGNORED_KEYS note above. Two
consequences, both deliberate:

- read → write → read is exactly equal: the reader's canonical minting is
  deterministic and the writer preserves document order.
- an *edited* model's ordinal values may renumber across save → reopen
  (edits mint `max + 1`; the reader re-mints 1..n) while every pairing is
  preserved. Tests that compare an edited in-memory model against its
  reopened self must therefore compare *resolved anchors*, not raw ordinal
  values — `resolveAnchorBlockIndex` is the assertion helper, and this is
  the same class of rule as SPEC-051's "mask `dropped` before comparing".

### Tests
- `drawingAnchors.spec.ts`: ordinal resolution, both fallback walks, `null`,
  minting over dangling references.
- `shapes.spec.ts` / `charts.spec.ts`: every command rule above — in
  particular `[para, d1, d2]` re-anchor of d2 leaving d1's pairing, offsets
  and resolved anchor untouched, and `duplicateShape` pairing the copy.
- `documentModel.spec.ts` / `textEditing.spec.ts`: split keeps the ordinal
  on the first half only.
- `pageRender.spec.ts`: hoisting order, column adoption, shared-anchor
  ordering, anchor-not-on-page no-op.
- `docx/roundTrip.spec.ts` and `odf/roundTrip.spec.ts`: the containment
  cycle for one and for two drawings on a paragraph; the legacy path
  emitting today's bytes for ordinal-less models; foreign
  drawing-before-text order stability.
- `e2e/SPEC-078.spec.ts`: two shapes anchored to one paragraph — markers
  both point at it; dragging one to another paragraph moves only it (the
  other's box is pixel-identical); save as `.docx`, reopen, both positions
  survive. Plus the strengthened SPEC-051 draw-position assertion (x *and*
  y).

### Specs this changes
- `specs/051-drawing-shapes-and-text-boxes.md` — the "Anchoring and wrap"
  convention and the "a separate anchor index would be a second copy of a
  fact the block list already holds" design note are superseded as described
  here; the block-list placement rule itself stands. The e2e's pinned
  known-bug comment about drawn-position y is resolved.
- `specs/052-charts.md` — charts inherit the pairing through
  `DrawingPlacement`, no chart-specific behaviour beyond the command list
  above.
- `specs/008-round-trip-preservation.md` — no change: containment writing
  applies only to modelled drawings; opaque blocks keep their block-index
  anchoring, which absorption does not renumber.
