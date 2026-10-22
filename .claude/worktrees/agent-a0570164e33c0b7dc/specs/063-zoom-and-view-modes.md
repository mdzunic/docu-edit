---
id: SPEC-063
title: Zoom and view modes
status: draft
area: shell
depends_on: [SPEC-027, SPEC-047, SPEC-058]
---
## Why
The document renders at exactly one size: `PX_PER_CM` (96 CSS px per inch,
`src/lib/core/units.ts`) times its page size, forever. There is no zoom
control anywhere in the product — the word "zoom" appears in `src/` only
inside `src/lib/pdf/domExtract.ts`. On a 1280px window with the outline
pane open, an A4 page at 100% leaves the page abutting the pane on one
side and a wide empty gutter on the other; on a 4K display it is a stamp in
the middle of the screen; and a user who wants to see two pages at once, or
to read at 150% because the body font is 10pt, has no way to ask.

Zoom is also the prerequisite for two things people expect and this product
cannot do: seeing more than one page at a time, and getting the chrome out
of the way to write. And it has to be *cheap*, because SPEC-047 made
pagination the app's most expensive operation and its performance baseline
is guarded. That constraint decides the design: zoom is a **view
transform**, not a re-layout. The model, the pagination and every stored
geometry stay in centimetres at 96dpi; only the presentation scales. A
document zoomed to 150% must not repaginate, must not reflow, and must
break no page at a different place than at 100% — which is exactly what
distinguishes a word processor's zoom from a browser's.

## Acceptance criteria
- [ ] A zoom control sits at the right of the status bar: a `−` button, the current percentage as a menu button, and a `+` button
- [ ] The zoom menu offers 50%, 75%, 90%, 100%, 125%, 150%, 200%, then Fit width and Fit page, with the active entry checked
- [ ] `−` and `+` step through the same preset ladder, clamped to 50% and 200%; the percentage is also directly editable, accepting 25–400%
- [ ] Ctrl/Cmd+`+`, Ctrl/Cmd+`−` and Ctrl/Cmd+`0` (reset to 100%) are registered in `src/lib/shortcuts.ts` as `view.zoomIn`, `view.zoomOut`, `view.zoomReset`, and appear in the shortcut reference
- [ ] Ctrl/Cmd + mouse wheel, and pinch on a trackpad, zoom the document and are `preventDefault`ed so the browser's own page zoom does not also fire
- [ ] Zoom **never repaginates**: page breaks, line breaks and every stored geometry are identical at every zoom level, asserted by a test that paginates at 100%, zooms to 200% and 50%, and compares the page-break model exactly
- [ ] Zoom is applied as a scale transform on the pages container per "How zoom is applied", so changing it does not invalidate the SPEC-047 render cache and does not schedule a pagination pass
- [ ] Every pointer interaction stays correct under zoom: clicking to place the caret, dragging a selection, dragging a table boundary, drawing and resizing a shape, and dragging a ruler marker each land where the pointer is, at 50% and at 200% — because client-to-model conversion divides by the scale, per "Coordinate conversion"
- [ ] Every overlay stays registered with the content it annotates: comment anchors, find highlights, the shape anchor marker, table handles, selection outlines and the caret
- [ ] Fit width sizes the page to the available canvas width minus a fixed margin, and recomputes when the window resizes or a side pane opens or closes; Fit page fits one whole page in the visible height
- [ ] Zoom persists per document under `docu-edit:zoom:<documentId>`, so reopening a document restores how you were reading it; a document with no stored value opens at 100%
- [ ] Three view modes are offered from `View → Layout`: **Print layout** (today's paginated pages, the default), **Web layout** (one continuous column at the content width, no page boundaries, no headers/footers), and **Focus** (print layout with the chrome hidden), per "View modes"
- [ ] Focus mode hides the menu bar, toolbar and open panes, keeps the status bar, is entered and left by Ctrl/Cmd+Shift+F and by Escape, and restores exactly the panes that were open before
- [ ] Print layout shows multiple pages side by side when the zoom leaves room for them, laid out in a centred grid with the page gap between them; a caret move to another page scrolls the grid, not just the column
- [ ] Web layout changes only the view: the model, its section page sizes and everything saved are untouched, and switching back to print layout restores identical pagination
- [ ] Print, print preview and PDF export always render at 100% print layout, whatever the screen zoom and mode
- [ ] The ruler (`DocumentRuler.vue`) scales with the document, so a ruler centimetre is a document centimetre at every zoom level, and dragging its markers stays accurate
- [ ] The status bar shows the current zoom at all times, and Focus mode's status bar keeps the page indicator and the zoom control

## Out of scope
Any change to pagination, line breaking or page geometry (SPEC-027,
SPEC-047 own those). A per-section or per-page zoom. Zooming an individual
object. Outline view as an editing mode — SPEC-037 put that out of scope
and it stays out. Reading-mode reflow (re-laying out text to the window,
which *would* change line breaks). Presentation/full-screen document
projection. Touch pinch-zoom as a small-screen layout strategy — SPEC-068
owns touch.

## Reference

### The rendering this scales
`src/lib/pagination.ts` computes page geometry in CSS pixels from the
model's centimetres, through `PX_PER_CM` (`src/lib/core/units.ts`, 96/2.54
≈ 37.795). Pages render as `DocPage.vue` inside the `.pages` container that
`EditorContext.pageRef` points at (`src/views/document/context.ts`);
overlays position themselves relative to that container. SPEC-047 makes
pagination incremental and cached, with `idleScheduler` deferral, and pins
a performance baseline.

The whole point of this design is that none of the above changes. Zoom is a
property of the *presentation* of an already-computed layout.

### How zoom is applied
One CSS transform on the pages container:

```css
.pages {
  transform: scale(var(--zoom));
  transform-origin: top center;
}
```

with the container's laid-out size compensated so the scroll extent matches
the scaled content (a wrapper of the scaled dimensions, or
`width/height: calc(<natural> * var(--zoom))` on a positioned wrapper —
the implementation's choice, as long as scrollbars reflect the scaled
size).

Consequences that must be honoured rather than discovered:

- **Text stays crisp.** A CSS scale on live DOM re-rasterises text at the
  scaled size; it does not blur. This is why the transform is on the
  container and not a canvas bitmap.
- **The caret and selection are the browser's**, inside the transformed
  subtree, so they scale with it and need no special handling.
- **Canvas-painted content** (SPEC-051 shapes, SPEC-052 charts) is a
  raster inside the transform, so it *would* blur. Those canvases
  therefore multiply their backing-store size by
  `devicePixelRatio * zoom` and paint at that resolution, keeping their
  CSS size unscaled-by-them; the container transform does the rest. This
  is the one place zoom reaches into a painter.
- **Position: fixed overlays** — the contextual toolbars, the link chip,
  menus — are outside the transform and read their anchors through the
  conversion below.

### Coordinate conversion
Everything that turns a client coordinate into model space today does it by
subtracting a `getBoundingClientRect()` origin and dividing by `PX_PER_CM`
— `useContextMenus.ts`, `useCharts.ts`, `useComments.ts`,
`useFindReplace.ts`, the canvas gesture layer and the table resize code all
do exactly this. Under zoom every one of them must also divide by the
scale.

Rather than patching each call site, the conversion moves behind one pair
of helpers on the editor context, and the call sites use them:

```ts
// src/views/document/context.ts — added to EditorContext
/** The live view scale: 1 at 100%. */
zoom: Readonly<Ref<number>>;
/** Client point -> px within the unscaled pages container. */
clientToContentPx(x: number, y: number): { x: number; y: number };
/** Unscaled container px -> client point (for placing fixed overlays). */
contentPxToClient(x: number, y: number): { x: number; y: number };
```

`getBoundingClientRect()` on an element *inside* the transform already
returns scaled client pixels, so the rule is simple and worth stating
because getting it backwards is the likely bug: **rects measured inside the
transform are scaled; model geometry is not.** Divide measured rects by
`zoom` before comparing them with model values; multiply model values by
`zoom` before placing anything in client space.

A unit test drives `clientToContentPx` at 0.5, 1 and 2 and asserts
round-tripping through `contentPxToClient` is the identity.

### View modes
```ts
export type ViewMode = "print" | "web" | "focus";
```

- **Print layout** — today's behaviour, and the only mode print/PDF ever
  use. Multi-page grid when width allows: pages are laid out in a centred
  flex/grid of `floor(availableWidth / (pageWidth * zoom + gap))` columns,
  minimum one.
- **Web layout** — the same paginated model rendered as one continuous
  column: page boundaries, page gaps, headers, footers and page numbers
  are hidden; content width is the first section's content width; the page
  indicator reads "Page n of m" from the underlying pagination, which
  still runs. Nothing about the model or its section properties changes,
  so switching back is exact. This is a rendering flag, not a second
  layout engine.
- **Focus** — print layout with `--height-menu-bar` and `--height-toolbar`
  regions and all panes hidden. The pane set is captured on entry and
  restored on exit. Focus is a view state, not persisted across reloads.

Mode is persisted per document under `docu-edit:view-mode:<documentId>`,
except `focus`, which always starts off.

### Menu and status-bar placement
`View → Layout` holds three `menuitemradio` items (Print layout, Web
layout, Focus). `View → Zoom` holds the ladder plus Fit width and Fit page.
All register in the `editorCommands` Map as `view.layout-print`,
`view.layout-web`, `view.focus`, `view.zoom-in`, `view.zoom-out`,
`view.zoom-reset`, `view.zoom-fit-width`, `view.zoom-fit-page` and
`view.zoom-<n>`, per SPEC-049's rule that a new command names its menu.

The status-bar control renders after the page indicator and before the
word count, as `data-testid="zoom-control"` with
`zoom-out`/`zoom-in`/`zoom-percent` children.

### Specs this changes
- `specs/047-large-document-performance.md` — no mechanism change, and its
  baseline must not move: a zoom change schedules no pagination work, which
  this spec asserts directly rather than leaving to the baseline test.
- `specs/051-drawing-shapes-and-text-boxes.md` / `specs/052-charts.md` —
  their canvases take the view scale into their backing-store sizing, as
  above; nothing about their model or painting logic changes.
- `specs/034-print-and-export-pdf.md` — restated, not changed: export
  always renders print layout at 100%.
- `specs/021-keyboard-shortcuts-and-reference.md` — the reference gains
  Ctrl/Cmd+`+`/`−`/`0` and Ctrl/Cmd+Shift+F.
