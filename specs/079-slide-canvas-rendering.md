---
id: SPEC-079
title: Slide rendering — the slide canvas, placeholders and the filmstrip
status: draft
area: layout
depends_on: [SPEC-058, SPEC-069]
---
## Why
The presentations preparation milestone (2026-08-24, see
`docs/presentations-architecture.md`) ships `.pptx`/`.odp` open, library
listing, a byte-preserving download — and a *text outline* at `/slides/:id`.
Nobody can look at that outline and know what their deck looks like. This
spec replaces the outline with the real thing: each slide rendered
WYSIWYG — geometry, placeholder inheritance, text with its effective
formatting, images, shapes and backgrounds — plus a filmstrip to move
between slides. Still read-only: no command surface, no writer changes, a
download stays byte-identical. Every editing spec that follows (SPEC-080+)
manipulates the model this spec defines and re-renders through the renderer
this spec builds, so the model here is the presentation product's
counterpart of SPEC-002's document model: **this file is the normative
source for the slide model** the way `specs/002-create-blank-document.md`
is for the text model.

## Acceptance criteria
- [ ] `PresentationModel.slides` becomes `Slide[]` (frames with effective geometry and text formatting, per the Reference), and the model gains `slideWidthCm`/`slideHeightCm`; `SlidePreview` disappears as a stored shape — the outline (`title` + `lines`) is derived by a pure `slideOutline(slide)` so the library's slide count and existing outline expectations keep one source of truth
- [ ] Opening the in-repo `mini-deck.pptx` and `mini-deck.odp` fixtures yields the same frame structure from both readers: same slide count, per-slide frame kinds in z-order, titles and body lines equal to the outline the prep milestone asserted, and geometry in cm within 0.05cm across the two formats
- [ ] A `.pptx` placeholder with no `a:xfrm` of its own takes its position and size from its slide layout's matching placeholder (`type` + `idx`), else from the slide master's; a placeholder carrying its own `a:xfrm` keeps it — asserted by a fixture slide exercising all three levels
- [ ] Effective text formatting is resolved at read: run properties from the run's `a:rPr`, else the shape's `a:lstStyle`, else the layout/master placeholder chain (`p:txStyles` by placeholder role and paragraph level), else theme + built-in defaults (minor latin typeface, 18pt, `#000000`); `a:schemeClr` resolves through the master's `p:clrMap` into the theme's `a:clrScheme`, with `lumMod`/`lumOff`/`shade`/`tint` applied per the Reference formulas
- [ ] The `.odp` reader resolves the same effective values through ODF style chains: `style:style` `parent-style-name` chains across `content.xml` automatic styles and `styles.xml`, `fo:font-size`/`fo:font-weight`/`fo:font-style`/`style:text-underline-style`/`fo:color`/`style:font-name`, and placeholder frames (`presentation:class`) missing geometry fall back to the master page's matching placeholder
- [ ] Slide backgrounds render: explicit slide background, else the layout's (pptx), else the master's / drawing-page style's (both formats); solid fills exactly, gradient fills as their first stop's colour (documented approximation), image/pattern fills as white — never an error
- [ ] Images on slides render from `model.media` (populated by both readers; parts stay in `opaque` too so a no-edit download remains byte-identical), through the store's existing media split machinery — a deck with a 2MB photo does not duplicate those bytes into the non-media record
- [ ] Shapes with preset geometry render through `core/shapeCatalogue` + `core/shapeCanvas` (`paintShapeIntoCanvas`), with fill, outline colour/width, adjustment values and shape text; `.pptx` `prst` names and `.odp` shape elements map through the DrawingML tables hoisted core-ward per the Reference; rotation renders via CSS transform from `rotationDeg`
- [ ] `p:grpSp` / `draw:g` groups become a `group` frame: children flattened for rendering with the group's child-offset mapping applied (pptx `chOff`/`chExt`), rendered read-only, preserved verbatim on write (the group's source XML rides along); tables, charts, media and other unmodelled objects become `foreign` frames that render a neutral outlined placeholder box with a kind label ("Table", "Chart", "Media", "Object") at their bounds when geometry is parseable, and nothing otherwise
- [ ] `/slides/:id` shows the current slide on a stage that scales to fit the workspace (letterboxed, correct aspect ratio, crisp at devicePixelRatio ≥ 2 for canvas-painted shapes) and a filmstrip of numbered slide thumbnails; clicking a thumbnail, ArrowUp/ArrowDown/Home/End on the focused filmstrip, and PageUp/PageDown anywhere outside a text input change the current slide
- [ ] The filmstrip is a proper listbox for tests and assistive tech: `role="listbox"`/`role="option"`, `aria-selected`, each option labelled "Slide N of M — <title or 'untitled'>", and the stage announces slide changes via `aria-live="polite"`
- [ ] A presentation stored by the preparation milestone (its record still holding the outline-shaped `slides`) upgrades on open: the view re-parses from the retained package via the format reader, replaces the stored model under the same id/title, and renders — no dead "old format" state, no data loss, asserted by a unit test that stores a prep-shaped record and opens it
- [ ] Opening a deck, rendering it, and downloading it without edits still re-emits every package part byte-identical (`equalBytes` per part), `.odp` `mimetype` first and STOREd — the prep milestone's round-trip suite keeps passing with the media-split adjustment described in the Reference
- [ ] A slide whose XML is unparseable still renders (empty stage, its filmstrip slot present) — reading stays as tolerant as the prep readers
- [ ] e2e (`e2e/SPEC-079.spec.ts`): open `mini-deck.pptx` → stage shows slide 1's title text at larger rendered size than its body text; filmstrip shows the fixture's slide count; keyboard-navigate to slide 2 and its body line is visible; the read-only note from the prep milestone is gone

## Out of scope
Any editing (SPEC-080+). Slide size changes. Notes (SPEC-083), playback
(SPEC-084), transitions (SPEC-085), animations (SPEC-086) — `p:transition`,
`p:timing` and `anim:*`/`presentation:notes` content is captured verbatim
into the preservation fields the Reference defines, and otherwise ignored.
Vertical text, columns in text bodies, autofit shrink-to-fit (`normAutofit`
font scaling is applied when declared with an explicit `fontScale`,
otherwise text may overflow its frame — accepted), RTL layout, tab stops
inside slide text (SPEC-093 now models slide-number/date/footer placeholders
as live deck fields, superseding SPEC-079's frozen cached text; an `a:fld`/
`text:date` outside a field placeholder still renders its literal cached
text). Gradient/image fills beyond the stated
approximations. Embedded fonts. Hyperlinks on slides render as plain
styled text. Charts/tables/SmartArt/video render as labelled placeholder
boxes only. Editing or rendering master/layout views.

## Reference

### Dependencies restated
- The prep milestone's model and I/O:
  `src/lib/presentation/model.ts` (`PresentationModel`, `ParsedPresentation`),
  readers `presentation/pptx/reader.ts` (`parsePptx`) and
  `presentation/odp/reader.ts` (`parseOdp`), writers re-emitting
  `model.opaque` verbatim, assembly in `presentation/open.ts`. Round-trip
  contract: part-level `equalBytes` (`src/lib/core/equality.ts`), ZIP
  container bytes may differ; `.odp` `mimetype` first + STOREd.
- Boundary rules (enforced by `presentation/boundary.spec.ts` +
  `core/boundary.spec.ts`): `presentation/` imports only presentation
  siblings and `core/`; never `documentModel.ts`, `docx/`, `odf/`, the
  store, or Vue. New Vue components live under `src/components/` /
  `src/views/`, not in `lib/`.
- SPEC-058 design tokens and `AppShell`; SPEC-069's shape catalogue —
  already product-neutral in `core/shapeCatalogue.ts` (`SHAPE_DEFINITIONS`,
  `shapeDefinition`, `resolvedShapeAdjustments`), painted by
  `core/shapeCanvas.ts` (`paintShapeIntoCanvas`). Units:
  `core/units.ts` (`emuToCm`, `cmToPx`, `PX_PER_CM`, `halfPointsToPt`).
- Media: `core/media.ts` `MediaPart { mediaId, contentType, data(base64) }`;
  the store splits `model.media` into separate IDB records
  (`MEDIA_KEY_PREFIX` machinery in `documentStore.ts`) — presentations get
  this for free because `PresentationModel.media` already types over it.

### The slide model (normative — everything later builds on this)
Added to `src/lib/presentation/model.ts`; plain serialisable data, like the
text model. All lengths cm, sizes pt, colours `#rrggbb` lowercase hex.

```ts
export type PlaceholderRole = "title" | "subtitle" | "body" | "other";

export type SlideFill =
  | { type: "solid"; color: string }
  | { type: "none" };

export interface SlideRun {
  text: string;
  bold?: true;
  italic?: true;
  underline?: true;
  fontFamily?: string;   // effective; absent = renderer default stack
  sizePt?: number;       // effective; absent = 18
  color?: string;        // effective; absent = #000000
}

export interface SlideParagraph {
  runs: SlideRun[];
  align?: "left" | "center" | "right" | "justify";  // absent = left
  /** Bullet marker: "disc" renders •, "number" renders 1./2./… restarting
   *  per contiguous same-level run of numbered paragraphs. */
  bullet?: "disc" | "number";
  level?: number;        // 0..8 indent level; absent = 0
}

interface SlideFrameBase {
  /** Minted at read (crypto.randomUUID()); stable for the model's life —
   *  selection, undo and (SPEC-086) animation targets hang off it. Never
   *  serialised to either file format. */
  id: string;
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
  rotationDeg?: number;  // clockwise, like DrawingPlacement's convention
}

export interface TextFrame extends SlideFrameBase {
  kind: "text";
  role?: PlaceholderRole;          // present = it is a placeholder
  /** pptx only: the source `p:ph`'s type/idx attributes, written back
   *  verbatim on rebuild so Office keeps the inheritance chain. */
  phType?: string;
  phIdx?: number;
  paragraphs: SlideParagraph[];
  fill?: SlideFill;                // frame background; absent = none
}

export interface ImageFrame extends SlideFrameBase {
  kind: "image";
  mediaId: string;                 // key into model.media
}

export interface ShapeFrame extends SlideFrameBase {
  kind: "shape";
  geometry: ShapeGeometry;         // core/shapes id, e.g. "rect", "ellipse"
  adjustments?: number[];          // core/shapeCatalogue fractions
  fill?: SlideFill;
  stroke?: { color: string; widthPt: number };
  paragraphs?: SlideParagraph[];   // shape text, centred per bodyPr default
}

export interface GroupFrame extends SlideFrameBase {
  kind: "group";
  children: SlideFrame[];          // coordinates already in slide space
  /** The whole source element (`p:grpSp` / `draw:g`), verbatim — what the
   *  writers re-emit; children exist for rendering only until SPEC-082's
   *  ungroup command drops the xml and promotes them. */
  xml: string;
}

export interface ForeignFrame extends SlideFrameBase {
  kind: "foreign";
  xml: string;                     // verbatim source element
  /** Placeholder-box label: "Table" | "Chart" | "Media" | "Object". */
  label: string;
  /** True when x/y/width/height could not be parsed — render nothing. */
  noBounds?: true;
}

export type SlideFrame =
  | TextFrame | ImageFrame | ShapeFrame | GroupFrame | ForeignFrame;

export interface Slide {
  id: string;                      // minted at read, never serialised
  frames: SlideFrame[];            // z-order, back to front
  background?: SlideFill;          // effective (inheritance resolved)
  /** pptx: this slide's part path ("ppt/slides/slide2.xml"); ties the
   *  model slide to its verbatim part in `opaque`. Absent for .odp and
   *  (later) from-scratch slides. */
  partName?: string;
  /** odp: the source `draw:page` element markup, verbatim. SPEC-081's
   *  writer splices it back for untouched pages. */
  sourceXml?: string;
  /** Verbatim XML captured for rebuild-time retention (SPEC-081 uses it;
   *  this spec only fills it): pptx `p:bg`, `p:transition`, `p:timing`,
   *  `p:clrMapOvr`, `p:extLst` and the root `p:sld` attribute string; odp
   *  `presentation:notes`, `anim:*` children and the `draw:page` attribute
   *  string. Keys are those element names. */
  preserved?: Record<string, string>;
}
```

`PresentationModel` changes: `slides: Slide[]`; new `slideWidthCm: number`
and `slideHeightCm: number` (pptx `p:sldSz` cx/cy via `emuToCm`, default
33.867×19.05 when absent; odp from the master page's `fo:page-width`/
`fo:page-height`). `slideCount` and the derived `slideOutline(slide):
{ title?: string; lines: string[] }` (title = joined text of the first
`role:"title"` frame; lines = remaining frames' paragraph texts in order)
keep the prep milestone's outline observable — `roundTrip.spec.ts`'s
outline assertions are re-expressed over `slideOutline` with unchanged
expected values.

### Colour resolution (pptx)
`a:srgbClr val` → `#hex`. `a:sysClr` → its `lastClr`. `a:schemeClr val` →
master's `p:clrMap` attribute mapping (`tx1→dk1` etc.) → theme
`a:clrScheme` entry. Transforms, applied in order of appearance to the
resolved sRGB colour, each channel c ∈ [0,1], val = attribute/100000:
`shade`: c·val; `tint`: c·val + (1−val); `lumMod`/`lumOff`: convert to HSL,
L' = clamp(L·lumMod + lumOff), back to sRGB. All other transforms ignored.
The resolver lives in `presentation/pptx/color.ts` with direct unit tests
(known PowerPoint values: accent1 + lumMod 60% + lumOff 40% etc.).

### DrawingML reuse (enabling move, two-phase)
`docx/shape.ts` owns `GEOMETRY_TO_PRST`/`PRST_TO_GEOMETRY` and the pure
`a:xfrm`/`a:solidFill`/`a:ln`/`a:avLst` parsing helpers. Phase one: hoist
the tables and DOM-reading helpers (they take `Element`s, no docx types)
into `core/drawingMl.ts`; `docx/shape.ts` re-imports them so the docx
round-trip suites pin behaviour unchanged. Phase two: the pptx reader
consumes `core/drawingMl.ts`. Same pattern as the 2026-08-24 core moves
(memory: two-phase move). The ODF equivalent: the length parser
(`"2.54cm"`-style strings → cm) moves from `odf/` into `core/odfXml.ts`
(alongside `PRESENTATION_NAMESPACE`) for the odp reader.

### Reader rules
pptx, per slide (resolved through the slide's `.rels` → layout → its
`.rels` → master → theme, never assuming part paths):
- `p:sp` with `p:ph` → `TextFrame` with `role` (`title`/`ctrTitle` →
  "title", `subTitle` → "subtitle", `body` → "body", else "other"),
  `phType`/`phIdx` kept. `p:sp` without `p:ph`: `prstGeom prst="rect"` (or
  no prstGeom) with a `p:txBody` → `TextFrame`; any other known `prst` →
  `ShapeFrame`; unknown `prst`/`custGeom` → `ForeignFrame` label "Object".
- `p:pic` → `ImageFrame` (blip r:embed → media part → `MediaPart` with
  `mediaId` = part path; unresolvable embed → `ForeignFrame` "Media").
- `p:graphicFrame` → `ForeignFrame` ("Table" for `a:tbl`, "Chart" for
  chart graphic data, else "Object"); `p:cxnSp` with
  `prst="line"|"straightConnector1"` → `ShapeFrame` geometry "line", other
  connectors → `ForeignFrame` "Object".
- Bullets: explicit `a:buChar` → "disc", `a:buAutoNum` → "number",
  `a:buNone` → none; body-role paragraphs with no explicit bullet
  declaration in their chain default to "disc" (PowerPoint's body
  default); title/subtitle default to none.
- Empty-text placeholders (no `a:t` content) yield a `TextFrame` with
  `paragraphs: []` — rendered empty, no prompt text.

odp, per `draw:page`: `draw:frame` with `draw:text-box` → `TextFrame`
(role from `presentation:class`); with `draw:image` → `ImageFrame`
(`xlink:href` → package part, `mediaId` = part path); `draw:custom-shape`
→ `ShapeFrame` when its `draw:enhanced-geometry` `draw:type` maps through
the same core table (`odf/shape.ts`'s mapping hoisted with it), else
`ForeignFrame`; `draw:rect`/`draw:ellipse`/`draw:line` → `ShapeFrame`;
`draw:g` → `GroupFrame`; `table:table` in a frame → `ForeignFrame`
"Table"; anything else → `ForeignFrame` "Object".

### Rendering
- `src/components/presentation/SlideStage.vue` — one slide at a given
  scale: a stage div sized `cmToPx(slideWidthCm) × cmToPx(slideHeightCm)`,
  CSS `transform: scale(…)` to fit its container (the same
  fit-don't-reflow trick the text editor's zoom uses, SPEC-063). Frames
  are absolutely-positioned children in z-order: text → nested divs/spans
  with inline styles (pt → px via 96/72), bullets as marker spans with
  level indents of 0.95cm per level; images → `<img>` with the media
  `data:` URI resolver from `core/media.ts`; shapes → a `<canvas>` sized
  at `devicePixelRatio`, painted by `paintShapeIntoCanvas`, plus a
  centred text overlay div when the shape has paragraphs; groups →
  children rendered recursively; foreign → the labelled placeholder box
  (dashed `--border-subtle` outline, `--text-secondary` label) unless
  `noBounds`.
- `SlideFilmstrip.vue` — the listbox of thumbnails, each a scaled-down
  `SlideStage` (`pointer-events: none`) with the slide number; keyboard
  per AC. Both components are dumb: props + emits, no store access.
- `PresentationView.vue` wires them (current-slide state, the upgrade
  path, stage scaling via a `ResizeObserver`) and drops the outline list
  and the read-only note. The view keeps `useTheme()`; the stage's slide
  surface is always rendered light (a deck's colours are its own — white
  background unless the deck says otherwise), while the surrounding
  workspace uses the app tokens in both themes.

### Migration and media split
`getPresentation` returning a record whose model lacks `slideWidthCm`
marks it legacy; `PresentationView` (and later editing entry points)
await `upgradePresentationModel(model)` in `src/lib/presentation/open.ts`:
serialise the retained package via the format's writer (pure part
re-emission), re-parse via the new reader, keep `id`/`title`, store the
result via `storePresentation`. Media parts referenced by `ImageFrame`s
are copied into `model.media` **and remain in `opaque`** — the writers
keep re-emitting `opaque` alone, so the no-edit round trip stays
byte-identical and nothing about the writers changes in this spec.
(SPEC-081's rebuild writers take over media ownership; until then the
duplication is bounded by the store's media split, which persists
`model.media` once per record.)

### Tests
- `presentation/slideModel.spec.ts`: `slideOutline` derivation; frame
  kind classification tables (both formats).
- `presentation/pptx/color.spec.ts`: scheme resolution + transform math.
- `presentation/roundTrip.spec.ts`: extended fixtures (see below) —
  inheritance levels, effective text values, media split, byte-identical
  no-edit round trip (part-level `equalBytes`), legacy-record upgrade.
- `presentation/fixtures/build.ts` grows both mini decks: a slide with a
  layout-inherited placeholder, an explicitly-positioned shape
  (`ellipse`, solid fill, adjustment), an image (reuse the 2×2 PNG bytes
  convention from the docx fixtures), a group of two rects, a table →
  foreign, and a themed colour with `lumMod` — same content in pptx and
  odp so the cross-format equality AC bites. `e2e/fixtures/mini-deck.*`
  are regenerated via the existing
  `npx vite-node scripts/build-presentation-fixtures.ts`.
- `views/PresentationView.spec.ts`: filmstrip keyboard nav, aria
  labelling, upgrade-on-open.
- `e2e/SPEC-079.spec.ts` per the AC. Note (memory: e2e constraints): use
  relative size comparisons (title font > body font), never absolute
  pixel widths — Mac and CI fonts differ.

### Specs this changes
- The prep milestone's outline assertions in
  `presentation/roundTrip.spec.ts` and `e2e/presentations.spec.ts` are
  re-expressed over `slideOutline`/the stage with the same expected
  strings — coverage is preserved, never weakened.
- `specs/000-roadmap.md` gains this spec's phase; nothing else changes.
