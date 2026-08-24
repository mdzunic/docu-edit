---
id: SPEC-088
title: Export a presentation to HTML — self-contained, with transitions and animations
status: draft
area: export
depends_on: [SPEC-084, SPEC-085, SPEC-086]
---
## Why
A deck's most common fate is being *sent to someone* — someone with no
Keynote, no PowerPoint, and no account on anything. A single `.html`
file that opens from disk in any browser and presents itself — slides
rendered faithfully, keyboard/click/touch navigation, the deck's
transitions and builds playing exactly as they do in SPEC-084 playback —
is the most portable deliverable this product can produce, and it
exercises everything Phase 13 built: the model, the transition and
animation subsets, and the playback semantics contract. The export is a
serialiser over the model (not a DOM snapshot), so it is unit-testable
in Node and stays in `src/lib/` per the repo's layering rule.

## Acceptance criteria
- [ ] The deck header's Download menu gains "Export to HTML (.html)" (testid `presentation-export-html`), enabled for any deck with ≥1 slide; it downloads `<title>.html` via `core/triggerDownload`
- [ ] The exported file is fully self-contained: one HTML document, inline CSS and JS, images as `data:` URIs, no external requests of any kind (fonts included — the export uses system font stacks with the model's family names first), works from `file://` and offline; the e2e asserts zero non-`file:` requests while driving it
- [ ] The rendered output is model-faithful: every slide is an absolutely-positioned scene at the deck's native geometry (cm → px at 96dpi), text frames with their effective formatting and bullets, image frames from media, shape frames as inline SVG built from the same `core/shapeCatalogue` geometry the canvas painter uses (fill, stroke, adjustments, rotation), group children rendered, foreign frames as their labelled placeholder boxes — visually consistent with the SPEC-079 stage
- [ ] The viewer scales the current slide to fit the window (letterboxed on black, aspect preserved, recomputed on resize), starting at slide 1 or the slide named by the URL hash (`#5` = slide 5, clamped); the hash tracks navigation via `history.replaceState` so a reload resumes and a copied URL deep-links
- [ ] Navigation reproduces SPEC-084's semantics from the shared key-map data: advance (click, ArrowRight, ArrowDown, Space, PageDown, Enter, N), back (ArrowLeft, ArrowUp, PageUp, P, Backspace), Home/End, number+Enter, touch tap + horizontal swipe; advance steps through animation builds before leaving a slide; back lands fully built; past the last slide, SPEC-084's end screen ("End of slide show — click or press any key to exit" — "exit" here closes into a final black state; a further advance restarts at slide 1 so kiosk loops are possible manually)
- [ ] Slide transitions play on forward slide changes per the SPEC-085 model (fade/push/wipe/cover/uncover/split, direction and duration honoured); `advanceAfterMs` auto-advances; input mid-transition completes it instantly; backward is instant
- [ ] Animations play per the SPEC-086 model: entrance-targeted frames hidden until built, click groups, withPrevious/afterPrevious timing, durations/delays, exits hiding their targets, mid-build input completing instantly
- [ ] `prefers-reduced-motion` renders every transition and build as an instant cut in the exported file too
- [ ] Speaker notes are included only when the export dialog's "Include speaker notes" checkbox is on (a small options dialog opens before export: notes checkbox, default off; testids `export-html-dialog`, `export-html-notes`, `export-html-confirm`); when included, pressing V in the viewer toggles a notes drawer for the current slide, and when excluded the notes text appears nowhere in the file's bytes
- [ ] The document is presentable and accessible as a document: `<title>` is the deck title (escaped), `lang` set, the visible slide's container is `aria-live`-announced ("Slide N of M"), non-visible slides are `display:none` + `inert` semantics (`aria-hidden`), and images carry empty `alt` (no alt authoring exists yet — documented)
- [ ] Untrusted content is escaped everywhere it lands (text, titles, font family names, colours are re-validated against `#rrggbb`): a deck whose text contains `</script><script>alert(1)</script>` exports to a file where that string renders as text and no script injection occurs (unit test asserts the payload is entity-escaped in the emitted bytes)
- [ ] The export runs entirely from the model + media (DOM-free `src/lib/presentation/htmlExport.ts`, unit-tested in Node); output for the fixture decks is deterministic (stable across two calls) and under 16MB for decks whose media total ≤ 10MB (base64 overhead documented at ~4/3)
- [ ] e2e (`e2e/SPEC-088.spec.ts`): open the pptx fixture (which carries a transition + animations from SPEC-085/086 fixtures) → export with notes off → capture the download, `page.goto(file://…)` → slide 1 visible; advance: build appears before slide change; slide change shows a transition intermediate state; deep-link `#2` reload lands on slide 2; request log shows no network fetches; repeat export with notes on → V toggles the note text visible
- [ ] Exporting never mutates the model (no dirty flags, no persistence writes — asserted)

## Out of scope
Print stylesheets / PDF-via-HTML, editing anything in the exported file,
embedding video/audio (foreign frames stay placeholders), font
embedding/subsetting, notes rendered for a second screen, analytics or
remote control, service workers, an index/overview grid, exporting a
subset of slides, `.zip` multi-file export (single file only), RTL
mirroring, exposing the export through SPEC-045's generation API
(future follow-up — the serialiser's purity makes it trivial later).

## Reference

### Dependencies restated
- SPEC-079 model + rendering semantics (fonts pt→px ×96/72, bullet
  indents 0.95cm/level, prompt text never exported, foreign placeholder
  boxes); `cmToPx`/`PX_PER_CM` from `core/units.ts`.
- SPEC-084: `playbackKeys.ts` shared key tables (imported by the
  serialiser and embedded as generated JS data — the tables are the
  single source; the export test asserts the emitted table equals the
  imported one), end-screen text, announce format, step machine
  semantics (`advance`/`back`/`goTo`, builds before slides).
- SPEC-085 transition model + per-type visual definitions (two stacked
  scenes, transform/opacity/clip-path, `transitionend`+timeout commit);
  SPEC-086 animation model + build grouping (`presentation/animations.ts`
  `groupIntoBuilds` is reused by the exporter to precompute build data).
- SPEC-083 `Slide.notes`. Media: `MediaPart { contentType, data }` —
  already base64; `data:` URI = `data:${contentType};base64,${data}`.
- Shape SVG: add `shapeSvgPath(geometry, widthPx, heightPx,
  adjustments): string` to `core/shapeCatalogue.ts` (pure; built from
  the same `ShapeDefinition` outline data `paintShape` consumes —
  polygon definitions become `M/L/Z` paths, ellipse/rounded cases their
  arc equivalents). The text product may later reuse it; its unit tests
  live with core.

### Output structure (normative)
```
<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{deck title}</title><style>…one static stylesheet…</style></head>
<body data-docu-edit-export="1">
 <main id="stage" role="region" aria-roledescription="slide show" aria-label="{title}">
   <section class="slide" id="s1" data-transition="push" data-transition-dir="left"
            data-transition-ms="600" data-advance-after="0">
     …frames: <div class="frame text">…, <img class="frame">, <svg class="frame">…
      animated frames carry data-anim-* attributes (effect, dir, ms, delay,
      build index, trigger)…
   </section>…
 </main>
 <div id="hud">…counter…</div><div id="end" hidden>…end screen…</div>
 <aside id="notes" hidden>…per-slide notes when included…</aside>
 <script>/* viewer: key tables (generated), step machine, transition
   engine, build engine, hash sync, resize scaler, reduced-motion */</script>
</body></html>
```
All dynamic content is emitted through one `esc()` helper (the
serialiser's, mirrored in nothing — the viewer script never injects
HTML). The stylesheet and viewer script are static template literals in
`htmlExport.ts` — no minification, readable output is a feature.
Numbers land in `data-*` attributes; the script reads them, so the
generated JS contains no per-deck code, only per-deck data.

### The serialiser
`exportPresentationHtml(model: PresentationModel, opts: { includeNotes:
boolean }): string`. Pure; throws on zero slides (the menu item is
disabled, the guard is for API callers). Frame emission mirrors the
stage renderer's decisions exactly where a browser is involved
(inline-style font sizing, marker spans for bullets); parity is pinned
by a unit test rendering the fixture deck's outline strings into both
the stage (jsdom) and the export string and comparing text content
per frame.

### The viewer script's contract (tested via e2e, mirrored in unit
tests of the emitted data)
State: `{ slide, build }`; `advance()` = next build else next slide
(playing its transition) else end; `back()` = previous slide, built;
number entry buffers digits 3s. Transitions clone SPEC-085's class
names so the static stylesheet is shared knowledge, not copied logic:
class `t-push t-left` etc. on the incoming/outgoing sections. Builds:
frames with `data-anim-build > current` get `.hidden`; playing a build
applies `.anim-fadeIn`-style classes with inline
`transition-duration`/`transition-delay`. `matchMedia("(prefers-reduced-
motion: reduce)")` short-circuits both engines. Auto-advance timers
clear on any manual input and on `visibilitychange`.

### UI
The export dialog is a small `BaseDialog` flow off the Download menu
(SPEC-080's menu); export happens on confirm. File name: the same
sanitisation `pptxFileName` applies, with `.html`.

### Tests
- `presentation/htmlExport.spec.ts`: escaping (the script-injection
  payload, quotes, `]]>`), determinism, self-containment (no `http`
  substring outside the doctype/lang boilerplate — assert via regex
  allowlist), notes inclusion/exclusion byte-level, data-attribute
  correctness for transitions/animations from a fixture model, key
  table equality with `playbackKeys.ts`, size accounting, no-mutation
  (deep-freeze the model input).
- `core/shapeCatalogue` spec additions: `shapeSvgPath` against known
  polygons + ellipse.
- e2e per AC (Playwright `download.saveAs` → `page.goto(file://)`;
  request assertions via `page.on("request")`).

### Specs this changes
- SPEC-080's Download menu gains this item (the menu was specified as
  the future home for Export entries).
- SPEC-084/085/086 gain an external consumer of their contracts; their
  shared modules (`playbackKeys.ts`, `groupIntoBuilds`, transition
  class vocabulary) are now explicitly public within the product.
