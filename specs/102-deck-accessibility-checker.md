---
id: SPEC-102
title: Deck accessibility — alt text, reading order and the checker
status: draft
area: editing
depends_on: [SPEC-043, SPEC-079, SPEC-081, SPEC-082, SPEC-086, SPEC-089, SPEC-091, SPEC-094, SPEC-095, SPEC-097]
---
## Why
SPEC-082's out of scope names alt-text authoring, and SPEC-079 gives no frame a
`descr`, a title or even a name. The consequence is worse than a missing feature:
`imageFrameXml` (`src/lib/presentation/pptx/serialize.ts:261`) emits
`<p:cNvPr id="{n}" name="Picture {n}"/>` and `textFrameXml` (line 192) emits
`name="{Role}"`, so the first edit to any slide silently strips every frame's alt
text, its author-given name and its whole `a:extLst` — `a16:creationId` included —
out of the rebuilt part. A preserve-don't-destroy violation hiding inside a feature
gap, and it must close before anything can author alt text at all.

A deck is also where accessibility fails hardest: reading order *is* z-order, so the
order assistive technology announces a slide in is a side-effect of what was drawn
last; images arrive with nothing to announce; contrast is whatever the theme gave the
deck; a slide with no title placeholder is an unlabelled stop in the screen reader's
slide list. The word processor already ships the answer's architecture —
`specs/043-accessibility-checker.md` with `src/lib/accessibilityCheck.ts`,
`src/views/document/useAccessibilityCheck.ts` and `src/components/AccessibilityPane.vue`:
a rule catalogue with assigned severities, a one-sentence explanation per issue, a
mechanical fix where one exists, an idle-chunked scan. This file reuses it, hoists
its two product-neutral pieces into `src/lib/core`, and adds the deck rules on top.

**This file is the normative source for the frame accessibility model (`altText`,
`titleText`, `decorative`, `name`, `nvExtLst`), the `p:cNvPr` / `svg:title` +
`svg:desc` wire mappings, the deck rule catalogue with its severities, and the rule
that reading order is z-order.** SPEC-088's export consumes the alt text.

## Acceptance criteria
- [ ] `SlideFrameBase` in `src/lib/presentation/model.ts` gains `altText?: string`, `titleText?: string`, `decorative?: true`, `name?: string` and `nvExtLst?: string` per the Reference — inherited by every frame kind (text, image, shape, table, chart, media, group, foreign); `""` is never stored (absent means unset), `MAX_ALT_TEXT_CHARS = 250` is a soft cap that warns and never truncates, and all five stay plain serialisable strings so `JSON.parse(JSON.stringify(model))` still `toEqual`s the model
- [ ] Both readers populate them and dirty nothing: pptx from the frame's first `p:cNvPr` (`@descr`→`altText`, `@title`→`titleText`, `@name`→`name`, its `a:extLst` through `serializeElement` verbatim into `nvExtLst`, a descendant whose **local name** is `decorative` with `val` neither `"0"` nor `"false"` → `decorative: true`); odp from the `draw:frame`'s `svg:desc`/`svg:title` children read position-agnostically (`firstChildElement(frameEl, SVG_NAMESPACE, "desc")`, the rule `src/lib/odf/drawing.ts:95` already uses) and `@draw:name`; neither reader sets `dirty`/`notesDirty`/`structureDirty`, mutates `opaque` or rewrites `preserved`
- [ ] Both writers stop discarding them: `p:cNvPr` emits `id`, then `name` (`frame.name` else today's generated `"Picture {id}"`/`"Shape {id}"`/capitalised role), then `title`, then `descr`, with `nvExtLst` re-emitted verbatim as its child element; the odp serialisers emit `svg:title` before `svg:desc` at the position the golden pins; a frame with none of the five fields set writes byte-for-byte what it writes today
- [ ] Alt text on a `GroupFrame`/`ForeignFrame` is a **surgical splice** into `frame.xml`, never a regeneration: the command rewrites or inserts `descr`/`title` on the first `p:cNvPr`, or replaces/inserts `svg:title`/`svg:desc` under the root `draw:frame`/`draw:g`, leaving every other byte of `xml` identical — asserted by a substring-identity test, and the only sanctioned mutation of a preserved verbatim string in the codebase
- [ ] `src/lib/presentation/accessibility.ts` (framework-free; presentation siblings and `lib/core` only) exports `setFrameAltText`, `setFrameTitleText`, `setFrameDecorative`, `setReadingOrder(model, slideIndex, frameIds): boolean`, `DECK_ACCESSIBILITY_RULES`, `SEVERITY_BY_DECK_KIND`, `buildDeckScanTargets`, `scanDeckTarget`, `finalizeDeckScan` and `checkDeckAccessibility(model): DeckAccessibilityIssue[]`; every command ends with `markSlideEdited(slide, model.format)` and is one undo entry ("Alt text", "Mark decorative", "Reading order", "Sort reading order")
- [ ] An "Accessibility" toolbar toggle (`slides-accessibility-toggle`, `aria-pressed`, in `PresentationToolbar.vue`'s right-aligned `div.tb-group.tb-slide-actions` beside `slides-animate-toggle`, prop `accessibilityOpen`, emit `accessibility-toggle`) opens `slides-accessibility-panel` — a `<section aria-label="Accessibility for slide {n}">`, max-height 45%, `overflow-y: auto`, open state the persisted preference `presentationAccessibilityOpen` added to `src/lib/preferences.ts` in all three places (declaration, default `false`, `isBoolean` validator) — holding three stacked sections, the first offering the selected non-text frame `a11y-alt-text` (textarea + `a11y-alt-count`), `a11y-alt-title` and an `a11y-decorative` checkbox that disables and clears both, `a11y-alt-empty` with nothing selected and `a11y-alt-text-note` on a text frame, committing on blur or after SPEC-081's `TYPING_BURST_MS = 1000` burst
- [ ] Reading order section: `a11y-order-list` lists the slide's top-level frames **back-to-front**, numbered from 1 = announced first = backmost, each row `a11y-order-row` naming the frame through `frameAccessibleName`, with `a11y-order-up`/`a11y-order-down` calling `reorderFrame(model, slideIndex, frameId, "backward" | "forward")` (`src/lib/presentation/frames.ts:394`); the stage draws `a11y-order-badge-{n}` beside each frame while the section is open, and a move that would put a frame behind one whose `rotatedBounds` intersect it raises `a11y-order-overlap-warning` naming that frame — a warning, never a block, because reading order and z-order are one axis (see the Reference)
- [ ] Check section: `a11y-check-run` scans the deck into issues grouped by slide (`a11y-group-slide-{n}`, ordered by slide then reading order), each `a11y-issue` carrying its kind label, its one-sentence explanation and an `a11y-issue-fix` button only where a mechanical fix exists; `a11y-summary` shows the count and `a11y-clean` the "No accessibility issues found" state; selecting an issue sets `currentIndex` to its slide, calls `selectFrame(issue.frameId)` and scrolls the filmstrip to it
- [ ] The catalogue is exactly the ten rules of the Reference table with the severities it assigns — `missing-alt-text`, `low-contrast`, `slide-no-title`, `table-no-header-row` (Error); `duplicate-slide-title`, `tiny-font`, `media-no-captions`, `link-bare-url` (Warning); `empty-placeholder`, `reading-order-suspect` (Tip) — each assigned by SPEC-043's principle rather than by whichever text-side row it resembles; a decorative frame is exempt from `missing-alt-text` and from nothing else
- [ ] Contrast is judged by WCAG 2.1 AA — 4.5:1, or 3:1 for large text (effective 18 pt or more, or 14 pt or more when the run is bold) — the ratio rounded to two decimals before comparison and the foreground/background chain resolved exactly as the Reference states through SPEC-079's already-resolved `Slide.background`; the checker never invents a constant of its own and never composites an overlapping frame behind the text
- [ ] The six mechanical fixes are one undo entry each and re-run the whole scan: `add-alt-text` (focus `a11y-alt-text` for that frame), `contrast-ink` (set the run's colour to whichever of `#000000`/`#ffffff` scores the higher ratio against the resolved background), `add-slide-title` (mint a title `TextFrame` at SPEC-089's title slot geometry, else the Reference's fallback band, and start editing), `mark-header-row` (SPEC-091's `TableFrame.firstRow`), `remove-empty-placeholder` (`deleteFrame`) and `sort-reading-order` re-ordering one slide's frames by (`yCm`, then `xCm`) with a 0.5 cm tolerance band on y; the other four kinds render no button
- [ ] The scan chunks through `scheduleIdleWork` (`src/lib/core/idleScheduler.ts`) with SPEC-043's own `{ timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 }`, one frame per target plus a per-slide and a per-deck finalize, so a 200-slide deck finishes under 2 s without blocking the UI; `a11y-export` writes the report through DOM-free `exportAccessibilityReportHtml(model, issues): string` (`src/lib/presentation/a11yReport.ts`) and `triggerBlobDownload(bytes, fileName, "text/html")` (`src/lib/core/triggerDownload.ts`) as `{title} — accessibility.html`: self-contained, no external requests, one section per slide, a totals table and a `@media print` sheet, so the same file is the printed form
- [ ] The two product-neutral pieces move two-phase per the Reference — `core/contrast.ts` and `core/accessibility.ts` — each with its own spec; `src/lib/accessibilityCheck.ts` re-imports and re-exports so `accessibilityCheck.spec.ts` and every text-side expectation pass **unchanged**, and both `src/lib/core/boundary.spec.ts` and `src/lib/presentation/boundary.spec.ts` stay green — coverage preserved, never weakened
- [ ] Round trip: a no-edit download stays byte-identical part-for-part under `equalBytes` (`src/lib/core/equality.ts`, part-level after `unzip`, never container bytes); an alt-text edit on slide 2 leaves every other part byte-identical and writing the reopened model twice yields identical slide-part bytes; edit → save → reopen reproduces `altText`, `titleText`, `name`, `nvExtLst` and `decorative` in both formats under field-level `toEqual` plus `equalStructural` (`src/lib/equality.ts`) — whose `IGNORED_KEYS` drops frame `id`s and whose `XML_STRING_KEYS` contains `xml`, so a spliced group compares canonically while `nvExtLst` compares raw; `equalSemantic`/`resolve` are typed `(model: DocumentModel)` and unusable on a `PresentationModel`, and no criterion here claims them
- [ ] Fixtures: `src/lib/presentation/fixtures/build.ts` gains `buildPptxA11yFixtureBytes`/`buildOdpA11yFixtureBytes` and `EXPECTED_A11Y_ISSUES` under a new banner comment, never mutating the four mini-deck slides, each deck carrying one image with `descr` and one without, a run at 3.1:1 on its frame fill, a slide with no title, two slides sharing a title, an 8 pt run, an empty body placeholder and one decorative shape; `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec102-accessibility.pptx` and `.odp`
- [ ] e2e (`e2e/SPEC-102.spec.ts`): open `spec102-accessibility.pptx` → Accessibility → Check → the report groups the untagged image, the low-contrast run and the untitled slide by slide → click the image issue → that slide shows with the frame selected → type alt text → the issue disappears → fix the contrast issue in one click → reading order lists three rows; move the picture to position 1 and the overlap warning appears → export the report → download, reopen → the alt text, the recoloured run and the new order all come back and the untouched slides stay byte-identical

## Out of scope
Automatic alt-text generation (SPEC-043's exclusion carries over), OCR of text inside
images, WCAG conformance reporting, alt text on a group's *children* (a `GroupFrame`
is one target; the checker never descends into `children`), per-run language tags and
pronunciation, a tab order distinct from reading order (both formats have one order),
accessible names for SPEC-091 table cells or SPEC-092 chart series, caption tracks for
SPEC-095 media (the rule flags their absence, it cannot author them), remediating
SPEC-088's export or any PDF path, contrast against an overlapping frame's fill or a
background image, and scanning the editor chrome — SPEC-043 owns that.

## Reference

### Dependencies restated
- `specs/043-accessibility-checker.md` owns the checker; `src/lib/accessibilityCheck.ts`
  is what is reused: `AccessibilitySeverity = "error" | "warning" | "tip"`,
  `SEVERITY_BY_KIND`, `AccessibilityIssue { id; kind; severity; explanation; at:
  CaretPosition; endAt: CaretPosition; fix?; fixHeadingLevel? }`,
  `contrastRatio(hexA, hexB): number` (relative luminance, rounded to two decimals),
  `isLargeText(effectiveFontSizePt, bold): boolean`, and the
  `buildAccessibilityScanTargets`/`scanAccessibilityTarget`/
  `finalizeAccessibilityScan`/`checkAccessibility` chunking split.
  `AccessibilityIssue` is caret-addressed and **text-typed** — it does not move; the
  deck issue is its own shape. Its severity principle carries over verbatim: Error
  when a screen-reader user cannot reach the content or is given the wrong content,
  Warning when reachable but likely to mislead, Tip when correct but better
  structure is available.
- `specs/079-slide-canvas-rendering.md` owns the slide model: `SlideFrameBase` (not
  exported) `{ id: string; xCm; yCm; widthCm; heightCm: number; rotationDeg? }`,
  `id` minted `crypto.randomUUID()` at read and never serialised; `SlideFrame =
  TextFrame | ImageFrame | ShapeFrame | GroupFrame | ForeignFrame`; `SlideRun
  { text; bold?; italic?; underline?; fontFamily?; sizePt?; color? }` (absent
  `sizePt` = 18, absent `color` = `#000000`); `SlideFill = { type: "solid"; color } |
  { type: "none" }`; `Slide.background` is the **already-resolved effective**
  background (slide → layout → master / drawing-page style, at read). Lengths cm,
  sizes pt, colours lowercase `#rrggbb`. `slideOutline(slide): { title?: string;
  lines: string[] }` (`presentation/slideModel.ts:34`) is a slide title's only source.
- `specs/081-slide-text-editing.md` owns slide-part serialisation:
  `serializeSlideXml(slide, rels, slidePartPath)` with `modelled ?? preserved ?? ""`
  at every slot, `serializePage(slide, registry, usedNames, master, styleSource)`,
  the wire ids (pptx `p:cNvPr id` 2, 3, … in z-order; odp `de-f{n}` on `xml:id` and
  `draw:id`) and `markSlideEdited(slide, format)` (`presentation/text.ts:340` —
  `dirty = true`, and for odp `delete slide.sourceXml`); `TYPING_BURST_MS = 1000`.
- `specs/082-slide-objects-images-shapes-text-boxes.md` owns frame commands:
  `reorderFrame(model, slideIndex, frameId, target: "front"|"forward"|"backward"|
  "back"): string | null` and `deleteFrame` (`presentation/frames.ts`);
  `frameNoun`/`frameAccessibleName` (`views/presentation/useFrameManipulation.ts` —
  "Image"/"Group"/`frame.label || "Object"`/the shape catalogue label/
  "Placeholder"/"Text box"); `useFrameSelection` → `{ selectedFrameId,
  selectedFrame, select, clear, cycle }`; undo is `record(label)` **before** the
  mutation, `commit()` after.
- `specs/086-slide-animations.md` owns the golden-fixture-first process for bytes this
  repo cannot confirm — a hand-authored golden module, the serialiser asserted equal to
  it, a tolerant reader, cross-app checking as manual QA — in its four-test shape.
- Four models not yet built, each read for one thing only:
  `specs/089-slide-layouts-and-masters.md` — `layoutFor(model, slide)`'s `role:
  "title"` `LayoutSlot { xCm; yCm; widthCm; heightCm }`, which `add-slide-title` mints
  into; absent, it falls back to `DEFAULT_TITLE_BAND`, fractions `(0.08, 0.08, 0.84,
  0.16)` of `slideWidthCm`/`slideHeightCm` (SPEC-089's own `titleAndBody` title slot).
  `specs/091-tables-on-slides.md` — `TableFrame.firstRow?`, what `mark-header-row`
  sets. `specs/095-audio-and-video-on-slides.md` — `MediaFrame { kind: "media"; media:
  "video" | "audio" }`, which carries no caption or track field, exactly what
  `media-no-captions` reports. With either of those absent the frame stays a
  `ForeignFrame` `label: "Table"`/`"Media"` and only `missing-alt-text` applies.
  `specs/097-slide-text-depth.md` — `SlideRun.linkHref`, which exempts a run from
  `link-bare-url`; where it has not landed no run carries the field, so every bare-URL
  run fires — the same rule read against the narrower model.
- Equality: `equalBytes(a: Uint8Array, b: Uint8Array): boolean` from
  `src/lib/core/equality.ts`; `equalStructural(a: unknown, b: unknown)` from
  `src/lib/equality.ts`; `equalSemantic` is typed `(model: DocumentModel)` and **not
  usable on a `PresentationModel`**.

### The model (additions to SPEC-079's `SlideFrameBase`)
```ts
// src/lib/presentation/model.ts — inherited by EVERY SlideFrame kind
  /** What assistive technology announces. Absent = unset; "" is never stored.
   *  Soft cap MAX_ALT_TEXT_CHARS = 250 (warn, never truncate). */
  altText?: string;
  titleText?: string;   // pptx p:cNvPr@title; odp svg:title
  /** Skip me. Mutually exclusive with altText/titleText — setting it clears both. */
  decorative?: true;
  name?: string;        // pptx p:cNvPr@name, odp draw:name; absent = generated
  /** pptx only: the source p:cNvPr's whole `a:extLst`, verbatim, re-emitted
   *  unchanged so a16:creationId and every unknown extension survive a rebuild.
   *  Raw string under equalStructural (not in XML_STRING_KEYS). */
  nvExtLst?: string;

export interface DeckAccessibilityIssue {
  id: string;                      // stable within one scan; re-minted per scan
  kind: DeckAccessibilityKind;
  severity: AccessibilitySeverity; // from src/lib/core/accessibility.ts
  explanation: string;             // one sentence
  slideIndex: number;
  frameId?: string;                // absent for the two slide-level kinds
  fix?: DeckAccessibilityFix;
}
```

### Wire mapping (both formats)
**pptx.** `p:cNvPr` carries `id`, `name`, `descr`, `title` (ECMA-376 §19.3.1.12 /
§20.1.2.2.8); the reader takes attributes in any order, the writer emits the fixed
order. The word processor already does the identical thing one level up —
`docPr?.getAttribute("descr")` at `src/lib/docx/shape.ts:185`, `descr="…"` at line
341 — so the attribute name is confirmed in-repo, not guessed. **Decorative is the
one uncertain detail and is not invented here.** Office marks it with an
`<a:extLst><a:ext uri="{…}"><adec:decorative val="1"/>` extension whose GUID and
namespace URI this spec does not assert. Hence: the reader matches by **local name**
`decorative` anywhere inside `nvExtLst`, never by URI, and the writer's canonical
bytes are captured golden-first exactly as SPEC-086 established — hand-author
`presentation/pptx/a11yGolden.ts` (`DECORATIVE_EXT_GOLDEN`, `CNVPR_GOLDEN`,
`GOLDEN_A11Y_FRAMES`), assert serialiser output equals it, make the reader accept it
plus a PowerPoint-flavoured variant with reordered attributes; cross-app verification
is manual QA recorded in the PR, not CI. Setting `decorative` splices that one
`a:ext` into `nvExtLst`; clearing it removes it, every sibling byte-identical.
**odp.** `draw:frame` takes `svg:title`/`svg:desc` element children;
`src/lib/odf/drawing.ts:144` and `src/lib/odf/shape.ts:313` already write
`<svg:desc>` (last, after the content child) and read it position-agnostically at
line 95. **Where ODF puts them in `draw:frame`'s content model, and whether ODF has
any decorative marker at all** (`draw:decorative`? `loext:decorative`?) are not
confirmed in this repo: same golden-first process, `presentation/odp/a11yGolden.ts`
against ODF 1.2 §10.4.2/§14, tolerant reader, Impress QA in the PR. **If the golden
shows ODF has no such marker**, an odp deck writes neither child for a decorative
frame, the flag does not survive reopening, and the panel says so through
`a11y-decorative-odp-note` — a documented visible degradation, never a silent loss.
In both formats an alt-text edit is an ordinary content edit: `markSlideEdited` sets
`dirty` and drops the odp page's `sourceXml`; no new writer path is introduced.

### Reading order is z-order
In both formats the announcement order is document order of the shape tree —
`p:spTree`'s children, `draw:page`'s children — which is also paint order.
`Slide.frames` is that array, back to front. There is no second axis to expose, so
the panel is honest: row 1 is the backmost frame, and moving a row **up** is
`reorderFrame(…, "backward")`, which moves it further back visually. Where frames
overlap that is visible, so before a move the panel intersects the moved frame's
`rotatedBounds` (`src/lib/core/shapeGeometry.ts`) with those of the frame it would
pass and, on any non-empty intersection, shows `a11y-order-overlap-warning` naming it
("Moving Picture 3 earlier will put it behind Title"). It informs, never blocks.

### The rule catalogue, its severities and contrast
| Kind | Severity | Fires when | Fix |
| --- | --- | --- | --- |
| `missing-alt-text` | Error | a non-text, non-decorative frame has no non-blank `altText` | `add-alt-text` |
| `low-contrast` | Error | a run's ratio is below its AA threshold | `contrast-ink` |
| `slide-no-title` | Error | `slideOutline(slide).title` is `undefined` | `add-slide-title` |
| `table-no-header-row` | Error | a `TableFrame` has `firstRow !== true` | `mark-header-row` |
| `duplicate-slide-title` | Warning | two titles match after trim + lowercase + whitespace collapse | — |
| `tiny-font` | Warning | an effective `sizePt` below `MIN_SLIDE_FONT_PT = 12` | — |
| `media-no-captions` | Warning | a `MediaFrame` (video or audio) exists | — |
| `link-bare-url` | Warning | a run's text alone matches SPEC-043's `BARE_URL_RE` **and** the run carries no `linkHref` | — |
| `empty-placeholder` | Tip | a `TextFrame` with a `role` whose text is blank | `remove-empty-placeholder` |
| `reading-order-suspect` | Tip | `frames` order differs from (`yCm`, then `xCm`), 0.5 cm y-tolerance | `sort-reading-order` |

**Contrast.** The AA thresholds and the two-decimal rounding are AC-9's, computed by
SPEC-043's `contrastRatio` (`(Lmax + 0.05) / (Lmin + 0.05)`); what this file fixes is
the colour chain. **Foreground**, first that is set: `run.color` → `#000000`.
**Background**, first that is set: the frame's own `fill` when `{ type: "solid" }`
(an explicit `{ type: "none" }` falls through like unset, as SPEC-043's `highlight:
false` does) → `slide.background` when `{ type: "solid" }` → `#ffffff`. It terminates
by construction and is deliberately the renderer's own — `surfaceBackground(slide)` at
`src/components/presentation/slideRender.ts:120` returns `bg.type === "solid" ?
bg.color : "#ffffff"`, and a checker that disagreed with the screen would be worse than
none. Where SPEC-094 has landed, colours are `ModelColor` and the checker reaches a
literal only through its `resolveColor`.

### Reuse, and the two two-phase core moves
`src/lib/core` is already far richer than a text-only layer, so only what genuinely
sits outside it moves — each phase landing with its own tests and no call-site
behaviour change:
1. **`src/lib/core/contrast.ts`** (new). Phase one lifts `hexToRgb`,
   `srgbChannelToLinear`, `relativeLuminance`, `contrastRatio` and `isLargeText` out
   of `src/lib/accessibilityCheck.ts` (lines 94–131) unchanged, adding
   `WCAG_AA_NORMAL = 4.5`, `WCAG_AA_LARGE = 3` and `bestContrastingInk(background:
   string, inks: string[]): string`; `accessibilityCheck.ts` re-imports and
   re-exports so `accessibilityCheck.spec.ts` passes untouched. Phase two has
   `presentation/accessibility.ts` import them. `core/colors.ts` keeps the palettes
   — nothing in it computes luminance today.
2. **`src/lib/core/accessibility.ts`** (new). Phase one holds the shared vocabulary
   only — `AccessibilitySeverity`, `ACCESSIBILITY_SEVERITY_ORDER = ["error",
   "warning", "tip"]`, `severityLabel(severity)` — with `accessibilityCheck.ts`
   re-exporting `AccessibilitySeverity` so `SEVERITY_BY_KIND` and every text-side
   import are unchanged; phase two has the deck catalogue import it.
   `AccessibilityIssue`, `AccessibilityIssueKind`, `AccessibilityFixKind` and the
   detection engine stay text-side, typed against `DocumentModel`/`CaretPosition`.
`AccessibilityPane.vue` is reused as a **pattern only** — severity grouping,
one-sentence explanations, per-issue fix button, summary line, clean state, its
`asAccessibilityIssues(value: unknown)` runtime-prop cast — because its props are
typed to caret-addressed issues.

### UI
`src/views/presentation/useSlideAccessibility.ts`, deps `{ model, currentSlide,
currentIndex, record, commit, selectedFrameId, selectFrame, navigate }`, returning
`{ open, toggle, frame, altText, titleText, decorative, setAltText, setTitleText,
setDecorative, orderRows, moveUp, moveDown, overlapWarning, sortByPosition, issues,
scanning, summary, rescan, goToIssue, applyFix, exportReport }`; `orderRows` and
`issues` are fresh-array computeds (the `useAnimations`
`(currentSlide.animations ?? []).slice()` rule) or they never re-evaluate behind the
model's `shallowRef`. `src/components/presentation/AccessibilityPanel.vue` declares
exactly `{ controller: { type: Object, required: true }, slideNumber: { type: Number,
default: 1 } }`, emits nothing, aliases `const c = props.controller;` and uses runtime
`defineProps({ … })` only — ESLint parses `.vue` scripts with espree, which rejects
TS-only syntax; typed transforms go in a sibling `a11yRender.ts`.
`PresentationView.vue` gains **one** `useSlideAccessibility({ … })` call, its template
hookup and one toolbar prop/emit pair, paid for by extracting the stage-fit block
(`stageScale`, `recomputeScale`, the `resize` listener, the `ResizeObserver` and the
three `keydown` registrations) into `src/views/presentation/useStageFit.ts` — SPEC-094
owns that module, so where it has already landed an equivalent extraction of the same
size pays instead. `src/views/PresentationView.size.spec.ts`'s `LINE_BUDGET` then drops
by at least 10 from whatever it stands at when this lands (550 → 540 today) with a
ledger row appended; it never rises.

### Tests
- `src/lib/core/contrast.spec.ts` and `core/accessibility.spec.ts` cover the moved
  pieces on their new surface; `src/lib/accessibilityCheck.spec.ts` and every
  `docx`/`odf` suite keep every existing expectation through the re-exports; both
  `boundary.spec.ts` files stay green.
- `src/lib/presentation/accessibility.spec.ts` — each of the ten rules fired and not
  fired, the severity table pinned, `setReadingOrder`, the group/foreign splice's
  substring identity, `markSlideEdited` called by every command, the five fixes.
- `pptx/serialize.spec.ts` / `odp/serialize.spec.ts` — `p:cNvPr` attribute order,
  `nvExtLst` re-emitted verbatim, `svg:title`/`svg:desc` position, and a frame with
  no accessibility fields writing today's exact bytes; `pptx/a11yGolden.ts` and
  `odp/a11yGolden.ts` asserted both ways plus a flavoured variant, in SPEC-086's
  four-test shape.
- New `a11yRoundTrip.spec.ts` (per the round-trip criterion, `equalBytes` per part,
  `equalStructural` over the modelled subset) and `a11yFixtures.spec.ts`;
  `roundTrip.spec.ts` gains assertions and loses none.
- `a11yReport.spec.ts` (DOM-free, no external URLs in the output),
  `useSlideAccessibility.spec.ts`, `AccessibilityPanel.spec.ts`,
  `PresentationView.size.spec.ts`, `e2e/SPEC-102.spec.ts`.

### Specs this changes
- SPEC-079's frame model gains five inherited fields and its "no alt text, descr or
  accessibility title on any frame kind" gap closes; SPEC-081's serialisers gain the
  `p:cNvPr`/`svg:*` emission and stop discarding `@name`, `@descr`, `@title` and
  `a:extLst` on a rebuild — a fidelity fix, and every existing serialiser
  expectation for a frame without these fields is unchanged.
- SPEC-082's out of scope loses "alt-text authoring"; `reorderFrame` gains a second
  caller and no new semantics. SPEC-043's severity vocabulary and contrast maths move
  into `src/lib/core` behind re-exports, its catalogue, thresholds and pane behaviour
  untouched — coverage preserved, never weakened.
- SPEC-088's export renders `altText` into `alt`/`aria-label`, a decorative frame as
  `alt=""`. `specs/000-roadmap.md` gains this spec's phase row; nothing else changes.
