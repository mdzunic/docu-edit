---
id: SPEC-106
title: Presentation generation API, and Markdown to slides
status: draft
area: export
depends_on: [SPEC-045, SPEC-050, SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-083, SPEC-088]
---
## Why
SPEC-045 shipped a real generation surface for the word processor — a fluent
`DocumentBuilder`, a JSON document description with a published schema, byte readers
and three runnable examples — and every buyer for it has the same second question:
*and the deck?* Quarterly reviews, release notes and status decks are generated from
data far more often than they are typed. Today `src/lib/generate/` cannot produce a
single slide, and `PresentationModel` can only be reached by opening a file.

The second half is Markdown. SPEC-050 made `.md` a first-class **document** format —
open it, edit it on the page canvas, save it back — and deliberately said nothing
about slides. But "one deck per file, `---` between slides, `#` is the title, `-` is a
bullet" is how most decks that are written rather than dragged actually start, and the
inline vocabulary is the one SPEC-050's hand-written parser already implements.
SPEC-088 named the third piece itself, deferring "exposing the export through
SPEC-045's generation API" to a follow-up: this file is it. **It is the normative
source for `PresentationBuilder`, the JSON deck-description format, the slide-Markdown
grammar in both directions, and the rule that generation adds no serialisation code
of its own.**

## Acceptance criteria
- [ ] `src/lib/generate/presentation.ts` exports `PresentationBuilder` per the Reference: `new PresentationBuilder(options?: { title?: string; format?: "pptx" | "odp"; size?: "16:9" | "4:3" | { width: LengthInput; height: LengthInput } })` starting from `createBlankPresentation()`'s model with **no** first slide, every `add*`/`set*` returning `this`, and every argument validated on entry through `generate/validate.ts`'s existing `expect*`/`rejectUnknownKeys` helpers so a `DocumentGenerationError` naming the offending path (`"slides[2].bullets[0]: expected string or run array"`) always precedes any byte
- [ ] Slides are added by intent, not geometry: `addTitleSlide(title, subtitle?)`, `addBulletSlide(title, bullets, options?)` (`bullets` a nested array or `{ text, level }` rows, level 0–8), `addImageSlide(title, image)` and `addBlankSlide()` each append one `Slide` whose placeholder frames come from `slides.ts` itself, through two additive exports it gains — `titleSlideFrames(w, h)` (the 8/32/84/18 title + 8/52/84/12 subtitle pair `createBlankPresentation` already builds) and `titleBodyFrames(w, h)` (the 8/8/84/16 + 8/30/84/60 pair `addSlide` already builds), extracted with no behaviour change — so a generated deck's geometry is the editor's, never a second table (today the fractions are inline literals and `placeholder` is module-private)
- [ ] Per-slide content methods act on the slide most recently added and throw when there is none: `addBullets(items)`, `addTextBox(content, rect)`, `addImage(spec)` (`{ data: base64, contentType, x, y, width, height?, altText? }`, landed through SPEC-082's exported `insertImageFrame`, whose `{dir}image{n}.{ext}` minting rule is unchanged and whose private `mintMediaId` stays private), `setNotes(paragraphs)` (SPEC-083's `Slide.notes`, trailing empties stripped, `[]` never stored) and `setBackground(color)`
- [ ] Lengths and runs are SPEC-045's, unchanged: every rect field takes a `LengthInput` (`{ cm } | { pt } | { in } | { emu }`) through `generate/length.ts`'s `lengthToEmu` and lands in the model as centimetres via its `emuToCm`, rounded by `slides.ts`'s newly exported `r3`; every text field takes `RunSpec` **as `runs.ts` declares it** — a string, or `{ text, bold?, italic?, underline?, strikethrough?, fontFamily?, fontSizePt?, color?, highlight? }`, whose `RUN_KEYS` `rejectUnknownKeys` guard makes `sizePt` a `DocumentGenerationError` and not a synonym (`sizePt` is the *model*-side name on `SlideRun`) — mapped by a new `slideRuns.ts` beside `runs.ts` that renames `fontSizePt` → `SlideRun.sizePt` and drops `strikethrough`/`highlight`, which `SlideRun` has no field for, with a count, exactly as AC-9 does on the Markdown path; the pt/cm/`#rrggbb` conventions are SPEC-079's, and no unit helper is invented here
- [ ] Output mirrors `DocumentBuilder`'s four output methods (`builder.ts`: `toDoc`, `toBytes`, `toBlob`, `toBuffer`) with two deliberate, stated departures: `toDeck(): PresentationModel`, `toBytes(format?, options?): Promise<Uint8Array>`, `toBlob(...)` and `toBuffer(...)` take `format` **optionally**, because the deck builder carries one, where SPEC-045's are positional and required; and `fileName(format?)` is **new here** — `DocumentBuilder` has no such method and gains none — delegating to `pptxFileName`/`odpFileName`. Everything delegates to `writePptx`/`writeOdp` and **adds no serialisation code** — a construct the writers cannot emit is fixed in the writer, per SPEC-045's own rule; `options.modified?: Date` overrides the stamped timestamp, and `toBuffer` throws a clear `DocumentGenerationError` where `Buffer` is not global
- [ ] A generated deck carries no retained package (`opaque` absent, no `partName`/`sourceXml`/`preserved` on any slide), so both writers take SPEC-080's from-scratch path; generating the same deck twice with the same `options.modified` yields `equalBytes`-identical packages (`src/lib/core/equality.ts`, part-level after `unzip`), and the model stays plain JSON — `JSON.parse(JSON.stringify(deck))` `toEqual`s it
- [ ] `src/lib/generate/deckDescription.ts` exports `DeckDescription` and `deckFromDescription(description): PresentationModel`, a plain-JSON format mapping **1:1** onto the builder per the Reference table, published as `schema/deck-description.json` beside SPEC-045's `schema/document-description.json`; `PresentationBuilder.fromDescription(description)` is the static factory, unknown top-level or per-slide keys are validation errors, and an object with no `slides` describes an empty deck
- [ ] `src/lib/generate/slideMarkdown.ts` exports `parseSlideMarkdown(text, options?): PresentationModel` over the grammar in "Slide Markdown": `---` on its own line splits slides, `#`/`##` is the title placeholder, `-`/`*`/`+` and `1.` rows are body bullets nested by two-space indent, a standalone `![alt](data:…)` is an image frame, a `Notes:` line opens speaker notes to the end of the slide, and YAML front matter sets `title`/`format`/`size` and nothing else
- [ ] Inline formatting is SPEC-050's, reused rather than re-implemented: `src/lib/md/reader.ts`'s module-private `parseInline` gains a thin exported wrapper `parseInlineRuns(text): Run[]` — additive, no call site or expectation changed, so `md/reader.spec.ts` and `md/roundTrip.spec.ts` pass untouched — and `slideMarkdown.ts` maps each `Run`'s `text` + `formatting.bold`/`italic`/`underline`/`strikethrough` onto a `SlideRun`, dropping what a slide run has no field for and counting it
- [ ] Anything outside the grammar is **reported, never guessed**: a block quote, a table, a fenced code block, a raw HTML block or a heading deeper than `##` is dropped with a `DeckDropRecord { kind: DeckDropKind; location; reason }` — SPEC-008's `DropRecord` *shape*, re-declared in a new `src/lib/generate/deckDrops.ts` over its own closed `DeckDropKind` union (`shape`, `group`, `foreignFrame`, `table`, `chart`, `media`, `transition`, `animation`, `background`, `fill`, `stroke`, `runFormatting`, `slideGeometry`, `blockQuote`, `codeFence`, `rawHtml`, `heading`, `remoteImage`, `frontMatterKey`), because `src/lib/opaque.ts`'s `DropKind` is a closed text-side union with no member for any of them and importing it would pull `documentModel.ts` across the product boundary — returned as `SlideMarkdownParse { deck: PresentationModel; dropped: DeckDropRecord[] }` from a sibling `parseSlideMarkdownWithReport`; nothing is silently swallowed and nothing half-parses
- [ ] The inverse exists and is honest: `writeSlideMarkdown(model): { text: string; dropped: DeckDropRecord[] }` emits the canonical spelling of the grammar above from `slideOutline(slide)` plus each slide's body paragraphs, images as `![alt](data:…)` and notes under `Notes:`, and reports every construct Markdown cannot carry — shapes, groups, foreign frames, transitions (SPEC-085), animations (SPEC-086), fills, strokes, per-run colour and size, and slide geometry; a `.md` → deck → `.md` round trip satisfies `equalStructural` on the **model**, and `equalBytes` on the file is explicitly not asserted (SPEC-050's rule, restated)
- [ ] `src/lib/generate/readDeck.ts` mirrors `read.ts`: `readPptx(bytes, title?)` and `readOdp(bytes, title?)` wrap SPEC-080's byte readers behind `nodeDom.ts`'s `ensureNodeXmlGlobals()`, so generated bytes are re-read with no `File`, no disk and no browser; generate → write → read reproduces every construct this API can emit under `equalStructural` (`src/lib/equality.ts`) in both formats, geometry within ±0.01 cm
- [ ] SPEC-088's follow-up lands: `PresentationBuilder.toHtml(options?: { includeNotes?: boolean }): string` and a top-level `exportDeckHtml(model, options?)` wrap `presentation/htmlExport.ts`'s DOM-free `exportPresentationHtml` unchanged — same self-contained, no-external-request output — so a generated deck reaches HTML without a browser, and SPEC-088's own dialog and testids are untouched. SPEC-088 declares no default (its `opts: { includeNotes: boolean }` is required), so **this file chooses one**: `includeNotes` defaults to `false`, passed through explicitly; and because `exportPresentationHtml` throws on zero slides, `toHtml()` on a deck with no slides is refused at validation time with a `DocumentGenerationError` naming the empty deck, never SPEC-088's raw throw
- [ ] The boundary holds: everything new lives under `src/lib/generate/`, which is a **composition point** (the class `presentation/boundary.spec.ts`'s header already names alongside `openDocument.ts`), so `src/lib/core/boundary.spec.ts` and `src/lib/presentation/boundary.spec.ts` both stay green — nothing under `lib/core` or `lib/presentation` gains a `generate/` import, the deck builder never imports `documentModel.ts` or the `.docx`/`.odt` layers, and `src/lib/generate/index.ts` re-exports both builders while each stays separately importable
- [ ] Generation is DOM-free and fast: no module on the deck path touches `document`, `window` or a canvas (`slideMarkdown.ts` and `presentation.ts` are asserted importable and runnable under Node with no polyfill beyond `ensureNodeXmlGlobals` on the *read* path), and a 100-slide deck of title + six bullets builds and writes in under 2 seconds in Node — SPEC-045's own performance bar, one spec per format
- [ ] Documentation and examples match SPEC-045's bar: `docs/generate-api.md` gains a "Presentations" part covering every method, the deck description and the Markdown grammar (or a sibling `docs/generate-deck-api.md` it links, if the file would otherwise pass 800 lines), and `examples/` gains two runnable programs beside `invoice.js`/`report.js`/`letter.js` under the same `_impl/*.ts` + `lib/runTs.js` wrapper pattern — `deck.js` (a data-driven quarterly review) and `slides-from-markdown.js`
- [ ] e2e (`e2e/SPEC-106.spec.ts`): a Node-side unit path proves the bytes, and the browser path proves the hand-off — `page.evaluate` builds a three-slide deck through `PresentationBuilder`, stores it with `storePresentation(model)` — `storeDocument` is typed `(doc: DocumentModel)` and will not take a deck — and navigates to `/slides/:id`, whose guard reads `documentKindOf(id)` off the stored record, where the stage shows the title, the filmstrip shows three thumbnails, the notes panel shows the generated note, and a download reopens with the same three slides

## Out of scope
A hosted HTTP service (SPEC-045's exclusion carries over). Opening a `.md` file **as a
deck** from the library, the picker or a drop — `.md` stays SPEC-050's text format, and
`documentKinds.ts`, `openDocument.ts` and `OPENABLE_EXTENSIONS` are untouched; slide
Markdown is a generation input, reached through this API. Charts (SPEC-092), tables
(SPEC-091), media (SPEC-095), transitions (SPEC-085) and animations (SPEC-086) as
generatable constructs — each is reported when a deck-to-Markdown export meets one, and
each is a one-method follow-up over its own model once landed. Layouts, masters and
themes (SPEC-089/094): a generated deck uses the from-scratch set. Deck-level headers
and footers (SPEC-093). Reading an existing deck *into* the description format.
Rendering to PDF or images (compose with SPEC-104), and Keynote output.

## Reference

### Dependencies restated
- `specs/045-document-generation-api.md` owns the generation surface this extends, in
  `src/lib/generate/`: `DocumentBuilder` (`builder.ts`), `docFromDescription`
  (`description.ts`), `readDocx`/`readOdt` (`read.ts`), `DocumentGenerationError`
  (`errors.ts`), `LengthInput` + `lengthToEmu(value, path)` + `emuToCm(emu)`
  (`length.ts`), `RunSpec`/`RUN_KEYS`/`buildRuns` (`runs.ts`), the `expect*` +
  `rejectUnknownKeys` helpers (`validate.ts`) and `ensureNodeXmlGlobals()`
  (`nodeDom.ts`). Its rules — validate on entry, never emit an invalid file, **add no
  serialisation code** — carry over verbatim.
- `specs/079-slide-canvas-rendering.md` owns the deck model: `PresentationModel { kind:
  "presentation"; id; title; format: "pptx"|"odp"; slides: Slide[]; slideWidthCm;
  slideHeightCm; opaque?; media?; properties?; structureDirty?; odpContentEnvelope? }`;
  `Slide { id, frames, background?, notes?, … }`; `SlideFrame` = text/image/shape/group/
  foreign over `SlideFrameBase { id; xCm; yCm; widthCm; heightCm; rotationDeg? }`, `id`
  minted at read and never serialised. Lengths cm, sizes pt, colours lowercase
  `#rrggbb`, byte payloads base64; `slideOutline(slide)` is in `slideModel.ts`.
- `specs/080-create-presentation-and-slide-management.md` owns the writers and the
  from-scratch path: `createBlankPresentation()` (16:9 — `DEFAULT_SLIDE_WIDTH_CM =
  33.867` / `DEFAULT_SLIDE_HEIGHT_CM = 19.05` — `format: "pptx"`, one title + subtitle
  slide, **no** `opaque`), `addSlide(model, afterIndex)` and the placeholder fractions
  both use, `stripRetention`, and `writePptx`/`writeOdp` → `Promise<{ bytes:
  Uint8Array }>` with `pptxFileName`/`odpFileName`. With no retained package the
  writers build a fresh minimal one, which is exactly what a generated deck wants.
- `specs/081-slide-text-editing.md` owns the text model — `SlideParagraph { runs:
  SlideRun[]; align?; bullet?: "disc"|"number"; level? /* 0..8 */ }`, `SlideRun { text;
  bold?; italic?; underline?; fontFamily?; sizePt?; color? }` (note `sizePt`, not the
  API's `fontSizePt`), `DEFAULT_SIZE_PT = 18`, `MAX_INDENT_LEVEL = 8`; `specs/082-…`
  owns `insertImageFrame(model, slideIndex, media: MediaPart, rect)` and the
  `{dir}image{n}.{ext}` minting rule, with `MediaPart { mediaId; contentType; data }`
  in `src/lib/core/media.ts`; `specs/083-…` owns `Slide.notes?: string[]` (trailing
  empties stripped, `[]` never stored).
- `specs/050-markdown-open-edit-and-convert.md` owns document Markdown, and stays the
  only path a `.md` **file** opens by. Its hand-written parser is `src/lib/md/reader.ts`
  (private `parseInline(text, resolveCodeFormatting)` → `Run[]`, `HEADING_RE`,
  `LIST_ITEM_RE`, `STANDALONE_IMAGE_RE`, `FENCE_RE`, `splitFrontMatter`); its writer is
  `src/lib/md/writer.ts` (`MD_MIME_TYPE`, `mdFileName`, the private `escapeInline`);
  its equality rule is model-level `equalStructural`, never `equalBytes` on the file,
  because the serialiser normalises syntax.
- `specs/088-export-presentation-to-html.md` owns DOM-free `exportPresentationHtml(
  model, opts: { includeNotes: boolean }): string` (`presentation/htmlExport.ts`),
  where both `opts` and `includeNotes` are required and a zero-slide deck throws —
  the two facts AC-13 answers — plus its no-external-requests rule and the
  out-of-scope line this file resolves.
- Equality: `equalBytes(a, b)` from `src/lib/core/equality.ts`, part-level after
  `unzip`, never container bytes; `equalStructural(a, b)` from `src/lib/equality.ts`.
  `equalSemantic`/`resolve` there are typed `(model: DocumentModel)` and are **not**
  usable on a `PresentationModel`; no criterion here claims them.

### Where this lives, and why the boundary holds
`src/lib/generate/` already imports `documentModel.ts`, `docx/writer.ts` and
`odf/writer.ts`; adding the deck builder makes it import `lib/presentation` too, which
is legal and only there — `presentation/boundary.spec.ts`'s own header names the
app-layer composition points (`openDocument.ts`, `documentStore.ts`) and `generate/` is
one by the same definition. The discipline that keeps this honest is directional:
`presentation.ts`, `deckDescription.ts`, `deckDrops.ts` and `readDeck.ts` import
`lib/presentation` and `lib/core` only, and `slideMarkdown.ts` is the single module
touching both products — through exactly one added export (`parseInlineRuns`) and one
type (`Run`), which is why the drop record is declared in `generate/` rather than
imported from `opaque.ts`. Nothing under `lib/core` or `lib/presentation` imports
`generate/`, and the two builders never import each other: `index.ts` re-exports both,
so a text-only Node script keeps importing `./builder`.

### `PresentationBuilder`
```ts
const deck = new PresentationBuilder({ title: "Q3 review", format: "pptx" });
deck.addTitleSlide("Q3 review", "Prepared 30 September")
    .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%", "APAC flat"]])
    .setNotes(["Lead with Europe."])
    .addImageSlide("Pipeline", { data: png, contentType: "image/png" });
const bytes = await deck.toBytes();          // .pptx
```
A nested array in `bullets` is one level deeper than its parent, clamped at
`MAX_INDENT_LEVEL`; `{ text, level }` is the explicit form. `addImageSlide` fits its
image into the body rect at its own aspect, centred; `addImage` takes an explicit rect
and clamps it into the slide. `size: "4:3"` is SPEC-090's 25.4 × 19.05 cm, restated.

### The JSON deck-description format
1:1 with the methods above, validated by `deckFromDescription` and published as
`schema/deck-description.json`. Top level: `title?`, `format?`, `size?`, `slides?`; each
slide object carries `layout` plus that layout's fields, any other key being an error:

| `layout` | Fields | Method |
| --- | --- | --- |
| `"title"` | `title`, `subtitle?`, `notes?` | `addTitleSlide` |
| `"bullets"` | `title`, `bullets` (required), `notes?` | `addBulletSlide` |
| `"image"` | `title`, `image` (required), `notes?` | `addImageSlide` |
| `"blank"` | `textBoxes?`, `images?`, `background?`, `notes?` | `addBlankSlide` + per-slide methods |

`bullets` rows are a string, a nested array, or `{ text | runs, level? }`; `image` is
`{ data, contentType, width?, height?, altText? }`; `textBoxes` rows carry `text`/`runs`
plus `x`, `y`, `width`, `height` as `LengthInput`s.

### Slide Markdown
```md
---
title: Q3 review
---
# Q3 review
Prepared 30 September

---

## Highlights
- Revenue up 12%
  - Europe up 20%

Notes: Lead with Europe.
```
Rules, and only these. A `---` line **between** blocks starts a new slide (the front
matter's own delimiters are consumed first by `splitFrontMatter`, so a leading `---`
never mints an empty slide). The first `#` or `##` on a slide fills the title
placeholder; a slide whose first block is not a heading gets no title frame and its
content lands in the body. Plain paragraphs before any list become the subtitle on the
first slide and body paragraphs after it. List rows become body paragraphs with
`bullet: "disc"` (`-`/`*`/`+`) or `"number"` (`1.`), `level` = indent ÷ 2 clamped 0–8.
A standalone `![alt](src)` is an image frame when `src` is a `data:` URI and a dropped
`remoteImage` otherwise — nothing is fetched, per SPEC-088's self-containment rule —
and `Notes:` at the start of a line opens speaker notes to the end of the slide.
Everything else (tables, fences, block quotes, raw HTML, reference definitions, `###`+)
is dropped and reported, and front matter accepts `title`, `format` and `size`, any
other key being preserved in neither direction and reported once.

### Writing slide Markdown, and what it costs
`writeSlideMarkdown` is the inverse in canonical spelling — ATX headings, `-` bullets,
two-space nesting, `_italic_`/`**bold**` through `md/writer.ts`'s escaping discipline —
and its `dropped` list is the point of the feature, not a footnote: a deck is far
richer than its outline. One `DeckDropRecord` per shape, group, foreign frame, table,
chart, media frame, transition, animation, background, fill, stroke and per-run
colour/size/font, each naming its slide by 1-based number ("slide 4"). The round trip
asserted is `.md` → deck → `.md`, on the model.

### Tests
- `src/lib/generate/presentation.spec.ts` — every method's model output and validation
  errors (path text included, `sizePt` among the rejected keys), placeholder frames
  coming from `slides.ts`'s exports, length conversion, `fontSizePt` → `sizePt` run
  mapping with its drop count, `toDeck` purity.
- `deckDescription.spec.ts` — each `layout` row, unknown-key rejection, the empty
  description, plus a schema-shape test beside SPEC-045's.
- `slideMarkdown.spec.ts` — the grammar row by row, nesting, notes, data-URI images,
  every drop case, front matter, and the `.md` → deck → `.md` `equalStructural` round
  trip; `src/lib/md/reader.spec.ts`, `md/writer.spec.ts` and `md/roundTrip.spec.ts`
  pass **unchanged**, proving `parseInlineRuns` purely additive — coverage preserved,
  never weakened.
- `deckRoundTrip.spec.ts` — build → `writePptx`/`writeOdp` → `readPptx`/`readOdp` →
  `equalStructural`; byte determinism with a fixed `modified`; `equalBytes` per part
  on a second write.
- `deckPerformance.spec.ts` (the shape `performance.spec.ts` already has) for the
  100-slide bar; `deckHtml.spec.ts` for `toHtml`, its `includeNotes` default and the
  empty-deck refusal; `src/lib/core/boundary.spec.ts` and
  `src/lib/presentation/boundary.spec.ts` re-run unchanged; `e2e/SPEC-106.spec.ts`.

### Specs this changes
- `specs/088-export-presentation-to-html.md` — its out-of-scope line "exposing the
  export through SPEC-045's generation API" is resolved here; `exportPresentationHtml`
  itself, its dialog and its testids are untouched, and it gains no default.
- `specs/045-document-generation-api.md` — `src/lib/generate/` gains a second product
  and `docs/generate-api.md` a second part; `DocumentBuilder` (which gains no
  `fileName` and no optional `format`), the document description,
  `schema/document-description.json`, the three existing examples and every one of its
  acceptance criteria are unchanged.
- `specs/050-markdown-open-edit-and-convert.md` — `md/reader.ts` gains one additive
  export and nothing else; `.md` remains a text-document format everywhere in the
  product, and its supported subset, opaque-block rule and equality rule are unchanged.
- `specs/080-create-presentation-and-slide-management.md` — its from-scratch writer
  path gains a second caller and no new semantics, and `slides.ts` gains three purely
  additive exports with no behaviour change: `titleSlideFrames`, `titleBodyFrames` and
  `r3`, the frame arrays and rounding `createBlankPresentation`/`addSlide` already
  build inline. `src/lib/opaque.ts` and SPEC-008 are **not** touched: the deck side
  declares its own `DeckDropKind`, so `DropKind` stays the closed text-side union it
  is and `slideMarkdown.ts` still reaches the text side through one export, one type.
- `specs/000-roadmap.md` gains this spec's phase row and closes the backlog line "Deck
  export to PDF; HTML export via the generation API" (SPEC-104 took the first half).
