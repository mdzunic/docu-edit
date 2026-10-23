# Presentation generation API (SPEC-106)

The second product of the generation API (see [generate-api.md](generate-api.md)
for documents): a fluent `PresentationBuilder`, a JSON deck-description
format, and a slide-Markdown grammar in both directions, for generating
`.pptx`/`.odp` decks from data — no editor, no UI. It composes the product's
own slide model (`PresentationModel`, `src/lib/presentation/model.ts`) out of
the editor's own pieces — `createBlankPresentation()`'s model, the
`titleSlideFrames`/`titleBodyFrames` placeholder geometry `slides.ts` builds
for a blank deck and a new slide, SPEC-082's `insertImageFrame` and
`insertTextBoxFrame` — and delegates every byte to SPEC-080's writers
(`writePptx`/`writeOdp`). It adds no serialisation code of its own: a
construct the writers cannot emit is fixed in the writer.

Everything lives in `src/lib/generate/`, imported like any other `src/lib`
module:

```ts
import { PresentationBuilder } from "./src/lib/generate";        // both builders
import { PresentationBuilder } from "./src/lib/generate/presentation"; // the deck builder alone
```

The deck builder never imports the text model (`documentModel.ts`) or the
`.docx`/`.odt` layers, and the two builders never import each other:
`index.ts` re-exports both, and each stays separately importable.
Generation is DOM-free — it runs in plain Node with no polyfill; only
*reading* bytes back (`readPptx`/`readOdp`) borrows a `DOMParser` there.

Runnable examples: `node examples/deck.js` and
`node examples/slides-from-markdown.js` — see [Examples](#examples).

## Quick start

```ts
import { PresentationBuilder } from "./src/lib/generate";

const deck = new PresentationBuilder({ title: "Q3 review", format: "pptx" });
deck
  .addTitleSlide("Q3 review", "Prepared 30 September")
  .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%", "APAC flat"]])
  .setNotes(["Lead with Europe."])
  .addImageSlide("Pipeline", { data: pngBase64, contentType: "image/png" });

const bytes = await deck.toBytes();          // .pptx (the deck's own format)
const odp = await deck.toBuffer("odp");      // Node
const html = deck.toHtml();                  // one self-contained .html
```

Every `add*`/`set*` method validates its arguments on entry and throws a
`DocumentGenerationError` naming the offending path — e.g.
`slides[2].bullets[0]: expected string or run array …` — before it changes
anything, so an error always precedes any byte and never leaves a
half-built slide behind.

## `PresentationBuilder`

### `new PresentationBuilder(options?)`

| Option | Type | Default |
| --- | --- | --- |
| `title` | `string` | `"Untitled presentation"` |
| `format` | `"pptx" \| "odp"` | `"pptx"` |
| `size` | `"16:9" \| "4:3" \| { width: LengthInput; height: LengthInput }` | `"16:9"` |

The builder starts from `createBlankPresentation()`'s model **with no first
slide**. `"16:9"` is 33.867 × 19.05 cm and `"4:3"` SPEC-090's 25.4 × 19.05 cm;
an explicit size takes two [lengths](#lengths-and-runs), each between
2.54 cm and 142.24 cm (the deck-setup dialog's bounds). `format` decides what
`toBytes()`/`fileName()` produce when called without one, and which package
directory inserted pictures are named into (`ppt/media/` or `Pictures/`).

Every `add*`/`set*` method returns `this`.

### Adding slides

Slides are added by intent, never by geometry: each method appends one slide
whose placeholders come from `slides.ts` itself, so a generated deck's layout
is exactly what the editor gives a new slide.

#### `addTitleSlide(title, subtitle?)`

The title-slide pair `createBlankPresentation` builds: a title placeholder
(x 8% y 32% w 84% h 18% of the slide) and a subtitle placeholder (x 8% y 52%
w 84% h 12%). `title` and `subtitle` are [text content](#lengths-and-runs).
Without a subtitle the subtitle placeholder is still there, empty — as on a
blank deck's first slide.

#### `addBulletSlide(title, bullets, options?)`

The title + body pair `addSlide` builds: a title (x 8% y 8% w 84% h 16%) and
a body (x 8% y 30% w 84% h 60%) holding one bulleted paragraph per row.
`bullets` is an array of rows, each one of:

- a **string** — a row at the current depth;
- a **nested array** of rows — one level deeper than the array holding it,
  clamped at level 8 (`MAX_INDENT_LEVEL`);
- a **`{ text | runs, level? }` row** — the explicit form, `level` 0–8
  (absolute).

```ts
deck.addBulletSlide("Highlights", [
  "Revenue up 12%",                         // level 0
  ["Europe up 20%", ["Germany"]],           // levels 1 and 2
  { runs: [{ text: "Costs", bold: true }, " flat"], level: 0 },
]);
```

`options.style` is `"bullet"` (the default: every row a disc) or `"number"`
(numbered rows). A row with no text adds no paragraph — neither format's
reader keeps an empty one.

#### `addImageSlide(title, image)`

The title placeholder of `addSlide`'s pair, and an image fitted into the body
rectangle at its own aspect ratio, centred. `image` is
`{ data, contentType, width?, height?, altText? }`:

- `data`: base64 bytes (raw `Uint8Array` bytes are accepted too, outside JSON);
- `contentType`: `"image/png"`, `"image/jpeg"` or `"image/gif"` — the
  image's pixel size must decode, or the call is refused;
- `width`/`height`: optional [lengths](#lengths-and-runs). With neither, the
  image takes the largest size that fits the body at its aspect; with one,
  the other follows from the aspect; with both, that size. A size larger than
  the body is shrunk at its aspect to fit.
- `altText`: validated and then reported in [`dropped()`](#dropped) — the
  slide model's image frame has no alt-text field to keep it in.

The picture lands through SPEC-082's `insertImageFrame`, named by its
unchanged `{dir}image{n}.{ext}` rule (`ppt/media/image1.png`,
`Pictures/image1.png`, …).

#### `addBlankSlide()`

A slide with no frames, for the per-slide methods to fill.

### Per-slide content

These act on the slide most recently added, and throw a
`DocumentGenerationError` (`addBullets: no slide to act on — add one first …`)
when there is none.

#### `addBullets(items, options?)`

Appends rows — the same row shapes and `style` option as `addBulletSlide` —
to the slide's body placeholder. A slide without one (a title slide, a blank
slide) gains a body placeholder at `addSlide`'s geometry.

#### `addTextBox(content, rect)`

A free text box (no placeholder role) holding one paragraph of
[text content](#lengths-and-runs) at `rect = { x, y, width, height }` — four
lengths; `width`/`height` must be positive.

#### `addImage(spec)`

An image at an explicit rectangle: `{ data, contentType, x, y, width,
height?, altText? }`. Without `height`, the image's aspect gives it. The rect
is clamped into the slide: shrunk at its aspect when larger than the slide,
then moved fully onto it.

#### `setNotes(paragraphs)`

The slide's speaker notes (SPEC-083's `Slide.notes`): one string per
paragraph; a string holding a newline is split into paragraphs, as the
editor's notes panel does. Trailing empty paragraphs are stripped, and notes
that are all empty leave the slide with no notes field at all — `[]` is never
stored. A second call replaces the first.

#### `setBackground(color)`

A solid `#rrggbb` background (stored lowercase). Both writers emit it for a
generated deck (pptx as the slide's own `p:bg`, odp as its drawing-page
fill), and both readers read it back.

### Output

`toBytes`/`toBlob`/`toBuffer` mirror `DocumentBuilder`'s output methods, with
two deliberate departures: their `format` is **optional** (the deck carries
one), and `fileName(format?)` is new — `DocumentBuilder` has none.

| Method | Returns |
| --- | --- |
| `toDeck()` | `PresentationModel` — a deep copy, plain JSON (`JSON.parse(JSON.stringify(deck))` equals it) |
| `toBytes(format?, options?)` | `Promise<Uint8Array>` — the package bytes |
| `toBlob(format?, options?)` | `Promise<Blob>` typed `application/vnd.openxmlformats-officedocument.presentationml.presentation` or `application/vnd.oasis.opendocument.presentation` |
| `toBuffer(format?, options?)` | `Promise<Buffer>` — Node only; throws a `DocumentGenerationError` where `Buffer` is not global (use `toBlob`/`toBytes` there) |
| `fileName(format?)` | `"<title>.pptx"`/`"<title>.odp"` via `pptxFileName`/`odpFileName` (`"Untitled presentation"` for a blank title) |
| `toHtml(options?)` | `string` — see [HTML export](#html-export) |
| `dropped()` | `DeckDropRecord[]` — see [below](#dropped) |

`format` defaults to the deck's own. `options.modified?: Date` overrides the
created/modified timestamp stamped into `docProps/core.xml` / `meta.xml`
(default: now); with it fixed, generating the same deck twice yields
byte-identical parts. Writing to the other format is the same from-scratch
write the app's own "Download a copy as…" makes.

A generated deck carries no retained package — no `opaque` parts and no
`partName`/`sourceXml`/`preserved` on any slide — so both writers take
SPEC-080's from-scratch path and build a fresh, minimal package.

#### `dropped()`

What the builder accepted but the slide model cannot hold, as
[drop records](#drop-records): a run's `strikethrough` or `highlight`
(counted per slide), an image's alt text.

### `PresentationBuilder.fromDescription(description)`

A builder holding the deck a [JSON deck description](#json-deck-description-format)
describes, ready for more calls or for output. `deckFromDescription(description)`
returns the same deck directly as a `PresentationModel`.

## Lengths and runs

Both are SPEC-045's, unchanged — see [Lengths](generate-api.md#lengths) and
[Run specs](generate-api.md#run-specs). Every rect field takes a
`LengthInput` (`{ cm }`, `{ pt }`, `{ in }` or `{ emu }`), converted through
`lengthToEmu` and stored in centimetres (`emuToCm`, rounded to 3 decimals by
`slides.ts`'s `r3`).

Every text field — a title, a subtitle, a row's `runs`, a text box — is
*text content*: a plain string or an array of `RunSpec`s (a string, or
`{ text, bold?, italic?, underline?, strikethrough?, fontFamily?, fontSizePt?,
color?, highlight? }`). On a slide run, `fontSizePt` becomes the model's
`sizePt`; `bold`/`italic`/`underline`, `fontFamily` and `color` carry over;
`strikethrough` and `highlight`, which a slide run has no field for, are
dropped and counted in `dropped()`. The model-side name `sizePt` is **not** a
synonym: like any unknown run key it is a `DocumentGenerationError`
(`slides[0].title[0].sizePt: unknown property`).

## JSON deck-description format

A plain-JSON object mapping 1:1 onto the builder, validated by
`deckFromDescription` and published as
[`schema/deck-description.json`](../schema/deck-description.json) (beside
SPEC-045's `schema/document-description.json`).

```json
{
  "title": "Q3 review",
  "format": "pptx",
  "size": "16:9",
  "slides": [
    { "layout": "title", "title": "Q3 review", "subtitle": "Prepared 30 September" },
    {
      "layout": "bullets",
      "title": "Highlights",
      "bullets": ["Revenue up 12%", ["Europe up 20%", "APAC flat"]],
      "notes": ["Lead with Europe."]
    },
    { "layout": "image", "title": "Pipeline", "image": { "data": "iVBORw0…", "contentType": "image/png" } },
    {
      "layout": "blank",
      "textBoxes": [{ "text": "Scorecard", "x": { "cm": 2 }, "y": { "cm": 2 }, "width": { "cm": 20 }, "height": { "cm": 3 } }],
      "images": [{ "data": "iVBORw0…", "contentType": "image/png", "x": { "cm": 22 }, "y": { "cm": 2 }, "width": { "cm": 8 } }],
      "background": "#f4f6f8"
    }
  ]
}
```

Top level: `title?`, `format?`, `size?` (as the builder's options) and
`slides?` — an object with no `slides` describes an empty deck. Each slide
object carries `layout` plus that layout's fields; any other key, at the top
level or on a slide, is a validation error.

| `layout` | Fields | Builder calls |
| --- | --- | --- |
| `"title"` | `title`, `subtitle?`, `notes?` | `addTitleSlide(title, subtitle)` |
| `"bullets"` | `title`, `bullets` (required), `style?`, `notes?` | `addBulletSlide(title, bullets, { style })` |
| `"image"` | `title`, `image` (required), `notes?` | `addImageSlide(title, image)` |
| `"blank"` | `textBoxes?`, `images?`, `background?`, `notes?` | `addBlankSlide()`, then `addTextBox` / `addImage` per entry and `setBackground` |

`notes` is `setNotes(notes)` on every layout. `bullets` rows are a string, a
nested array, or `{ text | runs, level? }`; `image` is `{ data, contentType,
width?, height?, altText? }`; `images` entries are `addImage`'s spec;
`textBoxes` rows carry `text` or `runs` plus `x`, `y`, `width`, `height`.
`style` on the bullets layout is `addBulletSlide`'s options (`"bullet"` or
`"number"`), so the mapping stays 1:1. Errors are named from the
description's root: `slides[2].bullets[0]: expected string or run array …`.

## Slide Markdown

One deck per text — the way decks that are written rather than dragged
usually start. Slide Markdown is a **generation input**, reached through this
API: a `.md` *file* still opens only as a text document (SPEC-050).

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

### The grammar

These rules, and only these:

| Construct | Becomes |
| --- | --- |
| `---` on its own line, between blocks | the start of a new slide. A leading or trailing `---` mints no slide; front matter's own delimiters are consumed first; a `---` inside a fenced block does not split (one in speaker notes ends them with the slide — write `\---` for a literal line). An empty stretch between two separators is a blank slide. |
| `#` or `##` as a slide's **first** block | the title placeholder. A slide whose first block is not a heading has no title frame, and its content lands in the body. |
| Plain paragraphs | on the **first** slide, before any list: the subtitle (a title slide — `addTitleSlide`'s pair); otherwise body paragraphs. Soft-wrapped lines join with a space; a hard break (two trailing spaces or a trailing `\`) starts a new paragraph (a slide paragraph is one rendered line). |
| `-`, `*`, `+` rows / `1.` (or `1)`) rows | body paragraphs with `bullet: "disc"` / `bullet: "number"`, `level` = indent ÷ 2, clamped 0–8. Each row is one line. |
| A standalone `![alt](data:…)` line | an image frame fitted into the body area, as `addImageSlide` fits one. Base64 and percent-encoded `data:` URIs of PNG, JPEG or GIF are decoded. |
| `Notes:` at the start of a line | speaker notes to the end of the slide: the rest of that line, then every following line verbatim (trailing whitespace trimmed; a leading `\` escapes the line). |
| YAML front matter | `title`, `format` (`pptx`/`odp`) and `size` (`16:9`/`4:3`) — nothing else. |

Inline formatting is SPEC-050's own parser, reused through `md/reader.ts`'s
`parseInlineRuns`: `**bold**`, `*italic*`/`_italic_`, `[links](https://…)`
and `\` escapes carry onto the slide run (bold, italic, `linkHref`);
`~~strikethrough~~` and a `` `code` `` span's formatting have no slide-run
field — their text is kept, the formatting dropped and counted.

The slide layouts are the builder's: a first slide with a title and plain
paragraphs is a title slide (with a body placeholder too if it also holds
list rows); any other slide with a title is a title + body slide, except one
holding only pictures, which keeps just its title (as `addImageSlide` does).

### What is reported, never guessed

Anything outside the grammar is dropped and reported as a
[drop record](#drop-records) — nothing is silently swallowed, and nothing
half-parses:

| Source | `kind` |
| --- | --- |
| a block quote | `blockQuote` |
| a fenced code block | `codeFence` |
| a table | `table` |
| a raw HTML block, a link reference definition, a thematic break other than `---` | `rawHtml` |
| a `###`+ heading, or a heading that does not open its slide | `heading` |
| an image whose `src` is not a `data:` URI (nothing is fetched — SPEC-088's self-containment rule) | `remoteImage` |
| an image that is not PNG/JPEG/GIF or does not decode; an image's alt text | `media` |
| a front-matter key other than `title`/`format`/`size`, a value they cannot take, a comment | `frontMatterKey` |
| strikethrough, code-span formatting | `runFormatting` |

### `parseSlideMarkdown(text, options?)` / `parseSlideMarkdownWithReport(text, options?)`

`parseSlideMarkdown` returns the `PresentationModel`;
`parseSlideMarkdownWithReport` returns `{ deck, dropped }`
(`SlideMarkdownParse`). `options` — `{ title?, format?, size? }`, as the
builder's — fill in whatever the front matter does not set.

```ts
import { parseSlideMarkdownWithReport, readPptx } from "./src/lib/generate";
import { writePptx } from "./src/lib/presentation/pptx/writer";

const { deck, dropped } = parseSlideMarkdownWithReport(markdown, { title: "Release notes" });
for (const record of dropped) console.warn(`${record.location}: ${record.reason}`);
const { bytes } = await writePptx(deck);
```

### `writeSlideMarkdown(model)`

The inverse: `{ text, dropped }`. It writes the canonical spelling of the
grammar — front matter (`title`, `format`, `size`), `#` for the first slide's
title and `##` for the rest, `-` bullets and `1.` numbers nested by two
spaces, `**bold**`/`_italic_`/`[links](…)` escaped as `md/writer.ts` escapes
them, every picture as `![](data:…)`, and notes under `Notes:` — from each
slide's outline (`slideOutline`) and body paragraphs.

Its `dropped` list is the point, not a footnote: a deck is far richer than
its outline. One record per shape, group, foreign frame, table, chart, media
frame, transition, animation, background, fill, stroke and per-run colour/
size/font/underline, plus `slideGeometry` for a text box or picture whose
position Markdown cannot keep and for a slide size other than 16:9 or 4:3 —
each naming its slide (`"slide 4"`), or `"deck"`.

### The round trip

`.md` → deck → `.md` is asserted with `equalStructural` on the **model**: the
deck parsed from the written text equals the deck parsed from the original.
`equalBytes` on the text is explicitly not asserted — the writer normalises
spelling (SPEC-050's rule) — but its output is canonical: writing a deck
parsed from it again gives the same text.

## Drop records

```ts
interface DeckDropRecord {
  kind: DeckDropKind;
  location: string;   // "slide 4", "front matter" or "deck"
  reason: string;
}
```

SPEC-008's `DropRecord` shape, over the deck side's own closed union
(`src/lib/generate/deckDrops.ts`): `shape`, `group`, `foreignFrame`, `table`,
`chart`, `media`, `transition`, `animation`, `background`, `fill`, `stroke`,
`runFormatting`, `slideGeometry`, `blockQuote`, `codeFence`, `rawHtml`,
`heading`, `remoteImage`, `frontMatterKey`. (The text product's `DropKind`
has no member for any of these, and importing it would cross the product
boundary.)

## HTML export

`builder.toHtml(options?)` and the top-level `exportDeckHtml(model, options?)`
wrap SPEC-088's DOM-free `exportPresentationHtml` unchanged: one
self-contained `.html` document — inline CSS and viewer script, pictures as
`data:` URIs, no external requests — reached without a browser.

- `options.includeNotes` defaults to **`false`** (SPEC-088 declares no
  default) and is passed through explicitly; `true` embeds the speaker-notes
  overlay.
- A deck with no slides is refused at validation time with a
  `DocumentGenerationError` (`deck: an empty deck (no slides) cannot be
  exported to HTML …`), never SPEC-088's own throw.

## Reading generated decks

`readPptx(bytes, title?)` / `readOdp(bytes, title?)` parse a `Uint8Array`/
`ArrayBuffer`/`Buffer` straight into a `PresentationModel` with SPEC-080's
byte readers — the assembly the app's own open path uses — with no `File`,
no disk and no browser. `title` defaults the way a blank deck's does. As with
`readDocx`/`readOdt`, outside a browser they first borrow `DOMParser`/
`XMLSerializer`/`Node` from the project's `jsdom` dependency
(`ensureNodeXmlGlobals`); writing never needs this.

Generate → write → read reproduces every construct this API emits under
`equalStructural`, in both formats, with geometry within ±0.01 cm.

**Opening a generated deck in the editor**, without a download: store the
model and navigate to it — `storePresentation`, not `storeDocument` (which is
typed for the text model):

```ts
import { storePresentation } from "./src/lib/documentStore";

const model = deck.toDeck();
storePresentation(model);
// router.push(`/slides/${model.id}`) — the route's guard reads the stored kind
```

## Validation

Every method — fluent, JSON or Markdown options — validates its input before
anything is written: unknown properties, wrong types, out-of-range numbers,
malformed lengths and colours, undecodable images and a per-slide call with
no slide all throw a `DocumentGenerationError` whose message names the path
(`slides[1].image.contentType: expected one of …`). Slide *Markdown* is
different by design: text outside the grammar is not an error but a
[drop record](#drop-records).

## Examples

| File | Demonstrates |
| --- | --- |
| `examples/deck.js` | a data-driven quarterly review: title, highlights with notes, one slide per region from the numbers, a text-box scorecard on a coloured background; writes `.pptx`, `.odp` and `.html` |
| `examples/slides-from-markdown.js` | Markdown in, deck out: parses a file (`node examples/slides-from-markdown.js deck.md`) or a built-in sample, prints what was dropped, writes the deck in its front matter's format |

```sh
node examples/deck.js
node examples/slides-from-markdown.js
```

Output lands in `examples/output/`. Like the document examples, each
`examples/<name>.js` is a small bootstrap running `examples/_impl/<name>.ts`
through `examples/lib/runTs.js`.
