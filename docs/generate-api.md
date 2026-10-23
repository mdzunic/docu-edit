# Document generation API (SPEC-045)

A fluent JavaScript/TypeScript API and a JSON document-description format
for generating `.docx`/`.odt` files from data — no editor, no UI. It is a
front end over the product's own canonical document model and delegates to
the existing writers (`src/lib/docx/writer.ts`, `src/lib/odf/writer.ts`); it
adds no new file-format code of its own.

The module lives at `src/lib/generate/` and is imported the same way any
other `src/lib` module is:

```ts
import { DocumentBuilder } from "../lib/generate";
```

Runnable, self-contained examples live in `examples/` — see [Examples](#examples)
below. Decks (`.pptx`/`.odp`) have their own builder in the same module — see
[Presentations](#presentations).

## Quick start

```ts
import { DocumentBuilder } from "./src/lib/generate";

const builder = new DocumentBuilder({ title: "Invoice #1001" });

builder
  .addHeading("Invoice", 1)
  .addParagraph([{ text: "Bill to: ", bold: true }, "Acme Corp"])
  .addTable({
    header: true,
    rows: [
      ["Item", "Qty", "Price"],
      ["Widget", "3", "$9.00"],
    ],
  })
  .addPageBreak()
  .addTableOfContents();

const buffer = await builder.toBuffer("docx"); // Node
const blob = await builder.toBlob("docx"); // browser
```

Every `add*`/`set*` method validates its arguments immediately and throws a
`DocumentGenerationError` naming the offending value's path (e.g.
`"body[3].rows[1]: expected 2 cells"`) — nothing is written until
`toBytes`/`toBlob`/`toBuffer` is called, and invalid input never produces a
file.

## `DocumentBuilder`

`new DocumentBuilder(options?: { title?: string })` creates an empty
document (one implicit section, matching the common single-section case —
see [Sections](#sections-and-page-setup) below). Every `add*` method appends
to that section's body and returns `this`, so calls chain.

### `addHeading(content, level, options?)`

Appends a heading paragraph.

- `content`: a plain string, or an array of [run specs](#run-specs) for
  formatted heading text.
- `level`: an integer 1–6.
- `options.alignment?`: `"left" | "center" | "right" | "justify"`.

The heading's appearance (bold, size, spacing) is baked from the product's
built-in `Heading{level}` style, the same look applying that style in the
editor produces — it never references a style by id, so a generated document
never needs a style catalogue of its own.

```ts
builder.addHeading("Chapter One", 1);
builder.addHeading([{ text: "Section ", }, { text: "1.1", italic: true }], 2);
```

### `addParagraph(content, options?)`

Appends a paragraph.

- `content`: a plain string, or an array of [run specs](#run-specs).
- `options.alignment?`: `"left" | "center" | "right" | "justify"`.
- `options.spaceBeforePt?` / `options.spaceAfterPt?`: non-negative points.

```ts
builder.addParagraph("Plain text.");
builder.addParagraph([{ text: "Warning: ", bold: true, color: "#cc0000" }, "read carefully."]);
```

### `addList(items, options?)`

Appends one paragraph per item, all at the same (single) level — nested
lists are out of scope.

- `items`: an array, each entry a plain string or an array of
  [run specs](#run-specs).
- `options.style?`: `"bullet"` (default) or `"number"`.
- `options.startAt?`: an integer 1–9999, the number the list counts from.

```ts
builder.addList(["First", "Second", "Third"]);
builder.addList(["Fourth", "Fifth"], { style: "number", startAt: 4 });
```

**Adjacency note:** two `addList` calls of a *different* `style` placed
directly next to each other (nothing between them) are read back as **one**
list using the first call's style — this is how the product's own
multi-level list numbering plan works (a contiguous run of list paragraphs
shares one definition per level; see `src/lib/lists.ts`'s `planListLevels`),
not a bug in this API. Put a paragraph (or anything else) between two lists
of different style if they should stay visually distinct.

### `addTable(spec)`

Appends a table.

- `spec.rows`: a non-empty array of rows; each row a non-empty array of
  cells, all rows the same length. A cell is a plain string, an array of
  [run specs](#run-specs), or `{ text }` / `{ runs }`.
- `spec.columnWidths?`: an array of [lengths](#lengths), one per column;
  defaults to equal widths spanning the page's text area.
- `spec.header?`: `true` marks the first row as a repeating header row.
- `spec.alignment?`: `"left" | "center" | "right"`.
- `spec.widthMode?`: `"auto" | "fixed" | "percent"` (default `"fixed"`).
- `spec.widthPercent?`: 0–100, meaningful only when `widthMode` is
  `"percent"`.

```ts
builder.addTable({
  header: true,
  rows: [
    ["Item", "Qty", "Price"],
    ["Widget", "3", "$9.00"],
    [{ runs: [{ text: "Total", bold: true }] }, "", "$9.00"],
  ],
  columnWidths: [{ cm: 8 }, { cm: 3 }, { cm: 4 }],
});
```

### `addImage(spec)`

Appends an inline image.

- `spec.data`: a base64 string, or (Node/browser JS only, not the JSON
  format) a `Uint8Array` / `ArrayBuffer` / `Buffer`.
- `spec.contentType`: `"image/png"`, `"image/jpeg"` or `"image/gif"` — SVG
  is not supported (the product converts SVG to PNG on insert via a
  browser-only canvas path this API deliberately avoids, to stay DOM-free
  in Node — see `src/lib/svgToPng.ts`).
- `spec.width?` / `spec.height?`: [lengths](#lengths); default to the
  image's natural size scaled down to fit the page's text area, aspect
  ratio preserved.
- `spec.altText?`: a string.
- `spec.wrap?`: `"inline"` (default) `| "square" | "behind" | "infront"`.

The image's natural pixel dimensions are decoded from its own bytes (no
`<img>`/canvas decode), so this works identically in Node and the browser.

```ts
import { readFile } from "node:fs/promises";
const logo = await readFile("./logo.png");
builder.addImage({ data: logo, contentType: "image/png", width: { cm: 4 }, altText: "Company logo" });
```

### `addPageBreak()`

Appends a manual page break. Takes no arguments.

### `addTableOfContents(options?)`

Appends a table of contents. Its entries are resolved from the document's
headings **at output time** (`toDoc`/`toBytes`/`toBlob`/`toBuffer`), so a
TOC can be added before headings that come later in the body.

- `options.levels?`: an array of heading levels (1–6) to include; default
  `[1, 2, 3]`.
- `options.showPageNumbers?` / `options.hyperlinksOnly?`: booleans; default
  `true` / `false`.

**Page numbers note:** this API has no live pagination, so every entry's
page number is always unresolved (`null`) — opening the file in Word or
LibreOffice and updating the TOC (or opening it in this app, which paginates
live) is what actually fills page numbers in. Both readers infer
`showPageNumbers`/`hyperlinksOnly` from whether any entry actually has a
cached page number, so a TOC this API writes with `showPageNumbers: true`
reads back as `hyperlinksOnly: true` once reopened — pass
`{ hyperlinksOnly: true, showPageNumbers: false }` if round-tripping through
this app's own reader matters for your use case.

### `setHeader(content, kind?)` / `setFooter(content, kind?)`

Sets the section's header/footer content, replacing any previous content
for that `kind`.

- `content`: the same shorthand as a paragraph (a string or an array of
  [run specs](#run-specs)) for a single line, or an array of block-spec
  objects (`{ type: "paragraph" | "heading" | "list" | "table" | "image", ... }`,
  the same shapes `body` accepts in the [JSON format](#json-document-description-format))
  for anything richer. Page breaks and TOCs are not valid here — see the
  [JSON format reference](#header--footer-content) for why.
- `kind?`: `"default"` (every page, the default) `| "first" | "even"`.

```ts
builder.setHeader("Acme Corp — Confidential");
builder.setFooter([{ type: "paragraph", text: "Page footer", alignment: "center" }]);
```

### Output

- `toDoc(): DocumentModel` — the assembled canonical model, with every TOC's
  entries resolved. Useful for tests, or for handing straight to this app's
  own `storeDocument`/router, skipping serialisation entirely.
- `toBytes(format, options?): Promise<Uint8Array>` — `format` is `"docx"` or
  `"odt"`; `options.modified?: Date` overrides the written timestamp
  (defaults to now).
- `toBlob(format, options?): Promise<Blob>` — works in the browser and in
  Node (Node has had a global `Blob` since v18).
- `toBuffer(format, options?): Promise<Buffer>` — Node only; throws a clear
  `DocumentGenerationError` if `Buffer` is not global (e.g. an unbundled
  browser script).

### `DocumentBuilder.fromDescription(description)`

A static factory that interprets a [JSON document description](#json-document-description-format)
into a builder ready for output — see below.

### Sections and page setup

This API always produces a document with **one** section (the common case
every one of its examples needs). Page size, margins, columns and per-
section headers/footers are section-level constructs this API does not
expose; use the editor's own page-setup UI (SPEC-027) after opening a
generated document if you need them, or extend the model directly via
`toDoc()`.

## Lengths

Every width/height this API accepts (table column widths, image size) is a
"length helper": an object with **exactly one** of four unit keys, converted
to an integer Emu (English Metric Units, the canonical model's own length
unit — 914400 per inch) immediately on entry.

```ts
{ cm: 2.54 }
{ pt: 12 }
{ in: 1 }
{ emu: 914400 }
```

`cm`/`pt`/`in` may be any non-negative number (rounded to the nearest whole
Emu); `emu` must itself be a non-negative **integer** — it is already the
target unit, so there is nothing to round. A negative value, a non-integer
`emu`, or an object with zero or more than one unit key is a validation
error.

## Run specs

A "run" is a span of text with its own formatting — the unit every block
type that carries text (heading, paragraph, list item, table cell) accepts
as an array. Each entry is either a plain string (unformatted) or an object:

```ts
{
  text: string,
  bold?: boolean,
  italic?: boolean,
  underline?: boolean,
  strikethrough?: boolean,
  fontFamily?: string,
  fontSizePt?: number,   // > 0
  color?: string,        // "#rrggbb"
  highlight?: string,    // "#rrggbb"
}
```

A block that carries text also accepts the plain-string shorthand directly
(equivalent to a single unformatted run) instead of an array.

## JSON document-description format

A plain-JSON object mapping **1:1** onto the builder API above — every
block type in `body` corresponds to exactly one `add*` method, with the
same fields. Its schema is published at
[`schema/document-description.json`](../schema/document-description.json)
as a JSON Schema document for external tooling; this section is the
narrative reference.

```json
{
  "title": "Invoice #1001",
  "body": [
    { "type": "heading", "text": "Invoice", "level": 1 },
    { "type": "paragraph", "runs": [{ "text": "Bill to: ", "bold": true }, "Acme Corp"] },
    { "type": "list", "items": ["First", "Second"], "style": "bullet" },
    {
      "type": "table",
      "header": true,
      "rows": [["Item", "Qty", "Price"], ["Widget", "3", "$9.00"]]
    },
    { "type": "image", "data": "<base64>", "contentType": "image/png", "width": { "cm": 4 } },
    { "type": "pageBreak" },
    { "type": "toc", "levels": [1, 2], "hyperlinksOnly": true, "showPageNumbers": false }
  ],
  "header": "Acme Corp — Confidential",
  "footer": [{ "type": "paragraph", "text": "Page 1", "alignment": "center" }]
}
```

Generate directly from a description:

```ts
import { DocumentBuilder } from "./src/lib/generate";

const builder = DocumentBuilder.fromDescription(description);
const buffer = await builder.toBuffer("docx");
```

### Top-level fields

| Field | Type | Notes |
| --- | --- | --- |
| `title` | string | optional; defaults to "Untitled document" |
| `body` | array of [block](#block-types) | optional; omitted or `[]` describes an empty document |
| `header` | [header/footer content](#header--footer-content) | optional |
| `footer` | [header/footer content](#header--footer-content) | optional |

Any other top-level key is a validation error.

### Block types

Every block object has a `type` field plus the fields listed for that type
below — these are exactly the fields the corresponding builder method
takes, flattened into one object with `type` added. Any field not listed
for a given `type` is a validation error.

| `type` | Fields | Corresponds to |
| --- | --- | --- |
| `"heading"` | `text` or `runs`, `level` (required), `alignment?` | `addHeading` |
| `"paragraph"` | `text` or `runs`, `alignment?`, `spaceBeforePt?`, `spaceAfterPt?` | `addParagraph` |
| `"list"` | `items` (required), `style?`, `startAt?` | `addList` |
| `"table"` | `rows` (required), `columnWidths?`, `header?`, `alignment?`, `widthMode?`, `widthPercent?` | `addTable` |
| `"image"` | `data` (required, base64), `contentType` (required), `width?`, `height?`, `altText?`, `wrap?` | `addImage` |
| `"pageBreak"` | none | `addPageBreak` |
| `"toc"` | `levels?`, `showPageNumbers?`, `hyperlinksOnly?` | `addTableOfContents` |

A block's `text`/`runs` fields are mutually exclusive — specify at most one.

### Header / footer content

`header`/`footer` accept the same paragraph-content shorthand as a
paragraph (a plain string or an array of [run specs](#run-specs)) for a
single line, or an array of block-spec objects for anything richer — the
same shapes as `body` above, **except** `"pageBreak"` and `"toc"`: both are
document/section-level constructs the underlying writers already ignore
inside a header/footer part, so this API rejects them there with a clear
error instead of silently emitting nothing.

## Validation

Every method — fluent or JSON — validates its input before anything is
written: unrecognised properties, wrong types, out-of-range numbers and
malformed lengths all throw a `DocumentGenerationError` whose message names
the offending path (e.g. `body[2].rows[1]: expected 2 cells`) and the
expected type/constraint. No partial or invalid file is ever produced.

## Reading generated output

`readDocx(bytes, title?)` / `readOdt(bytes, title?)` parse a `Uint8Array`/
`ArrayBuffer`/`Buffer` straight into the canonical model — no `File`, no
disk, and no download round trip — using the exact same reader
(`parseDocx`/`parseOdt`) and model-assembly logic
(`src/lib/openDocument.ts`'s `assembleDocumentModel`) the app's own "Open
document" picker uses.

```ts
import { readDocx } from "./src/lib/generate";
const model = await readDocx(await builder.toBuffer("docx"));
```

**Node note:** the existing readers parse XML with the browser's
`DOMParser`/`XMLSerializer`/`Node`, which a genuine (non-jsdom) Node.js
process does not have globally. `readDocx`/`readOdt` detect this and lazily
borrow them from the project's own `jsdom` test dependency — already
installed, nothing new to add — only when running outside a browser; a real
browser never takes that branch. Writing (`toBytes`/`toBlob`/`toBuffer`)
never needs this: a freshly-generated document has no preserved source XML
to parse, so the write path has no DOM dependency at all, in Node or the
browser (SPEC-045 AC-3).

**Opening a generated Blob in the editor** (SPEC-002), without a download
round trip: wrap the Blob in a `File` and feed it through the same functions
the app's own file picker uses —

```ts
import { openDocumentFile } from "./src/lib/openDocument";
import { storeDocument } from "./src/lib/documentStore";

const blob = await builder.toBlob("docx");
const file = new File([blob], "Report.docx", { type: blob.type });
const model = await openDocumentFile(file);
storeDocument(model);
// router.push(`/doc/${model.id}`)
```

## Presentations

The same module generates decks (SPEC-106): `PresentationBuilder` builds a
`.pptx`/`.odp` slide by slide —
`addTitleSlide`/`addBulletSlide`/`addImageSlide`/`addBlankSlide`, the
per-slide `addBullets`/`addTextBox`/`addImage`/`setNotes`/`setBackground`, and
`toDeck`/`toBytes`/`toBlob`/`toBuffer`/`fileName`/`toHtml` — from the same
lengths and run specs as `DocumentBuilder`, delegating every byte to the
presentation writers. A JSON deck-description format
(`deckFromDescription`, `schema/deck-description.json`), a slide-Markdown
grammar in both directions (`parseSlideMarkdown`/`writeSlideMarkdown`), HTML
export (`exportDeckHtml`) and byte readers (`readPptx`/`readOdp`) complete it.

```ts
import { PresentationBuilder } from "./src/lib/generate";

const deck = new PresentationBuilder({ title: "Q3 review" })
  .addTitleSlide("Q3 review", "Prepared 30 September")
  .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%"]]);
const bytes = await deck.toBytes(); // .pptx
```

The full reference — every method, the deck description and the Markdown
grammar — is [generate-deck-api.md](generate-deck-api.md), a sibling file
because this one would otherwise pass 800 lines. The two builders never
import each other: `src/lib/generate/index.ts` re-exports both, and a
text-only script can keep importing `./builder` alone.

## Examples

Five complete, runnable programs live in `examples/`:

| File | Demonstrates |
| --- | --- |
| `examples/invoice.js` | a table, mixed run formatting |
| `examples/report.js` | multiple headings, page breaks, a table of contents |
| `examples/letter.js` | an inline image (letterhead) |
| `examples/deck.js` | a data-driven quarterly review deck (SPEC-106) |
| `examples/slides-from-markdown.js` | slide Markdown in, a deck out (SPEC-106) |

Each is executable directly:

```sh
node examples/invoice.js
node examples/report.js
node examples/letter.js
node examples/deck.js
node examples/slides-from-markdown.js
```

Every example writes its output to `examples/output/`. Each
`examples/<name>.js` is a small bootstrap that runs the actual example logic
(`examples/_impl/<name>.ts`) through Vite's own module runner (`runnerImport`
from the `vite` package — already a real project dependency), because the
generation API is TypeScript like the rest of `src/lib/`, and plain Node
cannot import a `.ts` file without flags this command line intentionally
never passes. See `examples/lib/runTs.js` for the short bridge.
