---
id: SPEC-145
title: Convert between products — documents, decks and workbooks
status: draft
area: platform
complexity: normal
story_points: 8
depends_on: [SPEC-142, SPEC-018, SPEC-019, SPEC-022, SPEC-024, SPEC-039, SPEC-045, SPEC-050, SPEC-055, SPEC-091, SPEC-106, SPEC-118, SPEC-121, SPEC-126, SPEC-140]
---
## Why
Some of the most common conversions a person asks a shell for cross the line
between this product's products: the outline in `plan.docx` as a first-draft
deck, the deck someone presented as a handout document, a workbook's numbers
as a table in a report, the tables in a Word file as a spreadsheet, a
Markdown talk as `.pptx`. The app cannot do any of these today, and the
products are not allowed to import each other — so no product can.

This is what `src/lib/convert/` is for. Like `src/lib/generate/`, it is a
composition point above the products, and the generation API it may call is
already the validated way to *build* a deck or a document from content.
Every mapping here is a pure function from one product's model to another's,
run in Node, followed by SPEC-141's `writeAs` — so no conversion here writes a
byte of any format itself.

Every one of these conversions loses something, by nature: a deck has no
footnotes, a document has no formulas, a spreadsheet has no paragraphs. So
each mapping is pinned exactly — what becomes what, and what is dropped and
reported — rather than left to "best effort", and every drop reaches
SPEC-141's report, so `--strict` refuses a lossy one. **This file is the
normative source for the four cross-product mappings, the Markdown-to-deck
rule, and their `convert` options.**

## Acceptance criteria
- [ ] Node-tier `CONVERSIONS` rows: text → `pptx`, `odp`; presentation → `docx`, `odt`, and `md` (slide Markdown); workbook → `docx`, `odt`, `md`; text → `xlsx`, `ods`, `csv`, `tsv`; presentation ↔ workbook is not offered, and asking for it names `docx` as the bridge (`can't convert a presentation to .xlsx directly — convert it to .docx first`); each row goes `openBytes` → a mapping in `src/lib/convert/crossProduct/` → `writeAs`, and the mapping's drops and notices are the report
- [ ] Text → deck per "Document to deck": tracked changes accepted on a copy (`acceptAllChanges`, SPEC-039), slides started at the slide level (`--slide-level <1–9>`, default the shallowest heading level in the document), shallower headings becoming section title slides, deeper headings, paragraphs and list items becoming bullet rows at the pinned levels with their bold, italic, underline, strikethrough and colour (through `RunSpec`) and external links (set as `SlideRun.linkHref` on the built deck per "Links and images"), more than `MAX_ROWS_PER_SLIDE` (8, new) rows continuing on `<title> (cont.)` slides, each image on its own image slide, and the opening title slide from the Title/Subtitle paragraphs or the document title — built through SPEC-106's `PresentationBuilder` (`addTitleSlide`, `addBulletSlide`, `addImageSlide`), so a generated deck passes its validation
- [ ] Deck → text per "Deck to document": per slide (hidden ones too, unless `--skip-hidden`) a Heading 1 from `slideOutline(slide).title` or `Slide <n>`; text frames' paragraphs in frame order as paragraphs or list items at their indent level with their run formatting and links; picture frames as inline images no wider than the text column (PNG, JPEG and GIF only — the content types SPEC-045's image block accepts; any other is reported); SPEC-091 table frames as tables, one paragraph per cell holding the cell's paragraphs joined by a space; speaker notes after the slide's content as a bold `Speaker notes` paragraph and the notes in italics, unless `--no-notes`; `--page-per-slide` puts a page break before each slide after the first; charts, media, shapes' geometry, animations and transitions are reported by kind
- [ ] Deck → `md` is SPEC-106's `writeSlideMarkdown(model)` and its drop records, so `.md` → deck → `.md` keeps SPEC-106's `equalStructural` round trip; deck → `docx`/`odt` is the outline document above (a `.md` *document* of a deck is therefore never produced — documented, with the reason)
- [ ] Markdown to deck per "Markdown to deck": the decision is made on the input's raw text **before** it is opened — a `.md` input with a slide separator (a line that is exactly `---`, preceded by a blank line or the start of the body, outside YAML front matter and fenced code; a `---` directly under a text line is a setext heading underline, not a separator) is parsed with SPEC-106's `parseSlideMarkdownWithReport`; any other `.md` input goes through `openBytes` (SPEC-050's reader) and "Document to deck"; `--md-slides` / `--md-outline` force either, and each is a usage error for a non-Markdown input
- [ ] Workbook → text per "Workbook to document": the sheets SPEC-142's "Sheet selection" picks (the active sheet by default, `--sheet`, or every visible sheet with `--all-sheets`; `--range <A1:F20>` narrows a single selected sheet), each a Heading 1 named after the sheet when more than one is written, then a table of its used range with each cell's displayed text — the grid's text in a column wide enough to hold it: `formatValue(value, numberFormat, ctx).text` with the context under "Workbook to document" — bold and italic kept, the first row marked as a repeating header row; merged cells written unmerged with their value in the top-left cell; formula cells as their cached values with one notice; an empty sheet skipped with a notice; a sheet table past `MAX_TABLE_CELLS` (20,000, new) refused (exit 3) with `<sheet> has <n> cells — more than a document table can hold here (20,000). Narrow it with --sheet and --range, or convert to .csv.`
- [ ] Text → workbook per "Document to workbook": each top-level table becomes one sheet named by its SPEC-055 caption's user text (`captionUserText` of the paragraph `captionForObject` finds), else `Table <n>`, made valid and unique under SPEC-121's `validateSheetName` rules; each cell's plain text becomes a typed value through SPEC-126's `inferCsvField(text, locale, "1900")` — **never** a formula (a cell reading `=SUM(A1:A2)` stays text, as SPEC-140 rules for imported data) — with its number format, bold where every run of the cell is bold, merges from the table's spans, and column widths from its grid; nested tables and everything outside tables are reported; a document with no table is refused (exit 3, `<file> has no tables to put in a workbook`); a `csv`/`tsv` target writes `--table <n>` (default 1) and, when the document has more, the notice `only table 1 of <n> was written — choose another with --table`
- [ ] Option scoping: `--slide-level`, `--md-slides`, `--md-outline` apply to text → deck; `--skip-hidden`, `--no-notes`, `--page-per-slide` to deck → text; `--sheet`, `--all-sheets`, `--range`, `--locale` to workbook → text; `--table`, `--locale` to text → workbook; any of them given for another conversion is exit 1 for that input, naming the conversions it applies to
- [ ] Every mapping is pure, DOM-free and deterministic: it takes a model and options and returns `{ subject: ConvertSubject; report: ConversionReport }` without touching its input (asserted by deep-freezing the input in tests), uses no `Date`, `Math.random` or `crypto` beyond SPEC-141's pinned clock and the builders' own id minting, and two runs with the same `--modified` write `equalBytes` files
- [ ] No chaining into the render tier: `--to pdf`/`png`/`jpeg`/`webp` always render the input's own product (SPEC-143/144 — a `.docx` to `png` is its pages, never slides); `convert plan.docx --to pptx` followed by `convert plan.pptx --to pdf` is the documented way to a deck PDF of a document
- [ ] `docs/cli.md` gains "Between products": the mapping tables in prose, what is dropped per direction, the Markdown rule, and the options
- [ ] e2e (`e2e/SPEC-145.spec.ts`): `convert word-toc.docx --to pptx` opens with a slide titled with each of the document's Heading 1 texts, in document order, after the title slide; `convert spec091-tables.pptx --to docx` opens with a Heading 1 per slide and a table whose first cell matches the slide table's; `convert <tmp>/talk.md --to pptx` (three `---`-separated slides) has three slides; `convert spec118-formats.xlsx --to docx --sheet 1` opens with a table whose text, for three sampled cells including a currency-formatted one (cells the grid shows in full, not as `###`), equals what the app's grid shows for them; `convert table.docx --to xlsx` opens with a sheet `Table 1` whose `A1` equals the table's first cell text; `convert table.docx --to pptx --strict` is refused (exit 3) because the table is reported as dropped

## Out of scope
Presentation ↔ workbook directly. Tables, charts, equations, shapes, text
boxes, footnotes, headers and footers carried *into* a deck (the builder has
no table or chart method yet — each is a follow-up over SPEC-091/092 and a
builder extension). Charts and pictures carried from a workbook into a
document. Slide design: a deck converted from a document uses the blank
deck's default theme and layouts, never a chosen template. Formulas or
cell styles beyond bold and italic in either direction. Inferring slide breaks
from a document without headings — such a document is all front matter, so
its body becomes 8-row slides titled with the document title. HTML or plain text as a target. Mail merge.

## Reference

### Dependencies restated
- `specs/141-…`/`specs/142-…` — `CONVERSIONS`, `ConvertSubject` (`{ kind:
  "text"; model: DocumentModel } | { kind: "presentation"; model } | { kind:
  "spreadsheet"; model }`), `writeAs`, `ConversionReport { dropped: { kind;
  location; reason }[]; notices: string[] }`, `openBytes`, `--locale`
  (one of `SPREADSHEET_LOCALE_TAGS`), "Sheet selection" (`--sheet`,
  `--all-sheets`), output naming, `--strict`
  (exit 3 on any drop or notice), exit codes.
- Text model (SPEC-011 onward): sections of blocks; `Paragraph { runs;
  headingLevel?: number; list?: ListFormatting { type: "bullet" | "number";
  level } ; style }`; `Table` blocks with rows of cells of blocks (SPEC-022)
  and spans (SPEC-023); inline image runs (SPEC-024); the built-in Title and
  Subtitle paragraph styles (SPEC-019); captions (SPEC-055) attached to
  tables; `acceptAllChanges(model)` (`src/lib/trackChanges.ts`, SPEC-039);
  captions (SPEC-055) are Caption-styled paragraphs bound to the object
  directly after or before them: `captionForObject(model, address) →
  ParagraphAddress | null` and `captionUserText(paragraph)` (the text past
  the label and number) in `src/lib/captions.ts`; `headingAnchor(text,
  ordinal)` in `src/lib/links.ts` derives a heading's link anchor.
  `sanitizeHref(url): string | null` in `src/lib/core/urls.ts`.
- `specs/045-document-generation-api.md` — `DocumentBuilder`
  (`addHeading(content, level)`, `addParagraph`, `addList`, `addTable({ rows;
  header?; columnWidths? })`, `addImage({ data; contentType; width? })`,
  `addPageBreak`, `toDoc()`), `RunSpecObject { text; bold?; italic?;
  underline?; strikethrough?; fontFamily?; fontSizePt?; color?; highlight? }`
  — no link field, and unknown keys are rejected; its image block accepts
  `image/png`, `image/jpeg` and `image/gif` only; a table cell holds one
  paragraph. Where a construct needs list levels deeper than `addList`'s, or
  a link, the mapping sets it on the built model (`Paragraph.list.level` as
  SPEC-018/032 define it; `Run.linkHref` / `Run.linkAnchor`, SPEC-025).
- `specs/106-presentation-generation-api-and-markdown.md` —
  `PresentationBuilder` (`addTitleSlide(title, subtitle?)`,
  `addBulletSlide(title, bullets: BulletRow[], options?)`,
  `addImageSlide(title, image: SlideImageSpec)`, `setNotes(paragraphs)`,
  `toDeck()`), `BulletRow = string | { text?; runs?; level? } | BulletRow[]`
  (levels 0–8, clamped at `MAX_INDENT_LEVEL`), `SlideImageSpec { data;
  contentType: "image/png" | "image/jpeg" | "image/gif"; width?; height?;
  altText? }`; `parseSlideMarkdownWithReport(text, options?) → { deck;
  dropped: DeckDropRecord[] }`; `writeSlideMarkdown(model) → { text; dropped
  }`; the `.md` → deck → `.md` round trip under `equalStructural`.
- Presentation model (SPEC-079 onward): `Slide { frames; notes?; hidden? }`;
  `slideOutline(slide): { title?; lines }` (`presentation/slideModel.ts`);
  text frames of `SlideParagraph`s with an indent level, a bullet or number
  marker and `SlideRun`s (`linkHref`, always `sanitizeHref`-clean, or
  `linkSlide`, SPEC-097); picture frames referencing a media part; SPEC-091
  table frames; SPEC-092 chart frames; SPEC-095 media frames.
- Workbook model (SPEC-116 onward): sheets with `hidden`, sparse rows of
  cells (`value`, `formula?`, style index), `merges`, `ColumnSpan.widthPx`,
  `dateSystem`; `formatValue(value, code, ctx: FormatContext):
  FormattedValue` (`spreadsheet/numberFormat/render.ts`, SPEC-118) with
  `FormatContext { locale: SpreadsheetLocale; dateSystem; widthPx?;
  digitWidthPx?; dateQuirks?: "excel" | "proleptic"; palette? }` — the grid
  passes the cell's width, which can turn a number into `###`;
  `inferCsvField(field, locale, dateSystem) → { value; numberFormat? }`
  (SPEC-126); `validateSheetName(model, name)` (SPEC-121: 1–31 characters,
  none of `: \ / ? * [ ]`, no leading/trailing `'`, not "History", unique
  ignoring case); `WorkbookBuilder` (SPEC-140) — which the text → workbook
  mapping uses to build, gaining its validation.
- Equality: `equalStructural` (`src/lib/equality.ts`) for the slide Markdown
  round trip; field-level expectations on built models elsewhere;
  `equalBytes` (`src/lib/core/equality.ts`) for determinism.

### Document to deck
| Source | Becomes |
| --- | --- |
| Front matter (blocks before the first heading at or above the slide level) | A title slide: title = first front-matter paragraph in the Title style, else the document title; subtitle = first in the Subtitle style. Other non-empty front-matter blocks → a bullet slide titled with the deck title |
| Heading shallower than the slide level | A title-only section slide |
| Heading at the slide level | A new bullet slide, titled with its text |
| Heading deeper than the slide level | A bullet row at level 0, every run bold |
| Paragraph (non-empty) | A bullet row at level 0 |
| List item at list level `k` (1-based) | A bullet row at level `k − 1` (clamped to 8), numbered lists as `{ style: "number" }` slides only when every row of that slide is numbered |
| Inline image (PNG, JPEG, GIF) | An image slide titled like the slide it falls in, after that slide; any other content type is reported |
| Run formatting | bold, italic, underline, strikethrough, colour kept; links per "Links and images"; font, size and highlight dropped silently |
| Fields | their displayed text |
| Tables, equations, drawings, footnotes/endnotes, comments, headers/footers, TOC | dropped and reported, one record per item, `location` "paragraph <n>" / "table <n>" |
| Page and section breaks, empty paragraphs | ignored |

More than 8 rows on a slide continue on a slide titled `<title> (cont.)`,
splitting between rows (never inside one). A document with no heading at or
above the slide level is all front matter.

### Deck to document
| Source | Becomes |
| --- | --- |
| Deck title | Document title |
| Slide | Heading 1: outline title, else `Slide <n>` (n is 1-based among all slides) |
| Text frame paragraph, level 0, no marker | Paragraph |
| Paragraph with a bullet/number marker or level > 0 | List item, `type` from the marker, `level` = indent level + 1 |
| Picture frame | Inline image paragraph, frame width capped at the section's text width, aspect kept, alt text kept |
| Table frame | Table; each cell one paragraph, its paragraphs joined by a space |
| Shape with text | Its paragraphs, as text frames are; geometry reported once per shape |
| Notes | A paragraph `Speaker notes` (bold) then each notes paragraph in italics |
| Chart, media, animation, transition | Reported by kind; nothing written |

### Markdown to deck
Run on the decoded input text (UTF-8, BOM removed), before `openBytes`:
1. Strip a leading YAML front matter block (`---` … `---`) from consideration.
2. Ignore lines inside fenced code (```` ``` ```` or `~~~` fences).
3. Any remaining line that is exactly `---` (no other characters, any line
   ending) whose previous line is blank — or which is the first line after
   step 1 — → slide Markdown; otherwise → outline (SPEC-050 reader, then
   "Document to deck"). A `---` right under a text line is SPEC-050's setext
   heading and never counts.

### Links and images
- Document → deck: a run's `linkHref` that `sanitizeHref` accepts becomes the
  built `SlideRun.linkHref`; one it rejects, and every `linkAnchor` (decks
  cannot target a heading), is dropped and reported.
- Deck → document: `SlideRun.linkHref` becomes `Run.linkHref`; `linkSlide: n`
  becomes `Run.linkAnchor = headingAnchor(<slide n's heading text>, n - 1)` —
  each slide is exactly one Heading 1 and the converted document has no other
  headings, so slide n's heading is ordinal n − 1.
- Images in either direction are carried only as `image/png`, `image/jpeg` or
  `image/gif` (what both builders accept); SVG, WebP, EMF, TIFF and anything
  else is dropped and reported with its location.

### Constants
`MAX_ROWS_PER_SLIDE = 8` and `MAX_TABLE_CELLS = 20_000` are new, exported
from `src/lib/convert/crossProduct/limits.ts`.

### Workbook to document
Display text uses `formatValue(value, numberFormat, { locale:
SPREADSHEET_LOCALES[<--locale tag>], dateSystem: model.dateSystem,
dateQuirks: model.format === "ods" ? "proleptic" : "excel" }).text` — no
`widthPx`, so a value is never shown as `###` — and an error cell its error
code. The used range is the bounding box of stored cells with a value or a
non-default style, from SPEC-116's sparse rows. Column widths become the
table's grid widths (px → cm at 96 dpi), scaled down proportionally when
their sum exceeds the text width. Hidden rows and columns are skipped and reported once per
sheet. Charts, images, shapes, comments, validation and conditional formats
are reported by kind.

### Document to workbook
Sheets appear in document order of their tables. Row `r`, column `c` of the
table's grid is cell `(r, c)` from `A1`; a horizontally or vertically spanned
cell becomes a merge over the cells it covers, its text in the top-left.
Column widths: the table's grid widths in cm → `{ px }` at 96 dpi. A cell
whose text is empty stores no value. The workbook's title is the document
title, its format the target's, its date system 1900.

### Tests
- `docToDeck.spec.ts` — every row of its table against a hand-built model,
  slide-level override, continuation splitting at 8, the title-slide rules,
  report records, input untouched.
- `deckToDoc.spec.ts` — every row against `mini-deck.pptx`,
  `spec091-tables.pptx` and `spec097-text-depth.pptx`, `--skip-hidden`
  with `spec104-hidden.pptx`, `--no-notes`, `--page-per-slide`.
- `markdownToDeck.spec.ts` — the three-step rule (front matter `---`, fenced
  `---`, a real separator), both overrides, the `equalStructural` round trip.
- `sheetsToDoc.spec.ts` — display text for currency, percent, date and error
  cells in `en-US` and `de-DE`, merges, hidden rows, empty sheets, the cell
  ceiling message, `--range`.
- `docToSheets.spec.ts` — caption and default names, name collisions and
  invalid characters, typed values, the `=` text rule, merges, widths, the
  no-tables refusal, `--table` and the csv notice.
- `crossProductCli.spec.ts` (in-memory `CliIo`) — every row reachable,
  option scoping errors, the presentation ↔ workbook refusal text,
  determinism; `e2e/SPEC-145.spec.ts`.

### Specs this changes
- `specs/141-…` — the registry gains the rows above and `docs/cli.md` a
  section; `writeAs` is unchanged (mappings hand it a model of the target's
  product).
- `specs/045-…`, `specs/106-…`, `specs/140-…` — their builders gain a caller;
  no method changes.
- `specs/000-roadmap.md` — Phase 19 row.
