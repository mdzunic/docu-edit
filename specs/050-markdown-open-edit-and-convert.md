---
id: SPEC-050
title: Markdown — open, edit visually, and convert to and from .docx/.odt
status: draft
area: io
depends_on: [SPEC-008, SPEC-009, SPEC-019]
---
## Why
Markdown is where a lot of writing starts and almost none of it finishes.
People draft in it, then need a document — for a client, a committee, a
printer — and the trip from `.md` to `.docx` currently runs through some other
tool. Going the other way is worse: a `.docx` someone sends you is a dead end
for anyone whose pipeline is text files and version control.

This makes `.md` a first-class format alongside `.docx` and `.odt`: open one
and it renders on the page canvas as a document, with headings that are real
Heading styles and lists that are real lists — not a syntax-highlighted source
pane. Save a document as `.md` and get clean, readable markdown back.

It is queued last deliberately. Markdown is the lossiest target the product
has, and the honest list of what a conversion drops can only be written once
the features that get dropped exist. Running this before comments, footnotes,
columns and the table of contents would mean shipping a conversion report that
goes quietly out of date every time a later spec lands.

## Acceptance criteria
- [ ] Opening a `.md` or `.markdown` file — from the library, the file picker, or a drop — renders it on the page canvas as an editable document, with no raw-source view anywhere in the product
- [ ] The supported markdown subset is exactly the one named in "Supported subset" below; the parser is hand-written and no markdown dependency is added
- [ ] Markdown constructs outside that subset (raw HTML blocks, reference definitions, anything unrecognised) are retained as opaque blocks anchored in place and re-emitted verbatim on a `.md` save, per the SPEC-008 rule
- [ ] YAML front matter is retained verbatim and re-emitted at the top of the file on a `.md` save, unparsed
- [ ] A markdown document opens with the default page setup — markdown carries no page geometry, and the canvas needs some; the applied defaults are never written back to the `.md`
- [ ] Saving any open document as `.md` produces the `WriteResult` shape SPEC-008 defines, with `dropped` naming everything markdown cannot carry, and raises the same conversion report dialog
- [ ] Opening a `.md`, editing nothing, and saving it back as `.md` satisfies `equalStructural(openedModel, reopenedModel)` — the serialiser normalises syntax, so `equalBytes` on the file is explicitly *not* asserted
- [ ] A `.docx` → `.md` → `.docx` round trip satisfies `equalSemantic` against the source model masked by `dropped`, using SPEC-008's `withoutDropped(source, dropped)` transform and never a hand-rolled subset comparison
- [ ] Saving a `.docx`-backed document as `.md` does not change what that document is: it stays a `.docx` document in the library, keeps its preserved package bytes, and a later save as `.docx` still round-trips them intact
- [ ] Format detection admits markdown by file extension, documented as the deliberate exception to the content-sniffing rule in "Detection" below
- [ ] The unsupported-file-type message names all three formats, replacing the string SPEC-009 asserts — see "Specs this changes"
- [ ] A test opens a `.md` fixture exercising every construct in the supported subset, saves as `.docx`, reopens, and asserts `equalSemantic` against the masked source

## Out of scope
CommonMark or GFM conformance — the subset below is the contract, and a
construct outside it is preserved, not interpreted. A raw markdown source view
or a split editor; the whole point is the visual one. Editing preserved raw
HTML. Markdown flavours as a user setting. Pandoc-style extensions. Converting
markdown *to* `.pdf` directly — that path is `.md` → document → SPEC-034.
Front-matter interpretation, including mapping it to document properties; it is
carried, not read. Two-way sync with a file on disk.

## Reference

### Supported subset
Everything else is preserved opaquely rather than parsed.

| Markdown | Document model |
| --- | --- |
| ATX headings `#`–`######` | Heading 1–6 paragraph styles (SPEC-019) |
| Paragraph, blank-line separated | Paragraph |
| Two-space or backslash line end | LineBreak |
| `**bold**`, `*italic*`, `~~strike~~` | Run formatting (SPEC-012) |
| `` `code` `` | Run with the Code character style |
| Fenced code block, with info string | Paragraphs in the CodeBlock style; info string kept as the style's language |
| `-`/`*`/`+` and `1.` lists, nested | Bulleted and numbered lists (SPEC-018, SPEC-032) |
| `> ` blockquote | Blockquote paragraph style |
| `---` thematic break | Horizontal-rule paragraph |
| `[text](url)` | Hyperlink run (SPEC-025) |
| `![alt](src)` | Inline image (SPEC-024), `src` resolved relative to the file |
| GFM pipe table | Table without merged cells (SPEC-022) |

### What markdown cannot carry
Each of these produces a `DropRecord` on a `.md` save and appears in the
conversion report. This list is the exclusion list the cross-format test masks
with — it is committed alongside the test, per SPEC-008.

Page setup and section breaks · headers and footers · page numbers · columns ·
footnotes and endnotes · comments · the table of contents (emitted as its
rendered text, not as a field) · bookmarks · fonts, sizes, colours and
highlight · alignment, indents and spacing · merged table cells and table
styling · preserved `.docx`/`.odt` opaque parts and nodes · tracked changes.

`DropRecord['kind']` as SPEC-008 defines it is a closed union of package-level
kinds. This spec widens it with the content-level kinds above; the union stays
the single source both the report and the test read.

### Detection
`src/lib/openDocument.ts` detects format from content, never the extension —
an ODF package is recognised by its `mimetype` entry, and anything else that is
a ZIP takes the `.docx` path. Markdown has no magic bytes and no container, so
that rule cannot decide it. The deliberate exception:

> A file that is not a ZIP package is markdown if its name ends `.md` or
> `.markdown`, and unsupported otherwise. ZIP packages continue to be
> identified by content exactly as they are today.

This keeps content sniffing authoritative wherever it can work and confines
extension-trust to the one format that admits nothing else.

### Existing I/O surface
```ts
// src/lib/openDocument.ts   openDocumentFile(file): Promise<DocumentModel>
// src/lib/docx/reader.ts    parseDocx(bytes): Promise<ParsedDocxDocument>   — SPEC-004
// src/lib/odf/reader.ts     parseOdt(bytes):  Promise<ParsedOdtDocument>    — SPEC-005
// src/lib/opaque.ts         WriteResult = { bytes: Uint8Array; dropped: DropRecord[] }  — SPEC-008
```
This spec adds `src/lib/md/reader.ts` and `src/lib/md/writer.ts` alongside the
existing two pairs, matching their signatures, and extends the model's format
tag — today `"docx" | "odt"`, passed to `buildModel(name, format, parsed)` —
with `"md"`. Per `CLAUDE.md`, both live in `src/lib/` as plain TypeScript with
no Vue imports and are unit-testable without mounting a component.

### Equality
Defined normatively in `specs/002-create-blank-document.md` ("Equality"),
implemented in `src/lib/equality.ts`. Two of the three relations apply:
- **`equalStructural`** — for the `.md` → `.md` no-edit round trip, on the
  models. Not on the bytes: the serialiser emits one canonical spelling per
  construct, so a file using `*italic*` and setext headings comes back as
  `_italic_` and ATX headings with the same meaning. Asserting `equalBytes`
  here would be asserting that we echo input we deliberately normalise.
- **`equalSemantic`** — for every cross-format save, paired with the writer's
  `dropped` list exactly as SPEC-008 specifies.

`equalBytes` is used nowhere in this spec. Markdown has no opaque *parts*, only
opaque blocks inside one text file, and those are compared as text.

### Specs this changes
- `specs/009-drag-drop-and-open-from-url.md` asserts the literal error
  `"We can only open .docx and .odt files"`. That string is also pinned in
  `e2e/SPEC-009.spec.ts` and `src/components/AppShell.spec.ts`. All three change
  together to name markdown. This is a deliberate amendment, not a test being
  weakened to fit.
- `specs/008-round-trip-preservation.md` — `DropRecord['kind']` widens as above.

### Fixtures
Adds `fixtures/md/kitchen-sink.md`, covering every row of the supported subset
plus one raw HTML block and YAML front matter, with a manifest entry naming the
constructs that must survive a `.md` round trip. Reuses the existing `.docx`
fixtures for the conversion direction.
