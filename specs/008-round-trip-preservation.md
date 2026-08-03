---
id: SPEC-008
title: Round-trip preservation of unsupported content
status: draft
area: io
depends_on: [SPEC-006, SPEC-007]
---
## Why
Every editor that rewrites a file it only partly understands eventually
destroys someone's work — a chart, a content control, a custom XML part gone
because we did not have a feature for it. The rule for this product: what we
cannot edit, we carry through untouched, and we tell the user when we cannot.

## Acceptance criteria
- [ ] Opening a document retains the original package bytes for the lifetime of the document
- [ ] Package parts the reader does not model (e.g. `customXml/`, embedded objects, charts, themes, `word/webSettings.xml`) are written back byte-identical on save to the same format
- [ ] XML elements inside `document.xml` / `content.xml` that the reader does not model are retained as opaque nodes in the document model, anchored to their position, and re-emitted in place on save
- [ ] Opening a file and saving it to the same format without editing produces `equalStructural(openedModel, reopenedModel)` and an `equalBytes` result for every preserved opaque part — see "Equality" below
- [ ] Editing one paragraph and saving does not alter unrelated opaque content elsewhere in the document
- [ ] Content that cannot be preserved when converting between formats is listed in a "Conversion report" dialog shown once after a cross-format save, naming each dropped item type and its approximate location (e.g. "Chart on page 3 — not converted")
- [ ] The conversion report has a "Don't show again for this document" checkbox
- [ ] A test opens a `.docx` containing a chart, a content control and a custom XML part, edits one paragraph, saves, and asserts all three are still present and intact in the output

## Out of scope
Rendering the preserved content (a chart still shows as a placeholder box).
Editing preserved content. Macro (`vbaProject.bin`) preservation is included by
the part rule above but macro execution is never supported.

## Reference

### Canonical document model
Defined normatively in `specs/002-create-blank-document.md` ("Canonical
document model"), and in `docs/document-model.md` once SPEC-002 is merged.
**Read one of those before starting.** Two fields exist for this spec:
```ts
type Doc = { …; opaque: OpaquePart[] }        // package parts we do not model
type Block  = Paragraph | OpaqueBlock          // unmodelled block-level XML
type Inline = Run | LineBreak | OpaqueInline   // unmodelled inline XML
```
`OpaqueBlock` / `OpaqueInline` hold the original XML subtree verbatim plus its
position; `OpaquePart` holds the original part bytes plus its part name,
content type and relationship entries. Opaque nodes are compared with XML
canonicalization (C14N) by `canonicalize`, so insignificant whitespace and
attribute order do not cause false failures.

### Existing I/O surface
```ts
// src/io/docx/read.ts   readDocx(bytes): Promise<ReadResult>      — SPEC-004
// src/io/odt/read.ts    readOdt(bytes):  Promise<ReadResult>      — SPEC-005
// src/io/docx/write.ts  writeDocx(doc, opts?): Promise<Blob>      — SPEC-006
// src/io/odt/write.ts   writeOdt(doc, opts?): Promise<Blob>       — SPEC-007
type WriteOptions = { preserveOpaque?: boolean }
```
This spec makes `preserveOpaque` the default behaviour rather than an option,
and extends both readers to produce opaque nodes instead of dropping content.
Read `specs/006-save-as-docx.md` and `specs/007-save-as-odt.md` first.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json` (SPEC-004). This
spec adds `fixtures/docx/word-charts-and-controls.docx` — containing a chart, a
content control and a custom XML part — and `fixtures/odt/lo-embedded.odt`,
with manifest entries listing the parts that must survive.

### Equality
Defined normatively in `specs/002-create-blank-document.md` ("Equality"),
implemented in `src/model/equality.ts`. All three relations are used here:
- **`equalBytes`** — for each preserved opaque part, before and after a save.
  This is the only place byte equality is required.
- **`equalStructural`** — for the no-edit same-format round trip, and for
  asserting that editing one paragraph left every other node untouched.
- **`equalSemantic`** — for cross-format saves, paired with the exclusion list
  that this spec surfaces to the user as the conversion report. The report's
  contents and the test's exclusion list are generated from the same source, so
  they cannot drift apart.
