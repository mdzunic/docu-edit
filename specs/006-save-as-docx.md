---
id: SPEC-006
title: Save as .docx
status: draft
area: io
depends_on: [SPEC-004]
---
## Why
Opening is half a product. Users need the file back, and it has to open cleanly
in Word and LibreOffice — a file that triggers Word's repair dialog is worse
than no file at all.

## Acceptance criteria
- [ ] A "Download" action in the app bar offers "Word (.docx)" and downloads `<title>.docx`
- [ ] The written package is a valid OPC ZIP containing `[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, `word/styles.xml` and `docProps/core.xml`
- [ ] All document model content written by SPEC-004's reader survives a save: text, run formatting, paragraph formatting, paragraph styles, and section page setup
- [ ] `docProps/core.xml` carries the document title and a `dcterms:modified` timestamp
- [ ] Files produced by the writer open in Microsoft Word and LibreOffice Writer with no repair prompt and no lost text — verified by a documented manual check recorded in the PR, plus an automated validation step
- [ ] The writer is exposed as `writeDocx` in `src/io/docx/write.ts` with the signature in "Writer API" below, with no DOM dependency so it runs in a worker and in Node tests
- [ ] An automated test round-trips every `.docx` fixture (open → save → re-open) and asserts `equalStructural(first, second)` — see "Equality" below — reporting the first differing path on failure
- [ ] An automated test asserts the written XML validates against the WordprocessingML schema for the parts the writer emits
- [ ] Saving a document that was opened from `.odt` produces a valid `.docx` with content converted, not a renamed ODF file, and a test asserts `equalSemantic` between the source model and the re-opened `.docx` model
- [ ] Download while the document is empty still produces a valid one-paragraph file

## Out of scope
Saving back to the original file handle in place (SPEC-010 covers autosave;
File System Access API is a later concern). PDF export (SPEC-034).
Preservation of parts the reader does not understand (SPEC-008).

## Reference

### Canonical document model
Defined normatively in `specs/002-create-blank-document.md` ("Canonical
document model"), and in `docs/document-model.md` once SPEC-002 is merged.
**Read one of those before starting.** This writer consumes that model and
converts out of its units: `Emu → twips` (÷635) for `w:ind`, `w:spacing` and
`w:pgMar`, `Emu → twentieths of a point` where OOXML requires it, and
half-points straight through to `w:sz`. Rounding is half-up and must be
symmetric with the SPEC-004 reader so a round trip is lossless.

### Writer API — the surface later specs build on
```ts
// src/io/docx/write.ts
writeDocx(doc: Doc, opts?: WriteOptions): Promise<Blob>
type WriteOptions = { preserveOpaque?: boolean }   // SPEC-008 uses this
```
SPEC-007 adds `writeOdt` with the identical shape. The reader side already
exists from SPEC-004: `readDocx(bytes): Promise<{ doc, warnings }>` in
`src/io/docx/read.ts`, plus `detectFormat` in `src/io/detect.ts`.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json`, created by
SPEC-004 (`docx/word-basic.docx`, `docx/libreoffice-basic.docx`,
`docx/gdocs-basic.docx`) and extended by SPEC-005 with the `.odt` files. The
round-trip test iterates the manifest rather than a hard-coded list, so
fixtures added by later specs are covered automatically.

### Equality
Defined normatively in `specs/002-create-blank-document.md` ("Equality"):
- **`equalStructural`** — deep equality after canonicalization (ids dropped,
  default-valued props dropped, adjacent equal-prop runs merged, zero-length
  runs dropped, colours lowercased, lengths as integer `Emu`, keys sorted).
  This is the relation for the same-format round trip above. It is blind to
  representation choices the writer is free to make, and strict about content.
- **`equalSemantic`** — `equalStructural` over the style-resolved model, for
  the cross-format assertion above, where formatting may legitimately move
  between a style and a run.

Do not write a bespoke subset comparison for these tests; use the exported
helpers from `src/model/equality.ts`.
