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
- [ ] The writer is exposed as `writeDocx` in `src/lib/docx/writer.ts` with the signature in "Writer API" below, with no DOM dependency so it runs in a worker and in Node tests
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
converts out of its units: centimetres → twips (× 567) for `w:ind` and
`w:pgMar`, points → twentieths of a point (× 20) for `w:spacing`, and points →
half-points (× 2) for `w:sz`. Rounding is half-up and must be symmetric with
the SPEC-004 reader so a round trip is lossless within SPEC-014's 0.01 cm
tolerance.

### Writer API — the surface later specs build on
```ts
// src/lib/docx/writer.ts
writeDocx(model: DocumentModel, options?: WriteDocxOptions): Promise<WriteResult>
type WriteDocxOptions = { modified?: Date }   // injectable clock, so tests can
                                              // assert an exact core.xml
```
SPEC-008 widens the return type from a bare package to `WriteResult`, defined
in `src/lib/opaque.ts`:

```ts
interface WriteResult { bytes: Uint8Array; dropped: DropRecord[] }
```
`dropped` names what a cross-format save could not represent. That array — not
any comparison of models — is the source of both the conversion report and the
cross-format test's exclusion list, because only the writer knows an item's
type and location at the point it drops it. Keep the conversion path structured
so that list can be collected rather than reconstructed later.

Note `WriteResult.bytes` is a `Uint8Array`, not a `Blob`: the writers run
without DOM (the criterion above), and turning bytes into a `Blob` and an
object URL is the download caller's job.

Preserving opaque content is **not** an option flag. SPEC-008 makes it the
default and unconditional behaviour, which is why `WriteDocxOptions` carries
only the injectable timestamp.

SPEC-007 adds `writeOdt(model, options?)` in `src/lib/odf/writer.ts` with the
identical shape. The reader side already exists from SPEC-004:
`parseDocx(data): Promise<ParsedDocxDocument>` in `src/lib/docx/reader.ts`,
with format detection by content in `src/lib/openDocument.ts`.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json`, created by
SPEC-004 (`docx/word-basic.docx`, `docx/libreoffice-basic.docx`,
`docx/gdocs-basic.docx`) and extended by SPEC-005 with the `.odt` files. The
round-trip test iterates the manifest rather than a hard-coded list, so
fixtures added by later specs are covered automatically.

### Equality
Defined normatively in `specs/002-create-blank-document.md` ("Equality"):
- **`equalStructural`** — deep equality with `id` and three UI-state keys
  skipped, absent/`{}`/`[]` treated as one value, and XML-valued fields
  compared canonically via `canonicalizeXml`. It does **not** merge runs, drop
  default-valued props or normalise colour case — the readers and the editing
  primitives keep the model in that shape, and equality compares what the model
  holds. This is the relation for the same-format round trip above.
- **`equalSemantic`** — `equalStructural` over the style-resolved model, for
  the cross-format assertion above, where formatting may legitimately move
  between a style and a run. It returns a boolean and produces no diff or
  exclusion list; SPEC-008 pairs it with the writer's `dropped` array.

Do not write a bespoke subset comparison for these tests; use the exported
helpers from `src/lib/equality.ts`. Note that `equalSemantic` and `resolve` are
specified by SPEC-002 but **not yet implemented** — see that spec's
"Implementation status".
