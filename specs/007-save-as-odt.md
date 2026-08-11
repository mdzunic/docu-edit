---
id: SPEC-007
title: Save as .odt
status: draft
area: io
depends_on: [SPEC-005, SPEC-006]
---
## Why
Symmetry with SPEC-006, and the point at which the product becomes a converter:
open a `.docx`, save it as `.odt`, and the content survives.

## Acceptance criteria
- [ ] The "Download" menu offers "OpenDocument Text (.odt)" and downloads `<title>.odt`
- [ ] The written package contains an uncompressed `mimetype` entry stored first, plus `content.xml`, `styles.xml`, `meta.xml`, `settings.xml` and `META-INF/manifest.xml`
- [ ] The `mimetype` entry contains exactly `application/vnd.oasis.opendocument.text` and is written with ZIP STORE (no deflate), as ODF requires
- [ ] All document model content survives a save: text, run and paragraph formatting, paragraph styles, page layout
- [ ] Run and paragraph formatting is emitted as automatic styles referenced by name, not as invalid inline attributes
- [ ] `meta.xml` carries the title and a modification timestamp
- [ ] Files produced by the writer open in LibreOffice Writer without warnings — verified by a documented manual check recorded in the PR
- [ ] The writer is exposed as `writeOdt` in `src/lib/odf/writer.ts`, matching the `writeDocx` signature in "Writer API" below, with no DOM dependency
- [ ] An automated test round-trips every `.odt` fixture (open → save → re-open) and asserts `equalStructural(first, second)` — see "Equality" below — reporting the first differing path on failure
- [ ] Cross-format tests assert `equalSemantic` for `.docx` → save-as-`.odt` → open and for `.odt` → save-as-`.docx` → open, paired with a committed exclusion list naming what each direction is permitted to lose. `equalSemantic` yields no such list itself — it is a boolean — so the list is authored per direction here and becomes writer-reported in SPEC-008
- [ ] The exclusion list is empty for the content this spec covers — text, run formatting, paragraph formatting, styles and page layout convert without loss in both directions

## Out of scope
Flat ODF (`.fodt`). Digital signatures. Encryption.

## Decisions
If the document has no title, go with default Document 1.

## Reference

### Canonical document model
Defined normatively in `specs/002-create-blank-document.md` ("Canonical
document model"), and in `docs/document-model.md` once SPEC-002 is merged.
**Read one of those before starting.** In short: `DocumentModel → Section →
Block → Paragraph → Run`, plain JSON, blocks discriminated by a `type` string,
lengths in centimetres and type sizes in points (fields named `…Cm` / `…Pt`),
colours lowercase `"#rrggbb"`. This writer emits ODF's unit-suffixed `cm`/`pt`
strings from those numbers, rounded symmetrically with the SPEC-005 reader so a
round trip is lossless.

### Existing I/O surface
Both readers and the `.docx` writer already exist when this spec runs:
```ts
// src/lib/docx/reader.ts  parseDocx(data): Promise<ParsedDocxDocument>     — SPEC-004
// src/lib/odf/reader.ts   parseOdt(data):  Promise<ParsedOdtDocument>      — SPEC-005
// src/lib/docx/writer.ts  writeDocx(model, options?): Promise<WriteResult> — SPEC-006
// src/lib/odf/writer.ts   writeOdt(model, options?):  Promise<WriteResult> — this spec
// src/lib/openDocument.ts openDocumentFile(file): Promise<DocumentModel>

type WriteOdtOptions = { modified?: Date }   // injectable clock, mirroring
                                             // WriteDocxOptions
```
Preserving opaque content is not an option flag — SPEC-008 makes it default and
unconditional, which is why the options type carries only the timestamp.
SPEC-008 widens both writers to `Promise<WriteResult>`, where
`WriteResult = { bytes: Uint8Array; dropped: DropRecord[] }`
(`src/lib/opaque.ts`) — bytes rather than a `Blob` because the writers carry no
DOM dependency. `dropped` names each item
the writer could not represent in the target format, and is the single source of
both the conversion report and the cross-format exclusion list — only the writer
knows an item's type and location at the moment it drops it, which no comparison
of the resulting models can recover. Structure the conversion path so that list
can be collected as it goes.

Their specs are `specs/005-open-odt-text-and-paragraphs.md` and
`specs/006-save-as-docx.md` — read both; this spec is the ODF mirror of
SPEC-006 and must not diverge in API shape, error handling or test structure.

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json`, created by
SPEC-004 and extended by SPEC-005 (`odt/libreoffice-basic.odt`,
`odt/openoffice-legacy.odt`). The round-trip test iterates the manifest rather
than a hard-coded list.

### Equality
Defined normatively in `specs/002-create-blank-document.md` ("Equality"),
implemented in `src/lib/equality.ts`:
- **`equalStructural(a, b)`** — deep equality with `id` and three UI-state
  keys skipped, absent/`{}`/`[]` treated as one value, and XML-valued fields
  compared canonically via `canonicalizeXml`. It does not merge runs, drop
  default-valued props or normalise colour case. **Use this for same-format
  round trips.**
- **`equalSemantic(a, b)`** — `equalStructural` over the style-resolved model,
  so formatting that moves between a style and a run still compares equal.
  **Use this for cross-format conversion**, always with an explicit exclusion
  list. It returns a boolean and produces no diff, so the exclusion list is
  authored alongside the test, not derived from the comparison; SPEC-008 moves
  it to the writer's `dropped` array and renders that as the user-facing
  conversion report.

Do not hand-roll a subset comparison — use the exported helpers.
