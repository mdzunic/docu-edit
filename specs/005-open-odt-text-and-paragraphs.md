---
id: SPEC-005
title: Open .odt — text and paragraphs
status: draft
area: io
depends_on: [SPEC-004]
---
## Why
ODF is the other half of the promise, and it is what LibreOffice, public-sector
users and several European governments hand out. The same document model must
accept both formats so nothing downstream cares which one a document came from.

## Acceptance criteria
- [ ] The file picker accepts `.odt` alongside `.docx`, and the file's format is detected from its contents (`mimetype` entry) rather than its extension
- [ ] The reader parses `content.xml` and `styles.xml` from the ODF package
- [ ] Paragraphs (`text:p`), headings (`text:h` with `text:outline-level`) and spans (`text:span`) map to the document model
- [ ] ODF whitespace handling is correct: `text:s` (with `text:c`), `text:tab` and `text:line-break` produce the same text content a user sees in LibreOffice
- [ ] Automatic styles from `content.xml` and named styles from `styles.xml` are resolved, including style inheritance via `style:parent-style-name`
- [ ] Bold, italic, underline, strikethrough, font family, size, colour and highlight resolve to the same run formatting model used by the `.docx` reader
- [ ] Paragraph alignment, indents and spacing resolve to the same paragraph formatting model
- [ ] Page layout (`style:page-layout`) sets page size, orientation and margins
- [ ] The format badge reads ODT
- [ ] A `.fodt` (flat XML ODF) file is rejected with "Flat ODF files aren't supported yet"
- [ ] The reader is exposed as `readOdt` in `src/io/odt/read.ts` with the same `ReadResult` shape as `readDocx`, and reports anything it cannot map as a `Warning`
- [ ] `fixtures/odt/libreoffice-basic.odt` and `fixtures/odt/openoffice-legacy.odt` are added to the shared fixture set with `manifest.json` entries
- [ ] Fixture tests open both and assert text content, heading levels and run formatting against the manifest expectations
- [ ] A test asserts that `libreoffice-basic.odt` and `word-basic.docx` — the same document saved from both applications — produce `equalSemantic` models

## Out of scope
Tables, images, lists, headers/footers and other structural content (later
specs). `.ott` templates (SPEC-046). Encrypted ODF.

## Reference

### Canonical document model
Defined normatively in `specs/002-create-blank-document.md` ("Canonical
document model"), and in `docs/document-model.md` in the repo once SPEC-002 is
merged. **Read one of those before starting.** Key point for this spec: the
model is format-neutral. ODF's `cm`/`in`/`pt` values convert to integer `Emu`
(360000/cm, 914400/in, 12700/pt) on read, and ODF point font sizes convert to
integer half-points. Nothing ODF-specific may leak into the model.

### Reader API — established by SPEC-004
`specs/004-open-docx-text-and-paragraphs.md` defines the I/O surface this spec
mirrors:
```ts
// src/io/detect.ts
detectFormat(bytes: ArrayBuffer): 'docx' | 'odt' | 'doc' | 'fodt' | 'unknown'

// src/io/odt/read.ts   (this spec; src/io/docx/read.ts already exists)
readOdt(bytes: ArrayBuffer): Promise<ReadResult>
type ReadResult = { doc: Doc; warnings: Warning[] }
type Warning    = { code: string; message: string; path?: string }
```

### Fixtures
The shared set at `fixtures/` with `fixtures/manifest.json`, created by
SPEC-004. This spec adds the `fixtures/odt/` files listed above and their
manifest entries; it does not create a parallel fixture scheme.

### Equality
`equalSemantic` is defined in `specs/002-create-blank-document.md`
("Equality"): structural equality over the style-resolved model. It is the
right relation for the cross-application assertion above, because Word and
LibreOffice legitimately differ on whether formatting sits in a style or on the
run.
