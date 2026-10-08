---
id: SPEC-004
title: Open .docx — text and paragraphs
status: done
area: io
depends_on: [SPEC-002, SPEC-003]
---
## Why
`.docx` is the format most documents arrive in. This spec gets real Word
content onto the canvas — text, paragraphs, and the direct formatting that
carries most of a document's meaning — and establishes the OOXML reader that
later specs extend.

## Acceptance criteria
- [ ] "Open document" opens a file picker limited to `.docx`; choosing a file parses it and opens it at `/doc/:id`
- [ ] The reader unzips the OPC package and parses `word/document.xml` via the relationships in `word/_rels/document.xml.rels` rather than assuming fixed part names
- [ ] Paragraphs (`w:p`) and runs (`w:r`) map to the document model, preserving text content including spaces (`xml:space="preserve"`), tabs, and line breaks (`w:br`)
- [ ] Run properties bold, italic, underline, strikethrough, font family, font size, text colour and highlight are read into run formatting
- [ ] Paragraph properties alignment, left/right/first-line indent, space before/after and line spacing are read into paragraph formatting
- [ ] Paragraph style references (`w:pStyle`) are resolved against `word/styles.xml`, and built-in heading styles render at their defined size and weight
- [ ] Page size, orientation and margins from `w:sectPr` are applied to the rendered page
- [ ] The document title in the app bar comes from the file name without extension; the format badge reads DOCX
- [ ] A file that is not a valid ZIP, or is a ZIP without `word/document.xml`, shows the error "This file isn't a Word document we can open" and leaves the previous state untouched
- [ ] A `.doc` (binary Word 97) file is detected and rejected with "Old .doc files aren't supported — save as .docx first"
- [ ] The reader is exposed as `parseDocx` in `src/lib/docx/reader.ts` with the signature given in "Reader API" below, and does no DOM work so it can run in a worker and in Node tests
- [ ] Anything the reader cannot map is preserved rather than thrown or silently dropped — as opaque parts and nodes once SPEC-008 lands, and never as a thrown error short of an unopenable package
- [ ] The fixture set described in "Fixtures" below is committed with its `manifest.json`, and every fixture's provenance (which application produced it) is recorded there
- [ ] Fixture-based tests open `word-basic.docx`, `libreoffice-basic.docx` and `gdocs-basic.docx` and assert paragraph count, text content, heading levels and run formatting against the expectations in the manifest

## Out of scope
Tables, images, headers/footers, footnotes, lists, comments and tracked changes
— each has its own spec; unsupported parts are handled by SPEC-008. Password-
protected documents. Macros.

## Reference

### Canonical document model
Defined normatively in `specs/002-create-blank-document.md` ("Canonical
document model") and, once SPEC-002 is merged, in `docs/document-model.md` in
the repo. **Read one of those before starting.** In short: a plain-JSON tree
`DocumentModel → Section → Block → Paragraph → Run`, blocks discriminated by a
`type` string and runs holding plain text with `"\t"` and `"\n"` for tabs and
forced breaks; lengths are centimetres and type sizes points, in fields named
`…Cm` and `…Pt`; colours are lowercase `"#rrggbb"`. This reader converts
OOXML's twips and half-points into those units — it never stores twips.

### Reader API
```ts
// src/lib/docx/reader.ts
parseDocx(data: ArrayBuffer | Uint8Array): Promise<ParsedDocxDocument>

// src/lib/openDocument.ts — the caller that picks a reader
openDocumentFile(file: File): Promise<DocumentModel>
```
`ParsedDocxDocument` is the parsed package: `sections`, `defaultStyles`,
`styles`, `opaque`, `sourceXml`, `media` and the per-feature fields later specs
add. `openDocumentFile` turns it into a `DocumentModel` via `buildModel(name,
format, parsed)`, so the reader itself never constructs the model's identity or
title.

**Format detection is by content, not extension**, and lives in the caller
rather than in a separate module: `openDocumentFile` calls `detectOdf(bytes)`
(`src/lib/odf/reader.ts`), which returns `"package" | "flat" | "none"` by
reading the ODF `mimetype` entry. `"package"` takes the `parseOdt` path,
`"flat"` raises `FlatOdfError`, and anything else falls through to `parseDocx`,
which recognises a legacy `.doc` itself. SPEC-050 documents markdown as the one
deliberate exception to content sniffing.

`parseDocx` does not return warnings. Content it does not understand is
preserved as opaque parts and nodes (SPEC-008) rather than reported, and it
throws for a package it cannot open at all — `LegacyDocError` for a Word
97-2003 `.doc` and `DocxOpenError` otherwise, each carrying the user-facing
message the criteria above name (`LEGACY_DOC_MESSAGE`, `INVALID_DOCX_MESSAGE`).
SPEC-005 adds `parseOdt` with the same shape (`OdtOpenError`,
`INVALID_ODT_MESSAGE`); SPEC-006 adds `writeDocx`.

### Fixtures
Committed under `fixtures/`, shared by every I/O spec:
```
fixtures/
  docx/  word-basic.docx  libreoffice-basic.docx  gdocs-basic.docx
  odt/                                    (SPEC-005 adds these)
  manifest.json
```
`manifest.json` maps each file to its provenance and its expected assertions
(paragraph count, heading outline, sample run formatting), so later specs
extend the set by adding a file plus an entry rather than by inventing their
own expectations. Later specs add fixtures for the features they introduce —
`word-tables.docx`, `word-footnotes.docx` and so on — under the same rule.
