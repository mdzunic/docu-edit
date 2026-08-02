---
id: SPEC-004
title: Open .docx — text and paragraphs
status: draft
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
- [ ] Fixture-based tests open at least three real `.docx` files (Word-authored, LibreOffice-authored, and Google-Docs-exported) and assert paragraph count, text content, and heading levels

## Out of scope
Tables, images, headers/footers, footnotes, lists, comments and tracked changes
— each has its own spec; unsupported parts are handled by SPEC-008. Password-
protected documents. Macros.
