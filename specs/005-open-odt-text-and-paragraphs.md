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
- [ ] Fixture tests open at least two real `.odt` files (LibreOffice Writer and an OpenOffice-era file) and assert text content, heading levels and formatting

## Out of scope
Tables, images, lists, headers/footers and other structural content (later
specs). `.ott` templates (SPEC-046). Encrypted ODF.
