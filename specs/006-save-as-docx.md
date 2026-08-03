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
- [ ] An automated test round-trips each `.docx` fixture (open → save → re-open) and asserts the re-opened model equals the first-opened model
- [ ] An automated test asserts the written XML validates against the WordprocessingML schema for the parts the writer emits
- [ ] Saving a document that was opened from `.odt` produces a valid `.docx` with content converted, not a renamed ODF file
- [ ] Download while the document is empty still produces a valid one-paragraph file

## Out of scope
Saving back to the original file handle in place (SPEC-010 covers autosave;
File System Access API is a later concern). PDF export (SPEC-034).
Preservation of parts the reader does not understand (SPEC-008).

## Decisions
Question: AC-14: The spec depends on SPEC-004, but requires saving a document opened from .odt. If SPEC-004 does not support reading .odt files, it is unclear how an .odt document can be opened to be saved. Please clarify if .odt reading is supported by SPEC-004 or if this AC belongs to a different spec.
Answer: There is no AC-14 for this spec, and it has nothing to do with .odt, but with .docx. This belongs to different spec - skip it.
