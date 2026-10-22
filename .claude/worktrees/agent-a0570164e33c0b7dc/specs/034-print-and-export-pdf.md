---
id: SPEC-034
title: Print and export to PDF
status: draft
area: export
depends_on: [SPEC-028, SPEC-029]
---
## Why
A document that cannot be printed or sent as a PDF is not finished. This is
also the honest test of the layout engine: what comes out of the printer either
matches the screen or it does not.

## Acceptance criteria
- [ ] Ctrl/Cmd+P opens a print preview showing the paginated document exactly as it will print
- [ ] The preview offers page range, and printing goes through the browser's print dialog with our page size and margins applied
- [ ] Printed output matches on-screen pagination: the same content appears on the same page numbers
- [ ] Headers, footers, page numbers, footnotes and columns all print in their correct positions
- [ ] Screen-only affordances (caret, selection, comment markers, formatting marks, resize handles, the update-TOC control) do not appear in print
- [ ] "Download → PDF (.pdf)" produces a PDF of the whole document with selectable, searchable text — not a raster image
- [ ] The PDF has document title metadata and PDF bookmarks/outline generated from Heading 1–3
- [ ] Hyperlinks in the document are clickable links in the PDF
- [ ] Fonts used are embedded or subset in the PDF so it renders identically on a machine without them
- [ ] A 100-page document exports to PDF in under 10 seconds with a progress indicator, and the UI stays responsive
- [ ] Export failures show "Couldn't create the PDF" with a reason and leave the document untouched

## Out of scope
PDF/A archival conformance. Tagged-PDF accessibility structure beyond outline
bookmarks. Password-protecting the PDF. Printing selected text only.
