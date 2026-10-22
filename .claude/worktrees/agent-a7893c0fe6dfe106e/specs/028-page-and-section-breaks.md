---
id: SPEC-028
title: Page and section breaks
status: draft
area: layout
depends_on: [SPEC-027]
---
## Why
Pagination is what makes this a word processor rather than a text editor. Once
content flows across real pages, headers, page numbers, footnotes and print all
become possible — and sections are how one document holds two page layouts.

## Acceptance criteria
- [ ] Content flows across multiple rendered pages, breaking at the bottom margin, with a visible gap between pages
- [ ] A page count is shown in the status bar as "Page N of M" reflecting the caret's page
- [ ] Ctrl/Cmd+Enter inserts a manual page break at the caret and the following content starts a new page
- [ ] An "Insert → Break" menu offers page break, column break, and section breaks (next page, continuous)
- [ ] Deleting a manual break rejoins the flow; breaks are visible when formatting marks are shown
- [ ] A "Show formatting marks" toggle (Ctrl/Cmd+Shift+P) displays paragraph marks, spaces, tabs and break markers
- [ ] A section break lets the following section have its own page size, orientation and margins, editable via Page setup's "Apply to: this section / whole document"
- [ ] Paragraph "keep with next", "keep lines together" and "page break before" options are settable and honoured by pagination
- [ ] Widow/orphan control is on by default and prevents single lines stranded across a page boundary
- [ ] Breaks and section properties round-trip through `.docx` (`w:br type="page"`, `w:sectPr`) and `.odt` (`fo:break-before`, master pages)
- [ ] Pagination of a 200-page document completes without freezing the UI, and typing in the middle of the document re-paginates only the affected region

## Out of scope
Headers and footers (SPEC-029). Columns (SPEC-030). Line numbers. Print
(SPEC-034).
