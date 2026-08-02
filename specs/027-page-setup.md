---
id: SPEC-027
title: Page setup — size, margins, orientation
status: draft
area: layout
depends_on: [SPEC-002, SPEC-004, SPEC-005]
---
## Why
Until now the page has been a rectangle we chose. Documents arrive as Letter or
A4, portrait or landscape, with margins that matter to whoever printed them —
and if we render them at the wrong width, every line break in the document is
wrong.

## Acceptance criteria
- [ ] A "Page setup" dialog sets paper size (A4, Letter, Legal, A3, A5, Tabloid, Custom), orientation, and the four margins
- [ ] Custom size accepts width and height with a unit selector (cm, mm, in, pt)
- [ ] Changing any value re-renders the page and re-flows text immediately on OK
- [ ] The dialog's unit follows a document-level measurement-unit preference, and that preference is settable and persisted
- [ ] Values from the opened document populate the dialog, including documents whose size is a non-standard custom size
- [ ] Margins that leave a text area narrower than 1 cm are rejected with "Margins are too large for this page size"
- [ ] Page setup round-trips through `.docx` (`w:sectPr/w:pgSz`, `w:pgMar`) and `.odt` (`style:page-layout-properties`) with no unit drift
- [ ] The rendered page shows the text area accurately: content starts exactly at the top margin and wraps at the right margin
- [ ] A "Show/hide ruler" toggle displays a horizontal ruler with margin markers that reflect the current setup

## Out of scope
Different page setup per section (SPEC-028). Gutters, mirrored margins and
book-fold layout. Page borders and background colour. Print scaling (SPEC-034).
