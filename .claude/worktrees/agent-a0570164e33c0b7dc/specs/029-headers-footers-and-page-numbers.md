---
id: SPEC-029
title: Headers, footers and page numbers
status: draft
area: layout
depends_on: [SPEC-028]
---
## Why
Every report, contract and letterhead has them, and page numbers are the first
computed field in the product — the mechanism that later carries the table of
contents and cross-references.

## Acceptance criteria
- [ ] Double-clicking the top or bottom page margin, or "Insert → Header/Footer", enters header/footer editing with the body content dimmed
- [ ] Header and footer content supports text, formatting, tables and images like body content
- [ ] Escape or double-clicking the body returns to body editing
- [ ] Header/footer content repeats on every page of its section and edits apply to all of them at once
- [ ] Options for "Different first page" and "Different odd and even pages" are available per section and produce distinct header/footer sets
- [ ] A "Insert page number" control offers position (header/footer, left/centre/right) and format (1, 2, 3 / i, ii, iii / a, b, c), plus "Page X of Y"
- [ ] Page numbers render the correct value on every page and update when pagination changes
- [ ] "Start numbering at" is settable per section, and a section can be set to continue from the previous section
- [ ] Headers and footers round-trip through `.docx` (`header1.xml`/`footer1.xml` parts with `w:headerReference`, `PAGE`/`NUMPAGES` fields) and `.odt` (master page styles with `text:page-number`/`text:page-count`)
- [ ] Opening a fixture with different first-page and odd/even headers renders them on the right pages

## Out of scope
Watermarks. Fields other than page number and page count (date, author, etc.).
Chapter-based numbering ("2-1").
