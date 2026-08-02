---
id: SPEC-033
title: Table of contents
status: draft
area: structure
depends_on: [SPEC-019, SPEC-029]
---
## Why
The payoff for having real heading styles and real pagination. A generated,
updatable TOC is also the feature people cite when explaining why they cannot
use a browser editor for long documents.

## Acceptance criteria
- [ ] "Insert → Table of contents" inserts a TOC at the caret built from Heading 1–6 paragraphs
- [ ] The TOC shows heading text, a dot leader, and the page number the heading falls on
- [ ] A TOC options dialog sets which heading levels to include (1–6), whether to show page numbers, and whether to use dot leaders or hyperlinks only
- [ ] Each entry is clickable and scrolls the caret to that heading
- [ ] The TOC block shows an "Update table of contents" control when the document's headings or pagination have changed since it was generated
- [ ] Updating offers "Update page numbers only" and "Update entire table"
- [ ] TOC content is not directly editable; attempting to type inside it shows the update hint instead
- [ ] Deleting the TOC leaves the headings untouched
- [ ] The TOC is written to `.docx` as a `TOC` field with cached result text so Word shows it without an update prompt, and to `.odt` as `text:table-of-content` with an index body
- [ ] Opening a fixture containing a TOC created in Word renders the existing entries and allows regeneration
- [ ] A TOC for a document with 300 headings generates in under 1 second

## Out of scope
Table of figures, index, and bibliography. Custom TOC entry styles beyond the
built-in TOC 1–6. Manual TOC entries.
