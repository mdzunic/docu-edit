---
id: SPEC-045
title: Document generation API
status: draft
area: export
depends_on: [SPEC-006, SPEC-007, SPEC-022, SPEC-024]
---
## Why
"Generate new documents from scratch" is a first-class requirement, and the
buyers for it are systems, not people: invoicing, HR letters, reports. The
writers already exist — this exposes them as a stable, documented interface.

## Acceptance criteria
- [ ] A documented JavaScript API creates a document, appends headings, paragraphs with formatted runs, lists, tables, images, page breaks, headers and footers, and produces a `.docx` or `.odt` as a Blob/Buffer
- [ ] The API is usable both in the browser and in Node, with no DOM dependency in the generation path
- [ ] A JSON document-description format maps 1:1 to the API, so a document can be generated from data without writing imperative code, and its schema is published
- [ ] Generation validates its input and throws with a message naming the offending path (e.g. "body[3].cells: expected array") rather than emitting an invalid file
- [ ] Generated files open in Word and LibreOffice with no repair prompt, verified by fixtures covering each supported block type
- [ ] Every construct the API can generate can be re-opened by this product's own readers with no loss — an automated round-trip test asserts it
- [ ] A generated document can be opened in the editor directly (`generate → open`) without a download round trip
- [ ] The API is documented with a README and at least three runnable examples: an invoice with a table, a multi-page report with a TOC, and a letter with a letterhead image
- [ ] Generation of a 50-page document completes in under 2 seconds in Node

## Out of scope
A hosted HTTP service or authentication for it. Markdown or HTML as an input
format. Rendering to PDF from the API (compose with SPEC-034). Charts.
