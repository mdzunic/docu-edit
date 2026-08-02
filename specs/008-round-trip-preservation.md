---
id: SPEC-008
title: Round-trip preservation of unsupported content
status: draft
area: io
depends_on: [SPEC-006, SPEC-007]
---
## Why
Every editor that rewrites a file it only partly understands eventually
destroys someone's work — a chart, a content control, a custom XML part gone
because we did not have a feature for it. The rule for this product: what we
cannot edit, we carry through untouched, and we tell the user when we cannot.

## Acceptance criteria
- [ ] Opening a document retains the original package bytes for the lifetime of the document
- [ ] Package parts the reader does not model (e.g. `customXml/`, embedded objects, charts, themes, `word/webSettings.xml`) are written back byte-identical on save to the same format
- [ ] XML elements inside `document.xml` / `content.xml` that the reader does not model are retained as opaque nodes in the document model, anchored to their position, and re-emitted in place on save
- [ ] Opening a file and saving it to the same format without editing produces a document that is semantically unchanged: a documented comparison (text, styles, part list) shows no differences
- [ ] Editing one paragraph and saving does not alter unrelated opaque content elsewhere in the document
- [ ] Content that cannot be preserved when converting between formats is listed in a "Conversion report" dialog shown once after a cross-format save, naming each dropped item type and its approximate location (e.g. "Chart on page 3 — not converted")
- [ ] The conversion report has a "Don't show again for this document" checkbox
- [ ] A test opens a `.docx` containing a chart, a content control and a custom XML part, edits one paragraph, saves, and asserts all three are still present and intact in the output

## Out of scope
Rendering the preserved content (a chart still shows as a placeholder box).
Editing preserved content. Macro (`vbaProject.bin`) preservation is included by
the part rule above but macro execution is never supported.
