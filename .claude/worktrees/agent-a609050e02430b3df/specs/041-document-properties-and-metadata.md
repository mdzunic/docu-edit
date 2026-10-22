---
id: SPEC-041
title: Document properties and metadata
status: draft
area: io
depends_on: [SPEC-006, SPEC-007]
---
## Why
Metadata travels with the file into document management systems, and it is also
a privacy liability: an author name or a comment left in a file that goes to a
counterparty. Users need to see and control both.

## Acceptance criteria
- [ ] "File → Properties" shows a dialog with General (file name, format, size, created, modified), Description (title, subject, keywords, comments, author, last modified by) and Statistics (from SPEC-036) tabs
- [ ] Editable fields write to `.docx` `docProps/core.xml` and `docProps/app.xml`, and to `.odt` `meta.xml`
- [ ] Metadata on an opened document is shown as it was in the source file, including fields we do not edit
- [ ] Custom properties (name, type, value) can be added, edited and removed, and round-trip through both formats
- [ ] A "Remove personal information on save" toggle strips author, last-modified-by, comment authors and revision timestamps when saving, and states in the dialog exactly what it removes
- [ ] With that toggle on, tracked-change and comment authors are replaced with "Author 1", "Author 2" consistently rather than deleted
- [ ] The document's identity/author name used by comments and track changes is settable here and persists across sessions
- [ ] Metadata fields are treated as untrusted text on display — no markup or formula injection, covered by a test

## Out of scope
Document classification labels and DLP integration. Digital signatures. Custom
XML part editing.
