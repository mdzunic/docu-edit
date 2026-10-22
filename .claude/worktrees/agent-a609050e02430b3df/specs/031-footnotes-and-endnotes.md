---
id: SPEC-031
title: Footnotes and endnotes
status: draft
area: layout
depends_on: [SPEC-028]
---
## Why
Academic, legal and government documents are unusable without them, and
footnotes are the hardest interaction between content and pagination: a note
must sit on the page its reference lands on, and adding one can change that.

## Acceptance criteria
- [ ] "Insert → Footnote" (Ctrl/Cmd+Alt+F) inserts a superscript reference mark at the caret and places the caret in a new footnote area at the bottom of the current page
- [ ] "Insert → Endnote" (Ctrl/Cmd+Alt+D) does the same at the end of the document
- [ ] The footnote area is separated from body text by a separator line and holds fully editable, formattable content
- [ ] Footnotes are numbered automatically in document order and renumber immediately when one is inserted or deleted
- [ ] Deleting a reference mark deletes its note; deleting the note's text does not orphan the mark
- [ ] A footnote whose reference moves to another page moves with it during re-pagination
- [ ] A footnote too long for the remaining page space splits across pages rather than overflowing
- [ ] Hovering a reference mark shows the note text in a tooltip; clicking it jumps to the note, and clicking the note's number jumps back
- [ ] A footnote settings dialog sets numbering format (1/i/a/*), start number, and restart per page or per section
- [ ] Footnotes and endnotes round-trip through `.docx` (`footnotes.xml`, `endnotes.xml` with `w:footnoteReference`) and `.odt` (`text:note` with `text:note-citation` and `text:note-body`)
- [ ] Opening a fixture with 20+ footnotes across several pages places every note on the page of its reference

## Out of scope
Citations and bibliography management. Converting footnotes to endnotes in
bulk. Continuation notices ("continued on next page").
