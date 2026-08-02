---
id: SPEC-002
title: Create a new blank document
status: draft
area: io
depends_on: [SPEC-001]
---
## Why
The cheapest way to prove the document model works end to end is to make one
from nothing. This spec defines the in-memory document model that every later
feature reads and writes.

## Acceptance criteria
- [ ] Clicking "New document" in the empty state creates a document and navigates to `/doc/:id`
- [ ] A new document contains exactly one empty paragraph and renders a page-shaped canvas (white page, visible margins) centred in the canvas area
- [ ] The default title is "Untitled document" and appears in the app bar
- [ ] Clicking the title in the app bar makes it editable; typing a new title and pressing Enter or blurring updates it, and Escape reverts it
- [ ] A title of only whitespace reverts to "Untitled document"
- [ ] The document model is a serialisable tree with a documented shape: a document has sections, a section has block-level children, a paragraph has inline runs, and a run has text plus a formatting record
- [ ] The model carries a `defaultStyles` record (page size, margins, default font family and size) initialised to A4, 2.54cm margins, and the product default body font
- [ ] Unit tests cover creating a document, reading and writing its title, and serialising the model to JSON and back without loss

## Out of scope
Typing into the document (SPEC-011). Saving to disk (SPEC-006, SPEC-007).
Persistence across reloads (SPEC-003, SPEC-010). Templates (SPEC-046).
