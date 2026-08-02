---
id: SPEC-001
title: App shell and document workspace
status: draft
area: shell
depends_on: []
---
## Why
Everything else in this product lives inside one screen: a page canvas with a
toolbar above it. Building that frame once, empty, means every later feature
adds a control to a known place instead of inventing its own layout.

## Acceptance criteria
- [ ] Visiting `/` shows the app shell: a top app bar, a toolbar strip below it, and a document canvas area
- [ ] The app bar shows the product name and the current document title; with no document open the title reads "No document"
- [ ] The canvas area with no document open shows an empty state with the text "No document open" and two buttons, "New document" and "Open document"
- [ ] Both buttons are present and visible but may be non-functional in this spec
- [ ] Visiting `/doc/:id` for an id that does not exist shows a "Document not found" message and a link back to `/`
- [ ] The toolbar strip is a horizontal container that scrolls horizontally when its contents overflow, rather than wrapping or clipping
- [ ] The layout fills the viewport height with no page-level vertical scrollbar; only the canvas area scrolls
- [ ] At a 1024px-wide viewport the toolbar and canvas are both usable — no overlap, no content cut off

## Out of scope
Any document parsing, editing, or persistence. Mobile/touch layout. Theming or
dark mode. Authentication.

## Decisions
Question: For a valid document ID at /doc/:id, what should be displayed in the canvas and as the document title in the app bar, given that document parsing and persistence are explicitly out of scope?
Answer: use Datetime timestamp (from first opening/creating doc) for now, it's simply the current client-side time on page load

Question: How should the system determine if a document ID at /doc/:id is valid or non-existent, given that document persistence is explicitly out of scope?
Answer: Treat all IDs as valid for now

Question: The acceptance criteria state that visiting /doc/:id for a non-existent ID should show a 'Document not found' message, but the Decisions section says to treat all IDs as valid for now. Should the 'Document not found' state be implemented, or is it impossible to reach given the decision to treat all IDs as valid?
Answer: implement the 'Document not found' message, but don't test it as long there is on persistence
