---
id: SPEC-037
title: Document outline navigator
status: draft
area: shell
depends_on: [SPEC-019, SPEC-028]
---
## Why
Scrolling is a bad way to move around a 60-page document. The outline pane is
how people navigate long documents in both Word and Docs, and reordering by
dragging a heading is the fastest restructuring tool there is.

## Acceptance criteria
- [ ] A collapsible left sidebar shows the document outline built from Heading 1–6, indented by level
- [ ] The heading containing the caret is highlighted, and the highlight follows scrolling and caret movement
- [ ] Clicking an outline entry scrolls to that heading and places the caret at its start
- [ ] Outline entries with children can be collapsed and expanded in the pane
- [ ] Dragging an outline entry moves that heading and all its content (including sub-headings) to the new position as one undo step
- [ ] A document with no headings shows "Headings you add will appear here"
- [ ] The pane can be toggled from the View menu and its open/closed state persists per user
- [ ] Collapsing a heading in the document body itself (a chevron next to the heading) hides its content until expanded, and the hidden state is never written to the file
- [ ] The pane is keyboard navigable — arrow keys move, Enter jumps — and is exposed as a navigation landmark with a tree role

## Out of scope
Thumbnail page navigation. Outline view as an editing mode (promote/demote from
the pane). Search within the outline.
