---
id: SPEC-038
title: Comments
status: draft
area: collab
depends_on: [SPEC-011, SPEC-025]
---
## Why
Documents are reviewed more often than they are written. Comments are the
review workflow, and both formats store them, so a document commented in Word
must show its comments here and vice versa.

## Acceptance criteria
- [ ] Selecting text and choosing "Insert → Comment" (Ctrl/Cmd+Alt+M) highlights the range and opens a comment composer in a right-hand margin pane
- [ ] Submitting adds the comment with an author name and timestamp; the author name comes from a settable identity, defaulting to "You"
- [ ] Comment threads support replies, and the thread is shown anchored beside its highlighted range
- [ ] Clicking a highlighted range focuses its thread, and clicking a thread scrolls to and highlights its range
- [ ] A thread can be resolved, which removes the highlight and moves the thread to a "Resolved" section that can be reopened
- [ ] A comment's author can edit or delete their own comments; deleting a thread removes its anchor
- [ ] Editing text inside a commented range keeps the comment anchored to the remaining text; deleting the entire range deletes the comment with an undoable warning
- [ ] Comments round-trip through `.docx` (`comments.xml`, `w:commentRangeStart/End`, `w:commentReference`) and `.odt` (`office:annotation` with `office:annotation-end`), including replies where the format supports them
- [ ] Opening a Word document with comments shows every comment, its author and its date, correctly anchored
- [ ] A comments-only view lists all threads in document order with filters for open/resolved and by author
- [ ] Comment text is rendered as plain text — no markup injection — covered by a security test

## Out of scope
@-mentions and notifications. Suggested edits (SPEC-039). Real-time comment
sync between users (SPEC-044). Rich text inside comments.
