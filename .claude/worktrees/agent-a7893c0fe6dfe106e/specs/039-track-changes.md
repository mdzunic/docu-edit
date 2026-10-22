---
id: SPEC-039
title: Track changes
status: draft
area: collab
depends_on: [SPEC-015, SPEC-038]
---
## Why
Legal, HR and procurement documents move through redlines. This is the single
biggest feature gap between browser editors and Word, and it must be right:
silently dropping tracked changes on save would be a data-loss bug in a
document someone is negotiating over.

## Acceptance criteria
- [ ] A "Track changes" toggle (Ctrl/Cmd+Shift+E) turns recording on and off, with clear on/off state in the toolbar and status bar
- [ ] With tracking on, inserted text renders underlined in the author's colour and deleted text renders struck through rather than disappearing
- [ ] Formatting changes are recorded and shown as a change bar in the margin plus a description in the review pane ("Formatted: Bold")
- [ ] Each change records author and timestamp, shown on hover and in the review pane
- [ ] A review pane lists all changes in document order with Accept and Reject per change, and "Accept all" / "Reject all"
- [ ] Next/previous change navigation moves the caret between changes
- [ ] A display mode selector offers "All markup", "Simple markup" (change bars only), "No markup" (as if all accepted) and "Original" (as if all rejected)
- [ ] Accepting or rejecting is undoable, and rejecting a deletion restores the exact original text and formatting
- [ ] Editing over an existing tracked insertion by another author produces a nested change, not a lost one
- [ ] Tracked changes round-trip through `.docx` (`w:ins`, `w:del`, `w:delText`, `w:rPrChange`, `w:pPrChange`) and `.odt` (`text:tracked-changes` with `text:changed-region`), preserving author and date
- [ ] A document with existing tracked changes opened from Word shows every change with its original author, and saving without touching them leaves them byte-equivalent
- [ ] Tracked deletions are excluded from word count, find/replace results and spell check
- [ ] Changes made by list, table and image operations are tracked, not silently applied

## Out of scope
Comparing two documents to generate a redline. Merging changes from multiple
copies. Locking tracking on so it cannot be disabled. Per-author filtering
beyond the review pane list.
