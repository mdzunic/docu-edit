---
id: SPEC-040
title: Version history and restore
status: draft
area: collab
depends_on: [SPEC-010]
---
## Why
Autosave protects the last state; version history protects the ones before it.
Users need to answer "what did this look like on Tuesday" and get back to it
without keeping a folder of `final-v3-REALLY-final.docx` files.

## Acceptance criteria
- [ ] The app snapshots the document periodically during editing and at every explicit save/download, storing snapshots locally
- [ ] "File → Version history" opens a pane listing versions with timestamp, author identity and word-count delta, newest first
- [ ] Selecting a version shows it read-only in the canvas with additions and deletions since the previous version highlighted
- [ ] "Restore this version" replaces the current content and is itself undoable and recorded as a new version
- [ ] "Name this version" labels a version, and named versions can be filtered to on their own
- [ ] "Download this version" exports the selected snapshot as `.docx` or `.odt` without changing the current document
- [ ] Snapshots are stored as deltas so history for a 100-page document edited over a week stays under a documented storage budget, and the oldest unnamed versions are pruned first when the budget is hit
- [ ] Pruning never removes named versions or the last 10 versions
- [ ] Closing and reopening the document preserves the full history

## Out of scope
Server-side or cross-device history. Restoring a single paragraph from an old
version. Branching. Diff between two arbitrary versions.
