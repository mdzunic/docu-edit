---
id: SPEC-010
title: Autosave and crash recovery
status: draft
area: io
depends_on: [SPEC-003, SPEC-011]
---
## Why
Nobody presses Ctrl+S in a browser any more, and a tab that dies mid-sentence
should not cost an afternoon. Autosave is also the safety net that makes every
later editing feature safe to ship.

## Acceptance criteria
- [ ] Edits are written to local storage automatically, at most 2 seconds after the user stops typing and at least once every 30 seconds during continuous typing
- [ ] The app bar shows a save state: "Saving…" while writing, "All changes saved" when idle and clean, and "Not saved — <reason>" on failure
- [ ] Closing the tab with unsaved changes triggers the browser's leave-confirmation prompt; a fully saved document does not
- [ ] Reopening a document restores the exact content, caret position and scroll position from the last autosave
- [ ] If the tab crashes or is force-closed mid-edit, reopening the document shows a bar "Recovered unsaved changes from <time>" with "Keep" and "Discard" actions
- [ ] "Discard" restores the last cleanly saved state and the bar disappears
- [ ] Autosave writes do not block typing: input remains responsive during a save of a 200-page document
- [ ] Two tabs open on the same document detect the conflict; the second tab shows "This document is open in another tab" and opens read-only until the other tab is closed
- [ ] Storage quota errors surface as "Not saved — browser storage is full" with a link to download the file

## Out of scope
Server-side autosave. Saving back to the user's original file on disk. Version
history (SPEC-040). Merge of concurrent edits (SPEC-044).
