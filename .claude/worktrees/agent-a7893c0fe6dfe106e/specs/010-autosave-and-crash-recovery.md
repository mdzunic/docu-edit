---
id: SPEC-010
title: Autosave and crash recovery
status: done
area: io
depends_on: [SPEC-003, SPEC-011]
---
## Why
Nobody presses Ctrl+S in a browser any more, and a tab that dies mid-sentence
should not cost an afternoon. Autosave is also the safety net that makes every
later editing feature safe to ship.

## Acceptance criteria
- [ ] Edits are written to the IndexedDB-backed document store from SPEC-003 automatically, at most 2 seconds after the user stops typing and at least once every 30 seconds during continuous typing — see "Storage" below; `window.localStorage` is not used for document content
- [ ] The app bar shows a save state: "Saving…" while writing, "All changes saved" when idle and clean, and "Not saved — <reason>" on failure
- [ ] Closing the tab with unsaved changes triggers the browser's leave-confirmation prompt; a fully saved document does not. "Unsaved" means a draft write is in flight or has failed — see "Leave-confirmation" below
- [ ] Reopening a document after a clean close restores the exact content, caret position and scroll position from the last autosave
- [ ] The stored record carries the caret position and scroll position alongside the model, so both survive a reload without being reconstructed
- [ ] If the tab crashes or is force-closed mid-edit, reopening the document shows a bar "Recovered unsaved changes from <time>" with "Keep" and "Discard" actions — the crash is detected by the draft record being newer than the committed one, per "Persisted state" below
- [ ] "Keep" promotes the draft to committed and the bar disappears
- [ ] "Discard" deletes the draft, loads the committed record — the last clean state as defined in "Persisted state" below — and the bar disappears
- [ ] Autosave writes do not block typing: input remains responsive during a save of a 200-page document
- [ ] Two tabs open on the same document detect the conflict; the second tab shows "This document is open in another tab" and opens read-only until the other tab is closed
- [ ] Storage quota errors surface as "Not saved — browser storage is full" with a link to download the file, and replace SPEC-003's "Changes will not be saved in this browser" banner rather than stacking with it — see "Storage failure" below
- [ ] The storage-failure state clears once a write succeeds again: the message disappears and the save state returns to "All changes saved"

## Out of scope
Server-side autosave. Saving back to the user's original file on disk. Version
history (SPEC-040). Merge of concurrent edits (SPEC-044).

## Reference

### Storage
"Local storage" here means client-side storage generally, not the
`window.localStorage` API. SPEC-003 is normative: documents live in IndexedDB
keyed by document id. That store already exists — `src/lib/idb.ts` behind the
write-through map in `src/lib/documentStore.ts` — and autosave extends it rather
than adding a second store beside it.

`window.localStorage` is ruled out twice over: its ~5 MB budget cannot hold a
200-page document, and it is synchronous on the main thread, so using it would
violate the "autosave writes do not block typing" criterion by construction.

`StoredDocumentRecord` currently holds `id`, `format`, `lastModified` and
`model`; this spec adds the caret and scroll position and the committed/draft
split below.

### Persisted state — committed and draft
The store keeps two records per document:

- **committed** — the last clean state. This is what "Discard" restores.
- **draft** — the autosave journal, written on the 2 s / 30 s cadence.

A state becomes committed at exactly four points:

1. the document is created or imported (so a committed record always exists,
   including a crash during the first session — "Discard" is never left with
   nothing to restore);
2. the document is downloaded as `.docx` or `.odt` (SPEC-006 / SPEC-007) — what
   SPEC-003 already calls an explicit save;
3. the tab closes cleanly and the final flush completes;
4. the user clicks "Keep" on the recovery bar.

Nothing else promotes a draft. On open, a draft newer than the committed record
means the previous session ended without a clean close, which is the crash
signal the recovery bar is keyed to; a draft that is absent or not newer means
the last close was clean and the document opens with no bar.

This split is what makes "the last cleanly saved state" a real thing to restore.
With a single write-through record the same bytes would be both the restored
document and the "unsaved changes" being recovered, and "Discard" would have no
target — server-side save and disk save are both out of scope, and a download
emits a `Blob` without touching the store.

### Leave-confirmation
Because a draft is at most 2 seconds behind the caret, the leave-confirmation
prompt fires only while a write is in flight or after one has failed. It is
therefore a storage-failure signal in practice, not an unsaved-work signal — the
expected behaviour is that it almost never appears, and appears persistently
once storage is full. Do not widen the definition of "unsaved" to make the
prompt fire more often.

### Storage failure
SPEC-003 shows a persistent banner "Changes will not be saved in this browser"
when the store is unavailable, and its `isStorageAvailable()` latches one-way and
never recovers. This spec supersedes both: quota and write failures show
"Not saved — browser storage is full" with a download link, in the app bar's
save state rather than as a second banner, and the state clears when a write
succeeds again. SPEC-003's banner remains only for the case it was written for —
storage unavailable at startup, e.g. IndexedDB blocked by the browser.
