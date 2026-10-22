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
- [ ] The app snapshots the document periodically during editing and at every explicit save/download, storing snapshots locally — "periodically" means the fixed cadence in "Snapshot cadence" below
- [ ] "File → Version history" opens a pane listing versions with timestamp, author identity and word-count delta, newest first
- [ ] Selecting a version shows it read-only in the canvas with additions and deletions since the previous version highlighted
- [ ] "Restore this version" replaces the current content and is itself undoable and recorded as a new version
- [ ] "Name this version" labels a version, and named versions can be filtered to on their own
- [ ] "Download this version" exports the selected snapshot as `.docx` or `.odt` without changing the current document
- [ ] Snapshots are stored as deltas over periodic full baselines, and every baseline is written because one of the four triggers in "Baselines and deltas" below fired — never at implementation discretion
- [ ] Image bytes and the verbatim parts SPEC-008 preserves are stored once per document and shared across versions by content hash, not copied into every baseline — see "Shared blobs" below
- [ ] History for one document stays within a 50 MB budget, and a 100-page document edited across a week of working days fits inside it without any pruning — see "The 50 MB budget" below
- [ ] When the budget is hit, the oldest unnamed versions are pruned first
- [ ] Pruning never removes named versions or the last 10 versions, and never leaves a retained version unreconstructable — see "Pruning and chain integrity" below
- [ ] When nothing is prunable and the budget is still exceeded, snapshotting stops with a visible notice while the document keeps autosaving — see "When the budget cannot be met" below
- [ ] Reconstructing any retained version takes under a second for a 100-page document at any history depth, because the delta chain is capped
- [ ] Closing and reopening the document preserves the full history

## Out of scope
Server-side or cross-device history. Restoring a single paragraph from an old
version. Branching. Diff between two arbitrary versions. Sharing blobs between
different documents' histories.

## Reference

### Snapshot cadence
A version is taken:

1. every 5 minutes of editing, and only if the model changed since the last
   version — the timer runs off the same activity signal as the autosave
   scheduler (`src/lib/autosaveScheduler.ts`), one order of magnitude coarser
   than its 2 s / 30 s draft cadence;
2. at each of SPEC-010's commit points that the user drives — download as
   `.docx`/`.odt`, and a clean tab close;
3. immediately before and after a restore, so restoring is itself a point in
   history and is undoable against a recorded state.

This cadence is not a tuning knob. It is the input to the storage arithmetic
below: change it and the constants stop adding up to 50 MB.

### Where history lives
Version records go in a **second IndexedDB object store**, `versions`, beside
the existing `documents` store — a `DB_VERSION` bump and one more
`onupgradeneeded` branch in `src/lib/idb.ts`.

They must not go in `documents`. `initLibrary` hydrates the library with a
single `getAll()` over that store and skips drafts by key prefix — but only
after reading them — so 50 MB of version payloads per document would be read
into memory at every startup.

Two record shapes:

- `versions:<docId>` — one index record per document holding, per version: id,
  timestamp, author, word count, name, chain kind (baseline or delta), parent
  id and encoded size. Small; loaded when the pane opens. It is enough to
  render the list, compute word-count deltas and run the budget arithmetic
  without touching a single payload.
- `version:<docId>:<versionId>` — the payload, baseline or delta. Loaded only
  when a version is selected, restored, downloaded, or replayed as part of a
  chain.

Payloads are stored as JSON strings rather than structured-cloned objects, so
"encoded size" is a number the code actually has, and the budget is enforced on
a measurement rather than an estimate.

### Baselines and deltas
A version record is one of two kinds:

- **baseline** — the whole model, standalone.
- **delta** — a patch against the *immediately preceding* version. Versions
  therefore form a chain: baseline → delta → delta → … Reconstructing version
  *n* means loading its chain's baseline and applying every delta up to *n*.

Deltas are block-granular. Every command returns a fresh model that leaves
untouched blocks at the same array *reference* (see the header comments in
`src/lib/commands.ts` and `src/lib/trackChangesDiff.ts`), so a delta is computed
by walking sections and blocks with reference equality and recording only the
spans that differ, each changed block carried in full. There is no
intra-paragraph patching: a paragraph is small, and `applyTrackChanges` — the
diff engine fine enough to produce one — exists for display (the highlighting
in AC-3), not for storage.

**A snapshot is written as a baseline when any of the following holds, checked
in this order, and as a delta otherwise:**

1. **No chain to extend** — it is the document's first version, or the
   preceding version's payload is missing or unreadable.
2. **Chain cap** — the chain already holds 50 deltas. This is the
   reconstruction bound: at most 51 records are ever read and applied.
3. **Chain outweighs a copy** — the delta bytes accumulated since the chain's
   baseline have reached that baseline's own encoded size. Past that point the
   chain costs more to store than a fresh copy would, so the copy is written.
4. **Wholesale rewrite** — the delta just computed encodes to 50% or more of
   the chain's baseline. Select-all-and-paste, replace-all and "Restore this
   version" land here; storing them as deltas saves nothing and lengthens every
   later reconstruction.

Rule 4 requires encoding the delta and then discarding it when the baseline
wins. That is one wasted encode on a rare path, and it is what stops a
rewrite from being stored at nearly full size *and* forcing a long replay.

Nothing here is "as needed". A baseline written at any other point is a bug,
and the constants — 50 deltas, 100%, 50% — are fixed by the budget below rather
than tuned per document or per device.

Deltas chain against their predecessor rather than fanning out from the
baseline: fanning makes each delta grow with distance from the baseline, which
costs more over a long editing session than the capped replay costs. Reverse
deltas (newest stored whole, older ones as backward patches) were the other
candidate and are rejected because the current model already lives in the
committed record from SPEC-010 — the newest version needs no special treatment.

### Shared blobs
`media`, `opaque` and `sourceXml` — image bytes (SPEC-024) and the unmodelled
package parts SPEC-008 keeps verbatim — are the bulk of most real documents and
change rarely. Baselines do not copy them. Each is written once under
`blob:<docId>:<sha256>` (hashed with `crypto.subtle.digest`, no new dependency)
and both baselines and deltas reference the hash. A blob is deleted when the
last version referencing it is pruned and the current document no longer
references it. Keys are scoped per document so pruning one document's history
never has to reason about another's references.

Without this, a 12 MB document of photographs would exhaust the whole budget in
four baselines, and history would be useless exactly where it matters most.

### The 50 MB budget
History for one document may occupy 50 MB — the same ceiling SPEC-009 puts on a
single file the app will open, so a document's history can never cost more than
re-opening the document itself. What counts: every version record for that
document, plus every blob referenced only by history.

The worked case behind the constants — a 100-page document edited across a week
of working days:

| Quantity | Value |
| --- | --- |
| Model JSON, 100 pages, blobs excluded | ~2 MB |
| Editing time | 40 h (5 × 8 h, continuous — a deliberate worst case) |
| Versions at one per 5 min | ~480 |
| Baselines at a 50-delta cap | ~10 → ~20 MB |
| Deltas, ~20 KB each (a 5-minute burst touches a handful of paragraphs) | ~10 MB |
| **Total** | **~30 MB** |

So the AC holds with roughly 20 MB of headroom for the size triggers firing on
large edits, and that week never prunes. The two size triggers bind at
different times by design: with ~20 KB deltas the cap (rule 2) fires first,
while a session of large structural edits trips rule 3 or rule 4 well before
50 deltas accumulate.

### Pruning and chain integrity
The budget is checked after each snapshot write. While the document's history
exceeds 50 MB:

1. Take the oldest **prunable** version — unnamed, and not among the newest 10.
   If there is none, stop; see below.
2. If a retained version's chain passes through it, re-encode the oldest such
   retained version as a baseline first, so nothing retained depends on
   anything older.
3. Delete the version record, and any blob whose last reference went with it.
4. Re-measure.

Step 2 can raise usage before it lowers it — a 2 MB baseline replacing a 20 KB
delta — so the loop must re-measure after every step rather than pre-computing
how many versions to drop. It terminates because each iteration deletes exactly
one version and creates none.

### When the budget cannot be met
If nothing is prunable and the total is still over budget — ten named 5 MB
versions, say — history stops rather than eating into the document. No new
snapshots are written, the pane shows "Version history paused — 50 MB used.
Delete a version, or remove a name, to resume", and the document keeps
autosaving normally. Removing a name is what makes a version prunable again.

Likewise, a quota error on a history write abandons that write; history never
frees space by evicting the document's own committed or draft records.
SPEC-010's "Not saved — browser storage is full" state belongs to the document,
and history must never be the reason it appears.
