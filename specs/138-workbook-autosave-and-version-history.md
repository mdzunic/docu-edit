---
id: SPEC-138
title: Autosave, drafts and version history for workbooks
status: draft
area: io
complexity: normal
story_points: 8
depends_on: [SPEC-010, SPEC-040, SPEC-103, SPEC-121, SPEC-122, SPEC-129, SPEC-137]
---
## Why
Type into a workbook for an hour, close the tab mid-edit, and SPEC-121's
write-through has kept every cell — but there is nothing between "the last
gesture" and "nothing": no journal to recover a crashed session from, no
committed state for "Discard" to return to, and no way to see what the
budget looked like on Tuesday. SPEC-116 made sure of it on purpose:
`saveDraft` and `getOrCreateDocument` throw on a spreadsheet id and
`useAutosave.beginSession` bails, because the draft machinery could not yet
tell a workbook from a document and silently overwriting one is the
unforgivable failure.

The machinery itself is no longer the problem. SPEC-010 built the
draft/committed split and the recovery bar; SPEC-040 the version chain;
SPEC-103 generalised both beyond text (`DraftState`, the `opaque` split,
`src/lib/core/versionChain.ts`, a history panel driven by a controller).
This spec generalises them once more, for the fourth kind, and the one thing
that is genuinely new is **rows**. A workbook's content does not live in its
library record: SPEC-137 stores each sheet's `rows` as 512-row
`sheetrows:` chunk records and the retained sources as one `sheetsource:`
record, each written only when its reference changes. A draft that wrote
those records would already have destroyed the state "Discard" restores —
exactly the reasoning behind `DRAFT_MEDIA_KEY_PREFIX`.
And a version delta at the text side's granularity would carry a whole
sheet whenever one cell changed, so workbook deltas are row windows over row
identity, which SPEC-116's copy-on-write invariant makes sound.

Decks skipped the two-tab edit lock (SPEC-103's out of scope). Workbooks
cannot: two tabs editing one workbook would race on the same row chunk
records, so the lock SPEC-010 gave the word processor (`core/editLock.ts`)
comes with this spec. **This file is the normative source for the workbook
draft record (`DraftState.sheetView`, the draft chunk and source records
and their manifest), the workbook version payloads and their row-window deltas,
workbook restore semantics, and the workbook edit lock.** It owns no model
concept: SPEC-116 owns the model and the row split, SPEC-121 the writers and
the undo stack, SPEC-122 the sheet writer.

## Acceptance criteria
- [ ] The spreadsheet guard SPEC-116 put in `saveDraft` and the `isSpreadsheetModel` bail in `useAutosave.beginSession` are removed, while `getOrCreateDocument`'s spreadsheet guard **stays** (creating a text document over a workbook's id still throws); `replaceSpreadsheetModel(id, model: SpreadsheetModel)` joins `replaceDocumentModel`/`replacePresentationModel` over SPEC-103's private `replaceModel`, and a new `replaceStoredModel(id, model: StoredModel)` dispatches on `kind`, so `beginSession` and `discardRecovered` make one call for every kind instead of growing a third branch — `src/composables/useAutosave.ts` stays within its `src/composableSize.spec.ts` entry (404 today), extracting rather than raising it
- [ ] `DraftState` gains `sheetView?: SheetViewState` (`{ sheetId; active: CellRef; topRow; leftCol }`, declared in `src/lib/spreadsheet/selection.ts`), `StoredDocumentRecord` and `DraftRecord` gain `sheetView?`, `UseAutosaveOptions` gains `getSheetView?: () => SheetViewState | undefined` beside SPEC-103's `getSlideIndex?`, and `RestoredState` gains `sheetView?`; scroll is restored as the top-left visible row and column, never pixels, per "The workbook draft record"
- [ ] Workbook edits stop writing the committed record through: the last step of SPEC-121's edit pipeline, `deps.onApplied()` (run after every `commit`, `undo` and `redo`), calls `noteEdit()` where it called `touchDocument(model.id)`, putting edits on the SPEC-010 2 s / 30 s draft cadence, and a workbook is committed at exactly SPEC-010's four points — create/import, download (`spreadsheet-download` and "Download a copy as…"), clean close (`close-document` through `flushForClose`, and tab close) and "Keep"; opening a workbook, switching tabs, moving the selection and SPEC-120's open-time recalculation never call `noteEdit()`
- [ ] Draft rows never touch committed rows: each draft chunk record shadows one SPEC-137 committed chunk record under the key `draftRowsKey(committedKey) = "draft" + committedKey` (`draftsheetrows:<docId>:<sheetId>:<chunk4>`, `DRAFT_SHEET_ROWS_KEY_PREFIX = "draftsheetrows:"`, in the same `documents` store), and `persistDraftSheetRows(docId, model)` applies `persistDraftMedia`'s latch rules per chunk, against SPEC-137's per-chunk committed latch: reference-equal to the last committed write → no draft record and any stale draft copy deleted; reference-equal to the last draft write → skipped; otherwise written, the latch set before the write and cleared on failure — so no draft write can ever reach a committed chunk record
- [ ] `DraftRecord` gains `rowKeys?: Record<string, string[]>`, the manifest naming, per sheet id, the ordered record keys whose `rows` concatenate to that sheet's rows (a committed key for a unit unchanged since the last commit, its draft key otherwise); `loadDraft` rejoins every sheet from its manifest, and a manifest naming a record that is missing or unreadable makes `loadDraft` return `undefined` — a half-rejoined workbook is never offered for recovery
- [ ] A workbook draft shell carries no bulk: `rows` are `[]` (they ride the chunk records), `opaque` is `[]` and no `sourceXml` is inline, as in SPEC-137's shell; the retained sources — `opaque`, each sheet's `sourceXml`, `odsContentEnvelope` — stay in SPEC-137's committed `sheetsource:<docId>` record, which a draft never writes, while they are reference-equal to what that record last stored; once one differs, the draft writes a draft source record `draftsheetsource:<docId>` (`DRAFT_SHEET_SOURCE_KEY_PREFIX`, SPEC-137's `WorkbookSourceRecord` shape) and `loadDraft` rejoins from it, else from the committed record; they are tracked by reference identity, not assumed immutable, because SPEC-124/125's retained-reference rewrite replaces them, per "The workbook draft record"
- [ ] Draft cost is asserted by operation counts at the `idb` seam, not timings, on a three-sheet workbook: the first tick after editing one cell of sheet 2 writes the draft shell and that cell's draft chunk record; a second edit in the same chunk writes the shell and that chunk again; an edit to sheet 3 after it writes the shell and sheet 3's chunk only; a structural edit that rewrites a retained source also writes `draftsheetsource:<docId>`; a commit writes the changed committed chunks, the source record when changed and the slim shell in one SPEC-137 `writeBatch`, then deletes every draft record of the document
- [ ] Draft chunk records of sheets deleted mid-session never leak: `deleteDraft`, `commitDocument` and `deleteDocument` delete the document's whole draft range in one operation — SPEC-137's `deleteRange("documents", "draftsheetrows:<docId>:", "draftsheetrows:<docId>:\uffff")` plus `draftsheetsource:<docId>` — and SPEC-137's `loadAllRecords` key ranges also exclude the `draftsheetrows:` and `draftsheetsource:` prefixes, so `initLibrary` reads no draft chunk or source record, as it reads no committed one
- [ ] `loadDraft` runs `upgradeSpreadsheetModel` (SPEC-116) on a recovered workbook as on every other read path, so a draft written before a reader-version bump recovers upgraded; recovered and discarded models are both re-identified (`model.id = id`) and handed to the view through `replaceStoredModel`
- [ ] Reopening a workbook whose draft is newer than its committed record shows SPEC-103's `RecoveryBar.vue` (`recovery-bar`, `recovery-keep`, `recovery-discard`, copy "Recovered unsaved changes from <time>") over the recovered cells with the drafted sheet, active cell and scroll position restored; "Keep" commits and hides the bar; "Discard" deletes the draft, restores the committed model and its `sheetView`, and calls `useSpreadsheetUndo.reset()` (the stack's by-reference snapshots point at rows that no longer exist)
- [ ] `SpreadsheetAppBar.vue` shows the save state through `SaveStateBadge.vue` (`save-state`, `data-save-state`, class `sheet-save-state`) with the words of `src/lib/core/saveState.ts`'s `saveStateText` — "Saving…", "All changes saved", "Not saved — browser storage is full", and `SAVE_STATE_RECOVERED` "Not saved yet" while the recovery bar is open — and on `error` a `sheet-save-state-download` link running the workbook download; any write-through readout SPEC-121 added is replaced, not doubled
- [ ] A second tab on the same workbook opens read-only: `acquireEditLock(id, onStatus)` from `src/lib/core/editLock.ts` runs on open; while the lock is `"busy"` a `sheet-conflict-bar` reads "This workbook is open in another tab", the grid carries `aria-readonly="true"`, the cell editor, every edit command and every history action that writes are disabled, the save badge is hidden and no autosave session starts; when the other tab closes, the queued lock grants, `beginSession` adopts the other tab's commit through `refreshDocumentFromStorage` (rows rejoined), the undo stack resets and editing enables — no reload; the lock is released on Close and on unmount
- [ ] Workbook versions live in `src/lib/spreadsheetVersions.ts` (top level, beside `documentStore.ts` and SPEC-103's `presentationVersions.ts` — `lib/spreadsheet` may not import `./idb`) in the **existing** `versions` object store under the same `versions:`/`version:`/`blob:` key prefixes with **no `DB_VERSION` bump**; payloads are `WorkbookBaselinePayload`/`WorkbookDeltaPayload` per "Workbook version payloads", and `VersionIndexEntry` gains `sheets?: number` and `cells?: number` (additive, as SPEC-103 added `slides?`)
- [ ] Row deltas are windows over row identity: each sheet's rows are cut into bands of `VERSION_ROW_BAND = 4096` row indices and `windowOf` (`src/lib/core/arrayWindow.ts`) trims each band's reference-equal prefix and suffix, so an edit at row 5 and one at row 900,000 cost two one-row windows, not a 900,000-row window; a sheet whose `rows` array is reference-equal to its predecessor's contributes nothing, and a structural edit that re-indexes every row lands, correctly, in SPEC-040's wholesale-rewrite rule as a baseline
- [ ] `opaque` parts, `media` parts, each sheet's `sourceXml` and the `odsContentEnvelope` are content-addressed blobs (`blob:<docId>:<sha256>`, `sha256Hex` from `core/versionChain.ts`), each hashed at most once per reference (a per-`docId:sheetId` cache compares the string by identity first) and carried by a delta only when its reference — hence possibly its hash — changed since the parent version, which happens when SPEC-124/125's retained-reference rewrite replaces it; a 30 MB `.ods` table's source therefore costs one hash and one blob per distinct content, not per version
- [ ] Snapshots fire on SPEC-040's cadence and triggers: `createVersionCadenceTimer` armed by `noteEdit()` (5 min of editing, only when the workbook changed), each download, clean close via `onCleanClose`, and `pre-restore`/`post-restore`; the `keepVersionHistory` preference gates `cadence`/`commit` only
- [ ] Over budget a workbook degrades to no history and **never fails the edit**: a workbook with more than `WORKBOOK_HISTORY_MAX_CELLS = 250_000` stored cells is refused before any encode, and a smaller one whose blob set plus baseline would exceed `HISTORY_BUDGET_BYTES` is refused before any write; either latches `paused` with `WORKBOOK_HISTORY_TOO_LARGE_MESSAGE` in `history-paused-notice`; normal exhaustion uses core's `HISTORY_PAUSED_MESSAGE`; a history failure never calls `markStorageUnavailable`, never touches the workbook's `documents`-store records and never blocks a commit
- [ ] File ▸ "Version history" (`sheet-menu-version-history`) and an app-bar toggle (`sheet-history-toggle`) toggle the `"history"` pane of SPEC-129's sheet side panel (`useSheetSidePanel().toggle("history")`; this spec widens `SheetPaneId` with `"history"` and adds its `SHEET_PANES` entry), whose body is `sheet-history-panel`, built on the product-neutral `src/components/HistoryPanel.vue` extracted from SPEC-103's `DeckHistoryPanel.vue` with every `history-*` testid unchanged; it lists versions newest first with timestamp, author (`preferences.authorName`), "N sheets · M cells" and a cell-count delta ("+12 cells", "−3 cells", "no change in cell count"), and offers Name/Remove name, named-only, Delete, Preview, Restore and Download (`.xlsx`/`.ods` per `model.format` + `packageVariant`, through `spreadsheetDownload`)
- [ ] Preview renders the reconstructed workbook in the grid read-only (`aria-readonly="true"`, its own sheet tabs, opening on the version's active sheet) under `history-preview-banner` "Previewing an earlier version, read-only" with `history-exit-preview` "Back to current"; there is no cell diff highlighting, and exiting returns the live workbook, selection and undo stack untouched
- [ ] Restore is a **replace**, pinned in "Restore": `pre-restore` snapshot → reconstruct → `upgradeSpreadsheetModel` → `replaceStoredModel` → `useSpreadsheetUndo.reset()` → commit → `post-restore` snapshot → a `useToasts` toast "Version restored" whose action "Undo restore" (`toast-action`) restores the `pre-restore` version the same way; Ctrl+Z after a restore does nothing (the stack is empty), and the panel says so: "Restoring replaces the workbook and clears undo — the version saved just before is at the top of the list"
- [ ] Restore fidelity: open `mini-book.xlsx`, snapshot, edit cells on two sheets and add a sheet, restore the snapshot, download — every part of the download is `equalBytes` (`src/lib/core/equality.ts`, part-level after `unzip`) with a download taken at snapshot time, under a fixed clock (`setClock`, `src/lib/core/clock.ts`); the same holds for `mini-book.ods`, and a no-edit workbook snapshotted and restored still downloads byte-identical per part to its source
- [ ] `SpreadsheetView.vue` gains exactly two wiring calls — `useWorkbookSaving({ … })` (autosave, recovery, lock) and `useWorkbookHistory({ … })` (the history controller), each a `src/views/spreadsheet/` composable within `composableSize.spec.ts`'s default 400 lines — plus template hookup; `SpreadsheetView.size.spec.ts`'s budget never rises (extract first if the hookup would cross it)
- [ ] Tests: `src/lib/documentStore.autosave.spec.ts` (lifted guard, workbook draft/commit/recover, the draft-chunk latch in all three states, the manifest and its missing-record refusal, retained-source stripping and the draft source record, operation counts per tick, range cleanup on commit/discard/delete — existing expectations untouched), `src/lib/spreadsheetVersions.spec.ts` (new), `src/lib/spreadsheet/roundTrip.spec.ts` (restore fidelity), `src/views/spreadsheet/useWorkbookSaving.spec.ts` and `useWorkbookHistory.spec.ts` (new), `src/views/SpreadsheetView.spec.ts`, `SpreadsheetView.size.spec.ts`
- [ ] e2e (`e2e/SPEC-138.spec.ts`): open `mini-book.xlsx` → type into `B3` on "Data" → reload → the value is there with no bar → inject a synthetic newer `draft:<id>` shell (with its `rowKeys` manifest) plus one draft chunk record (keyed by `draftRowsKey`) via `page.evaluate` + `indexedDB` (SPEC-010's technique) → reopen: `recovery-bar` visible, the drafted value shown and `name-box` reading the drafted active cell → Discard restores the committed value → edit, download (a commit-point snapshot) → open `sheet-history-panel`, name the version, preview it, exit, edit again, restore the named version → the grid shows the restored value and the list gains the restore's two snapshots → open the same `/sheets/:id` in a second page: `sheet-conflict-bar` visible and typing does nothing → close the first page → the second page's bar disappears and typing works

## Out of scope
Server-side or cross-device drafts and history. Sharing blobs between two
workbooks' histories. Branching, comparing two arbitrary versions, cell-level
diff highlighting in preview, and restoring one sheet or one range from an
old version (the unit of restore is the workbook). Persisting the undo stack
(SPEC-121 excluded it; unchanged). Autosaving back to the user's file on
disk. Merging concurrent edits from two tabs (SPEC-044's territory; here the
second tab is read-only). Compressing history payloads. Version history for
CSV/TSV workbooks beyond what their model already is (SPEC-126's model is a
`SpreadsheetModel` and gets exactly this). Changing SPEC-137's chunking or
residency — this spec drafts against them as they are.

## Reference

### Dependencies restated
- `specs/010-autosave-and-crash-recovery.md` — two records per document:
  **committed** (the last clean state "Discard" restores) and **draft** (the
  journal on the 2 s debounce / 30 s deadline / 5 s retry cadence of
  `createAutosaveScheduler`, `src/lib/core/autosaveScheduler.ts`). Committed
  only at create/import, download, clean close and "Keep"; recovery when
  `draft.savedAt > record.lastModified`. `DRAFT_KEY_PREFIX = "draft:"`,
  `MEDIA_KEY_PREFIX = "media:"`, `DRAFT_MEDIA_KEY_PREFIX = "draftmedia:"` in
  `src/lib/documentStore.ts`, whose comment states the rule this spec
  copies: an autosave tick must never overwrite the payload the committed
  state depends on. `putRecordYielding` awaits a macrotask first because
  `IDBObjectStore.put()` structured-clones synchronously. The two-tab rule:
  "This document is open in another tab", read-only until the other closes.
- `specs/040-version-history-and-restore.md` — baseline or delta against the
  immediately preceding version; the four baseline triggers in order (no
  chain, `CHAIN_DELTA_CAP` 50 deltas, chain outweighs a copy, delta ≥
  `WHOLESALE_REWRITE_RATIO` 0.5 of the baseline); reference-counted blobs at
  `blob:<docId>:<sha256>`; 50 MB per document; prune the oldest unnamed
  version outside the newest `RETAINED_RECENT_COUNT` 10, re-baselining any
  retained version that chained through it, re-measuring each iteration.
- `specs/103-presentation-autosave-and-version-history.md` (not built when
  this was written; it lands first) — `DraftState { caret: CaretPosition |
  null; scroll: number; slideIndex?: number }`; `saveDraft(id, state)` /
  `commitDocument(id, state)`; `DraftRecord.model: StoredModel`;
  `replacePresentationModel` over a private `replaceModel`;
  `OPAQUE_KEY_PREFIX = "opaque:"`, `DRAFT_OPAQUE_KEY_PREFIX = "draftopaque:"`,
  `persistOpaque`/`persistDraftOpaque`/`stripOpaque`/`rejoinOpaque` with the
  media latch rules; `useAutosave` serving both products
  (`RestoredState.model: StoredModel`, `getSlideIndex?`);
  `src/components/RecoveryBar.vue`; `src/lib/core/versionChain.ts` with
  `VersionChainKind`, `SnapshotTrigger`, `VersionIndexEntry`,
  `VersionsIndexRecord`, `VersionBlobRecord`, `ChainInfo`,
  `gatherChainInfo`, `decideBaseline`, `measureBytes`, `sha256Hex`, the
  constants and `HISTORY_PAUSED_MESSAGE`; `presentationVersions.ts`
  (`recordDeckVersion`, `reconstructDeckVersion`); `DeckHistoryPanel.vue`,
  controller-driven, with testids `history-list`, `history-empty`,
  `history-version`, `history-name`, `history-name-input`,
  `history-name-submit`, `history-remove-name`, `history-delete`,
  `history-named-only`, `history-preview`, `history-preview-banner`,
  `history-exit-preview`, `history-restore`, `history-download`,
  `history-paused-notice`. **If SPEC-103 landed any of these under another
  name, this spec uses the landed name**; the extraction in "View wiring"
  starts from whichever component renders the `history-*` testids.
- `specs/116-workbooks-in-the-library.md` — `SpreadsheetModel { kind:
  "spreadsheet"; id; title; format; packageVariant?; readerVersion; sheets;
  activeSheetIndex; dateSystem; odsNullDate?; odsContentEnvelope?: { prefix;
  suffix }; opaque: OpaquePart[]; preview?; properties?; media? }`, `Sheet {
  id; name; visibility; kind; rows: Row[]; partName?; sourceXml? }`.
  **Copy-on-write rows**: no command mutates a `Row`, a `Cell` or a sheet's
  `rows` array; untouched rows keep their identity. **The source is the
  splice base**: `opaque`, `sourceXml` and the envelope are kept for the
  record's life — replaced, never mutated, by SPEC-124/125's
  retained-reference rewrite. `upgradeSpreadsheetModel(model)` runs on
  every read path. `spreadsheetDownload(model, targetFormat = model.format)`
  → `Promise<{ bytes; fileName; mimeType; notices: string[] }>` (as SPEC-121
  extends it).
- `specs/137-large-workbooks-performance.md` — the storage layout this spec
  drafts against, all in the `documents` store (no new object store, no
  `DB_VERSION` bump): committed rows as `SHEET_ROWS_CHUNK = 512`-row chunk
  records `sheetrows:<docId>:<sheetId>:<chunk4>` (`{ id, docId, sheetId,
  chunk, rows }`), the retained sources as one `sheetsource:<docId>` record
  (`WorkbookSourceRecord { opaque; odsContentEnvelope?; sheetSources }`), a
  slim shell (no rows, `opaque: []`, no `sourceXml`) carrying the
  `sheetStore { chunks; storedBytes }` manifest; a persist rewrites a chunk
  only when its row slice differs by identity from the last persisted slice
  (the per-chunk latch) and the source record only when a source reference
  changed, all in one `writeBatch`; `core/idb.ts` `getAllInRange`,
  `deleteRange`, `writeBatch`; `loadAllRecords` reads only the key ranges
  outside the chunk and source prefixes; residency — `getSpreadsheetShell`,
  `getSpreadsheet` (resident only), `loadSpreadsheet`, `releaseSpreadsheet`.
- `specs/129-…` — `useSheetSidePanel()` (`src/views/spreadsheet/
  useSheetSidePanel.ts`): `SheetPaneId` (widened per pane), `SHEET_PANES`,
  `{ activePane; open(id); close(); toggle(id); side; widthPx; collapsed;
  tabs }`, persisted under `docu-edit:sheet-side-panel`, rendered by
  `SidePanel.vue` with the active pane's component in its default slot.
- `specs/117-…` — `SheetSelection { active: CellRef; ranges; activeRangeIndex }`
  and `Sheet.selection?` (view state from the file, never marks dirty);
  `GridMetrics.rowTop(r)`/`colLeft(c)` map indices to pixels.
- `specs/121-create-workbook-and-sheet-management.md` —
  `useSpreadsheetUndo` `{ record(label, scope: UndoScope), commit(),
  undo(), redo(), reset(), canUndo, canRedo }`, per open view, never
  persisted; a snapshot holds every sheet (fields in
  `SHEET_BY_REFERENCE_FIELDS` — `rows`, `sourceXml` — by reference) plus
  `WORKBOOK_SNAPSHOT_FIELDS` (`activeSheetIndex`, `structureDirty`,
  `styles`, `opaque`, `odsContentEnvelope`, and whatever later specs
  append) by reference. The edit pipeline ends in
  `deps.onApplied()` = `triggerRef` + `touchDocument(model.id)` — the
  write-through AC-3 replaces. `Sheet.dirty?`/`structureDirty?` drive the
  writers' rebuild; a clean workbook re-emits every part byte-identical;
  `writeXlsx(model, { modified?, fullCalcOnLoad? })`/`writeOds(model, {
  modified? })`; downloads run through `useWorkbookDownload.ts`. `styles`
  is append-only for the life of a workbook (SPEC-122).
- **Equality**: `equalBytes` from `src/lib/core/equality.ts`, part-level
  after `unzip`; workbook-model comparisons are field-level `toEqual` (the
  text side's `equalSemantic` is `DocumentModel`-typed).

### The workbook draft record
```ts
// src/lib/spreadsheet/selection.ts (type only)
export interface SheetViewState {
  sheetId: string;     // the tab that was showing
  active: CellRef;     // the active cell (the selection collapses to it on restore)
  topRow: number;      // first fully visible row, 0-based
  leftCol: number;     // first fully visible column, 0-based
}
// src/lib/documentStore.ts
export interface DraftState { caret: CaretPosition | null; scroll: number;
  slideIndex?: number; sheetView?: SheetViewState }
export const DRAFT_SHEET_ROWS_KEY_PREFIX = "draftsheetrows:";
export const DRAFT_SHEET_SOURCE_KEY_PREFIX = "draftsheetsource:";
// DraftRecord gains: sheetView?: SheetViewState; rowKeys?: Record<string, string[]>;
//                    sourceKey?: string   (`sheetsource:<docId>` or `draftsheetsource:<docId>`)
// StoredDocumentRecord gains: sheetView?: SheetViewState
```
A workbook's state is `{ caret: null, scroll: 0, sheetView }`. Indices, not
pixels: a pixel offset means nothing once a column is resized or the zoom
changes, and `GridMetrics` turns the indices back into an offset. On open
the view applies, in order: the record's (or draft's) `sheetView`; else the
file's `activeSheetIndex` with that sheet's `Sheet.selection`; else `A1`.
A `sheetView` naming a sheet that no longer exists is ignored.

| Record | Key | Written by | Holds |
| --- | --- | --- | --- |
| committed shell (137) | `<docId>` | `persist`, `commitDocument` | model minus rows and sources, + `sheetStore` |
| committed chunks (137) | `sheetrows:<docId>:<sheetId>:<chunk4>` | `persist`, `commitDocument` | 512 rows of one sheet |
| committed sources (137) | `sheetsource:<docId>` | `persist`, `commitDocument` | `opaque`, envelope, every `sourceXml` |
| draft shell | `draft:<docId>` | `saveDraft` | model minus the same, + `rowKeys`, `sheetView`, `sourceKey` |
| draft chunks | `draftRowsKey(k)` = `draftsheetrows:<docId>:<sheetId>:<chunk4>` | `saveDraft` only | 512 rows of one sheet |
| draft sources | `draftsheetsource:<docId>` (`DRAFT_SHEET_SOURCE_KEY_PREFIX = "draftsheetsource:"`) | `saveDraft` only | as the committed source record |

**Retained sources.** `opaque`, `odsContentEnvelope` and each sheet's
`sourceXml` are kept for the record's life (116's invariant 3) but are
**not** immutable: SPEC-124's and SPEC-125's retained-reference rewrite
replaces them — never mutates them — when a structural edit or a move
re-points a reference inside unmodelled XML, and a SPEC-121 duplicate mints
a new sheet with a source. So every one is tracked by reference identity,
exactly like rows, against SPEC-137's source latch (the references the
committed `sheetsource:<docId>` record last stored). While all are
reference-equal to it the draft shell's `sourceKey` names
`sheetsource:<docId>` and nothing is written; once any differs, the tick
writes the whole current set to `draftsheetsource:<docId>` (under its own
draft latch, so an unchanged set is not rewritten on the next tick) and
`sourceKey` names that record. `loadDraft` rejoins the sources from the
record `sourceKey` names. A large `.ods` table's source is the size of its
XML, so a tick clones it only between a structural edit (or a duplicate)
and the next commit.

**Rows.** `persistDraftSheetRows` walks `model.sheets`, cuts each sheet's
`rows` into SPEC-137's chunks and, per chunk with committed key `k`, does
what `persistDraftMedia` does for media: the chunk's row slice
reference-equal (row by row) to SPEC-137's last committed slice for `k` →
`k` goes into `rowKeys[sheet.id]`, and a draft copy the draft latch still
holds is dropped and `draftRowsKey(k)` deleted; equal to the last draft
slice for `k` → `draftRowsKey(k)` goes into `rowKeys[sheet.id]`, nothing
written; otherwise `{ id: draftRowsKey(k), docId, sheetId, chunk, rows }`
is written through `putRecordYielding`, latch set first, cleared on
failure. `rowKeys[sheetId]` lists the sheet's chunk keys in chunk order and
is the only thing `loadDraft` trusts: it never infers a key, and rejoin is
the concatenation of the listed records' `rows`. Any `persist` of the
record (a rename, a star) still writes the live rows as committed, as it
writes the live text model today; the next tick then finds reference
equality and drops its draft copy. A tick after one cell edit costs one
512-row chunk.

**Residency.** SPEC-137's rule applies: a draft, commit or version capture
runs only on a resident workbook, and `useWorkbookSaving.flushForClose`
settles before `useWorkbookLoad` calls `releaseSpreadsheet`. Version
history operates on the in-memory `rows` arrays and is indifferent to
chunking.

### Workbook version payloads (`src/lib/spreadsheetVersions.ts`)
```ts
interface BlobRef { hash: string }                       // bytes in blob:<docId>:<hash>
// OpaqueRef { name; hash } and MediaRef { mediaId; contentType; hash } are SPEC-103's.
type SheetMeta = Omit<Sheet, "rows" | "sourceXml">;     // name, visibility, columns, merges, pivots…, dirty
interface SheetSnapshot { id: string; meta: SheetMeta; rows: Row[]; sourceXml?: BlobRef }
// WorkbookFields: every top-level SpreadsheetModel field except id, sheets,
// opaque, media, preview, odsContentEnvelope — title, format, packageVariant,
// readerVersion, activeSheetIndex, dateSystem, odsNullDate, styles, theme,
// defaultStyleId, properties, structureDirty and whatever later specs add.
export interface WorkbookBaselinePayload { kind: "workbook-baseline";
  sheets: SheetSnapshot[]; fields: WorkbookFields;
  opaque: OpaqueRef[]; media: MediaRef[]; envelope?: BlobRef }
export interface RowWindow { band: number; prefixLen: number; suffixLen: number; replaced: Row[] }
export type SheetDelta =
  | { id: string; kind: "added"; snapshot: SheetSnapshot }
  | { id: string; kind: "changed"; meta?: SheetMeta; rows?: RowWindow[]; sourceXml?: BlobRef };
export interface WorkbookDeltaPayload { kind: "workbook-delta"; parentId: string;
  order: string[];                    // every sheet id, in order; absence = deleted
  sheets: SheetDelta[]; fields: Partial<WorkbookFields>;
  opaque?: OpaqueRef[]; media?: MediaRef[]; envelope?: BlobRef }   // present only when changed
export async function recordWorkbookVersion(options): Promise<RecordVersionResult>;
export async function reconstructWorkbookVersion(docId, versionId): Promise<SpreadsheetModel | null>;
export function workbookVersionEntries(index): WorkbookVersionListEntry[];   // + cellDelta
```
**Delta.** For each sheet id in the new order: absent from the previous
version → `added`; `rows` and every `SheetMeta` key reference-equal and the
same `sourceXml` reference → nothing; otherwise `changed`, carrying `meta`
whole when any meta key differs by reference, `sourceXml` when its hash
differs, and `rows` windows when the arrays differ. Windows: partition both
arrays by `band = Math.floor(row.index / VERSION_ROW_BAND)` (binary search
on `index`; rows are sorted), and for each band whose slices differ emit
`windowOf(beforeBand, afterBand)` as `{ band, prefixLen, suffixLen,
replaced: afterMid }`. Applying: per band, `before.slice(0, prefixLen) ++
replaced ++ before.slice(before.length − suffixLen)`, bands re-concatenated
in order. Top-level fields compare by reference and are carried whole —
`styles` included, which only grows. `opaque` and the envelope compare by
reference too and, when replaced (a structural edit or move rewrote a
retained reference), the delta carries their new blob refs — unchanged
parts keep their hashes, so only the rewritten parts add blobs. Like
SPEC-040, this relies on the previous version's model being kept live
(`cachedPrev`); after a reload every row looks new, the delta trips rule 4,
and a baseline is written — a safe degradation, not a special case.

**Recording loop.** The deck spec writes the loop (load index → chain info
→ build delta → `decideBaseline` → ensure blob refs → write payload and
index → enforce budget and prune) in `presentationVersions.ts`. This spec
moves that loop to `src/lib/versionRecorder.ts` behind
`VersionCodec<M> { baseline(model); delta(prev, next, parentId); hydrate(
docId, payload, parent); blobHashes(model); embeddedHashes(payload);
stats(model) }`; `presentationVersions.ts` keeps every export and
delegates (its spec passes unchanged), and `spreadsheetVersions.ts` is the
second codec. The text side's `versionHistory.ts` is not touched.

### The budget, and the honest degrade
`HISTORY_BUDGET_BYTES` = 50 MB per workbook. A stored cell costs about 45
bytes of payload JSON (`{"col":12,"value":12345.67,"styleId":3}` plus its
share of the row wrapper). The worked case — a 20,000-cell workbook from a
400 KB `.xlsx`, edited 8 h/day for five days:

| Quantity | Value |
| --- | --- |
| Blob set (`opaque`, base64, written once) | ≈ 0.6 MB |
| Baseline (20,000 × 45 B) | ≈ 0.9 MB |
| Versions at one per 5 min over 40 h; baselines at the 50-delta cap | ≈ 480; ≈ 10 baselines → 9 MB |
| Deltas ≈ 470 × 5 KB (a few one-row windows each) | ≈ 2.4 MB |
| **Total** | **≈ 12 MB — never prunes** |

At `WORKBOOK_HISTORY_MAX_CELLS = 250_000` a baseline is ≈ 11 MB: three
baselines and their deltas fit, history prunes from the second day and
keeps working. Above it a baseline alone would take a fifth of the budget
and its `JSON.stringify` a visible main-thread pause every five minutes, so
the count — a sum of `row.cells.length`, no encode — refuses first:
```ts
export const WORKBOOK_HISTORY_TOO_LARGE_MESSAGE =
  "Version history off — this workbook is larger than the 50 MB history budget";
```
Encoding is preceded by `yieldToMainThread()`; nothing history does is
awaited on the commit path.

### Restore
A restore swaps every top-level field, every sheet and the retained
sources. `useSpreadsheetUndo` snapshots every sheet plus
`WORKBOOK_SNAPSHOT_FIELDS` — not `title`, `theme`, `properties`,
`readerVersion` or any workbook-level field a later spec forgets to
append — so an undoable restore would revert part of a workbook and leave
the rest from the version, a state that never existed. Restore is therefore a replace, and
its undo is history's own: the `pre-restore` version is written first, the
toast's "Undo restore" action (`toast-action`, for the toast's
`TOAST_DURATION_MS` 4 s; afterwards the panel's own Restore on that entry
does the same) restores it, and both restores appear in the list. The restored
model takes the live record's `id`; it passes `upgradeSpreadsheetModel`
because a version may predate a reader bump. A preview or restore of a
version whose blobs are missing shows "This version can't be opened — part
of it is missing" in the panel and changes nothing.

### Edit lock
`acquireEditLock(docId, onStatus)` names the lock `docu-edit:doc:<docId>`;
ids are unique across kinds, so a workbook and a document never contend.
`useWorkbookSaving` holds `lockState: "pending" | "acquired" | "busy"`;
editing is enabled only at `"acquired"` — while `"pending"` the grid is
read-only for the instant before the lock resolves, as DocumentView does.
A busy tab still renders the workbook it loaded (rows rejoined from the
committed records), so reading works; it never writes a draft, a commit or
a version.

### View wiring and UI
`src/views/spreadsheet/useWorkbookSaving.ts` takes `WorkbookSavingDeps {
model: ShallowRef<SpreadsheetModel | null>; workbookId: Ref<string>;
getSheetView: () => SheetViewState | undefined; applySheetView: (v:
SheetViewState | undefined) => void; resetUndo: () => void; download: () =>
void }` and returns `{ saveState, saveStateText, recovery, recoveryLabel,
keepRecovered, discardRecovered, noteEdit, flushForClose, lockState }`; it
composes `useAutosave` and the lock. SPEC-121's `onApplied` dependency
calls `noteEdit()` in place of `touchDocument` (through the view's wiring,
so the undo composable gains no import). `src/views/spreadsheet/useWorkbookHistory.ts` is the
controller SPEC-103's panel contract expects (`open`, `toggle`, `versions`,
`visibleVersions`, `namedOnly`, `pausedMessage`, `previewModel`, `preview`,
`exitPreview`, `restore`, `undoRestore`, `nameVersion`, `removeName`,
`deleteVersion`, `downloadVersion`, plus `statLabel(entry)` and
`downloadLabels`), its `open`/`toggle` being SPEC-129's
`useSheetSidePanel().open("history")`/`toggle("history")` — the side panel's
own persisted state is the only record of whether the pane is open.

`src/components/HistoryPanel.vue` is SPEC-103's deck panel with the two
deck-specific facts lifted into the controller (`statLabel`, the download
labels) and the root testid a prop; `DeckHistoryPanel.vue` becomes a
wrapper passing `slides-history-panel`, so SPEC-103's e2e paths do not
change. `src/components/spreadsheet/WorkbookHistoryPanel.vue` passes
`sheet-history-panel` and is the `"history"` pane's body in SPEC-129's sheet
side panel (`SHEET_PANES.history = { title: "Version history", icon:
"history" }`). The text side's
`VersionHistoryPane.vue` (testids `version-history-*`) is untouched.

### Tests
- `documentStore.autosave.spec.ts` — per AC-23; the manifest's committed
  fallback for an untouched sheet; a sheet deleted mid-session leaves no
  record after commit; `draftRowsKey` maps
  `sheetrows:<docId>:<sheetId>:<chunk4>` to `draftsheetrows:…`; the draft
  source record is written only after a source reference changed.
- `spreadsheetVersions.spec.ts` — baseline/delta round trip, band windows
  (edits at rows 5 and 900,000 → two windows), `added`/deleted sheets via
  `order`, meta and `sourceXml` deltas, blob dedupe and hash caching, the
  cell-count and byte budget refusals, pruning with re-baselining,
  `workbookVersionEntries` deltas; `versionRecorder.spec.ts` for the moved
  loop with both codecs; `presentationVersions.spec.ts` passes unchanged.
- `spreadsheet/roundTrip.spec.ts` — restore fidelity per AC-21.
- `useWorkbookSaving.spec.ts` / `useWorkbookHistory.spec.ts` — lock states
  with an injected `LockManagerLike`, recovery, `sheetView` precedence,
  restore + undo-restore, preview isolation; `e2e/SPEC-138.spec.ts`.

### Specs this changes
- `specs/116-…` — the `saveDraft` guard and the `beginSession` bail go;
  the `getOrCreateDocument` guard stays.
- `specs/121-…` — `onApplied`'s `touchDocument` becomes `noteEdit()`: the
  pipeline's write-through becomes the SPEC-010 draft cadence, and its
  store assertion (exactly the changed sheets' rows records are written)
  holds at the commit points instead of after every edit; the undo stack,
  its labels and its cap are unchanged.
- `specs/103-…` — `DeckHistoryPanel.vue` wraps the extracted
  `HistoryPanel.vue`; its recording loop moves to `versionRecorder.ts`;
  `replaceStoredModel` joins its replace functions. No behaviour or testid
  changes.
- `specs/010-…`/`specs/040-…` — `DraftState`, `DraftRecord` and
  `VersionIndexEntry` gain additive fields; behaviour and copy unchanged.
- `specs/137-…` — the `documents` store gains the draft families
  `draftsheetrows:<docId>:<sheetId>:<chunk4>` and `draftsheetsource:<docId>`,
  which `loadAllRecords`' key ranges exclude and `deleteDocument` deletes;
  its per-chunk and source latches gain draft twins; its residency rule is
  honoured by every draft, commit and version capture.
- `specs/129-…` — `SheetPaneId` gains `"history"`.
- `specs/000-roadmap.md` gains this spec's row.
