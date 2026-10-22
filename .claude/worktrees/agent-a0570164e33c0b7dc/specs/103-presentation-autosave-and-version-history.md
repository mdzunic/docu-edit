---
id: SPEC-103
title: Autosave, drafts and version history for presentations
status: draft
area: io
complexity: complex
story_points: 13
depends_on: [SPEC-009, SPEC-010, SPEC-040, SPEC-079, SPEC-080]
---
## Why
Edit a deck for an hour, close the tab, lose the hour. SPEC-080's out of scope says it
outright — "Autosave/drafts (the store's draft machinery stays text-only, per the prep
milestone's `saveDraft` guard)" — and the guard is still there
(`src/lib/documentStore.ts:651`): `if (isPresentationModel(record.model)) { throw new
Error(\`Document ${id} is a presentation\`); }`, with a twin in
`src/composables/useAutosave.ts#beginSession`. A slide edit today reaches only
`touchDocument(model.id)` from `useSlideRender.commit()` — a write-through on browser
idle, no journal — so there is nothing to recover from and nothing to go back to.

The word processor solved this twice and both solutions are product-neutral in
substance: SPEC-010's draft/committed split with a recovery bar, and SPEC-040's version
chain (baselines, deltas, content-addressed blobs, a 50 MB budget, a prune loop). This
spec lifts the guard, generalises the one state machine instead of writing a second,
and moves SPEC-040's pure chain arithmetic into `src/lib/core`. What is genuinely new is
size: a deck's model carries `opaque` — *every source part, verbatim, base64* — so a
naive 2 s draft tick would structured-clone ten megabytes off the main thread's back,
every two seconds, forever. **This file is the normative source for the presentation
draft record, the deck version-history record shapes and chain rules, and the
over-budget degrade.** It owns no slide-model concept: SPEC-079 owns the slide model,
SPEC-080 the writers' rebuild rules and the undo stack.

## Acceptance criteria
- [ ] `saveDraft`'s presentation guard (`src/lib/documentStore.ts:651-654`) and `useAutosave.beginSession`'s twin are removed while `getOrCreateDocument`'s guard (line 413) **stays** — creating a text document over a deck's id still fails loudly; `DraftRecord.model` widens `DocumentModel` → `StoredModel`, `saveDraft(id, state)`/`commitDocument(id, state)` take `DraftState { caret: CaretPosition | null; scroll: number; slideIndex?: number }` in place of three positional arguments, `StoredDocumentRecord`/`DraftRecord` each gain `slideIndex?: number`, and `replacePresentationModel(id, model: PresentationModel): PresentationModel` joins `replaceDocumentModel` over one shared private `replaceModel`
- [ ] `PresentationModel.opaque` splits out of the persisted shell exactly as `media` does: `OPAQUE_KEY_PREFIX = "opaque:"`, `DRAFT_OPAQUE_KEY_PREFIX = "draftopaque:"`, `OpaqueRecord { id; docId; opaque: OpaquePart[] }`, `persistOpaque`/`persistDraftOpaque` with the same reference latch and failed-write latch clearing as `persistMedia`/`persistDraftMedia`, `stripOpaque`/`rejoinOpaque`, `isOpaqueRecord`/`isDraftOpaqueRecord`, `initLibrary` skip-and-rejoin, two extra `deleteRecord` calls in `deleteDocument` (`opaqueKey(id)`, `draftOpaqueKey(id)`, taking it from four to six); a record with no `OpaqueRecord` is legacy-inline and leaves the latch **unseeded** so the next persist migrates it (the media rule at `documentStore.ts:317-327`), because a deck reloading without `opaque` would silently save as a from-scratch rebuild and destroy SPEC-008 preservation
- [ ] Draft ticks after the first write only the shell: `useSlideRender`'s `commit()` stops re-slicing `model.media` unconditionally and does so only via a new `bumpMedia()`, and `frames.ts#insertImageFrame` assigns `model.media = [...model.media, part]` instead of `model.media.push(part)` (line 175) so the reference latch changes exactly when media changed; asserted by `putRecord` call counts, not timings — tick 1 writes shell+opaque+media, ticks 2..n write shell only
- [ ] `useAutosave` serves both products: `RestoredState.model` widens to `StoredModel`, options gain `getSlideIndex?: () => number`, and `beginSession` routes a `PresentationModel` through `replacePresentationModel`; cadence, `SAVE_STATE_*` strings, `hasUnsavedChanges`, `beforeunload`/`pagehide`, `flushForClose`, `commitNow` and the `keepRecovered`/`discardRecovered` pair are reused unchanged
- [ ] The deck header shows the save state (`deck-save-state` with `data-save-state`, text from the exported `SAVE_STATE_SAVING`/`SAVE_STATE_SAVED`/`SAVE_STATE_ERROR`) and, on `error`, a `deck-save-state-download` link running `deckActions.download()`; the recovery bar is the text side's markup extracted to `src/components/RecoveryBar.vue` with testids `recovery-bar`/`recovery-keep`/`recovery-discard` and its copy ("Recovered unsaved changes from &lt;time&gt;") preserved verbatim, so no existing e2e path changes
- [ ] Reopening a deck whose draft is newer than its committed record shows the bar over the recovered slides; "Keep" commits and clears it, "Discard" deletes the draft, restores the committed model and re-selects `slideIndex`; `PresentationPlayView` never writes a draft, a commit or a version
- [ ] SPEC-040's pure arithmetic moves to `src/lib/core/versionChain.ts` in the two-phase pattern (phase 1 add + re-export, phase 2 re-point and delete the copies): `VersionChainKind`, `SnapshotTrigger`, `VersionIndexEntry`, `VersionsIndexRecord`, `VersionBlobRecord`, `ChainInfo`, `gatherChainInfo`, `decideBaseline`, `measureBytes`, `sha256Hex` (module-private at `src/lib/versionHistory.ts:241` today, so phase one **exports** it from `src/lib/core/versionChain.ts` and `versionHistory.ts` imports it back; its `hashMediaPart`/`hashOpaquePart` wrappers and their `WeakMap` caches at `versionHistory.ts:251-274` stay text-side, only the digest itself being product-neutral), prune-candidate selection, `CHAIN_DELTA_CAP` 50, `WHOLESALE_REWRITE_RATIO` 0.5, `RETAINED_RECENT_COUNT` 10, `HISTORY_BUDGET_BYTES` and `HISTORY_PAUSED_MESSAGE`; `HISTORY_BUDGET_BYTES` becomes a literal `50 * 1024 * 1024` in core (core cannot import `openDocument`) and `versionHistory.spec.ts` gains an assertion pinning it equal to `MAX_FILE_SIZE_BYTES`
- [ ] `src/lib/presentationVersions.ts` (top level, beside `documentStore.ts` — `lib/presentation` may not import `./idb`) records deck versions into the **existing** `versions` object store under the same `versions:`/`version:`/`blob:` key prefixes with **no `DB_VERSION` bump**: baselines carry `slides` verbatim (including `dirty`, `partName`, `sourceXml`, `preserved`), deltas carry a `windowOf`-trimmed slide window, and `opaque`/`media` are content-addressed per part per the Reference
- [ ] Snapshots fire on SPEC-040's cadence and triggers: `createVersionCadenceTimer` (5 min of editing activity, only when the deck changed), each download (`useDeckActions.download`/`downloadCopy`), clean close via `onCleanClose`, and `pre-restore`/`post-restore`; the `keepVersionHistory` preference gates `cadence`/`commit` only
- [ ] The panel lists versions newest-first with timestamp, author (`preferences.authorName`), slide count and slide-count delta, and offers Name/Remove name, named-only filter, Delete, Preview, Restore and Download (`.pptx`/`.odp` for `model.format`), testids `slides-history-toggle`, `slides-history-panel`, `history-list`, `history-empty`, `history-version`, `history-name`, `history-name-input`, `history-name-submit`, `history-remove-name`, `history-delete`, `history-named-only`, `history-preview`, `history-preview-banner`, `history-exit-preview`, `history-restore`, `history-download`, `history-paused-notice`
- [ ] Preview renders the reconstructed model on the SPEC-079 stage read-only (`<SlideStage :interactive="false">` plus its own filmstrip index) with **no diff highlighting** — `applyTrackChanges` is `DocumentModel`-typed and there is no presentation equivalent — and exiting preview returns the live deck untouched
- [ ] Restore records `pre-restore`, then `record("Restore version")` → swap → `commit()` → `post-restore`; Ctrl+Z afterwards reverts only `slides`/`structureDirty` (all `usePresentationUndo`'s snapshot holds, per SPEC-080), and the panel copy says so, the pre-restore version being one click away
- [ ] Round trip: open a deck, snapshot, edit, restore the snapshot, download — every part of the download matches a download taken at snapshot time under `equalBytes` from `src/lib/core/equality.ts`, compared part-level after `unzip`; a no-edit deck that is snapshotted and restored still downloads byte-identical per part
- [ ] Over budget the deck degrades to no-history and **never fails the edit**: a snapshot whose blob set plus baseline would exceed `HISTORY_BUDGET_BYTES` is refused before any write, `paused` latches with `DECK_HISTORY_TOO_LARGE_MESSAGE` in `history-paused-notice`, and a history write failure never calls `markStorageUnavailable`, never touches the deck's `documents`-store records and never blocks `commit()`
- [ ] `PresentationView.vue` gains exactly one `useDeckSaving({ … })` call plus template hookup; `useSlideCommands` absorbs the view's `onFilmstripCommand`/`onFilmstripReorder`, and `src/views/PresentationView.size.spec.ts`'s `LINE_BUDGET` drops by at least 5 from whatever it stands at when this lands (550 → 545 today) with a ledger row appended — the budget never rises
- [ ] e2e (`e2e/SPEC-103.spec.ts`): open `mini-deck.pptx` → edit a title → reload → the edit is there with no bar → inject a synthetic newer `draft:<id>` record via `page.evaluate` + `indexedDB` (SPEC-010's technique) → reopen: `recovery-bar` visible over the drafted slides → Discard restores the committed deck → edit, download (a commit-point snapshot) → open `slides-history-panel`, name the version, preview it, exit preview, restore it → the stage shows the restored slides and the panel lists the restore's own two snapshots

## Out of scope
Server-side or cross-device drafts and history; sharing blobs between two decks'
histories; branching, diff between two arbitrary versions, and restoring a single slide
from an old version (the unit of restore is the deck). Diff highlighting in preview. A
deck-level two-tab edit lock (`core/editLock.ts` stays wired to DocumentView only).
Persisting the undo stack — SPEC-080 excluded it and this spec does not change that.
Draft or version state for the play view. Compressing history payloads. Autosaving back
to the user's file on disk.

## Reference

### Dependencies restated
- `specs/010-autosave-and-crash-recovery.md` — the committed/draft split. `records` in
  `src/lib/documentStore.ts` is the in-memory source of truth; write-through is best
  effort. `DRAFT_KEY_PREFIX = "draft:"`, `MEDIA_KEY_PREFIX = "media:"`,
  `DRAFT_MEDIA_KEY_PREFIX = "draftmedia:"`; `DraftRecord { id; docId; savedAt: number;
  model; caret: CaretPosition | null; scroll: number }`; `MediaRecord { id; docId; media:
  MediaPart[] }`. Committed at four points only (create/import, download, clean close,
  "Keep"); recovery triggers on `draft.savedAt > record.lastModified`.
  `createAutosaveScheduler` (`core/autosaveScheduler.ts`) gives 2 s debounce / 30 s
  deadline / 5 s retry; `putRecordYielding` awaits a real macrotask first because
  `IDBObjectStore.put()` structured-clones synchronously.
- `specs/040-version-history-and-restore.md` — baseline or delta against the
  immediately preceding version; four baseline triggers in order (no chain,
  `CHAIN_DELTA_CAP` 50, chain outweighs a copy, delta ≥ `WHOLESALE_REWRITE_RATIO` 0.5
  of the baseline); reference-counted blobs at `blob:<docId>:<sha256>`; 50 MB per
  document; pruning takes the oldest unnamed version outside the newest
  `RETAINED_RECENT_COUNT` 10, re-baselines anything retained that chained through it
  and re-measures each iteration; `createVersionCadenceTimer`
  (`VERSION_CADENCE_DEFAULT_MS` = 5 min) in `src/lib/core/versionCadenceScheduler.ts`.
- `specs/079-slide-canvas-rendering.md` — `PresentationModel { kind: "presentation"; id;
  title; format: "pptx"|"odp"; slides: Slide[]; slideWidthCm; slideHeightCm; opaque?:
  OpaquePart[]; media?: MediaPart[]; properties?: DocumentProperties; structureDirty?:
  true; odpContentEnvelope?: { prefix; suffix } }`; `Slide { id; frames; background?;
  notes?; dirty?; notesDirty?; transition?; animations?; partName?; sourceXml?;
  preserved? }`; `slideCount(model)`. `OpaquePart { name; data }` and `MediaPart
  { mediaId; contentType; data }` are base64 (`core/packageParts.ts`, `core/media.ts`).
- `specs/080-…-slide-management.md` — a deck with no `dirty`/`notesDirty`/
  `structureDirty` slide re-zips its parts verbatim; `usePresentationUndo` snapshots
  `structuredClone(model.slides)` plus `structureDirty` only, `MAX_UNDO_STEPS` 100,
  `record(label)` **before** the mutation; `useSlideRender.commit()` re-identifies the
  current slide, re-slices `slides` (and today `media`), calls `touchDocument(model.id)`
  and `triggerRef`s. `specs/009-…-open-from-url.md` — a dropped or `/open?url=`-fetched
  deck is stored in the library like any other and is not re-fetched on reload;
  `MAX_FILE_SIZE_BYTES` is 50 MB.
- **Equality**: `equalBytes(a: Uint8Array, b: Uint8Array): boolean` lives in
  `src/lib/core/equality.ts` (not `src/lib/equality.ts`) and is used part-level after
  `unzip`, never on container bytes. `equalStructural` is in `src/lib/equality.ts`; its
  `equalSemantic`/`resolve` are typed `(model: DocumentModel)` and **cannot** take a
  `PresentationModel`. So this spec states its own presentation-side relation: *restore
  fidelity* means per-part `equalBytes` over the two downloads plus field-level
  `toEqual` on `slides`/`slideWidthCm`/`slideHeightCm`/`properties`.

### The lifted guard, and splitting `opaque`
Both guards go; nothing else about the draft machinery changes shape. Deck state is
`{ caret: null, scroll: 0, slideIndex }` — a deck has no caret and its scroll is a
fit-to-workspace scale, not a position. `opaque` holds every source part verbatim — a
7 MB decompressed package is ≈ 9.4 MB of base64, structured-cloned on every 2 s tick if
left in the shell. It splits out under its own prefix with the same reference latch, and
because nothing may mutate `opaque` at read *or* edit its reference never changes for a
deck's whole open: written once per session, skipped thereafter.

`media` already has that machinery, but two things defeat
it on the deck path and both are fixed here — `useSlideRender.commit()` re-slices
`model.media` on **every** gesture, handing the latch a fresh reference each time, while
`frames.ts#insertImageFrame` does `model.media.push(part)`, keeping the reference
identical exactly when the bytes *did* change. Gating the re-slice behind `bumpMedia()`
without also making the insert assign a fresh array would silently stop persisting
inserted images — the two are one change. (SPEC-095 wants the same gate for a different
reason; they converge.)

### Deck version history — records, chain and blobs
`src/lib/presentationVersions.ts`, over core's record shapes:

```ts
interface OpaqueRef { name: string; hash: string }        // bytes live in blob:<docId>:<hash>
interface MediaRef { mediaId: string; contentType: string; hash: string }
// DeckFields: title, format, slideWidthCm, slideHeightCm, properties,
// structureDirty, odpContentEnvelope — all optional; a delta carries only what changed.
export interface DeckBaselinePayload { kind: "deck-baseline"; slides: Slide[];
  fields: DeckFields; opaque: OpaqueRef[]; media: MediaRef[] }
export interface DeckDeltaPayload {
  kind: "deck-delta"; parentId: string; fields: DeckFields;
  slides: { prefixLen: number; suffixLen: number; replaced: Slide[] };
  opaque?: OpaqueRef[]; media?: MediaRef[];   // present only when the ref changed
}
export interface DeckVersionPayloadRecord { id: string; store: "payload";
  docId: string; versionId: string; payload: DeckBaselinePayload | DeckDeltaPayload }
export async function recordDeckVersion(options): Promise<RecordVersionResult>;
export async function reconstructDeckVersion(docId, versionId): Promise<PresentationModel | null>;
export function deckVersionEntries(index): DeckVersionListEntry[];  // + slideDelta
```

`VersionIndexEntry` gains `slides?: number` (additive — renaming `words` would orphan
every existing `versions:<docId>` record); `src/lib/idb.ts`'s three version helpers
retype to `VersionStoreRecord | DeckVersionStoreRecord`, both type-only imports, and the
`versions` store already exists at `DB_VERSION = 3`, so there is no migration. Slide
deltas ride the text side's reference-equality trick: `commit()` re-identifies
only the current slide and its frames, so untouched slides keep their object identity
and `windowOf` (`src/lib/core/arrayWindow.ts`) trims them away. After an undo every
slide is a `structuredClone`, the delta grows, and the wholesale-rewrite rule writes a
baseline instead — a safe degradation, exactly as SPEC-040 documents for a post-reload
session. `opaque` and `media` are content-addressed with `crypto.subtle.digest(
"SHA-256", …)` over the base64 string, one blob per part; a deck's `media` deliberately
duplicates image parts already inside `opaque`, and identical bytes hash identically,
so content-addressing collapses that duplication to one blob on its own.

### The budget, and the honest degrade
Reuses `HISTORY_BUDGET_BYTES` = 50 MB per deck. The worked case — a 30-slide deck from a
6 MB `.pptx` (≈ 7 MB decompressed), edited 8 h/day for five days:

| Quantity | Value |
| --- | --- |
| Blob set (`opaque` ∪ `media`, base64, written once) | ≈ 9.4 MB |
| Versions at one per 5 min over 40 h → baselines at the 50-delta cap | ≈ 480 → ~10 |
| Baselines ~10 × 240 KB (≈ 8 KB/slide, blobs excluded) + deltas ~470 × 10 KB | ≈ 7.1 MB |
| **Total** | **≈ 16.5 MB** — never prunes |

The ceiling is set by the blob set, not the deltas: `opaque` base64 runs ≈ 1.37× the
decompressed package, so a package past roughly **36 MB** cannot hold even one
baseline, while `openDocument` opens it happily up to 50 MB. That is detected **before**
any write — `blobBytes + measureBytes(baseline) > HISTORY_BUDGET_BYTES` — and latches
`paused: true` with

```ts
export const DECK_HISTORY_TOO_LARGE_MESSAGE =
  "Version history off — this deck's package is larger than the 50 MB history budget";
```

rendered in `history-paused-notice` beside the empty list; normal exhaustion (nothing
prunable, still over) uses core's `HISTORY_PAUSED_MESSAGE` unchanged. Either way the edit
proceeds: history never evicts the deck's own `documents`-store records, never latches
the storage banner, and never awaits anything on the `commit()` path.

### View wiring and UI
One composable, `src/views/presentation/useDeckSaving.ts`, taking `DeckSavingDeps
{ model: ShallowRef<PresentationModel|null>; presentationId: Ref<string>; currentIndex:
Ref<number>; record: (label?: string) => void; commit: () => void; reportError:
(message: string) => void }` and returning `{ saveState, saveStateText, recovery,
keepRecovered, discardRecovered, noteEdit, flushForClose, history }`. It composes
`useAutosave` (drafts, recovery, tab close) and the history controller: `open` = the
persisted preference `presentationHistoryOpen` (added to `src/lib/preferences.ts` in all
three places — declaration, default `false`, `isBoolean` validator), plus `toggle`,
`versions`, `visibleVersions`, `namedOnly`, `pausedMessage`, `previewModel`,
`previewIndex`, `preview`, `exitPreview`, `restore`, `nameVersion`, `removeName`,
`deleteVersion`, `downloadVersion`.

`SlideRenderDeps` gains `onCommit?: () => void`, called at the end of `commit()` and
wired to `noteEdit`; `refresh()` deliberately does not fire it, so a mid-drag preview
never arms the cadence. `DeckActions.vue` gains the two save-state elements and the
`slides-history-toggle` button beside Notes (`EditorSaveStatus.vue` is not reused: it
bundles the text editor's `close-document` control, and moving a control between
surfaces rewrites every e2e path reaching it).
`src/components/presentation/DeckHistoryPanel.vue` copies SPEC-085/086's panel contract
exactly — props `{ controller: { type: Object, required: true }, slideNumber: { type:
Number, default: 1 } }`, emits nothing, `const c = props.controller;`, runtime
`defineProps` only (espree lints `.vue` scripts), typed transforms in a sibling
`deckHistoryRender.ts` — and mounts in `.stage-column` under `v-if=
"saving.history.open.value"`. Undo label: `"Restore version"`.

### Tests
- `src/lib/core/versionChain.spec.ts` (new): chain gathering, the four baseline triggers,
  prune-candidate selection, `measureBytes`, `sha256Hex`.
- `src/lib/documentStore.autosave.spec.ts`: the lifted guard, deck draft/commit/recover,
  the opaque split and both latches, legacy-inline rejoin, `putRecord` call counts per
  tick, `deleteDocument` cleaning all six keys; its existing expectations re-expressed
  against `DraftState` with identical values — coverage preserved, never weakened.
- `src/lib/presentationVersions.spec.ts` (new): baseline/delta encode-decode, `windowOf`
  slide trimming, blob refcounting and `opaque`∩`media` dedupe, the budget refusal,
  pruning with re-baselining, `deckVersionEntries` slide deltas.
- `src/lib/presentation/roundTrip.spec.ts`: restore fidelity — snapshot → edit → restore
  → write, per-part `equalBytes` against the snapshot-time write, and the no-edit deck
  still byte-identical per part after a snapshot/restore cycle.
- `src/lib/versionHistory.spec.ts`: `HISTORY_BUDGET_BYTES === MAX_FILE_SIZE_BYTES`.
- `src/views/presentation/useDeckSaving.spec.ts` (new) and
  `src/views/PresentationView.spec.ts`: recovery bar, save-state text, panel wiring;
  `src/views/PresentationView.size.spec.ts` at its ratcheted-down budget;
  `e2e/SPEC-103.spec.ts` per AC-16.

### Specs this changes
- `specs/000-roadmap.md` gains this spec's phase row; nothing else changes there.
- SPEC-080's out-of-scope line "Autosave/drafts (the store's draft machinery stays
  text-only…)" is superseded here; its "Persisting the undo stack" exclusion still stands.
- SPEC-010's `saveDraft`/`commitDocument` signatures and `useAutosave`'s `RestoredState`
  widen to both products; its behaviour and copy are unchanged.
- SPEC-040's chain arithmetic and constants move to `src/lib/core/versionChain.ts` and
  are re-exported; the text pane, its cadence and its numbers are untouched.
- SPEC-095, if it lands first, already introduces `bumpMedia()`; whichever lands second
  adopts the other's version rather than adding a second gate.
