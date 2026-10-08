---
id: SPEC-137
title: Large workbooks — performance
status: draft
area: platform
complexity: complex
story_points: 13
depends_on: [SPEC-047, SPEC-120, SPEC-122]
---
## Why
Everything in Phases 17–18 works at a thousand rows. The workbooks people
actually bring are the 25,000-row export, the year of transactions, the price
list with a running total down one column — and on those the product as specced
does three things it must not. It parses on the main thread, so the tab freezes
for the length of the parse. It writes each sheet's rows as **one** IndexedDB
record (SPEC-116's `sheetrows:` split), so typing one number re-clones half a
million cells; and because that record sits in the same object store the
library's `getAll` reads at start-up, and the shell still carries the retained
package (`opaque`, base64) and every `.ods` table's `sourceXml`, the landing page
of a library holding twenty large workbooks reads every cell of all twenty before
it can draw a card. And nothing measures any of it.

SPEC-047 is the precedent and the method: one budget table, a generated fixture,
a median-of-5 benchmark in its own Playwright project with an absolute ceiling
and a 15% relative gate against the container-calibrated `e2e/perf-baseline.json`.
Its own amendment explains why the word processor could not parse off-thread —
its XML layer is `DOMParser`, which workers lack. The spreadsheet readers are
different by construction: SPEC-116 put every O(cells) part (worksheet parts,
shared strings, `content.xml`) on the forward-only `core/xmlScanner.ts`, so the
heavy half of a workbook read **can** leave the main thread, and this spec moves
it.

**This file is the normative source for the workbook performance budgets and
their benchmark, the reader's phase split and the off-thread parse protocol, the
spreadsheet storage layout (row chunks, the source record, lazy residency), the
grid's rendering ceiling, and the decision on `MAX_WORKBOOK_CELLS`.**

## Acceptance criteria
- [ ] Every workbook budget is one row of "The budgets", asserted by `e2e/SPEC-137.perf.spec.ts` on the median of 5 runs under SPEC-047's two gates — the absolute ceiling, and the checked-in `e2e/perf-baseline.json` value for that row's key + 15% — with each row's key, ceiling and direction restated as a `BUDGETS` object at the top of the file, exactly as `e2e/SPEC-047.perf.spec.ts` does
- [ ] The fixtures are generated, not committed: `scripts/build-spec137-fixtures.ts` (run with `npx vite-node`, ending in `process.exit(0)` because the idle scheduler keeps Node alive) builds `e2e/fixtures/spec137-large.xlsx` (500,000 cells, 25,000 rows × 20 columns) and `spec137-library.xlsx` (100,000 cells) through `spreadsheet/fixtures/build.ts`, deterministically (seeded PRNG, fixed timestamps — same script, same bytes), Excel-shaped per "The benchmark fixtures"; both paths join `.gitignore` beside `spec-047-fixture.docx`, and the benchmark builds them on demand when absent
- [ ] `playwright.config.ts` gains a `perf-workbook` project (`testMatch: /SPEC-137\.perf\.spec\.ts$/`, `bypassCSP: true`, Desktop Chrome) whose `dependencies` name the benchmark that is the chain's tail at implementation time (today `perf-large-document`), so it runs alone at its own level after the functional suite and every other benchmark; `src/perfBenchmarkProjects.spec.ts` adds `"perf-workbook"` to its `benchmarks` list and asserts its level is greater than its predecessor's — no existing assertion is weakened, and no project ever shares a level with a benchmark
- [ ] The reader splits into the phases of "The read phases" in `src/lib/spreadsheet/open.ts` — `prepareWorkbookSources` (DOMParser-free), `readWorkbookShell` (main thread, may use `parseXml`), `scanSheet` (DOMParser-free, emits `RowChunk`s) and `assembleWorkbook` — with `readSpreadsheetBytes` keeping its SPEC-116 signature and running them inline; a new `src/lib/spreadsheet/open.workerSafety.spec.ts` runs the two DOMParser-free phases over every spreadsheet fixture with `DOMParser`, `XMLSerializer` and `document` stubbed to `undefined` and asserts the assembled model `toEqual`s `readSpreadsheetBytes`' — the rule every later reader change keeps: per-sheet reading never touches `DOMParser`
- [ ] Opening parses off the main thread: `src/workers/spreadsheetParse.worker.ts` (a module worker, `self` narrowed locally like `wordCount.worker.ts`) hosts `src/lib/spreadsheet/parseSession.ts`, and `src/lib/workers/spreadsheetParseClient.ts` exports `readSpreadsheetOffThread(parts, format, title)` over `createWorkerRpcClient` speaking exactly the protocol in "The parse protocol"; `openAnyDocumentFile`'s spreadsheet branch and the store's upgrade re-read both use it (`upgradeSpreadsheetModel` gains an optional `read` parameter defaulting to `readSpreadsheetBytes`, because the product layer cannot import the worker client — the store passes `readSpreadsheetOffThread`); rows cross the boundary as `SHEET_ROWS_CHUNK`-row chunks pulled one per request, so no single message deserialises more than one chunk on the main thread
- [ ] The off-thread path degrades, never fails: when `isWorkerSupported()` is false (Vitest/jsdom) the same phases run inline with `yieldToMainThread()` every `SCAN_YIELD_ROWS = 2_000` rows; a worker that errors or terminates mid-open is disposed and the open retried once inline; errors cross the boundary as `{ name, message, offset? }` and are re-thrown as the same classes (`XmlScanError`, the readers' parse error, the `WORKBOOK_TOO_LARGE_MESSAGE` refusal, counted in the worker across all sheets) so every open-funnel message is unchanged; `parseSession.spec.ts` drives the session's message handler directly and proves its model equals the inline one for every fixture
- [ ] Rows persist in chunks per "The storage layout": the `documents` store holds one record per `SHEET_ROWS_CHUNK = 512`-row band of row indices per sheet, keyed `sheetrows:<docId>:<sheetId>:<chunk>` (chunk zero-padded to 4 digits — SPEC-116's per-sheet key gains one segment, and keeps the `sheetrows:` prefix SPEC-138's draft keys are derived from), plus one `sheetsource:<docId>` source record holding `opaque`, `odsContentEnvelope` and every sheet's `sourceXml`; the shell record holds **no cell, row or source byte** — each sheet's `rows: []`, `opaque: []`, no `sourceXml` — plus a store-side manifest `sheetStore { chunks: Record<sheetId, number[]>; storedBytes: number }`
- [ ] A persist rewrites only what changed: for a sheet whose `rows` reference changed it takes `windowOf(lastPersisted, current)` (`core/arrayWindow.ts`), maps the window's row-index span to chunk indices, and rewrites a chunk only when its row slice differs by identity from the last persisted slice, deletes chunks that became empty and drops a deleted sheet's chunks; the source record is rewritten only when `opaque` or a `sourceXml` reference changed; all of it — changed chunks, deletions, the source record and the shell — commits in **one** readwrite transaction through `core/idb.ts`'s new `writeBatch`, which issues one operation per success callback so its structured clones never fuse into one long task; editing one cell of the 500,000-cell fixture writes exactly one chunk record and the shell (unit-tested against the `idb` seam)
- [ ] Workbooks are resident only while needed: `getSpreadsheetShell(id)` (sync, always — library, guards, cards), `getSpreadsheet(id)` (sync, the full model only when resident, else `undefined`), `loadSpreadsheet(id): Promise<SpreadsheetModel | undefined>` (rejoins the source record and every chunk with one `get` per chunk and `yieldToMainThread()` between them; idempotent, concurrent callers share one promise) and `releaseSpreadsheet(id)` (drops rows and source from memory once pending persistence has flushed, keeping the shell); `useWorkbookLoad` awaits `loadSpreadsheet` and releases on unmount, `useLibrary.downloadDocuments` and the library's duplicate load-then-release workbooks, and the store never persists rows or source for a workbook that is not resident
- [ ] The landing page reads no cell: `loadAllRecords` (`src/lib/idb.ts`) reads the `documents` store through `getAllInRange` over the key ranges that exclude the `sheetrows:` and `sheetsource:` prefixes (and any draft-chunk prefix SPEC-138 adds), so a library of 20 stored `spec137-library` workbooks lists, sorts, searches and renders its `SheetMiniature` cards from shells and `preview` alone (asserted against the `idb` seam: no chunk or source record is read by `initLibrary`, `listDocuments` or a card render), and `libraryStorageUsage` reports a workbook from its manifest's `storedBytes` — it never `JSON.stringify`s a spreadsheet
- [ ] SPEC-116-era records upgrade lazily and once: a shell with no `sheetStore` manifest is a 116-era record — `loadSpreadsheet` rejoins it from its per-sheet `sheetrows:<docId>:<sheetId>` records, then rewrites it as chunk records, a source record and a slim shell and deletes the legacy records, in one `writeBatch`; an idle pass after `initLibrary` (one workbook per `runWhenIdle` callback) migrates the 116-era workbooks nobody opens, so their heavy shells stop being read at start-up
- [ ] A missing piece is reported, never papered over: a manifest naming a chunk or source record that storage does not return opens the workbook read-only with `WORKBOOK_STORAGE_INCOMPLETE_MESSAGE` ("Part of this workbook couldn't be loaded from browser storage — download it or reopen the original file") and the record is never written back from that partial model
- [ ] Recalculation stays on the main thread — SPEC-120 leaves the choice to this spec, decided in "Where the work runs" — driven exactly as SPEC-120 pins it: `startRecalc` → `RecalcPass.step(RECALC_SLICE_MS)` (12 ms) through `scheduleEagerWork` in `useRecalc`, the open-time graph build sliced in `runWhenIdle` after first paint, a newer pass cancelling and inheriting an older one; this spec changes no engine API and only adds `data-calc-state="idle" | "calculating"` on `spreadsheet-grid`, set by `useRecalc` the moment a pass starts and cleared when it completes (independent of the 250 ms `CALCULATING_STATUS_DELAY_MS` before `sheet-status-calculating` shows), so the benchmark has a deterministic end-of-recalc signal
- [ ] Saving stays responsive: SPEC-122's xlsx sheet serialiser and ODS table serialiser build parts as arrays of string pieces joined once and yield whenever `SERIALISE_SLICE_MS = 8` ms of work has elapsed, and large retained parts are decoded through `base64ToBytesChunked` (new in `core/packageParts.ts`: ≤ 4 MiB of base64 per `atob`, a yield between slices) instead of the synchronous `base64ToBytes`
- [ ] The grid holds a rendering ceiling per "The grid's ceiling": `SheetCellLayer.vue` (SPEC-117) keys its `sheet-cell` elements by viewport **slot**, not by address, so a scroll step patches existing elements instead of mounting new ones; never more than `GRID_MAX_RENDERED_CELLS = 5_000` cell elements exist (overscan shrinks to zero first); when the viewport alone would exceed it the grid switches to overview rendering — `SheetOverviewCanvas.vue` (`sheet-overview-canvas`) paints fills, borders, gridlines, headers and text (greeked below `GRID_GREEK_BELOW_PX = 6`), only the active cell keeps a `sheet-cell` element, and pointer hit-testing goes through `GridMetrics` — and the whole `spreadsheet-grid` subtree stays under `GRID_MAX_DOM_NODES = 12_000` elements at any zoom
- [ ] Readers allocate lean cells: a `Cell` is created with `col` then `value` first and optional fields only when present (never an `undefined`-valued key), a `Row` with `index` then `cells`, integral numbers stay integral, and every cell that reads one shared string holds the same string instance — `Object.keys` of a plain numeric cell from either reader is exactly `["col", "value"]`
- [ ] `MAX_WORKBOOK_CELLS` stays **2,000,000** per "Memory and the cell ceiling", and the benchmark proves it affordable: with `H0` the settled heap of an empty workbook and `H` that of `spec137-large.xlsx`, it asserts `H0 + 4 × (H − H0) ≤ 400` MB (SPEC-047's whole-app heap ceiling); if that projection fails when this spec is implemented, the implementing run lowers the constant to the largest multiple of 250,000 that passes and updates `WORKBOOK_TOO_LARGE_MESSAGE` and its tests in the same change — this spec never raises it
- [ ] Opening a large workbook shows progress: from the file choice until navigation, the library shows `spreadsheet-open-progress` (`role="status"`) reading "Opening {title}… {n}%" as chunks arrive, and the main thread stays interactive throughout (the open row of "The budgets")
- [ ] The baseline is calibrated honestly: the implementing run records each row in `e2e/perf-baseline.json` under the key named in "The budgets", with a `_comment_sheet` entry saying whether it is a container median of 5 or a dev-machine projection (Mac medians × 1.24, × 2.5 for the three open-path rows, unscaled for the heap rows — the factory-setup and perf-gotcha rules), and a quiet factory-container run replaces any projection in an explicit commit before this spec's gate; no baseline is ever rewritten by the harness
- [ ] e2e functional (`e2e/SPEC-137.spec.ts`, in the normal `chromium` project, no timing assertions): open `spec137-library.xlsx` → the grid paints → scroll to row 4,000 and count ≤ 5,000 `sheet-cell` and ≤ 12,000 elements under `spreadsheet-grid` → zoom to 10% → `sheet-overview-canvas` is visible and exactly one `sheet-cell` remains → back to 100% → edit `B2` → reload → `B2` holds the edit, and the `documents` store holds that workbook's `sheetrows:` chunk records and a shell whose sheets have empty `rows` → open it from the library again → it renders → the library storage panel lists it with a non-zero size
- [ ] Tests: `parseSession.spec.ts`, `open.workerSafety.spec.ts`, `open.spec.ts` (phase split equals the old single pass), `documentStore.sheets.spec.ts` (chunk keys and bands, the 116-era migration, identity-latched rewrite, empty-chunk deletion, sheet deletion, one `writeBatch` per persist, residency and release, the incomplete-storage refusal, `initLibrary` reading no chunk or source record, `storedBytes`), `core/idb.spec.ts` (`writeBatch` ordering, `getAllInRange`, `deleteRange`), `core/packageParts.spec.ts` (`base64ToBytesChunked` equals `base64ToBytes`), `SheetGrid.spec.ts` (slot keys stable across a scroll step, ceiling, overview switch), `perfBenchmarkProjects.spec.ts`, `e2e/SPEC-137.spec.ts` and `e2e/SPEC-137.perf.spec.ts`

## Out of scope
Recalculation in a worker (see "Where the work runs"; Backlog). Streaming a
workbook from a server, or paging rows from IndexedDB on demand while editing
(the whole workbook is resident while open). Workbooks above the cell ceiling —
refused, as SPEC-116 states. Four costs earlier specs deferred here are
deferred again, explicitly, to the Backlog — the benchmark fixture carries none
of them, so a budget for them would be a guess: parsing `.csv` in a worker and
streaming it (SPEC-126's synchronous `readCsvBytes`), conditional-format
evaluation beyond the viewport-driven bound (SPEC-129), sheets with thousands
of drawings (SPEC-131), and incremental `GridMetrics` invalidation (SPEC-117 —
its full rebuild per rows change is O(rows) and fits the edit budget at this
scale). Pivot refresh (SPEC-139) and chart recomputation (SPEC-132) on large
ranges. Mobile and touch performance. Choosing a CI provider (SPEC-047's
position, unchanged). A 30-minute leak session for workbooks (SPEC-047's leak
test covers the shared app; a workbook one is Backlog).

## Reference

### Dependencies restated
- `specs/047-large-document-performance.md` — the method: a budget table where
  every row is asserted; median of 5; an absolute ceiling plus a 15% relative gate
  against `e2e/perf-baseline.json`, which only an explicit commit changes; long
  tasks from the Long Tasks API (so a block under 50 ms reads as 0); heap as CDP
  `Performance.getMetrics` → `JSHeapUsedSize` after `HeapProfiler.collectGarbage`;
  open timings start when the file is chosen, with `t0` taken only after
  `open-document-input` is attached (the 5bf8ccc barrier — `page.goto` resolving is
  not the app being mounted). `src/lib/workers/workerRpc.ts`:
  `isWorkerSupported()`, `createWorkerRpcClient<TRequest, TResponse>(factory)` →
  `{ call(request, transfer?), terminate() }`, lazy worker creation, every pending
  call rejected on `error`/`terminate`; `core/idleScheduler.ts`: `yieldToMainThread`,
  `runWhenIdle`, `scheduleIdleWork`. The benchmark chain in `playwright.config.ts`:
  `perf-quiet` → `perf-charts` → `perf-find` → `perf-large-document`, one benchmark
  per dependency level, never hung off `chromium`, pinned by
  `src/perfBenchmarkProjects.spec.ts`. The factory container runs ~1.24× slower than
  the Mac overall and 2–3× on the open path; heap readings agree across machines.
- `specs/116-…` — the model and its invariants (copy-on-write rows is what makes
  the identity-latched chunk rewrite sound), `readSpreadsheetBytes`,
  `upgradeSpreadsheetModel`, `core/xmlScanner.ts`, `MAX_WORKBOOK_CELLS`,
  `WORKBOOK_TOO_LARGE_MESSAGE`, the shell/rows split this spec replaces
  (`SHEET_ROWS_KEY_PREFIX = "sheetrows:"`, one record per sheet), `buildPreview`,
  `GRID_OVERSCAN_ROWS = 8`, `GRID_OVERSCAN_COLS = 4`, `MAX_SCROLL_EXTENT_PX`.
  `src/lib/documentStore.ts` today: `initLibrary` → `loadAllRecords()` (`getAll`
  over `documents`), the media split (`MEDIA_KEY_PREFIX`, `lastPersistedMedia`),
  `libraryStorageUsage` (JSON length per record), `duplicateDocument` (sync).
- `specs/117-…` `GridMetrics` (`colLeft`, `colWidth`, `colAt`, `rowTop`,
  `rowHeight`, `rowAt`), zoom 10–400%, Ctrl/Cmd+End = last used cell;
  `specs/119-…` shared formulas expand per cell at read, and `parseFormulaAt(text,
  host)` in `formula/parse.ts` caches ASTs in an LRU of `FORMULA_AST_CACHE_SIZE =
  50_000` keyed by `relativeFormulaKey`, so one shared block shares one AST;
  `specs/120-…` `recalculate(model, changed: CellKey[], options?: RecalcOptions):
  RecalcResult`, `startRecalc(model, changed, options?)` →
  `RecalcPass { step(budgetMs); cancel(); result }`, `RECALC_SLICE_MS = 12`,
  `CALCULATING_STATUS_DELAY_MS = 250`, `sheet-status-calculating`, `useRecalc.ts`,
  the open order (paint with cached values, then the sliced graph build in
  `runWhenIdle`; a full pass only for `fullCalcOnLoad` or a missing cache),
  `evaluate` DOM-free; `specs/121-…` write-through persistence and the by-reference undo
  snapshot; `specs/122-…` the sheet writers.

### The budgets
Measured on `spec137-large.xlsx` in Desktop Chrome (1280 × 720) at 100% zoom.

| Key (`perf-baseline.json`) | Ceiling | How it is measured |
| --- | --- | --- |
| `sheetOpenFirstPaintMs` | 6,000 ms | `t0` once `open-document-input` is attached; `setInputFiles` → first `sheet-cell` with text inside `spreadsheet-grid` |
| `sheetMainThreadBlockDuringOpenMs` | 150 ms | longest task from `t0` to first paint + 2 s (initial persistence included) |
| `sheetReopenFirstPaintMs` | 3,000 ms | after a reload (not resident), click its `library-row` → first `sheet-cell` with text |
| `sheetScrollFps` | ≥ 30 fps | frame timestamps over a 5 s programmatic fling (2,000 px wheel steps, down then up) |
| `sheetMainThreadBlockDuringScrollMs` | 50 ms | longest task during the fling |
| `sheetJumpToEndMs` | 500 ms | Ctrl/Cmd+End keydown → `sheet-cell[data-ref="T25000"]` attached and `name-box` reads `T25000` |
| `sheetChainRecalcMs` | 2,000 ms | type in `S1`, Enter → `spreadsheet-grid[data-calc-state="idle"]`, then (untimed) `T10000` holds the new total |
| `sheetMainThreadBlockDuringRecalcMs` | 50 ms | longest task in that window |
| `sheetSaveMs` | 5,000 ms | after the chain edit (sheet dirty): `spreadsheet-download` → download bytes resolved |
| `sheetMainThreadBlockDuringSaveMs` | 100 ms | longest task in that window |
| `sheetLibraryLandingMs` | 1,500 ms | 20 `spec137-library` workbooks stored; reload `/` → `performance.now()` in the page when the 20th `library-row` attaches (navigation-relative) |
| `sheetLibraryHeapDeltaMB` | 15 MB | settled heap of that landing minus an empty library's |
| `sheetHeapAfterSettleMB` | 300 MB | settled heap with `spec137-large.xlsx` open |

The library rows are seeded once per benchmark run (20 opens of
`spec137-library.xlsx`, outside the median loop). `sheetLibraryHeapDeltaMB` is
the row that proves no cell was read: twenty resident 100,000-cell workbooks
would cost far more than 15 MB.

### The benchmark fixtures
`spec137-large.xlsx`: one sheet "Data", A1:T25000. Columns A–S: A text from
1,000 distinct shared strings, B–H numbers (integers and two-decimal amounts), I
dates (serials with `s` pointing at a numFmt-14 xf), J booleans, K–S mixed numbers
and strings; column T rows 1–10,000 a dependency chain — `T1` `=S1`, `T2:T10000`
one shared formula `T(n−1)+S(n)` (`t="shared" ref="T2:T10000" si="0"`) — every
formula with its cached value, rows 10,001–25,000 of T plain numbers; no
`fullCalcOnLoad`, no `calcChain` drift. Excel's shapes: `spans` and
`x14ac:dyDescent` on rows, `<dimension ref="A1:T25000">`, `mc:Ignorable`, a theme
and a styles part with ~10 `cellXfs`. `spec137-library.xlsx`: the same generator at
5,000 rows, no chain.

### The read phases (`src/lib/spreadsheet/open.ts`)
```ts
export interface WorkbookSources {            // phase 1 output — plain data, DOMParser-free
  format: SpreadsheetFormat;
  sharedStrings?: SharedStringTable;          // xlsx, scanned once (xlsx/sharedStrings.ts's table type)
  odsEnvelope?: { prefix: string; suffix: string };
  odsTables?: { start: number; end: number }[];   // UTF-16 offsets of each table's sourceXml in content.xml
}
export interface SheetScanJob { sheetIndex: number; partName: string; range?: { start: number; end: number } }
export interface ScanContext { /* plain JSON built by the shell phase: date system, style maps, … */ }
export interface RowChunk { sheetIndex: number; chunk: number; rows: Row[] }   // rows with index in [chunk·512, chunk·512+512)
export interface SheetScanResult { sheetIndex: number; fields: Partial<Sheet>; cells: number }   // everything but rows
export function prepareWorkbookSources(texts: Map<string, string>, format: SpreadsheetFormat): WorkbookSources;
export function readWorkbookShell(parts: Map<string, Uint8Array>, sources: WorkbookSources, title: string):
  Promise<{ shell: SpreadsheetModel; jobs: SheetScanJob[]; context: ScanContext }>;
export function scanSheet(job: SheetScanJob, text: string, sources: WorkbookSources, context: ScanContext,
  emit: (chunk: RowChunk) => void, budget: { cellsLeft: number }): SheetScanResult;
export function assembleWorkbook(shell: SpreadsheetModel, results: SheetScanResult[], chunks: RowChunk[]): SpreadsheetModel;
```
Phase 2 reads the small parts (`workbook.xml` and rels, `styles.xml`, theme,
`docProps`, `settings.xml`, `meta.xml`, `styles.xml` of an ods, the ods envelope
prefix's automatic styles, and other workbook-level parts later specs add) and
builds `ScanContext`. Phase 3 reads one worksheet part or one ods table, **all of
it** — `sheetData` and every sheet-level element later specs model (columns,
merges, views, hyperlinks, page setup …) — returning those in `fields`. A later
spec that must resolve something per cell against a workbook-level table adds it
to `ScanContext`; it never parses a part with `DOMParser` inside phase 3.

### The parse protocol
One `createWorkerRpcClient<SheetParseRequest, SheetParseResponse>` per open,
terminated when the open settles. `src/lib/workers/` is app-layer
infrastructure (`wordCountClient.ts` already imports the text model), so the
client may import `lib/spreadsheet`; the worker file imports only
`parseSession.ts`, whose imports are spreadsheet siblings and core.
```ts
type SheetParseRequest =
  | { op: "prepare"; format: SpreadsheetFormat; sources: { name: string; bytes: Uint8Array }[] }   // copies, not transfers
  | { op: "scan"; session: number; jobs: SheetScanJob[]; context: ScanContext }
  | { op: "chunk"; session: number; index: number }
  | { op: "dispose"; session: number };
type SheetParseResponse =
  | { op: "prepared"; session: number; sources: Omit<WorkbookSources, "sharedStrings"> }
  | { op: "scanned"; results: SheetScanResult[]; chunkCount: number }
  | { op: "chunk"; chunk: RowChunk }
  | { op: "disposed" }
  | { op: "error"; name: string; message: string; offset?: number };
```
The main thread unzips (the open funnel already does, with its yields), posts
**copies** of the source parts (xlsx: `sharedStrings.xml` and the worksheet parts;
ods: `content.xml`) so `opaque` keeps its own bytes, runs phase 2 itself, then pulls
chunks in order, appending each into the shell with a `yieldToMainThread()`
between pulls and reporting `index / chunkCount` as the open progress. The
worker decodes, holds the shared-string table and source texts for the session,
and counts cells across every job against `MAX_WORKBOOK_CELLS`. Rows travel by
structured clone: a columnar binary encoding would be transferable, but decoding
it into `Cell` objects on the main thread costs what the native deserialiser
already costs, so chunking — not encoding — is what bounds the block.

### The storage layout
```ts
interface SheetRowsChunkRecord { id: string /* `sheetrows:${docId}:${sheetId}:${chunk4}` */; docId: string;
  sheetId: string; chunk: number; rows: Row[] }
interface WorkbookSourceRecord { id: string /* `sheetsource:${docId}` */; docId: string; opaque: OpaquePart[];
  odsContentEnvelope?: { prefix: string; suffix: string }; sheetSources: Record<string, string> }
// the shell's stored record gains:
//   sheetStore: { chunks: Record<string /* sheetId */, number[]>; storedBytes: number }
export const SHEET_ROWS_CHUNK = 512;
export const SHEET_SOURCE_KEY_PREFIX = "sheetsource:";
```
All of it lives in the existing `documents` store — no `DB_VERSION` bump, one
store per transaction, and SPEC-138's draft records (`"draft"` + the committed
key) sit beside the records they shadow. `SHEET_ROWS_CHUNK = 512`: at this
fixture's 20 columns a chunk is ~10,000 cells, a few milliseconds of structured
clone, and 49 records per 25,000 rows; chunks are bands of row **index**, not
array position, so inserting one sparse row rewrites one chunk rather than every
chunk after it. Sheet ids are minted from `[A-Za-z0-9_-]`, so keys stay
unambiguous and a 116-era key (three segments) is told from a chunk key (four).
A workbook's records are the ranges `sheetrows:${docId}:` …
`sheetrows:${docId}:\uffff` and `sheetsource:${docId}`; `deleteDocument` and
emptying the trash delete them. `storedBytes` is estimated at persist time (per
chunk: 16 bytes per cell plus string lengths; the source record: its base64
lengths). `core/idb.ts` gains `getAllInRange(store, lower?, upper?, { lowerOpen?,
upperOpen? })`, `deleteRange(store, lower, upper)` and `writeBatch(store, ops:
({ put: unknown } | { delete: string } | { deleteRange: [string, string] })[])`,
which issues each operation from the previous one's success callback inside one
readwrite transaction. Document ids are UUIDs, so no document key begins with
`sheet`; the start-up read is three ranges — below `sheetrows:`, between the two
prefixes, above `sheetsource:\uffff`.

### Where the work runs
- **Parse** — worker, per the protocol. **Small parts** stay on the main thread:
  they need `DOMParser` and are O(1) in the cell count.
- **Recalculation** — main thread, time-sliced (SPEC-120 explicitly left the
  worker open to this spec; it is declined). A worker would need the model:
  either cloned per edit (500,000 cells is a clone of the same order as the work)
  or mirrored and kept in sync per edit (twice the memory and a second copy of
  every command). The 10,000-cell chain is tens of milliseconds of evaluation;
  SPEC-120's 12 ms slices meet the 50 ms block at the cost of a few frames.
- **Save** — main thread, time-sliced, as SPEC-047 found for `writeDocx`: the
  model lives here, and the zip's deflate already runs through
  `CompressionStream`.
- **Persistence** — main thread, one transaction with one operation per success
  callback.

### The grid's ceiling
SPEC-117 applies zoom as a scale transform over unzoomed `GridMetrics`, so at
10% the viewport spans ten times the grid pixels — the reason a ceiling is
needed at all. A rendered cell's key is `${row mod R}:${col mod C}`, where R × C
is the render window's capacity; Vue therefore patches text, attributes and position of the
same elements as the window slides. The window is the viewport plus overscan;
overscan shrinks before the ceiling is crossed. At 1280 × 720 and 100% the window
is ~1,500 cells; the ceiling is crossed only by the viewport itself at low zoom
(below ~45% at 1280 × 720, ~65% at 1920 × 1080), where text is too small to edit
and the overview canvas takes over: it repaints at most once per animation frame
from `GridMetrics`, the same style resolution and `formatValue` the cells use,
greeks text whose rendered size is under 6 px as a 1 px bar in its colour, and
returns to DOM cells as soon as the viewport fits again. Row and column headers
follow the same rule.

### Memory and the cell ceiling
`MAX_WORKBOOK_CELLS` stays 2,000,000. The estimate behind it: a lean `Cell` is
~24–40 bytes of V8 heap plus its array slot, a row ~120 bytes, so 500,000 cells
are ~25–40 MB of model, and the retained worksheet part adds its base64 (~33 MB
for this fixture's ~25 MB of XML) — around 70–80 MB on top of the app's ~49 MB
baseline. Four times that is under SPEC-047's 400 MB whole-app ceiling with
little room to spare, which is why the benchmark asserts the projection rather
than trusting this paragraph, why the fixture sits at a quarter of the ceiling
(SPEC-047's 300-of-500-pages rule), and why this spec may lower the number but
never raises it. Formula ASTs do not scale with the cell count: SPEC-119's
`parseFormulaAt` LRU holds one AST for the chain's 9,999 shared-formula cells
(their `relativeFormulaKey`s are equal), and the LRU's 50,000-entry bound caps
the cache for workbooks of distinct formulas. That cache is a memo, so nothing
of it crosses the worker boundary: a parse check the reader runs in the worker
fills the worker's own LRU, which dies with the worker, and the main thread's
fills on first evaluation. The 50 MB file cap still applies first; Google
Sheets' 10M-cell limit is not a target for a browser-resident model.

### Tests
Per the Tests AC. The perf file's header restates SPEC-047's validity notes
(quiet machine, its own project, a full-suite-only red is confirmed by a solo run,
host contention moves every row together).

### Specs this changes
- `specs/116-…` — the per-sheet `sheetrows:` record becomes per-chunk
  `sheetrows:` records in the same store, and `loadAllRecords` skips them;
  `opaque` and `sourceXml` leave the shell for `sheetsource:`, and the shell
  record gains the `sheetStore` manifest; "every read path rejoins" becomes
  lazy residency — `getSpreadsheet(id)` now returns the full model only while
  it is resident (else `undefined`), beside the new `getSpreadsheetShell`,
  `loadSpreadsheet` and `releaseSpreadsheet`; `upgradeSpreadsheetModel` gains
  an optional `read`; `readSpreadsheetBytes` is phased. `specs/120-…` — its "SPEC-137 may move the pass into a worker" is
  answered no; `useRecalc` exposes `data-calc-state`. `specs/122-…` — the
  serialisers yield.
- `specs/047-…` — the benchmark chain gains `perf-workbook`;
  `e2e/perf-baseline.json` gains the `sheet*` rows.
- `specs/138-…` (later; drafts and versions for workbooks) builds on this
  layout: its draft keys are `"draft"` + these keys (`draftsheetrows:…`,
  `draftsheetsource:…`), its latch is this spec's per-chunk latch, and a draft
  or version capture runs only on a resident workbook; nothing here waits for
  it.
- `specs/117-…` — `SheetCellLayer.vue` keys cells by viewport slot, and
  `GRID_MAX_MOUNTED_CELLS = 10_000` with its stored-cells-only fallback is
  replaced by `GRID_MAX_RENDERED_CELLS = 5_000` and the overview canvas.
- `specs/117-…`, `specs/126-…`, `specs/129-…`, `specs/131-…` — the items each
  deferred to this spec move to the Backlog (see Out of scope).
