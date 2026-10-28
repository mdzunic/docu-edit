// The document library (SPEC-003): an in-memory map of document records that
// writes through to IndexedDB. The in-memory map is the source of truth for
// the running session, so every read stays synchronous; persistence is a
// best-effort mirror. When IndexedDB is unavailable or a write fails (e.g.
// quota exceeded), the app keeps working in-memory and the storage status
// flips so the UI can show its warning banner.
import { createDocument, type DocumentModel } from "./documentModel";
import { kindOfFormat, type DocumentKind, type LibraryFormat } from "./documentKinds";
import type { PresentationModel } from "./presentation/model";
import type { MediaModel } from "./media/model";
import type { PackageVariant, Row, SheetPreview, SpreadsheetModel } from "./spreadsheet/model";
import { spreadsheetLocaleFor } from "./spreadsheet/numberFormat/locale";
import { SPREADSHEET_READER_VERSION } from "./spreadsheet/readerVersion";
import { buildPreview } from "./spreadsheet/preview";
import type { OpenedFile } from "./openDocument";
import { loadPreferences, MARGIN_PRESETS, setPreferencesStorageFailureHandler } from "./preferences";
import { PAPER_SIZES } from "./core/pageSetup";
import { encodePendingMediaInBackground, type MediaPart } from "./core/media";
import type { OpaquePart } from "./core/packageParts";
import { normaliseShapeFill } from "./core/shapes";
import { deleteRecord, deleteRecordRange, getRecord, loadAllRecords, putRecord, writeBatch, type WriteBatchOp } from "./idb";
import {
  DRAFT_SHEET_ROWS_KEY_PREFIX,
  DRAFT_SHEET_SOURCE_KEY_PREFIX,
  draftRowsKey,
  draftRowsRange,
  draftSourceKey,
  emptyLatch,
  joinWorkbook,
  latchMatches,
  legacySheetRowsKey,
  legacyWorkbookShell,
  planWorkbookWrite,
  sameRows,
  sameSource,
  sheetRowsChunkKey,
  SHEET_ROWS_KEY_PREFIX,
  SHEET_SOURCE_KEY_PREFIX,
  sourceLatchOf,
  sourceRecordOf,
  workbookRowsRange,
  workbookShell,
  workbookSourceKey,
  workbookStoredBytes,
  type SheetRowsChunkRecord,
  type SheetRowsRecord,
  type SheetStoreManifest,
  type SourceLatch,
  type WorkbookLatch,
  type WorkbookSourceRecord,
} from "./workbookStorage";
import { chunkRows, rowsInChunk } from "./spreadsheet/rowChunks";
import type { SheetViewState } from "./spreadsheet/selection";

// SPEC-137's storage layout lives in workbookStorage.ts; the store re-exports
// its names so callers keep one import.
export {
  DRAFT_SHEET_ROWS_KEY_PREFIX,
  DRAFT_SHEET_SOURCE_KEY_PREFIX,
  draftRowsKey,
  SHEET_ROWS_CHUNK,
  SHEET_ROWS_KEY_PREFIX,
  SHEET_SOURCE_KEY_PREFIX,
} from "./workbookStorage";
export type { SheetRowsChunkRecord, SheetRowsRecord, SheetStoreManifest, WorkbookSourceRecord } from "./workbookStorage";
import { runWhenIdle, yieldToMainThread } from "./core/idleScheduler";
import type { CaretPosition } from "./textEditing";

/** Where a document's content last came from or went to. "new" means it was
 * created in the app and has never been opened from or saved as a file.
 * Spans every kind the library holds (see documentKinds.ts). */
export type DocumentFormat = LibraryFormat;

/** Everything a library record's `model` can hold: the word processor's
 * document tree, a presentation, a media document or a workbook.
 * Discriminated by `kind` — `DocumentModel` is the one without it. */
export type StoredModel = DocumentModel | PresentationModel | MediaModel | SpreadsheetModel;

/** Narrows a stored model to the presentation product's. */
export function isPresentationModel(model: StoredModel): model is PresentationModel {
  return "kind" in model && model.kind === "presentation";
}

/** Narrows a stored model to the media product's (SPEC-107). */
export function isMediaModel(model: StoredModel): model is MediaModel {
  return "kind" in model && model.kind === "media";
}

/** Narrows a stored model to the spreadsheet product's (SPEC-116). */
export function isSpreadsheetModel(model: StoredModel): model is SpreadsheetModel {
  return "kind" in model && model.kind === "spreadsheet";
}

/** What the library persists per document: the model plus its metadata.
 * This is the COMMITTED record (SPEC-010): the last clean state, which the
 * recovery bar's "Discard" restores. */
export interface StoredDocumentRecord {
  id: string;
  format: DocumentFormat;
  /** Epoch milliseconds of the last modification. */
  lastModified: number;
  model: StoredModel;
  /** Caret position at the last commit, so reopening restores it without
   * reconstruction (SPEC-010). Absent on records from before autosave. */
  caret?: CaretPosition | null;
  /** Canvas scrollTop in pixels at the last commit. */
  scroll?: number;
  /** A deck's selected slide at the last commit (SPEC-103) — what "Discard"
   * re-selects. Absent on text documents and on decks committed before it. */
  slideIndex?: number;
  /** A workbook's tab, active cell and top-left visible cell at the last
   * commit (SPEC-138) — what reopening and "Discard" put back. */
  sheetView?: SheetViewState;
  /** Starred by the user (SPEC-067). Absent means not starred. */
  starred?: boolean;
  /** When the document was moved to the trash (SPEC-067), as an epoch
   * millisecond value. Absent means it is in the library. A record with this
   * set is excluded from listDocuments() and from every enumeration; it keeps
   * its id and content, so restoreDocument returns it intact. */
  trashedAt?: number;
  /** A workbook shell's manifest (SPEC-137 "The storage layout"): the chunk
   * records storage holds per sheet and the estimated stored bytes. Absent
   * on a workbook never persisted and on a SPEC-116-era record (whose rows
   * live in per-sheet records and sources inline — migrated on load). */
  sheetStore?: SheetStoreManifest;
}

/** The autosave journal (SPEC-010): written on the 2 s / 30 s cadence,
 * stored beside the committed record under a prefixed key. A draft newer
 * than its committed record on open means the previous session did not end
 * with a clean close — the crash signal the recovery bar keys on. Both
 * editing products keep one (SPEC-103 lifted the deck's guard). */
export interface DraftRecord {
  /** The IndexedDB key: DRAFT_KEY_PREFIX + docId. */
  id: string;
  docId: string;
  /** Epoch milliseconds when the draft was written. */
  savedAt: number;
  model: StoredModel;
  caret: CaretPosition | null;
  scroll: number;
  /** The deck's selected slide when the draft was written (SPEC-103). */
  slideIndex?: number;
  /** The workbook's view when the draft was written (SPEC-138). */
  sheetView?: SheetViewState;
  /** A workbook draft's manifest (SPEC-138): per sheet id, the ordered keys
   * of the records whose `rows` concatenate to that sheet's rows — a
   * committed chunk key for a chunk unchanged since the last commit, its
   * draft key otherwise. The only thing loadDraft trusts. */
  rowKeys?: Record<string, string[]>;
  /** The record a workbook draft's retained sources rejoin from:
   * `sheetsource:<docId>` (unchanged since the last commit) or
   * `draftsheetsource:<docId>`. */
  sourceKey?: string;
}

/** The editor state a draft or a commit records beside the model (SPEC-010,
 * widened by SPEC-103 for decks and SPEC-138 for workbooks). A deck has no
 * caret and its scroll is a fit-to-workspace scale, not a position — so a
 * deck's state is `{ caret: null, scroll: 0, slideIndex }`, and a workbook's
 * `{ caret: null, scroll: 0, sheetView }`. */
export interface DraftState {
  caret: CaretPosition | null;
  scroll: number;
  slideIndex?: number;
  sheetView?: SheetViewState;
}

/** A document's images, persisted as their own record (SPEC-047 AC-15): a
 * committed record or draft's `model.media` is stripped to `[]` before
 * either goes through `putRecord`, because `IDBObjectStore.put()`
 * structured-clones its argument synchronously, and the fixture's ~500
 * images (~50MB of base64) turned every metadata-only write (a title
 * rename, a 2s autosave tick, the download flow's own `touchDocument`) into
 * a multi-hundred-ms main-thread block. Split out here and written only
 * when `media`'s array reference actually changed (see `persistMedia`) —
 * images.ts never mutates `model.media` in place (SPEC-024's own
 * invariant), so an edit that never touches an image keeps the same
 * reference and this record is skipped entirely. */
export interface MediaRecord {
  /** The IndexedDB key: MEDIA_KEY_PREFIX + docId. */
  id: string;
  docId: string;
  media: MediaPart[];
}

/** A deck's retained source package (`PresentationModel.opaque`), persisted
 * as its own record exactly as `media` is (SPEC-103): every source part,
 * verbatim and base64 — a 7 MB package is ≈ 9.4 MB here — which left in the
 * shell would be structured-cloned on every 2 s autosave tick. Nothing
 * mutates `opaque` at read or edit, so its reference never changes for a
 * deck's whole open: written once per session (see persistOpaque), skipped
 * thereafter. Text documents keep their `opaque` inline. */
export interface OpaqueRecord {
  /** The IndexedDB key: OPAQUE_KEY_PREFIX + docId. */
  id: string;
  docId: string;
  opaque: OpaquePart[];
}

/** Everything the documents object store can hold. A workbook (SPEC-137) is
 * a slim shell, one SheetRowsChunkRecord per 512-row band of each sheet and
 * one WorkbookSourceRecord; a SPEC-116-era workbook still awaiting migration
 * has a SheetRowsRecord per sheet. */
export type PersistedRecord =
  | StoredDocumentRecord
  | DraftRecord
  | MediaRecord
  | OpaqueRecord
  | SheetRowsRecord
  | SheetRowsChunkRecord
  | WorkbookSourceRecord;

export const DRAFT_KEY_PREFIX = "draft:";
export const MEDIA_KEY_PREFIX = "media:";
/** A DRAFT's media copy (see persistDraftMedia): its own key, because the
 * committed record rejoins from MEDIA_KEY_PREFIX and an autosave tick must
 * never be able to overwrite the media the committed state depends on —
 * "Discard" after a crash restores the committed shell, and if drafts
 * shared the committed media record, a draft that deleted an image would
 * already have destroyed the committed copy of it. */
export const DRAFT_MEDIA_KEY_PREFIX = "draftmedia:";
/** A deck's committed package record (SPEC-103, see OpaqueRecord). */
export const OPAQUE_KEY_PREFIX = "opaque:";
/** A DRAFT's package copy, under its own key for DRAFT_MEDIA_KEY_PREFIX's
 * reason: an autosave tick must never overwrite the package the committed
 * state rejoins from. */
export const DRAFT_OPAQUE_KEY_PREFIX = "draftopaque:";
/** A document's library search index record (SPEC-114) — a cache of its
 * text, never read or written here except to delete it with its document.
 * idb.ts excludes the prefix from the start-up read. */
export const INDEX_KEY_PREFIX = "index:";

function draftKey(docId: string): string {
  return DRAFT_KEY_PREFIX + docId;
}

function mediaKey(docId: string): string {
  return MEDIA_KEY_PREFIX + docId;
}

function draftMediaKey(docId: string): string {
  return DRAFT_MEDIA_KEY_PREFIX + docId;
}

function opaqueKey(docId: string): string {
  return OPAQUE_KEY_PREFIX + docId;
}

function draftOpaqueKey(docId: string): string {
  return DRAFT_OPAQUE_KEY_PREFIX + docId;
}

export function isDraftRecord(record: PersistedRecord): record is DraftRecord {
  return record.id.startsWith(DRAFT_KEY_PREFIX);
}

export function isMediaRecord(record: PersistedRecord): record is MediaRecord {
  return record.id.startsWith(MEDIA_KEY_PREFIX);
}

export function isDraftMediaRecord(record: PersistedRecord): record is MediaRecord {
  return record.id.startsWith(DRAFT_MEDIA_KEY_PREFIX);
}

export function isOpaqueRecord(record: PersistedRecord): record is OpaqueRecord {
  return record.id.startsWith(OPAQUE_KEY_PREFIX);
}

export function isDraftOpaqueRecord(record: PersistedRecord): record is OpaqueRecord {
  return record.id.startsWith(DRAFT_OPAQUE_KEY_PREFIX);
}

/** A workbook rows record: a chunk (SPEC-137) or a SPEC-116 per-sheet one. */
export function isSheetRowsRecord(record: PersistedRecord): record is SheetRowsRecord | SheetRowsChunkRecord {
  return record.id.startsWith(SHEET_ROWS_KEY_PREFIX);
}

export function isSheetRowsChunkRecord(record: PersistedRecord): record is SheetRowsChunkRecord {
  return isSheetRowsRecord(record) && "chunk" in record && typeof record.chunk === "number";
}

export function isWorkbookSourceRecord(record: PersistedRecord): record is WorkbookSourceRecord {
  return record.id.startsWith(SHEET_SOURCE_KEY_PREFIX);
}

/** A workbook draft's chunk or source record (SPEC-138) — never a library
 * entry, and outside the start-up read's key ranges anyway. */
export function isDraftSheetRecord(record: PersistedRecord): boolean {
  return record.id.startsWith(DRAFT_SHEET_ROWS_KEY_PREFIX) || record.id.startsWith(DRAFT_SHEET_SOURCE_KEY_PREFIX);
}

/** Row-level metadata the library list renders; no document content. */
export interface LibraryEntry {
  id: string;
  title: string;
  format: DocumentFormat;
  /** Which product the entry belongs to — drives the row's route and the
   * kind-specific card details. Derived from `format` (see documentKinds). */
  kind: DocumentKind;
  lastModified: number;
  /** Whether the user has starred the document (SPEC-067). */
  starred: boolean;
  /** Epoch ms the document was trashed — present only on entries returned by
   * listTrashedDocuments(); always absent from listDocuments(). */
  trashedAt?: number;
  /** A workbook's package variant (SPEC-116) — its badge reads XLSM, XLTX,
   * XLTM or OTS by it; absent for every other entry. */
  packageVariant?: PackageVariant;
}

export const MAX_TITLE_LENGTH = 200;

const records = new Map<string, StoredDocumentRecord>();

// --- storage status --------------------------------------------------------

let storageAvailable = true;
const storageListeners = new Set<() => void>();

/** False once loading or a write has failed; it never flips back — the
 * warning banner is persistent for the rest of the session. */
export function isStorageAvailable(): boolean {
  return storageAvailable;
}

/** Notifies when the storage status changes. Returns an unsubscribe. */
export function onStorageStatusChange(listener: () => void): () => void {
  storageListeners.add(listener);
  return () => storageListeners.delete(listener);
}

/** Latches the storage-unavailable state and shows the SPEC-003 warning
 * banner. Exported so the preferences module (SPEC-066) surfaces the same
 * warning when a preference write hits a blocked or full `localStorage`. */
export function markStorageUnavailable(): void {
  if (!storageAvailable) {
    return;
  }
  storageAvailable = false;
  for (const listener of storageListeners) {
    listener();
  }
}

// Wire the preferences store's SPEC-003 warning hook (AC-21) to the same
// banner state, without preferences importing documentStore (which would
// create an initialisation cycle).
setPreferencesStorageFailureHandler(markStorageUnavailable);

// The media array reference last written for each doc, so persistMedia can
// skip a rewrite when nothing image-related changed (see MediaRecord's own
// comment for the invariant this relies on). Seeded from storage on load
// (initLibrary) so the first post-load write doesn't pay one redundant
// round trip.
const lastPersistedMedia = new Map<string, MediaPart[] | undefined>();

/** A real macrotask yield (SPEC-047 AC-15/AC-16) before handing a record to
 * `putRecord`: `IDBObjectStore.put()` structured-clones its argument
 * synchronously, and `putRecord`'s own `await getDb()` resolves via
 * microtask once the db handle is cached (true for every write after the
 * first) — a microtask never hands control back to the browser's
 * input/render queue, so without a real yield first, that clone runs in
 * the SAME task as whatever triggered the write (open, a keystroke, a
 * save's `touchDocument`). Same class of gap zipWriter.ts's `deflateRaw`
 * comment documents for a different await. */
async function putRecordYielding(record: PersistedRecord): Promise<void> {
  await yieldToMainThread();
  await putRecord(record);
}

/** Writes a doc's media as its own record — but only when `media`'s array
 * reference differs from what was last written for it. Returns the write's
 * promise so a caller that must reject on storage failure (saveDraft,
 * commitDocument) can await it; callers that are already fire-and-forget
 * (persist) can leave it unawaited. Returns null when skipped — decided
 * synchronously, before any yield, so callers can tell "skipped" from "in
 * flight" without awaiting either way. */
function persistMedia(docId: string, media: MediaPart[] | undefined): Promise<void> | null {
  if (lastPersistedMedia.get(docId) === media) {
    return null;
  }
  lastPersistedMedia.set(docId, media);
  // SPEC-047: the interactive open defers a large `.docx`'s image encode off
  // the first-paint path (only realised pages force-encode as they render).
  // This first idle persist is the background finaliser: encode any parts
  // still deferred, in byte-bounded chunks with a yield between them so the
  // ~1s whole-document encode never lands as one long task (AC-16), then
  // write the complete media record — reload must never find empty `data`.
  const record: MediaRecord = { id: mediaKey(docId), docId, media: media ?? [] };
  return encodePendingMediaInBackground(media)
    .then(() => putRecordYielding(record))
    .catch((error: unknown) => {
    // The latch was set optimistically before the write; a failed write must
    // clear it, or every later persist would skip the media (reference
    // unchanged, "already persisted") while still stripping the shell —
    // leaving the images nowhere in storage even after space frees up.
    if (lastPersistedMedia.get(docId) === media) {
      lastPersistedMedia.delete(docId);
    }
    throw error;
  });
}

// The draft's own copy of the latch — see persistDraftMedia.
const lastDraftPersistedMedia = new Map<string, MediaPart[] | undefined>();


/** The draft-flavoured persistMedia (SPEC-047 AC-15): writes the draft's
 * media under DRAFT_MEDIA_KEY_PREFIX so an autosave can never overwrite the
 * committed media record ("Discard" restores the committed shell, which
 * rejoins from mediaKey — that record must only ever be written by
 * persist/commitDocument). When the draft's media IS the committed one
 * (reference-equal to the last committed write — the overwhelmingly common
 * case, most edits touch no image), nothing draft-specific is stored and
 * any stale draft copy is dropped so loadDraft falls back to the committed
 * record. */
function persistDraftMedia(docId: string, media: MediaPart[] | undefined): Promise<void> | null {
  if (lastPersistedMedia.get(docId) === media) {
    if (lastDraftPersistedMedia.has(docId)) {
      lastDraftPersistedMedia.delete(docId);
      deleteRecord(draftMediaKey(docId)).catch(() => {});
    }
    return null;
  }
  if (lastDraftPersistedMedia.get(docId) === media) {
    return null;
  }
  lastDraftPersistedMedia.set(docId, media);
  // SPEC-047: as persistMedia — a draft autosave must store real image bytes,
  // never a still-deferred part's empty `data`. Chunked so it never blocks.
  const record: MediaRecord = { id: draftMediaKey(docId), docId, media: media ?? [] };
  return encodePendingMediaInBackground(media)
    .then(() => putRecordYielding(record))
    .catch((error: unknown) => {
    // Same failed-write rule as persistMedia's latch above.
    if (lastDraftPersistedMedia.get(docId) === media) {
      lastDraftPersistedMedia.delete(docId);
    }
    throw error;
  });
}

// A deck's package latch (SPEC-103) — persistMedia's rule for `opaque`: the
// array reference last written for each doc, seeded on load only when an
// OpaqueRecord exists (initLibrary's legacy-inline rule).
const lastPersistedOpaque = new Map<string, OpaquePart[] | undefined>();

/** Writes a deck's package as its own record — only when `opaque`'s array
 * reference differs from what was last written for it; null when skipped
 * (decided synchronously, like persistMedia). The latch is set optimistically
 * and cleared again if the write fails, so the next persist retries it. */
function persistOpaque(docId: string, opaque: OpaquePart[] | undefined): Promise<void> | null {
  if (lastPersistedOpaque.get(docId) === opaque) {
    return null;
  }
  lastPersistedOpaque.set(docId, opaque);
  const record: OpaqueRecord = { id: opaqueKey(docId), docId, opaque: opaque ?? [] };
  return putRecordYielding(record).catch((error: unknown) => {
    // persistMedia's failed-write rule: never leave a write that did not
    // land latched as "already persisted".
    if (lastPersistedOpaque.get(docId) === opaque) {
      lastPersistedOpaque.delete(docId);
    }
    throw error;
  });
}

// The draft's own copy of the package latch — see persistDraftOpaque.
const lastDraftPersistedOpaque = new Map<string, OpaquePart[] | undefined>();

/** The draft-flavoured persistOpaque, persistDraftMedia's rule exactly: while
 * the draft's package IS the committed one (every tick of an ordinary session
 * — nothing edits `opaque`), nothing draft-specific is stored and any stale
 * draft copy is dropped; otherwise it writes once under
 * DRAFT_OPAQUE_KEY_PREFIX and skips while the reference holds. */
function persistDraftOpaque(docId: string, opaque: OpaquePart[] | undefined): Promise<void> | null {
  if (lastPersistedOpaque.get(docId) === opaque) {
    if (lastDraftPersistedOpaque.has(docId)) {
      lastDraftPersistedOpaque.delete(docId);
      deleteRecord(draftOpaqueKey(docId)).catch(() => {});
    }
    return null;
  }
  if (lastDraftPersistedOpaque.get(docId) === opaque) {
    return null;
  }
  lastDraftPersistedOpaque.set(docId, opaque);
  const record: OpaqueRecord = { id: draftOpaqueKey(docId), docId, opaque: opaque ?? [] };
  return putRecordYielding(record).catch((error: unknown) => {
    if (lastDraftPersistedOpaque.get(docId) === opaque) {
      lastDraftPersistedOpaque.delete(docId);
    }
    throw error;
  });
}

/** The package write a model's persist owes: a deck's `opaque` through its
 * latch; nothing for every other kind (a text document keeps its `opaque`
 * inline). */
function persistModelOpaque(docId: string, model: StoredModel): Promise<void> | null {
  return isPresentationModel(model) ? persistOpaque(docId, model.opaque) : null;
}

/** A model with `media` stripped to `[]` for the shell record `persist`/
 * `saveDraft`/`commitDocument` hand to `putRecord` — the array itself
 * always goes through `persistMedia` instead. */
function stripMedia<M extends StoredModel>(model: M): M {
  return { ...model, media: [] };
}

/** A deck's model with `opaque` stripped to `[]` for the shell, as
 * stripMedia does for `media` (the package goes through persistOpaque). A
 * deck with no package (a from-scratch deck) and every other kind pass
 * through unchanged. */
function stripOpaque<M extends StoredModel>(model: M): M {
  return isPresentationModel(model) && model.opaque !== undefined ? { ...model, opaque: [] } : model;
}

/** Rejoins a deck's package (read back via `opaqueKey`/`draftOpaqueKey`) into
 * a shell that had it stripped — the inverse of stripOpaque. Undefined (no
 * OpaqueRecord) leaves the shell's own `opaque`: a legacy-inline record's
 * package, or a from-scratch deck's absence of one. */
function rejoinOpaque(model: StoredModel, opaque: OpaquePart[] | undefined): StoredModel {
  if (opaque !== undefined && isPresentationModel(model)) {
    model.opaque = opaque;
  }
  return model;
}

/** Rejoins a doc's media (read back via `mediaKey`) into a model whose
 * stored shell had it stripped — the inverse of `stripMedia`, applied on
 * every read path (initLibrary, refreshDocumentFromStorage, loadDraft).
 * Absent when nothing was ever persisted for it (a document with no
 * images never gets a MediaRecord — see persistMedia's skip). */
function rejoinMedia(model: StoredModel, media: MediaPart[] | undefined): StoredModel {
  normaliseStoredShapeFills(model);
  if (media === undefined || media.length === 0) {
    return model;
  }
  model.media = media;
  return model;
}

/** The shell record's model: media stripped (stripMedia), and a deck's
 * package too (stripOpaque). A workbook's shell goes through writeWorkbook
 * instead (SPEC-137). */
function shellModel<M extends StoredModel>(model: M): M {
  return stripOpaque(stripMedia(model));
}

/** A committed shell write that waits for this persist's own package write
 * (SPEC-103): once it lands the shell goes out stripped; if it failed, the
 * shell keeps the package inline instead — the legacy layout initLibrary
 * migrates on a later persist — so storage never holds a stripped deck whose
 * package is nowhere, which would reload and save as a from-scratch rebuild
 * and destroy SPEC-008's preservation. With no package write owed it is
 * exactly the plain shell write. */
async function putShell(record: StoredDocumentRecord, opaqueWrite: Promise<void> | null): Promise<void> {
  const split = opaqueWrite === null || (await opaqueWrite.then(() => true, () => false));
  await putRecordYielding({ ...record, model: split ? shellModel(record.model) : stripMedia(record.model) });
}

/** A workbook's card preview in the user's number locale (SPEC-118: the
 * app layer hands buildPreview the current preference; a later change of it
 * leaves stored previews alone until their next rebuild). */
function workbookPreview(model: SpreadsheetModel): SheetPreview {
  return buildPreview(model, spreadsheetLocaleFor(loadPreferences().spreadsheetLocale));
}

// --- workbook residency and persistence (SPEC-137) --------------------------
//
// A workbook is RESIDENT while its record's model is the whole workbook —
// rows and retained sources in memory; otherwise the record holds the slim
// shell (workbookStorage.ts), which is all the library, the guards and the
// cards ever read. loadSpreadsheet makes it resident (one get per chunk,
// yielding between them), releaseSpreadsheet drops rows and sources again
// once storage holds them. A persist of a resident workbook writes exactly
// what changed since the last one — the chunks whose slice differs by
// identity, the deletions, the source record when a source changed — and
// the shell, in ONE writeBatch; a workbook that is not resident writes its
// shell alone, never rows or sources.

interface WorkbookState {
  resident: boolean;
  /** Stored in SPEC-116's layout (no manifest): the next persist of the
   * resident model migrates it, in the same one batch. */
  legacy: boolean;
  /** The load found a record the manifest names missing: read-only, and
   * never written back (WORKBOOK_STORAGE_INCOMPLETE_MESSAGE). */
  incomplete: boolean;
  /** Bumped by every load request and store: a release that started before
   * the latest one stands down. */
  requests: number;
  releasing: boolean;
  /** What storage holds (see WorkbookLatch); meaningful while resident. */
  latch: WorkbookLatch;
  loading: Promise<SpreadsheetModel | undefined> | null;
  /** The persist running, and the one queued behind it (persists of one
   * workbook never overlap, so the latch always describes storage). */
  persisting: Promise<boolean> | null;
  queued: Promise<boolean> | null;
  /** SPEC-138: bumped by every replaceSpreadsheetModel (each edit, a
   * recovered draft, Discard, a restore); `committedEdits` is its value when
   * the last persist was planned. While they differ the resident model holds
   * changes only the draft has — a view-state write keeps them in memory and
   * a release drops them back to what storage commits (uncommitted()). */
  edits: number;
  committedEdits: number;
}

const workbooks = new Map<string, WorkbookState>();

function workbookState(id: string, resident = false): WorkbookState {
  let state = workbooks.get(id);
  if (!state) {
    state = {
      resident,
      legacy: false,
      incomplete: false,
      requests: 0,
      releasing: false,
      latch: emptyLatch(),
      loading: null,
      persisting: null,
      queued: null,
      edits: 0,
      committedEdits: 0,
    };
    workbooks.set(id, state);
  }
  return state;
}

/** Whether the resident model holds changes the committed records do not
 * (SPEC-138: an edit since the last commit, or recovered draft content). */
function uncommitted(state: WorkbookState): boolean {
  return state.edits !== state.committedEdits;
}

/** Whether storage holds everything the resident model has: the latch
 * matches it (and, in the chunked layout, a manifest exists). */
function storageHolds(state: WorkbookState, record: StoredDocumentRecord): boolean {
  if (!isSpreadsheetModel(record.model) || (!state.legacy && record.sheetStore === undefined)) {
    return false;
  }
  return latchMatches(state.latch, record.model);
}

/** Queues one persist of workbook `id` behind the running one (a persist
 * already queued is shared: it reads the record when it starts). Resolves
 * true once written, false on a storage failure (already latched). */
function persistWorkbook(id: string): Promise<boolean> {
  const state = workbookState(id);
  if (state.queued) {
    return state.queued;
  }
  const after = state.persisting ?? Promise.resolve(true);
  const run: Promise<boolean> = after.then(async () => {
    state.queued = null;
    state.persisting = run;
    try {
      return await writeWorkbook(id, state);
    } finally {
      if (state.persisting === run) {
        state.persisting = null;
      }
    }
  });
  state.queued = run;
  return run;
}

/** One persist (see persistWorkbook): plans against the latch and commits
 * the plan and the shell in one writeBatch; the latch and manifest advance
 * only once it committed. */
async function writeWorkbook(id: string, state: WorkbookState): Promise<boolean> {
  // SPEC-047 AC-15/16: never clone in the task that triggered the write.
  await yieldToMainThread();
  const record = records.get(id);
  if (!record || !isSpreadsheetModel(record.model) || workbooks.get(id) !== state || state.incomplete) {
    // Gone, or a partial model that must never be written back.
    return true;
  }
  const model = record.model;
  const edits = state.edits;
  let ops: WriteBatchOp[];
  let adopt: (() => void) | null = null;
  if (!state.resident) {
    // Only the shell: the rows and sources stay exactly as storage holds
    // them (a 116-era shell keeps its inline sources).
    ops = [{ put: { ...record, model: stripMedia(state.legacy ? legacyWorkbookShell(model) : workbookShell(model)) } }];
  } else {
    const plan = planWorkbookWrite(id, model, { manifest: record.sheetStore, latch: state.latch, legacy: state.legacy });
    if (plan.rowsChanged) {
      model.preview = workbookPreview(model);
    }
    const manifest: SheetStoreManifest = { chunks: plan.chunks, storedBytes: workbookStoredBytes(model) };
    ops = [...plan.ops, { put: { ...record, model: stripMedia(workbookShell(model)), sheetStore: manifest } }];
    adopt = () => {
      state.latch = plan.latch;
      state.legacy = false;
      state.committedEdits = edits;
      record.sheetStore = manifest;
    };
  }
  try {
    await writeBatch(ops);
    adopt?.();
    return true;
  } catch {
    markStorageUnavailable();
    return false;
  }
}

/** A stored record by key; a read that fails counts as not returned. */
async function readStored(key: string): Promise<PersistedRecord | undefined> {
  try {
    return await getRecord(key);
  } catch {
    return undefined;
  }
}

/** Reads a workbook's rows and sources back (one get per record, a yield
 * between them) and joins them into its shell. */
async function readWorkbook(
  id: string,
  shell: SpreadsheetModel,
  manifest: SheetStoreManifest | undefined,
): Promise<{ model: SpreadsheetModel; latch: WorkbookLatch; incomplete: boolean }> {
  const rows = new Map<string, Row[]>();
  let incomplete = false;
  if (!manifest) {
    // SPEC-116's layout: each sheet's rows in its own record, the sources
    // inline in the shell. An absent record leaves the shell's rows (a
    // record stored before rows were split out); a failed read is missing.
    for (const sheet of shell.sheets) {
      await yieldToMainThread();
      try {
        const stored = await getRecord(legacySheetRowsKey(id, sheet.id));
        rows.set(sheet.id, stored && isSheetRowsRecord(stored) ? stored.rows : sheet.rows);
      } catch {
        incomplete = true;
        rows.set(sheet.id, sheet.rows);
      }
    }
    const model = { ...shell, sheets: shell.sheets.map((sheet) => ({ ...sheet, rows: rows.get(sheet.id) ?? [] })) };
    return { model, latch: { rows, source: sourceLatchOf(model) }, incomplete };
  }
  const stored = await readStored(workbookSourceKey(id));
  const source = stored && isWorkbookSourceRecord(stored) ? stored : undefined;
  incomplete = source === undefined;
  for (const sheet of shell.sheets) {
    const parts: Row[][] = [];
    for (const chunk of manifest.chunks[sheet.id] ?? []) {
      await yieldToMainThread();
      const record = await readStored(sheetRowsChunkKey(id, sheet.id, chunk));
      if (record && isSheetRowsChunkRecord(record)) {
        parts.push(record.rows);
      } else {
        incomplete = true;
      }
    }
    rows.set(sheet.id, parts.length === 1 ? parts[0]! : parts.flat());
  }
  const model = joinWorkbook(shell, rows, source);
  return { model, latch: { rows, source: source ? sourceLatchOf(model) : null }, incomplete };
}

/** Brings a resident workbook up to the current reader version in place
 * (SPEC-116 "Reader versions"): only the model is replaced, so the record
 * keeps its id, starred and trash state, and upgradeSpreadsheetModel keeps
 * the title and every sheet's id. The re-read runs off the main thread
 * (SPEC-137: readSpreadsheetOffThread; both modules load on demand, so the
 * store never pulls the readers into the start-up graph). A record that
 * cannot be re-read stays exactly as it was. Resolves whether it upgraded. */
async function upgradeResidentWorkbook(record: StoredDocumentRecord): Promise<boolean> {
  const model = record.model;
  if (!isSpreadsheetModel(model) || model.readerVersion >= SPREADSHEET_READER_VERSION) {
    return false;
  }
  try {
    const [{ upgradeSpreadsheetModel }, { readSpreadsheetOffThread }] = await Promise.all([
      import("./spreadsheet/open"),
      import("./workers/spreadsheetParseClient"),
    ]);
    const upgraded = await upgradeSpreadsheetModel(model, readSpreadsheetOffThread);
    if (records.get(record.id) === record && record.model === model) {
      record.model = upgraded;
      return true;
    }
  } catch {
    // Kept as it was.
  }
  return false;
}

/** The workbook stored under `id` as its shell — always, resident or not
 * (SPEC-137): the library, the routing guards and the cards read this. A
 * shell that is not resident has no rows and no sources. Undefined for an
 * absent id and for every other kind. */
export function getSpreadsheetShell(id: string): SpreadsheetModel | undefined {
  const model = records.get(id)?.model;
  return model !== undefined && isSpreadsheetModel(model) ? model : undefined;
}

/** Whether the workbook `id` is resident (its rows and sources in memory). */
export function isSpreadsheetResident(id: string): boolean {
  return workbooks.get(id)?.resident === true && getSpreadsheetShell(id) !== undefined;
}

/** Whether the workbook `id`'s last load came back incomplete — some record
 * its manifest names was not returned by storage. The view shows
 * WORKBOOK_STORAGE_INCOMPLETE_MESSAGE and refuses edits; the store never
 * writes that partial model back. */
export function isSpreadsheetStorageIncomplete(id: string): boolean {
  const state = workbooks.get(id);
  return state?.resident === true && state.incomplete;
}

/** Makes the workbook `id` resident and resolves to the whole model: rejoins
 * the source record and every chunk the manifest names (one get per chunk,
 * yieldToMainThread() between them), migrates a SPEC-116-era record (its
 * per-sheet records rejoined, then rewritten as chunks, a source record and
 * a slim shell, the legacy records deleted — one writeBatch), and runs the
 * reader-version upgrade when behind. Idempotent; concurrent callers share
 * one promise. Undefined for an absent id and for every other kind. */
export function loadSpreadsheet(id: string): Promise<SpreadsheetModel | undefined> {
  const record = records.get(id);
  if (!record || !isSpreadsheetModel(record.model)) {
    return Promise.resolve(undefined);
  }
  const state = workbookState(id);
  state.requests += 1;
  if (state.loading) {
    return state.loading;
  }
  if (state.resident && (state.incomplete || record.model.readerVersion >= SPREADSHEET_READER_VERSION)) {
    return Promise.resolve(record.model);
  }
  const loading: Promise<SpreadsheetModel | undefined> = (async () => {
    if (!state.resident) {
      const shell = record.model as SpreadsheetModel;
      const read = await readWorkbook(id, shell, state.legacy ? undefined : record.sheetStore);
      if (records.get(id) !== record || workbooks.get(id) !== state || record.model !== shell) {
        return getSpreadsheet(id);
      }
      record.model = read.model;
      state.resident = true;
      state.incomplete = read.incomplete;
      state.latch = read.latch;
    }
    if (!state.incomplete && ((await upgradeResidentWorkbook(record)) || state.legacy)) {
      void persist(record);
    }
    return getSpreadsheet(id);
  })().finally(() => {
    if (state.loading === loading) {
      state.loading = null;
    }
  });
  state.loading = loading;
  return loading;
}

/** Drops the workbook `id`'s rows and sources from memory, keeping its
 * shell — once its pending persistence has flushed and storage holds what
 * the model has. Stands down when the workbook was requested again
 * meanwhile, and keeps a workbook storage could not take (memory is then its
 * only copy). */
export async function releaseSpreadsheet(id: string): Promise<void> {
  const state = workbooks.get(id);
  if (!state) {
    return;
  }
  const requests = state.requests;
  state.releasing = true;
  try {
    if (state.loading) {
      await state.loading;
    }
    if (!state.resident) {
      return;
    }
    if (!state.incomplete && uncommitted(state)) {
      // SPEC-138: changes no commit took (a session that closed with the
      // recovery bar open, or with Autosave off) live in the draft, never in
      // the committed records — memory drops back to what storage commits.
      await (state.queued ?? state.persisting);
      if (workbooks.get(id) === state && state.requests === requests && state.resident && uncommitted(state)) {
        const stored = await readStored(id);
        if (stored && isCommittedRecord(stored) && isSpreadsheetModel(stored.model) && workbooks.get(id) === state && state.requests === requests) {
          await adoptStoredRecord(id, stored);
        }
      }
      return;
    }
    if (!state.incomplete) {
      await (state.queued ?? state.persisting);
      const current = records.get(id);
      if (current && state.resident && !storageHolds(state, current)) {
        await persistWorkbook(id);
      }
    }
    const record = records.get(id);
    if (
      !record ||
      !isSpreadsheetModel(record.model) ||
      workbooks.get(id) !== state ||
      state.requests !== requests ||
      !state.resident ||
      (!state.incomplete && !storageHolds(state, record))
    ) {
      return;
    }
    record.model = state.legacy ? legacyWorkbookShell(record.model) : workbookShell(record.model);
    state.resident = false;
    state.incomplete = false;
    state.latch = emptyLatch();
  } finally {
    state.releasing = false;
  }
}

/** Load-then-release (SPEC-137): runs `use` over the whole workbook `id`,
 * loading it when it is not resident and releasing it afterwards — unless
 * something else held it resident before, or asked for it meanwhile. The
 * library's download and duplicate and the idle migration go through this. */
export async function withLoadedSpreadsheet<T>(
  id: string,
  use: (model: SpreadsheetModel) => T | Promise<T>,
): Promise<T | undefined> {
  const before = workbooks.get(id);
  const held = before?.resident === true && !before.releasing;
  const model = await loadSpreadsheet(id);
  const requests = workbooks.get(id)?.requests;
  try {
    return model ? await use(model) : undefined;
  } finally {
    if (!held && workbooks.get(id)?.requests === requests) {
      await releaseSpreadsheet(id);
    }
  }
}

/** The idle migration pass (SPEC-137): after initLibrary, each SPEC-116-era
 * workbook nobody opens is loaded (which rewrites it in the chunked layout)
 * and released, one per runWhenIdle callback, so its heavy shell stops being
 * read at start-up. */
function migrateLegacyWorkbooksWhenIdle(ids: readonly string[]): void {
  const queue = [...ids];
  const next = (): void => {
    const id = queue.shift();
    if (id !== undefined) {
      runWhenIdle(() => {
        void migrateLegacyWorkbook(id).finally(next);
      }, 3000);
    }
  };
  next();
}

async function migrateLegacyWorkbook(id: string): Promise<void> {
  const state = workbooks.get(id);
  if (!state?.legacy || state.resident || state.loading) {
    // Deleted, or opened meanwhile — that load migrates it.
    return;
  }
  await withLoadedSpreadsheet(id, () => undefined);
}

/** SPEC-073 migration (AC-30): a model stored by an earlier version carries
 * legacy `"#rrggbb"` string shape fills; normalise them to the `ShapeFill`
 * union here on the document-store read path (rejoinMedia runs on every
 * hydration: initLibrary, refreshDocumentFromStorage, loadDraft), so nothing
 * downstream — commands, render, writers — ever sees the string form.
 * Mutates in place, like rejoinMedia itself: these run on freshly
 * deserialised records nothing else holds yet. */
function normaliseStoredShapeFills(model: StoredModel): void {
  if (isPresentationModel(model) || isMediaModel(model) || isSpreadsheetModel(model) || !Array.isArray(model.sections)) {
    return;
  }
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "shape") {
        const fill = normaliseShapeFill(block.fill);
        if (fill !== block.fill) {
          if (fill) {
            block.fill = fill;
          } else {
            delete block.fill;
          }
        }
      }
    }
  }
}

// Fire-and-forget write-through: a failed write must not break the in-memory
// session — it only flips the storage status.
/** Resolves once both writes have settled. It never rejects — a failure has
 * already latched the storage-unavailable state, which `isStorageAvailable()`
 * reports and the warning banner shows — so a caller that only wants to know
 * "is it written yet" can await this and read that latch afterwards. Callers
 * with nothing to wait for ignore the promise, as they always have. */
function persist(record: StoredDocumentRecord): Promise<void> {
  const media = persistMedia(record.id, record.model.media)?.catch(() => markStorageUnavailable());
  if (isSpreadsheetModel(record.model)) {
    // SPEC-137: the shell, changed chunks and sources go in one batch.
    return Promise.all([media, persistWorkbook(record.id)]).then(() => undefined);
  }
  const opaqueWrite = persistModelOpaque(record.id, record.model);
  const opaque = opaqueWrite?.catch(() => markStorageUnavailable());
  const shell = putShell(record, opaqueWrite).catch(() => markStorageUnavailable());
  return Promise.all([media, opaque, shell]).then(() => undefined);
}

// --- hydration -------------------------------------------------------------

let initPromise: Promise<void> | null = null;

/** Loads persisted documents into the in-memory map. Idempotent, and never
 * rejects: a failure marks storage unavailable and leaves the store empty,
 * so callers can always `await` it and then read synchronously. main.ts
 * awaits this before mounting the app. */
export function initLibrary(): Promise<void> {
  initPromise ??= loadAllRecords()
    .then((stored) => {
      const mediaByDocId = new Map<string, MediaPart[]>();
      const opaqueByDocId = new Map<string, OpaquePart[]>();
      for (const record of stored) {
        // Strictly committed media only — a draft's media copy
        // (isDraftMediaRecord) must never rejoin the committed shell.
        if (!isDraftMediaRecord(record) && isMediaRecord(record)) {
          mediaByDocId.set(record.docId, record.media);
        }
        // The same for a deck's package (SPEC-103).
        if (!isDraftOpaqueRecord(record) && isOpaqueRecord(record)) {
          opaqueByDocId.set(record.docId, record.opaque);
        }
      }
      const legacyWorkbooks: string[] = [];
      for (const record of stored) {
        // Drafts, media and a workbook's rows and source records live in the
        // same object store but are not library entries; the editor loads a
        // draft explicitly via loadDraft, media rejoins its document record
        // below, and a workbook's rows and sources (SPEC-137: outside the
        // start-up read's key ranges anyway) wait for loadSpreadsheet.
        if (
          isDraftRecord(record) ||
          isMediaRecord(record) ||
          isDraftMediaRecord(record) ||
          isOpaqueRecord(record) ||
          isDraftOpaqueRecord(record) ||
          isSheetRowsRecord(record) ||
          isWorkbookSourceRecord(record) ||
          isDraftSheetRecord(record)
        ) {
          continue;
        }
        if (!records.has(record.id)) {
          const media = mediaByDocId.get(record.id);
          rejoinMedia(record.model, media);
          const opaque = opaqueByDocId.get(record.id);
          rejoinOpaque(record.model, opaque);
          if (isSpreadsheetModel(record.model)) {
            // Its shell only — not resident. A shell without a manifest is
            // SPEC-116's layout, migrated by the idle pass or its first load.
            const state = workbookState(record.id);
            state.legacy = record.sheetStore === undefined;
            if (state.legacy) {
              legacyWorkbooks.push(record.id);
            }
          }
          // Seed the skip-map ONLY when a MediaRecord actually exists in
          // storage. A record written before the media split carries its
          // media inline in the shell and has no MediaRecord — seeding from
          // it would make the first persist skip the media write while
          // stripping the shell, silently deleting every image from storage.
          // Left unseeded, that first persist migrates the document instead.
          if (media !== undefined) {
            lastPersistedMedia.set(record.id, record.model.media);
          }
          // The media rule above, for a deck's package (SPEC-103): a deck
          // stored before the split carries `opaque` inline and has no
          // OpaqueRecord. Seeding from it would make the next persist skip
          // the package write while stripping the shell — a reload would then
          // find no package and save as a from-scratch rebuild. Unseeded, the
          // next persist migrates it.
          if (opaque !== undefined && isPresentationModel(record.model)) {
            lastPersistedOpaque.set(record.id, record.model.opaque);
          }
          records.set(record.id, record);
        }
      }
      migrateLegacyWorkbooksWhenIdle(legacyWorkbooks);
    })
    .catch(() => markStorageUnavailable());
  return initPromise;
}

// --- reads -----------------------------------------------------------------

/** Returns the stored text document for an id, or undefined if none exists
 * — including when the id belongs to a presentation (callers of this
 * accessor are the word processor's paths; they must never receive the
 * other product's model). */
export function getDocument(id: string): DocumentModel | undefined {
  const model = records.get(id)?.model;
  return model !== undefined && !isPresentationModel(model) && !isMediaModel(model) && !isSpreadsheetModel(model)
    ? model
    : undefined;
}

/** The presentation counterpart of `getDocument`: undefined for absent ids
 * and for the other kinds alike. */
export function getPresentation(id: string): PresentationModel | undefined {
  const model = records.get(id)?.model;
  return model !== undefined && isPresentationModel(model) ? model : undefined;
}

/** The media counterpart of `getDocument`/`getPresentation` (SPEC-107):
 * undefined for absent ids and for the other kinds alike, so a media id can
 * never leak into a document-only path — nor the reverse. */
export function getMedia(id: string): MediaModel | undefined {
  const model = records.get(id)?.model;
  return model !== undefined && isMediaModel(model) ? model : undefined;
}

/** The workbook counterpart of `getDocument`/`getPresentation`/`getMedia`
 * (SPEC-116): undefined for absent ids and for every other kind — and
 * (SPEC-137) for a workbook that is not resident: only then is the model the
 * whole workbook. getSpreadsheetShell reads any; loadSpreadsheet makes one
 * resident. */
export function getSpreadsheet(id: string): SpreadsheetModel | undefined {
  return isSpreadsheetResident(id) ? getSpreadsheetShell(id) : undefined;
}

/** Which product owns a stored id — the routing guard's question ("is
 * /doc/:id actually a presentation?"); undefined when nothing is stored
 * under the id. */
export function documentKindOf(id: string): DocumentKind | undefined {
  const record = records.get(id);
  return record ? kindOfFormat(record.format) : undefined;
}

/** The full in-memory record (metadata + model) for an id. */
export function getStoredRecord(id: string): StoredDocumentRecord | undefined {
  return records.get(id);
}

/** Applies the "new documents" preferences (SPEC-066) to a freshly-created
 * blank model: font, size, page size, margins and language seed the default
 * styles, and the three per-document app-preference fields (measurement unit,
 * ruler and formatting-mark visibility) are set from their app-level defaults
 * so a new document inherits them (AC-24..27). The last three are written only
 * when they differ from the model's implicit default, so a document created
 * with default preferences is identical to one created before this spec — and
 * an existing document, never passed through here, keeps its own stored value.
 * The model defaults stay the fallback when no preference is set (SPEC-002). */
function applyNewDocumentPreferences(model: DocumentModel): void {
  const prefs = loadPreferences();
  const styles = model.defaultStyles;
  styles.fontFamily = prefs.defaultFontFamily;
  styles.fontSizePt = prefs.defaultFontSizePt;
  styles.language = prefs.defaultLanguage;
  const pageSize = PAPER_SIZES[prefs.defaultPageSizeId as keyof typeof PAPER_SIZES];
  if (pageSize) {
    styles.pageSize = { ...pageSize };
  }
  const margins = MARGIN_PRESETS[prefs.defaultMarginsId];
  if (margins) {
    styles.margins = { ...margins };
  }
  if (prefs.measurementUnit !== "cm") {
    model.measurementUnit = prefs.measurementUnit;
  }
  if (prefs.showRuler !== true) {
    // One "show rulers by default" preference seeds both per-document flags;
    // View → Horizontal/Vertical ruler then toggles them apart.
    model.showRuler = prefs.showRuler;
    model.showVerticalRuler = prefs.showRuler;
  }
  if (prefs.showFormattingMarks !== false) {
    model.showFormattingMarks = prefs.showFormattingMarks;
  }
}

/** Returns the stored document for an id, creating and storing a fresh one
 * if none exists yet. */
export function getOrCreateDocument(id: string): DocumentModel {
  const existing = records.get(id);
  if (existing) {
    if (isPresentationModel(existing.model)) {
      // Unreachable through the views (DocumentView redirects presentation
      // ids to /slides/:id before loading), and creating a text document
      // over a presentation's id would destroy it — so fail loudly.
      throw new Error(`Document ${id} is a presentation`);
    }
    if (isMediaModel(existing.model)) {
      // Same rule for media (SPEC-107): silently overwriting the stored
      // bytes with a blank text document is the one unforgivable failure.
      throw new Error(`Document ${id} is a media file`);
    }
    if (isSpreadsheetModel(existing.model)) {
      // And for workbooks (SPEC-116).
      throw new Error(`Document ${id} is a spreadsheet`);
    }
    return existing.model;
  }
  const model = createDocument(id);
  applyNewDocumentPreferences(model);
  const record: StoredDocumentRecord = {
    id,
    format: "new",
    lastModified: Date.now(),
    model,
  };
  records.set(id, record);
  persist(record);
  return model;
}

function toLibraryEntry(record: StoredDocumentRecord): LibraryEntry {
  const entry: LibraryEntry = {
    id: record.id,
    format: record.format,
    kind: kindOfFormat(record.format),
    lastModified: record.lastModified,
    title: record.model.title,
    starred: record.starred === true,
    trashedAt: record.trashedAt,
  };
  if (isSpreadsheetModel(record.model) && record.model.packageVariant !== undefined) {
    entry.packageVariant = record.model.packageVariant;
  }
  return entry;
}

/** Library rows sorted by last modified, newest first. Excludes trashed
 * documents (SPEC-067): a record with `trashedAt` set is not enumerated
 * here, nor anywhere else the library lists documents. */
export function listDocuments(): LibraryEntry[] {
  return [...records.values()]
    .filter((record) => record.trashedAt === undefined)
    .map(toLibraryEntry)
    .toSorted((a, b) => b.lastModified - a.lastModified);
}

/** Trashed documents only (SPEC-067), most recently trashed first. The
 * counterpart to listDocuments(): together they partition the store on
 * `trashedAt`. */
export function listTrashedDocuments(): LibraryEntry[] {
  return [...records.values()]
    .filter((record) => record.trashedAt !== undefined)
    .map(toLibraryEntry)
    .toSorted((a, b) => (b.trashedAt ?? 0) - (a.trashedAt ?? 0));
}

// --- mutations -------------------------------------------------------------

/** Records that a document was just modified (content already updated on its
 * model, e.g. via the app bar's title editor) and persists it. */
export function touchDocument(id: string): Promise<void> {
  const record = records.get(id);
  if (!record) {
    return Promise.resolve();
  }
  record.lastModified = Date.now();
  return persist(record);
}

/** Returns `title` unless it is already taken, else the first free numbered
 * variant: `title (2)`, `title (3)`, … */
export function uniqueTitle(title: string, takenTitles: string[]): string {
  const taken = new Set(takenTitles);
  if (!taken.has(title)) {
    return title;
  }
  let n = 2;
  while (taken.has(`${title} (${n})`)) {
    n += 1;
  }
  return `${title} (${n})`;
}

/** Renames a document: trims, rejects empty/over-long titles (returns null),
 * deduplicates against the other documents' titles, bumps lastModified and
 * persists. Returns the title actually applied. */
export function renameDocument(
  id: string,
  requestedTitle: string,
): string | null {
  const record = records.get(id);
  if (!record) {
    return null;
  }
  const trimmed = requestedTitle.trim();
  if (trimmed.length === 0 || trimmed.length > MAX_TITLE_LENGTH) {
    return null;
  }
  const otherTitles = [...records.values()]
    .filter((other) => other.id !== id)
    .map((other) => other.model.title);
  record.model.title = uniqueTitle(trimmed, otherTitles);
  record.lastModified = Date.now();
  persist(record);
  return record.model.title;
}

/** Creates an independent copy titled "<title> (copy)" with a fresh id and a
 * lastModified of now (so it sorts to the top of the library). */
export function duplicateDocument(
  id: string,
): StoredDocumentRecord | undefined {
  const source = records.get(id);
  if (!source) {
    return undefined;
  }
  const sourceState = workbooks.get(id);
  if (isSpreadsheetModel(source.model) && (!sourceState?.resident || sourceState.incomplete)) {
    // SPEC-137: a workbook that is not resident is only its shell in memory —
    // cloned, it would store an empty copy — and a partial one must never be
    // written back. The library loads it first (duplicateStoredDocument).
    return undefined;
  }
  const copyId = crypto.randomUUID();
  const model = structuredClone(source.model);
  model.id = copyId;
  model.title = `${source.model.title} (copy)`;
  const copy: StoredDocumentRecord = {
    id: copyId,
    format: source.format,
    lastModified: Date.now(),
    model,
  };
  records.set(copyId, copy);
  if (isSpreadsheetModel(model)) {
    workbookState(copyId, true);
  }
  persist(copy);
  return copy;
}

/** The library's duplicate: a workbook is loaded first (SPEC-137
 * load-then-release), duplicated whole, and released again — its copy too,
 * once written. Every other kind duplicates synchronously. */
export async function duplicateStoredDocument(id: string): Promise<StoredDocumentRecord | undefined> {
  if (getSpreadsheetShell(id) === undefined) {
    return duplicateDocument(id);
  }
  const copy = await withLoadedSpreadsheet(id, () => duplicateDocument(id));
  if (copy) {
    void releaseSpreadsheet(copy.id);
  }
  return copy;
}

/** Moves a document to the trash (SPEC-067): sets `trashedAt` so it drops
 * out of listDocuments() and every enumeration, but keeps its id, model,
 * version history and draft intact so restoreDocument returns it whole. Does
 * NOT bump lastModified — trashing is not an edit — so a restored document
 * sorts back exactly where it was. This is the reversible delete; only
 * "Delete permanently" and "Empty trash" reach deleteDocument. */
export function trashDocument(id: string): void {
  const record = records.get(id);
  if (!record || record.trashedAt !== undefined) {
    return;
  }
  record.trashedAt = Date.now();
  persist(record);
}

/** Restores a trashed document to the library (SPEC-067) by clearing
 * `trashedAt`. The inverse of trashDocument; the toast's Undo and the Trash
 * view's Restore both call it. */
export function restoreDocument(id: string): void {
  const record = records.get(id);
  if (!record || record.trashedAt === undefined) {
    return;
  }
  delete record.trashedAt;
  persist(record);
}

/** Sets (or clears) a document's starred flag (SPEC-067) and persists.
 * Absent means not starred, so clearing removes the field. */
export function setStarred(id: string, starred: boolean): void {
  const record = records.get(id);
  if (!record) {
    return;
  }
  if (starred) {
    record.starred = true;
  } else {
    delete record.starred;
  }
  persist(record);
}

/** How much room the library occupies (SPEC-067): the approximate byte size
 * of every record — trashed included, since a trashed document still costs
 * space until it is permanently deleted — and a per-document breakdown so the
 * storage panel can name the largest. Sizes are derived from the in-memory
 * model (the persisted shell plus its media), which is a close proxy for the
 * IndexedDB footprint without a round trip. Sorted largest first. */
export function libraryStorageUsage(): {
  totalBytes: number;
  byDocument: Array<{ id: string; bytes: number }>;
} {
  let totalBytes = 0;
  const byDocument: Array<{ id: string; bytes: number }> = [];
  for (const record of records.values()) {
    const bytes = approximateRecordBytes(record);
    totalBytes += bytes;
    byDocument.push({ id: record.id, bytes });
  }
  byDocument.sort((a, b) => b.bytes - a.bytes);
  return { totalBytes, byDocument };
}

const byteCounter =
  typeof TextEncoder === "function" ? new TextEncoder() : null;

/** UTF-8 byte length of a document record's persisted content — its model
 * serialised as JSON. Media (base64) dominates when present, which is exactly
 * what makes a document large on disk. Falls back to the string length when
 * TextEncoder is unavailable, and to 0 if the model is not serialisable. */
function approximateRecordBytes(record: StoredDocumentRecord): number {
  // A media document knows its exact source size (SPEC-107) — more honest
  // than the base64-inflated JSON length, and it makes the library's size
  // column read formatFileSize(byteLength) for media rows with no new path.
  if (isMediaModel(record.model)) {
    return record.model.byteLength;
  }
  // A workbook (SPEC-137) is never serialised to be sized: its manifest
  // carries the estimate made when it was last persisted; one never
  // persisted is estimated the same way from what memory holds.
  if (isSpreadsheetModel(record.model)) {
    return record.sheetStore?.storedBytes ?? workbookStoredBytes(record.model);
  }
  return jsonBytes(record.model);
}

/** UTF-8 byte length of a value's JSON (0 when not serialisable). */
function jsonBytes(value: unknown): number {
  try {
    const json = JSON.stringify(value);
    return byteCounter ? byteCounter.encode(json).length : json.length;
  } catch {
    return 0;
  }
}

/** Removes a document from the session and from storage, along with any
 * autosave draft and its search index record. */
export function deleteDocument(id: string): void {
  const model = records.get(id)?.model;
  if (!records.delete(id)) {
    return;
  }
  lastPersistedMedia.delete(id);
  lastDraftPersistedMedia.delete(id);
  lastPersistedOpaque.delete(id);
  lastDraftPersistedOpaque.delete(id);
  if (model && isSpreadsheetModel(model)) {
    // SPEC-137: the shell, every rows record (chunks and any SPEC-116
    // per-sheet record — one range) and the source record, in one batch;
    // SPEC-138: and the draft's chunk range and source copy.
    workbooks.delete(id);
    void deleteWorkbookRecords(id);
    deleteWorkbookDraftRecords(id);
  } else {
    deleteRecord(id).catch(() => markStorageUnavailable());
  }
  // Draft/media/package cleanup is best-effort and not a SPEC-003
  // write-through failure.
  deleteRecord(draftKey(id)).catch(() => {});
  deleteRecord(mediaKey(id)).catch(() => {});
  deleteRecord(draftMediaKey(id)).catch(() => {});
  deleteRecord(opaqueKey(id)).catch(() => {});
  deleteRecord(draftOpaqueKey(id)).catch(() => {});
  // The search index's copy of its text goes with it (SPEC-114) — a delete,
  // flag or no flag: a permanently deleted document leaves no text behind.
  deleteRecord(INDEX_KEY_PREFIX + id).catch(() => {});
}

async function deleteWorkbookRecords(id: string): Promise<void> {
  try {
    await writeBatch([{ delete: id }, { deleteRange: workbookRowsRange(id) }, { delete: workbookSourceKey(id) }]);
  } catch {
    markStorageUnavailable();
  }
}

// --- autosave (SPEC-010) ---------------------------------------------------
//
// Unlike the SPEC-003 write-through above, these reject on failure instead of
// latching the storage banner: the autosave UI owns its own failure state
// ("Not saved — browser storage is full" in the app bar) and recovers the
// moment a write succeeds again.

/** Writes the autosave draft for a document — a text document or a deck
 * (SPEC-103) — from its current in-memory model, with the editor state
 * beside it. Rejects on any storage failure. */
export async function saveDraft(id: string, state: DraftState): Promise<void> {
  const record = records.get(id);
  if (!record) {
    throw new Error(`No document with id ${id}`);
  }
  if (isMediaModel(record.model)) {
    // A media document has no draft state at all (SPEC-107): its bytes are
    // immutable, and an autosave draft would be a path to overwriting them.
    throw new Error(`Document ${id} is a media file`);
  }
  if (isSpreadsheetModel(record.model)) {
    // SPEC-138 lifted SPEC-116's guard: a workbook drafts into records of its
    // own, never the committed ones.
    return saveWorkbookDraft(id, record.model, state);
  }
  // A deck's package and media are skipped while their references hold, so
  // every tick after the first writes the shell alone (SPEC-103).
  const mediaWrite = persistDraftMedia(id, record.model.media);
  const opaqueWrite = isPresentationModel(record.model) ? persistDraftOpaque(id, record.model.opaque) : null;
  const draft: DraftRecord = {
    id: draftKey(id),
    docId: id,
    savedAt: Date.now(),
    model: shellModel(record.model),
    caret: state.caret,
    scroll: state.scroll,
    ...(state.slideIndex !== undefined ? { slideIndex: state.slideIndex } : {}),
  };
  await Promise.all([mediaWrite ?? Promise.resolve(), opaqueWrite ?? Promise.resolve(), putRecordYielding(draft)]);
}

// --- workbook drafts (SPEC-138 "The workbook draft record") -------------------
//
// A workbook draft never touches a committed record. Its shell (`draft:<id>`)
// is SPEC-137's slim shell — rows `[]`, `opaque` `[]`, no source text — plus
// the `rowKeys` manifest and the `sourceKey`. Each 512-row chunk whose slice
// is row-for-row the committed one is listed by its committed key and costs
// nothing; any other is written under its draft key (draftRowsKey) — once,
// under the draft's own per-chunk latch — so a tick after one cell edit costs
// that one chunk. The retained sources are tracked by reference against
// SPEC-137's source latch: while they are the committed ones the shell names
// `sheetsource:<id>`; once one was replaced (SPEC-124/125's retained-reference
// rewrite, a duplicated sheet) the whole set is written once to
// `draftsheetsource:<id>`. A commit and a discard clear the lot in one range.

/** The draft's per-chunk latch: per doc, draft key → the slice last written. */
const lastDraftChunks = new Map<string, Map<string, Row[]>>();
/** Per doc and sheet: the rows array the last tick listed, and its keys — an
 * untouched sheet costs nothing to list again. */
const lastDraftSheets = new Map<string, Map<string, { rows: Row[]; keys: string[] }>>();
/** The draft's source latch: per doc, the sources last written to its copy. */
const lastDraftSource = new Map<string, SourceLatch>();

function forgetWorkbookDraft(docId: string): void {
  lastDraftChunks.delete(docId);
  lastDraftSheets.delete(docId);
  lastDraftSource.delete(docId);
}

/** Deletes every draft chunk record of a workbook in one range — a sheet
 * deleted mid-session leaves none behind — and its draft source copy.
 * Best-effort, like every draft cleanup. */
function deleteWorkbookDraftRecords(docId: string): void {
  forgetWorkbookDraft(docId);
  try {
    const [lower, upper] = draftRowsRange(docId);
    deleteRecordRange(lower, upper).catch(() => {});
    deleteRecord(draftSourceKey(docId)).catch(() => {});
  } catch {
    // Storage without range deletes: the next commit retries.
  }
}

/** What storage commits for the workbook, as far as a draft may lean on it:
 * the chunked layout's latch and manifest (a SPEC-116-era record has no
 * committed chunk to name). */
function committedStorage(state: WorkbookState, record: StoredDocumentRecord): { latch: WorkbookLatch; manifest: SheetStoreManifest } | null {
  return !state.legacy && record.sheetStore ? { latch: state.latch, manifest: record.sheetStore } : null;
}

/** persistDraftMedia's latch rules, per chunk, against SPEC-137's per-chunk
 * committed latch: a slice reference-equal to the committed one is listed by
 * its committed key (and a stale draft copy dropped); one equal to the last
 * draft write is listed by its draft key, unwritten; any other is written,
 * the latch set first and cleared on failure. Returns the writes and the
 * `rowKeys` manifest. No write can reach a committed key. */
function persistDraftSheetRows(
  docId: string,
  model: SpreadsheetModel,
  committed: { latch: WorkbookLatch; manifest: SheetStoreManifest } | null,
): { writes: Promise<void>[]; rowKeys: Record<string, string[]> } {
  const drafted = lastDraftChunks.get(docId) ?? new Map<string, Row[]>();
  lastDraftChunks.set(docId, drafted);
  const listed = lastDraftSheets.get(docId) ?? new Map<string, { rows: Row[]; keys: string[] }>();
  lastDraftSheets.set(docId, listed);
  const writes: Promise<void>[] = [];
  const rowKeys: Record<string, string[]> = {};
  for (const sheet of model.sheets) {
    const committedRows = committed?.latch.rows.get(sheet.id);
    const dropDraftCopy = (key: string): void => {
      if (drafted.delete(draftRowsKey(key))) {
        deleteRecord(draftRowsKey(key)).catch(() => {});
      }
    };
    if (committed && committedRows === sheet.rows) {
      // The sheet's rows are the committed ones: every chunk by its key.
      const keys = (committed.manifest.chunks[sheet.id] ?? []).map((chunk) => sheetRowsChunkKey(docId, sheet.id, chunk));
      keys.forEach(dropDraftCopy);
      rowKeys[sheet.id] = keys;
      listed.set(sheet.id, { rows: sheet.rows, keys });
      continue;
    }
    const previous = listed.get(sheet.id);
    if (previous?.rows === sheet.rows) {
      rowKeys[sheet.id] = previous.keys;
      continue;
    }
    const keys: string[] = [];
    for (const [chunk, slice] of chunkRows(sheet.rows)) {
      const key = sheetRowsChunkKey(docId, sheet.id, chunk);
      if (committedRows !== undefined && sameRows(rowsInChunk(committedRows, chunk), slice)) {
        keys.push(key);
        dropDraftCopy(key);
        continue;
      }
      const draftKeyOfChunk = draftRowsKey(key);
      keys.push(draftKeyOfChunk);
      const last = drafted.get(draftKeyOfChunk);
      if (last !== undefined && sameRows(last, slice)) {
        continue;
      }
      drafted.set(draftKeyOfChunk, slice);
      const chunkRecord: SheetRowsChunkRecord = { id: draftKeyOfChunk, docId, sheetId: sheet.id, chunk, rows: slice };
      writes.push(
        putRecordYielding(chunkRecord).catch((error: unknown) => {
          if (drafted.get(draftKeyOfChunk) === slice) {
            drafted.delete(draftKeyOfChunk);
          }
          if (listed.get(sheet.id)?.rows === sheet.rows) {
            listed.delete(sheet.id);
          }
          throw error;
        }),
      );
    }
    rowKeys[sheet.id] = keys;
    listed.set(sheet.id, { rows: sheet.rows, keys });
  }
  return { writes, rowKeys };
}

/** The retained sources' draft half: while every source is the committed
 * record's (by reference — SPEC-137's source latch), nothing is written and
 * the shell names `sheetsource:<id>`; once one differs the whole set goes to
 * `draftsheetsource:<id>` under its own latch. */
function persistDraftSources(
  docId: string,
  model: SpreadsheetModel,
  committed: { latch: WorkbookLatch } | null,
): { write: Promise<void> | null; sourceKey: string } {
  const current = sourceLatchOf(model);
  const committedSource = committed?.latch.source ?? null;
  if (committedSource && sameSource(committedSource, current)) {
    if (lastDraftSource.delete(docId)) {
      deleteRecord(draftSourceKey(docId)).catch(() => {});
    }
    return { write: null, sourceKey: workbookSourceKey(docId) };
  }
  const key = draftSourceKey(docId);
  const last = lastDraftSource.get(docId);
  if (last && sameSource(last, current)) {
    return { write: null, sourceKey: key };
  }
  lastDraftSource.set(docId, current);
  const write = putRecordYielding(sourceRecordOf(docId, model, key)).catch((error: unknown) => {
    if (lastDraftSource.get(docId) === current) {
      lastDraftSource.delete(docId);
    }
    throw error;
  });
  return { write, sourceKey: key };
}

/** saveDraft for a workbook: its chunk and source records first, then the
 * shell naming them (SPEC-137's residency rule: only a whole workbook is
 * drafted — a shell or an incomplete load never is). Rejects on failure. */
async function saveWorkbookDraft(id: string, model: SpreadsheetModel, state: DraftState): Promise<void> {
  const workbook = workbooks.get(id);
  const record = records.get(id);
  if (!record || !workbook?.resident || workbook.incomplete) {
    throw new Error(`Workbook ${id} is not loaded`);
  }
  const committed = committedStorage(workbook, record);
  const mediaWrite = persistDraftMedia(id, model.media);
  const rows = persistDraftSheetRows(id, model, committed);
  const sources = persistDraftSources(id, model, committed);
  await Promise.all([mediaWrite ?? Promise.resolve(), ...rows.writes, sources.write ?? Promise.resolve()]);
  const draft: DraftRecord = {
    id: draftKey(id),
    docId: id,
    savedAt: Date.now(),
    model: stripMedia(workbookShell(model)),
    caret: null,
    scroll: 0,
    rowKeys: rows.rowKeys,
    sourceKey: sources.sourceKey,
    ...(state.sheetView !== undefined ? { sheetView: state.sheetView } : {}),
  };
  await putRecordYielding(draft);
}

/** A row record a draft manifest may name: a committed or a draft chunk of
 * `sheetId`. */
function isChunkOf(record: PersistedRecord, docId: string, sheetId: string): record is SheetRowsChunkRecord {
  const chunk = record as Partial<SheetRowsChunkRecord>;
  return (
    (record.id.startsWith(SHEET_ROWS_KEY_PREFIX) || record.id.startsWith(DRAFT_SHEET_ROWS_KEY_PREFIX)) &&
    chunk.docId === docId &&
    chunk.sheetId === sheetId &&
    Array.isArray(chunk.rows)
  );
}

/** A workbook draft rejoined from its manifest: every sheet's rows from the
 * records `rowKeys` lists (never an inferred key), the sources from the
 * record `sourceKey` names, then the reader-version upgrade every read path
 * runs. Undefined when the manifest names a record that is missing or
 * unreadable — a half-rejoined workbook is never offered for recovery. */
async function rejoinWorkbookDraft(draft: DraftRecord, shell: SpreadsheetModel): Promise<SpreadsheetModel | undefined> {
  const manifest = draft.rowKeys;
  if (!manifest || draft.sourceKey === undefined) {
    return undefined;
  }
  const rows = new Map<string, Row[]>();
  for (const sheet of shell.sheets) {
    const keys = manifest[sheet.id];
    if (!keys) {
      return undefined;
    }
    const parts: Row[][] = [];
    for (const key of keys) {
      await yieldToMainThread();
      const chunk = await getRecord(key);
      if (!chunk || !isChunkOf(chunk, draft.docId, sheet.id)) {
        return undefined;
      }
      parts.push(chunk.rows);
    }
    rows.set(sheet.id, parts.length === 1 ? parts[0]! : parts.flat());
  }
  const source = await getRecord(draft.sourceKey);
  const sourceRecord = source as Partial<WorkbookSourceRecord> | undefined;
  if (!source || sourceRecord?.docId !== draft.docId || !Array.isArray(sourceRecord.opaque) || typeof sourceRecord.sheetSources !== "object") {
    return undefined;
  }
  return upgradeRecoveredWorkbook(joinWorkbook(shell, rows, source as WorkbookSourceRecord));
}

/** SPEC-116's reader-version upgrade on a recovered draft (a draft written
 * before a reader bump recovers upgraded); kept as it was when it fails. */
async function upgradeRecoveredWorkbook(model: SpreadsheetModel): Promise<SpreadsheetModel> {
  if (model.readerVersion >= SPREADSHEET_READER_VERSION) {
    return model;
  }
  try {
    const [{ upgradeSpreadsheetModel }, { readSpreadsheetOffThread }] = await Promise.all([
      import("./spreadsheet/open"),
      import("./workers/spreadsheetParseClient"),
    ]);
    return await upgradeSpreadsheetModel(model, readSpreadsheetOffThread);
  } catch {
    return model;
  }
}

/** Reads a document's autosave draft; undefined when there is none or the
 * store is unreadable (either way there is nothing to recover). */
export async function loadDraft(id: string): Promise<DraftRecord | undefined> {
  try {
    const stored = await getRecord(draftKey(id));
    if (!stored || !isDraftRecord(stored)) {
      return undefined;
    }
    // A draft whose media diverged from the committed state has its own
    // copy under draftMediaKey; otherwise the committed media record is the
    // right one (persistDraftMedia never wrote, or dropped its stale copy).
    const draftMedia = await getRecord(draftMediaKey(stored.docId));
    if (draftMedia && isDraftMediaRecord(draftMedia)) {
      rejoinMedia(stored.model, draftMedia.media);
    } else {
      const mediaRecord = await getRecord(mediaKey(stored.docId));
      rejoinMedia(stored.model, mediaRecord && isMediaRecord(mediaRecord) ? mediaRecord.media : undefined);
    }
    if (isPresentationModel(stored.model)) {
      rejoinOpaque(stored.model, await draftPackage(stored.docId));
    }
    if (isSpreadsheetModel(stored.model)) {
      const workbook = await rejoinWorkbookDraft(stored, stored.model);
      if (!workbook) {
        return undefined;
      }
      stored.model = workbook;
    }
    return stored;
  } catch {
    return undefined;
  }
}

/** The package a deck draft rejoins (SPEC-103), by persistDraftOpaque's
 * rule: its own copy when it diverged from the committed one, else the
 * committed OpaqueRecord, else — a legacy-inline committed record, whose
 * package was never split out — the committed model's own `opaque`. */
async function draftPackage(docId: string): Promise<OpaquePart[] | undefined> {
  const draftOpaque = await getRecord(draftOpaqueKey(docId));
  if (draftOpaque && isDraftOpaqueRecord(draftOpaque)) {
    return draftOpaque.opaque;
  }
  const committed = await getRecord(opaqueKey(docId));
  if (committed && isOpaqueRecord(committed)) {
    return committed.opaque;
  }
  const memory = records.get(docId)?.model;
  return memory && isPresentationModel(memory) ? memory.opaque : undefined;
}

/** Removes a document's autosave draft. Best-effort: a draft that survives
 * is harmless once the committed record is newer, because recovery only
 * triggers on a draft NEWER than the committed record. */
export async function deleteDraft(id: string): Promise<void> {
  try {
    lastDraftPersistedMedia.delete(id);
    lastDraftPersistedOpaque.delete(id);
    deleteRecord(draftMediaKey(id)).catch(() => {});
    deleteRecord(draftOpaqueKey(id)).catch(() => {});
    if (workbooks.has(id)) {
      deleteWorkbookDraftRecords(id);
    }
    await deleteRecord(draftKey(id));
  } catch {
    // Ignored — see above.
  }
}

/** Promotes the current in-memory state to the committed record — the last
 * clean state "Discard" restores — recording the editor state (caret and
 * scroll; a deck's slide) alongside the model, and clears the draft. Rejects
 * when the committed write fails. */
export async function commitDocument(id: string, state: DraftState): Promise<void> {
  const record = records.get(id);
  if (!record) {
    return;
  }
  record.lastModified = Date.now();
  record.caret = state.caret;
  record.scroll = state.scroll;
  if (state.slideIndex !== undefined) {
    record.slideIndex = state.slideIndex;
  }
  if (state.sheetView !== undefined) {
    record.sheetView = state.sheetView;
  }
  const mediaWrite = persistMedia(id, record.model.media);
  if (isSpreadsheetModel(record.model)) {
    const [written] = await Promise.all([persistWorkbook(id), mediaWrite ?? Promise.resolve()]);
    if (!written) {
      throw new Error(`Workbook ${id} could not be written`);
    }
  } else {
    const opaqueWrite = persistModelOpaque(id, record.model);
    await Promise.all([mediaWrite ?? Promise.resolve(), opaqueWrite ?? Promise.resolve(), putShell(record, opaqueWrite)]);
  }
  await deleteDraft(id);
}

/** Re-reads a document's committed record from storage and adopts it when it
 * is newer than the in-memory one — the case where another tab committed
 * after this tab loaded. Keeps the in-memory record on read failure or when
 * memory is already as new (never clobbers unsaved in-memory edits with
 * stale storage). Returns the record now in memory. */
export async function refreshDocumentFromStorage(
  id: string,
): Promise<StoredDocumentRecord | undefined> {
  try {
    const stored = await getRecord(id);
    const current = records.get(id);
    if (stored && isCommittedRecord(stored) && (!current || stored.lastModified > current.lastModified)) {
      await adoptStoredRecord(id, stored);
    }
  } catch {
    // Fall through to the in-memory record.
  }
  return records.get(id);
}

/** A stored committed record — not a draft, media, package, rows or source
 * record. */
function isCommittedRecord(record: PersistedRecord): record is StoredDocumentRecord {
  return (
    !isDraftRecord(record) &&
    !isMediaRecord(record) &&
    !isOpaqueRecord(record) &&
    !isSheetRowsRecord(record) &&
    !isWorkbookSourceRecord(record) &&
    !isDraftSheetRecord(record)
  );
}

/** Makes a committed record read back from storage the in-memory one: its
 * media and a deck's package rejoined, their latches seeded only when found,
 * and a workbook adopted as its shell, not resident — loadSpreadsheet
 * rejoins its rows and sources (and upgrades it) on demand (SPEC-137). */
async function adoptStoredRecord(id: string, stored: StoredDocumentRecord): Promise<void> {
  const mediaRecord = await getRecord(mediaKey(id));
  const media = mediaRecord && isMediaRecord(mediaRecord) ? mediaRecord.media : undefined;
  rejoinMedia(stored.model, media);
  // A deck's package (SPEC-103): the same found-or-legacy rule.
  const opaqueRecord = isPresentationModel(stored.model) ? await getRecord(opaqueKey(id)) : undefined;
  const opaque = opaqueRecord && isOpaqueRecord(opaqueRecord) ? opaqueRecord.opaque : undefined;
  rejoinOpaque(stored.model, opaque);
  if (isSpreadsheetModel(stored.model)) {
    workbooks.delete(id);
    workbookState(id).legacy = stored.sheetStore === undefined;
  }
  // Same rule as initLibrary: only a found MediaRecord may seed the
  // skip-map. A legacy record (inline media, no MediaRecord) must leave
  // the map unseeded so the next persist writes the media out rather
  // than stripping the shell and skipping the media write.
  if (media !== undefined) {
    lastPersistedMedia.set(id, stored.model.media);
  } else {
    lastPersistedMedia.delete(id);
  }
  if (opaque !== undefined && isPresentationModel(stored.model)) {
    lastPersistedOpaque.set(id, stored.model.opaque);
  } else {
    lastPersistedOpaque.delete(id);
  }
  records.set(id, stored);
}

/** Swaps a text document's in-memory model — applying recovered draft
 * content, or restoring the committed model on "Discard". Deliberately does
 * NOT bump lastModified or persist: neither direction of recovery is a
 * commit. */
export function replaceDocumentModel(id: string, model: DocumentModel): DocumentModel {
  return replaceModel(id, model);
}

/** The deck counterpart of replaceDocumentModel (SPEC-103): recovery, its
 * "Discard", and a version restore swap the deck's model through here. */
export function replacePresentationModel(id: string, model: PresentationModel): PresentationModel {
  return replaceModel(id, model);
}

/** The workbook counterpart (SPEC-138): recovery, its "Discard", a version
 * restore — and every edit, which hands the view's live model to the store
 * here instead of writing it through. The workbook becomes resident, and
 * holds changes only a draft or the next commit will write. */
export function replaceSpreadsheetModel(id: string, model: SpreadsheetModel): SpreadsheetModel {
  return replaceModel(id, model);
}

/** Swaps any editable model through its own kind's path — so a session
 * makes one call for every kind (SPEC-138). A media file has no editable
 * model: it throws, as replaceModel's loud rule does. */
export function replaceStoredModel(id: string, model: StoredModel): StoredModel {
  if (isPresentationModel(model)) {
    return replacePresentationModel(id, model);
  }
  if (isSpreadsheetModel(model)) {
    return replaceSpreadsheetModel(id, model);
  }
  if (isMediaModel(model)) {
    throw new Error(`Document ${id} is a media file`);
  }
  return replaceDocumentModel(id, model);
}

/** The words a kind-crossing refusal names a stored model by. */
function kindPhrase(model: StoredModel): string {
  if (isPresentationModel(model)) {
    return "a presentation";
  }
  if (isMediaModel(model)) {
    return "a media file";
  }
  return isSpreadsheetModel(model) ? "a spreadsheet" : "a text document";
}

/** The one swap the editing products share. Never crosses kinds: replacing a
 * media file's model with editor state would destroy the stored bytes (the
 * loud-throw rule getOrCreateDocument keeps), and a deck, a workbook and a
 * text document must never overwrite each other. Never bumps lastModified or
 * persists: neither direction of recovery is a commit. */
function replaceModel<M extends DocumentModel | PresentationModel | SpreadsheetModel>(id: string, model: M): M {
  model.id = id;
  const record = records.get(id);
  if (record && (isMediaModel(record.model) || kindPhrase(record.model) !== kindPhrase(model))) {
    throw new Error(`Document ${id} is ${kindPhrase(record.model)}`);
  }
  if (!record) {
    const created: StoredDocumentRecord = {
      id,
      format: model.format ?? "new",
      lastModified: Date.now(),
      model,
    };
    records.set(id, created);
  } else {
    record.model = model;
  }
  if (isSpreadsheetModel(model)) {
    // SPEC-137/138: the view only ever holds a whole workbook.
    const state = workbookState(id);
    state.resident = true;
    state.edits += 1;
  }
  return model;
}

/** Stores an already-built document (e.g. one parsed from an opened file)
 * under its own id, overwriting anything previously stored there. Wraps it in
 * a library record (format from the model, defaulting to "new") and writes
 * through to storage like every other mutation. */
export function storeDocument(doc: DocumentModel): void {
  storeRecord(doc, doc.format ?? "new");
}

/** The presentation counterpart of `storeDocument` — the open path stores a
 * freshly parsed `.pptx`/`.odp` through the same record machinery. */
export function storePresentation(model: PresentationModel): void {
  storeRecord(model, model.format);
}

/** The media counterpart (SPEC-107) — an opened PDF/image stores through the
 * same record machinery, its payload split into the media record like every
 * other kind's. */
export function storeMedia(model: MediaModel): void {
  storeRecord(model, model.format);
}

/** Writes an edited media document through (SPEC-112's Save and Flatten):
 * `model` — a new model object, the image editor never mutates the stored
 * one — becomes the record's model (its format too: a flattened SVG is a
 * PNG now), `lastModified` bumps (so the library thumbnail, keyed by id +
 * lastModified, rebuilds), and the media record then the shell are written.
 * One write-through, and a STRICT one: unlike touchDocument's
 * fire-and-forget persist, a failed write rejects and puts the previous
 * model back, so "Saved" or "Flattened" is never reported over storage that
 * still holds the old state. The media record goes first: if the shell
 * write then fails, storage holds the new pixels under the old shell (an
 * edited shell whose render is gone fails closed), never a flattened shell
 * over the original's bytes. A missing or non-media id is a no-op. */
export async function storeMediaEdit(model: MediaModel): Promise<void> {
  const record = records.get(model.id);
  if (!record || !isMediaModel(record.model)) {
    return;
  }
  const previous = { model: record.model, format: record.format, lastModified: record.lastModified };
  record.model = model;
  record.format = model.format;
  record.lastModified = Date.now();
  try {
    await persistMedia(model.id, model.media);
    await putRecordYielding({ ...record, model: shellModel(model) });
  } catch (error) {
    markStorageUnavailable();
    if (records.get(model.id) === record && record.model === model) {
      record.model = previous.model;
      record.format = previous.format;
      record.lastModified = previous.lastModified;
    }
    throw error;
  }
}

/** The workbook counterpart (SPEC-116) — an opened `.xlsx`/`.ods` stores
 * through the same record machinery, its sheets' rows split into their own
 * records like every other kind's payload. */
export function storeSpreadsheet(model: SpreadsheetModel): void {
  model.preview = workbookPreview(model);
  const previous = records.get(model.id);
  // SPEC-137: a freshly opened workbook is resident; re-stored over an id,
  // it keeps what storage holds so the first persist writes only the diff.
  const state = workbookState(model.id);
  state.resident = true;
  state.incomplete = false;
  state.requests += 1;
  storeRecord(model, model.format);
  if (previous?.sheetStore) {
    records.get(model.id)!.sheetStore = previous.sheetStore;
  }
}

/** Writes a workbook's view state (SPEC-117 "View state": each sheet's
 * selection, zoom, gridlines and top-left cell, and the active sheet) back
 * to its stored record — the shell record only, keeping the existing
 * `lastModified`: viewing is not editing, so it never touches the document
 * (no touchDocument), never marks a sheet dirty and, because a view-state
 * change replaces sheet objects but keeps each sheet's `rows` array, never
 * writes a chunk record (the per-chunk latch skips an unchanged reference).
 * `model` replaces the record's model; a missing or non-workbook id is a
 * no-op. Resolves once the write has settled (never rejects). */
export function storeSpreadsheetViewState(model: SpreadsheetModel): Promise<void> {
  const record = records.get(model.id);
  if (!record || !isSpreadsheetModel(record.model) || workbooks.get(model.id)?.incomplete) {
    return Promise.resolve();
  }
  record.model = model;
  // The view only ever holds a whole workbook (loadSpreadsheet's).
  const state = workbookState(model.id);
  state.resident = true;
  if (uncommitted(state)) {
    // SPEC-138: the model holds edits only the draft has — writing the view
    // state would commit them under a "Discard" that must still undo them.
    // The view fields ride in the next draft and commit instead.
    return Promise.resolve();
  }
  return persist(record);
}

/** Writes an edited workbook through (SPEC-121 "The edit pipeline"): the
 * live `model` becomes the record's model (a view-state write may have left
 * a copy there) and touchDocument bumps recency and persists the shell plus
 * exactly the rows records whose reference changed. A missing or
 * non-workbook id is a no-op. */
export function storeSpreadsheetEdit(model: SpreadsheetModel): Promise<void> {
  const record = records.get(model.id);
  if (!record || !isSpreadsheetModel(record.model) || workbooks.get(model.id)?.incomplete) {
    // SPEC-137: a workbook whose storage came back incomplete is read-only.
    return Promise.resolve();
  }
  record.model = model;
  workbookState(model.id).resident = true;
  return touchDocument(model.id);
}

/** Stores whatever the kind-agnostic open path returned, dispatching on
 * `kind` — the shared tail of the picker, drag-drop and URL open paths. A
 * Keynote import's non-blocking notice (SPEC-087 AC-37) is handed to
 * `notify` (the caller's toast) after the model is stored, so it survives
 * the navigation into the editor. */
export function storeOpenedFile(opened: OpenedFile, notify?: (message: string) => void): void {
  if (opened.kind === "presentation") {
    storePresentation(opened.model);
    if (opened.importNotice !== undefined) {
      notify?.(opened.importNotice);
    }
  } else if (opened.kind === "media") {
    storeMedia(opened.model);
  } else if (opened.kind === "spreadsheet") {
    storeSpreadsheet(opened.model);
  } else {
    storeDocument(opened.model);
  }
}

function storeRecord(model: StoredModel, format: DocumentFormat): void {
  const record: StoredDocumentRecord = {
    id: model.id,
    format,
    lastModified: Date.now(),
    model,
  };
  records.set(model.id, record);
  // Deferred to browser idle (SPEC-047 AC-3/AC-16): storeDocument runs on
  // the open path right before navigating into the document, and the
  // IndexedDB put's synchronous structured clone of a large model was one
  // of the main-thread blocks between choosing the file and first paint.
  // The in-memory record above is what the view reads — only the
  // write-through waits. The guard re-checks the record is still current at
  // fire time so a document deleted (or re-stored) in the deferral window
  // is not resurrected by a stale write.
  runWhenIdle(() => {
    if (records.get(record.id) === record) {
      persist(record);
    }
  }, 1500);
}
