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
import type { PackageVariant, Row, SpreadsheetModel } from "./spreadsheet/model";
import { SPREADSHEET_READER_VERSION, upgradeSpreadsheetModel } from "./spreadsheet/open";
import { buildPreview } from "./spreadsheet/preview";
import type { OpenedFile } from "./openDocument";
import { loadPreferences, MARGIN_PRESETS, setPreferencesStorageFailureHandler } from "./preferences";
import { PAPER_SIZES } from "./core/pageSetup";
import { encodePendingMediaInBackground, type MediaPart } from "./core/media";
import { normaliseShapeFill } from "./core/shapes";
import { deleteRecord, getRecord, loadAllRecords, putRecord } from "./idb";
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
  /** Starred by the user (SPEC-067). Absent means not starred. */
  starred?: boolean;
  /** When the document was moved to the trash (SPEC-067), as an epoch
   * millisecond value. Absent means it is in the library. A record with this
   * set is excluded from listDocuments() and from every enumeration; it keeps
   * its id and content, so restoreDocument returns it intact. */
  trashedAt?: number;
}

/** The autosave journal (SPEC-010): written on the 2 s / 30 s cadence,
 * stored beside the committed record under a prefixed key. A draft newer
 * than its committed record on open means the previous session did not end
 * with a clean close — the crash signal the recovery bar keys on. */
export interface DraftRecord {
  /** The IndexedDB key: DRAFT_KEY_PREFIX + docId. */
  id: string;
  docId: string;
  /** Epoch milliseconds when the draft was written. */
  savedAt: number;
  model: DocumentModel;
  caret: CaretPosition | null;
  scroll: number;
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

/** One workbook sheet's rows, persisted as their own record (SPEC-116): the
 * shell record stores every sheet with `rows: []`, and each sheet's rows go
 * here — written only when that array's reference changed since the last
 * write (see persistSheetRows). The model's copy-on-write invariant (an edit
 * replaces the rows array it touches and nothing else) is what makes the
 * reference check sound, so a rename, a star or a library listing never
 * structured-clones a cell. */
export interface SheetRowsRecord {
  /** The IndexedDB key: SHEET_ROWS_KEY_PREFIX + docId + ":" + sheetId. */
  id: string;
  docId: string;
  sheetId: string;
  rows: Row[];
}

/** Everything the documents object store can hold. */
export type PersistedRecord = StoredDocumentRecord | DraftRecord | MediaRecord | SheetRowsRecord;

export const DRAFT_KEY_PREFIX = "draft:";
export const MEDIA_KEY_PREFIX = "media:";
/** A DRAFT's media copy (see persistDraftMedia): its own key, because the
 * committed record rejoins from MEDIA_KEY_PREFIX and an autosave tick must
 * never be able to overwrite the media the committed state depends on —
 * "Discard" after a crash restores the committed shell, and if drafts
 * shared the committed media record, a draft that deleted an image would
 * already have destroyed the committed copy of it. */
export const DRAFT_MEDIA_KEY_PREFIX = "draftmedia:";
export const SHEET_ROWS_KEY_PREFIX = "sheetrows:";

function draftKey(docId: string): string {
  return DRAFT_KEY_PREFIX + docId;
}

function mediaKey(docId: string): string {
  return MEDIA_KEY_PREFIX + docId;
}

function draftMediaKey(docId: string): string {
  return DRAFT_MEDIA_KEY_PREFIX + docId;
}

function sheetRowsKey(docId: string, sheetId: string): string {
  return `${SHEET_ROWS_KEY_PREFIX}${docId}:${sheetId}`;
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

export function isSheetRowsRecord(record: PersistedRecord): record is SheetRowsRecord {
  return record.id.startsWith(SHEET_ROWS_KEY_PREFIX);
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

// The rows array last written for each sheet (keyed by its record id), so
// persistSheetRows skips every sheet whose rows reference is unchanged — the
// same latch discipline as lastPersistedMedia, seeded from storage on load
// only for records that actually exist there.
const lastPersistedRows = new Map<string, Row[]>();
// Per document, the sheet-rows record ids stored (or being written) — so the
// next persist deletes the record of a sheet that no longer exists, and
// deleteDocument finds every one.
const sheetRowsKeys = new Map<string, Set<string>>();

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

/** A model with `media` stripped to `[]` for the shell record `persist`/
 * `saveDraft`/`commitDocument` hand to `putRecord` — the array itself
 * always goes through `persistMedia` instead. */
function stripMedia<M extends StoredModel>(model: M): M {
  return { ...model, media: [] };
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

/** The shell record's model: media stripped (stripMedia), and for a workbook
 * every sheet's rows emptied — they live in their own SheetRowsRecords. */
function shellModel<M extends StoredModel>(model: M): M {
  const shell = stripMedia(model);
  if (isSpreadsheetModel(shell)) {
    shell.sheets = shell.sheets.map((sheet) => ({ ...sheet, rows: [] }));
  }
  return shell;
}

/** Writes a workbook's sheet rows as their own records — only the sheets
 * whose `rows` reference differs from the last one written (SPEC-116's
 * copy-on-write invariant makes that sound) — and deletes the record of any
 * sheet the model no longer has. When any rows changed, the card preview is
 * recomputed first so the shell written after this carries it. Returns the
 * writes' promises; a failed write clears its latch (so the next persist
 * retries) and is handed to `onFailure`. */
function persistSheetRows(docId: string, model: StoredModel, onFailure: (error: unknown) => void): Promise<void>[] {
  if (!isSpreadsheetModel(model)) {
    return [];
  }
  const writes: Promise<void>[] = [];
  const current = new Set<string>();
  for (const sheet of model.sheets) {
    const key = sheetRowsKey(docId, sheet.id);
    current.add(key);
    const rows = sheet.rows;
    if (lastPersistedRows.get(key) === rows) {
      continue;
    }
    lastPersistedRows.set(key, rows);
    const record: SheetRowsRecord = { id: key, docId, sheetId: sheet.id, rows };
    writes.push(
      putRecordYielding(record).catch((error: unknown) => {
        if (lastPersistedRows.get(key) === rows) {
          lastPersistedRows.delete(key);
        }
        onFailure(error);
      }),
    );
  }
  for (const key of sheetRowsKeys.get(docId) ?? []) {
    if (!current.has(key)) {
      lastPersistedRows.delete(key);
      deleteRecord(key).catch(() => {});
    }
  }
  sheetRowsKeys.set(docId, current);
  if (writes.length > 0) {
    model.preview = buildPreview(model);
  }
  return writes;
}

/** Rejoins a workbook's sheet rows (read back from their own records) into a
 * model whose stored shell had them emptied — applied on every read path.
 * Seeds the latch only for rows that really exist in storage, and records
 * every stored key for the document (orphans included, so the next persist
 * removes them). */
function rejoinSheetRows(docId: string, model: SpreadsheetModel, stored: readonly SheetRowsRecord[]): void {
  const bySheet = new Map(stored.map((record) => [record.sheetId, record.rows]));
  for (const sheet of model.sheets) {
    const rows = bySheet.get(sheet.id);
    if (rows !== undefined) {
      sheet.rows = rows;
      lastPersistedRows.set(sheetRowsKey(docId, sheet.id), rows);
    }
  }
  sheetRowsKeys.set(docId, new Set(stored.map((record) => record.id)));
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
  const rows = persistSheetRows(record.id, record.model, () => markStorageUnavailable());
  const shell = putRecordYielding({ ...record, model: shellModel(record.model) }).catch(() =>
    markStorageUnavailable(),
  );
  return Promise.all([media, ...rows, shell]).then(() => undefined);
}

/** Brings a stored workbook up to the current reader version in place
 * (SPEC-116 "Reader versions"): only the model is replaced, so the record
 * keeps its id, starred and trash state, and upgradeSpreadsheetModel keeps
 * the title and every sheet's id; the upgraded record is written back once.
 * A record that cannot be re-read stays exactly as it was. */
async function upgradeSpreadsheetRecord(record: StoredDocumentRecord): Promise<void> {
  const model = record.model;
  if (!isSpreadsheetModel(model) || model.readerVersion >= SPREADSHEET_READER_VERSION) {
    return;
  }
  let upgraded: SpreadsheetModel;
  try {
    upgraded = await upgradeSpreadsheetModel(model);
  } catch {
    return;
  }
  if (records.get(record.id) === record && record.model === model) {
    record.model = upgraded;
    void persist(record);
  }
}

/** The view's half of the upgrade path (useWorkbookLoad): upgrades the
 * workbook stored under `id` if it is behind the reader version, resolving
 * to the current model — undefined for a missing or non-workbook id. */
export async function upgradeStoredSpreadsheet(id: string): Promise<SpreadsheetModel | undefined> {
  const record = records.get(id);
  if (record) {
    await upgradeSpreadsheetRecord(record);
  }
  return getSpreadsheet(id);
}

// --- hydration -------------------------------------------------------------

let initPromise: Promise<void> | null = null;

/** Loads persisted documents into the in-memory map. Idempotent, and never
 * rejects: a failure marks storage unavailable and leaves the store empty,
 * so callers can always `await` it and then read synchronously. main.ts
 * awaits this before mounting the app. */
export function initLibrary(): Promise<void> {
  initPromise ??= loadAllRecords()
    .then(async (stored) => {
      const mediaByDocId = new Map<string, MediaPart[]>();
      const sheetRowsByDocId = new Map<string, SheetRowsRecord[]>();
      for (const record of stored) {
        // Strictly committed media only — a draft's media copy
        // (isDraftMediaRecord) must never rejoin the committed shell.
        if (isDraftMediaRecord(record)) {
          continue;
        }
        if (isMediaRecord(record)) {
          mediaByDocId.set(record.docId, record.media);
        } else if (isSheetRowsRecord(record)) {
          const forDoc = sheetRowsByDocId.get(record.docId) ?? [];
          forDoc.push(record);
          sheetRowsByDocId.set(record.docId, forDoc);
        }
      }
      const behindReaderVersion: StoredDocumentRecord[] = [];
      for (const record of stored) {
        // Drafts, media and sheet-rows records live in the same object
        // store but are not library entries; the editor loads a draft
        // explicitly via loadDraft, and media and rows rejoin their
        // document record below.
        if (isDraftRecord(record) || isMediaRecord(record) || isDraftMediaRecord(record) || isSheetRowsRecord(record)) {
          continue;
        }
        if (!records.has(record.id)) {
          const media = mediaByDocId.get(record.id);
          rejoinMedia(record.model, media);
          if (isSpreadsheetModel(record.model)) {
            rejoinSheetRows(record.id, record.model, sheetRowsByDocId.get(record.id) ?? []);
            if (record.model.readerVersion < SPREADSHEET_READER_VERSION) {
              behindReaderVersion.push(record);
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
          records.set(record.id, record);
        }
      }
      // SPEC-116: a workbook stored by an older reader is re-read from its
      // retained package before anything reads it.
      for (const record of behindReaderVersion) {
        await upgradeSpreadsheetRecord(record);
      }
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
 * (SPEC-116): undefined for absent ids and for every other kind. */
export function getSpreadsheet(id: string): SpreadsheetModel | undefined {
  const model = records.get(id)?.model;
  return model !== undefined && isSpreadsheetModel(model) ? model : undefined;
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
  persist(copy);
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
  // A workbook (SPEC-116) counts its shell plus every sheet's rows record —
  // each rows array measured once per reference (copy-on-write: an
  // unchanged sheet is never re-serialised to be sized).
  if (isSpreadsheetModel(record.model)) {
    return record.model.sheets.reduce((total, sheet) => total + rowsBytes(sheet.rows), jsonBytes(shellModel(record.model)));
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

const rowsByteCache = new WeakMap<Row[], number>();

function rowsBytes(rows: Row[]): number {
  let bytes = rowsByteCache.get(rows);
  if (bytes === undefined) {
    bytes = jsonBytes(rows);
    rowsByteCache.set(rows, bytes);
  }
  return bytes;
}

/** Removes a document from the session and from storage, along with any
 * autosave draft. */
export function deleteDocument(id: string): void {
  const model = records.get(id)?.model;
  if (!records.delete(id)) {
    return;
  }
  lastPersistedMedia.delete(id);
  lastDraftPersistedMedia.delete(id);
  deleteRecord(id).catch(() => markStorageUnavailable());
  // Draft/media/rows cleanup is best-effort and not a SPEC-003 write-through
  // failure.
  deleteRecord(draftKey(id)).catch(() => {});
  deleteRecord(mediaKey(id)).catch(() => {});
  deleteRecord(draftMediaKey(id)).catch(() => {});
  const rowsKeys = new Set(sheetRowsKeys.get(id));
  if (model && isSpreadsheetModel(model)) {
    for (const sheet of model.sheets) {
      rowsKeys.add(sheetRowsKey(id, sheet.id));
    }
  }
  for (const key of rowsKeys) {
    lastPersistedRows.delete(key);
    deleteRecord(key).catch(() => {});
  }
  sheetRowsKeys.delete(id);
}

// --- autosave (SPEC-010) ---------------------------------------------------
//
// Unlike the SPEC-003 write-through above, these reject on failure instead of
// latching the storage banner: the autosave UI owns its own failure state
// ("Not saved — browser storage is full" in the app bar) and recovers the
// moment a write succeeds again.

/** Writes the autosave draft for a document from its current in-memory
 * model. Rejects on any storage failure. */
export async function saveDraft(
  id: string,
  caret: CaretPosition | null,
  scroll: number,
): Promise<void> {
  const record = records.get(id);
  if (!record) {
    throw new Error(`No document with id ${id}`);
  }
  if (isPresentationModel(record.model)) {
    // Drafts are a text-editor feature; nothing edits presentations yet.
    throw new Error(`Document ${id} is a presentation`);
  }
  if (isMediaModel(record.model)) {
    // A media document has no draft state at all (SPEC-107): its bytes are
    // immutable, and an autosave draft would be a path to overwriting them.
    throw new Error(`Document ${id} is a media file`);
  }
  if (isSpreadsheetModel(record.model)) {
    // Nothing edits a workbook yet (SPEC-116; SPEC-121 brings editing), and
    // a text-editor draft over its id would destroy it.
    throw new Error(`Document ${id} is a spreadsheet`);
  }
  const mediaWrite = persistDraftMedia(id, record.model.media);
  const draft: DraftRecord = {
    id: draftKey(id),
    docId: id,
    savedAt: Date.now(),
    model: stripMedia(record.model),
    caret,
    scroll,
  };
  await Promise.all([mediaWrite ?? Promise.resolve(), putRecordYielding(draft)]);
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
    return stored;
  } catch {
    return undefined;
  }
}

/** Removes a document's autosave draft. Best-effort: a draft that survives
 * is harmless once the committed record is newer, because recovery only
 * triggers on a draft NEWER than the committed record. */
export async function deleteDraft(id: string): Promise<void> {
  try {
    lastDraftPersistedMedia.delete(id);
    deleteRecord(draftMediaKey(id)).catch(() => {});
    await deleteRecord(draftKey(id));
  } catch {
    // Ignored — see above.
  }
}

/** Promotes the current in-memory state to the committed record — the last
 * clean state "Discard" restores — recording caret and scroll alongside the
 * model, and clears the draft. Rejects when the committed write fails. */
export async function commitDocument(
  id: string,
  caret: CaretPosition | null,
  scroll: number,
): Promise<void> {
  const record = records.get(id);
  if (!record) {
    return;
  }
  record.lastModified = Date.now();
  record.caret = caret;
  record.scroll = scroll;
  const mediaWrite = persistMedia(id, record.model.media);
  const failures: unknown[] = [];
  const rowsWrites = persistSheetRows(id, record.model, (error) => failures.push(error));
  await Promise.all([
    mediaWrite ?? Promise.resolve(),
    ...rowsWrites,
    putRecordYielding({ ...record, model: shellModel(record.model) }),
  ]);
  if (failures.length > 0) {
    throw failures[0];
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
    if (
      stored &&
      !isDraftRecord(stored) &&
      !isMediaRecord(stored) &&
      !isSheetRowsRecord(stored) &&
      (!current || stored.lastModified > current.lastModified)
    ) {
      const mediaRecord = await getRecord(mediaKey(id));
      const media = mediaRecord && isMediaRecord(mediaRecord) ? mediaRecord.media : undefined;
      rejoinMedia(stored.model, media);
      if (isSpreadsheetModel(stored.model)) {
        const rowsRecords: SheetRowsRecord[] = [];
        for (const sheet of stored.model.sheets) {
          const rowsRecord = await getRecord(sheetRowsKey(id, sheet.id));
          if (rowsRecord && isSheetRowsRecord(rowsRecord)) {
            rowsRecords.push(rowsRecord);
          }
        }
        rejoinSheetRows(id, stored.model, rowsRecords);
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
      records.set(id, stored);
    }
  } catch {
    // Fall through to the in-memory record.
  }
  // SPEC-116: whichever record is now current, a workbook behind the reader
  // version is upgraded before it is handed out.
  const record = records.get(id);
  if (record) {
    await upgradeSpreadsheetRecord(record);
  }
  return records.get(id);
}

/** Swaps a document's in-memory model — applying recovered draft content, or
 * restoring the committed model on "Discard". Deliberately does NOT bump
 * lastModified or persist: neither direction of recovery is a commit. */
export function replaceDocumentModel(
  id: string,
  model: DocumentModel,
): DocumentModel {
  model.id = id;
  const record = records.get(id);
  if (record && isMediaModel(record.model)) {
    // SPEC-107: replacing a media document's model with text-editor state
    // would destroy the stored bytes — the loud-throw rule applies.
    throw new Error(`Document ${id} is a media file`);
  }
  if (record && isSpreadsheetModel(record.model)) {
    // SPEC-116: the same rule for a workbook.
    throw new Error(`Document ${id} is a spreadsheet`);
  }
  if (!record) {
    const created: StoredDocumentRecord = {
      id,
      format: model.format ?? "new",
      lastModified: Date.now(),
      model,
    };
    records.set(id, created);
    return model;
  }
  record.model = model;
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

/** The workbook counterpart (SPEC-116) — an opened `.xlsx`/`.ods` stores
 * through the same record machinery, its sheets' rows split into their own
 * records like every other kind's payload. */
export function storeSpreadsheet(model: SpreadsheetModel): void {
  storeRecord(model, model.format);
}

/** Writes a workbook's view state (SPEC-117 "View state": each sheet's
 * selection, zoom, gridlines and top-left cell, and the active sheet) back
 * to its stored record — the shell record only, keeping the existing
 * `lastModified`: viewing is not editing, so it never touches the document
 * (no touchDocument), never marks a sheet dirty and, because a view-state
 * change replaces sheet objects but keeps each sheet's `rows` array, never
 * writes a rows record (persistSheetRows skips an unchanged reference).
 * `model` replaces the record's model; a missing or non-workbook id is a
 * no-op. Resolves once the write has settled (never rejects). */
export function storeSpreadsheetViewState(model: SpreadsheetModel): Promise<void> {
  const record = records.get(model.id);
  if (!record || !isSpreadsheetModel(record.model)) {
    return Promise.resolve();
  }
  record.model = model;
  return persist(record);
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
