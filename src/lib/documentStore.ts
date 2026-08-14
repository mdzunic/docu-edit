// The document library (SPEC-003): an in-memory map of document records that
// writes through to IndexedDB. The in-memory map is the source of truth for
// the running session, so every read stays synchronous; persistence is a
// best-effort mirror. When IndexedDB is unavailable or a write fails (e.g.
// quota exceeded), the app keeps working in-memory and the storage status
// flips so the UI can show its warning banner.
import { createDocument, type DocumentModel, type MediaPart } from "./documentModel";
import { deleteRecord, getRecord, loadAllRecords, putRecord } from "./idb";
import { runWhenIdle, yieldToMainThread } from "./idleScheduler";
import type { CaretPosition } from "./textEditing";

/** Where a document's content last came from or went to. "new" means it was
 * created in the app and has never been opened from or saved as a file. */
export type DocumentFormat = "docx" | "odt" | "md" | "new";

/** What the library persists per document: the model plus its metadata.
 * This is the COMMITTED record (SPEC-010): the last clean state, which the
 * recovery bar's "Discard" restores. */
export interface StoredDocumentRecord {
  id: string;
  format: DocumentFormat;
  /** Epoch milliseconds of the last modification. */
  lastModified: number;
  model: DocumentModel;
  /** Caret position at the last commit, so reopening restores it without
   * reconstruction (SPEC-010). Absent on records from before autosave. */
  caret?: CaretPosition | null;
  /** Canvas scrollTop in pixels at the last commit. */
  scroll?: number;
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

/** Everything the documents object store can hold. */
export type PersistedRecord = StoredDocumentRecord | DraftRecord | MediaRecord;

export const DRAFT_KEY_PREFIX = "draft:";
export const MEDIA_KEY_PREFIX = "media:";
/** A DRAFT's media copy (see persistDraftMedia): its own key, because the
 * committed record rejoins from MEDIA_KEY_PREFIX and an autosave tick must
 * never be able to overwrite the media the committed state depends on —
 * "Discard" after a crash restores the committed shell, and if drafts
 * shared the committed media record, a draft that deleted an image would
 * already have destroyed the committed copy of it. */
export const DRAFT_MEDIA_KEY_PREFIX = "draftmedia:";

function draftKey(docId: string): string {
  return DRAFT_KEY_PREFIX + docId;
}

function mediaKey(docId: string): string {
  return MEDIA_KEY_PREFIX + docId;
}

function draftMediaKey(docId: string): string {
  return DRAFT_MEDIA_KEY_PREFIX + docId;
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

/** Row-level metadata the library list renders; no document content. */
export interface LibraryEntry {
  id: string;
  title: string;
  format: DocumentFormat;
  lastModified: number;
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

function markStorageUnavailable(): void {
  if (!storageAvailable) {
    return;
  }
  storageAvailable = false;
  for (const listener of storageListeners) {
    listener();
  }
}

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
  const record: MediaRecord = { id: mediaKey(docId), docId, media: media ?? [] };
  return putRecordYielding(record).catch((error: unknown) => {
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
  const record: MediaRecord = { id: draftMediaKey(docId), docId, media: media ?? [] };
  return putRecordYielding(record).catch((error: unknown) => {
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
function stripMedia(model: DocumentModel): DocumentModel {
  return { ...model, media: [] };
}

/** Rejoins a doc's media (read back via `mediaKey`) into a model whose
 * stored shell had it stripped — the inverse of `stripMedia`, applied on
 * every read path (initLibrary, refreshDocumentFromStorage, loadDraft).
 * Absent when nothing was ever persisted for it (a document with no
 * images never gets a MediaRecord — see persistMedia's skip). */
function rejoinMedia(model: DocumentModel, media: MediaPart[] | undefined): DocumentModel {
  if (media === undefined || media.length === 0) {
    return model;
  }
  model.media = media;
  return model;
}

// Fire-and-forget write-through: a failed write must not break the in-memory
// session — it only flips the storage status.
function persist(record: StoredDocumentRecord): void {
  persistMedia(record.id, record.model.media)?.catch(() => markStorageUnavailable());
  putRecordYielding({ ...record, model: stripMedia(record.model) }).catch(() => markStorageUnavailable());
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
      for (const record of stored) {
        // Strictly committed media only — a draft's media copy
        // (isDraftMediaRecord) must never rejoin the committed shell.
        if (isDraftMediaRecord(record)) {
          continue;
        }
        if (isMediaRecord(record)) {
          mediaByDocId.set(record.docId, record.media);
        }
      }
      for (const record of stored) {
        // Drafts and media records live in the same object store but are
        // not library entries; the editor loads a draft explicitly via
        // loadDraft, and media rejoins its document record below.
        if (isDraftRecord(record) || isMediaRecord(record) || isDraftMediaRecord(record)) {
          continue;
        }
        if (!records.has(record.id)) {
          const media = mediaByDocId.get(record.id);
          rejoinMedia(record.model, media);
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
    })
    .catch(() => markStorageUnavailable());
  return initPromise;
}

// --- reads -----------------------------------------------------------------

/** Returns the stored document for an id, or undefined if none exists. */
export function getDocument(id: string): DocumentModel | undefined {
  return records.get(id)?.model;
}

/** The full in-memory record (metadata + model) for an id. */
export function getStoredRecord(id: string): StoredDocumentRecord | undefined {
  return records.get(id);
}

/** Returns the stored document for an id, creating and storing a fresh one
 * if none exists yet. */
export function getOrCreateDocument(id: string): DocumentModel {
  const existing = records.get(id);
  if (existing) {
    return existing.model;
  }
  const record: StoredDocumentRecord = {
    id,
    format: "new",
    lastModified: Date.now(),
    model: createDocument(id),
  };
  records.set(id, record);
  persist(record);
  return record.model;
}

/** Library rows sorted by last modified, newest first. */
export function listDocuments(): LibraryEntry[] {
  return [...records.values()]
    .map(({ id, format, lastModified, model }) => ({
      id,
      format,
      lastModified,
      title: model.title,
    }))
    .sort((a, b) => b.lastModified - a.lastModified);
}

// --- mutations -------------------------------------------------------------

/** Records that a document was just modified (content already updated on its
 * model, e.g. via the app bar's title editor) and persists it. */
export function touchDocument(id: string): void {
  const record = records.get(id);
  if (!record) {
    return;
  }
  record.lastModified = Date.now();
  persist(record);
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

/** Removes a document from the session and from storage, along with any
 * autosave draft. */
export function deleteDocument(id: string): void {
  if (!records.delete(id)) {
    return;
  }
  lastPersistedMedia.delete(id);
  lastDraftPersistedMedia.delete(id);
  deleteRecord(id).catch(() => markStorageUnavailable());
  // Draft/media cleanup is best-effort and not a SPEC-003 write-through
  // failure.
  deleteRecord(draftKey(id)).catch(() => {});
  deleteRecord(mediaKey(id)).catch(() => {});
  deleteRecord(draftMediaKey(id)).catch(() => {});
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
  await Promise.all([
    mediaWrite ?? Promise.resolve(),
    putRecordYielding({ ...record, model: stripMedia(record.model) }),
  ]);
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
      (!current || stored.lastModified > current.lastModified)
    ) {
      const mediaRecord = await getRecord(mediaKey(id));
      const media = mediaRecord && isMediaRecord(mediaRecord) ? mediaRecord.media : undefined;
      rejoinMedia(stored.model, media);
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
      return stored;
    }
  } catch {
    // Fall through to the in-memory record.
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
  const record: StoredDocumentRecord = {
    id: doc.id,
    format: doc.format ?? "new",
    lastModified: Date.now(),
    model: doc,
  };
  records.set(doc.id, record);
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
