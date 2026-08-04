// The document library (SPEC-003): an in-memory map of document records that
// writes through to IndexedDB. The in-memory map is the source of truth for
// the running session, so every read stays synchronous; persistence is a
// best-effort mirror. When IndexedDB is unavailable or a write fails (e.g.
// quota exceeded), the app keeps working in-memory and the storage status
// flips so the UI can show its warning banner.
import { createDocument, type DocumentModel } from "./documentModel";
import { deleteRecord, getRecord, loadAllRecords, putRecord } from "./idb";
import type { CaretPosition } from "./textEditing";

/** Where a document's content last came from or went to. "new" means it was
 * created in the app and has never been opened from or saved as a file. */
export type DocumentFormat = "docx" | "odt" | "new";

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

/** Everything the documents object store can hold. */
export type PersistedRecord = StoredDocumentRecord | DraftRecord;

export const DRAFT_KEY_PREFIX = "draft:";

function draftKey(docId: string): string {
  return DRAFT_KEY_PREFIX + docId;
}

export function isDraftRecord(record: PersistedRecord): record is DraftRecord {
  return record.id.startsWith(DRAFT_KEY_PREFIX);
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

// Fire-and-forget write-through: a failed write must not break the in-memory
// session — it only flips the storage status.
function persist(record: StoredDocumentRecord): void {
  putRecord(record).catch(() => markStorageUnavailable());
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
      for (const record of stored) {
        // Drafts live in the same object store but are not library entries;
        // the editor loads them explicitly via loadDraft.
        if (isDraftRecord(record)) {
          continue;
        }
        if (!records.has(record.id)) {
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
  deleteRecord(id).catch(() => markStorageUnavailable());
  // Draft cleanup is best-effort and not a SPEC-003 write-through failure.
  deleteRecord(draftKey(id)).catch(() => {});
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
  const draft: DraftRecord = {
    id: draftKey(id),
    docId: id,
    savedAt: Date.now(),
    model: record.model,
    caret,
    scroll,
  };
  await putRecord(draft);
}

/** Reads a document's autosave draft; undefined when there is none or the
 * store is unreadable (either way there is nothing to recover). */
export async function loadDraft(id: string): Promise<DraftRecord | undefined> {
  try {
    const stored = await getRecord(draftKey(id));
    return stored && isDraftRecord(stored) ? stored : undefined;
  } catch {
    return undefined;
  }
}

/** Removes a document's autosave draft. Best-effort: a draft that survives
 * is harmless once the committed record is newer, because recovery only
 * triggers on a draft NEWER than the committed record. */
export async function deleteDraft(id: string): Promise<void> {
  try {
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
  await putRecord(record);
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
      (!current || stored.lastModified > current.lastModified)
    ) {
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
  persist(record);
}
