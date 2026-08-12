// Thin promise wrappers around the IndexedDB persistence for the document
// library. Kept as its own module so documentStore can be unit-tested with
// this seam mocked, and so nothing else touches IndexedDB directly.
//
// Every function rejects when IndexedDB is missing, blocked, or over quota —
// callers (documentStore) treat any rejection as "storage unavailable" and
// keep working in-memory.
import type { PersistedRecord } from "./documentStore";
import type { VersionStoreRecord } from "./versionHistory";
import type { StoredTemplateRecord } from "./templateStore";

const DB_NAME = "docu-edit";
const DB_VERSION = 3;
const STORE_NAME = "documents";
// SPEC-040: version history's own store, kept separate from `documents` so
// rendering the library never reads 50 MB of version payloads per document
// into memory at startup (see versionHistory.ts's header comment).
const VERSIONS_STORE_NAME = "versions";
// SPEC-046: user templates get their own store too, for the same reason —
// the gallery reads every template's metadata without touching the
// `documents` store's much larger population.
const TEMPLATES_STORE_NAME = "templates";

let dbPromise: Promise<IDBDatabase> | null = null;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    // If `indexedDB` is absent (jsdom, locked-down browsers) this throws
    // synchronously inside the executor, which rejects the promise.
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(VERSIONS_STORE_NAME)) {
        db.createObjectStore(VERSIONS_STORE_NAME, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(TEMPLATES_STORE_NAME)) {
        db.createObjectStore(TEMPLATES_STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.addEventListener("error", () =>
      reject(request.error ?? new Error("IndexedDB open failed")),
    );
  });
}

function getDb(): Promise<IDBDatabase> {
  dbPromise ??= openDb();
  return dbPromise;
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.addEventListener("error", () =>
      reject(request.error ?? new Error("IndexedDB request failed")),
    );
  });
}

/** Reads every persisted record — committed documents and autosave drafts
 * alike; callers split them by key prefix. */
export async function loadAllRecords(): Promise<PersistedRecord[]> {
  const db = await getDb();
  const store = db.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME);
  return requestToPromise(store.getAll() as IDBRequest<PersistedRecord[]>);
}

/** Reads one record by key, resolving undefined when absent. */
export async function getRecord(
  id: string,
): Promise<PersistedRecord | undefined> {
  const db = await getDb();
  const store = db.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME);
  return requestToPromise(
    store.get(id) as IDBRequest<PersistedRecord | undefined>,
  );
}

/** Writes (inserts or replaces) one record. */
export async function putRecord(record: PersistedRecord): Promise<void> {
  const db = await getDb();
  const store = db.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME);
  await requestToPromise(store.put(record));
}

/** Removes one document record by id. Resolves even if the id is absent. */
export async function deleteRecord(id: string): Promise<void> {
  const db = await getDb();
  const store = db.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME);
  await requestToPromise(store.delete(id));
}

// --- version history (SPEC-040) --------------------------------------------
// A separate object store, and separate helpers, so a version-history write
// or read is always its own IndexedDB transaction on its own store — a
// quota failure here can never abort a `documents` store transaction (AC-40).

/** Reads one version-history record (index, payload or blob) by key. */
export async function getVersionsRecord(id: string): Promise<VersionStoreRecord | undefined> {
  const db = await getDb();
  const store = db.transaction(VERSIONS_STORE_NAME, "readonly").objectStore(VERSIONS_STORE_NAME);
  return requestToPromise(store.get(id) as IDBRequest<VersionStoreRecord | undefined>);
}

/** Writes (inserts or replaces) one version-history record. */
export async function putVersionsRecord(record: VersionStoreRecord): Promise<void> {
  const db = await getDb();
  const store = db.transaction(VERSIONS_STORE_NAME, "readwrite").objectStore(VERSIONS_STORE_NAME);
  await requestToPromise(store.put(record));
}

/** Removes one version-history record by id. Resolves even if absent. */
export async function deleteVersionsRecord(id: string): Promise<void> {
  const db = await getDb();
  const store = db.transaction(VERSIONS_STORE_NAME, "readwrite").objectStore(VERSIONS_STORE_NAME);
  await requestToPromise(store.delete(id));
}

// --- user templates (SPEC-046) ----------------------------------------------

/** Reads every persisted user template. */
export async function loadAllTemplateRecords(): Promise<StoredTemplateRecord[]> {
  const db = await getDb();
  const store = db.transaction(TEMPLATES_STORE_NAME, "readonly").objectStore(TEMPLATES_STORE_NAME);
  return requestToPromise(store.getAll() as IDBRequest<StoredTemplateRecord[]>);
}

/** Writes (inserts or replaces) one user template record. */
export async function putTemplateRecord(record: StoredTemplateRecord): Promise<void> {
  const db = await getDb();
  const store = db.transaction(TEMPLATES_STORE_NAME, "readwrite").objectStore(TEMPLATES_STORE_NAME);
  await requestToPromise(store.put(record));
}

/** Removes one user template record by id. Resolves even if absent. */
export async function deleteTemplateRecord(id: string): Promise<void> {
  const db = await getDb();
  const store = db.transaction(TEMPLATES_STORE_NAME, "readwrite").objectStore(TEMPLATES_STORE_NAME);
  await requestToPromise(store.delete(id));
}
