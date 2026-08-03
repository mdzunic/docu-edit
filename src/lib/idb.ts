// Thin promise wrappers around the IndexedDB persistence for the document
// library. Kept as its own module so documentStore can be unit-tested with
// this seam mocked, and so nothing else touches IndexedDB directly.
//
// Every function rejects when IndexedDB is missing, blocked, or over quota —
// callers (documentStore) treat any rejection as "storage unavailable" and
// keep working in-memory.
import type { StoredDocumentRecord } from "./documentStore";

const DB_NAME = "docu-edit";
const DB_VERSION = 1;
const STORE_NAME = "documents";

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

/** Reads every persisted document record. */
export async function loadAllRecords(): Promise<StoredDocumentRecord[]> {
  const db = await getDb();
  const store = db.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME);
  return requestToPromise(store.getAll() as IDBRequest<StoredDocumentRecord[]>);
}

/** Writes (inserts or replaces) one document record. */
export async function putRecord(record: StoredDocumentRecord): Promise<void> {
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
