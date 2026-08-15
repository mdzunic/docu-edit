// A generic promise layer over IndexedDB: open a named database with a set
// of keyPath'd object stores, then get/put/delete/getAll against one store.
// Product-neutral — lib/idb.ts wires docu-edit's own schema (documents,
// versions, templates) through this; a Slides product would open its own
// database the same way.
//
// Every function rejects when IndexedDB is missing, blocked, or over quota —
// callers treat any rejection as "storage unavailable" and keep working
// in-memory.

/** One object store in a database's schema, created on upgrade. */
export interface StoreSchema {
  name: string;
  keyPath: string;
}

/** A lazily opened IndexedDB database: the connection is established on
 * first use and shared by every subsequent call. */
export interface IdbDatabase {
  /** The open (or opening) connection — exposed for callers that need a
   * transaction shape these helpers do not cover. */
  db(): Promise<IDBDatabase>;
  getAll<T>(storeName: string): Promise<T[]>;
  get<T>(storeName: string, key: string): Promise<T | undefined>;
  put(storeName: string, record: unknown): Promise<void>;
  delete(storeName: string, key: string): Promise<void>;
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.addEventListener("error", () =>
      reject(request.error ?? new Error("IndexedDB request failed")),
    );
  });
}

/** Opens (lazily, once) a database, creating any missing `stores` on
 * upgrade. Never throws synchronously: an absent `indexedDB` (jsdom,
 * locked-down browsers) rejects the first operation instead. */
export function openIdbDatabase(name: string, version: number, stores: StoreSchema[]): IdbDatabase {
  let dbPromise: Promise<IDBDatabase> | null = null;

  function openDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      // If `indexedDB` is absent this throws synchronously inside the
      // executor, which rejects the promise.
      const request = indexedDB.open(name, version);
      request.onupgradeneeded = () => {
        const db = request.result;
        for (const store of stores) {
          if (!db.objectStoreNames.contains(store.name)) {
            db.createObjectStore(store.name, { keyPath: store.keyPath });
          }
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

  return {
    db: getDb,
    async getAll<T>(storeName: string): Promise<T[]> {
      const db = await getDb();
      const store = db.transaction(storeName, "readonly").objectStore(storeName);
      return requestToPromise(store.getAll() as IDBRequest<T[]>);
    },
    async get<T>(storeName: string, key: string): Promise<T | undefined> {
      const db = await getDb();
      const store = db.transaction(storeName, "readonly").objectStore(storeName);
      return requestToPromise(store.get(key) as IDBRequest<T | undefined>);
    },
    async put(storeName: string, record: unknown): Promise<void> {
      const db = await getDb();
      const store = db.transaction(storeName, "readwrite").objectStore(storeName);
      await requestToPromise(store.put(record));
    },
    async delete(storeName: string, key: string): Promise<void> {
      const db = await getDb();
      const store = db.transaction(storeName, "readwrite").objectStore(storeName);
      await requestToPromise(store.delete(key));
    },
  };
}
