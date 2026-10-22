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

/** Which ends of a key range are open (excluded); both closed by default. */
export interface KeyRangeBounds {
  lowerOpen?: boolean;
  upperOpen?: boolean;
}

/** One operation of a writeBatch: a put (insert or replace) of a record, a
 * delete of one key, or a delete of every key in the closed range
 * [lower, upper]. */
export type WriteBatchOp = { put: unknown } | { delete: string } | { deleteRange: [string, string] };

/** A lazily opened IndexedDB database: the connection is established on
 * first use and shared by every subsequent call. */
export interface IdbDatabase {
  /** The open (or opening) connection — exposed for callers that need a
   * transaction shape these helpers do not cover. */
  db(): Promise<IDBDatabase>;
  getAll<T>(storeName: string): Promise<T[]>;
  /** SPEC-137: every record whose key lies between `lower` and `upper`
   * (either side absent = unbounded), in key order. */
  getAllInRange<T>(storeName: string, lower?: string, upper?: string, bounds?: KeyRangeBounds): Promise<T[]>;
  get<T>(storeName: string, key: string): Promise<T | undefined>;
  put(storeName: string, record: unknown): Promise<void>;
  delete(storeName: string, key: string): Promise<void>;
  /** SPEC-137: deletes every record whose key lies in [lower, upper]. */
  deleteRange(storeName: string, lower: string, upper: string): Promise<void>;
  /** SPEC-137: applies `ops` in order inside ONE readwrite transaction —
   * all of them commit or none does — issuing each operation from the
   * previous one's success callback, so their structured clones run in
   * separate tasks instead of fusing into one long one. Resolves when the
   * transaction completes; rejects (the transaction aborted, nothing
   * written) on any failure. */
  writeBatch(storeName: string, ops: readonly WriteBatchOp[]): Promise<void>;
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.addEventListener("error", () =>
      reject(request.error ?? new Error("IndexedDB request failed")),
    );
  });
}

/** The IDBKeyRange for a (lower, upper) pair, either side optional;
 * undefined (every key) when both are absent. */
function keyRange(lower: string | undefined, upper: string | undefined, bounds: KeyRangeBounds = {}): IDBKeyRange | undefined {
  if (lower !== undefined && upper !== undefined) {
    return IDBKeyRange.bound(lower, upper, bounds.lowerOpen === true, bounds.upperOpen === true);
  }
  if (lower !== undefined) {
    return IDBKeyRange.lowerBound(lower, bounds.lowerOpen === true);
  }
  if (upper !== undefined) {
    return IDBKeyRange.upperBound(upper, bounds.upperOpen === true);
  }
  return undefined;
}

/** Issues one batch operation against `store` — the put's structured clone
 * happens here, synchronously. */
function issue(store: IDBObjectStore, op: WriteBatchOp): IDBRequest {
  if ("put" in op) {
    return store.put(op.put);
  }
  if ("delete" in op) {
    return store.delete(op.delete);
  }
  return store.delete(IDBKeyRange.bound(op.deleteRange[0], op.deleteRange[1]));
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
    async getAllInRange<T>(storeName: string, lower?: string, upper?: string, bounds?: KeyRangeBounds): Promise<T[]> {
      const db = await getDb();
      const store = db.transaction(storeName, "readonly").objectStore(storeName);
      return requestToPromise(store.getAll(keyRange(lower, upper, bounds)) as IDBRequest<T[]>);
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
    async deleteRange(storeName: string, lower: string, upper: string): Promise<void> {
      const db = await getDb();
      const store = db.transaction(storeName, "readwrite").objectStore(storeName);
      await requestToPromise(store.delete(IDBKeyRange.bound(lower, upper)));
    },
    async writeBatch(storeName: string, ops: readonly WriteBatchOp[]): Promise<void> {
      if (ops.length === 0) {
        return;
      }
      const db = await getDb();
      await new Promise<void>((resolve, reject) => {
        const transaction = db.transaction(storeName, "readwrite");
        const store = transaction.objectStore(storeName);
        transaction.oncomplete = () => resolve();
        // A failed request's error event bubbles here before the abort;
        // whichever comes first settles the promise.
        transaction.addEventListener("error", () => reject(transaction.error ?? new Error("IndexedDB batch failed")));
        transaction.addEventListener("abort", () => reject(transaction.error ?? new Error("IndexedDB batch aborted")));
        let next = 0;
        // One operation per success callback: the transaction stays active
        // inside a request's callback, and nothing else is pending while the
        // next operation's clone runs — so a 49-chunk batch is 49 short tasks,
        // never one long one. After the last success the callback issues
        // nothing and the transaction commits.
        const issueNext = (): void => {
          const op = ops[next];
          if (op === undefined) {
            return;
          }
          next += 1;
          let request: IDBRequest;
          try {
            request = issue(store, op);
          } catch (error) {
            // A synchronous refusal (DataCloneError, a bad key): roll back
            // what this batch already wrote.
            try {
              transaction.abort();
            } catch {
              // Already finished — nothing to roll back.
            }
            reject(error);
            return;
          }
          request.onsuccess = issueNext;
        };
        issueNext();
      });
    },
  };
}
