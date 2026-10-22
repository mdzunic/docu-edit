// TEST-ONLY: a small in-memory IndexedDB for the idb specs (SPEC-137) —
// jsdom has none and the project takes no fake-indexeddb dependency. Never
// imported by the app. It covers what core/idb.ts uses, string keys only:
// open/upgrade, readonly and readwrite transactions on one store,
// get/getAll(range)/put/delete(key or range), IDBKeyRange.bound/lowerBound/
// upperBound. Requests complete asynchronously, one per macrotask; a
// transaction is active only while it is being created and inside a
// request's success callback (issuing at any other time throws, as a real
// transaction does), commits when a callback issues nothing more, and rolls
// back on abort. Values are structured-cloned on put and on read, and every
// read and write is logged so a spec can assert what was touched.

export interface FakeIdbEvent {
  kind: "issue" | "success" | "read" | "complete" | "abort";
  op?: "get" | "getAll" | "put" | "delete";
  key?: string;
  transaction: number;
}

class FakeKeyRange {
  constructor(
    readonly lower: string | undefined,
    readonly upper: string | undefined,
    readonly lowerOpen: boolean,
    readonly upperOpen: boolean,
  ) {}
  static bound(lower: string, upper: string, lowerOpen = false, upperOpen = false): FakeKeyRange {
    return new FakeKeyRange(lower, upper, lowerOpen, upperOpen);
  }
  static lowerBound(lower: string, open = false): FakeKeyRange {
    return new FakeKeyRange(lower, undefined, open, false);
  }
  static upperBound(upper: string, open = false): FakeKeyRange {
    return new FakeKeyRange(undefined, upper, false, open);
  }
  includes(key: string): boolean {
    if (this.lower !== undefined && (this.lowerOpen ? key <= this.lower : key < this.lower)) {
      return false;
    }
    return !(this.upper !== undefined && (this.upperOpen ? key >= this.upper : key > this.upper));
  }
}

type Listener = (event: { target: unknown }) => void;

class FakeRequest {
  result: unknown;
  error: Error | null = null;
  onsuccess: Listener | null = null;
  onerror: Listener | null = null;
  onupgradeneeded: Listener | null = null;
  private listeners: Record<string, Listener[]> = {};
  addEventListener(type: string, listener: Listener): void {
    (this.listeners[type] ??= []).push(listener);
  }
  fire(type: "success" | "error" | "upgradeneeded"): void {
    const handler = type === "success" ? this.onsuccess : type === "error" ? this.onerror : this.onupgradeneeded;
    handler?.({ target: this });
    for (const listener of this.listeners[type] ?? []) {
      listener({ target: this });
    }
  }
}

/** The whole fake: install() puts it on globalThis; `stores` is the data. */
export class FakeIndexedDb {
  readonly stores = new Map<string, Map<string, unknown>>();
  readonly log: FakeIdbEvent[] = [];
  private transactions = 0;
  /** Makes the next put of a key starting with this prefix fail. */
  failPutPrefix: string | null = null;

  install(): void {
    Object.assign(globalThis, { indexedDB: this.factory(), IDBKeyRange: FakeKeyRange });
  }

  uninstall(): void {
    const target = globalThis as Record<string, unknown>;
    delete target.indexedDB;
    delete target.IDBKeyRange;
  }

  /** Every key read (by get or getAll) — what a spec asserts was touched. */
  readKeys(): string[] {
    return this.log.filter((event) => event.kind === "read").map((event) => event.key!);
  }

  private factory() {
    const fake = this;
    return {
      open(_name: string, _version: number) {
        const request = new FakeRequest();
        setTimeout(() => {
          const db = fake.database();
          request.result = db;
          request.fire("upgradeneeded");
          request.fire("success");
        }, 0);
        return request;
      },
    };
  }

  private database() {
    const fake = this;
    return {
      objectStoreNames: { contains: (name: string) => fake.stores.has(name) },
      createObjectStore(name: string) {
        fake.stores.set(name, new Map());
      },
      transaction(name: string, mode: "readonly" | "readwrite") {
        return fake.transaction(name, mode);
      },
    };
  }

  private transaction(name: string, mode: "readonly" | "readwrite") {
    const fake = this;
    const id = (this.transactions += 1);
    const data = this.stores.get(name)!;
    const journal = new Map<string, unknown>();
    const queue: { request: FakeRequest; run: () => unknown }[] = [];
    let active = true;
    let finished = false;
    const listeners: Record<string, Listener[]> = {};
    const transaction = {
      error: null as Error | null,
      oncomplete: null as Listener | null,
      addEventListener(type: string, listener: Listener) {
        (listeners[type] ??= []).push(listener);
      },
      objectStore: () => store,
      abort() {
        if (finished) {
          throw new Error("InvalidStateError");
        }
        finish("abort");
      },
    };
    const finish = (kind: "complete" | "abort"): void => {
      finished = true;
      if (kind === "abort") {
        for (const [key, before] of journal) {
          if (before === undefined) {
            data.delete(key);
          } else {
            data.set(key, before);
          }
        }
      }
      fake.log.push({ kind, transaction: id });
      if (kind === "complete") {
        transaction.oncomplete?.({ target: transaction });
      }
      for (const listener of listeners[kind] ?? []) {
        listener({ target: transaction });
      }
    };
    const remember = (key: string): void => {
      if (!journal.has(key)) {
        journal.set(key, data.get(key));
      }
    };
    const step = (): void => {
      setTimeout(() => {
        if (finished) {
          return;
        }
        const next = queue.shift();
        if (!next) {
          finish("complete");
          return;
        }
        try {
          next.request.result = next.run();
        } catch (error) {
          next.request.error = error as Error;
          transaction.error = error as Error;
          next.request.fire("error");
          for (const listener of listeners.error ?? []) {
            listener({ target: transaction });
          }
          finish("abort");
          return;
        }
        active = true;
        next.request.fire("success");
        // Active through the callback's microtasks, inactive from the next task.
        setTimeout(() => {
          active = false;
          step();
        }, 0);
      }, 0);
    };
    const enqueue = (op: FakeIdbEvent["op"], key: string | undefined, run: () => unknown): FakeRequest => {
      if (!active || finished) {
        throw new Error("TransactionInactiveError");
      }
      fake.log.push({ kind: "issue", op, key, transaction: id });
      const request = new FakeRequest();
      queue.push({
        request,
        run: () => {
          const result = run();
          fake.log.push({ kind: "success", op, key, transaction: id });
          return result;
        },
      });
      return request;
    };
    const keysIn = (range: FakeKeyRange | string | undefined): string[] =>
      [...data.keys()].filter((key) => (range === undefined ? true : typeof range === "string" ? key === range : range.includes(key))).toSorted();
    const read = (key: string): unknown => {
      fake.log.push({ kind: "read", key, transaction: id });
      return structuredClone(data.get(key));
    };
    const store = {
      get: (key: string) => enqueue("get", key, () => (data.has(key) ? read(key) : undefined)),
      getAll: (range?: FakeKeyRange) => enqueue("getAll", undefined, () => keysIn(range).map(read)),
      put: (value: { id: string }) => {
        if (mode !== "readwrite") {
          throw new Error("ReadOnlyError");
        }
        const copy = structuredClone(value);
        return enqueue("put", value.id, () => {
          if (fake.failPutPrefix !== null && value.id.startsWith(fake.failPutPrefix)) {
            fake.failPutPrefix = null;
            throw new Error("QuotaExceededError");
          }
          remember(value.id);
          data.set(value.id, copy);
          return value.id;
        });
      },
      delete: (key: string | FakeKeyRange) =>
        enqueue("delete", typeof key === "string" ? key : `${key.lower}..${key.upper}`, () => {
          for (const match of keysIn(key)) {
            remember(match);
            data.delete(match);
          }
        }),
    };
    setTimeout(() => {
      active = false;
      step();
    }, 0);
    return transaction;
  }
}
