import { describe, it, expect, vi } from "vitest";
import { acquireEditLock, type LockManagerLike } from "./editLock";

/** The initial value of a callback slot filled in later (a promise executor,
 * a registration call) — shared so the slot's initialiser is not a nested
 * function expression of its own. */
const noop = (): void => {};

async function microtasks(): Promise<void> {
  for (let i = 0; i < 10; i += 1) {
    await Promise.resolve();
  }
}

/** An in-memory LockManager: one exclusive lock per name, FIFO queue,
 * ifAvailable and AbortSignal semantics — the parts acquireEditLock uses. */
function fakeLockManager(): LockManagerLike {
  const holders = new Map<string, Array<() => void>>();

  function locked(name: string): boolean {
    return (holders.get(name) ?? []).length > 0;
  }

  async function grant(
    name: string,
    callback: (lock: unknown) => Promise<void> | void,
  ): Promise<void> {
    let done: () => void = noop;
    const release = new Promise<void>((resolve) => {
      done = resolve;
    });
    const queue = holders.get(name) ?? [];
    holders.set(name, [...queue, done]);
    try {
      await callback({ name });
    } finally {
      const current = holders.get(name) ?? [];
      const next = current.slice(1);
      holders.set(name, next);
      done();
      void release;
    }
  }

  const waiting = new Map<string, Array<() => void>>();

  function releaseNext(name: string): void {
    const queue = waiting.get(name) ?? [];
    const next = queue.shift();
    if (next) {
      next();
    }
  }

  return {
    async request(name, options, callback) {
      if (options.ifAvailable && locked(name)) {
        await callback(null);
        return;
      }
      if (locked(name)) {
        // Queue until the holder finishes.
        await new Promise<void>((resolve, reject) => {
          const entry = (): void => resolve();
          waiting.set(name, [...(waiting.get(name) ?? []), entry]);
          options.signal?.addEventListener("abort", () => {
            const queue = waiting.get(name) ?? [];
            waiting.set(
              name,
              queue.filter((w) => w !== entry),
            );
            reject(new DOMException("aborted", "AbortError"));
          });
        });
        if (options.signal?.aborted) {
          throw new DOMException("aborted", "AbortError");
        }
      }
      await grant(name, async (lock) => {
        await callback(lock);
        releaseNext(name);
      });
    },
  };
}

describe("acquireEditLock", () => {
  it("acquires immediately when the lock is free", async () => {
    const locks = fakeLockManager();
    const onStatus = vi.fn();
    const handle = acquireEditLock("doc-1", onStatus, locks);
    await microtasks();
    expect(onStatus).toHaveBeenCalledWith("acquired");
    handle.release();
  });

  it("reports busy while another tab holds the lock, acquired once it releases", async () => {
    const locks = fakeLockManager();
    const first = acquireEditLock("doc-2", vi.fn(), locks);
    await microtasks();

    const onStatus = vi.fn();
    const second = acquireEditLock("doc-2", onStatus, locks);
    await microtasks();
    expect(onStatus).toHaveBeenCalledWith("busy");
    expect(onStatus).not.toHaveBeenCalledWith("acquired");

    first.release();
    await microtasks();
    expect(onStatus).toHaveBeenCalledWith("acquired");
    second.release();
  });

  it("locks per document id — different documents do not conflict", async () => {
    const locks = fakeLockManager();
    const a = acquireEditLock("doc-a", vi.fn(), locks);
    const onStatus = vi.fn();
    const b = acquireEditLock("doc-b", onStatus, locks);
    await microtasks();
    expect(onStatus).toHaveBeenCalledWith("acquired");
    a.release();
    b.release();
  });

  it("release before the queued grant never reports acquired", async () => {
    const locks = fakeLockManager();
    const first = acquireEditLock("doc-3", vi.fn(), locks);
    await microtasks();

    const onStatus = vi.fn();
    const second = acquireEditLock("doc-3", onStatus, locks);
    await microtasks();
    second.release();
    first.release();
    await microtasks();
    expect(onStatus).toHaveBeenCalledWith("busy");
    expect(onStatus).not.toHaveBeenCalledWith("acquired");
  });

  it("releasing lets the next queued tab acquire", async () => {
    const locks = fakeLockManager();
    const first = acquireEditLock("doc-4", vi.fn(), locks);
    await microtasks();
    const onStatus = vi.fn();
    const second = acquireEditLock("doc-4", onStatus, locks);
    await microtasks();

    first.release();
    await microtasks();
    expect(onStatus).toHaveBeenLastCalledWith("acquired");
    second.release();

    // And after the second releases too, a third acquires immediately.
    const third = vi.fn();
    const handle = acquireEditLock("doc-4", third, locks);
    await microtasks();
    expect(third).toHaveBeenCalledWith("acquired");
    handle.release();
  });

  it("degrades to acquired when no Web Locks API exists (jsdom, old browsers)", () => {
    const onStatus = vi.fn();
    const handle = acquireEditLock("doc-5", onStatus, undefined);
    expect(onStatus).toHaveBeenCalledWith("acquired");
    handle.release();
  });
});
