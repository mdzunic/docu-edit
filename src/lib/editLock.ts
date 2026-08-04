// Cross-tab single-writer lock per document (SPEC-010), built on the Web
// Locks API. Chosen over BroadcastChannel/localStorage heartbeats because a
// web lock is released by the browser the instant its tab closes OR crashes
// — no stale-lock cleanup — and a queued request resolves the moment the
// holder goes away, so the read-only tab becomes editable event-driven, with
// no polling.

export type EditLockState = "acquired" | "busy";

export interface EditLockHandle {
  /** Releases a held lock and/or cancels a queued request. Idempotent. */
  release(): void;
}

/** The slice of LockManager this module uses; injectable so unit tests can
 * run without a real Web Locks implementation (jsdom has none). */
export interface LockManagerLike {
  request(
    name: string,
    options: { ifAvailable?: boolean; signal?: AbortSignal },
    callback: (lock: unknown) => Promise<void> | void,
  ): Promise<unknown>;
}

/** Typed `null` seed for lock-handle variables in `.vue` scripts, which
 * espree bars from writing the `EditLockHandle | null` annotation. */
export function noEditLock(): EditLockHandle | null {
  return null;
}

function defaultLockManager(): LockManagerLike | undefined {
  if (typeof navigator !== "undefined" && "locks" in navigator) {
    return navigator.locks;
  }
  return undefined;
}

/** Requests the exclusive edit lock for a document. Reports "acquired" when
 * this tab holds the lock (possibly later, after another tab releases it) or
 * "busy" when another tab currently holds it. Without a Web Locks API the
 * lock degrades to always-acquired — single-tab editing still works.
 *
 * The lock is held until release() — or until the tab dies, at which point
 * the browser releases it for us (that is the crash story). */
export function acquireEditLock(
  docId: string,
  onStatus: (state: EditLockState) => void,
  locks: LockManagerLike | undefined = defaultLockManager(),
): EditLockHandle {
  if (!locks) {
    onStatus("acquired");
    return { release() {} };
  }
  const name = `docu-edit:doc:${docId}`;
  let released = false;
  let releaseHeld = (): void => {};
  // Resolving this promise is what ends the lock-holding callback below,
  // handing the lock back to the browser.
  const held = new Promise<void>((resolve) => {
    releaseHeld = resolve;
  });
  const controller = new AbortController();

  void locks
    .request(name, { ifAvailable: true }, (lock) => {
      if (released) {
        return undefined;
      }
      if (lock) {
        onStatus("acquired");
        return held;
      }
      onStatus("busy");
      // Queue behind the current holder; grants when that tab closes.
      void locks
        .request(name, { signal: controller.signal }, (queuedLock) => {
          if (released || !queuedLock) {
            return undefined;
          }
          onStatus("acquired");
          return held;
        })
        .catch(() => {
          // AbortError when release() cancels the queued request.
        });
      return undefined;
    })
    .catch(() => {
      // A failing lock manager degrades like a missing one, minus the
      // acquired signal — the caller stays in its pre-lock state.
    });

  return {
    release(): void {
      if (released) {
        return;
      }
      released = true;
      releaseHeld();
      controller.abort();
    },
  };
}
