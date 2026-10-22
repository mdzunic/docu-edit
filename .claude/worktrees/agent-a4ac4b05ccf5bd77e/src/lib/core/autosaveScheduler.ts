// The autosave cadence (SPEC-010): a flush fires 2 s after the last change
// (debounce) and no later than 30 s after the first unflushed change (the
// deadline that guarantees a save at least every 30 s during continuous
// typing). After a failed save with no typing to re-trigger the debounce, a
// retry timer keeps attempting so the failure state can clear on its own
// once storage recovers.
//
// The scheduler only decides WHEN to flush; the caller owns the actual
// write and reports back via notifySaveResult.

export const AUTOSAVE_DEBOUNCE_MS = 2000;
export const AUTOSAVE_MAX_INTERVAL_MS = 30_000;
export const AUTOSAVE_RETRY_MS = 5000;

export interface AutosaveSchedulerOptions {
  /** Called when a save should happen. */
  flush: () => void;
  /** The debounce interval, or a getter read at each change so a live change
   * to the "Autosave delay" preference (SPEC-066) takes effect on the next
   * edit without recreating the scheduler. */
  debounceMs?: number | (() => number);
  maxIntervalMs?: number;
  retryMs?: number;
}

export interface AutosaveScheduler {
  /** Report an edit; (re)arms the debounce and, if none is pending, the
   * 30 s deadline. */
  noteChange(): void;
  /** Report the outcome of a flush; a failure arms the retry timer unless
   * further typing already re-armed the debounce. */
  notifySaveResult(ok: boolean): void;
  /** Fire the flush immediately, cancelling all pending timers. */
  flushNow(): void;
  dispose(): void;
}

type Timer = ReturnType<typeof setTimeout> | null;

export function createAutosaveScheduler(
  options: AutosaveSchedulerOptions,
): AutosaveScheduler {
  const debounceOption = options.debounceMs;
  const resolveDebounceMs = (): number =>
    typeof debounceOption === "function" ? debounceOption() : (debounceOption ?? AUTOSAVE_DEBOUNCE_MS);
  const maxIntervalMs = options.maxIntervalMs ?? AUTOSAVE_MAX_INTERVAL_MS;
  const retryMs = options.retryMs ?? AUTOSAVE_RETRY_MS;

  let debounceTimer: Timer = null;
  let deadlineTimer: Timer = null;
  let retryTimer: Timer = null;
  let disposed = false;

  function clearAll(): void {
    if (debounceTimer !== null) {
      clearTimeout(debounceTimer);
      debounceTimer = null;
    }
    if (deadlineTimer !== null) {
      clearTimeout(deadlineTimer);
      deadlineTimer = null;
    }
    if (retryTimer !== null) {
      clearTimeout(retryTimer);
      retryTimer = null;
    }
  }

  function fire(): void {
    clearAll();
    options.flush();
  }

  return {
    noteChange(): void {
      if (disposed) {
        return;
      }
      if (debounceTimer !== null) {
        clearTimeout(debounceTimer);
      }
      debounceTimer = setTimeout(fire, resolveDebounceMs());
      if (deadlineTimer === null) {
        deadlineTimer = setTimeout(fire, maxIntervalMs);
      }
    },
    notifySaveResult(ok: boolean): void {
      if (disposed || ok) {
        return;
      }
      if (debounceTimer === null && retryTimer === null) {
        retryTimer = setTimeout(fire, retryMs);
      }
    },
    flushNow(): void {
      if (disposed) {
        return;
      }
      fire();
    },
    dispose(): void {
      disposed = true;
      clearAll();
    },
  };
}
