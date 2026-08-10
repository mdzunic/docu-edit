// The version-history cadence timer (SPEC-040 AC-1): a deadline armed once
// on the first edit after being idle, firing exactly `intervalMs` later
// regardless of how much editing happens in between — the same "deadline"
// half of autosaveScheduler.ts's pattern (its 30 s max-interval timer), one
// order of magnitude coarser and with no debounce component, since a
// debounce that reset on every keystroke would let continuous typing push
// the cadence out indefinitely instead of firing "every 5 minutes of
// editing activity".

export interface VersionCadenceTimerOptions {
  /** Called when the deadline fires. */
  onDue: () => void;
  intervalMs?: number;
}

export interface VersionCadenceTimer {
  /** Report an edit; arms the deadline if none is already pending. */
  noteChange(): void;
  /** Clears any pending deadline without disabling future ones — a document
   * switch's session-close hook, so a timer left over from the previous
   * document never fires against the next one. */
  reset(): void;
  dispose(): void;
}

export const VERSION_CADENCE_DEFAULT_MS = 5 * 60 * 1000;

export function createVersionCadenceTimer(options: VersionCadenceTimerOptions): VersionCadenceTimer {
  const intervalMs = options.intervalMs ?? VERSION_CADENCE_DEFAULT_MS;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let disposed = false;

  return {
    noteChange(): void {
      if (disposed || timer !== null) {
        return;
      }
      timer = setTimeout(() => {
        timer = null;
        options.onDue();
      }, intervalMs);
    },
    reset(): void {
      if (timer !== null) {
        clearTimeout(timer);
        timer = null;
      }
    },
    dispose(): void {
      disposed = true;
      if (timer !== null) {
        clearTimeout(timer);
        timer = null;
      }
    },
  };
}
