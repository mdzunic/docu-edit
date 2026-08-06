// requestIdleCallback with a requestAnimationFrame fallback (Safari,
// jsdom) — shared by the pagination controller and the find-highlight
// measurement (SPEC-028), so expensive layout reads run in browser idle
// time instead of competing with user input. In a lib because the
// espree-parsed `.vue` scripts cannot write the feature-detection casts
// themselves.

/** A cancellable slot of scheduled background work. */
export interface IdleHandle {
  cancel(): void;
}

/** Typed `null` seed for idle-handle variables in `.vue` scripts (espree
 * bars the annotation — see domCaret.noElement). */
export function noIdleHandle(): IdleHandle | null {
  return null;
}

/** Runs `callback` in the next idle period (forced after `timeoutMs` under
 * sustained load), passing the slot's time budget in ms — the idle
 * deadline's remaining time clamped to `maxBudgetMs`, or `fallbackBudgetMs`
 * on the rAF fallback path. */
export function scheduleIdleWork(
  callback: (budgetMs: number) => void,
  options: { timeoutMs: number; fallbackBudgetMs: number; maxBudgetMs: number },
): IdleHandle {
  // lib.dom types requestIdleCallback on Window; Safari/jsdom lack it at
  // runtime, hence the typeof guard rather than a plain call.
  if (typeof window.requestIdleCallback === "function") {
    const handle = window.requestIdleCallback(
      (deadline) =>
        callback(Math.min(options.maxBudgetMs, Math.max(1, deadline.timeRemaining()))),
      { timeout: options.timeoutMs },
    );
    return {
      cancel() {
        if (typeof window.cancelIdleCallback === "function") {
          window.cancelIdleCallback(handle);
        }
      },
    };
  }
  const handle = window.requestAnimationFrame(() => callback(options.fallbackBudgetMs));
  return {
    cancel() {
      window.cancelAnimationFrame(handle);
    },
  };
}
