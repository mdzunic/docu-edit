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

// One shared channel with a FIFO of scheduled slots: each postMessage runs
// exactly one queued slot (a cancelled slot runs as a no-op). Slots are
// one-shot (the callback reschedules), matching scheduleIdleWork's shape.
let eagerChannel: MessageChannel | null = null;
const eagerQueue: Array<() => void> = [];

/** Runs `callback` on the next macrotask via a MessageChannel — the
 * highest-duty way to chain background slots: unlike requestIdleCallback
 * there is no frame-aligned wait between slots, while input events and
 * rendering still interleave between macrotasks. `budgetMs` is passed
 * through as the slot's time budget; the caller sizes it to bound how long
 * input can be delayed. Falls back to requestAnimationFrame where
 * MessageChannel is unavailable. */
export function scheduleEagerWork(
  callback: (budgetMs: number) => void,
  budgetMs: number,
): IdleHandle {
  if (typeof MessageChannel !== "function") {
    const handle = window.requestAnimationFrame(() => callback(budgetMs));
    return {
      cancel() {
        window.cancelAnimationFrame(handle);
      },
    };
  }
  if (!eagerChannel) {
    eagerChannel = new MessageChannel();
    eagerChannel.port1.addEventListener("message", () => {
      const run = eagerQueue.shift();
      if (run) {
        run();
      }
    });
    // addEventListener (unlike onmessage) does not start the port.
    eagerChannel.port1.start();
  }
  let cancelled = false;
  const slot = (): void => {
    if (!cancelled) {
      callback(budgetMs);
    }
  };
  eagerQueue.push(slot);
  // eslint-disable-next-line unicorn/require-post-message-target-origin -- MessagePort.postMessage has no targetOrigin parameter (its second argument is a transfer list).
  eagerChannel.port2.postMessage(null);
  return {
    cancel() {
      cancelled = true;
    },
  };
}

/** A single macrotask yield (SPEC-047 AC-15/AC-16's main-thread-block
 * ceilings): `await`ed inside a long synchronous loop (parseDocx's body
 * walk, the zip writer's entry loop) so a large document's parse/save
 * reads as many short tasks instead of one long one. Unlike awaiting a
 * promise that resolves via microtask (which never actually hands control
 * back to the browser's input/render queue), scheduleEagerWork's
 * MessageChannel post is a real macrotask boundary. */
export function yieldToMainThread(): Promise<void> {
  return new Promise((resolve) => {
    scheduleEagerWork(() => resolve(), 1);
  });
}
