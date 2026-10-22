// The find bar's walk driver (SPEC-134 "The walk", "Scope and order"): runs
// a lib/spreadsheet/find.ts SheetSearch one slot at a time — each a
// `schedule` call (core's scheduleEagerWork in the app) stepping until
// FIND_SLICE_BUDGET_MS by `now` — restarts it after FIND_DEBOUNCE_MS when
// what it walked over changed (`key`: the scope and every sheet's `rows`
// reference, compared item by item), and answers the one waiting request —
// the first hit at or after / after / the last before a cell, or "keep" (the
// hit at that cell stays current, nothing moves) — as soon as the walk can.
// Plain callbacks, no Vue: useSheetFind owns the state it reports into.
import { FIND_DEBOUNCE_MS, FIND_SLICE_BUDGET_MS, type FindHit, type SheetSearch } from "../../lib/spreadsheet/find";
import type { IdleHandle } from "../../lib/core/idleScheduler";

export type FindRequestMode = "atOrAfter" | "after" | "before" | "keep";

export interface FindRequest {
  mode: FindRequestMode;
  from: FindHit;
}

/** What a request came to: no hit, the hit at the cell kept current (or
 * none, -1), or a hit to select (`wrapped` past an end). */
export type FindAnswer = { kind: "none" } | { kind: "keep"; index: number } | { kind: "select"; index: number; wrapped: boolean };

export interface FindWalkDeps {
  /** A fresh search over the current query, options, model and selection;
   * null when there is nothing to search. */
  create(): SheetSearch | null;
  /** What the walk depends on — equal keys mean nothing to restart for. */
  key(): readonly unknown[];
  /** The active cell, where a request starts by default. */
  from(): FindHit | null;
  /** The search, its hits or its state changed. */
  changed(): void;
  answered(answer: FindAnswer): void;
  schedule: (callback: (budgetMs: number) => void, budgetMs: number) => IdleHandle;
  now: () => number;
}

const sameKey = (a: readonly unknown[], b: readonly unknown[]): boolean => a.length === b.length && a.every((item, i) => item === b[i]);

export function createFindWalk(deps: FindWalkDeps) {
  let search: SheetSearch | null = null;
  let slot: IdleHandle | null = null;
  let debounce: ReturnType<typeof setTimeout> | null = null;
  /** What a restart after the debounce should answer (none: keep). */
  let queued: FindRequestMode | undefined;
  let request: FindRequest | null = null;
  let walkedKey: readonly unknown[] = [];

  function stop(): void {
    slot?.cancel();
    slot = null;
    if (debounce !== null) {
      clearTimeout(debounce);
      debounce = null;
    }
    queued = undefined;
  }

  /** Answers the waiting request once the walk can. */
  function answer(): void {
    if (!search || !request) {
      return;
    }
    const pending = request;
    const found = search.locate(pending.from, pending.mode === "keep" ? "atOrAfter" : pending.mode);
    if (found.kind === "pending") {
      return;
    }
    request = null;
    if (found.kind === "none") {
      deps.answered(found);
      return;
    }
    if (pending.mode !== "keep") {
      deps.answered({ kind: "select", index: found.index, wrapped: found.wrapped });
      return;
    }
    const hit = search.hits[found.index]!;
    const same = hit.sheetId === pending.from.sheetId && hit.row === pending.from.row && hit.col === pending.from.col;
    deps.answered({ kind: "keep", index: same ? found.index : -1 });
  }

  /** One slot: step until the budget, then answer and schedule the next. */
  function run(): void {
    const walk = search;
    slot = null;
    if (!walk) {
      return;
    }
    const start = deps.now();
    walk.step(() => deps.now() - start >= FIND_SLICE_BUDGET_MS);
    deps.changed();
    answer();
    if (!walk.done && search === walk) {
      slot = deps.schedule(run, FIND_SLICE_BUDGET_MS);
    }
  }

  /** Starts afresh; `mode` from `from` is what to answer — default: a
   * waiting request, else "keep" at the active cell. */
  function begin(mode?: FindRequestMode, from: FindHit | null = deps.from()): void {
    const waiting = mode === undefined ? request : null;
    stop();
    walkedKey = deps.key();
    search = deps.create();
    request = search && from ? (waiting ?? { mode: mode ?? "keep", from }) : null;
    deps.changed();
    if (search) {
      slot = deps.schedule(run, FIND_SLICE_BUDGET_MS);
    }
  }

  /** Restarts after the debounce — always for a `mode`, else only when the
   * key changed. A queued mode survives a later plain call. */
  function soon(mode?: FindRequestMode): void {
    queued = mode ?? queued;
    if (debounce !== null) {
      clearTimeout(debounce);
    }
    debounce = setTimeout(() => {
      const next = queued;
      debounce = null;
      queued = undefined;
      if (next !== undefined || !sameKey(walkedKey, deps.key())) {
        begin(next);
      }
    }, FIND_DEBOUNCE_MS);
  }

  /** A navigation request: answered now if the walk can, else when it can
   * (a walk not started yet, or stale — a query or option change waiting, or
   * what it walked over changed — starts afresh now). */
  function ask(mode: "after" | "before", from: FindHit): void {
    if (!search || queued !== undefined || !sameKey(walkedKey, deps.key())) {
      begin(mode, from);
      return;
    }
    request = { mode, from };
    answer();
  }

  /** Ends the walk: no search, nothing waiting. */
  function clear(): void {
    stop();
    search = null;
    request = null;
    deps.changed();
  }

  return {
    get search(): SheetSearch | null {
      return search;
    },
    begin,
    soon,
    ask,
    clear,
    stop,
  };
}

export type FindWalk = ReturnType<typeof createFindWalk>;
