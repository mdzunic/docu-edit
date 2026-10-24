// The deck search walk (SPEC-098 "Performance rules"), split out of
// useDeckSearch.ts the way sheetFindWalk.ts serves the sheet's find. A deck of
// up to WALK_BATCH_THRESHOLD slides walks at once; a bigger one walks
// WALK_BATCH_SLIDES slides per batch inside scheduleIdleWork slots, handing
// each slot's hits on as they arrive so the first group paints immediately.
// The walk is findInSlide over plain model arrays — no commit(), no
// re-identification — and it stamps the clock SPEC-098's budgets read: from
// the instant it starts (the input debounce firing) to the first hits and to
// settling.
import { scheduleIdleWork, type IdleHandle } from "../../lib/core/idleScheduler";
import type { CompiledQuery } from "../../lib/core/findQuery";
import type { PresentationModel } from "../../lib/presentation/model";
import { findInDeck, findInSlide, type DeckHit } from "../../lib/presentation/deckSearch";

/** Decks over this many slides walk in idle slices. */
export const WALK_BATCH_THRESHOLD = 40;
/** Slides per batch of a sliced walk. */
export const WALK_BATCH_SLIDES = 10;

/** performance.now() stamps: the walk's start, its first hits and settling. */
export interface WalkTiming {
  startedAt: number;
  firstHitsAt?: number;
  settledAt?: number;
}

export interface DeckWalkDeps {
  /** Test seams: the idle-slot scheduler and the clock. */
  schedule?: (callback: (budgetMs: number) => void) => IdleHandle;
  now?: () => number;
}

export interface DeckWalk {
  /** Starts a walk, replacing any running one: `update` receives the hits so
   * far (a fresh array) and whether the walk has settled. */
  start(model: PresentationModel, compiled: CompiledQuery, update: (hits: DeckHit[], done: boolean, timing: WalkTiming) => void): void;
  cancel(): void;
  readonly running: boolean;
}

export function createDeckWalk(deps: DeckWalkDeps = {}): DeckWalk {
  const now = deps.now ?? ((): number => performance.now());
  const schedule =
    deps.schedule ?? ((callback: (budgetMs: number) => void): IdleHandle => scheduleIdleWork(callback, { timeoutMs: 50, fallbackBudgetMs: 8, maxBudgetMs: 12 }));
  let slot: IdleHandle | null = null;
  let token = 0;

  function cancel(): void {
    token += 1;
    slot?.cancel();
    slot = null;
  }

  function start(model: PresentationModel, compiled: CompiledQuery, update: (hits: DeckHit[], done: boolean, timing: WalkTiming) => void): void {
    cancel();
    const timing: WalkTiming = { startedAt: now() };
    const total = model.slides.length;
    if (total <= WALK_BATCH_THRESHOLD) {
      const hits = findInDeck(model, compiled);
      timing.firstHitsAt = timing.settledAt = now();
      update(hits, true, timing);
      return;
    }
    const mine = token;
    let hits: DeckHit[] = [];
    let next = 0;
    const run = (budgetMs: number): void => {
      if (mine !== token) return;
      const deadline = now() + budgetMs;
      const found: DeckHit[] = [];
      // Whole batches only; as many as the slot's budget allows.
      do {
        for (let index = next; index < Math.min(total, next + WALK_BATCH_SLIDES); index++) found.push(...findInSlide(model, index, compiled));
        next += WALK_BATCH_SLIDES;
      } while (next < total && now() < deadline);
      hits = [...hits, ...found];
      if (timing.firstHitsAt === undefined && hits.length > 0) timing.firstHitsAt = now();
      const done = next >= total;
      if (done) {
        slot = null;
        timing.settledAt = now();
        timing.firstHitsAt ??= timing.settledAt;
      } else {
        slot = schedule(run);
      }
      update(hits, done, { ...timing });
    };
    update(hits, false, { ...timing });
    slot = schedule(run);
  }

  return {
    start,
    cancel,
    get running() {
      return slot !== null;
    },
  };
}
