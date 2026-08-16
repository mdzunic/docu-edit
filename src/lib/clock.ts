// The one injectable source of "now" for field evaluation (SPEC-053
// "Determinism"). Every field whose value depends on the wall clock — `date`
// and `time` — reads the current instant through `now()` and nothing calls
// `Date.now()` or `new Date()` directly. `cached` field results are compared
// by equality.ts's `equalStructural` (they are document content Word and
// Writer display without recalculating), so a field evaluated from the real
// wall clock would make every round-trip test that touches a date field flake
// whenever the suite runs across midnight. Unit tests substitute a fixed
// instant via `setClock`; e2e tests set it before loading the page.

/** The override installed by tests, or `null` to read the real clock. */
let fixedInstant: Date | null = null;

/** The current instant. The single seam the field evaluator reads the clock
 * through — see the module comment. */
export function now(): Date {
  return fixedInstant !== null ? new Date(fixedInstant.getTime()) : new Date();
}

/** Pins `now()` to a fixed instant (unit and e2e tests). Passing `null`
 * restores the real clock. */
export function setClock(instant: Date | null): void {
  fixedInstant = instant === null ? null : new Date(instant.getTime());
}
