// The presenter view's elapsed-time timer (SPEC-100) — pure and clock-injected.
// A timer is two numbers: when the running stretch began (null while paused)
// and the time banked by earlier stretches. Every operation takes the caller's
// `nowMs` and returns a new state, so nothing here reads the wall clock and no
// test couples to it; the view ticks a display ref once a second and asks
// `elapsedMs` what to show.

export interface TimerState {
  /** When the current running stretch began, or null while paused. */
  startedAtMs: number | null;
  /** Time banked by earlier stretches (ms). */
  accumulatedMs: number;
}

/** A paused timer at zero. */
export const IDLE_TIMER: TimerState = { startedAtMs: null, accumulatedMs: 0 };

export function isTimerRunning(state: TimerState): boolean {
  return state.startedAtMs !== null;
}

/** Starts (or resumes) the timer at `nowMs`; a running timer is unchanged. */
export function startTimer(state: TimerState, nowMs: number): TimerState {
  return state.startedAtMs !== null ? state : { startedAtMs: nowMs, accumulatedMs: state.accumulatedMs };
}

/** Pauses at `nowMs`, banking the running stretch; a paused timer is unchanged. */
export function pauseTimer(state: TimerState, nowMs: number): TimerState {
  return state.startedAtMs === null ? state : { startedAtMs: null, accumulatedMs: elapsedMs(state, nowMs) };
}

/** Zeroes the timer, keeping whether it runs: a running timer restarts from
 * `nowMs`, a paused one stays paused at zero. */
export function resetTimer(state: TimerState, nowMs: number): TimerState {
  return { startedAtMs: state.startedAtMs === null ? null : nowMs, accumulatedMs: 0 };
}

/** The time shown at `nowMs`. Never negative — a clock that steps backwards
 * reads as no time passing, not as a negative timer. */
export function elapsedMs(state: TimerState, nowMs: number): number {
  const running = state.startedAtMs === null ? 0 : Math.max(0, nowMs - state.startedAtMs);
  return Math.max(0, state.accumulatedMs + running);
}

/** `M:SS` below one hour, `H:MM:SS` at or above it (whole seconds, floored;
 * a non-finite or negative input reads 0:00). */
export function formatElapsed(ms: number): string {
  const total = Number.isFinite(ms) && ms > 0 ? Math.floor(ms / 1000) : 0;
  const hours = Math.floor(total / 3600);
  const minutes = Math.floor((total % 3600) / 60);
  const seconds = String(total % 60).padStart(2, "0");
  return hours > 0 ? `${hours}:${String(minutes).padStart(2, "0")}:${seconds}` : `${minutes}:${seconds}`;
}
