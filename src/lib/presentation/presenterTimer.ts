// The presenter view's elapsed-time timer (SPEC-100) — pure and clock-injected.
// A timer is two numbers: when the running stretch began (null while paused)
// and the time banked by earlier stretches. Every operation takes the caller's
// `nowMs` and returns a new state, so nothing here reads the wall clock and no
// test couples to it; the view ticks a display ref once a second and asks
// `elapsedMs` what to show.
//
// SPEC-146 makes it the SHOW timer (the normative source is that spec): three
// phases on the same two numbers, a Stopwatch or Countdown reading of them
// (`timerReadout`), the talk-length vocabulary the preferences validate with,
// and the validator every TimerState from another window or from
// sessionStorage passes through (`parseTimerState`).

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

// --- SPEC-146: the show timer's phases, modes and readout --------------------
// One TimerState always measures elapsed time; the mode only changes what is
// displayed, so switching mode or changing the talk length mid-show never
// resets or loses the count.

export type TimerPhase = "stopped" | "running" | "paused";

/** running ⇔ startedAtMs !== null; stopped ⇔ not running and accumulatedMs === 0; else paused. */
export function timerPhase(state: TimerState): TimerPhase {
  if (state.startedAtMs !== null) return "running";
  return state.accumulatedMs === 0 ? "stopped" : "paused";
}

/** stopped → running from 0 at nowMs; running → paused (banked); paused → running. */
export function toggleTimer(state: TimerState, nowMs: number): TimerState {
  const phase = timerPhase(state);
  if (phase === "stopped") return startTimer(IDLE_TIMER, nowMs);
  return phase === "running" ? pauseTimer(state, nowMs) : startTimer(state, nowMs);
}

/** Any phase → IDLE_TIMER (a fresh object). */
export function stopTimer(): TimerState {
  return { ...IDLE_TIMER };
}

export type TimerMode = "stopwatch" | "countdown";
export const TIMER_MODES: readonly TimerMode[] = ["stopwatch", "countdown"];
export const COUNTDOWN_MINUTES_DEFAULT = 20;
export const COUNTDOWN_MINUTES_MIN = 1;
export const COUNTDOWN_MINUTES_MAX = 240;

/** A whole number of minutes in [MIN, MAX] (the preference validator). */
export function isCountdownMinutes(raw: unknown): raw is number {
  return typeof raw === "number" && Number.isInteger(raw) && raw >= COUNTDOWN_MINUTES_MIN && raw <= COUNTDOWN_MINUTES_MAX;
}

/** What a person typed → whole minutes, rounded and clamped into [MIN, MAX];
 * "", whitespace, or anything Number() rejects → null (the field restores). */
export function parseCountdownMinutes(text: string): number | null {
  const trimmed = text.trim();
  const value = trimmed === "" ? Number.NaN : Number(trimmed);
  if (Number.isNaN(value)) return null;
  return Math.min(COUNTDOWN_MINUTES_MAX, Math.max(COUNTDOWN_MINUTES_MIN, Math.round(value)));
}

export interface TimerReadout {
  text: string;
  overtime: boolean;
}

/** The minus sign overtime reads with (U+2212, not a hyphen). */
const MINUS = "−";

/** What the timer shows at nowMs. stopwatch: formatElapsed(elapsed), never
 * overtime. countdown: r = lengthMs − elapsed; s = Math.ceil(r / 1000);
 * s > 0 → formatElapsed(s·1000); s ≤ 0 → "0:00" when s is 0 (or −0), else
 * "−" (U+2212) + formatElapsed(−s·1000); overtime ⇔ r ≤ 0. */
export function timerReadout(state: TimerState, nowMs: number, mode: TimerMode, lengthMs: number): TimerReadout {
  const elapsed = elapsedMs(state, nowMs);
  if (mode === "stopwatch") return { text: formatElapsed(elapsed), overtime: false };
  const remaining = lengthMs - elapsed;
  const seconds = Math.ceil(remaining / 1000);
  if (seconds > 0) return { text: formatElapsed(seconds * 1000), overtime: false };
  return { text: seconds === 0 ? "0:00" : MINUS + formatElapsed(-seconds * 1000), overtime: true };
}

function isNonNegativeNumber(raw: unknown): raw is number {
  return typeof raw === "number" && Number.isFinite(raw) && raw >= 0;
}

/** Validates a TimerState from another window or from storage: startedAtMs is
 * null or a finite number ≥ 0; accumulatedMs a finite number ≥ 0; extra fields
 * dropped; anything else → null. Never throws. */
export function parseTimerState(raw: unknown): TimerState | null {
  try {
    if (raw === null || typeof raw !== "object" || Array.isArray(raw)) return null;
    const { startedAtMs, accumulatedMs } = raw as Record<string, unknown>;
    if (!isNonNegativeNumber(accumulatedMs)) return null;
    if (startedAtMs !== null && !isNonNegativeNumber(startedAtMs)) return null;
    return { startedAtMs, accumulatedMs };
  } catch {
    return null; // a hostile getter
  }
}

/** The sessionStorage key an owning window keeps its timer under: "docu-edit:show-timer:<deckId>". */
export function showTimerRecordKey(deckId: string): string {
  return `docu-edit:show-timer:${deckId}`;
}
