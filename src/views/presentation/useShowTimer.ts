// The show timer (SPEC-146) — one per window that OWNS a show: the presenter
// window always, the full-screen play view when it has no ?follow= or once it
// is promoted. A following window builds one but never begins, adopts or
// changes it, so it never shows a timer and never touches the record.
//
//   · lib/presentation/presenterTimer.ts holds the pure state; this feeds it
//     the injected clock (`nowMs`, default lib/core/clock's wall clock, which
//     every window on the machine shares). The display ticks once a second,
//     and every operation and every mode or length change re-stamps it, so the
//     readout changes on the key or click itself.
//   · `begin()` runs once per mount, when the deck lands in an owning window.
//     It restores this tab's sessionStorage record for the deck exactly (a
//     running one keeps counting through the reload; a paused or stopped one
//     stays so, whatever auto-start says), and otherwise starts stopped and
//     then runs at nowMs when `presentationTimerAutoStart` is on. A promoted
//     window `adopt`s the presenter's timer instead.
//   · Mode and talk length are per-device preferences, never per show:
//     setMode/setLengthMinutes write the same keys the Preferences dialog
//     writes, and only the readout recomputes — the elapsed time is kept.
//   · The record (`showTimerRecordKey`) is written on every change and
//     removed on unmount, which every in-app departure causes; a reload or a
//     crashed tab never unmounts, so the record survives exactly those. Every
//     storage access is guarded: blocked storage just means no persistence.
//   · `status` is ShowTimer's live region: the transition a key or click
//     made, and "Time is up" once when a running countdown crosses zero —
//     never on begin, adopt, a restore or an ordinary tick.
// Nothing here dirties, edits or writes the deck.
import { computed, onBeforeUnmount, onMounted, ref, shallowRef, watch, type Ref } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import { now } from "../../lib/core/clock";
import {
  IDLE_TIMER,
  parseCountdownMinutes,
  parseTimerState,
  showTimerRecordKey,
  startTimer,
  stopTimer,
  timerPhase,
  timerReadout,
  toggleTimer,
  type TimerMode,
  type TimerPhase,
  type TimerState,
} from "../../lib/presentation/presenterTimer";

/** The timer display ticks this often. */
export const TIMER_TICK_MS = 1000;

/** The two session actions the timer's keys and buttons send. */
export type ShowTimerAction = "timerToggle" | "timerStop";

type TimerStorage = Pick<Storage, "getItem" | "setItem" | "removeItem">;

export interface ShowTimerDeps {
  deckId: Ref<string>;
  nowMs?: () => number; // default () => now().getTime()
  autoStart?: () => boolean; // default: presentationTimerAutoStart
  mode?: () => TimerMode; // default: presentationTimerMode
  lengthMinutes?: () => number; // default: presentationTimerCountdownMinutes
  storage?: TimerStorage | null; // default window.sessionStorage (guarded); null = none
  onChange?: (state: TimerState) => void; // after begin/toggle/stop/adopt — the session posts `timer`
}

/** The tab's sessionStorage, or null where reading it throws (blocked). */
function sessionStore(): TimerStorage | null {
  try {
    return typeof window === "undefined" ? null : window.sessionStorage;
  } catch {
    return null;
  }
}

/** The timer keys (t toggles; s and r stop), in both windows. */
export function timerKeyAction(key: string): ShowTimerAction | null {
  const lower = key.toLowerCase();
  if (lower === "t") return "timerToggle";
  return lower === "s" || lower === "r" ? "timerStop" : null;
}

/** The full-screen show's timer keys, as playKeyHandler's `extraKey`: sent
 * through the session's `perform` (a follower forwards them); not consumed on
 * the end screen, where they fall through to its any-key exit; and Enter or
 * Space on a focused ShowTimer button is left to that button alone. */
export function showTimerKeys(deps: { atEnd: Ref<boolean>; perform: (action: ShowTimerAction) => void; onKey?: () => void }) {
  return (event: KeyboardEvent): boolean => {
    if (deps.atEnd.value) return false; // the HUD is hidden: SPEC-084's any-key exit
    const target = event.target;
    if ((event.key === "Enter" || event.key === " ") && target instanceof HTMLElement && target.closest("[data-show-timer-control]")) {
      return true; // the focused button's own activation, and nothing else
    }
    const action = timerKeyAction(event.key);
    if (!action) return false;
    event.preventDefault();
    deps.perform(action);
    deps.onKey?.();
    return true;
  };
}

const ANNOUNCE: Record<TimerPhase, string> = {
  stopped: "Timer started",
  running: "Timer paused",
  paused: "Timer resumed",
};

export function useShowTimer(deps: ShowTimerDeps) {
  const nowMs = deps.nowMs ?? (() => now().getTime());
  const { preferences, setPreference } = usePreferences();
  const autoStart = deps.autoStart ?? (() => preferences.presentationTimerAutoStart);
  const storage = deps.storage === undefined ? sessionStore() : deps.storage;

  const timer = shallowRef<TimerState>(IDLE_TIMER);
  const tickMs = ref(nowMs());
  const status = ref("");
  const phase = computed(() => timerPhase(timer.value));
  const mode = computed<TimerMode>(() => (deps.mode ? deps.mode() : preferences.presentationTimerMode));
  const lengthMinutes = computed<number>(() =>
    deps.lengthMinutes ? deps.lengthMinutes() : preferences.presentationTimerCountdownMinutes,
  );
  const readout = computed(() => timerReadout(timer.value, tickMs.value, mode.value, lengthMinutes.value * 60_000));
  const text = computed(() => readout.value.text);
  const overtime = computed(() => readout.value.overtime);
  /** Everything ShowTimer renders, for `v-bind`. */
  const display = computed(() => ({
    phase: phase.value,
    mode: mode.value,
    text: text.value,
    overtime: overtime.value,
    status: status.value,
    lengthMinutes: lengthMinutes.value,
  }));

  // "Time is up": a running countdown crossing zero, by tick, length or mode
  // change. Synchronous, so begin/adopt can land in overtime quietly.
  let quiet = false;
  watch(
    overtime,
    (isOver, wasOver) => {
      if (isOver && !wasOver && !quiet && phase.value === "running") status.value = "Time is up";
    },
    { flush: "sync" },
  );

  // --- the record (owning windows only) -----------------------------------------
  /** The key this window wrote (null: it never owned the show). Kept, because
   * by unmount the route — and so deckId — may already be the next page's. */
  let ownedKey: string | null = null;
  function readRecord(): TimerState | null {
    try {
      const raw = storage?.getItem(showTimerRecordKey(deps.deckId.value));
      return typeof raw === "string" ? parseTimerState(JSON.parse(raw)) : null;
    } catch {
      return null; // blocked storage or a corrupt record: no record
    }
  }
  function commit(): void {
    ownedKey = showTimerRecordKey(deps.deckId.value);
    try {
      storage?.setItem(ownedKey, JSON.stringify(timer.value));
    } catch {
      // Blocked or full: the timer runs on without persistence.
    }
    deps.onChange?.(timer.value);
  }
  /** Re-stamps the tick and sets the state it yields, announcing nothing
   * (begin, adopt, a restore). */
  function land(next: (at: number) => TimerState): void {
    quiet = true;
    try {
      tickMs.value = nowMs();
      timer.value = next(tickMs.value);
    } finally {
      quiet = false;
    }
  }

  // --- operations -----------------------------------------------------------------
  let begun = false;
  function begin(): void {
    if (begun) return;
    begun = true;
    const restored = readRecord();
    land((at) => restored ?? (autoStart() ? startTimer(IDLE_TIMER, at) : stopTimer()));
    commit();
  }
  function adopt(state: TimerState): void {
    begun = true;
    land(() => ({ startedAtMs: state.startedAtMs, accumulatedMs: state.accumulatedMs }));
    commit();
  }
  function toggle(): void {
    const from = phase.value;
    tickMs.value = nowMs();
    timer.value = toggleTimer(timer.value, tickMs.value);
    status.value = ANNOUNCE[from];
    commit();
  }
  function stop(): void {
    if (phase.value === "stopped") return;
    tickMs.value = nowMs();
    timer.value = stopTimer();
    status.value = "Timer stopped";
    commit();
  }
  function setMode(next: TimerMode): void {
    setPreference("presentationTimerMode", next);
    tickMs.value = nowMs();
  }
  function setLengthMinutes(minutes: number): void {
    const value = parseCountdownMinutes(String(minutes)); // the field's own rounding and clamping
    if (value === null) return;
    setPreference("presentationTimerCountdownMinutes", value);
    tickMs.value = nowMs();
  }

  let tick = 0;
  onMounted(() => {
    tick = window.setInterval(() => {
      tickMs.value = nowMs();
    }, TIMER_TICK_MS);
  });
  onBeforeUnmount(() => {
    if (tick) clearInterval(tick);
    if (ownedKey === null) return; // a follower never touched the record
    try {
      storage?.removeItem(ownedKey);
    } catch {
      // Blocked: nothing was persisted.
    }
  });

  return {
    timer: timer as Readonly<Ref<TimerState>>,
    phase,
    mode,
    lengthMinutes,
    text,
    overtime,
    status: status as Readonly<Ref<string>>,
    display,
    begin,
    adopt,
    toggle,
    stop,
    setMode,
    setLengthMinutes,
  };
}

export type ShowTimer = ReturnType<typeof useShowTimer>;
