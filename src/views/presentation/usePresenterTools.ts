// The presenter view's two side tools (SPEC-100), split out of
// usePresenterView: the elapsed-time timer and the speaker-notes scale.
//
//   · The timer is lib/presentation/presenterTimer.ts's pure state, fed the
//     injected clock (`nowMs`) — nothing here reads Date.now(). It starts with
//     the show (`start`), the display ref ticks once a second while mounted,
//     and toggle/reset re-stamp it so the readout changes on the click itself.
//   · The notes scale steps through PRESENTER_NOTES_SCALES and is persisted per
//     device as the `presentationPresenterNotesScale` preference.
import { computed, onBeforeUnmount, onMounted, ref } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import { PRESENTER_NOTES_SCALES } from "../../lib/preferences";
import {
  elapsedMs,
  formatElapsed,
  IDLE_TIMER,
  isTimerRunning,
  pauseTimer,
  resetTimer,
  startTimer,
} from "../../lib/presentation/presenterTimer";

/** The timer display ticks this often. */
export const TIMER_TICK_MS = 1000;

export function usePresenterTools(nowMs: () => number) {
  // --- the timer ---------------------------------------------------------------
  const timer = ref(IDLE_TIMER);
  const tickMs = ref(nowMs());
  let tick = 0;
  const timerText = computed(() => formatElapsed(elapsedMs(timer.value, tickMs.value)));
  const timerPaused = computed(() => !isTimerRunning(timer.value));
  /** Starts the timer from zero (the show began). */
  function start(): void {
    tickMs.value = nowMs();
    timer.value = startTimer(IDLE_TIMER, tickMs.value);
  }
  function toggleTimer(): void {
    tickMs.value = nowMs();
    timer.value = isTimerRunning(timer.value) ? pauseTimer(timer.value, tickMs.value) : startTimer(timer.value, tickMs.value);
  }
  function resetTimerNow(): void {
    tickMs.value = nowMs();
    timer.value = resetTimer(timer.value, tickMs.value);
  }
  onMounted(() => {
    tick = window.setInterval(() => {
      tickMs.value = nowMs();
    }, TIMER_TICK_MS);
  });
  onBeforeUnmount(() => {
    if (tick) clearInterval(tick);
  });

  // --- notes scale (per device) ------------------------------------------------
  const { preferences, setPreference } = usePreferences();
  const notesScale = computed(() => preferences.presentationPresenterNotesScale);
  const scaleIndex = computed(() => Math.max(0, PRESENTER_NOTES_SCALES.indexOf(notesScale.value)));
  function stepNotesScale(delta: number): void {
    const index = Math.min(Math.max(scaleIndex.value + delta, 0), PRESENTER_NOTES_SCALES.length - 1);
    setPreference("presentationPresenterNotesScale", PRESENTER_NOTES_SCALES[index]);
  }

  return {
    startTimer: start,
    timerText,
    timerPaused,
    toggleTimer,
    resetTimer: resetTimerNow,
    notesScale,
    canShrinkNotes: computed(() => scaleIndex.value > 0),
    canGrowNotes: computed(() => scaleIndex.value < PRESENTER_NOTES_SCALES.length - 1),
    stepNotesScale,
  };
}
