// The presenter view's side tool (SPEC-100), split out of usePresenterView:
// the speaker-notes scale. It steps through PRESENTER_NOTES_SCALES and is
// persisted per device as the `presentationPresenterNotesScale` preference.
// The timer half moved out in SPEC-146: it is the show timer now
// (useShowTimer), which both presenting windows run.
import { computed } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import { PRESENTER_NOTES_SCALES } from "../../lib/preferences";

export function usePresenterTools() {
  const { preferences, setPreference } = usePreferences();
  const notesScale = computed(() => preferences.presentationPresenterNotesScale);
  const scaleIndex = computed(() => Math.max(0, PRESENTER_NOTES_SCALES.indexOf(notesScale.value)));
  function stepNotesScale(delta: number): void {
    const index = Math.min(Math.max(scaleIndex.value + delta, 0), PRESENTER_NOTES_SCALES.length - 1);
    setPreference("presentationPresenterNotesScale", PRESENTER_NOTES_SCALES[index]);
  }

  return {
    notesScale,
    canShrinkNotes: computed(() => scaleIndex.value > 0),
    canGrowNotes: computed(() => scaleIndex.value < PRESENTER_NOTES_SCALES.length - 1),
    stepNotesScale,
  };
}
