// The speaker-notes editing session (SPEC-083) — the presentation view's thin
// wiring for the notes panel below the stage: the persisted open/closed state
// (a `presentationNotesOpen` preference via SPEC-066), the Cmd/Ctrl+Shift+H
// toggle, and per-input model edits that coalesce into undo bursts exactly like
// SPEC-081 typing. A notes edit marks the slide dirty; for `.pptx` it sets
// `notesDirty` so only the notes-slide part rebuilds (the slide part stays
// byte-verbatim), for `.odp` it rebuilds the whole page (there is no separate
// notes part), which `markSlideEdited` arranges by dropping `sourceXml`.
import { computed, watch, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { markSlideEdited } from "../../lib/presentation/text";
import { usePreferences } from "../../composables/usePreferences";

/** How long a pause closes a notes typing burst (SPEC-081 parity: ≥ 1s). */
export const NOTES_BURST_MS = 1000;

export interface NotesDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  /** The current slide index — a change commits the open burst (AC). */
  currentIndex: Ref<number>;
  /** Snapshot pre-edit state for undo (usePresentationUndo.record). */
  record: (label?: string) => void;
  /** Re-render the shallowRef + persist write-through (PresentationView). */
  commit: () => void;
}

export function useNotes(deps: NotesDeps) {
  const { preferences, setPreference } = usePreferences();

  /** Whether the panel is expanded — persisted per device (AC), default off. */
  const open = computed(() => preferences.presentationNotesOpen);
  function toggle(): void {
    closeBurst();
    setPreference("presentationNotesOpen", !open.value);
  }

  // --- undo burst (SPEC-081 parity) -----------------------------------------
  let burstActive = false;
  let burstTimer = 0;
  function noteTyping(): void {
    if (!burstActive) {
      deps.record("Notes");
      burstActive = true;
    }
    if (burstTimer) {
      clearTimeout(burstTimer);
    }
    burstTimer = window.setTimeout(closeBurst, NOTES_BURST_MS);
  }
  function closeBurst(): void {
    burstActive = false;
    if (burstTimer) {
      clearTimeout(burstTimer);
      burstTimer = 0;
    }
  }

  /** The current slide's notes as newline-joined text for the textarea. Split
   * then join by "\n" is an identity, so the value never fights the caret. */
  const notesText = computed(() => (deps.currentSlide.value?.notes ?? []).join("\n"));

  /** Applies a textarea edit to the model (AC: per input event). Empty text
   * removes the field (`[]` is never stored); otherwise the paragraphs are the
   * newline-split lines. */
  function onInput(value: string): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model) {
      return;
    }
    noteTyping();
    const paragraphs = value.split("\n");
    if (paragraphs.every((line) => line.length === 0)) {
      delete slide.notes;
    } else {
      slide.notes = paragraphs;
    }
    if (model.format === "odp") {
      // The whole page rebuilds; markSlideEdited drops sourceXml so the page is
      // regenerated (with the new notes) rather than re-emitted verbatim.
      markSlideEdited(slide, "odp");
    } else {
      slide.notesDirty = true;
    }
    deps.commit();
  }

  // Switching slides commits the current burst (AC).
  watch(deps.currentIndex, () => closeBurst());

  /** Global Cmd/Ctrl+Shift+H toggles the panel (AC). */
  function onGlobalKeydown(event: KeyboardEvent): void {
    if ((event.metaKey || event.ctrlKey) && event.shiftKey && !event.altKey && event.key.toLowerCase() === "h") {
      event.preventDefault();
      toggle();
    }
  }

  return { open, toggle, notesText, onInput, onGlobalKeydown, closeBurst };
}
