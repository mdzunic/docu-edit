// Frame selection for the slide stage (SPEC-081): a single click selects a
// frame (the stage draws a selection box; resize/move arrives in SPEC-082), a
// click outside any frame deselects, and Tab/Shift+Tab cycle selection through
// the current slide's frames in z-order. Editing is a separate concern
// (useSlideTextEditing); this composable only tracks WHICH frame is selected.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import type { Slide, SlideFrame } from "../../lib/presentation/model";

export interface FrameSelectionDeps {
  /** The slide whose frames are selectable (the stage's current slide). */
  currentSlide: ComputedRef<Slide | null>;
}

export interface FrameSelection {
  selectedFrameId: Ref<string | null>;
  selectedFrame: ComputedRef<SlideFrame | null>;
  select: (frameId: string | null) => void;
  clear: () => void;
  /** Cycle selection in z-order; `+1` forward (Tab), `-1` back (Shift+Tab). */
  cycle: (direction: 1 | -1) => void;
}

export function useFrameSelection(deps: FrameSelectionDeps): FrameSelection {
  const selectedFrameId = ref<string | null>(null);

  const selectedFrame = computed<SlideFrame | null>(() => {
    const slide = deps.currentSlide.value;
    if (!slide || selectedFrameId.value === null) return null;
    return slide.frames.find((frame) => frame.id === selectedFrameId.value) ?? null;
  });

  // A slide change drops any stale selection.
  watch(
    () => deps.currentSlide.value?.id,
    () => {
      selectedFrameId.value = null;
    },
  );

  function select(frameId: string | null): void {
    selectedFrameId.value = frameId;
  }

  function clear(): void {
    selectedFrameId.value = null;
  }

  function cycle(direction: 1 | -1): void {
    const slide = deps.currentSlide.value;
    if (!slide || slide.frames.length === 0) return;
    const ids = slide.frames.map((frame) => frame.id);
    const current = selectedFrameId.value === null ? -1 : ids.indexOf(selectedFrameId.value);
    const count = ids.length;
    const next = current === -1 ? (direction === 1 ? 0 : count - 1) : (current + direction + count) % count;
    selectedFrameId.value = ids[next];
  }

  return { selectedFrameId, selectedFrame, select, clear, cycle };
}
