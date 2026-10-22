// Frame selection for the slide stage (SPEC-081, multi-selection SPEC-096):
// a single click selects a frame, a click outside any frame deselects, and
// Tab/Shift+Tab cycle selection through the current slide's frames in
// z-order. SPEC-096 widens the selection to a SET — `selectedFrameIds` in
// z-order, top-level frames only — while `selectedFrameId` stays the
// single-frame Ref (the FIRST-selected anchor) so useAnimations' and
// useSlideTextEditing's deps are untouched. Editing is a separate concern
// (useSlideTextEditing); this composable only tracks WHICH frames are
// selected.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import type { Slide, SlideFrame } from "../../lib/presentation/model";

export interface FrameSelectionDeps {
  /** The slide whose frames are selectable (the stage's current slide). */
  currentSlide: ComputedRef<Slide | null>;
}

export interface FrameSelection {
  selectedFrameId: Ref<string | null>;
  selectedFrame: ComputedRef<SlideFrame | null>;
  /** Every selected frame id, z-order stable, top-level frames only. */
  selectedFrameIds: Ref<string[]>;
  /** The first-selected frame — what single-frame consumers anchor on. */
  anchorFrame: ComputedRef<SlideFrame | null>;
  select: (frameId: string | null) => void;
  /** Toggles one frame in or out of the selection (Shift+click). */
  toggle: (frameId: string) => void;
  /** Replaces (or, with `add`, extends) the selection with several frames. */
  selectMany: (frameIds: string[], add?: boolean) => void;
  /** Selects every frame on the slide (Cmd/Ctrl+A). */
  selectAll: () => void;
  isSelected: (frameId: string) => boolean;
  clear: () => void;
  /** Cycle selection in z-order; `+1` forward (Tab), `-1` back (Shift+Tab). */
  cycle: (direction: 1 | -1) => void;
}

export function useFrameSelection(deps: FrameSelectionDeps): FrameSelection {
  // The selection, kept z-order stable (slide order) — [0] is the anchor
  // frame single-frame consumers and the align "First selected" reference
  // read.
  const selectedFrameIds = ref<string[]>([]);

  // `selectedFrameId` keeps its SPEC-081 shape: a writable Ref<string|null>
  // holding selectedFrameIds[0], so existing single-frame consumers
  // (useAnimations' selectedFrameId/selectFrame, useSlideTextEditing) are
  // untouched. Writes collapse the whole set to that one frame.
  const selectedFrameId = computed<string | null>({
    get: () => selectedFrameIds.value[0] ?? null,
    set: (id) => {
      selectedFrameIds.value = id === null ? [] : [id];
    },
  });

  function frameById(id: string | null | undefined): SlideFrame | null {
    const slide = deps.currentSlide.value;
    if (!slide || id === null || id === undefined) return null;
    return slide.frames.find((frame) => frame.id === id) ?? null;
  }

  const selectedFrame = computed<SlideFrame | null>(() => frameById(selectedFrameId.value));
  const anchorFrame = computed<SlideFrame | null>(() => frameById(selectedFrameIds.value[0]));

  // A slide change drops any stale selection — the whole set.
  watch(
    () => deps.currentSlide.value?.id,
    () => {
      selectedFrameIds.value = [];
    },
  );

  function select(frameId: string | null): void {
    selectedFrameIds.value = frameId === null ? [] : [frameId];
  }

  /** Top-level slide order for `ids`, dropping ids the slide no longer has —
   * what keeps the set "z-order stable, top-level frames only". */
  function inZOrder(ids: readonly string[]): string[] {
    const slide = deps.currentSlide.value;
    if (!slide) return [];
    return slide.frames.filter((frame) => ids.includes(frame.id)).map((frame) => frame.id);
  }

  function toggle(frameId: string): void {
    if (selectedFrameIds.value.includes(frameId)) {
      selectedFrameIds.value = selectedFrameIds.value.filter((id) => id !== frameId);
    } else {
      selectedFrameIds.value = inZOrder([...selectedFrameIds.value, frameId]);
    }
  }

  function selectMany(frameIds: string[], add = false): void {
    const merged = add ? [...selectedFrameIds.value, ...frameIds] : frameIds;
    selectedFrameIds.value = inZOrder([...new Set(merged)]);
  }

  function selectAll(): void {
    const slide = deps.currentSlide.value;
    selectedFrameIds.value = slide ? slide.frames.map((frame) => frame.id) : [];
  }

  function isSelected(frameId: string): boolean {
    return selectedFrameIds.value.includes(frameId);
  }

  function clear(): void {
    selectedFrameIds.value = [];
  }

  function cycle(direction: 1 | -1): void {
    const slide = deps.currentSlide.value;
    if (!slide || slide.frames.length === 0) return;
    const ids = slide.frames.map((frame) => frame.id);
    const current = selectedFrameId.value === null ? -1 : ids.indexOf(selectedFrameId.value);
    const count = ids.length;
    const next = current === -1 ? (direction === 1 ? 0 : count - 1) : (current + direction + count) % count;
    selectedFrameIds.value = [ids[next]];
  }

  return {
    selectedFrameId,
    selectedFrame,
    selectedFrameIds,
    anchorFrame,
    select,
    toggle,
    selectMany,
    selectAll,
    isSelected,
    clear,
    cycle,
  };
}
