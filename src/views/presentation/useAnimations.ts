// The animations editing session (SPEC-086) — the presentation view's wiring
// for the Animations side panel: the persisted open/closed state (a
// `presentationAnimationsOpen` preference), the current slide's ordered effect
// list, the selected row and its field-edit commands (each one undo entry +
// persist), the confirm-to-replace flow for a slide with preserved out-of-
// subset timing, the on-stage numbered badges, and the "Play" preview state.
// The model mutations live in the framework-free lib (slideAnimations.ts); this
// composable sequences record → mutate → commit, mirrors the row selection onto
// the stage selection (AC-14), and holds the transient confirm/preview UI state.
import { computed, ref, shallowRef, type ComputedRef, type Ref } from "vue";
import type {
  AnimationEffect,
  AnimationTrigger,
  PresentationModel,
  Slide,
  SlideAnimation,
  TransitionDirection,
} from "../../lib/presentation/model";
import {
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  effectiveTrigger,
  hasPreservedTiming,
  MAX_ANIMATION_DELAY_MS,
  MAX_ANIMATION_DURATION_MS,
  MIN_ANIMATION_DELAY_MS,
  MIN_ANIMATION_DURATION_MS,
  usesAnimationDirection,
} from "../../lib/presentation/animations";
import { addAnimation, moveAnimation, removeAnimation, updateAnimation } from "../../lib/presentation/slideAnimations";
import { frameAccessibleName } from "./useFrameManipulation";
import { usePreferences } from "../../composables/usePreferences";

export interface AnimationsDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** The stage's selected frame id (SPEC-081) — the Add target, and what a row
   * selection highlights (AC-14). */
  selectedFrameId: Ref<string | null>;
  /** Sets the stage selection (used to highlight a row's target frame). */
  selectFrame: (frameId: string | null) => void;
}

/** The frame kinds an animation may target (AC out-of-scope: group/foreign
 * cannot be animated). */
function isAnimatableKind(kind: string): boolean {
  return kind === "text" || kind === "image" || kind === "shape";
}

export interface AnimationRow {
  index: number;
  number: number;
  effect: AnimationEffect;
  trigger: AnimationTrigger;
  frameName: string;
  frameId: string;
}

export function useAnimations(deps: AnimationsDeps) {
  const { preferences, setPreference } = usePreferences();

  const open = computed(() => preferences.presentationAnimationsOpen);
  function toggle(): void {
    setPreference("presentationAnimationsOpen", !open.value);
  }

  const animations = computed<SlideAnimation[]>(() => deps.currentSlide.value?.animations ?? []);
  const selectedIndex = ref(0);

  function frameName(frameId: string): string {
    const frame = deps.currentSlide.value?.frames.find((f) => f.id === frameId);
    return frame ? frameAccessibleName(frame) : "Frame";
  }

  const rows = computed<AnimationRow[]>(() =>
    animations.value.map((a, index) => ({
      index,
      number: index + 1,
      effect: a.effect,
      trigger: effectiveTrigger(animations.value, index),
      frameName: frameName(a.frameId),
      frameId: a.frameId,
    })),
  );

  /** The numbered badges drawn beside each animated frame while the panel is
   * open (AC-15). One per animation, in play order. */
  const badges = computed(() =>
    animations.value.map((a, index) => ({ index, number: index + 1, frameId: a.frameId })),
  );

  const selected = computed<SlideAnimation | null>(() => animations.value[selectedIndex.value] ?? null);
  const hasSelection = computed(() => selected.value !== null);
  const selectedEffect = computed<AnimationEffect | "">(() => selected.value?.effect ?? "");
  const selectedTrigger = computed<AnimationTrigger>(() =>
    selected.value ? effectiveTrigger(animations.value, selectedIndex.value) : "click",
  );
  const selectedDirection = computed<TransitionDirection>(() => selected.value?.direction ?? "left");
  const directionEnabled = computed(() => !!selected.value && usesAnimationDirection(selected.value.effect));
  // Row 1 always behaves as click (AC-3) — its trigger control is locked.
  const triggerLocked = computed(() => selectedIndex.value === 0);
  const durationSeconds = computed(() =>
    selected.value ? effectiveAnimationDurationMs(selected.value) / 1000 : 0.5,
  );
  const delaySeconds = computed(() => (selected.value ? effectiveAnimationDelayMs(selected.value) / 1000 : 0));

  const canMoveUp = computed(() => hasSelection.value && selectedIndex.value > 0);
  const canMoveDown = computed(() => hasSelection.value && selectedIndex.value < animations.value.length - 1);

  // --- Add target (AC-8) -----------------------------------------------------
  const targetFrame = computed(() => {
    const id = deps.selectedFrameId.value;
    return id ? (deps.currentSlide.value?.frames.find((f) => f.id === id) ?? null) : null;
  });
  const canAdd = computed(() => !!targetFrame.value && isAnimatableKind(targetFrame.value.kind));
  const addHelp = computed(() => {
    if (!targetFrame.value) return "Select a frame on the slide to animate it.";
    if (!isAnimatableKind(targetFrame.value.kind)) return "This kind of object can't be animated.";
    return "";
  });

  function selectRow(index: number): void {
    if (index < 0 || index >= animations.value.length) return;
    selectedIndex.value = index;
    // Highlight the row's target frame on the stage (AC-14).
    deps.selectFrame(animations.value[index].frameId);
  }

  // --- confirm-to-replace preserved timing (AC-17) ---------------------------
  const pendingReplaceFrameId = ref<string | null>(null);

  function add(): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model || !canAdd.value || !targetFrame.value) return;
    if (hasPreservedTiming(slide) && !slide.animations) {
      pendingReplaceFrameId.value = targetFrame.value.id;
      return;
    }
    doAdd(targetFrame.value.id);
  }
  function doAdd(frameId: string): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model) return;
    deps.record("Add animation");
    const index = addAnimation(slide, frameId, model.format);
    deps.commit();
    selectedIndex.value = index;
  }
  function confirmReplace(): void {
    const frameId = pendingReplaceFrameId.value;
    pendingReplaceFrameId.value = null;
    if (frameId) doAdd(frameId);
  }
  function cancelReplace(): void {
    pendingReplaceFrameId.value = null;
  }

  // --- field commands --------------------------------------------------------
  function edit(label: string, mutate: (slide: Slide, format: PresentationModel["format"]) => void): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model || !selected.value) return;
    deps.record(label);
    mutate(slide, model.format);
    deps.commit();
  }
  function setEffect(effect: AnimationEffect): void {
    edit("Change animation effect", (slide, format) => updateAnimation(slide, selectedIndex.value, { effect }, format));
  }
  function setDirection(direction: TransitionDirection): void {
    edit("Change animation direction", (slide, format) =>
      updateAnimation(slide, selectedIndex.value, { direction }, format),
    );
  }
  function setTrigger(trigger: AnimationTrigger): void {
    if (triggerLocked.value) return;
    edit("Change animation trigger", (slide, format) =>
      updateAnimation(slide, selectedIndex.value, { trigger }, format),
    );
  }
  function setDurationSeconds(seconds: number): void {
    const ms = Math.min(MAX_ANIMATION_DURATION_MS, Math.max(MIN_ANIMATION_DURATION_MS, Math.round(seconds * 1000)));
    edit("Change animation duration", (slide, format) =>
      updateAnimation(slide, selectedIndex.value, { durationMs: ms }, format),
    );
  }
  function setDelaySeconds(seconds: number): void {
    const ms = Math.min(MAX_ANIMATION_DELAY_MS, Math.max(MIN_ANIMATION_DELAY_MS, Math.round(seconds * 1000)));
    edit("Change animation delay", (slide, format) =>
      updateAnimation(slide, selectedIndex.value, { delayMs: ms }, format),
    );
  }
  function remove(): void {
    edit("Remove animation", (slide, format) => {
      selectedIndex.value = Math.max(0, removeAnimation(slide, selectedIndex.value, format));
    });
  }
  function moveUp(): void {
    if (!canMoveUp.value) return;
    edit("Reorder animation", (slide, format) => {
      selectedIndex.value = moveAnimation(slide, selectedIndex.value, -1, format);
    });
  }
  function moveDown(): void {
    if (!canMoveDown.value) return;
    edit("Reorder animation", (slide, format) => {
      selectedIndex.value = moveAnimation(slide, selectedIndex.value, 1, format);
    });
  }

  // --- Play preview (AC-11) --------------------------------------------------
  const previewState = shallowRef<Slide | null>(null);
  function play(): void {
    const slide = deps.currentSlide.value;
    if (slide && (slide.animations?.length ?? 0) > 0) {
      previewState.value = slide;
    }
  }
  function onPreviewDone(): void {
    previewState.value = null;
  }

  return {
    open,
    toggle,
    animations,
    rows,
    badges,
    selectedIndex,
    selected,
    hasSelection,
    selectedEffect,
    selectedTrigger,
    selectedDirection,
    directionEnabled,
    triggerLocked,
    durationSeconds,
    delaySeconds,
    canMoveUp,
    canMoveDown,
    canAdd,
    addHelp,
    pendingReplaceFrameId,
    selectRow,
    add,
    confirmReplace,
    cancelReplace,
    setEffect,
    setDirection,
    setTrigger,
    setDurationSeconds,
    setDelaySeconds,
    remove,
    moveUp,
    moveDown,
    previewState,
    play,
    onPreviewDone,
  };
}
