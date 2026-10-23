// The animations editing session (SPEC-086) — the presentation view's wiring
// for the Animations side panel: the persisted open/closed state (a
// `presentationAnimationsOpen` preference), the current slide's ordered effect
// list, the selected row and its field-edit commands (each one undo entry +
// persist), the confirm-to-replace flow for a slide with preserved out-of-
// subset timing, the on-stage numbered badges, and the "Play" preview state.
// The model mutations live in the framework-free lib (slideAnimations.ts); this
// composable sequences record → mutate → commit, mirrors the row selection onto
// the stage selection (AC-14), and holds the transient confirm/preview UI state.
//
// SPEC-105 grows the controller without changing its shape or the view's
// wiring: it creates and spreads in the build, repetition, trigger-frame,
// effect-option and motion-path commands (useAnimationOptions — each one undo
// entry in SPEC-080's voice through `patch`), the timeline view
// (useAnimationTimeline) and the motion-path editor (useMotionPathEdit), so
// PresentationView gains no wiring call.
import { computed, ref, shallowRef, type ComputedRef, type Ref } from "vue";
import type {
  AnimationEffect,
  AnimationTrigger,
  PresentationModel,
  Slide,
  SlideAnimation,
  TextBuildMode,
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
import { addAnimation, moveAnimation, removeAnimation, updateAnimation, type AnimationPatch } from "../../lib/presentation/slideAnimations";
import { frameAccessibleName } from "./useFrameManipulation";
import { useMotionPathEdit } from "./useMotionPathEdit";
import { isAnimatableKind, useAnimationOptions } from "./useAnimationOptions";
import { useAnimationTimeline } from "./useAnimationTimeline";
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

export interface AnimationRow {
  index: number;
  number: number;
  effect: AnimationEffect;
  trigger: AnimationTrigger;
  frameName: string;
  frameId: string;
  /** SPEC-105: the row's paragraph build, when it has one. */
  buildBy?: TextBuildMode;
  /** SPEC-105: the name of the frame an `onClickOf` row waits for. */
  triggerFrameName?: string;
}

export function useAnimations(deps: AnimationsDeps) {
  const { preferences, setPreference } = usePreferences();

  const open = computed(() => preferences.presentationAnimationsOpen);
  function toggle(): void {
    setPreference("presentationAnimationsOpen", !open.value);
  }

  // A fresh array each time the current slide re-identifies (commit bumps the
  // slide object), so the panel's derived computeds re-evaluate — the model is
  // plain data behind a shallowRef, so an in-place field edit is invisible
  // unless a dependency identity changes (the same trick the stage uses).
  const animations = computed<SlideAnimation[]>(() => (deps.currentSlide.value?.animations ?? []).slice());
  const selectedIndex = ref(0);
  /** The selected animation, read through the fresh array so field getters that
   * depend on `animations.value` re-evaluate after every edit. */
  function at(index: number): SlideAnimation | undefined {
    return animations.value[index];
  }

  function frameName(frameId: string): string {
    const frame = deps.currentSlide.value?.frames.find((f) => f.id === frameId);
    return frame ? frameAccessibleName(frame) : "Frame";
  }

  const rows = computed<AnimationRow[]>(() =>
    animations.value.map((a, index) => {
      const row: AnimationRow = {
        index,
        number: index + 1,
        effect: a.effect,
        trigger: effectiveTrigger(animations.value, index),
        frameName: frameName(a.frameId),
        frameId: a.frameId,
      };
      if (a.buildBy) row.buildBy = a.buildBy;
      if (a.triggerFrameId) row.triggerFrameName = frameName(a.triggerFrameId);
      return row;
    }),
  );

  /** The numbered badges drawn beside each animated frame while the panel is
   * open (AC-15). One per animation, in play order. */
  const badges = computed(() =>
    animations.value.map((a, index) => ({ index, number: index + 1, frameId: a.frameId })),
  );

  const selected = computed<SlideAnimation | null>(() => at(selectedIndex.value) ?? null);
  const hasSelection = computed(() => at(selectedIndex.value) !== undefined);
  const selectedEffect = computed<AnimationEffect | "">(() => at(selectedIndex.value)?.effect ?? "");
  const selectedTrigger = computed<AnimationTrigger>(() =>
    at(selectedIndex.value) ? effectiveTrigger(animations.value, selectedIndex.value) : "click",
  );
  const selectedDirection = computed<TransitionDirection>(() => at(selectedIndex.value)?.direction ?? "left");
  const directionEnabled = computed(() => {
    const a = at(selectedIndex.value);
    return !!a && usesAnimationDirection(a.effect);
  });
  // The first row of the MAIN sequence always behaves as click (AC-3) — its
  // with/after-previous choices are locked; it may still become "On click of"
  // (SPEC-105), which takes it out of the main sequence.
  const triggerLocked = computed(() => {
    const a = at(selectedIndex.value);
    return !!a && a.trigger !== "onClickOf" && animations.value.findIndex((x) => x.trigger !== "onClickOf") === selectedIndex.value;
  });
  const durationSeconds = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectiveAnimationDurationMs(a) / 1000 : 0.5;
  });
  const delaySeconds = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectiveAnimationDelayMs(a) / 1000 : 0;
  });

  // --- SPEC-105 field state and commands (useAnimationOptions) ----------------
  // The function declarations below (edit/patch) are hoisted, so the options
  // controller can take `patch` before they appear.
  const options = useAnimationOptions({ currentSlide: deps.currentSlide, animations, selectedIndex, patch });
  const { triggerFrameOptions } = options;

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
    if (triggerLocked.value && trigger !== "onClickOf" && trigger !== "click") return;
    if (trigger === "onClickOf") {
      // Waits for the first other frame until the user picks one.
      const triggerFrameId = at(selectedIndex.value)?.triggerFrameId ?? triggerFrameOptions.value[0]?.value;
      if (!triggerFrameId) return;
      edit("Change animation trigger", (slide, format) =>
        updateAnimation(slide, selectedIndex.value, { trigger, triggerFrameId }, format),
      );
      return;
    }
    edit("Change animation trigger", (slide, format) =>
      updateAnimation(slide, selectedIndex.value, { trigger }, format),
    );
  }
  /** One SPEC-105 field edit: one undo entry in SPEC-080's voice. */
  function patch(label: string, change: AnimationPatch): void {
    edit(label, (slide, format) => updateAnimation(slide, selectedIndex.value, change, format));
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

  // --- timeline (SPEC-105, useAnimationTimeline) ------------------------------
  const timelineView = useAnimationTimeline({ currentSlide: deps.currentSlide, animations, selectedIndex, selectRow, patch });

  // --- motion-path editor (SPEC-105) -------------------------------------------
  const pathEdit = useMotionPathEdit({
    model: deps.model,
    currentSlide: deps.currentSlide,
    selectedIndex,
    record: deps.record,
    commit: deps.commit,
  });

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
    ...options,
    ...timelineView,
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
    pathEdit,
    frameName,
    remove,
    moveUp,
    moveDown,
    previewState,
    play,
    onPreviewDone,
  };
}
