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
// SPEC-105 grows it without changing its shape: the build, repetition,
// trigger-frame, effect-option and motion-path commands (each one undo entry
// in SPEC-080's voice), the timeline view (`presentationAnimationTimelineOpen`,
// its geometry from components/presentation/animationTimeline.ts, an edge drag
// recording once on pointer-up) and the motion-path editor
// (useMotionPathEdit, created here so PresentationView gains no wiring).
import { computed, ref, shallowRef, type ComputedRef, type Ref } from "vue";
import type {
  AnimationEffect,
  AnimationTrigger,
  BuildStepTrigger,
  MotionPath,
  PresentationModel,
  Slide,
  SlideAnimation,
  SlideFrame,
  TextBuildMode,
  TransitionDirection,
} from "../../lib/presentation/model";
import {
  DEFAULT_EMPHASIS_COLOR,
  DEFAULT_PATH_DISTANCE,
  DEFAULT_SCALE_PERCENT,
  DEFAULT_SPIN_DEGREES,
  effectClass,
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  effectiveEmphasisColor,
  effectiveRepeatCount,
  effectiveScalePercent,
  effectiveSpinDegrees,
  effectiveTrigger,
  hasPreservedTiming,
  MAX_ANIMATION_DELAY_MS,
  MAX_ANIMATION_DURATION_MS,
  MIN_ANIMATION_DELAY_MS,
  MIN_ANIMATION_DURATION_MS,
  usesAnimationDirection,
} from "../../lib/presentation/animations";
import { addAnimation, moveAnimation, removeAnimation, updateAnimation, type AnimationPatch } from "../../lib/presentation/slideAnimations";
import { buildableParagraphs, canBuildByParagraph } from "../../lib/presentation/buildPlayback";
import {
  draggedDelayMs,
  draggedDurationMs,
  msFromDx,
  timelineModel,
  TIMELINE_STEP_MS,
  type TimelineBar,
} from "../../components/presentation/animationTimeline";
import { frameAccessibleName } from "./useFrameManipulation";
import { useMotionPathEdit } from "./useMotionPathEdit";
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
 * cannot be animated). SPEC-095 adds media — a clip can enter/exit like any
 * other frame (its own playback duration is never a build step). */
function isAnimatableKind(kind: string): boolean {
  return (
    kind === "text" || kind === "image" || kind === "shape" || kind === "table" || kind === "media"
  );
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

/** An edge of a timeline bar: its left (the delay) or right (the duration). */
export type TimelineEdge = "delay" | "duration";

/** The delay/duration an edge edit of `deltaMs` leaves (clamped, 100ms steps). */
function edgeValue(bar: TimelineBar, edge: TimelineEdge, deltaMs: number): number {
  return edge === "delay" ? draggedDelayMs(bar.delayMs, deltaMs) : draggedDurationMs(bar.singleMs, deltaMs, bar.passes);
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

  // --- SPEC-105 field state -------------------------------------------------
  const selectedFrame = computed<SlideFrame | null>(() => {
    const id = at(selectedIndex.value)?.frameId;
    return id ? (deps.currentSlide.value?.frames.find((f) => f.id === id) ?? null) : null;
  });
  const selectedClass = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectClass(a.effect) : "entrance";
  });
  const selectedBuildBy = computed<TextBuildMode | "whole">(() => at(selectedIndex.value)?.buildBy ?? "whole");
  /** A frame with fewer than two paragraphs (or no text) builds whole (AC-9). */
  const buildByEnabled = computed(() => canBuildByParagraph(selectedFrame.value ?? undefined));
  const buildByHelp = computed(() => {
    if (buildByEnabled.value) return "";
    return buildableParagraphs(selectedFrame.value ?? undefined)
      ? "Building by paragraph needs two or more paragraphs."
      : "Only text builds by paragraph.";
  });
  const buildTriggerShown = computed(() => !!at(selectedIndex.value)?.buildBy && buildByEnabled.value);
  const selectedBuildTrigger = computed<BuildStepTrigger>(() => at(selectedIndex.value)?.buildTrigger ?? "click");
  /** Repetition and auto-reverse belong to emphasis and path effects (and
   * stay editable on any row a file gave them). */
  const repeatShown = computed(() => {
    const a = at(selectedIndex.value);
    return !!a && (selectedClass.value === "emphasis" || selectedClass.value === "path" || a.repeatCount !== undefined || !!a.repeatUntilClick || !!a.autoReverse);
  });
  const selectedRepeatCount = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectiveRepeatCount(a) : 1;
  });
  const selectedRepeatUntilClick = computed(() => !!at(selectedIndex.value)?.repeatUntilClick);
  const selectedAutoReverse = computed(() => !!at(selectedIndex.value)?.autoReverse);
  /** Rewind applies to every effect but an instant appear. */
  const rewindShown = computed(() => {
    const a = at(selectedIndex.value);
    return !!a && (a.effect !== "appear" || !!a.rewindOnComplete);
  });
  const selectedRewind = computed(() => !!at(selectedIndex.value)?.rewindOnComplete);
  /** The slide's other animatable frames an "On click of" row may wait for. */
  const triggerFrameOptions = computed(() => {
    const own = at(selectedIndex.value)?.frameId;
    const frames = (deps.currentSlide.value?.frames ?? []).filter((f) => f.id !== own && isAnimatableKind(f.kind));
    const names = frames.map(frameAccessibleName);
    return frames.map((f, i) => ({
      value: f.id,
      label: names.filter((n) => n === names[i]).length > 1 ? `${names[i]} ${names.slice(0, i + 1).filter((n) => n === names[i]).length}` : names[i],
    }));
  });
  const triggerFrameEnabled = computed(() => at(selectedIndex.value)?.trigger === "onClickOf");
  const selectedTriggerFrameId = computed(() => at(selectedIndex.value)?.triggerFrameId ?? "");
  const selectedSpinDegrees = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectiveSpinDegrees(a) : DEFAULT_SPIN_DEGREES;
  });
  const selectedScalePercent = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectiveScalePercent(a) : DEFAULT_SCALE_PERCENT;
  });
  const selectedEmphasisColor = computed(() => {
    const a = at(selectedIndex.value);
    return a ? effectiveEmphasisColor(a) : DEFAULT_EMPHASIS_COLOR;
  });
  const selectedPath = computed<MotionPath | null>(() => at(selectedIndex.value)?.path ?? null);
  const selectedPathPreset = computed(() => selectedPath.value?.preset ?? "line");
  const pathDirectionShown = computed(() => ["line", "arc", "turn"].includes(selectedPathPreset.value));
  const selectedPathDirection = computed<TransitionDirection>(() => selectedPath.value?.direction ?? "right");
  const selectedPathDistancePercent = computed(() =>
    Math.round((selectedPath.value?.distanceFraction ?? DEFAULT_PATH_DISTANCE) * 100),
  );

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
  function setBuildBy(mode: TextBuildMode | "whole"): void {
    if (mode !== "whole" && !buildByEnabled.value) return;
    patch("Change animation build", mode === "whole" ? { clearBuildBy: true } : { buildBy: mode });
  }
  function setBuildTrigger(trigger: BuildStepTrigger): void {
    patch("Change animation build", { buildTrigger: trigger });
  }
  function setRepeatCount(count: number): void {
    if (!Number.isFinite(count)) return;
    patch("Change animation repeat", count < 2 ? { clearRepeat: true } : { repeatCount: count });
  }
  function setRepeatUntilClick(on: boolean): void {
    patch("Change animation repeat", { repeatUntilClick: on });
  }
  function setAutoReverse(on: boolean): void {
    patch("Change animation auto-reverse", { autoReverse: on });
  }
  function setRewind(on: boolean): void {
    patch("Change animation rewind", { rewindOnComplete: on });
  }
  function setTriggerFrame(frameId: string): void {
    if (!triggerFrameEnabled.value) return;
    patch("Change animation trigger frame", { triggerFrameId: frameId });
  }
  function setSpinDegrees(degrees: number): void {
    if (!Number.isFinite(degrees)) return;
    patch("Change animation spin", { spinDegrees: degrees });
  }
  function setScalePercent(percent: number): void {
    if (!Number.isFinite(percent)) return;
    patch("Change animation scale", { scalePercent: percent });
  }
  function setEmphasisColor(color: string): void {
    patch("Change animation colour", { emphasisColor: color });
  }
  /** Choosing a preset replaces the path (an imported custom one included —
   * the only way it changes); direction and distance carry over. */
  function setPathPreset(preset: MotionPath["preset"]): void {
    const current = selectedPath.value;
    if (preset === "custom" || !current || preset === current.preset) return;
    const next: MotionPath = { preset };
    if (current.direction && preset !== "loop") next.direction = current.direction;
    if (current.distanceFraction !== undefined) next.distanceFraction = current.distanceFraction;
    patch("Change motion path", { path: next });
  }
  function setPathDirection(direction: TransitionDirection): void {
    const current = selectedPath.value;
    if (!current || current.preset === "custom") return;
    patch("Change motion path", { path: { ...current, direction } });
  }
  function setPathDistancePercent(percent: number): void {
    const current = selectedPath.value;
    if (!current || current.preset === "custom" || !Number.isFinite(percent)) return;
    patch("Change motion path", { path: { ...current, distanceFraction: Math.round(percent) / 100 } });
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

  // --- timeline (SPEC-105) -----------------------------------------------------
  const timelineOpen = computed(() => preferences.presentationAnimationTimelineOpen);
  function toggleTimeline(): void {
    setPreference("presentationAnimationTimelineOpen", !timelineOpen.value);
  }
  /** A live edge drag: the row and the value it would write on pointer-up. */
  const timelineDraft = ref<{ animationIndex: number; edge: TimelineEdge; valueMs: number } | null>(null);
  let edgeDrag: { bar: TimelineBar; edge: TimelineEdge; startX: number } | null = null;
  /** The timeline of the current slide, a live drag's value applied. */
  const timeline = computed(() => {
    const slide = deps.currentSlide.value;
    if (!slide) return null;
    const draftValue = timelineDraft.value;
    const shown = draftValue
      ? {
          ...slide,
          animations: animations.value.map((a, i) =>
            i === draftValue.animationIndex
              ? { ...a, [draftValue.edge === "delay" ? "delayMs" : "durationMs"]: draftValue.valueMs }
              : a,
          ),
        }
      : { ...slide, animations: animations.value };
    return timelineModel(shown);
  });
  function writeEdge(animationIndex: number, edge: TimelineEdge, valueMs: number): void {
    const a = at(animationIndex);
    if (!a) return;
    const current = edge === "delay" ? effectiveAnimationDelayMs(a) : effectiveAnimationDurationMs(a);
    if (valueMs === current) return;
    selectedIndex.value = animationIndex;
    patch(
      edge === "delay" ? "Change animation delay" : "Change animation duration",
      edge === "delay" ? { delayMs: valueMs } : { durationMs: valueMs },
    );
  }
  /** A bar drives the same selection as its row. */
  function selectBar(bar: TimelineBar): void {
    selectRow(bar.animationIndex);
  }
  function startEdgeDrag(bar: TimelineBar, edge: TimelineEdge, event: PointerEvent): void {
    selectBar(bar);
    const target = event.target;
    if (target instanceof Element && typeof target.setPointerCapture === "function") {
      try {
        target.setPointerCapture(event.pointerId);
      } catch {
        // A synthetic pointer has nothing to capture.
      }
    }
    edgeDrag = { bar, edge, startX: event.clientX };
    timelineDraft.value = { animationIndex: bar.animationIndex, edge, valueMs: edgeValue(bar, edge, 0) };
  }
  function moveEdgeDrag(event: PointerEvent): void {
    if (!edgeDrag) return;
    const { bar, edge, startX } = edgeDrag;
    timelineDraft.value = { animationIndex: bar.animationIndex, edge, valueMs: edgeValue(bar, edge, msFromDx(event.clientX - startX)) };
  }
  /** Pointer-up ends the drag: one undo entry for the whole gesture. */
  function endEdgeDrag(): void {
    const draftValue = timelineDraft.value;
    edgeDrag = null;
    timelineDraft.value = null;
    if (draftValue) writeEdge(draftValue.animationIndex, draftValue.edge, draftValue.valueMs);
  }
  /** Left/Right on an edge: ∓100ms, one undo entry per key press. */
  function edgeKeydown(bar: TimelineBar, edge: TimelineEdge, event: KeyboardEvent): void {
    const delta = event.key === "ArrowLeft" ? -TIMELINE_STEP_MS : event.key === "ArrowRight" ? TIMELINE_STEP_MS : 0;
    if (delta === 0) return;
    event.preventDefault();
    writeEdge(bar.animationIndex, edge, edgeValue(bar, edge, edge === "duration" ? delta * bar.passes : delta));
  }

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
    selectedFrame,
    selectedClass,
    selectedBuildBy,
    buildByEnabled,
    buildByHelp,
    buildTriggerShown,
    selectedBuildTrigger,
    repeatShown,
    selectedRepeatCount,
    selectedRepeatUntilClick,
    selectedAutoReverse,
    rewindShown,
    selectedRewind,
    triggerFrameOptions,
    triggerFrameEnabled,
    selectedTriggerFrameId,
    selectedSpinDegrees,
    selectedScalePercent,
    selectedEmphasisColor,
    selectedPath,
    selectedPathPreset,
    pathDirectionShown,
    selectedPathDirection,
    selectedPathDistancePercent,
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
    setBuildBy,
    setBuildTrigger,
    setRepeatCount,
    setRepeatUntilClick,
    setAutoReverse,
    setRewind,
    setTriggerFrame,
    setSpinDegrees,
    setScalePercent,
    setEmphasisColor,
    setPathPreset,
    setPathDirection,
    setPathDistancePercent,
    timelineOpen,
    toggleTimeline,
    timeline,
    timelineDraft,
    selectBar,
    startEdgeDrag,
    moveEdgeDrag,
    endEdgeDrag,
    edgeKeydown,
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
