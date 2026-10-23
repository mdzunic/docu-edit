// The Animations panel's SPEC-105 field state and commands, split out of
// useAnimations (which creates it and spreads it into its controller, so the
// panel and PresentationView see one controller as before): which of the
// build, build-trigger, repeat, auto-reverse, rewind, trigger-frame, spin,
// scale, colour and motion-path controls the selected row shows, their
// values, and the edits — each through the controller's `patch`, one undo
// entry in SPEC-080's voice that dirties the slide and persists.
import { computed, type ComputedRef, type Ref } from "vue";
import type {
  BuildStepTrigger,
  MotionPath,
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
  effectiveEmphasisColor,
  effectiveRepeatCount,
  effectiveScalePercent,
  effectiveSpinDegrees,
} from "../../lib/presentation/animations";
import type { AnimationPatch } from "../../lib/presentation/slideAnimations";
import { buildableParagraphs, canBuildByParagraph } from "../../lib/presentation/buildPlayback";
import { frameAccessibleName } from "./useFrameManipulation";

/** The frame kinds an animation may target (AC out-of-scope: group/foreign
 * cannot be animated). SPEC-095 adds media — a clip can enter/exit like any
 * other frame (its own playback duration is never a build step). */
export function isAnimatableKind(kind: string): boolean {
  return (
    kind === "text" || kind === "image" || kind === "shape" || kind === "table" || kind === "media"
  );
}

export interface AnimationOptionsDeps {
  currentSlide: ComputedRef<Slide | null>;
  /** The current slide's rows (a fresh array per commit). */
  animations: ComputedRef<SlideAnimation[]>;
  selectedIndex: Ref<number>;
  /** One edit of the selected row: record(label) → update → commit. */
  patch: (label: string, change: AnimationPatch) => void;
}

export function useAnimationOptions(deps: AnimationOptionsDeps) {
  const { animations, selectedIndex, patch } = deps;
  function at(index: number): SlideAnimation | undefined {
    return animations.value[index];
  }

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
  return {
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
  };
}
