// The frame-animation editing commands (SPEC-086), kept in the framework-free
// lib layer so the panel composable stays thin and the add/edit/reorder/remove
// rules (each of which dirties the slide and drops any preserved out-of-subset
// timing) are unit-testable without a component. The panel decides WHEN to run
// these and gates the destructive replace behind a confirmation; this module
// decides WHAT they do to the model.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type {
  AnimationEffect,
  BuildStepTrigger,
  MotionPath,
  PresentationFormat,
  Slide,
  SlideAnimation,
  TextBuildMode,
} from "./model";
import { markSlideEdited } from "./text";
import {
  DEFAULT_EMPHASIS_COLOR,
  DEFAULT_PATH_DISTANCE,
  DEFAULT_SCALE_PERCENT,
  DEFAULT_SPIN_DEGREES,
  defaultAnimationDirection,
  dropPreservedTiming,
  isEmphasisColor,
  MAX_ANIMATION_DURATION_MS,
  MAX_PATH_DISTANCE,
  MAX_REPEAT_COUNT,
  MAX_SCALE_PERCENT,
  MAX_SPIN_DEGREES,
  MIN_PATH_DISTANCE,
  MIN_REPEAT_COUNT,
  MIN_SCALE_PERCENT,
  MIN_SPIN_DEGREES,
} from "./animations";
import { parseMotionPathD } from "./motionPaths";

/** A new animation targets `frameId` with the neutral default: fade in, on
 * click. */
function newAnimation(frameId: string): SlideAnimation {
  return { frameId, effect: "fadeIn", trigger: "click" };
}

/** Ensures the slide holds a (possibly empty) modelled animation array, dropping
 * any preserved out-of-subset timing the first time it becomes modelled
 * (confirm-to-replace has already run — AC-17). */
function ensureModelled(slide: Slide): SlideAnimation[] {
  if (!slide.animations) {
    slide.animations = [];
    dropPreservedTiming(slide);
  }
  return slide.animations;
}

/** Appends a fade-in/on-click animation targeting `frameId`, dirties the slide
 * and drops preserved timing. Returns the new animation's index. */
export function addAnimation(slide: Slide, frameId: string, format: PresentationFormat): number {
  const animations = ensureModelled(slide);
  animations.push(newAnimation(frameId));
  markSlideEdited(slide, format);
  return animations.length - 1;
}

/** The per-field edit a panel control makes to the selected animation. Absent
 * fields are left untouched; `direction`/`durationMs`/`delayMs` set to
 * `undefined` clear back to the effect default. SPEC-105 adds the build,
 * repetition, trigger-frame, effect-option and path fields; the three flags
 * take a boolean (only `true` is ever stored). */
export interface AnimationPatch {
  effect?: AnimationEffect;
  direction?: SlideAnimation["direction"];
  trigger?: SlideAnimation["trigger"];
  durationMs?: number;
  delayMs?: number;
  clearDirection?: true;
  clearDuration?: true;
  clearDelay?: true;
  /** The frame an `onClickOf` row waits for (with or without `trigger`). */
  triggerFrameId?: string;
  buildBy?: TextBuildMode;
  clearBuildBy?: true;
  buildTrigger?: BuildStepTrigger;
  /** 2–10 passes; clears `repeatUntilClick`. */
  repeatCount?: number;
  clearRepeat?: true;
  /** true clears `repeatCount`. */
  repeatUntilClick?: boolean;
  autoReverse?: boolean;
  rewindOnComplete?: boolean;
  spinDegrees?: number;
  scalePercent?: number;
  emphasisColor?: string;
  path?: MotionPath;
}

/** A path normalised to the stored form: defaults omitted, bounds clamped, a
 * custom `d` kept only when it parses (else the path is refused). */
function normalisedPath(path: MotionPath): MotionPath | undefined {
  if (path.preset === "custom") {
    return path.d !== undefined && parseMotionPathD(path.d) ? { preset: "custom", d: path.d } : undefined;
  }
  const next: MotionPath = { preset: path.preset };
  if (path.direction !== undefined && path.direction !== "right" && path.preset !== "loop") {
    next.direction = path.direction;
  }
  if (path.distanceFraction !== undefined) {
    const d = Math.min(MAX_PATH_DISTANCE, Math.max(MIN_PATH_DISTANCE, Number(path.distanceFraction.toFixed(4))));
    if (d !== DEFAULT_PATH_DISTANCE) {
      next.distanceFraction = d;
    }
  }
  return next;
}

/** Drops the effect-specific options the row's (new) effect no longer uses,
 * and seeds a motion path's default line (`path` present iff motionPath). */
function fitOptionsToEffect(animation: SlideAnimation): void {
  if (defaultAnimationDirection(animation.effect) === undefined) {
    delete animation.direction;
  }
  if (animation.effect !== "spin") delete animation.spinDegrees;
  if (animation.effect !== "growShrink") delete animation.scalePercent;
  if (animation.effect !== "colorChange") delete animation.emphasisColor;
  if (animation.effect === "motionPath") {
    animation.path ??= { preset: "line" };
  } else {
    delete animation.path;
  }
}

/** Edits the animation at `index`, dirtying the slide. Changing the effect
 * drops the options it no longer uses (a direction off fly/wipe, a spin angle
 * off spin, a path off motionPath …). Every SPEC-105 model constraint holds
 * after any patch: an `onClickOf` trigger is taken only with another frame on
 * the slide to wait for (any other trigger drops it), `repeatCount` and
 * `repeatUntilClick` replace each other, and a value equal to its default is
 * stored as absent. */
export function updateAnimation(
  slide: Slide,
  index: number,
  patch: AnimationPatch,
  format: PresentationFormat,
): void {
  const animation = slide.animations?.[index];
  if (!animation) {
    return;
  }
  if (patch.effect !== undefined) {
    animation.effect = patch.effect;
    fitOptionsToEffect(animation);
  }
  if (patch.direction !== undefined) {
    animation.direction = patch.direction;
  }
  if (patch.clearDirection) {
    delete animation.direction;
  }
  const validTriggerFrame = (id: string | undefined): id is string =>
    id !== undefined && id !== animation.frameId && slide.frames.some((f) => f.id === id);
  if (patch.trigger !== undefined) {
    if (patch.trigger === "onClickOf") {
      const triggerFrameId = patch.triggerFrameId ?? animation.triggerFrameId;
      if (validTriggerFrame(triggerFrameId)) {
        animation.trigger = "onClickOf";
        animation.triggerFrameId = triggerFrameId;
      }
    } else {
      animation.trigger = patch.trigger;
      delete animation.triggerFrameId;
    }
  } else if (patch.triggerFrameId !== undefined && animation.trigger === "onClickOf" && validTriggerFrame(patch.triggerFrameId)) {
    animation.triggerFrameId = patch.triggerFrameId;
  }
  if (patch.durationMs !== undefined) {
    animation.durationMs = patch.durationMs;
  }
  if (patch.clearDuration) {
    delete animation.durationMs;
  }
  if (patch.delayMs !== undefined) {
    animation.delayMs = patch.delayMs;
  }
  if (patch.clearDelay) {
    delete animation.delayMs;
  }
  if (patch.buildBy !== undefined) {
    animation.buildBy = patch.buildBy;
  }
  if (patch.clearBuildBy) {
    delete animation.buildBy;
    delete animation.buildTrigger;
  }
  if (patch.buildTrigger !== undefined && animation.buildBy !== undefined) {
    if (patch.buildTrigger === "click") {
      delete animation.buildTrigger;
    } else {
      animation.buildTrigger = patch.buildTrigger;
    }
  }
  if (patch.repeatCount !== undefined) {
    const n = Math.round(patch.repeatCount);
    if (n < MIN_REPEAT_COUNT) {
      delete animation.repeatCount;
    } else {
      animation.repeatCount = Math.min(MAX_REPEAT_COUNT, n);
      delete animation.repeatUntilClick;
    }
  }
  if (patch.clearRepeat) {
    delete animation.repeatCount;
  }
  if (patch.repeatUntilClick !== undefined) {
    if (patch.repeatUntilClick) {
      animation.repeatUntilClick = true;
      delete animation.repeatCount;
    } else {
      delete animation.repeatUntilClick;
    }
  }
  if (patch.autoReverse !== undefined) {
    if (patch.autoReverse) animation.autoReverse = true;
    else delete animation.autoReverse;
  }
  if (patch.rewindOnComplete !== undefined) {
    if (patch.rewindOnComplete) animation.rewindOnComplete = true;
    else delete animation.rewindOnComplete;
  }
  if (patch.spinDegrees !== undefined && animation.effect === "spin" && Number.isFinite(patch.spinDegrees)) {
    const deg = Math.min(MAX_SPIN_DEGREES, Math.max(MIN_SPIN_DEGREES, Math.round(patch.spinDegrees)));
    if (deg === DEFAULT_SPIN_DEGREES) delete animation.spinDegrees;
    else animation.spinDegrees = deg;
  }
  if (patch.scalePercent !== undefined && animation.effect === "growShrink" && Number.isFinite(patch.scalePercent)) {
    const pct = Math.min(MAX_SCALE_PERCENT, Math.max(MIN_SCALE_PERCENT, Math.round(patch.scalePercent)));
    if (pct === DEFAULT_SCALE_PERCENT) delete animation.scalePercent;
    else animation.scalePercent = pct;
  }
  if (patch.emphasisColor !== undefined && animation.effect === "colorChange") {
    const color = patch.emphasisColor.toLowerCase();
    if (isEmphasisColor(color)) {
      if (color === DEFAULT_EMPHASIS_COLOR) delete animation.emphasisColor;
      else animation.emphasisColor = color;
    }
  }
  if (patch.path !== undefined && animation.effect === "motionPath") {
    const path = normalisedPath(patch.path);
    if (path) {
      animation.path = path;
    }
  }
  // A colour change writes each half (to the colour, auto-reversed back) in
  // whole milliseconds, so its duration is kept in 2ms steps — an odd one
  // would come back 1ms longer from either format.
  if (animation.effect === "colorChange" && animation.durationMs !== undefined && animation.durationMs % 2 !== 0) {
    animation.durationMs = Math.min(MAX_ANIMATION_DURATION_MS, animation.durationMs + 1);
  }
  markSlideEdited(slide, format);
}

/** Removes the animation at `index`, dirtying the slide. Returns the index to
 * select next (the previous row, or 0), or -1 when none remain. */
export function removeAnimation(slide: Slide, index: number, format: PresentationFormat): number {
  const animations = slide.animations;
  if (!animations || index < 0 || index >= animations.length) {
    return animations && animations.length > 0 ? 0 : -1;
  }
  animations.splice(index, 1);
  markSlideEdited(slide, format);
  if (animations.length === 0) {
    return -1;
  }
  return Math.max(0, index - 1);
}

/** Moves the animation at `index` by `delta` (±1), dirtying the slide. Returns
 * the moved animation's new index (unchanged when the move is out of range). */
export function moveAnimation(slide: Slide, index: number, delta: number, format: PresentationFormat): number {
  const animations = slide.animations;
  if (!animations) {
    return index;
  }
  const target = index + delta;
  if (index < 0 || index >= animations.length || target < 0 || target >= animations.length) {
    return index;
  }
  const [moved] = animations.splice(index, 1);
  animations.splice(target, 0, moved);
  markSlideEdited(slide, format);
  return target;
}
