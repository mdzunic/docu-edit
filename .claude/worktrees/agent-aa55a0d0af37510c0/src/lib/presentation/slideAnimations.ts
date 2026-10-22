// The frame-animation editing commands (SPEC-086), kept in the framework-free
// lib layer so the panel composable stays thin and the add/edit/reorder/remove
// rules (each of which dirties the slide and drops any preserved out-of-subset
// timing) are unit-testable without a component. The panel decides WHEN to run
// these and gates the destructive replace behind a confirmation; this module
// decides WHAT they do to the model.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type { AnimationEffect, PresentationFormat, Slide, SlideAnimation } from "./model";
import { markSlideEdited } from "./text";
import { defaultAnimationDirection, dropPreservedTiming } from "./animations";

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
 * `undefined` clear back to the effect default. */
export interface AnimationPatch {
  effect?: AnimationEffect;
  direction?: SlideAnimation["direction"];
  trigger?: SlideAnimation["trigger"];
  durationMs?: number;
  delayMs?: number;
  clearDirection?: true;
  clearDuration?: true;
  clearDelay?: true;
}

/** Edits the animation at `index`, dirtying the slide. Changing the effect to a
 * non-directional one drops a now-meaningless direction. */
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
    // A direction is only meaningful for fly/wipe; drop it (or seed the default)
    // when the effect changes.
    if (defaultAnimationDirection(patch.effect) === undefined) {
      delete animation.direction;
    }
  }
  if (patch.direction !== undefined) {
    animation.direction = patch.direction;
  }
  if (patch.clearDirection) {
    delete animation.direction;
  }
  if (patch.trigger !== undefined) {
    animation.trigger = patch.trigger;
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
