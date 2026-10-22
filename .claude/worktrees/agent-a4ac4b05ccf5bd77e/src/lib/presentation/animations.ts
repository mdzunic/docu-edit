// The frame-animation model helpers (SPEC-086), kept in the framework-free lib
// layer so the panel composable stays thin and the grouping / retargeting /
// trigger-normalisation rules are unit-testable without a component. The
// per-format wire mappings live in `pptx/timing.ts` and `odp/anim.ts`; THIS
// module is the normative source for the format-neutral rules: how animations
// partition into build groups, how a duplicated slide's animations retarget to
// its copied frames, how deleting a frame drops its animations, how preserved
// (out-of-subset) timing keeps or loses its targets, and the defaults/bounds
// the panel and players read through.
//
// Framework-free (imports only the model types).
import type {
  AnimationEffect,
  AnimationTrigger,
  Slide,
  SlideAnimation,
  SlideFrame,
  SlideParagraph,
  TransitionDirection,
} from "./model";

export type { AnimationEffect, AnimationTrigger, SlideAnimation } from "./model";

export const DEFAULT_ANIMATION_DURATION_MS = 500;
export const MIN_ANIMATION_DURATION_MS = 100;
export const MAX_ANIMATION_DURATION_MS = 5000;
export const DEFAULT_ANIMATION_DELAY_MS = 0;
export const MIN_ANIMATION_DELAY_MS = 0;
export const MAX_ANIMATION_DELAY_MS = 10000;

/** The seven effects, in panel order. */
export const ANIMATION_EFFECTS: AnimationEffect[] = [
  "appear",
  "fadeIn",
  "flyIn",
  "wipeIn",
  "zoomIn",
  "fadeOut",
  "flyOut",
];

/** The three triggers, in panel order. */
export const ANIMATION_TRIGGERS: AnimationTrigger[] = ["click", "withPrevious", "afterPrevious"];

const ENTRANCE_EFFECTS = new Set<AnimationEffect>(["appear", "fadeIn", "flyIn", "wipeIn", "zoomIn"]);
const DIRECTIONAL_EFFECTS = new Set<AnimationEffect>(["flyIn", "flyOut", "wipeIn"]);

/** Whether an effect reveals its target (vs. an exit effect that hides it).
 * `type` is widened to `string` so untyped .vue reactive values can call it. */
export function isEntranceEffect(effect: string): boolean {
  return ENTRANCE_EFFECTS.has(effect as AnimationEffect);
}

/** Whether an effect's direction is user-choosable (fly/wipe). */
export function usesAnimationDirection(effect: string): boolean {
  return DIRECTIONAL_EFFECTS.has(effect as AnimationEffect);
}

/** A directional effect's default direction (all default to left); undefined
 * for the effects that ignore direction. */
export function defaultAnimationDirection(effect: string): TransitionDirection | undefined {
  return usesAnimationDirection(effect) ? "left" : undefined;
}

/** The direction an animation effectively plays with — its own, or the
 * per-effect default. */
export function effectiveAnimationDirection(a: SlideAnimation): TransitionDirection | undefined {
  return a.direction ?? defaultAnimationDirection(a.effect);
}

/** The animation's duration in ms, defaulting and clamping to the model bounds. */
export function effectiveAnimationDurationMs(a: SlideAnimation): number {
  const ms = a.durationMs ?? DEFAULT_ANIMATION_DURATION_MS;
  return Math.min(MAX_ANIMATION_DURATION_MS, Math.max(MIN_ANIMATION_DURATION_MS, ms));
}

/** The animation's delay in ms, defaulting and clamping to the model bounds. */
export function effectiveAnimationDelayMs(a: SlideAnimation): number {
  const ms = a.delayMs ?? DEFAULT_ANIMATION_DELAY_MS;
  return Math.min(MAX_ANIMATION_DELAY_MS, Math.max(MIN_ANIMATION_DELAY_MS, ms));
}

/** The trigger an animation effectively begins on: the first animation of a
 * slide always behaves as `click`, even if marked otherwise (both formats'
 * sequencers require an initiating click; the panel and players normalise
 * row 1 through here). */
export function effectiveTrigger(animations: SlideAnimation[], index: number): AnimationTrigger {
  return index === 0 ? "click" : animations[index].trigger;
}

/** Partitions animations into build groups at each `click` boundary (the first
 * animation always opens a group). Order within and across groups is play
 * order. The writers use this to structure the click groups; playback advances
 * one group per input. */
export function buildGroups(animations: SlideAnimation[]): SlideAnimation[][] {
  const groups: SlideAnimation[][] = [];
  animations.forEach((animation, index) => {
    if (index === 0 || effectiveTrigger(animations, index) === "click") {
      groups.push([animation]);
    } else {
      groups[groups.length - 1].push(animation);
    }
  });
  return groups;
}

/** Retargets a copied slide's animations onto its copied frames: each source
 * frame id maps to the copy's fresh id; animations whose frame did not survive
 * the copy are dropped (AC-19). */
export function retargetAnimations(
  animations: SlideAnimation[] | undefined,
  idMap: Map<string, string>,
): SlideAnimation[] | undefined {
  if (!animations) {
    return undefined;
  }
  const remapped: SlideAnimation[] = [];
  for (const animation of animations) {
    const nextId = idMap.get(animation.frameId);
    if (nextId) {
      remapped.push({ ...animation, frameId: nextId });
    }
  }
  return remapped;
}

/** Drops every animation targeting `frameId` (AC-18). Returns a fresh array;
 * a no-op returns the same reference so callers can detect "nothing removed".*/
export function removeAnimationsForFrame(
  animations: SlideAnimation[] | undefined,
  frameId: string,
): SlideAnimation[] | undefined {
  if (!animations || !animations.some((a) => a.frameId === frameId)) {
    return animations;
  }
  return animations.filter((a) => a.frameId !== frameId);
}

/** Whether a slide carries out-of-subset timing this editor kept verbatim
 * rather than modelling (AC-4/AC-17) — the honest "original animations kept"
 * state. pptx keeps it under `p:timing`, odp under `anim`. */
export function hasPreservedTiming(slide: Slide): boolean {
  return !!(slide.preserved && (slide.preserved["p:timing"] || slide.preserved["anim"]));
}

// --- preserved-timing targets ----------------------------------------------
// Preserved timing is verbatim XML that names its targets by SOURCE wire id.
// The readers record that id (`sourceWireId`) on each frame/paragraph it
// targets, and the writers keep the reference resolving after a rebuild — or
// drop the timing whole rather than emit a reference to nothing, since a
// preserved tree is never half-edited (AC-4). The rules below keep the model
// honest about it: a command that removes a target drops the timing at once
// (undo restores both, as with modelled animations — AC-18), and an id never
// rides onto a copy, which would make two frames claim one target.

/** Every paragraph a frame carries: text/shape bodies, table cells, and a
 * group's descendants'. */
function paragraphsOf(frame: SlideFrame): SlideParagraph[] {
  switch (frame.kind) {
    case "text":
      return frame.paragraphs;
    case "shape":
      return frame.paragraphs ?? [];
    case "table":
      return frame.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs));
    case "group":
      return frame.children.flatMap(paragraphsOf);
    default:
      return [];
  }
}

/** Whether a frame — or a paragraph or descendant inside it — is a target of
 * its slide's preserved timing. */
export function carriesSourceWireId(frame: SlideFrame): boolean {
  if (frame.sourceWireId !== undefined || paragraphsOf(frame).some((p) => p.sourceWireId !== undefined)) {
    return true;
  }
  return frame.kind === "group" && frame.children.some(carriesSourceWireId);
}

/** Clears every source wire id from these frames, their paragraphs and their
 * descendants — and the source paragraph indices beside them (a duplicate's
 * copy; a slide whose preserved timing is gone). */
export function clearSourceWireIds(frames: SlideFrame[]): void {
  for (const frame of frames) {
    delete frame.sourceWireId;
    for (const para of paragraphsOf(frame)) {
      delete para.sourceWireId;
      delete para.pptxSourceParagraph;
    }
    if (frame.kind === "group") {
      clearSourceWireIds(frame.children);
    }
  }
}

/** Keeps a source wire id only where `targeted` names it — on top-level frames
 * and their own paragraphs, the first claimant of each id — and clears every
 * other (a group's descendants always: only top-level frames are targets, as
 * for modelled animations). A text/shape frame that keeps an id in
 * `rangeTargeted` — one the timing addresses by paragraph POSITION, a pptx
 * `p:txEl` range — also keeps its lines' `pptxSourceParagraph`; every other
 * paragraph loses its. The readers record every candidate, then keep the
 * ones their preserved timing references (none when it is modelled). */
export function retainSourceWireIds(
  frames: SlideFrame[],
  targeted: ReadonlySet<string>,
  rangeTargeted: ReadonlySet<string> = new Set(),
): void {
  const claimed = new Set<string>();
  const claim = (id: string | undefined): boolean => {
    if (id === undefined || !targeted.has(id) || claimed.has(id)) {
      return false;
    }
    claimed.add(id);
    return true;
  };
  for (const frame of frames) {
    if (!claim(frame.sourceWireId)) {
      delete frame.sourceWireId;
    }
    if (frame.kind === "group") {
      clearSourceWireIds(frame.children);
      continue;
    }
    const byPosition =
      (frame.kind === "text" || frame.kind === "shape") &&
      frame.sourceWireId !== undefined &&
      rangeTargeted.has(frame.sourceWireId);
    for (const para of paragraphsOf(frame)) {
      if (!claim(para.sourceWireId)) {
        delete para.sourceWireId;
      }
      if (!byPosition) {
        delete para.pptxSourceParagraph;
      }
    }
  }
}

/** Drops a slide's preserved out-of-subset timing (either format's key) with
 * the source wire ids that tied it to the frames. Leaves `animations` alone:
 * the caller decides whether the slide becomes modelled. */
export function dropPreservedTiming(slide: Slide): void {
  if (slide.preserved) {
    delete slide.preserved["p:timing"];
    delete slide.preserved["anim"];
    if (Object.keys(slide.preserved).length === 0) {
      delete slide.preserved;
    }
  }
  clearSourceWireIds(slide.frames);
}

/** Drops the slide's preserved timing when one of `removed` (frames leaving
 * the slide's top level — deleted or grouped) is one of its targets: the
 * timing would otherwise name a shape the slide no longer has. */
export function dropPreservedTimingTargeting(slide: Slide, removed: readonly SlideFrame[]): void {
  if (slide.animations === undefined && hasPreservedTiming(slide) && removed.some(carriesSourceWireId)) {
    dropPreservedTiming(slide);
  }
}
