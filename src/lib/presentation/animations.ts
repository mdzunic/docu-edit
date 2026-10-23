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
  AnimationEffectClass,
  AnimationTrigger,
  MotionPath,
  Slide,
  SlideAnimation,
  SlideFrame,
  SlideParagraph,
  TransitionDirection,
} from "./model";

export type { AnimationEffect, AnimationEffectClass, AnimationTrigger, SlideAnimation } from "./model";

export const DEFAULT_ANIMATION_DURATION_MS = 500;
export const MIN_ANIMATION_DURATION_MS = 100;
export const MAX_ANIMATION_DURATION_MS = 5000;
export const DEFAULT_ANIMATION_DELAY_MS = 0;
export const MIN_ANIMATION_DELAY_MS = 0;
export const MAX_ANIMATION_DELAY_MS = 10000;

// SPEC-105 bounds and defaults for the widened subset.
export const MIN_REPEAT_COUNT = 2;
export const MAX_REPEAT_COUNT = 10;
export const DEFAULT_SPIN_DEGREES = 360;
export const MIN_SPIN_DEGREES = -3600;
export const MAX_SPIN_DEGREES = 3600;
export const DEFAULT_SCALE_PERCENT = 150;
export const MIN_SCALE_PERCENT = 10;
export const MAX_SCALE_PERCENT = 400;
export const DEFAULT_EMPHASIS_COLOR = "#ff0000";
/** A motion path's default reach: a quarter of the slide dimension. */
export const DEFAULT_PATH_DISTANCE = 0.25;
export const MIN_PATH_DISTANCE = 0.05;
export const MAX_PATH_DISTANCE = 1;

/** Every modelled effect, in panel order (SPEC-086's seven plus SPEC-105's
 * four emphasis effects and the motion path). */
export const ANIMATION_EFFECTS: AnimationEffect[] = [
  "appear",
  "fadeIn",
  "flyIn",
  "wipeIn",
  "zoomIn",
  "fadeOut",
  "flyOut",
  "pulse",
  "spin",
  "growShrink",
  "colorChange",
  "motionPath",
];

/** The four emphasis effects (SPEC-105), in panel order. */
export const EMPHASIS_EFFECTS: AnimationEffect[] = ["pulse", "spin", "growShrink", "colorChange"];

/** The motion-path presets (SPEC-105), in panel order. `custom` is read and
 * preserved, never authored. */
export const PATH_PRESETS: MotionPath["preset"][] = ["line", "arc", "turn", "loop", "custom"];

/** The four triggers, in panel order. */
export const ANIMATION_TRIGGERS: AnimationTrigger[] = ["click", "withPrevious", "afterPrevious", "onClickOf"];

const ENTRANCE_EFFECTS = new Set<AnimationEffect>(["appear", "fadeIn", "flyIn", "wipeIn", "zoomIn"]);
const EXIT_EFFECTS = new Set<AnimationEffect>(["fadeOut", "flyOut"]);
const EMPHASIS_SET = new Set<AnimationEffect>(EMPHASIS_EFFECTS);
const DIRECTIONAL_EFFECTS = new Set<AnimationEffect>(["flyIn", "flyOut", "wipeIn"]);

/** An effect's family: entrances reveal their target, exits hide it, emphasis
 * and path effects leave its visibility alone (SPEC-105). */
export function effectClass(effect: string): AnimationEffectClass {
  if (ENTRANCE_EFFECTS.has(effect as AnimationEffect)) return "entrance";
  if (EXIT_EFFECTS.has(effect as AnimationEffect)) return "exit";
  if (EMPHASIS_SET.has(effect as AnimationEffect)) return "emphasis";
  return effect === "motionPath" ? "path" : "exit";
}

/** Whether an effect reveals its target (vs. an exit effect that hides it).
 * False for every emphasis and path effect. `type` is widened to `string` so
 * untyped .vue reactive values can call it. */
export function isEntranceEffect(effect: string): boolean {
  return ENTRANCE_EFFECTS.has(effect as AnimationEffect);
}

/** Whether an effect changes its target's visibility (entrance or exit). */
export function changesVisibility(effect: string): boolean {
  const cls = effectClass(effect);
  return cls === "entrance" || cls === "exit";
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

/** How many passes an animation plays (SPEC-105): `repeatCount` (2–10), else
 * once — `repeatUntilClick` counts as exactly one pass, so an `afterPrevious`
 * chain behind it never stalls. */
export function effectiveRepeatCount(a: SlideAnimation): number {
  if (a.repeatUntilClick || a.repeatCount === undefined) {
    return 1;
  }
  return Math.min(MAX_REPEAT_COUNT, Math.max(MIN_REPEAT_COUNT, Math.round(a.repeatCount)));
}

/** A spin's angle in degrees, defaulted (360) and clamped to ±3600. */
export function effectiveSpinDegrees(a: SlideAnimation): number {
  const deg = a.spinDegrees ?? DEFAULT_SPIN_DEGREES;
  return Math.min(MAX_SPIN_DEGREES, Math.max(MIN_SPIN_DEGREES, deg));
}

/** A grow/shrink's peak scale in percent, defaulted (150) and clamped 10–400. */
export function effectiveScalePercent(a: SlideAnimation): number {
  const pct = a.scalePercent ?? DEFAULT_SCALE_PERCENT;
  return Math.min(MAX_SCALE_PERCENT, Math.max(MIN_SCALE_PERCENT, pct));
}

const HEX_COLOR = /^#[0-9a-f]{6}$/;

/** Whether a string is a model colour (`#rrggbb`, lower case). */
export function isEmphasisColor(value: string): boolean {
  return HEX_COLOR.test(value);
}

/** A colour change's target colour, defaulted (#ff0000). */
export function effectiveEmphasisColor(a: SlideAnimation): string {
  return a.emphasisColor !== undefined && isEmphasisColor(a.emphasisColor) ? a.emphasisColor : DEFAULT_EMPHASIS_COLOR;
}

/** The animations of the slide's main (advance-driven) sequence, in play
 * order — every row except the `onClickOf` ones (SPEC-105). */
export function mainSequence(animations: SlideAnimation[]): SlideAnimation[] {
  return animations.filter((a) => a.trigger !== "onClickOf");
}

/** One interactive sequence: the rows a click on `triggerFrameId` plays. */
export interface InteractiveSequence {
  triggerFrameId: string;
  animations: SlideAnimation[];
}

/** The slide's `onClickOf` rows grouped by trigger frame, each group and the
 * rows inside it in array order (SPEC-105). A row without a trigger frame
 * (never stored by the commands) plays nowhere. */
export function interactiveSequences(animations: SlideAnimation[]): InteractiveSequence[] {
  const sequences: InteractiveSequence[] = [];
  for (const animation of animations) {
    if (animation.trigger !== "onClickOf" || animation.triggerFrameId === undefined) {
      continue;
    }
    let sequence = sequences.find((s) => s.triggerFrameId === animation.triggerFrameId);
    if (!sequence) {
      sequence = { triggerFrameId: animation.triggerFrameId, animations: [] };
      sequences.push(sequence);
    }
    sequence.animations.push(animation);
  }
  return sequences;
}

/** The trigger an animation effectively begins on: the first animation of the
 * MAIN sequence always behaves as `click`, even if marked otherwise (both
 * formats' sequencers require an initiating click; the panel and players
 * normalise it through here). An `onClickOf` row keeps its trigger and never
 * opens a main-sequence group (SPEC-105). */
export function effectiveTrigger(animations: SlideAnimation[], index: number): AnimationTrigger {
  const animation = animations[index];
  if (animation.trigger === "onClickOf") {
    return "onClickOf";
  }
  const firstMain = animations.findIndex((a) => a.trigger !== "onClickOf");
  return index === firstMain ? "click" : animation.trigger;
}

/** Partitions the MAIN sequence into build groups at each `click` boundary
 * (its first animation always opens a group). Order within and across groups
 * is play order. `onClickOf` rows are not in any group (SPEC-105). Playback
 * advances one group per input — at row level; paragraph builds expand
 * further in `buildPlayback.ts`'s `slideBuildSteps`. */
export function buildGroups(animations: SlideAnimation[]): SlideAnimation[][] {
  const main = mainSequence(animations);
  const groups: SlideAnimation[][] = [];
  main.forEach((animation, index) => {
    if (index === 0 || effectiveTrigger(main, index) === "click") {
      groups.push([animation]);
    } else {
      groups[groups.length - 1].push(animation);
    }
  });
  return groups;
}

/** The SPEC-105 model constraints a stored row must meet, against the ids of
 * its slide's frames: `triggerFrameId` present iff the trigger is
 * `onClickOf`, naming another frame on the slide (never the row's own
 * target); `repeatCount` (2–10) and `repeatUntilClick` exclusive; `path`
 * present iff the effect is `motionPath`; and the per-field bounds. The
 * readers model a row only when it holds; the commands never break it. */
export function animationConstraintsHold(a: SlideAnimation, frameIds: ReadonlySet<string>): boolean {
  if (!frameIds.has(a.frameId)) return false;
  if ((a.trigger === "onClickOf") !== (a.triggerFrameId !== undefined)) return false;
  if (a.triggerFrameId !== undefined && (a.triggerFrameId === a.frameId || !frameIds.has(a.triggerFrameId))) return false;
  if (a.repeatCount !== undefined && a.repeatUntilClick) return false;
  if (
    a.repeatCount !== undefined &&
    (!Number.isInteger(a.repeatCount) || a.repeatCount < MIN_REPEAT_COUNT || a.repeatCount > MAX_REPEAT_COUNT)
  ) {
    return false;
  }
  if ((a.effect === "motionPath") !== (a.path !== undefined)) return false;
  if (a.path?.preset === "custom" ? a.path.d === undefined : a.path?.d !== undefined) return false;
  if (a.spinDegrees !== undefined && (a.spinDegrees < MIN_SPIN_DEGREES || a.spinDegrees > MAX_SPIN_DEGREES)) return false;
  if (a.scalePercent !== undefined && (a.scalePercent < MIN_SCALE_PERCENT || a.scalePercent > MAX_SCALE_PERCENT)) return false;
  if (a.emphasisColor !== undefined && !isEmphasisColor(a.emphasisColor)) return false;
  return true;
}

/** Retargets a copied slide's animations onto its copied frames: each source
 * frame id — the target and, for an `onClickOf` row, the trigger frame — maps
 * to the copy's fresh id; animations whose frame did not survive the copy are
 * dropped (AC-19). */
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
    const nextTrigger = animation.triggerFrameId === undefined ? undefined : idMap.get(animation.triggerFrameId);
    if (nextId && (animation.triggerFrameId === undefined || nextTrigger)) {
      remapped.push(nextTrigger ? { ...animation, frameId: nextId, triggerFrameId: nextTrigger } : { ...animation, frameId: nextId });
    }
  }
  return remapped;
}

/** Drops every animation targeting `frameId` (AC-18) — and every `onClickOf`
 * row it triggers, which would otherwise name a frame the slide no longer has
 * (SPEC-105). Returns a fresh array; a no-op returns the same reference so
 * callers can detect "nothing removed".*/
export function removeAnimationsForFrame(
  animations: SlideAnimation[] | undefined,
  frameId: string,
): SlideAnimation[] | undefined {
  const involves = (a: SlideAnimation): boolean => a.frameId === frameId || a.triggerFrameId === frameId;
  if (!animations || !animations.some(involves)) {
    return animations;
  }
  return animations.filter((a) => !involves(a));
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
