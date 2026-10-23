// The build-playback timing model (SPEC-086) — the format-neutral, framework-
// free rules the play view and the editor "Play" preview both drive: which
// frames are hidden between builds, and when each animation in a group starts.
// The CSS that renders these states lives beside the renderer
// (components/presentation/animationStyles.ts); this module is pure maths, so
// hidden-until-built and the with/after-previous chaining are unit-tested
// without a mount.
//
// SPEC-105 adds per-paragraph builds: `expandAnimation` turns a row into its
// BUILD STEPS — one per paragraph (or per level-0 paragraph) of its target,
// derived from the frame's current paragraphs and never stored, so editing
// text under a live build changes the step count and repairs nothing — and it
// is the single source of truth for both players and both serialisers. The
// main sequence partitions into step groups (`slideBuildSteps`), the
// `onClickOf` rows into one sequence per trigger frame
// (`slideInteractiveBuilds`), and visibility is tracked per build target key
// (`isBuildTargetHiddenAt`): a frame id, or `${frameId}#${paragraphIndex}`.
import type { MotionPath, Slide, SlideAnimation, SlideFrame, SlideParagraph } from "./model";
import {
  changesVisibility,
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  effectiveRepeatCount,
  effectiveSpinDegrees,
  effectiveTrigger,
  interactiveSequences,
  isEntranceEffect,
  mainSequence,
} from "./animations";
import { motionPathEnd, type PathPoint } from "./motionPaths";

/** One playable unit of a row: the whole frame, or one paragraph (by
 * paragraph) / one level-0 paragraph and the deeper ones after it (by
 * first-level paragraph). `key` names what it animates — `frameId`, or
 * `${frameId}#${paragraphIndex}` — and is what visibility is tracked by.
 * `paragraphEnd` is the last paragraph the step covers (inclusive). */
export interface BuildStep {
  animation: SlideAnimation;
  frameId: string;
  paragraphIndex?: number;
  paragraphEnd?: number;
  key: string;
}

/** The build key of a frame's paragraph. */
export function paragraphKey(frameId: string, paragraphIndex: number): string {
  return `${frameId}#${paragraphIndex}`;
}

/** The paragraphs a build can address: a text frame's, a shape's text. Any
 * other kind (an image, a table, media) has none — it builds whole. */
export function buildableParagraphs(frame: SlideFrame | undefined): SlideParagraph[] | undefined {
  if (frame?.kind === "text") return frame.paragraphs;
  if (frame?.kind === "shape") return frame.paragraphs ?? [];
  return undefined;
}

/** Whether a row's target can build by paragraph at all: a text or shape
 * frame with at least two paragraphs (fewer expand to one whole-frame step). */
export function canBuildByParagraph(frame: SlideFrame | undefined): boolean {
  return (buildableParagraphs(frame)?.length ?? 0) >= 2;
}

/** The build steps of one row against its target frame's CURRENT paragraphs:
 * no `buildBy`, a target with no paragraphs, or fewer than two paragraphs →
 * one whole-frame step (key = frameId, no paragraphIndex); `byParagraph` →
 * one step per paragraph in order; `byFirstLevelParagraph` → one step per
 * level-0 (or absent-level) paragraph covering every deeper paragraph up to
 * the next level-0 one, a leading deeper paragraph joining the first step. */
export function expandAnimation(a: SlideAnimation, frame: SlideFrame | undefined): BuildStep[] {
  const paragraphs = buildableParagraphs(frame);
  if (!a.buildBy || !paragraphs || paragraphs.length < 2) {
    return [{ animation: a, frameId: a.frameId, key: a.frameId }];
  }
  const step = (start: number, end: number): BuildStep => ({
    animation: a,
    frameId: a.frameId,
    paragraphIndex: start,
    paragraphEnd: end,
    key: paragraphKey(a.frameId, start),
  });
  if (a.buildBy === "byParagraph") {
    return paragraphs.map((_, i) => step(i, i));
  }
  const levelZero = paragraphs.flatMap((p, i) => ((p.level ?? 0) === 0 ? [i] : []));
  // The first step always starts at 0: a leading deeper paragraph joins it.
  const starts = [0, ...levelZero.slice(1)];
  return starts.map((start, k) => step(start, (starts[k + 1] ?? paragraphs.length) - 1));
}

/** A slide's frames by id, the lookup `expandAnimation` reads paragraphs
 * through. */
export function frameLookup(frames: readonly SlideFrame[]): (frameId: string) => SlideFrame | undefined {
  const byId = new Map(frames.map((f) => [f.id, f]));
  return (frameId) => byId.get(frameId);
}

/** The MAIN sequence of `animations` expanded and partitioned into build
 * groups (what `usePlayback` advances through, one group per input): each
 * row's first step opens a group exactly as SPEC-086's trigger rule does (the
 * first main row always), and each later step opens a new group when the
 * row's `buildTrigger` is "click" (the default) or joins the current one.
 * `frameFor` resolves a target to its frame (both serialisers pass theirs). */
export function mainBuildSteps(
  animations: SlideAnimation[],
  frameFor: (frameId: string) => SlideFrame | undefined,
): BuildStep[][] {
  const main = mainSequence(animations);
  const groups: BuildStep[][] = [];
  main.forEach((animation, index) => {
    expandAnimation(animation, frameFor(animation.frameId)).forEach((step, k) => {
      const opens =
        k === 0
          ? index === 0 || effectiveTrigger(main, index) === "click"
          : (animation.buildTrigger ?? "click") === "click";
      if (opens || groups.length === 0) {
        groups.push([step]);
      } else {
        groups[groups.length - 1].push(step);
      }
    });
  });
  return groups;
}

/** The slide's main sequence as build-step groups (`mainBuildSteps`). */
export function slideBuildSteps(slide: Slide): BuildStep[][] {
  return mainBuildSteps(slide.animations ?? [], frameLookup(slide.frames));
}

/** One interactive sequence's playable groups: each click on the trigger
 * frame plays the next group. */
export interface InteractiveBuild {
  triggerFrameId: string;
  groups: BuildStep[][];
}

/** The `onClickOf` rows of `animations` as one step sequence per trigger
 * frame (in order of first appearance). Every row starts with the click on
 * its trigger — the first opens the sequence's first group and each later row
 * joins the group in play, alongside its predecessor — and a by-paragraph
 * row's later steps follow its `buildTrigger` as in the main sequence (a
 * "click" step waits for the next click on the same trigger). */
export function interactiveBuildSteps(
  animations: SlideAnimation[],
  frameFor: (frameId: string) => SlideFrame | undefined,
): InteractiveBuild[] {
  return interactiveSequences(animations).map((sequence) => {
    const groups: BuildStep[][] = [];
    for (const animation of sequence.animations) {
      expandAnimation(animation, frameFor(animation.frameId)).forEach((step, k) => {
        const opens = groups.length === 0 || (k > 0 && (animation.buildTrigger ?? "click") === "click");
        if (opens) {
          groups.push([step]);
        } else {
          groups[groups.length - 1].push(step);
        }
      });
    }
    return { triggerFrameId: sequence.triggerFrameId, groups };
  });
}

/** The slide's interactive sequences (`interactiveBuildSteps`) — never
 * counted by `slideBuildCount`. */
export function slideInteractiveBuilds(slide: Slide): InteractiveBuild[] {
  return interactiveBuildSteps(slide.animations ?? [], frameLookup(slide.frames));
}

/** The slide's main-sequence build groups at row level: each step group's
 * distinct animations, in play order. */
export function slideBuildGroups(slide: Slide): SlideAnimation[][] {
  return slideBuildSteps(slide).map((group) => group.map((s) => s.animation).filter((a, i, all) => all.indexOf(a) === i));
}

/** How many build groups a slide has (what `usePlayback` advances through) —
 * main-sequence groups only, so SPEC-084's `buildCountAt` never waits on an
 * interactive group. */
export function slideBuildCount(slide: Slide): number {
  return slideBuildSteps(slide).length;
}

/** Whether a frame is hidden when the slide is settled with `groupsPlayed`
 * groups shown (no animation in flight). A frame targeted by an entrance effect
 * starts hidden and stays hidden until that effect's group plays; an exit
 * effect hides its target once its group plays. Only entrance and exit effects
 * change visibility — a played emphasis or path effect leaves it alone — and a
 * `rewindOnComplete` effect restores what it found. The last played animation
 * for the frame wins; a frame with no entrance and no played exit is visible
 * (AC-20/21). */
export function isFrameHiddenAt(groups: SlideAnimation[][], frameId: string, groupsPlayed: number): boolean {
  const all = groups.flat();
  let hidden = all.some((a) => a.frameId === frameId && isEntranceEffect(a.effect));
  for (let gi = 0; gi < groupsPlayed && gi < groups.length; gi++) {
    for (const a of groups[gi]) {
      if (a.frameId === frameId && changesVisibility(a.effect) && !a.rewindOnComplete) {
        hidden = !isEntranceEffect(a.effect);
      }
    }
  }
  return hidden;
}

/** Whether build target `key` (a frame id, or a paragraph's
 * `${frameId}#${index}`) is hidden once `played` of `all` the slide's steps
 * have run: hidden from the start when any entrance step addresses it, then
 * each played entrance/exit step for it shows/hides it — except one that
 * rewinds, which leaves the state it found. */
export function isBuildTargetHiddenAfter(all: readonly BuildStep[], played: readonly BuildStep[], key: string): boolean {
  let hidden = all.some((s) => s.key === key && isEntranceEffect(s.animation.effect));
  for (const s of played) {
    if (s.key === key && changesVisibility(s.animation.effect) && !s.animation.rewindOnComplete) {
      hidden = !isEntranceEffect(s.animation.effect);
    }
  }
  return hidden;
}

/** `isFrameHiddenAt` over build-step groups and target keys: whether `key` is
 * hidden when the slide is settled with `groupsPlayed` groups shown. */
export function isBuildTargetHiddenAt(groups: BuildStep[][], key: string, groupsPlayed: number): boolean {
  return isBuildTargetHiddenAfter(groups.flat(), groups.slice(0, Math.max(0, groupsPlayed)).flat(), key);
}

/** Where played emphasis/path steps leave target `key` once they have ended
 * (SPEC-105's settled end state): the net rotation of its spins and the net
 * offset (slide fractions) of its motion paths. An auto-reversed pass ends
 * where it began, a rewind restores the start, and pulse / grow-shrink /
 * colour change always return. */
export function settledMotion(played: readonly BuildStep[], key: string): { rotateDeg: number; offset: PathPoint } {
  let rotateDeg = 0;
  const offset = { x: 0, y: 0 };
  for (const s of played) {
    const a = s.animation;
    if (s.key !== key || a.rewindOnComplete || a.autoReverse) {
      continue;
    }
    if (a.effect === "spin") {
      rotateDeg += effectiveSpinDegrees(a) * effectiveRepeatCount(a);
    } else if (a.effect === "motionPath" && a.path) {
      const end = motionPathEnd(a.path as MotionPath);
      offset.x += end.x;
      offset.y += end.y;
    }
  }
  return { rotateDeg, offset };
}

/** One scheduled animation within a group: when it starts (ms from group start)
 * and how long it runs. `appear` runs instantly (0ms) so an `afterPrevious`
 * chained onto it does not wait. */
export interface TimelineEntry {
  animation: SlideAnimation;
  startMs: number;
  durationMs: number;
}

/** The animation's playback duration: a single pass (`appear` is instant) ×
 * its passes (`effectiveRepeatCount` — repeat-until-click counts once, so an
 * `afterPrevious` chain never stalls) × 2 when it auto-reverses. */
export function playbackDurationMs(a: SlideAnimation): number {
  const single = a.effect === "appear" ? 0 : effectiveAnimationDurationMs(a);
  return single * effectiveRepeatCount(a) * (a.autoReverse ? 2 : 1);
}

/** Schedules a build group's animations: the first begins at its own delay;
 * `withPrevious` begins with its predecessor (plus its delay); `afterPrevious`
 * chains after the predecessor ends (plus its delay). Returns the per-animation
 * timeline and the group's total running time. */
export function groupTimeline(group: SlideAnimation[]): { entries: TimelineEntry[]; totalMs: number } {
  const entries: TimelineEntry[] = [];
  let prevStart = 0;
  let prevDuration = 0;
  let total = 0;
  group.forEach((animation, i) => {
    const durationMs = playbackDurationMs(animation);
    const delay = effectiveAnimationDelayMs(animation);
    let startMs: number;
    if (i === 0) {
      startMs = delay;
    } else if (animation.trigger === "withPrevious") {
      startMs = prevStart + delay;
    } else {
      startMs = prevStart + prevDuration + delay;
    }
    entries.push({ animation, startMs, durationMs });
    prevStart = startMs;
    prevDuration = durationMs;
    total = Math.max(total, startMs + durationMs);
  });
  return { entries, totalMs: total };
}

/** One scheduled step within a group. */
export interface StepTimelineEntry {
  step: BuildStep;
  startMs: number;
  durationMs: number;
}

/** Schedules a group of build steps as `groupTimeline` schedules rows: the
 * first starts at its delay; a row's first step joins on its own trigger
 * (`afterPrevious` chains, anything else — `withPrevious`, an `onClickOf` row
 * joining its sequence's click — starts with its predecessor); a row's later
 * step on its `buildTrigger` ("withPrevious" alongside, otherwise after the
 * previous step). Every step carries the row's delay and duration. */
export function stepTimeline(steps: BuildStep[]): { entries: StepTimelineEntry[]; totalMs: number } {
  const entries: StepTimelineEntry[] = [];
  let prevStart = 0;
  let prevDuration = 0;
  let total = 0;
  steps.forEach((step, i) => {
    const a = step.animation;
    const durationMs = playbackDurationMs(a);
    const delay = effectiveAnimationDelayMs(a);
    let after: boolean;
    if (i === 0) {
      after = false;
    } else if (steps[i - 1].animation !== a) {
      after = a.trigger === "afterPrevious";
    } else {
      after = a.buildTrigger !== "withPrevious";
    }
    const startMs = i === 0 ? delay : (after ? prevStart + prevDuration : prevStart) + delay;
    entries.push({ step, startMs, durationMs });
    prevStart = startMs;
    prevDuration = durationMs;
    total = Math.max(total, startMs + durationMs);
  });
  return { entries, totalMs: total };
}

/** The trigger frame of one of the slide's interactive sequences under a
 * slide-space point (cm), topmost first — what a show's pointer-down hits. */
export function interactiveTriggerAt(slide: Slide, xCm: number, yCm: number): string | undefined {
  const triggers = new Set(interactiveSequences(slide.animations ?? []).map((s) => s.triggerFrameId));
  for (let i = slide.frames.length - 1; i >= 0; i--) {
    const f = slide.frames[i];
    if (xCm >= f.xCm && xCm <= f.xCm + f.widthCm && yCm >= f.yCm && yCm <= f.yCm + f.heightCm) {
      return triggers.has(f.id) ? f.id : undefined;
    }
  }
  return undefined;
}
