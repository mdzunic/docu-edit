// The build-playback timing model (SPEC-086) — the format-neutral, framework-
// free rules the play view and the editor "Play" preview both drive: which
// frames are hidden between builds, and when each animation in a group starts.
// The CSS that renders these states lives beside the renderer
// (components/presentation/animationStyles.ts); this module is pure maths, so
// hidden-until-built and the with/after-previous chaining are unit-tested
// without a mount.
import type { Slide, SlideAnimation } from "./model";
import {
  buildGroups,
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  isEntranceEffect,
} from "./animations";

/** The slide's animations partitioned into click-triggered build groups. */
export function slideBuildGroups(slide: Slide): SlideAnimation[][] {
  return buildGroups(slide.animations ?? []);
}

/** How many build groups a slide has (what `usePlayback` advances through). */
export function slideBuildCount(slide: Slide): number {
  return slideBuildGroups(slide).length;
}

/** Whether a frame is hidden when the slide is settled with `groupsPlayed`
 * groups shown (no animation in flight). A frame targeted by an entrance effect
 * starts hidden and stays hidden until that effect's group plays; an exit
 * effect hides its target once its group plays. The last played animation for
 * the frame wins; a frame with no entrance and no played exit is visible
 * (AC-20/21). */
export function isFrameHiddenAt(groups: SlideAnimation[][], frameId: string, groupsPlayed: number): boolean {
  const all = groups.flat();
  let hidden = all.some((a) => a.frameId === frameId && isEntranceEffect(a.effect));
  for (let gi = 0; gi < groupsPlayed && gi < groups.length; gi++) {
    for (const a of groups[gi]) {
      if (a.frameId === frameId) {
        hidden = !isEntranceEffect(a.effect);
      }
    }
  }
  return hidden;
}

/** One scheduled animation within a group: when it starts (ms from group start)
 * and how long it runs. `appear` runs instantly (0ms) so an `afterPrevious`
 * chained onto it does not wait. */
export interface TimelineEntry {
  animation: SlideAnimation;
  startMs: number;
  durationMs: number;
}

/** The animation's playback duration — `appear` is instant. */
export function playbackDurationMs(a: SlideAnimation): number {
  return a.effect === "appear" ? 0 : effectiveAnimationDurationMs(a);
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
