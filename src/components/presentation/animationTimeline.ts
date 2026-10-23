// The animation timeline's geometry (SPEC-105) — the second view of a slide's
// animation list, beside AnimationsPanel's rows: a ruler at 60px per second
// with one tick per second, one band per click group (the main sequence's,
// then each interactive sequence's), and one bar per build step offset by its
// `stepTimeline` start and sized by its playback duration (8px at least, so
// a zero-length appear still shows and overlaps stay visible). Framework-free
// and unit-tested; AnimationTimeline.vue only renders what this computes, and
// the drag/arrow edits go through the useAnimations controller.
import type { Slide } from "../../lib/presentation/model";
import {
  MAX_ANIMATION_DELAY_MS,
  MAX_ANIMATION_DURATION_MS,
  MIN_ANIMATION_DELAY_MS,
  MIN_ANIMATION_DURATION_MS,
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  effectiveRepeatCount,
} from "../../lib/presentation/animations";
import { slideBuildSteps, slideInteractiveBuilds, stepTimeline, type BuildStep } from "../../lib/presentation/buildPlayback";

/** The ruler's scale. */
export const PX_PER_SECOND = 60;
/** The narrowest a bar is drawn. */
export const MIN_BAR_PX = 8;
/** An edge drag or arrow key moves delay/duration in these steps. */
export const TIMELINE_STEP_MS = 100;

/** A bar's horizontal extent in px: offset by its start, sized by its
 * duration, never narrower than MIN_BAR_PX. */
export function barRect(startMs: number, durationMs: number): { left: number; width: number } {
  return {
    left: (startMs / 1000) * PX_PER_SECOND,
    width: Math.max(MIN_BAR_PX, (durationMs / 1000) * PX_PER_SECOND),
  };
}

/** One ruler tick per second from 0 through the last second `totalMs`
 * reaches (at least one second's worth). */
export function rulerTicks(totalMs: number): { seconds: number; x: number; label: string }[] {
  const last = Math.max(1, Math.ceil(totalMs / 1000));
  return Array.from({ length: last + 1 }, (_, seconds) => ({ seconds, x: seconds * PX_PER_SECOND, label: `${seconds}s` }));
}

/** A horizontal pointer travel (px) as milliseconds on the ruler, in whole
 * TIMELINE_STEP_MS steps. */
export function msFromDx(dx: number): number {
  return Math.round(((dx / PX_PER_SECOND) * 1000) / TIMELINE_STEP_MS) * TIMELINE_STEP_MS;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

/** The delay a left-edge drag of `dx` px (or an arrow step of ±100ms as
 * `deltaMs`) leaves, clamped to SPEC-086's 0–10000. */
export function draggedDelayMs(delayMs: number, deltaMs: number): number {
  return clamp(Math.round((delayMs + deltaMs) / TIMELINE_STEP_MS) * TIMELINE_STEP_MS, MIN_ANIMATION_DELAY_MS, MAX_ANIMATION_DELAY_MS);
}

/** The duration a right-edge drag leaves, clamped to SPEC-086's 100–5000.
 * The bar spans every pass (`passes` — repeats × auto-reverse), so the edge
 * moves the single pass by its share. */
export function draggedDurationMs(durationMs: number, deltaMs: number, passes = 1): number {
  const perPass = deltaMs / Math.max(1, passes);
  return clamp(
    Math.round((durationMs + perPass) / TIMELINE_STEP_MS) * TIMELINE_STEP_MS,
    MIN_ANIMATION_DURATION_MS,
    MAX_ANIMATION_DURATION_MS,
  );
}

/** How many single-duration spans a row's bar covers. */
export function barPasses(a: BuildStep["animation"]): number {
  return effectiveRepeatCount(a) * (a.autoReverse ? 2 : 1);
}

export interface TimelineBar {
  /** Running index over every bar on the timeline (the bar's test id). */
  barIndex: number;
  /** The row (`Slide.animations` index) the bar plays — what selecting it
   * selects. */
  animationIndex: number;
  /** 0-based paragraph the step builds, when it is a paragraph step. */
  paragraphIndex?: number;
  startMs: number;
  durationMs: number;
  delayMs: number;
  singleMs: number;
  passes: number;
  left: number;
  width: number;
}

export interface TimelineGroup {
  /** Running group index (the band's test id). */
  index: number;
  /** The interactive trigger frame, absent for a main-sequence group. */
  triggerFrameId?: string;
  bars: TimelineBar[];
}

export interface TimelineModel {
  groups: TimelineGroup[];
  totalMs: number;
  widthPx: number;
  ticks: { seconds: number; x: number; label: string }[];
}

/** The timeline of a slide: the main sequence's click groups, then each
 * interactive sequence's, every build step a bar placed by `stepTimeline`. */
export function timelineModel(slide: Slide): TimelineModel {
  const animations = slide.animations ?? [];
  const sequences: { groups: BuildStep[][]; triggerFrameId?: string }[] = [
    { groups: slideBuildSteps(slide) },
    ...slideInteractiveBuilds(slide),
  ];
  const groups: TimelineGroup[] = [];
  let barIndex = 0;
  let totalMs = 0;
  for (const sequence of sequences) {
    for (const steps of sequence.groups) {
      const timeline = stepTimeline(steps);
      totalMs = Math.max(totalMs, timeline.totalMs);
      const bars = timeline.entries.map((entry): TimelineBar => {
        const a = entry.step.animation;
        const bar: TimelineBar = {
          barIndex: barIndex++,
          animationIndex: animations.indexOf(a),
          startMs: entry.startMs,
          durationMs: entry.durationMs,
          delayMs: effectiveAnimationDelayMs(a),
          singleMs: effectiveAnimationDurationMs(a),
          passes: barPasses(a),
          ...barRect(entry.startMs, entry.durationMs),
        };
        if (entry.step.paragraphIndex !== undefined) {
          bar.paragraphIndex = entry.step.paragraphIndex;
        }
        return bar;
      });
      groups.push(
        sequence.triggerFrameId === undefined
          ? { index: groups.length, bars }
          : { index: groups.length, triggerFrameId: sequence.triggerFrameId, bars },
      );
    }
  }
  const ticks = rulerTicks(totalMs);
  const widthPx = Math.max(ticks[ticks.length - 1].x, ...groups.flatMap((g) => g.bars.map((b) => b.left + b.width)));
  return { groups, totalMs, widthPx, ticks };
}
