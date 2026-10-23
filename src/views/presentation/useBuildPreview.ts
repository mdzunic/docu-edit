// The build playback engine (SPEC-086) — the shared driver behind the play
// view's show and the editor "Play" preview. It owns a reactive `frameStyles`
// map (frame id → inline style) that renders the settled state between builds
// and the in-flight animation of a group. The timing/visibility maths is pure
// (lib/presentation/buildPlayback.ts) and the CSS is pure
// (components/presentation/animationStyles.ts); this composable is the
// imperative glue: schedule each animation in a group, cross-fade from→to over
// two frames, and settle when the group ends. Cancelling in-flight timers and
// settling to the target is how "input mid-build completes the group instantly"
// (AC-24) works.
import { ref, watch } from "vue";
import type { Slide } from "../../lib/presentation/model";
import type { PlaybackStep } from "./usePlayback";
import { groupTimeline, isFrameHiddenAt, slideBuildCount, slideBuildGroups } from "../../lib/presentation/buildPlayback";
import { fromStyle, settledStyle, toStyle, transitionFor } from "../../components/presentation/animationStyles";

type StyleMap = Record<string, Record<string, string>>;

// The settle runs this long after a group's last animation is scheduled to end:
// two frames for the from→to commit (`afterCommitFrame`) plus margin, so the
// settle never cuts a still-running transition short.
const SETTLE_SLACK_MS = 100;

/** The per-frame styles of `slide` settled with `groupsPlayed` groups shown —
 * the static visibility with no animation in flight (AC-20). */
export function settledFrameStyles(slide: Slide, groupsPlayed: number): StyleMap {
  const groups = slideBuildGroups(slide);
  const styles: StyleMap = {};
  for (const frame of slide.frames) {
    styles[frame.id] = settledStyle(isFrameHiddenAt(groups, frame.id, groupsPlayed));
  }
  return styles;
}

export function useBuildPreview() {
  const frameStyles = ref<StyleMap>({});
  const playing = ref(false);
  let timers: number[] = [];
  // Bumped by clearTimers so a frame callback scheduled before a stop, a
  // completion or the next group never lands a stale "to" style after it.
  let generation = 0;
  // The settled state a currently-animating group resolves to (its slide and the
  // group count shown once it finishes). Held so a mid-build input can snap
  // straight to it (AC-24).
  let pending: { slide: Slide; groupsPlayed: number } | null = null;

  function clearTimers(): void {
    for (const t of timers) {
      clearTimeout(t);
    }
    timers = [];
    generation += 1;
  }

  /** Runs `cb` once the "from" style just set has been rendered: two frames,
   * like SlideTransitionLayer. One is not enough — a rAF callback runs BEFORE
   * its frame's style recalc, so from and to would land in the same recalc and
   * the transition would start from the settled hidden style instead (flyIn and
   * zoomIn degrade to a fade, wipeIn to an appear) unless something else, such
   * as a mouse move's hit test, happened to flush style in between. */
  function afterCommitFrame(cb: () => void): void {
    const scheduled = generation;
    const run = (): void => {
      if (scheduled === generation) cb();
    };
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => requestAnimationFrame(run));
    } else {
      timers.push(window.setTimeout(run, 32));
    }
  }

  function setFrame(frameId: string, style: Record<string, string>): void {
    frameStyles.value = { ...frameStyles.value, [frameId]: style };
  }

  /** Renders the slide settled with `groupsPlayed` groups shown — the static
   * visibility with no animation in flight (AC-20). */
  function settle(slide: Slide, groupsPlayed: number): void {
    frameStyles.value = settledFrameStyles(slide, groupsPlayed);
  }

  /** Animates build group `groupIndex` (0-based): settles instantly to the
   * state before it (completing any prior in-flight group), then plays each
   * animation on its scheduled offset, settling to `groupIndex + 1` at the end.
   * Reduced motion skips straight to the settled after-state (AC-23). */
  function playGroup(slide: Slide, groupIndex: number, reducedMotion: boolean, onDone?: () => void): void {
    clearTimers();
    playing.value = false;
    pending = null;
    const groups = slideBuildGroups(slide);
    if (groupIndex < 0 || groupIndex >= groups.length) {
      settle(slide, Math.max(0, groupIndex));
      onDone?.();
      return;
    }
    if (reducedMotion) {
      settle(slide, groupIndex + 1);
      onDone?.();
      return;
    }
    settle(slide, groupIndex);
    playing.value = true;
    pending = { slide, groupsPlayed: groupIndex + 1 };
    const { entries, totalMs } = groupTimeline(groups[groupIndex]);
    for (const entry of entries) {
      const a = entry.animation;
      timers.push(
        window.setTimeout(() => {
          setFrame(a.frameId, { ...fromStyle(a), transition: "none" });
          afterCommitFrame(() => {
            setFrame(a.frameId, a.effect === "appear" ? toStyle(a) : { ...toStyle(a), transition: transitionFor(a) });
          });
        }, entry.startMs),
      );
    }
    timers.push(
      window.setTimeout(() => {
        settle(slide, groupIndex + 1);
        playing.value = false;
        pending = null;
        onDone?.();
      }, totalMs + SETTLE_SLACK_MS),
    );
  }

  /** Snap an in-flight build group to its completed state and hold there
   * (AC-24): a play-time input arriving mid-build finishes the current group
   * instantly rather than walking on to the next group or slide. Returns whether
   * a group was actually in flight — the caller consumes the input only then. */
  function complete(): boolean {
    if (!playing.value || !pending) return false;
    const { slide, groupsPlayed } = pending;
    clearTimers();
    playing.value = false;
    pending = null;
    settle(slide, groupsPlayed);
    return true;
  }

  function stop(): void {
    clearTimers();
    playing.value = false;
    pending = null;
  }

  /** Auto-plays every build group of a slide in sequence (the editor "Play"
   * preview, AC-11): hides unbuilt frames, then plays each group with a short
   * pause between, and calls `onFinished` after the last group settles. */
  function autoplay(slide: Slide, reducedMotion: boolean, onFinished: () => void): void {
    clearTimers();
    settle(slide, 0);
    const groups = slideBuildCount(slide);
    let index = 0;
    const step = (): void => {
      if (index >= groups) {
        timers.push(window.setTimeout(onFinished, 800));
        return;
      }
      const groupIndex = index;
      index += 1;
      playGroup(slide, groupIndex, reducedMotion, () => {
        timers.push(window.setTimeout(step, 500));
      });
    };
    timers.push(window.setTimeout(step, 400));
  }

  return { frameStyles, playing, settle, playGroup, complete, autoplay, stop };
}

/** The show's step → build wiring (SPEC-086), shared by the play view and the
 * presenter's current-slide pane (SPEC-100) so the two render the very same
 * build state for the same step: a new slide, or a backward/jump landing,
 * settles to the target instantly (back lands fully built — AC-22); a forward
 * build step on the same slide animates the group that just began. */
export function watchBuildSteps(
  preview: ReturnType<typeof useBuildPreview>,
  step: () => PlaybackStep,
  slide: () => Slide | null,
  reducedMotion: () => boolean,
): void {
  watch(step, (cur, prev) => {
    const current = slide();
    if (!current) return;
    preview.stop();
    if (!prev || cur.slideIndex !== prev.slideIndex || cur.buildIndex <= prev.buildIndex) {
      preview.settle(current, cur.buildIndex);
      return;
    }
    preview.playGroup(current, cur.buildIndex - 1, reducedMotion());
  });
}
