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
import { ref } from "vue";
import type { Slide } from "../../lib/presentation/model";
import { groupTimeline, isFrameHiddenAt, slideBuildCount, slideBuildGroups } from "../../lib/presentation/buildPlayback";
import { fromStyle, settledStyle, toStyle, transitionFor } from "../../components/presentation/animationStyles";

type StyleMap = Record<string, Record<string, string>>;

function nextFrame(cb: () => void): void {
  if (typeof requestAnimationFrame === "function") {
    requestAnimationFrame(() => cb());
  } else {
    setTimeout(cb, 16);
  }
}

export function useBuildPreview() {
  const frameStyles = ref<StyleMap>({});
  const playing = ref(false);
  let timers: number[] = [];
  // The settled state a currently-animating group resolves to (its slide and the
  // group count shown once it finishes). Held so a mid-build input can snap
  // straight to it (AC-24).
  let pending: { slide: Slide; groupsPlayed: number } | null = null;

  function clearTimers(): void {
    for (const t of timers) {
      clearTimeout(t);
    }
    timers = [];
  }

  function setFrame(frameId: string, style: Record<string, string>): void {
    frameStyles.value = { ...frameStyles.value, [frameId]: style };
  }

  /** Renders the slide settled with `groupsPlayed` groups shown — the static
   * visibility with no animation in flight (AC-20). */
  function settle(slide: Slide, groupsPlayed: number): void {
    const groups = slideBuildGroups(slide);
    const styles: StyleMap = {};
    for (const frame of slide.frames) {
      styles[frame.id] = settledStyle(isFrameHiddenAt(groups, frame.id, groupsPlayed));
    }
    frameStyles.value = styles;
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
          nextFrame(() => {
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
      }, totalMs + 40),
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
