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
import { groupTimeline, isFrameHiddenAt, slideBuildGroups } from "../../lib/presentation/buildPlayback";
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
  let timers: number[] = [];

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
        onDone?.();
      }, totalMs + 40),
    );
  }

  function stop(): void {
    clearTimers();
  }

  return { frameStyles, settle, playGroup, stop };
}
