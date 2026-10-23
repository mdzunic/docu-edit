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
//
// SPEC-105: the engine drives BUILD STEPS (`slideBuildSteps`) rather than
// rows, so a by-paragraph build reveals one paragraph per step through
// `paragraphStyles` (frame id → paragraph index → inline style), which the
// stage merges into each `slide-paragraph`; it tracks how far each of the
// slide's interactive sequences has played (`playTrigger` — a click on a
// trigger frame plays its next group, the main sequence untouched), keeps an
// until-click repeat looping until the next input, and settles emphasis and
// path effects to where they leave their target.
import { ref, watch } from "vue";
import type { Slide } from "../../lib/presentation/model";
import type { PlaybackStep } from "./usePlayback";
import {
  interactiveTriggerAt,
  isBuildTargetHiddenAfter,
  settledMotion,
  slideBuildCount,
  slideBuildSteps,
  slideInteractiveBuilds,
  stepTimeline,
  type BuildStep,
} from "../../lib/presentation/buildPlayback";
import {
  DEFAULT_STYLE_CONTEXT,
  fromStyle,
  settledTargetStyle,
  toStyle,
  transitionFor,
  type FrameStyle,
  type StyleContext,
} from "../../components/presentation/animationStyles";
import { clientToCm } from "./stageMath";

type StyleMap = Record<string, FrameStyle>;
type ParagraphStyleMap = Record<string, Record<number, FrameStyle>>;

/** How far a slide has built: main-sequence groups shown, and per trigger
 * frame how many of its interactive groups have played. */
interface BuildState {
  groupsPlayed: number;
  interactive: Record<string, number>;
}

/** The deck size a show resolves motion paths against. */
export interface DeckSize {
  widthCm: number;
  heightCm: number;
}

// The settle runs this long after a group's last animation is scheduled to end:
// two frames for the from→to commit (`afterCommitFrame`) plus margin, so the
// settle never cuts a still-running transition short.
const SETTLE_SLACK_MS = 100;

/** Every step of the slide (main and interactive) and the ones `state` has
 * played. */
function stepsAt(slide: Slide, state: BuildState): { all: BuildStep[]; played: BuildStep[] } {
  const main = slideBuildSteps(slide);
  const interactive = slideInteractiveBuilds(slide);
  return {
    all: [...main.flat(), ...interactive.flatMap((s) => s.groups.flat())],
    played: [
      ...main.slice(0, state.groupsPlayed).flat(),
      ...interactive.flatMap((s) => s.groups.slice(0, state.interactive[s.triggerFrameId] ?? 0).flat()),
    ],
  };
}

/** The per-frame and per-paragraph styles of `slide` settled at `state` — the
 * static visibility, with every played spin and path at its end, no
 * animation in flight (AC-20). */
export function settledBuildStyles(
  slide: Slide,
  state: BuildState,
  ctx: StyleContext = DEFAULT_STYLE_CONTEXT,
): { frameStyles: StyleMap; paragraphStyles: ParagraphStyleMap } {
  const { all, played } = stepsAt(slide, state);
  const styleOf = (key: string): FrameStyle =>
    settledTargetStyle(isBuildTargetHiddenAfter(all, played, key), settledMotion(played, key), ctx);
  const frameStyles: StyleMap = {};
  for (const frame of slide.frames) {
    frameStyles[frame.id] = styleOf(frame.id);
  }
  const paragraphStyles: ParagraphStyleMap = {};
  for (const step of all) {
    if (step.paragraphIndex === undefined) {
      continue;
    }
    const style = styleOf(step.key);
    const byIndex = (paragraphStyles[step.frameId] ??= {});
    for (let i = step.paragraphIndex; i <= (step.paragraphEnd ?? step.paragraphIndex); i++) {
      byIndex[i] = style;
    }
  }
  return { frameStyles, paragraphStyles };
}

/** The per-frame styles of `slide` settled with `groupsPlayed` groups shown —
 * the static visibility with no animation in flight (AC-20). */
export function settledFrameStyles(slide: Slide, groupsPlayed: number): StyleMap {
  return settledBuildStyles(slide, { groupsPlayed, interactive: {} }).frameStyles;
}

/** `deckSize` resolves motion paths (slide fractions) to surface px; absent,
 * a 16:9 deck. */
export function useBuildPreview(deckSize?: () => DeckSize | null | undefined) {
  const frameStyles = ref<StyleMap>({});
  const paragraphStyles = ref<ParagraphStyleMap>({});
  const playing = ref(false);
  let timers: number[] = [];
  // Bumped by clearTimers so a frame callback scheduled before a stop, a
  // completion or the next group never lands a stale "to" style after it.
  let generation = 0;
  // The settled state a currently-animating group resolves to (its slide and the
  // build state shown once it finishes). Held so a mid-build input can snap
  // straight to it (AC-24).
  let pending: { slide: Slide; state: BuildState } | null = null;
  // The slide and build state currently shown — interactive sequences keep
  // their progress while the main sequence advances on the same slide.
  let shown: { slideId: string; state: BuildState } | null = null;

  function context(): StyleContext {
    const size = deckSize?.();
    return size ? { slideWidthCm: size.widthCm, slideHeightCm: size.heightCm } : DEFAULT_STYLE_CONTEXT;
  }

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
    const fire = (): void => {
      if (scheduled === generation) cb();
    };
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => requestAnimationFrame(fire));
    } else {
      timers.push(window.setTimeout(fire, 32));
    }
  }

  /** Sets the inline style of a step's target: its frame, or every paragraph
   * the step covers. */
  function setTarget(step: BuildStep, style: FrameStyle): void {
    if (step.paragraphIndex === undefined) {
      frameStyles.value = { ...frameStyles.value, [step.frameId]: style };
      return;
    }
    const byIndex = { ...paragraphStyles.value[step.frameId] };
    for (let i = step.paragraphIndex; i <= (step.paragraphEnd ?? step.paragraphIndex); i++) {
      byIndex[i] = style;
    }
    paragraphStyles.value = { ...paragraphStyles.value, [step.frameId]: byIndex };
  }

  function settleState(slide: Slide, state: BuildState): void {
    const settled = settledBuildStyles(slide, state, context());
    frameStyles.value = settled.frameStyles;
    paragraphStyles.value = settled.paragraphStyles;
    shown = { slideId: slide.id, state };
  }

  /** The interactive progress to keep when settling `slide`: this slide's,
   * or none for a slide just arrived at. */
  function interactiveOf(slide: Slide): Record<string, number> {
    return shown && shown.slideId === slide.id ? shown.state.interactive : {};
  }

  /** Renders the slide settled with `groupsPlayed` groups shown — the static
   * visibility with no animation in flight (AC-20). A slide just arrived at
   * starts with none of its interactive sequences played. */
  function settle(slide: Slide, groupsPlayed: number): void {
    settleState(slide, { groupsPlayed, interactive: interactiveOf(slide) });
  }

  /** Animates `steps` from build state `before` to `after`: settles instantly
   * to `before` (completing any prior in-flight group), then plays each step
   * on its scheduled offset, settling to `after` at the end — an until-click
   * repeat still looping. Reduced motion skips straight to `after`, every
   * emphasis, repeat and path at its end state (AC-23). */
  function run(slide: Slide, steps: BuildStep[], before: BuildState, after: BuildState, reducedMotion: boolean, onDone?: () => void): void {
    clearTimers();
    playing.value = false;
    pending = null;
    if (reducedMotion) {
      settleState(slide, after);
      onDone?.();
      return;
    }
    settleState(slide, before);
    playing.value = true;
    pending = { slide, state: after };
    const ctx = context();
    const { played } = stepsAt(slide, before);
    const { entries, totalMs } = stepTimeline(steps);
    for (const entry of entries) {
      const step = entry.step;
      const a = step.animation;
      const stepCtx = { ...ctx, base: settledMotion(played, step.key) };
      timers.push(
        window.setTimeout(() => {
          setTarget(step, { ...fromStyle(a, stepCtx), transition: "none" });
          afterCommitFrame(() => {
            setTarget(step, a.effect === "appear" ? toStyle(a, stepCtx) : { ...toStyle(a, stepCtx), transition: transitionFor(a) });
          });
        }, entry.startMs),
      );
    }
    timers.push(
      window.setTimeout(() => {
        settleState(slide, after);
        // A repeat-until-click keeps looping past the group's end, until the
        // next input settles or plays anything else.
        for (const step of steps) {
          if (step.animation.repeatUntilClick) {
            setTarget(step, toStyle(step.animation, { ...ctx, base: settledMotion(played, step.key) }));
          }
        }
        playing.value = false;
        pending = null;
        onDone?.();
      }, totalMs + SETTLE_SLACK_MS),
    );
  }

  /** Animates build group `groupIndex` (0-based) of the main sequence:
   * settles instantly to the state before it (completing any prior in-flight
   * group), then plays each step on its scheduled offset, settling to
   * `groupIndex + 1` at the end. Reduced motion skips straight to the settled
   * after-state (AC-23). */
  function playGroup(slide: Slide, groupIndex: number, reducedMotion: boolean, onDone?: () => void): void {
    const groups = slideBuildSteps(slide);
    const interactive = interactiveOf(slide);
    if (groupIndex < 0 || groupIndex >= groups.length) {
      clearTimers();
      playing.value = false;
      pending = null;
      settle(slide, Math.max(0, groupIndex));
      onDone?.();
      return;
    }
    run(
      slide,
      groups[groupIndex],
      { groupsPlayed: groupIndex, interactive },
      { groupsPlayed: groupIndex + 1, interactive },
      reducedMotion,
      onDone,
    );
  }

  /** Plays the next group of the interactive sequence `triggerFrameId` starts
   * (SPEC-105) — after its last, a further click replays it from the first —
   * leaving the main sequence where it is. Completes any in-flight group
   * first. Returns whether the frame triggers anything on this slide. */
  function playTrigger(slide: Slide, triggerFrameId: string, reducedMotion: boolean): boolean {
    const sequence = slideInteractiveBuilds(slide).find((s) => s.triggerFrameId === triggerFrameId);
    if (!sequence || sequence.groups.length === 0) {
      return false;
    }
    complete();
    const current = shown && shown.slideId === slide.id ? shown.state : { groupsPlayed: 0, interactive: {} };
    const played = current.interactive[triggerFrameId] ?? 0;
    const index = played >= sequence.groups.length ? 0 : played;
    run(
      slide,
      sequence.groups[index],
      { groupsPlayed: current.groupsPlayed, interactive: { ...current.interactive, [triggerFrameId]: index } },
      { groupsPlayed: current.groupsPlayed, interactive: { ...current.interactive, [triggerFrameId]: index + 1 } },
      reducedMotion,
    );
    return true;
  }

  /** A show's pointer-down: when it lands on a visible trigger frame of the
   * slide (hit-tested in slide cm through the stage under the pointer), plays
   * that frame's interactive sequence and returns true — the caller then
   * consumes the input instead of advancing. */
  function playTriggerAt(slide: Slide | null, event: PointerEvent, reducedMotion: boolean): boolean {
    const target = event.target;
    const stage = target instanceof Element ? target.closest('[data-testid="slide-stage"]') : null;
    if (!slide || !stage) {
      return false;
    }
    const ctx = context();
    const { xCm, yCm } = clientToCm(event.clientX, event.clientY, stage.getBoundingClientRect(), ctx.slideWidthCm, ctx.slideHeightCm);
    const frameId = interactiveTriggerAt(slide, xCm, yCm);
    if (!frameId || frameStyles.value[frameId]?.visibility === "hidden") {
      return false;
    }
    return playTrigger(slide, frameId, reducedMotion);
  }

  /** Snap an in-flight build group to its completed state and hold there
   * (AC-24): a play-time input arriving mid-build finishes the current group
   * instantly rather than walking on to the next group or slide. Returns whether
   * a group was actually in flight — the caller consumes the input only then. */
  function complete(): boolean {
    if (!playing.value || !pending) return false;
    const { slide, state } = pending;
    clearTimers();
    playing.value = false;
    pending = null;
    settleState(slide, state);
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
    shown = null;
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

  return {
    frameStyles,
    paragraphStyles,
    playing,
    settle,
    playGroup,
    playTrigger,
    playTriggerAt,
    complete,
    autoplay,
    stop,
  };
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
