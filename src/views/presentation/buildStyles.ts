// The build engine's settled states (SPEC-086/105), framework-free: which
// steps of a slide a build state has played — main and interactive groups in
// the order they played — and the per-frame and per-paragraph inline styles
// that state settles to (visibility, plus where played spins and paths left
// their targets). useBuildPreview drives these; the play view's transition
// stage reads `fullyBuiltStyles` for the slide it leaves.
import type { Slide } from "../../lib/presentation/model";
import {
  isBuildTargetHiddenAfter,
  settledMotion,
  slideBuildCount,
  slideBuildSteps,
  slideInteractiveBuilds,
  type BuildStep,
} from "../../lib/presentation/buildPlayback";
import {
  DEFAULT_STYLE_CONTEXT,
  settledTargetStyle,
  type FrameStyle,
  type StyleContext,
} from "../../components/presentation/animationStyles";

export type StyleMap = Record<string, FrameStyle>;
export type ParagraphStyleMap = Record<string, Record<number, FrameStyle>>;

/** One interactive group played: its trigger, its index in that trigger's
 * sequence, and how many main-sequence groups were shown when it played. */
export interface InteractivePlay {
  trigger: string;
  group: number;
  at: number;
}

/** How far a slide has built: main-sequence groups shown, and the interactive
 * groups played, in the order they played. */
export interface BuildState {
  groupsPlayed: number;
  interactive: InteractivePlay[];
}



/** Every step of the slide (main and interactive) and the ones `state` has
 * played, in the order they played — each interactive group after the main
 * groups shown when it played — so the last to play wins. */
export function stepsAt(slide: Slide, state: BuildState): { all: BuildStep[]; played: BuildStep[] } {
  const main = slideBuildSteps(slide);
  const interactive = slideInteractiveBuilds(slide);
  const groupsOf = new Map(interactive.map((s) => [s.triggerFrameId, s.groups]));
  const played: BuildStep[] = [];
  for (let shown = 0; shown <= state.groupsPlayed; shown++) {
    for (const play of state.interactive) {
      if (Math.min(play.at, state.groupsPlayed) === shown) {
        played.push(...(groupsOf.get(play.trigger)?.[play.group] ?? []));
      }
    }
    if (shown < state.groupsPlayed && shown < main.length) {
      played.push(...main[shown]);
    }
  }
  return { all: [...main.flat(), ...interactive.flatMap((s) => s.groups.flat())], played };
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
  return settledBuildStyles(slide, { groupsPlayed, interactive: [] }).frameStyles;
}

/** The styles of `slide` fully built (every main group shown) — what a
 * transition's outgoing stage shows, paragraphs and settled motion included,
 * resolved against the deck's size. */
export function fullyBuiltStyles(
  slide: Slide | null | undefined,
  deck: { slideWidthCm: number; slideHeightCm: number } | null | undefined,
): { frameStyles: StyleMap; paragraphStyles: ParagraphStyleMap } {
  if (!slide) {
    return { frameStyles: {}, paragraphStyles: {} };
  }
  const ctx = deck ? { slideWidthCm: deck.slideWidthCm, slideHeightCm: deck.slideHeightCm } : DEFAULT_STYLE_CONTEXT;
  return settledBuildStyles(slide, { groupsPlayed: slideBuildCount(slide), interactive: [] }, ctx);
}
