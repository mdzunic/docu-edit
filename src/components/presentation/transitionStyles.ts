// The CSS a two-stage transition applies to its outgoing/incoming stages
// (SPEC-085) — kept framework-free so both the play view and the panel preview
// drive the same visuals, and so the geometry is unit-testable without mounting.
// Each transition renders as two stacked stages inside the letterbox; a class-
// free inline `transform`/`opacity`/`clip-path` animates from the "enter" state
// to the "active" state over the model duration.
import type { Slide, TransitionDirection, TransitionType } from "../../lib/presentation/model";
import {
  DEFAULT_TRANSITION_DURATION_MS,
  MAX_TRANSITION_DURATION_MS,
  MIN_TRANSITION_DURATION_MS,
  defaultDirectionFor,
  type SlideTransition,
} from "../../lib/presentation/transitions";

export type TransitionPhase = "enter" | "active";

export interface LayerStyle {
  transform?: string;
  opacity?: string;
  clipPath?: string;
}

export interface StagePair {
  outgoing: LayerStyle;
  incoming: LayerStyle;
  /** Whether the incoming stage stacks above the outgoing one. Uncover is the
   * one effect where the OUTGOING stage moves over a static incoming. */
  incomingOnTop: boolean;
}

/** The axis translate a directional slide uses to enter (incoming) — the
 * off-screen start position for its direction. */
function enterTranslate(direction: TransitionDirection): string {
  switch (direction) {
    case "left":
      return "translateX(100%)"; // enters from the right edge, travels left
    case "right":
      return "translateX(-100%)";
    case "up":
      return "translateY(100%)";
    case "down":
      return "translateY(-100%)";
  }
}

/** The axis translate a directional slide uses to exit (outgoing) in the same
 * direction as the incoming content travels. */
function exitTranslate(direction: TransitionDirection): string {
  switch (direction) {
    case "left":
      return "translateX(-100%)";
    case "right":
      return "translateX(100%)";
    case "up":
      return "translateY(-100%)";
    case "down":
      return "translateY(100%)";
  }
}

/** The collapsed clip-path a wipe starts from for its direction (revealed edge
 * grows opposite the travel direction). */
function wipeEnterClip(direction: TransitionDirection): string {
  switch (direction) {
    case "left":
      return "inset(0 0 0 100%)"; // reveal grows from the right edge leftward
    case "right":
      return "inset(0 100% 0 0)";
    case "up":
      return "inset(100% 0 0 0)";
    case "down":
      return "inset(0 0 100% 0)";
  }
}

const CLIP_OPEN = "inset(0 0 0 0)";

/** The two-stage styles for a transition type/direction at a given phase.
 * Returns static (empty) styles for `none`. `phase` is widened to `string` so
 * the .vue component (whose reactive refs are untyped plain JS) can call it. */
export function transitionStagePair(
  type: TransitionType | "none",
  direction: TransitionDirection,
  phase: string,
): StagePair {
  const enter = phase === "enter";
  switch (type) {
    case "fade":
      return { outgoing: {}, incoming: { opacity: enter ? "0" : "1" }, incomingOnTop: true };
    case "push":
      return {
        outgoing: { transform: enter ? "translateX(0)" : exitTranslate(direction) },
        incoming: { transform: enter ? enterTranslate(direction) : "translateX(0)" },
        incomingOnTop: true,
      };
    case "cover":
      return {
        outgoing: {},
        incoming: { transform: enter ? enterTranslate(direction) : "translateX(0)" },
        incomingOnTop: true,
      };
    case "uncover":
      return {
        outgoing: { transform: enter ? "translateX(0)" : exitTranslate(direction) },
        incoming: {},
        incomingOnTop: false,
      };
    case "wipe":
      return {
        outgoing: {},
        incoming: { clipPath: enter ? wipeEnterClip(direction) : CLIP_OPEN },
        incomingOnTop: true,
      };
    case "split":
      return {
        outgoing: {},
        incoming: { clipPath: enter ? "inset(0 50% 0 50%)" : CLIP_OPEN },
        incomingOnTop: true,
      };
    default:
      return { outgoing: {}, incoming: {}, incomingOnTop: true };
  }
}

/** Whether a transition should animate (vs. cut instantly): a real effect, a
 * previous slide to animate from, and motion allowed. */
export function transitionAnimates(
  transition: SlideTransition | null | undefined,
  reducedMotion: boolean,
  hasOutgoing: boolean,
): boolean {
  return !!transition && transition.type !== "none" && !reducedMotion && hasOutgoing;
}

/** The effective direction for a transition, defaulting per type. Loosely
 * typed so an untyped .vue prop object satisfies it. */
export function transitionDirection(transition: { type?: string; direction?: TransitionDirection }): TransitionDirection {
  return transition.direction ?? defaultDirectionFor(transition.type ?? "none") ?? "left";
}

/** The effective duration in ms, defaulted and clamped — loosely typed for the
 * .vue component's untyped transition prop. */
export function transitionDurationMs(transition: { durationMs?: number }): number {
  const ms = transition.durationMs ?? DEFAULT_TRANSITION_DURATION_MS;
  return Math.min(MAX_TRANSITION_DURATION_MS, Math.max(MIN_TRANSITION_DURATION_MS, ms));
}

/** The play view's two-stage transition state (outgoing + incoming + the model
 * transition), or null when no transition is animating. */
export interface TwoStageView {
  outgoing: Slide | null;
  incoming: Slide;
  transition: SlideTransition;
}

/** A typed `null` seed for the play view's `shallowRef`, so assigning a
 * `TwoStageView` later type-checks without TS syntax in the .vue script. */
export function emptyTwoStageView(): TwoStageView | null {
  return null;
}
