// The CSS mapping for animation playback (SPEC-086) — the visual counterpart of
// the wire mapping in lib/presentation. Each modelled effect becomes a
// from/to inline-style pair plus a `transition` shorthand the play engine
// toggles across two frames (set `from` with no transition, then `to` with it),
// honouring duration and delay. Kept beside the renderer (like
// transitionStyles.ts) because it is a pure rendering concern; framework-free,
// unit-tested without a mount.
import type { SlideAnimation, TransitionDirection } from "../../lib/presentation/model";
import {
  effectiveAnimationDelayMs,
  effectiveAnimationDirection,
  effectiveAnimationDurationMs,
  isEntranceEffect,
} from "../../lib/presentation/animations";

export type FrameStyle = Record<string, string>;

/** How far off-frame a fly starts/ends, per originate-edge direction. */
const FLY_OFFSET: Record<TransitionDirection, string> = {
  left: "translateX(-120%)",
  right: "translateX(120%)",
  up: "translateY(-120%)",
  down: "translateY(120%)",
};

/** The fully-hidden clip for a wipe, per reveal-origin direction (the reveal
 * grows from that edge to `inset(0)`). */
const WIPE_HIDDEN: Record<TransitionDirection, string> = {
  left: "inset(0 100% 0 0)",
  right: "inset(0 0 0 100%)",
  up: "inset(0 0 100% 0)",
  down: "inset(100% 0 0 0)",
};

/** The settled style for a frame between builds: hidden frames are removed from
 * view (opacity + visibility), visible frames carry no override. */
export function settledStyle(hidden: boolean): FrameStyle {
  return hidden ? { opacity: "0", visibility: "hidden" } : {};
}

/** The pre-animation ("from") style for an effect. */
export function fromStyle(a: SlideAnimation): FrameStyle {
  const direction = effectiveAnimationDirection(a) ?? "left";
  switch (a.effect) {
    case "appear":
      return { visibility: "visible" };
    case "fadeIn":
      return { visibility: "visible", opacity: "0" };
    case "fadeOut":
      return { visibility: "visible", opacity: "1" };
    case "flyIn":
      return { visibility: "visible", opacity: "1", transform: FLY_OFFSET[direction] };
    case "flyOut":
      return { visibility: "visible", opacity: "1", transform: "none" };
    case "wipeIn":
      return { visibility: "visible", opacity: "1", clipPath: WIPE_HIDDEN[direction] };
    case "zoomIn":
      return { visibility: "visible", opacity: "0", transform: "scale(0.2)" };
    default:
      return { visibility: "visible" };
  }
}

/** The post-animation ("to") style for an effect. */
export function toStyle(a: SlideAnimation): FrameStyle {
  const direction = effectiveAnimationDirection(a) ?? "left";
  switch (a.effect) {
    case "appear":
      return { visibility: "visible" };
    case "fadeIn":
      return { visibility: "visible", opacity: "1" };
    case "fadeOut":
      return { visibility: "visible", opacity: "0" };
    case "flyIn":
      return { visibility: "visible", opacity: "1", transform: "none" };
    case "flyOut":
      return { visibility: "visible", opacity: "0", transform: FLY_OFFSET[direction] };
    case "wipeIn":
      return { visibility: "visible", opacity: "1", clipPath: "inset(0)" };
    case "zoomIn":
      return { visibility: "visible", opacity: "1", transform: "none" };
    default:
      return { visibility: "visible" };
  }
}

/** The CSS `transition` shorthand for an effect's animated properties. */
export function transitionFor(a: SlideAnimation): string {
  const dur = effectiveAnimationDurationMs(a);
  const props =
    a.effect === "fadeIn" || a.effect === "fadeOut"
      ? ["opacity"]
      : a.effect === "wipeIn"
        ? ["clip-path"]
        : a.effect === "zoomIn"
          ? ["opacity", "transform"]
          : ["transform", "opacity"];
  return props.map((p) => `${p} ${dur}ms ease`).join(", ");
}

export { effectiveAnimationDelayMs, effectiveAnimationDurationMs, isEntranceEffect };
