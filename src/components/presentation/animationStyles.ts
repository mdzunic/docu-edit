// The CSS mapping for animation playback (SPEC-086) — the visual counterpart of
// the wire mapping in lib/presentation. Each modelled effect becomes a
// from/to inline-style pair plus a `transition` shorthand the play engine
// toggles across two frames (set `from` with no transition, then `to` with it),
// honouring duration and delay. Kept beside the renderer (like
// transitionStyles.ts) because it is a pure rendering concern; framework-free,
// unit-tested without a mount.
//
// SPEC-105's emphasis and path effects animate through the `de-anim-*`
// keyframes App.vue declares globally (an inline style cannot carry
// @keyframes): their "to" style names the keyframes and feeds them through
// custom properties — a scale, a rotation span, a colour, or the sampled
// points of a motion path resolved at the deck's cm size — and carries
// `animationIterationCount`/`animationDirection` for repetition and
// auto-reverse. They move the frame through the individual `scale`/`rotate`/
// `translate` properties, which compose with a rotated frame's own
// `transform` instead of replacing it.
import type { SlideAnimation, TransitionDirection } from "../../lib/presentation/model";
import {
  effectiveAnimationDelayMs,
  effectiveAnimationDirection,
  effectiveAnimationDurationMs,
  effectiveEmphasisColor,
  effectiveRepeatCount,
  effectiveScalePercent,
  effectiveSpinDegrees,
  effectClass,
  isEntranceEffect,
} from "../../lib/presentation/animations";
import { resolveMotionPath, samplePath, type PathPoint } from "../../lib/presentation/motionPaths";
import { cmToPx } from "../../lib/core/units";

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

/** The deck a path resolves against (slide fractions → surface px) and where
 * the target's earlier spins and paths left it (the start of its next one). */
export interface StyleContext {
  slideWidthCm: number;
  slideHeightCm: number;
  base?: { rotateDeg: number; offset: PathPoint };
}

/** A 16:9 deck — the context when a caller has none. */
export const DEFAULT_STYLE_CONTEXT: StyleContext = { slideWidthCm: 33.867, slideHeightCm: 19.05 };

/** A slide-fraction offset as a CSS `translate` value in surface px. */
function translateValue(offset: PathPoint, ctx: StyleContext): string {
  const x = offset.x * cmToPx(ctx.slideWidthCm);
  const y = offset.y * cmToPx(ctx.slideHeightCm);
  return `${Number(x.toFixed(2))}px ${Number(y.toFixed(2))}px`;
}

/** The settled style of a build target (SPEC-105): hidden or not, plus where
 * its played spins and paths left it — the end state reduced motion jumps to. */
export function settledTargetStyle(
  hidden: boolean,
  motion: { rotateDeg: number; offset: PathPoint },
  ctx: StyleContext = DEFAULT_STYLE_CONTEXT,
): FrameStyle {
  const style = settledStyle(hidden);
  if (motion.rotateDeg !== 0) {
    style.rotate = `${motion.rotateDeg}deg`;
  }
  if (motion.offset.x !== 0 || motion.offset.y !== 0) {
    style.translate = translateValue(motion.offset, ctx);
  }
  return style;
}

/** How many CSS iterations an emphasis/path effect runs: its passes, doubled
 * when it auto-reverses (each pass forward then back), or "infinite" until
 * the next click. */
export function animationIterationCount(a: SlideAnimation): string {
  if (a.repeatUntilClick) {
    return "infinite";
  }
  return String(effectiveRepeatCount(a) * (a.autoReverse ? 2 : 1));
}

/** The CSS animation direction: "alternate" for an auto-reversing effect. */
export function animationDirection(a: SlideAnimation): string {
  return a.autoReverse ? "alternate" : "normal";
}

/** The number of segments a motion path is sampled into for its keyframes
 * (App.vue's `de-anim-path` declares one stop per point). */
export const PATH_SAMPLES = 16;

/** The keyframe-driven "to" style of an emphasis/path effect: the
 * `de-anim-*` keyframes, their custom properties, and the repetition. A
 * rewinding effect does not hold its last keyframe. */
function keyframeStyle(a: SlideAnimation, ctx: StyleContext): FrameStyle {
  const dur = effectiveAnimationDurationMs(a);
  const base = ctx.base ?? { rotateDeg: 0, offset: { x: 0, y: 0 } };
  const style: FrameStyle = {
    animationDuration: `${dur}ms`,
    animationTimingFunction: "ease-in-out",
    animationIterationCount: animationIterationCount(a),
    animationDirection: animationDirection(a),
    animationFillMode: a.rewindOnComplete ? "none" : "forwards",
  };
  switch (a.effect) {
    case "pulse":
    case "growShrink":
      style.animationName = "de-anim-scale";
      style["--de-anim-scale"] = String(a.effect === "pulse" ? 1.1 : effectiveScalePercent(a) / 100);
      break;
    case "spin":
      style.animationName = "de-anim-spin";
      style["--de-anim-rotate-from"] = `${base.rotateDeg}deg`;
      style["--de-anim-rotate-to"] = `${base.rotateDeg + effectiveSpinDegrees(a)}deg`;
      break;
    case "colorChange":
      style.animationName = "de-anim-color";
      style["--de-anim-color"] = effectiveEmphasisColor(a);
      break;
    case "motionPath": {
      style.animationName = "de-anim-path";
      const commands = a.path ? resolveMotionPath(a.path) : undefined;
      const points = commands ? samplePath(commands, PATH_SAMPLES) : [];
      points.forEach((p, i) => {
        style[`--de-path-${i}`] = translateValue({ x: base.offset.x + p.x, y: base.offset.y + p.y }, ctx);
      });
      break;
    }
  }
  return style;
}

/** Whether an effect animates through keyframes (emphasis and path) — its
 * styles change no visibility. */
export function keyframed(a: SlideAnimation): boolean {
  const cls = effectClass(a.effect);
  return cls === "emphasis" || cls === "path";
}

/** The pre-animation ("from") style for an effect. Emphasis and path effects
 * start where their target rests, their keyframes not yet running — and
 * name no visibility: they leave a hidden target hidden (the engine layers
 * them over the target's current style). */
export function fromStyle(a: SlideAnimation, ctx: StyleContext = DEFAULT_STYLE_CONTEXT): FrameStyle {
  const direction = effectiveAnimationDirection(a) ?? "left";
  if (keyframed(a)) {
    const base = ctx.base;
    const style: FrameStyle = { animationName: "none" };
    if (base && base.rotateDeg !== 0) style.rotate = `${base.rotateDeg}deg`;
    if (base && (base.offset.x !== 0 || base.offset.y !== 0)) style.translate = translateValue(base.offset, ctx);
    return style;
  }
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

/** The post-animation ("to") style for an effect: for an emphasis or path
 * effect, its keyframes running (pulse/grow-shrink `scale`, spin `rotate`,
 * colour change `color`/`background-color`, motion path `translate` through
 * the path resolved at the deck's cm size). */
export function toStyle(a: SlideAnimation, ctx: StyleContext = DEFAULT_STYLE_CONTEXT): FrameStyle {
  const direction = effectiveAnimationDirection(a) ?? "left";
  if (keyframed(a)) {
    return keyframeStyle(a, ctx);
  }
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

/** The CSS `transition` shorthand for an effect's animated properties —
 * "none" for an emphasis or path effect, whose keyframes carry the motion. */
export function transitionFor(a: SlideAnimation): string {
  if (keyframed(a)) {
    return "none";
  }
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
