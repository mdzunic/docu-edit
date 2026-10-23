// SPEC-105 animation styles: the five new effects' from/to/transition cases
// (keyframes through custom properties), repetition and auto-reverse, and the
// settled end state of spins and paths.
import { describe, expect, it } from "vitest";
import type { SlideAnimation } from "../../lib/presentation/model";
import {
  animationDirection,
  animationIterationCount,
  fromStyle,
  PATH_SAMPLES,
  settledTargetStyle,
  toStyle,
  transitionFor,
} from "./animationStyles";

const ctx = { slideWidthCm: 25.4, slideHeightCm: 12.7 }; // 960 × 480 px
const row = (over: Partial<SlideAnimation>): SlideAnimation => ({ frameId: "f", effect: "spin", trigger: "click", ...over });

describe("SPEC-105 effect styles", () => {
  it("pulse and grow/shrink throb a scale and back", () => {
    expect(toStyle(row({ effect: "pulse" }), ctx)).toMatchObject({ animationName: "de-anim-scale", "--de-anim-scale": "1.1" });
    expect(toStyle(row({ effect: "growShrink", scalePercent: 250 }), ctx)).toMatchObject({ "--de-anim-scale": "2.5" });
  });

  it("spin rotates from where the target rests by its angle", () => {
    const style = toStyle(row({ spinDegrees: 90 }), { ...ctx, base: { rotateDeg: 45, offset: { x: 0, y: 0 } } });
    expect(style).toMatchObject({ animationName: "de-anim-spin", "--de-anim-rotate-from": "45deg", "--de-anim-rotate-to": "135deg" });
    expect(fromStyle(row({}), { ...ctx, base: { rotateDeg: 45, offset: { x: 0, y: 0 } } })).toMatchObject({ rotate: "45deg", animationName: "none" });
  });

  it("colour change goes to the colour and back", () => {
    expect(toStyle(row({ effect: "colorChange", emphasisColor: "#00aa00" }), ctx)).toMatchObject({
      animationName: "de-anim-color",
      "--de-anim-color": "#00aa00",
    });
  });

  it("a motion path translates through its points resolved at the deck's cm size", () => {
    const style = toStyle(row({ effect: "motionPath", path: { preset: "line" } }), ctx);
    expect(style.animationName).toBe("de-anim-path");
    expect(style["--de-path-0"]).toBe("0px 0px");
    expect(style[`--de-path-${PATH_SAMPLES}`]).toBe("240px 0px");
  });

  it("repeats and auto-reverses through the iteration count and direction; rewind does not hold", () => {
    expect([animationIterationCount(row({})), animationDirection(row({}))]).toEqual(["1", "normal"]);
    expect([animationIterationCount(row({ repeatCount: 3, autoReverse: true })), animationDirection(row({ autoReverse: true }))]).toEqual([
      "6",
      "alternate",
    ]);
    expect(animationIterationCount(row({ repeatUntilClick: true }))).toBe("infinite");
    expect(toStyle(row({ rewindOnComplete: true }), ctx).animationFillMode).toBe("none");
    expect(toStyle(row({}), ctx).animationFillMode).toBe("forwards");
    expect(transitionFor(row({}))).toBe("none");
    // SPEC-086's transitions are untouched.
    expect(transitionFor(row({ effect: "fadeIn", durationMs: 300 }))).toBe("opacity 300ms ease");
  });

  it("settles a spun and moved target at its end state", () => {
    expect(settledTargetStyle(false, { rotateDeg: 360, offset: { x: 0.25, y: -0.5 } }, ctx)).toEqual({
      rotate: "360deg",
      translate: "240px -240px",
    });
    expect(settledTargetStyle(true, { rotateDeg: 0, offset: { x: 0, y: 0 } }, ctx)).toEqual({ opacity: "0", visibility: "hidden" });
  });
});
