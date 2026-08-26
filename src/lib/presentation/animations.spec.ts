import { describe, expect, it } from "vitest";
import type { Slide, SlideAnimation } from "./model";
import {
  buildGroups,
  defaultAnimationDirection,
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  effectiveTrigger,
  hasPreservedTiming,
  isEntranceEffect,
  removeAnimationsForFrame,
  retargetAnimations,
  usesAnimationDirection,
} from "./animations";

function anim(frameId: string, over: Partial<SlideAnimation> = {}): SlideAnimation {
  return { frameId, effect: "fadeIn", trigger: "click", ...over };
}

describe("animation helpers", () => {
  it("classifies entrance vs exit effects", () => {
    expect(isEntranceEffect("fadeIn")).toBe(true);
    expect(isEntranceEffect("appear")).toBe(true);
    expect(isEntranceEffect("fadeOut")).toBe(false);
    expect(isEntranceEffect("flyOut")).toBe(false);
  });

  it("reports directional effects and their defaults", () => {
    expect(usesAnimationDirection("flyIn")).toBe(true);
    expect(usesAnimationDirection("wipeIn")).toBe(true);
    expect(usesAnimationDirection("flyOut")).toBe(true);
    expect(usesAnimationDirection("fadeIn")).toBe(false);
    expect(defaultAnimationDirection("flyIn")).toBe("left");
    expect(defaultAnimationDirection("fadeIn")).toBeUndefined();
  });

  it("clamps and defaults duration/delay", () => {
    expect(effectiveAnimationDurationMs(anim("a"))).toBe(500);
    expect(effectiveAnimationDurationMs(anim("a", { durationMs: 50 }))).toBe(100);
    expect(effectiveAnimationDurationMs(anim("a", { durationMs: 9000 }))).toBe(5000);
    expect(effectiveAnimationDelayMs(anim("a"))).toBe(0);
    expect(effectiveAnimationDelayMs(anim("a", { delayMs: 99999 }))).toBe(10000);
  });

  it("normalises the first animation's trigger to click", () => {
    const list = [anim("a", { trigger: "afterPrevious" }), anim("b", { trigger: "withPrevious" })];
    expect(effectiveTrigger(list, 0)).toBe("click");
    expect(effectiveTrigger(list, 1)).toBe("withPrevious");
  });

  it("partitions into build groups at click boundaries", () => {
    const list = [
      anim("a", { trigger: "afterPrevious" }), // first → click
      anim("b", { trigger: "withPrevious" }),
      anim("c", { trigger: "afterPrevious" }),
      anim("d", { trigger: "click" }),
      anim("e", { trigger: "withPrevious" }),
    ];
    const groups = buildGroups(list);
    expect(groups.map((g) => g.map((a) => a.frameId))).toEqual([["a", "b", "c"], ["d", "e"]]);
  });

  it("retargets duplicated animations, dropping unmapped frames", () => {
    const map = new Map([
      ["f1", "n1"],
      ["f2", "n2"],
    ]);
    const list = [anim("f1"), anim("f3"), anim("f2")];
    expect(retargetAnimations(list, map)).toEqual([anim("n1"), anim("n2")]);
    expect(retargetAnimations(undefined, map)).toBeUndefined();
  });

  it("removes animations for a deleted frame, else returns the same reference", () => {
    const list = [anim("f1"), anim("f2"), anim("f1", { effect: "flyIn" })];
    expect(removeAnimationsForFrame(list, "f1")).toEqual([anim("f2")]);
    expect(removeAnimationsForFrame(list, "gone")).toBe(list);
    expect(removeAnimationsForFrame(undefined, "f1")).toBeUndefined();
  });

  it("detects preserved out-of-subset timing", () => {
    const base: Slide = { id: "s", frames: [] };
    expect(hasPreservedTiming(base)).toBe(false);
    expect(hasPreservedTiming({ ...base, preserved: { "p:timing": "<p:timing/>" } })).toBe(true);
    expect(hasPreservedTiming({ ...base, preserved: { anim: "<anim:par/>" } })).toBe(true);
    expect(hasPreservedTiming({ ...base, preserved: { "@root": " x" } })).toBe(false);
  });
});
