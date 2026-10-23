import { describe, expect, it } from "vitest";
import type { Slide, SlideAnimation, SlideFrame, TextFrame } from "./model";
import {
  animationConstraintsHold,
  ANIMATION_EFFECTS,
  buildGroups,
  carriesSourceWireId,
  DEFAULT_EMPHASIS_COLOR,
  DEFAULT_PATH_DISTANCE,
  DEFAULT_SCALE_PERCENT,
  DEFAULT_SPIN_DEGREES,
  effectClass,
  effectiveEmphasisColor,
  effectiveRepeatCount,
  effectiveScalePercent,
  effectiveSpinDegrees,
  EMPHASIS_EFFECTS,
  interactiveSequences,
  mainSequence,
  MAX_REPEAT_COUNT,
  MIN_REPEAT_COUNT,
  PATH_PRESETS,
  clearSourceWireIds,
  defaultAnimationDirection,
  dropPreservedTimingTargeting,
  effectiveAnimationDelayMs,
  effectiveAnimationDurationMs,
  effectiveTrigger,
  hasPreservedTiming,
  isEntranceEffect,
  removeAnimationsForFrame,
  retainSourceWireIds,
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

/** A text frame with the given source id and one paragraph per `paraIds`
 * entry (each with that source id, or none). */
function text(id: string, sourceWireId: string | undefined, paraIds: (string | undefined)[]): TextFrame {
  const frame: TextFrame = {
    id,
    kind: "text",
    xCm: 0,
    yCm: 0,
    widthCm: 1,
    heightCm: 1,
    paragraphs: paraIds.map((p) => (p === undefined ? { runs: [] } : { runs: [], sourceWireId: p })),
  };
  if (sourceWireId !== undefined) frame.sourceWireId = sourceWireId;
  return frame;
}

/** `frame` with each line recording its own index as its source `a:p`. */
function indexed(frame: TextFrame): TextFrame {
  frame.paragraphs.forEach((p, i) => (p.pptxSourceParagraph = i));
  return frame;
}

// SPEC-086 AC-4: the source ids that tie preserved timing to its targets.
describe("preserved-timing targets", () => {
  it("keeps only the targeted ids, first claimant first, never inside a group", () => {
    const child = text("c", "8", ["p8"]);
    const frames: SlideFrame[] = [
      text("a", "3", ["p1", "p2"]),
      text("b", "3", ["p1"]), // a second claimant of 3 / p1
      { id: "g", kind: "group", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, children: [child], sourceWireId: "9" },
      text("d", "4", [undefined]),
    ];
    retainSourceWireIds(frames, new Set(["3", "p1", "8", "p8", "9"]));
    expect(frames.map((f) => f.sourceWireId)).toEqual(["3", undefined, "9", undefined]);
    expect((frames[0] as TextFrame).paragraphs.map((p) => p.sourceWireId)).toEqual(["p1", undefined]);
    expect((frames[1] as TextFrame).paragraphs[0].sourceWireId).toBeUndefined();
    expect(child.sourceWireId).toBeUndefined();
    expect(child.paragraphs[0].sourceWireId).toBeUndefined();
  });

  it("keeps lines' source paragraph indices only in a text/shape frame the timing addresses by range", () => {
    const byRange = indexed(text("a", "12", [undefined, undefined]));
    const wholeShape = indexed(text("b", "13", [undefined]));
    const untargeted = indexed(text("c", "14", [undefined]));
    const child = indexed(text("k", "16", [undefined]));
    const group: SlideFrame = { id: "g", kind: "group", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, children: [child], sourceWireId: "15" };
    retainSourceWireIds([byRange, wholeShape, untargeted, group], new Set(["12", "13", "15"]), new Set(["12", "14", "15", "16"]));
    expect(byRange.paragraphs.map((p) => p.pptxSourceParagraph)).toEqual([0, 1]);
    expect(wholeShape.paragraphs[0].pptxSourceParagraph).toBeUndefined();
    // Range-addressed but not a kept target (no record of its shape id).
    expect(untargeted.paragraphs[0].pptxSourceParagraph).toBeUndefined();
    expect(child.paragraphs[0].pptxSourceParagraph).toBeUndefined();
    // Read with no range-addressed timing (the default), nothing keeps one.
    const plain = indexed(text("p", "12", [undefined]));
    retainSourceWireIds([plain], new Set(["12"]));
    expect(plain.paragraphs[0].pptxSourceParagraph).toBeUndefined();
  });

  it("clears source paragraph indices with the source ids", () => {
    const frame = text("a", "12", [undefined]);
    frame.paragraphs[0].pptxSourceParagraph = 0;
    clearSourceWireIds([frame]);
    expect(frame.paragraphs[0]).toEqual({ runs: [] });
  });

  it("a frame carries a target through its own id or any of its paragraphs'", () => {
    expect(carriesSourceWireId(text("a", undefined, [undefined]))).toBe(false);
    expect(carriesSourceWireId(text("a", "3", []))).toBe(true);
    expect(carriesSourceWireId(text("a", undefined, [undefined, "p2"]))).toBe(true);
  });

  it("drops the preserved timing (and every source id) only when a removed frame was a target", () => {
    const targeted = text("t", "5", []);
    const slide = (): Slide => ({
      id: "s",
      frames: [text("k", "6", ["p6"])],
      preserved: { "p:timing": "<p:timing/>", "@root": " x" },
    });
    const untouched = slide();
    dropPreservedTimingTargeting(untouched, [text("u", undefined, [undefined])]);
    expect(untouched.preserved?.["p:timing"]).toBe("<p:timing/>");
    expect(untouched.frames[0].sourceWireId).toBe("6");

    const dropped = slide();
    dropPreservedTimingTargeting(dropped, [targeted]);
    expect(dropped.preserved).toEqual({ "@root": " x" });
    expect(dropped.frames[0].sourceWireId).toBeUndefined();
    expect((dropped.frames[0] as TextFrame).paragraphs[0].sourceWireId).toBeUndefined();

    // Modelled animations are their own rule (removeAnimationsForFrame).
    const modelled: Slide = { ...slide(), animations: [] };
    dropPreservedTimingTargeting(modelled, [targeted]);
    expect(modelled.preserved?.["p:timing"]).toBe("<p:timing/>");
  });
});

describe("SPEC-105 animation helpers", () => {
  it("classifies every effect: entrance, exit, emphasis, path", () => {
    expect(["appear", "fadeIn", "flyIn", "wipeIn", "zoomIn"].map(effectClass)).toEqual(Array(5).fill("entrance"));
    expect(["fadeOut", "flyOut"].map(effectClass)).toEqual(["exit", "exit"]);
    expect(EMPHASIS_EFFECTS.map(effectClass)).toEqual(Array(4).fill("emphasis"));
    expect(effectClass("motionPath")).toBe("path");
    expect(ANIMATION_EFFECTS).toHaveLength(12);
    expect(PATH_PRESETS).toEqual(["line", "arc", "turn", "loop", "custom"]);
  });

  it("isEntranceEffect is false for every emphasis and path effect", () => {
    for (const effect of [...EMPHASIS_EFFECTS, "motionPath"]) {
      expect(isEntranceEffect(effect), effect).toBe(false);
    }
  });

  it("exposes the new bounds and defaults", () => {
    expect([MIN_REPEAT_COUNT, MAX_REPEAT_COUNT]).toEqual([2, 10]);
    expect(DEFAULT_SPIN_DEGREES).toBe(360);
    expect(DEFAULT_SCALE_PERCENT).toBe(150);
    expect(DEFAULT_EMPHASIS_COLOR).toBe("#ff0000");
    expect(DEFAULT_PATH_DISTANCE).toBe(0.25);
    expect(effectiveSpinDegrees(anim("a", { effect: "spin" }))).toBe(360);
    expect(effectiveSpinDegrees(anim("a", { effect: "spin", spinDegrees: -9000 }))).toBe(-3600);
    expect(effectiveScalePercent(anim("a", { effect: "growShrink", scalePercent: 900 }))).toBe(400);
    expect(effectiveEmphasisColor(anim("a", { effect: "colorChange" }))).toBe("#ff0000");
  });

  it("effectiveRepeatCount: repeatCount ?? 1, an until-click repeat counting one pass", () => {
    expect(effectiveRepeatCount(anim("a"))).toBe(1);
    expect(effectiveRepeatCount(anim("a", { repeatCount: 4 }))).toBe(4);
    expect(effectiveRepeatCount(anim("a", { repeatUntilClick: true }))).toBe(1);
    expect(effectiveRepeatCount(anim("a", { repeatCount: 99 }))).toBe(10);
  });

  it("partitions rows into the main sequence and interactive sequences by trigger frame, in array order", () => {
    const rows = [
      anim("a"),
      anim("b", { trigger: "onClickOf", triggerFrameId: "t1" }),
      anim("c", { trigger: "afterPrevious" }),
      anim("d", { trigger: "onClickOf", triggerFrameId: "t2" }),
      anim("e", { trigger: "onClickOf", triggerFrameId: "t1" }),
    ];
    expect(mainSequence(rows).map((a) => a.frameId)).toEqual(["a", "c"]);
    expect(interactiveSequences(rows).map((s) => [s.triggerFrameId, s.animations.map((a) => a.frameId)])).toEqual([
      ["t1", ["b", "e"]],
      ["t2", ["d"]],
    ]);
  });

  it("effectiveTrigger forces click only on the first main-sequence row; an onClickOf row keeps its trigger", () => {
    const rows = [
      anim("a", { trigger: "onClickOf", triggerFrameId: "t" }),
      anim("b", { trigger: "afterPrevious" }),
      anim("c", { trigger: "withPrevious" }),
    ];
    expect(rows.map((_, i) => effectiveTrigger(rows, i))).toEqual(["onClickOf", "click", "withPrevious"]);
    // An onClickOf row never opens a main-sequence group.
    expect(buildGroups(rows).map((g) => g.map((a) => a.frameId))).toEqual([["b", "c"]]);
  });

  it("checks the model constraints", () => {
    const frames = new Set(["a", "b"]);
    expect(animationConstraintsHold(anim("a"), frames)).toBe(true);
    expect(animationConstraintsHold(anim("a", { trigger: "onClickOf" }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { trigger: "onClickOf", triggerFrameId: "b" }), frames)).toBe(true);
    expect(animationConstraintsHold(anim("a", { trigger: "onClickOf", triggerFrameId: "a" }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { trigger: "onClickOf", triggerFrameId: "zz" }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { triggerFrameId: "b" }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { repeatCount: 3, repeatUntilClick: true }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { repeatCount: 1 }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { effect: "motionPath" }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { effect: "motionPath", path: { preset: "line" } }), frames)).toBe(true);
    expect(animationConstraintsHold(anim("a", { path: { preset: "line" } }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { effect: "motionPath", path: { preset: "custom" } }), frames)).toBe(false);
    expect(animationConstraintsHold(anim("a", { emphasisColor: "red" }), frames)).toBe(false);
  });

  it("retargets and removes trigger frames with their rows", () => {
    const rows = [anim("a"), anim("b", { trigger: "onClickOf", triggerFrameId: "a" })];
    expect(retargetAnimations(rows, new Map([["a", "A"], ["b", "B"]]))).toEqual([
      anim("A"),
      anim("B", { trigger: "onClickOf", triggerFrameId: "A" }),
    ]);
    // The trigger frame did not survive the copy → the row goes with it.
    expect(retargetAnimations(rows, new Map([["b", "B"]]))).toEqual([]);
    // Deleting the trigger frame drops the rows it triggers.
    expect(removeAnimationsForFrame(rows, "a")).toEqual([]);
    expect(removeAnimationsForFrame([anim("b", { trigger: "onClickOf", triggerFrameId: "c" })], "c")).toEqual([]);
  });
});
