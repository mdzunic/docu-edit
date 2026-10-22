import { describe, expect, it } from "vitest";
import type { Slide, SlideAnimation, SlideFrame, TextFrame } from "./model";
import {
  buildGroups,
  carriesSourceWireId,
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
