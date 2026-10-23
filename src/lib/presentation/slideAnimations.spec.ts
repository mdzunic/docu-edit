import { describe, expect, it } from "vitest";
import type { Slide } from "./model";
import { addAnimation, moveAnimation, removeAnimation, updateAnimation } from "./slideAnimations";
import { deleteFrame } from "./frames";
import { duplicateSlide } from "./slides";
import type { PresentationModel } from "./model";

function slide(): Slide {
  return {
    id: "s1",
    frames: [
      { id: "f1", kind: "text", xCm: 0, yCm: 0, widthCm: 5, heightCm: 2, paragraphs: [] },
      { id: "f2", kind: "shape", xCm: 6, yCm: 0, widthCm: 5, heightCm: 2, geometry: "rect" },
    ],
  };
}

describe("slide animation commands", () => {
  it("adds a fade-in/click animation, dirties, and drops preserved timing", () => {
    const s = slide();
    s.preserved = { "p:timing": "<p:timing/>", "@root": " x" };
    const index = addAnimation(s, "f1", "pptx");
    expect(index).toBe(0);
    expect(s.animations).toEqual([{ frameId: "f1", effect: "fadeIn", trigger: "click" }]);
    expect(s.dirty).toBe(true);
    expect(s.preserved?.["p:timing"]).toBeUndefined();
    expect(s.preserved?.["@root"]).toBe(" x"); // unrelated preserved kept
  });

  it("updates fields and clears direction when effect goes non-directional", () => {
    const s = slide();
    addAnimation(s, "f1", "pptx");
    updateAnimation(s, 0, { effect: "flyIn", direction: "up", durationMs: 800, delayMs: 250 }, "pptx");
    expect(s.animations?.[0]).toEqual({
      frameId: "f1",
      effect: "flyIn",
      direction: "up",
      trigger: "click",
      durationMs: 800,
      delayMs: 250,
    });
    updateAnimation(s, 0, { effect: "appear" }, "pptx");
    expect(s.animations?.[0].direction).toBeUndefined();
  });

  it("removes and reorders animations", () => {
    const s = slide();
    addAnimation(s, "f1", "pptx");
    addAnimation(s, "f2", "pptx");
    expect(moveAnimation(s, 1, -1, "pptx")).toBe(0);
    expect(s.animations?.map((a) => a.frameId)).toEqual(["f2", "f1"]);
    expect(removeAnimation(s, 0, "pptx")).toBe(0);
    expect(s.animations?.map((a) => a.frameId)).toEqual(["f1"]);
    expect(removeAnimation(s, 0, "pptx")).toBe(-1);
    expect(s.animations).toEqual([]);
  });
});

describe("frame delete + slide duplicate retargeting", () => {
  function model(): PresentationModel {
    const s = slide();
    s.animations = [
      { frameId: "f1", effect: "fadeIn", trigger: "click" },
      { frameId: "f2", effect: "flyIn", trigger: "afterPrevious" },
    ];
    return {
      kind: "presentation",
      id: "m",
      title: "t",
      format: "pptx",
      slides: [s],
      slideWidthCm: 33.867,
      slideHeightCm: 19.05,
    };
  }

  it("deleting a frame drops its animations (AC-18)", () => {
    const m = model();
    deleteFrame(m, 0, "f1");
    expect(m.slides[0].animations).toEqual([{ frameId: "f2", effect: "flyIn", trigger: "afterPrevious" }]);
  });

  it("duplicating a slide retargets animations to the copied frames (AC-19)", () => {
    const m = model();
    const copyIndex = duplicateSlide(m, 0);
    const copy = m.slides[copyIndex];
    // Copy's frames have fresh ids; its animations reference those, not the source's.
    const copyIds = copy.frames.map((f) => f.id);
    expect(copy.animations?.map((a) => a.frameId)).toEqual(copyIds);
    expect(copy.animations?.map((a) => a.effect)).toEqual(["fadeIn", "flyIn"]);
    // The source is untouched.
    expect(m.slides[0].animations?.map((a) => a.frameId)).toEqual(["f1", "f2"]);
  });
});

describe("SPEC-105 widened AnimationPatch", () => {
  function withRow(over: Partial<import("./model").SlideAnimation> = {}): Slide {
    const s = slide();
    s.animations = [{ frameId: "f1", effect: "fadeIn", trigger: "click", ...over }];
    return s;
  }

  it("builds by paragraph, a build trigger only beside a build, click stored as absent", () => {
    const s = withRow();
    updateAnimation(s, 0, { buildTrigger: "afterPrevious" }, "pptx");
    expect(s.animations![0].buildTrigger).toBeUndefined(); // no build to trigger
    updateAnimation(s, 0, { buildBy: "byFirstLevelParagraph", buildTrigger: "withPrevious" }, "pptx");
    expect(s.animations![0]).toMatchObject({ buildBy: "byFirstLevelParagraph", buildTrigger: "withPrevious" });
    updateAnimation(s, 0, { buildTrigger: "click" }, "pptx");
    expect(s.animations![0].buildTrigger).toBeUndefined();
    updateAnimation(s, 0, { buildTrigger: "afterPrevious" }, "pptx");
    updateAnimation(s, 0, { clearBuildBy: true }, "pptx");
    expect(s.animations![0]).toEqual({ frameId: "f1", effect: "fadeIn", trigger: "click" });
    expect(s.dirty).toBe(true);
  });

  it("repeatCount and repeatUntilClick exclude each other; a single pass is stored as absent", () => {
    const s = withRow({ effect: "spin" });
    updateAnimation(s, 0, { repeatCount: 3 }, "pptx");
    expect(s.animations![0].repeatCount).toBe(3);
    updateAnimation(s, 0, { repeatUntilClick: true }, "pptx");
    expect(s.animations![0]).toMatchObject({ repeatUntilClick: true });
    expect(s.animations![0].repeatCount).toBeUndefined();
    updateAnimation(s, 0, { repeatCount: 25 }, "pptx");
    expect(s.animations![0].repeatCount).toBe(10);
    expect(s.animations![0].repeatUntilClick).toBeUndefined();
    updateAnimation(s, 0, { repeatCount: 1 }, "pptx");
    expect(s.animations![0].repeatCount).toBeUndefined();
  });

  it("stores only true for the three flags", () => {
    const s = withRow();
    updateAnimation(s, 0, { autoReverse: true, rewindOnComplete: true }, "pptx");
    expect(s.animations![0]).toMatchObject({ autoReverse: true, rewindOnComplete: true });
    updateAnimation(s, 0, { autoReverse: false, rewindOnComplete: false }, "pptx");
    expect(s.animations![0]).toEqual({ frameId: "f1", effect: "fadeIn", trigger: "click" });
  });

  it("takes onClickOf only with another frame on the slide; any other trigger drops the trigger frame", () => {
    const s = withRow();
    updateAnimation(s, 0, { trigger: "onClickOf" }, "pptx");
    expect(s.animations![0].trigger).toBe("click"); // no frame named
    updateAnimation(s, 0, { trigger: "onClickOf", triggerFrameId: "f1" }, "pptx");
    expect(s.animations![0].trigger).toBe("click"); // never its own target
    updateAnimation(s, 0, { trigger: "onClickOf", triggerFrameId: "nope" }, "pptx");
    expect(s.animations![0].trigger).toBe("click");
    updateAnimation(s, 0, { trigger: "onClickOf", triggerFrameId: "f2" }, "pptx");
    expect(s.animations![0]).toMatchObject({ trigger: "onClickOf", triggerFrameId: "f2" });
    updateAnimation(s, 0, { triggerFrameId: "f1" }, "pptx");
    expect(s.animations![0].triggerFrameId).toBe("f2");
    updateAnimation(s, 0, { trigger: "withPrevious" }, "pptx");
    expect(s.animations![0].triggerFrameId).toBeUndefined();
  });

  it("an effect change drops options it no longer uses and seeds a motion path's line", () => {
    const s = withRow({ effect: "spin", spinDegrees: 90 });
    updateAnimation(s, 0, { effect: "motionPath" }, "pptx");
    expect(s.animations![0]).toEqual({ frameId: "f1", effect: "motionPath", trigger: "click", path: { preset: "line" } });
    updateAnimation(s, 0, { effect: "growShrink", scalePercent: 200 }, "pptx");
    expect(s.animations![0]).toEqual({ frameId: "f1", effect: "growShrink", trigger: "click", scalePercent: 200 });
  });

  it("clamps the effect options and stores a default value as absent", () => {
    const s = withRow({ effect: "spin" });
    updateAnimation(s, 0, { spinDegrees: 99999 }, "pptx");
    expect(s.animations![0].spinDegrees).toBe(3600);
    updateAnimation(s, 0, { spinDegrees: 360 }, "pptx");
    expect(s.animations![0].spinDegrees).toBeUndefined();
    updateAnimation(s, 0, { scalePercent: 200 }, "pptx"); // not a grow/shrink: ignored
    expect(s.animations![0].scalePercent).toBeUndefined();
    updateAnimation(s, 0, { effect: "colorChange", emphasisColor: "#00AA00" }, "pptx");
    expect(s.animations![0].emphasisColor).toBe("#00aa00");
    updateAnimation(s, 0, { emphasisColor: "green" }, "pptx");
    expect(s.animations![0].emphasisColor).toBe("#00aa00");
    updateAnimation(s, 0, { emphasisColor: "#FF0000" }, "pptx");
    expect(s.animations![0].emphasisColor).toBeUndefined();
  });

  it("normalises a path: defaults omitted, distance clamped, an unparseable custom refused", () => {
    const s = withRow({ effect: "motionPath", path: { preset: "line" } });
    updateAnimation(s, 0, { path: { preset: "arc", direction: "right", distanceFraction: 3 } }, "pptx");
    expect(s.animations![0].path).toEqual({ preset: "arc", distanceFraction: 1 });
    updateAnimation(s, 0, { path: { preset: "loop", direction: "up", distanceFraction: 0.25 } }, "pptx");
    expect(s.animations![0].path).toEqual({ preset: "loop" });
    updateAnimation(s, 0, { path: { preset: "custom", d: "M 0 0 Q 1 1 2 2" } }, "pptx");
    expect(s.animations![0].path).toEqual({ preset: "loop" });
  });

  it("keeps a colour change's duration in 2ms steps (each half is written in whole ms)", () => {
    const s = withRow({ effect: "colorChange" });
    updateAnimation(s, 0, { durationMs: 555 }, "pptx");
    expect(s.animations![0].durationMs).toBe(556);
    const odd = withRow({ durationMs: 701 });
    updateAnimation(odd, 0, { effect: "colorChange" }, "pptx");
    expect(odd.animations![0].durationMs).toBe(702);
    const other = withRow({ effect: "spin" });
    updateAnimation(other, 0, { durationMs: 555 }, "pptx");
    expect(other.animations![0].durationMs).toBe(555);
  });

  it("dirties an odp slide by dropping its verbatim page", () => {
    const s = withRow();
    s.sourceXml = "<draw:page/>";
    updateAnimation(s, 0, { rewindOnComplete: true }, "odp");
    expect(s.sourceXml).toBeUndefined();
  });
});
