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
