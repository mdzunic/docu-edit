import { describe, expect, it } from "vitest";
import type { PresentationModel, Slide } from "./model";
import {
  applyTransitionToAll,
  applyTransitionToSlide,
  hasPreservedTransition,
  slidesWithPreservedTransition,
} from "./slideTransition";

function slide(overrides: Partial<Slide> = {}): Slide {
  return { id: crypto.randomUUID(), frames: [], ...overrides };
}

function deck(format: "pptx" | "odp", slides: Slide[]): PresentationModel {
  return {
    kind: "presentation",
    id: "d",
    title: "Deck",
    format,
    slides,
    slideWidthCm: 25,
    slideHeightCm: 14,
  };
}

describe("hasPreservedTransition", () => {
  it("detects a pptx preserved out-of-subset transition", () => {
    expect(hasPreservedTransition(slide({ preserved: { "p:transition": "<p:transition/>" } }))).toBe(true);
  });
  it("detects an odp preserved marker", () => {
    expect(hasPreservedTransition(slide({ preserved: { "draw:transition": "1" } }))).toBe(true);
  });
  it("is false for a plain slide or a modelled transition", () => {
    expect(hasPreservedTransition(slide())).toBe(false);
    expect(hasPreservedTransition(slide({ transition: { type: "fade" }, preserved: { "@root": " x" } }))).toBe(false);
  });
});

describe("applyTransitionToSlide", () => {
  it("sets a copy, marks pptx dirty, keeps sourceXml", () => {
    const s = slide({ sourceXml: "<p/>" });
    const t = { type: "push", direction: "left", durationMs: 1000 } as const;
    applyTransitionToSlide(s, t, "pptx");
    expect(s.transition).toEqual(t);
    expect(s.transition).not.toBe(t); // copied
    expect(s.dirty).toBe(true);
    expect(s.sourceXml).toBe("<p/>"); // pptx keeps it
  });

  it("drops the preserved out-of-subset transition it replaces", () => {
    const s = slide({ preserved: { "p:transition": "<p:transition><p:morph/></p:transition>", "@root": " x" } });
    applyTransitionToSlide(s, { type: "fade" }, "pptx");
    expect(s.preserved?.["p:transition"]).toBeUndefined();
    expect(s.preserved?.["@root"]).toBe(" x"); // other preserved kept
  });

  it("clears an odp preserved marker and drops sourceXml so the page rebuilds", () => {
    const s = slide({ preserved: { "draw:transition": "1" }, sourceXml: "<draw:page/>" });
    applyTransitionToSlide(s, { type: "split" }, "odp");
    expect(s.preserved).toBeUndefined(); // only key removed → whole map dropped
    expect(s.dirty).toBe(true);
    expect(s.sourceXml).toBeUndefined();
  });
});

describe("applyTransitionToAll", () => {
  it("copies the transition onto every slide, replacing preserved ones", () => {
    const model = deck("pptx", [
      slide({ transition: { type: "fade" } }),
      slide(),
      slide({ preserved: { "p:transition": "<p:transition><p:zoom/></p:transition>" } }),
    ]);
    applyTransitionToAll(model, { type: "push", direction: "up", durationMs: 800 });
    for (const s of model.slides) {
      expect(s.transition).toEqual({ type: "push", direction: "up", durationMs: 800 });
      expect(s.dirty).toBe(true);
    }
    expect(model.slides[2].preserved).toBeUndefined();
  });
});

describe("slidesWithPreservedTransition", () => {
  it("counts only slides holding a preserved out-of-subset transition", () => {
    const model = deck("pptx", [
      slide({ preserved: { "p:transition": "<p:transition/>" } }),
      slide(),
      slide({ preserved: { "draw:transition": "1" } }),
      slide({ transition: { type: "fade" } }),
    ]);
    expect(slidesWithPreservedTransition(model)).toBe(2);
  });
});
