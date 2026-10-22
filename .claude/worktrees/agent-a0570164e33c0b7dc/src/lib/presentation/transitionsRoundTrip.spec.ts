// SPEC-085 transition read/write round trips (AC-11/12/20/21), driven over the
// per-effect transition fixtures and models mutated through the editing
// commands. Mirrors roundTrip.spec.ts's read → mutate → write → re-read shape.
import { describe, expect, it } from "vitest";
import {
  buildOdpTransitionsFixtureBytes,
  buildPptxTransitionsFixtureBytes,
  EXPECTED_ODP_TRANSITIONS,
  EXPECTED_TRANSITIONS,
} from "./fixtures/build";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { readPresentationBytes } from "./open";
import { applyTransitionToSlide } from "./slideTransition";
import { markSlideEdited } from "./text";
import type { SlideTransition } from "./model";

const SIX: SlideTransition[] = [
  { type: "fade", durationMs: 2000 },
  { type: "push", direction: "right", durationMs: 2000, advanceAfterMs: 5000 },
  { type: "wipe", direction: "up", durationMs: 2000 },
  { type: "cover", direction: "down", durationMs: 2000 },
  { type: "uncover", direction: "left", durationMs: 2000 },
  { type: "split", durationMs: 2000 },
];

describe("AC-21: authored fixture markup parses to the expected model", () => {
  it("pptx: each mapped effect + the out-of-subset one", async () => {
    const parsed = await parsePptx(await buildPptxTransitionsFixtureBytes());
    parsed.slides.forEach((slide, i) => {
      const expected = EXPECTED_TRANSITIONS[i];
      if (expected === null) {
        expect(slide.transition).toBeUndefined();
        expect(slide.preserved?.["p:transition"]).toBeDefined(); // AC-3 preserved
      } else {
        expect(slide.transition).toEqual(expected);
        expect(slide.preserved?.["p:transition"]).toBeUndefined();
      }
    });
  });

  it("odp: each mapped effect + the out-of-subset one", async () => {
    const parsed = await parseOdp(await buildOdpTransitionsFixtureBytes());
    parsed.slides.forEach((slide, i) => {
      const expected = EXPECTED_ODP_TRANSITIONS[i];
      if (expected === null) {
        expect(slide.transition).toBeUndefined();
        expect(slide.preserved?.["draw:transition"]).toBe("1"); // AC-3 preserved marker
      } else {
        expect(slide.transition).toEqual(expected);
      }
    });
  });
});

describe("AC-20: round trip preserves the model in both formats", () => {
  it("pptx: quantises duration on reopen (the pinned delta), keeps type/direction/advance", async () => {
    const model = await readPresentationBytes(await buildPptxTransitionsFixtureBytes(), "pptx", "Deck");
    SIX.forEach((t, i) => applyTransitionToSlide(model.slides[i], t, "pptx"));
    const { writePptx } = await import("./pptx/writer");
    const bytes = (await writePptx(model)).bytes;
    const back = await readPresentationBytes(bytes, "pptx", "Deck");
    SIX.forEach((t, i) => {
      expect(back.slides[i].transition?.type).toBe(t.type);
      expect(back.slides[i].transition?.direction).toBe(t.direction);
      // 2000ms -> spd="slow" -> 1000ms on reopen (the single sanctioned delta).
      expect(back.slides[i].transition?.durationMs).toBe(1000);
    });
    // The auto-advance survives (push slide).
    expect(back.slides[1].transition?.advanceAfterMs).toBe(5000);
  });

  it("odp: preserves duration exactly (no quantisation)", async () => {
    const model = await readPresentationBytes(await buildOdpTransitionsFixtureBytes(), "odp", "Deck");
    SIX.forEach((t, i) => applyTransitionToSlide(model.slides[i], t, "odp"));
    const { writeOdp } = await import("./odp/writer");
    const bytes = (await writeOdp(model)).bytes;
    const back = await readPresentationBytes(bytes, "odp", "Deck");
    SIX.forEach((t, i) => {
      expect(back.slides[i].transition?.type).toBe(t.type);
      expect(back.slides[i].transition?.direction).toBe(t.direction);
      expect(back.slides[i].transition?.durationMs).toBe(2000); // exact
    });
    expect(back.slides[1].transition?.advanceAfterMs).toBe(5000);
  });
});

describe("AC-12: an untouched slide keeps its original transition bytes", () => {
  it("pptx: a text-dirtied slide whose transition was never touched re-emits it verbatim", async () => {
    const model = await readPresentationBytes(await buildPptxTransitionsFixtureBytes(), "pptx", "Deck");
    // Dirty the out-of-subset (Morph) slide for a NON-transition reason.
    const morph = model.slides[model.slides.length - 1];
    expect(morph.preserved?.["p:transition"]).toContain("p:morph");
    markSlideEdited(morph, "pptx");
    const { writePptx } = await import("./pptx/writer");
    const back = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const backMorph = back.slides[back.slides.length - 1];
    expect(backMorph.transition).toBeUndefined();
    expect(backMorph.preserved?.["p:transition"]).toContain("p:morph"); // survived
  });

  it("odp: a fade slide untouched by this spec keeps its transition through a rebuild", async () => {
    const model = await readPresentationBytes(await buildOdpTransitionsFixtureBytes(), "odp", "Deck");
    // Set a transition on slide 2 only; slide 1 (fade) is never touched but the
    // whole content.xml rebuilds — its original drawing-page style survives.
    applyTransitionToSlide(model.slides[1], { type: "cover", direction: "left", durationMs: 600 }, "odp");
    const { writeOdp } = await import("./odp/writer");
    const back = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    expect(back.slides[0].transition).toEqual({ type: "fade", durationMs: 800 });
    expect(back.slides[1].transition).toEqual({ type: "cover", direction: "left", durationMs: 600 });
  });
});
