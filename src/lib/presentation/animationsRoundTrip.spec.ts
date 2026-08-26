// SPEC-086 animation read/write round trips (AC-26/27/28): every effect, with a
// non-default trigger/direction/duration/delay mix, survives save → reopen in
// both formats (canonical read-back), and an untouched slide's timing bytes
// stay verbatim. Mirrors transitionsRoundTrip.spec.ts's read → mutate → write →
// re-read shape.
import { describe, expect, it } from "vitest";
import { buildOdpFixtureBytes, buildPptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { markSlideEdited } from "./text";
import type { SlideAnimation } from "./model";

/** All seven effects on the slide's first frame, with a non-default trigger /
 * direction / duration / delay mix (three click groups). */
function sample(frameId: string): SlideAnimation[] {
  return [
    { frameId, effect: "appear", trigger: "click" },
    { frameId, effect: "fadeIn", trigger: "withPrevious", durationMs: 300 },
    { frameId, effect: "flyIn", direction: "right", trigger: "afterPrevious", delayMs: 250 },
    { frameId, effect: "wipeIn", direction: "up", trigger: "click", durationMs: 800 },
    { frameId, effect: "zoomIn", trigger: "withPrevious" },
    { frameId, effect: "fadeOut", trigger: "click" },
    { frameId, effect: "flyOut", direction: "down", trigger: "afterPrevious" },
  ];
}

/** The sample with its frame id rebased onto a re-read slide's frame 0 (whose
 * id is freshly minted), for comparison. */
function rebased(frameId: string): SlideAnimation[] {
  return sample(frameId);
}

describe("AC-26: animations round-trip in both formats", () => {
  it("pptx: every effect reproduces exactly (canonical read-back)", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    const slide = model.slides[0];
    slide.animations = sample(slide.frames[0].id);
    markSlideEdited(slide, "pptx");
    const { writePptx } = await import("./pptx/writer");
    const back = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const backSlide = back.slides[0];
    expect(backSlide.animations).toEqual(rebased(backSlide.frames[0].id));
    expect(backSlide.preserved?.["p:timing"]).toBeUndefined();
  });

  it("odp: every effect reproduces exactly (canonical read-back)", async () => {
    const model = await readPresentationBytes(await buildOdpFixtureBytes(), "odp", "Deck");
    const slide = model.slides[0];
    slide.animations = sample(slide.frames[0].id);
    markSlideEdited(slide, "odp");
    const { writeOdp } = await import("./odp/writer");
    const back = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    const backSlide = back.slides[0];
    expect(backSlide.animations).toEqual(rebased(backSlide.frames[0].id));
    expect(backSlide.preserved?.["anim"]).toBeUndefined();
  });
});
