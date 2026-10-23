// SPEC-086 AC-29/27/28: the animation fixtures cover all three paths (plain,
// canonical, exotic-preserved), the canonical forms read back to the model, an
// exotic timing tree is preserved not parsed, and a canonically-written slide
// re-emits byte-stable timing / an untouched exotic slide keeps its bytes.
import { describe, expect, it } from "vitest";
import {
  buildOdpAnimationsFixtureBytes,
  buildOdpDepthFixtureBytes,
  buildPptxAnimationsFixtureBytes,
  buildPptxDepthFixtureBytes,
  EXPECTED_DEPTH_ANIMATIONS,
  EXPECTED_FIXTURE_ANIMATIONS,
  PPTX_CANONICAL_TIMING,
  SPEC105_BULLETS,
  SPEC105_TITLES,
  type DepthFixtureAnimation,
} from "./fixtures/build";
import { slideOutline } from "./slideModel";
import { hasPreservedTiming } from "./animations";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { readPresentationBytes } from "./open";
import { serializeSlideXml } from "./pptx/serialize";
import type { Slide, SlideAnimation } from "./model";

/** Strips the (minted) frame ids for comparison against the expected shape. */
function withoutIds(animations: SlideAnimation[] | undefined): Omit<SlideAnimation, "frameId">[] {
  return (animations ?? []).map(({ frameId: _frameId, ...rest }) => rest);
}

describe("AC-29: pptx animation fixture coverage", () => {
  it("plain / canonical / exotic slides take the three paths", async () => {
    const parsed = await parsePptx(await buildPptxAnimationsFixtureBytes());
    // Slide 1: plain, no animations, no preserved timing.
    expect(parsed.slides[0].animations).toBeUndefined();
    expect(parsed.slides[0].preserved?.["p:timing"]).toBeUndefined();
    // Slide 2: canonical forms model exactly, timing dropped from preserved.
    expect(withoutIds(parsed.slides[1].animations)).toEqual(EXPECTED_FIXTURE_ANIMATIONS);
    expect(parsed.slides[1].preserved?.["p:timing"]).toBeUndefined();
    // The animations target the body (frame 1) and shape (frame 2), by id.
    expect(parsed.slides[1].animations?.[0].frameId).toBe(parsed.slides[1].frames[1].id);
    expect(parsed.slides[1].animations?.[1].frameId).toBe(parsed.slides[1].frames[2].id);
    // Slide 3: exotic timing preserved, not modelled.
    expect(parsed.slides[2].animations).toBeUndefined();
    expect(parsed.slides[2].preserved?.["p:timing"]).toContain('presetClass="emph"');
  });

  it("AC-28: a canonically-read slide re-emits byte-stable timing", async () => {
    const model = await readPresentationBytes(await buildPptxAnimationsFixtureBytes(), "pptx", "Deck");
    const slide = model.slides[1];
    const xml = serializeSlideXml(slide, new Map(), slide.partName ?? "ppt/slides/slide2.xml");
    // The emitted timing is byte-identical to the fixture's canonical timing:
    // the serialiser mints body→3, shape→4, matching the fixture's ids.
    expect(xml).toContain(PPTX_CANONICAL_TIMING);
  });

  it("AC-27: an untouched exotic slide keeps its timing bytes through a rebuild", async () => {
    const model = await readPresentationBytes(await buildPptxAnimationsFixtureBytes(), "pptx", "Deck");
    // Dirty only the plain slide; the exotic slide is untouched.
    model.slides[0].animations = [{ frameId: model.slides[0].frames[1].id, effect: "appear", trigger: "click" }];
    model.slides[0].dirty = true;
    const { writePptx } = await import("./pptx/writer");
    const back = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    expect(back.slides[2].animations).toBeUndefined();
    expect(back.slides[2].preserved?.["p:timing"]).toContain('presetClass="emph"');
  });
});

describe("AC-29: odp animation fixture coverage", () => {
  it("plain / canonical / exotic slides take the three paths", async () => {
    const parsed = await parseOdp(await buildOdpAnimationsFixtureBytes());
    expect(parsed.slides[0].animations).toBeUndefined();
    expect(parsed.slides[0].preserved?.["anim"]).toBeUndefined();
    expect(withoutIds(parsed.slides[1].animations)).toEqual(EXPECTED_FIXTURE_ANIMATIONS);
    expect(parsed.slides[1].preserved?.["anim"]).toBeUndefined();
    expect(parsed.slides[1].animations?.[0].frameId).toBe(parsed.slides[1].frames[1].id);
    expect(parsed.slides[1].animations?.[1].frameId).toBe(parsed.slides[1].frames[2].id);
    expect(parsed.slides[2].animations).toBeUndefined();
    expect(parsed.slides[2].preserved?.["anim"]).toContain("ooo-emphasis-flash-bulb");
  });

  it("round-trips the canonical odp slide", async () => {
    const model = await readPresentationBytes(await buildOdpAnimationsFixtureBytes(), "odp", "Deck");
    const { writeOdp } = await import("./odp/writer");
    const back = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    expect(withoutIds(back.slides[1].animations)).toEqual(EXPECTED_FIXTURE_ANIMATIONS);
  });
});

/** A slide's animations with their frames as z-order indices (SPEC-105). */
function byFrameIndex(slide: Slide): DepthFixtureAnimation[] {
  const index = (id: string): number => slide.frames.findIndex((f) => f.id === id);
  return (slide.animations ?? []).map(({ frameId, triggerFrameId, ...rest }) =>
    triggerFrameId === undefined ? { ...rest, frame: index(frameId) } : { ...rest, frame: index(frameId), triggerFrame: index(triggerFrameId) },
  );
}

describe("SPEC-105 animation-depth fixtures", () => {
  for (const [format, build] of [
    ["pptx", buildPptxDepthFixtureBytes],
    ["odp", buildOdpDepthFixtureBytes],
  ] as const) {
    it(`${format}: a bullet slide, an emphasis + path slide, a trigger slide and an out-of-subset slide`, async () => {
      const model = await readPresentationBytes(await build(), format, "Deck");
      expect(model.slides.map((slide) => slideOutline(slide).title)).toEqual(SPEC105_TITLES);
      // 1: four bullets, no timing (the authoring target).
      expect(slideOutline(model.slides[0]).lines).toEqual([...SPEC105_BULLETS, "Press me"]);
      expect(model.slides[0].animations).toBeUndefined();
      expect(hasPreservedTiming(model.slides[0])).toBe(false);
      // 2 and 3: every SPEC-105 form models exactly.
      for (const index of [1, 2]) {
        expect(byFrameIndex(model.slides[index])).toEqual(EXPECTED_DEPTH_ANIMATIONS[index]);
        expect(hasPreservedTiming(model.slides[index])).toBe(false);
      }
      // 4: out of subset — preserved, never modelled.
      expect(model.slides[3].animations).toBeUndefined();
      expect(hasPreservedTiming(model.slides[3])).toBe(true);
    });
  }

  it("pptx: the out-of-subset slide is a partial p:pRg cover; odp: a flash bulb and a quadratic path", async () => {
    const pptx = await readPresentationBytes(await buildPptxDepthFixtureBytes(), "pptx", "Deck");
    expect(pptx.slides[3].preserved?.["p:timing"]).toMatch(/<p:pRg st="0" end="0"\/>[\s\S]*<p:pRg st="2" end="2"\/>/);
    const odp = await readPresentationBytes(await buildOdpDepthFixtureBytes(), "odp", "Deck");
    expect(odp.slides[3].preserved?.["anim"]).toContain("ooo-emphasis-flash-bulb");
    expect(odp.slides[3].preserved?.["anim"]).toContain("q 0.1 -0.2 0.2 0");
  });
});
