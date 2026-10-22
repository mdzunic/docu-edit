// SPEC-086 AC-29/27/28: the animation fixtures cover all three paths (plain,
// canonical, exotic-preserved), the canonical forms read back to the model, an
// exotic timing tree is preserved not parsed, and a canonically-written slide
// re-emits byte-stable timing / an untouched exotic slide keeps its bytes.
import { describe, expect, it } from "vitest";
import {
  buildOdpAnimationsFixtureBytes,
  buildPptxAnimationsFixtureBytes,
  EXPECTED_FIXTURE_ANIMATIONS,
  PPTX_CANONICAL_TIMING,
} from "./fixtures/build";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { readPresentationBytes } from "./open";
import { serializeSlideXml } from "./pptx/serialize";
import type { SlideAnimation } from "./model";

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
    expect(parsed.slides[2].preserved?.["anim"]).toContain("ooo-emphasis-spin");
  });

  it("round-trips the canonical odp slide", async () => {
    const model = await readPresentationBytes(await buildOdpAnimationsFixtureBytes(), "odp", "Deck");
    const { writeOdp } = await import("./odp/writer");
    const back = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    expect(withoutIds(back.slides[1].animations)).toEqual(EXPECTED_FIXTURE_ANIMATIONS);
  });
});
