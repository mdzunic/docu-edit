// SPEC-086 animation read/write round trips (AC-26/27/28): every effect, with a
// non-default trigger/direction/duration/delay mix, survives save → reopen in
// both formats (canonical read-back), and an untouched slide's timing bytes
// stay verbatim. Mirrors transitionsRoundTrip.spec.ts's read → mutate → write →
// re-read shape.
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import {
  buildOdpAnimationsFixtureBytes,
  buildOdpFixtureBytes,
  buildPptxAnimationsFixtureBytes,
  buildPptxFixtureBytes,
  PPTX_CANONICAL_TIMING,
} from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { markSlideEdited } from "./text";
import { hasPreservedTiming } from "./animations";
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

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** PowerPoint's "Fade, by paragraph" build on the body (spid 3): one click per
 * paragraph, each effect's behaviours narrowed to that paragraph by `p:txEl`. */
function paragraphFade(paragraph: number): string {
  const tgt = `<p:tgtEl><p:spTgt spid="3"><p:txEl><p:pRg st="${paragraph}" end="${paragraph}"/></p:txEl></p:spTgt></p:tgtEl>`;
  return (
    `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:set><p:cBhvr><p:cTn dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}` +
    `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
    `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn dur="500"/>${tgt}</p:cBhvr></p:animEffect>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
  );
}
const PARAGRAPH_BUILD_TIMING =
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
  paragraphFade(0) +
  paragraphFade(1) +
  `</p:childTnLst></p:cTn>` +
  `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
  `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
  `</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>` +
  `<p:bldLst><p:bldP spid="3" grpId="0" build="p"/></p:bldLst></p:timing>`;

/** The pptx animation fixture with slide 2's canonical timing swapped for the
 * paragraph build, and its body given the second paragraph the build targets. */
async function paragraphBuildDeck(): Promise<{ bytes: Uint8Array; slide2: string }> {
  const parts = await unzip(await buildPptxAnimationsFixtureBytes());
  const source = dec(parts.get("ppt/slides/slide2.xml")!);
  const bodyLine = "<a:t>Body line</a:t></a:r></a:p>";
  expect(source).toContain(PPTX_CANONICAL_TIMING);
  expect(source).toContain(bodyLine);
  const slide2 = source
    .replace(bodyLine, `${bodyLine}<a:p><a:r><a:t>Second line</a:t></a:r></a:p>`)
    .replace(PPTX_CANONICAL_TIMING, PARAGRAPH_BUILD_TIMING);
  parts.set("ppt/slides/slide2.xml", enc(slide2));
  return { bytes: await zip([...parts].map(([name, data]) => ({ name, data }))), slide2 };
}

/** The odp animation fixture's page 2, as written in `content.xml`. */
function page2Source(content: string): string {
  return content.slice(content.indexOf(`draw:name="page2"`), content.indexOf(`draw:name="page3"`));
}

describe("AC-4: a by-paragraph text build is preserved, never flattened to whole-frame effects", () => {
  it("pptx: reads as preserved timing, not N whole-frame animations", async () => {
    const { bytes } = await paragraphBuildDeck();
    const slide = (await readPresentationBytes(bytes, "pptx", "Deck")).slides[1];
    expect(slide.animations).toBeUndefined();
    expect(hasPreservedTiming(slide)).toBe(true);
  });

  it("pptx: an untouched slide keeps its part, timing included, byte-for-byte", async () => {
    const { bytes, slide2 } = await paragraphBuildDeck();
    const model = await readPresentationBytes(bytes, "pptx", "Deck");
    // Dirty another slide so the package is rebuilt around the untouched one.
    markSlideEdited(model.slides[0], "pptx");
    const { writePptx } = await import("./pptx/writer");
    const written = await unzip((await writePptx(model)).bytes);
    expect(dec(written.get("ppt/slides/slide2.xml")!)).toBe(slide2);
  });

  it("pptx: an edited slide re-emits the preserved timing byte-for-byte", async () => {
    const { bytes } = await paragraphBuildDeck();
    const model = await readPresentationBytes(bytes, "pptx", "Deck");
    markSlideEdited(model.slides[1], "pptx");
    const { writePptx } = await import("./pptx/writer");
    const saved = (await writePptx(model)).bytes;
    // The source element, plus only the `xmlns:p` it self-declares once
    // serialised standalone (the slide root binds the same prefix).
    const slideXml = dec((await unzip(saved)).get("ppt/slides/slide2.xml")!);
    expect(slideXml).toContain(PARAGRAPH_BUILD_TIMING.replace("<p:timing>", `<p:timing xmlns:p="${P_NS}">`));
    // Reopened, it is still the same preserved timing — never modelled.
    const back = (await readPresentationBytes(saved, "pptx", "Deck")).slides[1];
    expect(back.animations).toBeUndefined();
    expect(back.preserved?.["p:timing"]).toBe(model.slides[1].preserved?.["p:timing"]);
  });

  it("odp: a text-only (anim:sub-item) effect is preserved and an untouched page keeps its bytes", async () => {
    const parts = await unzip(await buildOdpAnimationsFixtureBytes());
    const content = dec(parts.get("content.xml")!);
    const target = `smil:targetElement="de-f1"`;
    const page2 = content.indexOf(`draw:name="page2"`);
    expect(content.indexOf(target, page2)).toBeGreaterThan(page2);
    const narrowed = content.slice(0, page2) + content.slice(page2).replace(target, `${target} anim:sub-item="text"`);
    parts.set("content.xml", enc(narrowed));
    const model = await readPresentationBytes(await zip([...parts].map(([name, data]) => ({ name, data }))), "odp", "Deck");
    expect(model.slides[1].animations).toBeUndefined();
    expect(hasPreservedTiming(model.slides[1])).toBe(true);

    markSlideEdited(model.slides[0], "odp");
    const { writeOdp } = await import("./odp/writer");
    const written = dec((await unzip((await writeOdp(model)).bytes)).get("content.xml")!);
    expect(page2Source(written)).toBe(page2Source(narrowed));
  });
});
