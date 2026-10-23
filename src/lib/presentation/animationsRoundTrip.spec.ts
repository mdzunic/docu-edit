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
  buildOdpDepthFixtureBytes,
  buildOdpFixtureBytes,
  buildPptxAnimationsFixtureBytes,
  buildPptxDepthFixtureBytes,
  buildPptxFixtureBytes,
  PPTX_CANONICAL_TIMING,
} from "./fixtures/build";
import { equalStructural } from "../equality";
import { equalBytes } from "../core/equality";
import { readPresentationBytes } from "./open";
import { markSlideEdited } from "./text";
import { hasPreservedTiming } from "./animations";
import type { PresentationModel, Slide, SlideAnimation } from "./model";

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
 * paragraph build, and its body given the second paragraph the build targets.
 * `uncovered` adds a third paragraph the build's ranges leave out — a partial
 * cover, out of SPEC-105's subset, so the build stays preserved (a build
 * covering every paragraph is modelled as one by-paragraph row). */
async function paragraphBuildDeck(uncovered = false): Promise<{ bytes: Uint8Array; slide2: string }> {
  const parts = await unzip(await buildPptxAnimationsFixtureBytes());
  const source = dec(parts.get("ppt/slides/slide2.xml")!);
  const bodyLine = "<a:t>Body line</a:t></a:r></a:p>";
  expect(source).toContain(PPTX_CANONICAL_TIMING);
  expect(source).toContain(bodyLine);
  const extra = uncovered ? `<a:p><a:r><a:t>Third line</a:t></a:r></a:p>` : "";
  const slide2 = source
    .replace(bodyLine, `${bodyLine}<a:p><a:r><a:t>Second line</a:t></a:r></a:p>${extra}`)
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
    // A partial cover (SPEC-105 models only a build covering every paragraph).
    const { bytes } = await paragraphBuildDeck(true);
    const slide = (await readPresentationBytes(bytes, "pptx", "Deck")).slides[1];
    expect(slide.animations).toBeUndefined();
    expect(hasPreservedTiming(slide)).toBe(true);
  });

  it("pptx: a build covering every paragraph models as ONE by-paragraph row, never N whole-frame rows (SPEC-105)", async () => {
    const { bytes } = await paragraphBuildDeck();
    const slide = (await readPresentationBytes(bytes, "pptx", "Deck")).slides[1];
    expect(slide.animations).toEqual([
      { frameId: slide.frames[1].id, effect: "fadeIn", trigger: "click", buildBy: "byParagraph" },
    ]);
    expect(hasPreservedTiming(slide)).toBe(false);
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
    const { bytes } = await paragraphBuildDeck(true);
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

// ---------------------------------------------------------------------------
// SPEC-105 round trips (AC-49/50 and the no-edit preserve path)
// ---------------------------------------------------------------------------

/** Slide animations with every frame reference as its z-order index — frame
 * ids are fresh UUIDs per read, and neither key is in IGNORED_KEYS. */
function byIndex(slide: Slide): unknown[] {
  const index = (id: string): number => slide.frames.findIndex((f) => f.id === id);
  return (slide.animations ?? []).map(({ frameId, triggerFrameId, ...rest }) =>
    triggerFrameId === undefined ? { ...rest, frameId: index(frameId) } : { ...rest, frameId: index(frameId), triggerFrameId: index(triggerFrameId) },
  );
}

/** Every new effect on the depth fixture's bullet slide (title 0, body 1,
 * rounded rectangle 2): a by-paragraph build, a repeat + auto-reverse +
 * rewind row, onClickOf rows between main rows, a motion path. */
function depthSample(slide: Slide): SlideAnimation[] {
  const [title, body, box] = slide.frames.map((f) => f.id);
  return [
    { frameId: body, effect: "fadeIn", trigger: "click", buildBy: "byParagraph" },
    { frameId: title, effect: "pulse", trigger: "withPrevious", repeatCount: 3, autoReverse: true, rewindOnComplete: true },
    { frameId: title, effect: "spin", trigger: "onClickOf", triggerFrameId: box, spinDegrees: -90 },
    { frameId: box, effect: "growShrink", trigger: "afterPrevious", scalePercent: 180, delayMs: 200 },
    { frameId: box, effect: "colorChange", trigger: "click", durationMs: 800, emphasisColor: "#3366cc", repeatUntilClick: true },
    { frameId: box, effect: "motionPath", trigger: "withPrevious", path: { preset: "arc", direction: "up", distanceFraction: 0.4 } },
    { frameId: body, effect: "fadeOut", trigger: "onClickOf", triggerFrameId: box, buildBy: "byParagraph", buildTrigger: "withPrevious" },
    { frameId: box, effect: "motionPath", trigger: "click", path: { preset: "loop" } },
  ];
}

describe("SPEC-105 AC-49/50: every new effect round-trips in both formats", () => {
  for (const [format, build] of [
    ["pptx", buildPptxDepthFixtureBytes],
    ["odp", buildOdpDepthFixtureBytes],
  ] as const) {
    async function writer(): Promise<(model: PresentationModel) => Promise<{ bytes: Uint8Array }>> {
      return format === "pptx" ? (await import("./pptx/writer")).writePptx : (await import("./odp/writer")).writeOdp;
    }
    const slidePart = format === "pptx" ? "ppt/slides/slide1.xml" : "content.xml";

    it(`${format}: save → reopen reproduces slide.animations exactly`, async () => {
      const source = await build();
      const model = await readPresentationBytes(source, format, "Deck");
      const slide = model.slides[0];
      slide.animations = depthSample(slide);
      markSlideEdited(slide, format);
      const saved = (await (await writer())(model)).bytes;
      const back = (await readPresentationBytes(saved, format, "Deck")).slides[0];
      expect(equalStructural(byIndex(back), byIndex(slide))).toBe(true);
      expect(byIndex(back)).toEqual(byIndex(slide));
      expect(back.animations).toEqual(depthSample(back));
      expect(hasPreservedTiming(back)).toBe(false);
    });

    it(`${format}: every part the edit did not touch is byte-equal; the reopened model writes stable bytes`, async () => {
      const source = await build();
      const model = await readPresentationBytes(source, format, "Deck");
      model.slides[0].animations = depthSample(model.slides[0]);
      markSlideEdited(model.slides[0], format);
      const write = await writer();
      const before = await unzip(source);
      const after = await unzip((await write(model)).bytes);
      expect([...after.keys()].toSorted()).toEqual([...before.keys()].toSorted());
      for (const [name, bytes] of before) {
        if (name !== slidePart) {
          expect(equalBytes(after.get(name)!, bytes), name).toBe(true);
        }
      }
      // Reopened and rebuilt, the slide part comes out the same twice.
      const reopened = await readPresentationBytes((await write(model)).bytes, format, "Deck");
      markSlideEdited(reopened.slides[0], format);
      const first = (await unzip((await write(reopened)).bytes)).get(slidePart)!;
      markSlideEdited(reopened.slides[0], format);
      const second = (await unzip((await write(reopened)).bytes)).get(slidePart)!;
      expect(equalBytes(first, second)).toBe(true);
      expect(dec(first)).toContain(format === "pptx" ? `nodeType="interactiveSeq"` : `presentation:node-type="interactive-sequence"`);
    });

    it(`${format}: a no-edit download is byte-identical part-for-part, the out-of-subset slide included`, async () => {
      const source = await build();
      const model = await readPresentationBytes(source, format, "Deck");
      expect(model.slides[3].animations).toBeUndefined();
      expect(hasPreservedTiming(model.slides[3])).toBe(true);
      const before = await unzip(source);
      const after = await unzip((await (await writer())(model)).bytes);
      expect([...after.keys()].toSorted()).toEqual([...before.keys()].toSorted());
      for (const [name, bytes] of before) {
        expect(equalBytes(after.get(name)!, bytes), name).toBe(true);
      }
    });

    it(`${format}: editing another slide keeps the out-of-subset slide's timing verbatim`, async () => {
      const model = await readPresentationBytes(await build(), format, "Deck");
      const preserved = { ...model.slides[3].preserved };
      markSlideEdited(model.slides[0], format);
      markSlideEdited(model.slides[3], format);
      const back = await readPresentationBytes((await (await writer())(model)).bytes, format, "Deck");
      expect(back.slides[3].animations).toBeUndefined();
      const key = format === "pptx" ? "p:timing" : "anim";
      expect(back.slides[3].preserved?.[key]).toBe(preserved[key]);
    });
  }

  it("odp: a built frame's paragraphs carry their de-f{n}-p{k} text:id, other pages none", async () => {
    const model = await readPresentationBytes(await buildOdpDepthFixtureBytes(), "odp", "Deck");
    model.slides[0].animations = depthSample(model.slides[0]);
    markSlideEdited(model.slides[0], "odp");
    const { writeOdp } = await import("./odp/writer");
    const content = dec((await unzip((await writeOdp(model)).bytes)).get("content.xml")!);
    const page1 = content.slice(content.indexOf("<draw:page"), content.indexOf("</draw:page>"));
    for (let k = 0; k < 4; k++) {
      expect(page1).toContain(`text:id="de-f2-p${k}" xml:id="de-f2-p${k}"`);
    }
    expect(page1).not.toContain("de-f1-p");
  });
});
