import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import type { SlideAnimation } from "../model";
import { readPptxTiming, writePptxTiming } from "./timing";
import { GOLDEN_ANIMATIONS, goldenSpidFor, TIMING_GOLDEN } from "./timingGolden";

const SPID_NS = ` xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
function parseTiming(xml: string): Element {
  return parseXml(`<root${SPID_NS}>${xml}</root>`).documentElement.getElementsByTagName("*")[0] as Element;
}
/** frame id "fN" → cNvPr id N; the inverse of goldenSpidFor. */
function resolve(spid: string): string | undefined {
  return `f${spid}`;
}

describe("pptx timing serialiser", () => {
  it("serialises the golden animations to the golden bytes", () => {
    expect(writePptxTiming(GOLDEN_ANIMATIONS, goldenSpidFor)).toBe(TIMING_GOLDEN);
  });

  it("emits nothing for an empty or absent animation list", () => {
    expect(writePptxTiming([], goldenSpidFor)).toBe("");
    expect(writePptxTiming(undefined, goldenSpidFor)).toBe("");
  });

  it("reads the golden back to the model (canonical read-back)", () => {
    const el = parseTiming(TIMING_GOLDEN);
    expect(readPptxTiming(el, resolve)).toEqual(GOLDEN_ANIMATIONS);
  });

  it("round-trips every single-effect form", () => {
    const cases: SlideAnimation[] = [
      { frameId: "f2", effect: "appear", trigger: "click" },
      { frameId: "f2", effect: "fadeIn", trigger: "click", durationMs: 300 },
      { frameId: "f2", effect: "flyIn", direction: "up", trigger: "click" },
      { frameId: "f2", effect: "wipeIn", direction: "down", trigger: "click", delayMs: 100 },
      { frameId: "f2", effect: "zoomIn", trigger: "click" },
      { frameId: "f2", effect: "fadeOut", trigger: "click" },
      { frameId: "f2", effect: "flyOut", direction: "right", trigger: "click" },
    ];
    for (const animation of cases) {
      const xml = writePptxTiming([animation], goldenSpidFor);
      const back = readPptxTiming(parseTiming(xml), resolve);
      expect(back).toEqual([animation]);
    }
  });

  it("tolerates PowerPoint-flavoured markup with reordered attributes", () => {
    // Same fade-in effect, attributes reordered and an extra behaviour attr the
    // reader must ignore.
    const xml =
      `<p:timing><p:tnLst><p:par>` +
      `<p:cTn nodeType="tmRoot" dur="indefinite" restart="never" id="1"><p:childTnLst>` +
      `<p:seq nextAc="seek" concurrent="1"><p:cTn nodeType="mainSeq" id="2" dur="indefinite"><p:childTnLst>` +
      `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:par><p:cTn presetClass="entr" nodeType="clickEffect" presetSubtype="0" fill="hold" presetID="10">` +
      `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:childTnLst><p:animEffect filter="fade" transition="in"><p:cBhvr>` +
      `<p:cTn fill="hold" dur="500"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:animEffect>` +
      `</p:childTnLst></p:cTn></p:par>` +
      `</p:childTnLst></p:cTn></p:par>` +
      `</p:childTnLst></p:cTn></p:par>` +
      `</p:childTnLst></p:cTn><p:prevCondLst/><p:nextCondLst/></p:seq>` +
      `</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    expect(readPptxTiming(parseTiming(xml), resolve)).toEqual([{ frameId: "f2", effect: "fadeIn", trigger: "click" }]);
  });

  it("preserves (returns undefined) for out-of-subset timing", () => {
    const unknownPreset = writePptxTiming(GOLDEN_ANIMATIONS, goldenSpidFor).replace('presetID="1"', 'presetID="99"');
    expect(readPptxTiming(parseTiming(unknownPreset), resolve)).toBeUndefined();

    const withAudio = writePptxTiming(GOLDEN_ANIMATIONS, goldenSpidFor).replace(
      "</p:seq>",
      "</p:seq><p:audio><p:cMediaNode/></p:audio>",
    );
    expect(readPptxTiming(parseTiming(withAudio), resolve)).toBeUndefined();

    // An unresolvable target (the slide doesn't contain that shape) → preserve.
    expect(readPptxTiming(parseTiming(TIMING_GOLDEN), () => undefined)).toBeUndefined();
  });
});
