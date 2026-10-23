import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import type { SlideAnimation, SlideFrame, SlideParagraph } from "../model";
import {
  preservedTimingSpids,
  preservedTimingTextTargets,
  readPptxTiming,
  rebuiltParagraphRange,
  retargetPreservedTiming,
  writePptxTiming,
  type PreservedTextRange,
} from "./timing";
import {
  DEPTH_ANIMATIONS,
  depthFrameFor,
  GOLDEN_ANIMATIONS,
  goldenSpidFor,
  TIMING_DEPTH_GOLDEN,
  TIMING_GOLDEN,
} from "./timingGolden";

const SPID_NS = ` xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
function parseTiming(xml: string): Element {
  return parseXml(`<root${SPID_NS}>${xml}</root>`).documentElement.getElementsByTagName("*")[0] as Element;
}
/** frame id "fN" → cNvPr id N; the inverse of goldenSpidFor. */
function resolve(spid: string): string | undefined {
  return `f${spid}`;
}

// PowerPoint's entrance shape: the instant `p:set` (dur="1", its own start
// condition) precedes the timed behaviour. Its dur is not the effect's.
const PPT_VISIBILITY_SET =
  `<p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>` +
  `</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
function pptEffect(presetID: number, subtype: number, behaviours: string): string {
  return (
    `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
    `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="5" presetID="${presetID}" presetClass="entr" presetSubtype="${subtype}" fill="hold" grpId="0" nodeType="clickEffect">` +
    `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${PPT_VISIBILITY_SET}${behaviours}</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  );
}
function pptFade(dur: number): string {
  return `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="${dur}"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:animEffect>`;
}
const PPT_FLY_X =
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="7" dur="750" fill="hold"/>` +
  `<p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:tavLst><p:tav tm="0"><p:val><p:strVal val="0-#ppt_w/2"/></p:val></p:tav>` +
  `<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>`;

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

  it("reads the duration from the timed behaviour when PowerPoint writes the visibility set first", () => {
    expect(readPptxTiming(parseTiming(pptEffect(10, 0, pptFade(500))), resolve)).toEqual([
      { frameId: "f2", effect: "fadeIn", trigger: "click" },
    ]);
    expect(readPptxTiming(parseTiming(pptEffect(10, 0, pptFade(2000))), resolve)).toEqual([
      { frameId: "f2", effect: "fadeIn", trigger: "click", durationMs: 2000 },
    ]);
    expect(readPptxTiming(parseTiming(pptEffect(2, 8, PPT_FLY_X)), resolve)).toEqual([
      { frameId: "f2", effect: "flyIn", trigger: "click", durationMs: 750 },
    ]);
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

  it("preserves a by-paragraph build of a modelled preset — never a whole-frame replay per paragraph", () => {
    // One effect's target narrowed to a paragraph is enough.
    const byParagraph = TIMING_GOLDEN.replace(
      /<p:spTgt spid="(\d+)"\/>/,
      `<p:spTgt spid="$1"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt>`,
    );
    expect(byParagraph).not.toBe(TIMING_GOLDEN);
    expect(readPptxTiming(parseTiming(byParagraph), resolve)).toBeUndefined();
  });
});

/** The range re-pointer for targets whose text did not move. */
const sameRange = (_spid: string, range: PreservedTextRange): { st: number; end: number } => ({ st: range.st, end: range.end });
/** A bare timing around `targets`, and one `p:txEl`-narrowed target. */
const timingWith = (targets: string): string => `<p:timing><p:tnLst>${targets}</p:tnLst></p:timing>`;
const rangeTarget = (spid: string, range: string): string => `<p:tgtEl><p:spTgt spid="${spid}"><p:txEl>${range}</p:txEl></p:spTgt></p:tgtEl>`;

// SPEC-086 AC-4: a preserved (out-of-subset) timing re-pointed at a rebuild's ids.
describe("preserved pptx timing retargeting", () => {
  const TIMING =
    `<p:timing><p:tnLst><p:par><p:cTn id="1"><p:childTnLst>` +
    `<p:set><p:cBhvr><p:cTn id="2" dur="1"/><p:tgtEl><p:spTgt spid="12"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:set>` +
    `<p:set><p:cBhvr><p:cTn id="3" dur="1"/><p:tgtEl><p:spTgt spid="7"><p:subSp spid="{0A1B2C3D-0000-0000-0000-000000000001}"/></p:spTgt></p:tgtEl></p:cBhvr></p:set>` +
    `</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="12" grpId="0" build="p"/></p:bldLst></p:timing>`;

  it("lists the numeric shape references — not time-node ids, not a sub-shape GUID", () => {
    expect(preservedTimingSpids(TIMING)).toEqual(new Set(["12", "7"]));
  });

  it("re-points every shape reference at once (no chained renumbering)", () => {
    const map = new Map([["12", 7], ["7", 2]]);
    const out = retargetPreservedTiming(TIMING, (spid) => map.get(spid), sameRange);
    expect(out).toBe(
      TIMING.replace(`spTgt spid="12"`, `spTgt spid="7"`)
        .replace(`spTgt spid="7"><p:subSp`, `spTgt spid="2"><p:subSp`)
        .replace(`bldP spid="12"`, `bldP spid="7"`),
    );
  });

  it("returns undefined when any reference has no shape left", () => {
    expect(retargetPreservedTiming(TIMING, (spid) => (spid === "12" ? 3 : undefined), sameRange)).toBeUndefined();
    expect(retargetPreservedTiming(`<p:timing/>`, () => undefined, () => undefined)).toBe(`<p:timing/>`);
  });

  it("lists the shapes addressed by text range — not the whole-shape and sub-shape targets", () => {
    expect(preservedTimingTextTargets(TIMING)).toEqual(new Set(["12"]));
  });

  it("re-points a paragraph range by its SOURCE shape, in the same pass as the shape ids", () => {
    const map = new Map([["12", 7], ["7", 2]]);
    const seen: [string, PreservedTextRange][] = [];
    const out = retargetPreservedTiming(TIMING, (spid) => map.get(spid), (spid, range) => {
      seen.push([spid, { kind: range.kind, st: range.st, end: range.end }]);
      return { st: 1, end: 2 };
    });
    expect(seen).toEqual([["12", { kind: "pRg", st: 0, end: 0 }]]);
    expect(out).toBe(
      TIMING.replace(`<p:spTgt spid="12"><p:txEl><p:pRg st="0" end="0"/>`, `<p:spTgt spid="7"><p:txEl><p:pRg st="1" end="2"/>`)
        .replace(`spTgt spid="7"><p:subSp`, `spTgt spid="2"><p:subSp`)
        .replace(`bldP spid="12"`, `bldP spid="7"`),
    );
  });

  it("re-points every range of a by-paragraph build, and a character range by its own kind", () => {
    const timing = timingWith(rangeTarget("4", `<p:pRg st="0" end="0"/>`) + rangeTarget("4", `<p:pRg st="1" end="1"/>`) + rangeTarget("4", `<p:charRg st="3" end="9"/>`));
    expect(retargetPreservedTiming(timing, () => 4, (_spid, { st, end }) => ({ st: st + 10, end: end + 20 }))).toBe(
      timingWith(rangeTarget("4", `<p:pRg st="10" end="20"/>`) + rangeTarget("4", `<p:pRg st="11" end="21"/>`) + rangeTarget("4", `<p:charRg st="13" end="29"/>`)),
    );
  });

  it("returns undefined when a range has no text left, or is not a pair of indices", () => {
    expect(retargetPreservedTiming(TIMING, () => 3, () => undefined)).toBeUndefined();
    const bad = TIMING.replace(`<p:pRg st="0" end="0"/>`, `<p:pRg st="x" end="0"/>`);
    expect(retargetPreservedTiming(bad, () => 3, sameRange)).toBeUndefined();
  });
});

/** Lines read from source paragraphs (`undefined` = a line an edit added). */
const lines = (...sources: (number | undefined)[]): SlideParagraph[] =>
  sources.map((source) => (source === undefined ? { runs: [] } : { runs: [], pptxSourceParagraph: source }));
const pRange = (st: number, end: number): PreservedTextRange => ({ kind: "pRg", st, end });

describe("a source paragraph range on a rebuilt body", () => {
  it("spans every line an a:br split off the source paragraph, and shifts past it", () => {
    // Source: a:p 0 = "A<a:br/>B", a:p 1 = "C" → lines A, B, C.
    const body = lines(0, 0, 1);
    expect(rebuiltParagraphRange(body, pRange(0, 0))).toEqual({ st: 0, end: 1 });
    expect(rebuiltParagraphRange(body, pRange(1, 1))).toEqual({ st: 2, end: 2 });
    expect(rebuiltParagraphRange(body, pRange(0, 1))).toEqual({ st: 0, end: 2 });
  });

  it("follows lines an edit added or removed around it", () => {
    // Source 0 deleted, a new line before source 2.
    const body = lines(1, undefined, 2, 3);
    expect(rebuiltParagraphRange(body, pRange(1, 1))).toEqual({ st: 0, end: 0 });
    expect(rebuiltParagraphRange(body, pRange(2, 2))).toEqual({ st: 2, end: 2 });
    expect(rebuiltParagraphRange(body, pRange(3, 3))).toEqual({ st: 3, end: 3 });
    // A range keeps what is left of it (and a line added inside it).
    expect(rebuiltParagraphRange(body, pRange(0, 2))).toEqual({ st: 0, end: 2 });
  });

  it("is undefined when no line from the range is left, and for a character range", () => {
    const body = lines(1, undefined, 2);
    expect(rebuiltParagraphRange(body, pRange(0, 0))).toBeUndefined();
    expect(rebuiltParagraphRange(body, pRange(5, 7))).toBeUndefined();
    expect(rebuiltParagraphRange([], pRange(0, 0))).toBeUndefined();
    expect(rebuiltParagraphRange(body, { kind: "charRg", st: 0, end: 3 })).toBeUndefined();
  });
});

/** PowerPoint's markup for a one-click-per-part fade-in build on shape 4: each
 * effect's behaviours target `<p:spTgt spid="4">{part}</p:spTgt>` ("" = the
 * whole shape). `cTnExtra` lands in the effect cTn before its childTnLst (where
 * PowerPoint puts `p:iterate`). */
function pptPartBuild(parts: string[], cTnExtra = ""): string {
  const effects = parts
    .map((part) => {
      const tgt = `<p:tgtEl><p:spTgt spid="4">${part}</p:spTgt></p:tgtEl>`;
      return (
        `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
        `<p:par><p:cTn fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
        `<p:par><p:cTn presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">` +
        `<p:stCondLst><p:cond delay="0"/></p:stCondLst>${cTnExtra}<p:childTnLst>` +
        `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn dur="500"/>${tgt}</p:cBhvr></p:animEffect>` +
        `<p:set><p:cBhvr><p:cTn dur="1" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
        `<p:to><p:strVal val="visible"/></p:to></p:set>` +
        `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
      );
    })
    .join("");
  const build = parts.some((part) => part.includes("txEl")) ? ` build="p"` : "";
  return (
    `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${effects}</p:childTnLst></p:cTn>` +
    `<p:prevCondLst/><p:nextCondLst/></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>` +
    `<p:bldLst><p:bldP spid="4" grpId="0"${build}/></p:bldLst></p:timing>`
  );
}
const pRg = (i: number) => `<p:txEl><p:pRg st="${i}" end="${i}"/></p:txEl>`;

describe("pptx timing reader: frame-level targets only", () => {
  it("models a whole-shape build (the control for the cases below)", () => {
    expect(readPptxTiming(parseTiming(pptPartBuild([""])), resolve)).toEqual([
      { frameId: "f4", effect: "fadeIn", trigger: "click" },
    ]);
  });

  it("preserves a by-paragraph text build (p:txEl) rather than replaying the whole frame per paragraph", () => {
    expect(readPptxTiming(parseTiming(pptPartBuild([pRg(0), pRg(1)])), resolve)).toBeUndefined();
    // One paragraph-targeted effect taints the whole timing (never half-parse).
    expect(readPptxTiming(parseTiming(pptPartBuild(["", pRg(0)])), resolve)).toBeUndefined();
    // A character range is a sub-target too.
    const charRg = `<p:txEl><p:charRg st="0" end="4"/></p:txEl>`;
    expect(readPptxTiming(parseTiming(pptPartBuild([charRg])), resolve)).toBeUndefined();
  });

  it("preserves every other p:spTgt sub-target", () => {
    const subTargets = [
      `<p:bg/>`,
      `<p:subSp spid="_x0000_s1"/>`,
      `<p:oleChartEl type="series" lvl="1"/>`,
      `<p:graphicEl><a:chart xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" seriesIdx="0" categoryIdx="-3" bldStep="series"/></p:graphicEl>`,
    ];
    for (const sub of subTargets) {
      expect(readPptxTiming(parseTiming(pptPartBuild([sub])), resolve), sub).toBeUndefined();
    }
  });

  it("preserves a by-word / by-letter text build (p:iterate)", () => {
    for (const type of ["wd", "lt"]) {
      const iterate = `<p:iterate type="${type}"><p:tmPct val="10000"/></p:iterate>`;
      expect(readPptxTiming(parseTiming(pptPartBuild([""], iterate)), resolve)).toBeUndefined();
    }
  });
});

// ---------------------------------------------------------------------------
// SPEC-105: the depth delta (builds, emphasis, paths, repetition, triggers)
// ---------------------------------------------------------------------------

const DEPTH_NS = ` xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
function parseDepth(xml: string): Element {
  return parseXml(`<root${DEPTH_NS}>${xml}</root>`).documentElement.getElementsByTagName("*")[0] as Element;
}
/** Writes then reads `animations` over DEPTH_FRAMES. */
function depthRoundTrip(animations: SlideAnimation[]): SlideAnimation[] | undefined {
  return readPptxTiming(parseDepth(writePptxTiming(animations, goldenSpidFor, depthFrameFor)), resolve, depthFrameFor);
}
/** DEPTH_FRAMES with frame `id`'s paragraphs replaced. */
function framesWith(id: string, paragraphs: SlideParagraph[]): (frameId: string) => SlideFrame | undefined {
  return (frameId) => {
    const frame = depthFrameFor(frameId);
    return frame && frameId === id && frame.kind === "text" ? { ...frame, paragraphs } : frame;
  };
}

describe("SPEC-105 depth golden", () => {
  it("serialises DEPTH_ANIMATIONS over DEPTH_FRAMES to TIMING_DEPTH_GOLDEN", () => {
    expect(writePptxTiming(DEPTH_ANIMATIONS, goldenSpidFor, depthFrameFor)).toBe(TIMING_DEPTH_GOLDEN);
  });

  it("reads TIMING_DEPTH_GOLDEN back to DEPTH_ANIMATIONS, interactive rows back in their model place", () => {
    expect(readPptxTiming(parseDepth(TIMING_DEPTH_GOLDEN), resolve, depthFrameFor)).toEqual(DEPTH_ANIMATIONS);
  });

  it("keeps the SPEC-086 golden byte-identical", () => {
    expect(writePptxTiming(GOLDEN_ANIMATIONS, goldenSpidFor, depthFrameFor)).toBe(TIMING_GOLDEN);
  });

  it("never claims emph:8, so SPEC-086's exotic emphasis tree stays out of subset", () => {
    const exotic = TIMING_GOLDEN.replace('presetID="1" presetClass="entr"', 'presetID="8" presetClass="emph"');
    expect(exotic).not.toBe(TIMING_GOLDEN);
    expect(readPptxTiming(parseDepth(exotic), resolve, depthFrameFor)).toBeUndefined();
  });
});

describe("SPEC-105 per-feature round trips", () => {
  const forms: SlideAnimation[][] = [
    [{ frameId: "f2", effect: "pulse", trigger: "click" }],
    [{ frameId: "f2", effect: "spin", trigger: "click", spinDegrees: 90, durationMs: 1200 }],
    [{ frameId: "f2", effect: "spin", trigger: "click", spinDegrees: -3600 }],
    [{ frameId: "f6", effect: "growShrink", trigger: "click", scalePercent: 10 }],
    [{ frameId: "f6", effect: "growShrink", trigger: "click", scalePercent: 400 }],
    [{ frameId: "f6", effect: "colorChange", trigger: "click", emphasisColor: "#123abc", durationMs: 2000 }],
    [{ frameId: "f6", effect: "colorChange", trigger: "click" }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "line" } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "line", direction: "down", distanceFraction: 1 } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "arc", direction: "up", distanceFraction: 0.05 } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "turn", direction: "left", distanceFraction: 0.4 } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "loop" } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "custom", d: "M 0 0 L 0.1 0.1 L 0.2 0" } }],
    [{ frameId: "f2", effect: "fadeIn", trigger: "click", repeatCount: 10, autoReverse: true, rewindOnComplete: true }],
    [{ frameId: "f2", effect: "spin", trigger: "click", repeatUntilClick: true }],
    [{ frameId: "f3", effect: "zoomIn", trigger: "click", buildBy: "byParagraph", buildTrigger: "afterPrevious", delayMs: 300 }],
    [{ frameId: "f4", effect: "wipeIn", direction: "down", trigger: "click", buildBy: "byFirstLevelParagraph" }],
    [
      { frameId: "f2", effect: "appear", trigger: "click" },
      { frameId: "f3", effect: "flyOut", trigger: "withPrevious", buildBy: "byParagraph", buildTrigger: "withPrevious" },
    ],
    [{ frameId: "f2", effect: "pulse", trigger: "onClickOf", triggerFrameId: "f5" }],
    [
      { frameId: "f3", effect: "fadeIn", trigger: "onClickOf", triggerFrameId: "f5", buildBy: "byParagraph", buildTrigger: "withPrevious" },
      { frameId: "f2", effect: "fadeIn", trigger: "click" },
      { frameId: "f6", effect: "spin", trigger: "onClickOf", triggerFrameId: "f5" },
    ],
  ];
  for (const form of forms) {
    it(`${form.map((a) => `${a.effect}${a.buildBy ? `/${a.buildBy}` : ""}${a.trigger === "onClickOf" ? "/onClickOf" : ""}`).join(" + ")}`, () => {
      expect(depthRoundTrip(form)).toEqual(form);
    });
  }

  it("expands builds over the frame's CURRENT paragraphs — more lines, more steps, nothing stored", () => {
    const row: SlideAnimation[] = [{ frameId: "f3", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" }];
    const two = framesWith("f3", [{ runs: [{ text: "a" }] }, { runs: [{ text: "b" }] }]);
    const xml = writePptxTiming(row, goldenSpidFor, two);
    expect(xml.match(/<p:pRg /g)).toHaveLength(4); // two steps × (effect + visibility set)
    expect(readPptxTiming(parseDepth(xml), resolve, two)).toEqual(row);
    // Under one paragraph the row plays (and writes) whole, with no p:bldP.
    const one = framesWith("f3", [{ runs: [{ text: "a" }] }]);
    const whole = writePptxTiming(row, goldenSpidFor, one);
    expect(whole).not.toContain("<p:txEl>");
    expect(whole).not.toContain("<p:bldLst>");
  });
});

/** A PowerPoint-authored by-paragraph build: every cTn numbered in document
 * order, the visibility `p:set` first with its own start condition, the
 * attributes reordered, the p:bldP's bldLvl left at its default. */
function pptByParagraph(spid: number, ranges: [number, number][], bldP: string): string {
  let id = 3;
  const effects = ranges
    .map(([st, end]) => {
      const tgt = `<p:tgtEl><p:spTgt spid="${spid}"><p:txEl><p:pRg st="${st}" end="${end}"/></p:txEl></p:spTgt></p:tgtEl>`;
      return (
        `<p:par><p:cTn id="${id++}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
        `<p:par><p:cTn id="${id++}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
        `<p:par><p:cTn id="${id++}" nodeType="clickEffect" grpId="0" fill="hold" presetSubtype="0" presetClass="entr" presetID="10">` +
        `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
        `<p:set><p:cBhvr><p:cTn id="${id++}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}` +
        `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
        `<p:animEffect filter="fade" transition="in"><p:cBhvr><p:cTn id="${id++}" dur="500"/>${tgt}</p:cBhvr></p:animEffect>` +
        `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
      );
    })
    .join("");
  return (
    `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${effects}</p:childTnLst></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
    `</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${bldP}</p:bldLst></p:timing>`
  );
}

/** One PowerPoint motion-path effect on shape 6 with `path`. */
function pptPath(presetID: number, path: string): string {
  return TIMING_GOLDEN.replace(
    /<p:par><p:cTn presetID="1" presetClass="entr"[\s\S]*?<\/p:cTn><\/p:par>/,
    `<p:par><p:cTn id="9" presetID="${presetID}" presetClass="path" presetSubtype="0" accel="50000" decel="50000" fill="hold" nodeType="clickEffect">` +
      `<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" rAng="0" ptsTypes="AA">` +
      `<p:cBhvr><p:cTn id="10" dur="2000" fill="hold"/><p:tgtEl><p:spTgt spid="6"/></p:tgtEl>` +
      `<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>` +
      `<p:rCtr x="0" y="0"/></p:animMotion></p:childTnLst></p:cTn></p:par>`,
  );
}

describe("SPEC-105 tolerant reading (PowerPoint-flavoured)", () => {
  it("models PowerPoint's numbered, reordered by-paragraph build — document order, visibility set first", () => {
    const xml = pptByParagraph(3, [[0, 0], [1, 1], [2, 2], [3, 3]], `<p:bldP spid="3" grpId="0" build="p"/>`);
    expect(readPptxTiming(parseDepth(xml), resolve, depthFrameFor)).toEqual([
      { frameId: "f3", effect: "fadeIn", trigger: "click", buildBy: "byParagraph" },
    ]);
  });

  it("reads a default-level (first-level) build over sub-points as by first-level paragraph", () => {
    const xml = pptByParagraph(4, [[0, 2], [3, 3]], `<p:bldP spid="4" grpId="0" build="p"/>`);
    expect(readPptxTiming(parseDepth(xml), resolve, depthFrameFor)).toEqual([
      { frameId: "f4", effect: "fadeIn", trigger: "click", buildBy: "byFirstLevelParagraph" },
    ]);
  });

  it("reads a motion path with PowerPoint's float noise and E end marker as its preset", () => {
    const xml = pptPath(42, "M -4.16667E-6 -4.07407E-6 L 0.25 0 E");
    const first = readPptxTiming(parseDepth(xml), resolve, depthFrameFor)?.[0];
    expect(first).toEqual({ frameId: "f6", effect: "motionPath", trigger: "click", durationMs: 2000, path: { preset: "line" } });
  });

  it("reads PowerPoint's custom (drawn) path as a custom path in canonical form", () => {
    const xml = pptPath(0, "M 0 0 C 0.05 0.1 0.1 0.1 0.15 0 E");
    expect(readPptxTiming(parseDepth(xml), resolve, depthFrameFor)?.[0].path).toEqual({
      preset: "custom",
      d: "M 0 0 C 0.05 0.1 0.1 0.1 0.15 0",
    });
  });
});

describe("SPEC-105 out of subset (→ preserved)", () => {
  const read = (xml: string, frameFor = depthFrameFor): SlideAnimation[] | undefined =>
    readPptxTiming(parseDepth(xml), resolve, frameFor);

  it("a partial p:pRg cover — a gap, an overlap, out of order", () => {
    const bld = `<p:bldP spid="3" grpId="0" build="p"/>`;
    expect(read(pptByParagraph(3, [[0, 0], [1, 1], [3, 3]], bld))).toBeUndefined();
    expect(read(pptByParagraph(3, [[0, 1], [1, 1], [2, 2], [3, 3]], bld))).toBeUndefined();
    expect(read(pptByParagraph(3, [[1, 1], [0, 0], [2, 2], [3, 3]], bld))).toBeUndefined();
    expect(read(pptByParagraph(3, [[0, 0], [1, 1], [2, 2]], bld))).toBeUndefined();
  });

  it("a p:bldP disagreeing with the effect pars", () => {
    const ranges: [number, number][] = [[0, 0], [1, 1], [2, 2], [3, 3]];
    // No p:bldP at all, or one that builds the shape whole.
    expect(read(pptByParagraph(3, ranges, ""))).toBeUndefined();
    expect(read(pptByParagraph(3, ranges, `<p:bldP spid="3" grpId="0"/>`))).toBeUndefined();
    // First-level declared, every paragraph built (f4 has sub-points).
    expect(read(pptByParagraph(4, ranges, `<p:bldP spid="4" grpId="0" build="p" bldLvl="1"/>`))).toBeUndefined();
    // A paragraph build declared for a shape the effects animate whole.
    const whole = writePptxTiming([{ frameId: "f3", effect: "fadeIn", trigger: "click" }], goldenSpidFor, depthFrameFor);
    expect(read(whole.replace("</p:tnLst>", `</p:tnLst><p:bldLst><p:bldP spid="3" grpId="0" build="p"/></p:bldLst>`))).toBeUndefined();
  });

  it("a frame whose lines are not its source paragraphs one-for-one (an a:br split)", () => {
    const xml = pptByParagraph(3, [[0, 0], [1, 1], [2, 2], [3, 3]], `<p:bldP spid="3" grpId="0" build="p"/>`);
    const split = framesWith("f3", [
      { runs: [{ text: "a" }], pptxSourceParagraph: 0 },
      { runs: [{ text: "b" }], pptxSourceParagraph: 0 },
      { runs: [{ text: "c" }], pptxSourceParagraph: 1 },
      { runs: [{ text: "d" }], pptxSourceParagraph: 2 },
    ]);
    expect(read(xml, split)).toBeUndefined();
  });

  it("an unresolvable trigger, or a trigger that is the target itself", () => {
    const xml = writePptxTiming([{ frameId: "f2", effect: "spin", trigger: "onClickOf", triggerFrameId: "f5" }], goldenSpidFor, depthFrameFor);
    expect(readPptxTiming(parseDepth(xml), (spid) => (spid === "5" ? undefined : `f${spid}`), depthFrameFor)).toBeUndefined();
    expect(read(xml.replace('<p:spTgt spid="5"/></p:tgtEl></p:cond></p:stCondLst>', '<p:spTgt spid="2"/></p:tgtEl></p:cond></p:stCondLst>'))).toBeUndefined();
  });

  it("a second sequence that is not an interactive one", () => {
    const xml = writePptxTiming(
      [
        { frameId: "f3", effect: "fadeIn", trigger: "click" },
        { frameId: "f2", effect: "spin", trigger: "onClickOf", triggerFrameId: "f5" },
      ],
      goldenSpidFor,
      depthFrameFor,
    );
    expect(read(xml)).toHaveLength(2); // the control
    expect(read(xml.replace('nodeType="interactiveSeq"', 'nodeType="mainSeq"'))).toBeUndefined();
    expect(read(xml.replace(' nodeType="interactiveSeq"', ""))).toBeUndefined();
    const twoMain = TIMING_GOLDEN.replace("</p:seq>", `</p:seq>${TIMING_GOLDEN.slice(TIMING_GOLDEN.indexOf("<p:seq "), TIMING_GOLDEN.indexOf("</p:seq>") + 8)}`);
    expect(read(twoMain)).toBeUndefined();
  });

  it("a malformed path, or one its preset does not reproduce", () => {
    expect(read(pptPath(0, "M 0 0 Q 0.1 0.1 0.2 0"))).toBeUndefined();
    expect(read(pptPath(0, "L 0.1 0"))).toBeUndefined();
    expect(read(pptPath(42, "M 0 0 L 0.2 0.2"))).toBeUndefined(); // a diagonal is no line preset
    expect(read(pptPath(40, "M 0 0 L 0.2 0"))).toBeUndefined(); // not a loop
    expect(read(pptPath(99, "M 0 0 L 0.25 0"))).toBeUndefined(); // unknown path preset
  });

  it("an emphasis id over behaviours this table does not write", () => {
    const spin = writePptxTiming([{ frameId: "f2", effect: "spin", trigger: "click" }], goldenSpidFor, depthFrameFor);
    expect(read(spin.replace(/<p:animRot by="\d+">/, "<p:animScale>").replace("</p:animRot>", "</p:animScale>"))).toBeUndefined();
    const color = writePptxTiming([{ frameId: "f6", effect: "colorChange", trigger: "click" }], goldenSpidFor, depthFrameFor);
    expect(read(color.replaceAll(' autoRev="1"', ""))).toBeUndefined(); // PowerPoint's held fill colour
    const grow = writePptxTiming([{ frameId: "f6", effect: "growShrink", trigger: "click" }], goldenSpidFor, depthFrameFor);
    expect(read(grow.replaceAll("#ppt_w*1.5", "#ppt_w*9"))).toBeUndefined(); // 900%: out of bounds
  });

  it("a repeat count that is not a whole number of passes, or indefinite with no click to end it", () => {
    const xml = writePptxTiming([{ frameId: "f2", effect: "spin", trigger: "click", repeatCount: 2 }], goldenSpidFor, depthFrameFor);
    expect(read(xml.replace('repeatCount="2000"', 'repeatCount="1500"'))).toBeUndefined();
    expect(read(xml.replace('repeatCount="2000"', 'repeatCount="indefinite"'))).toBeUndefined();
  });
});
