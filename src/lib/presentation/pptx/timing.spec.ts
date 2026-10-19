import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import type { SlideAnimation, SlideParagraph } from "../model";
import {
  preservedTimingSpids,
  preservedTimingTextTargets,
  readPptxTiming,
  rebuiltParagraphRange,
  retargetPreservedTiming,
  writePptxTiming,
  type PreservedTextRange,
} from "./timing";
import { GOLDEN_ANIMATIONS, goldenSpidFor, TIMING_GOLDEN } from "./timingGolden";

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
