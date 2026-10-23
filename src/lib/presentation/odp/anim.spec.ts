import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import type { SlideAnimation } from "../model";
import { paragraphWireId, preservedAnimReferences, readOdpAnim, writeOdpAnim, xmlIdsDefinedIn, type OdpTarget } from "./anim";
import { ANIM_DEPTH_GOLDEN, ANIM_GOLDEN, DEPTH_ANIMATIONS, depthFrameFor, GOLDEN_ANIMATIONS, goldenTargetFor } from "./animGolden";

const NS =
  ` xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"` +
  ` xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"` +
  ` xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"`;

/** Parses an `anim:*` fragment and returns its top-level element(s). */
function parseAnim(xml: string): Element[] {
  const root = parseXml(`<root${NS}>${xml}</root>`).documentElement;
  return Array.from(root.children);
}
/** wire id "de-fN" → frame id "fN"; the inverse of goldenTargetFor. */
function resolve(wireId: string): string | undefined {
  return wireId.startsWith("de-") ? wireId.slice(3) : undefined;
}

describe("odp animation serialiser", () => {
  it("serialises the golden animations to the golden bytes", () => {
    expect(writeOdpAnim(GOLDEN_ANIMATIONS, goldenTargetFor)).toBe(ANIM_GOLDEN);
  });

  it("emits nothing for an empty or absent animation list", () => {
    expect(writeOdpAnim([], goldenTargetFor)).toBe("");
    expect(writeOdpAnim(undefined, goldenTargetFor)).toBe("");
  });

  it("reads the golden back to the model (canonical read-back)", () => {
    expect(readOdpAnim(parseAnim(ANIM_GOLDEN), resolve)).toEqual(GOLDEN_ANIMATIONS);
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
      const xml = writeOdpAnim([animation], goldenTargetFor);
      expect(readOdpAnim(parseAnim(xml), resolve)).toEqual([animation]);
    }
  });

  it("tolerates Impress-flavoured markup with reordered attributes", () => {
    const xml =
      `<anim:par presentation:node-type="timing-root">` +
      `<anim:seq presentation:node-type="main-sequence">` +
      `<anim:par smil:begin="next"><anim:par smil:begin="0s">` +
      `<anim:par presentation:preset-id="ooo-entrance-fade-in" presentation:preset-class="entrance"` +
      ` smil:begin="0s" presentation:node-type="on-click">` +
      `<anim:transitionFilter smil:targetElement="de-f2" smil:dur="0.5s" smil:type="fade" smil:direction="forward"/>` +
      `</anim:par></anim:par></anim:par></anim:seq></anim:par>`;
    expect(readOdpAnim(parseAnim(xml), resolve)).toEqual([{ frameId: "f2", effect: "fadeIn", trigger: "click" }]);
  });

  it("preserves (returns undefined) for out-of-subset trees", () => {
    const unknown = ANIM_GOLDEN.replace("ooo-entrance-appear", "ooo-entrance-spiral");
    expect(readOdpAnim(parseAnim(unknown), resolve)).toBeUndefined();
    // Unresolvable target → preserve.
    expect(readOdpAnim(parseAnim(ANIM_GOLDEN), () => undefined)).toBeUndefined();
    // Not a timing-root → preserve.
    expect(readOdpAnim(parseAnim(`<anim:par presentation:node-type="other"/>`), resolve)).toBeUndefined();
  });
});

/** An Impress-shaped one-effect fade-in tree whose behaviours are `behaviours`. */
function impressFade(behaviours: string): string {
  return (
    `<anim:par presentation:node-type="timing-root"><anim:seq presentation:node-type="main-sequence">` +
    `<anim:par smil:begin="next"><anim:par smil:begin="0s">` +
    `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="entrance"` +
    ` presentation:preset-id="ooo-entrance-fade-in" presentation:group-id="0">${behaviours}</anim:par>` +
    `</anim:par></anim:par></anim:seq></anim:par>`
  );
}
function fadeOn(target: string, attrs = ""): string {
  return `<anim:transitionFilter smil:dur="0.5s" smil:targetElement="${target}"${attrs} smil:type="fade" smil:direction="forward"/>`;
}

describe("odp animation reader: frame-level targets only", () => {
  it("models a whole-frame effect, with or without an explicit whole sub-item", () => {
    const expected = [{ frameId: "f2", effect: "fadeIn", trigger: "click" }];
    expect(readOdpAnim(parseAnim(impressFade(fadeOn("de-f2"))), resolve)).toEqual(expected);
    expect(readOdpAnim(parseAnim(impressFade(fadeOn("de-f2", ` anim:sub-item="whole"`))), resolve)).toEqual(expected);
  });

  it("preserves an effect narrowed to the frame's text or background (anim:sub-item)", () => {
    for (const subItem of ["text", "background"]) {
      const xml = impressFade(fadeOn("de-f2", ` anim:sub-item="${subItem}"`));
      expect(readOdpAnim(parseAnim(xml), resolve), subItem).toBeUndefined();
    }
  });

  it("preserves a by-paragraph build: its targets are text:p ids, not frames", () => {
    expect(readOdpAnim(parseAnim(impressFade(fadeOn("id1", ` anim:sub-item="text"`))), resolve)).toBeUndefined();
  });

  it("preserves a by-word / by-letter text build (anim:iterate)", () => {
    for (const type of ["by-word", "by-letter"]) {
      const iterate =
        `<anim:iterate smil:targetElement="de-f2" anim:iterate-type="${type}" anim:iterate-interval="0.05s">` +
        `<anim:transitionFilter smil:dur="0.5s" smil:type="fade" smil:direction="forward"/></anim:iterate>`;
      expect(readOdpAnim(parseAnim(impressFade(iterate)), resolve), type).toBeUndefined();
    }
  });
});

// SPEC-086 AC-4: what a preserved (out-of-subset) anim string references.
describe("preserved odp anim references", () => {
  it("collects targetElement, master-element and id-qualified begin/end values, under any prefix", () => {
    const anim =
      `<anim:seq presentation:node-type="interactive-sequence" smil:begin="id9.click" smil:end="id9.click">` +
      `<anim:par smil:begin="next"><anim:par smil:begin="0s; id7.end+0.5s" presentation:master-element="id6">` +
      `<anim:set ns2:targetElement='id4' smil:begin="indefinite"/></anim:par></anim:par></anim:seq>`;
    expect(preservedAnimReferences(anim)).toEqual(new Set(["id9", "id7", "id6", "id4"]));
  });

  it("collects the ids a fragment defines, on any attribute whose local name is id", () => {
    const xml =
      `<draw:frame xml:id="id3" draw:id="id3" presentation:preset-id="ooo-x">` +
      `<text:p xml:id="id4"/></draw:frame><anim:par anim:id="n1" smil:targetElement="id9"/>`;
    expect(xmlIdsDefinedIn(xml)).toEqual(new Set(["id3", "id4", "n1"]));
  });
});

// ---------------------------------------------------------------------------
// SPEC-105: the depth delta (builds, emphasis, paths, repetition, triggers)
// ---------------------------------------------------------------------------

const DEPTH_NS = `${NS} xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"`;
function parseDepth(xml: string): Element[] {
  return Array.from(parseXml(`<root${DEPTH_NS}>${xml}</root>`).documentElement.children);
}
/** "de-fN" → frame "fN"; "de-fN-pK" → paragraph K of frame "fN". */
function resolveDepth(wireId: string): OdpTarget | undefined {
  const m = /^de-(f\d+)(?:-p(\d+))?$/.exec(wireId);
  if (!m) return undefined;
  return m[2] === undefined ? m[1] : { frameId: m[1], paragraphIndex: Number(m[2]) };
}
function depthRoundTrip(animations: SlideAnimation[]): SlideAnimation[] | undefined {
  return readOdpAnim(parseDepth(writeOdpAnim(animations, goldenTargetFor, depthFrameFor)), resolveDepth, depthFrameFor);
}
/** A main sequence of one click group holding `effects`, under the root. */
function impressTree(effects: string, extra = ""): string {
  return (
    `<anim:par presentation:node-type="timing-root"><anim:seq presentation:node-type="main-sequence">` +
    `<anim:par smil:begin="next"><anim:par smil:begin="0s">${effects}</anim:par></anim:par></anim:seq>${extra}</anim:par>`
  );
}

describe("SPEC-105 depth golden (odp)", () => {
  it("serialises DEPTH_ANIMATIONS over DEPTH_FRAMES to ANIM_DEPTH_GOLDEN", () => {
    expect(writeOdpAnim(DEPTH_ANIMATIONS, goldenTargetFor, depthFrameFor)).toBe(ANIM_DEPTH_GOLDEN);
  });

  it("reads ANIM_DEPTH_GOLDEN back to DEPTH_ANIMATIONS, interactive rows back in their model place", () => {
    expect(readOdpAnim(parseDepth(ANIM_DEPTH_GOLDEN), resolveDepth, depthFrameFor)).toEqual(DEPTH_ANIMATIONS);
  });

  it("keeps the SPEC-086 golden byte-identical", () => {
    expect(writeOdpAnim(GOLDEN_ANIMATIONS, goldenTargetFor, depthFrameFor)).toBe(ANIM_GOLDEN);
  });

  it("targets each step's paragraphs by their de-f{n}-p{k} text:p ids", () => {
    const xml = writeOdpAnim([{ frameId: "f4", effect: "fadeIn", trigger: "click", buildBy: "byFirstLevelParagraph" }], goldenTargetFor, depthFrameFor);
    expect(Array.from(xml.matchAll(/smil:targetElement="([^"]+)"/g), (m) => m[1])).toEqual([
      "de-f4-p0",
      "de-f4-p1",
      "de-f4-p2",
      "de-f4-p3",
    ]);
    expect(paragraphWireId("de-f4", 2)).toBe("de-f4-p2");
  });
});

describe("SPEC-105 per-feature round trips (odp)", () => {
  const forms: SlideAnimation[][] = [
    [{ frameId: "f2", effect: "pulse", trigger: "click" }],
    [{ frameId: "f2", effect: "spin", trigger: "click", spinDegrees: 90, durationMs: 1200 }],
    [{ frameId: "f6", effect: "growShrink", trigger: "click", scalePercent: 10 }],
    [{ frameId: "f6", effect: "colorChange", trigger: "click", emphasisColor: "#123abc", durationMs: 2000 }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "line", direction: "down", distanceFraction: 1 } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "arc", direction: "up", distanceFraction: 0.05 } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "turn", direction: "left", distanceFraction: 0.4 } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "loop" } }],
    [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "custom", d: "M 0 0 L 0.1 0.1 L 0.2 0" } }],
    [{ frameId: "f2", effect: "fadeIn", trigger: "click", repeatCount: 10, autoReverse: true, rewindOnComplete: true }],
    [{ frameId: "f2", effect: "spin", trigger: "click", repeatUntilClick: true }],
    [{ frameId: "f3", effect: "zoomIn", trigger: "click", buildBy: "byParagraph", buildTrigger: "afterPrevious", delayMs: 300 }],
    [{ frameId: "f4", effect: "fadeOut", trigger: "click", buildBy: "byFirstLevelParagraph", buildTrigger: "withPrevious" }],
    [
      { frameId: "f3", effect: "fadeIn", trigger: "onClickOf", triggerFrameId: "f5", buildBy: "byParagraph" },
      { frameId: "f2", effect: "fadeIn", trigger: "click" },
      { frameId: "f6", effect: "spin", trigger: "onClickOf", triggerFrameId: "f5" },
    ],
  ];
  for (const form of forms) {
    it(`${form.map((a) => `${a.effect}${a.buildBy ? `/${a.buildBy}` : ""}${a.trigger === "onClickOf" ? "/onClickOf" : ""}`).join(" + ")}`, () => {
      expect(depthRoundTrip(form)).toEqual(form);
    });
  }
});

describe("SPEC-105 tolerant reading (Impress-flavoured)", () => {
  it("reads both spellings: svg:type/smil:type on a rotate, svg:path/smil:path on a motion", () => {
    const spin = (typeAttr: string) =>
      impressTree(
        `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin">` +
          `<anim:animateTransform smil:dur="2s" smil:fill="hold" smil:targetElement="de-f2" ${typeAttr}="rotate" smil:by="180"/></anim:par>`,
      );
    const expected = [{ frameId: "f2", effect: "spin", trigger: "click", durationMs: 2000, spinDegrees: 180 }];
    expect(readOdpAnim(parseDepth(spin("svg:type")), resolveDepth, depthFrameFor)).toEqual(expected);
    expect(readOdpAnim(parseDepth(spin("smil:type")), resolveDepth, depthFrameFor)).toEqual(expected);
    const motion = (pathAttr: string) =>
      impressTree(
        `<anim:par smil:begin="0s" presentation:node-type="on-click" presentation:preset-class="motion-path" presentation:preset-id="ooo-motionpath-line">` +
          `<anim:animateMotion smil:dur="0.5s" smil:fill="hold" smil:targetElement="de-f6" ${pathAttr}="m 0 0 l 0 -0.25" presentation:additive="sum"/></anim:par>`,
      );
    const line = [{ frameId: "f6", effect: "motionPath", trigger: "click", path: { preset: "line", direction: "up" } }];
    expect(readOdpAnim(parseDepth(motion("svg:path")), resolveDepth, depthFrameFor)).toEqual(line);
    expect(readOdpAnim(parseDepth(motion("smil:path")), resolveDepth, depthFrameFor)).toEqual(line);
  });

  it("reads Impress's by-paragraph anim:iterate as a by-paragraph build, after previous, its interval the delay", () => {
    const xml = impressTree(
      `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-fade-in">` +
        `<anim:iterate smil:targetElement="de-f3" anim:sub-item="text" anim:iterate-type="by-paragraph" anim:iterate-interval="0.4s">` +
        `<anim:transitionFilter smil:dur="0.8s" smil:type="fade" smil:direction="forward"/></anim:iterate></anim:par>`,
    );
    expect(readOdpAnim(parseDepth(xml), resolveDepth, depthFrameFor)).toEqual([
      {
        frameId: "f3",
        effect: "fadeIn",
        trigger: "click",
        durationMs: 800,
        delayMs: 400,
        buildBy: "byParagraph",
        buildTrigger: "afterPrevious",
      },
    ]);
  });

  it("reads Impress's paragraph-targeted effects and interactive sequence (smil:end on the trigger)", () => {
    const xml = impressTree(
      `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin">` +
        `<anim:animateTransform smil:dur="2s" smil:fill="hold" smil:targetElement="de-f3-p0" svg:type="rotate" smil:by="360"/></anim:par>` +
        [1, 2, 3]
          .map(
            (k) =>
              `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="with-previous" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin">` +
              `<anim:animateTransform smil:dur="2s" smil:fill="hold" smil:targetElement="de-f3-p${k}" svg:type="rotate" smil:by="360"/></anim:par>`,
          )
          .join(""),
      `<anim:seq presentation:node-type="interactive-sequence" smil:begin="de-f5.click" smil:end="de-f5.click">` +
        `<anim:par smil:begin="0s"><anim:par smil:begin="0s">` +
        `<anim:par smil:begin="0s" smil:fill="hold" presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-appear">` +
        `<anim:set smil:begin="0s" smil:dur="0.001s" smil:fill="hold" smil:targetElement="de-f6" smil:attributeName="visibility" smil:to="visible"/>` +
        `</anim:par></anim:par></anim:par></anim:seq>`,
    );
    expect(readOdpAnim(parseDepth(xml), resolveDepth, depthFrameFor)).toEqual([
      { frameId: "f3", effect: "spin", trigger: "click", durationMs: 2000, buildBy: "byParagraph", buildTrigger: "withPrevious" },
      { frameId: "f6", effect: "appear", trigger: "onClickOf", triggerFrameId: "f5" },
    ]);
  });
});

/** A fade-in effect par on `target` with node type `node`. */
function fadeParOn(target: string, node = "with-previous"): string {
  return (
    `<anim:par presentation:node-type="${node}" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-fade-in" smil:begin="0s">` +
    `<anim:transitionFilter smil:type="fade" smil:direction="forward" smil:dur="0.5s" smil:targetElement="${target}"/></anim:par>`
  );
}

describe("SPEC-105 out of subset (odp → preserved)", () => {
  const read = (xml: string): SlideAnimation[] | undefined => readOdpAnim(parseDepth(xml), resolveDepth, depthFrameFor);

  it("a partial paragraph cover, or paragraphs out of order", () => {
    expect(read(impressTree(fadeParOn("de-f3-p0", "on-click") + fadeParOn("de-f3-p1") + fadeParOn("de-f3-p2")))).toBeUndefined();
    expect(read(impressTree(fadeParOn("de-f3-p1", "on-click") + fadeParOn("de-f3-p0") + fadeParOn("de-f3-p2") + fadeParOn("de-f3-p3")))).toBeUndefined();
    // The control: every paragraph, in order.
    expect(read(impressTree(fadeParOn("de-f3-p0", "on-click") + fadeParOn("de-f3-p1") + fadeParOn("de-f3-p2") + fadeParOn("de-f3-p3")))).toHaveLength(1);
  });

  it("an unresolvable trigger, a paragraph trigger, or a trigger that is the target", () => {
    const interactive = (trigger: string) =>
      impressTree(
        fadeParOn("de-f2", "on-click"),
        `<anim:seq presentation:node-type="interactive-sequence" smil:begin="${trigger}.click"><anim:par smil:begin="0s"><anim:par smil:begin="0s">${fadeParOn("de-f3", "on-click")}</anim:par></anim:par></anim:seq>`,
      );
    expect(read(interactive("de-f5"))).toHaveLength(2); // the control
    expect(read(interactive("id99"))).toBeUndefined();
    expect(read(interactive("de-f4-p0"))).toBeUndefined();
    expect(read(interactive("de-f3"))).toBeUndefined();
  });

  it("a second sequence that is not an interactive one", () => {
    const extra = `<anim:seq presentation:node-type="main-sequence"><anim:par smil:begin="next"><anim:par smil:begin="0s">${fadeParOn("de-f3", "on-click")}</anim:par></anim:par></anim:seq>`;
    expect(read(impressTree(fadeParOn("de-f2", "on-click"), extra))).toBeUndefined();
    expect(read(impressTree(fadeParOn("de-f2", "on-click"), extra.replace(` presentation:node-type="main-sequence"`, "")))).toBeUndefined();
  });

  it("a malformed path, or one its preset does not reproduce", () => {
    const path = (preset: string, d: string) =>
      impressTree(
        `<anim:par presentation:node-type="on-click" presentation:preset-class="motion-path" presentation:preset-id="${preset}" smil:begin="0s">` +
          `<anim:animateMotion svg:path="${d}" smil:dur="0.5s" smil:targetElement="de-f6"/></anim:par>`,
      );
    expect(read(path("ooo-motionpath-line", "M 0 0 L 0.25 0"))).toHaveLength(1); // the control
    expect(read(path("ooo-motionpath-custom", "M 0 0 Q 0.1 0.1 0.2 0"))).toBeUndefined();
    expect(read(path("ooo-motionpath-line", "M 0 0 L 0.2 0.2"))).toBeUndefined();
    expect(read(path("ooo-motionpath-loop", "M 0 0 L 0.2 0"))).toBeUndefined();
    expect(read(path("ooo-motionpath-4-point-star", "M 0 0 L 0.25 0"))).toBeUndefined();
  });

  it("an unmapped emphasis (flash bulb), or a mapped one over other behaviours", () => {
    expect(read(impressTree(fadeParOn("de-f2", "on-click").replace("ooo-entrance-fade-in", "ooo-emphasis-flash-bulb")))).toBeUndefined();
    const spin = writeOdpAnim([{ frameId: "f2", effect: "spin", trigger: "click" }], goldenTargetFor, depthFrameFor);
    expect(read(spin.replace('smil:type="rotate"', 'smil:type="scale"'))).toBeUndefined();
    const color = writeOdpAnim([{ frameId: "f6", effect: "colorChange", trigger: "click" }], goldenTargetFor, depthFrameFor);
    expect(read(color.replaceAll(' smil:autoReverse="true"', ""))).toBeUndefined(); // Impress's held fill colour
  });

  it("a by-word iterate stays out; a by-paragraph one on a paragraph target does too", () => {
    const iterate = (type: string, target: string) =>
      impressTree(
        `<anim:par presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-fade-in" smil:begin="0s">` +
          `<anim:iterate smil:targetElement="${target}" anim:iterate-type="${type}" anim:iterate-interval="0.05s">` +
          `<anim:transitionFilter smil:dur="0.5s" smil:type="fade" smil:direction="forward"/></anim:iterate></anim:par>`,
      );
    expect(read(iterate("by-word", "de-f3"))).toBeUndefined();
    expect(read(iterate("by-paragraph", "de-f3-p0"))).toBeUndefined();
  });
});
