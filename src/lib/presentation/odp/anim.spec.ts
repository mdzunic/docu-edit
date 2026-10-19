import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import type { SlideAnimation } from "../model";
import { preservedAnimReferences, readOdpAnim, writeOdpAnim, xmlIdsDefinedIn } from "./anim";
import { ANIM_GOLDEN, GOLDEN_ANIMATIONS, goldenTargetFor } from "./animGolden";

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
