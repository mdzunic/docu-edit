import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import type { SlideAnimation } from "../model";
import { readOdpAnim, writeOdpAnim } from "./anim";
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
