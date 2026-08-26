import { describe, expect, it } from "vitest";
import { parseXml } from "../core/xml";
import type { SlideTransition, TransitionDirection, TransitionType } from "./model";
import {
  DEFAULT_TRANSITION_DURATION_MS,
  defaultDirectionFor,
  effectiveDirection,
  effectiveDurationMs,
  msToOdfDuration,
  odpTransitionProps,
  parseOdfDurationToMs,
  pptxSpeedFor,
  readOdpTransition,
  readPptxTransition,
  usesDirection,
  writePptxTransition,
  type OdpTransitionAttrs,
} from "./transitions";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** Builds a `p:transition` element with the given attributes and inner XML. */
function transitionEl(attrs: string, inner = ""): Element {
  const doc = parseXml(
    `<p:transition xmlns:p="${P_NS}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ${attrs}>${inner}</p:transition>`,
  );
  return doc.documentElement;
}

describe("transition helpers", () => {
  it("knows which effects use direction", () => {
    expect(usesDirection("push")).toBe(true);
    expect(usesDirection("wipe")).toBe(true);
    expect(usesDirection("cover")).toBe(true);
    expect(usesDirection("uncover")).toBe(true);
    expect(usesDirection("fade")).toBe(false);
    expect(usesDirection("split")).toBe(false);
    expect(usesDirection("none")).toBe(false);
  });

  it("defaults every directional effect to left", () => {
    for (const type of ["push", "wipe", "cover", "uncover"] as TransitionType[]) {
      expect(defaultDirectionFor(type)).toBe("left");
    }
    expect(defaultDirectionFor("fade")).toBeUndefined();
    expect(defaultDirectionFor("split")).toBeUndefined();
  });

  it("resolves effective direction and duration", () => {
    expect(effectiveDirection({ type: "push" })).toBe("left");
    expect(effectiveDirection({ type: "push", direction: "up" })).toBe("up");
    expect(effectiveDurationMs({ type: "fade" })).toBe(DEFAULT_TRANSITION_DURATION_MS);
    expect(effectiveDurationMs({ type: "fade", durationMs: 2000 })).toBe(2000);
    // clamps to bounds
    expect(effectiveDurationMs({ type: "fade", durationMs: 50 })).toBe(100);
    expect(effectiveDurationMs({ type: "fade", durationMs: 9000 })).toBe(5000);
  });
});

describe("pptx quantisation", () => {
  it("maps durations to the nearest spd bucket", () => {
    expect(pptxSpeedFor(300)).toBe("fast");
    expect(pptxSpeedFor(450)).toBe("fast");
    expect(pptxSpeedFor(600)).toBe("med");
    expect(pptxSpeedFor(800)).toBe("med");
    expect(pptxSpeedFor(1000)).toBe("slow");
    expect(pptxSpeedFor(2000)).toBe("slow");
  });
});

describe("pptx read", () => {
  it("parses fade with no direction", () => {
    expect(readPptxTransition(transitionEl(`spd="med"`, `<p:fade/>`))).toEqual({ type: "fade", durationMs: 600 });
  });

  it("parses each directional effect in every direction", () => {
    const cases: [string, TransitionType, string, TransitionDirection][] = [
      ["push", "push", "l", "left"],
      ["push", "push", "r", "right"],
      ["push", "push", "u", "up"],
      ["push", "push", "d", "down"],
      ["wipe", "wipe", "l", "left"],
      ["cover", "cover", "d", "down"],
      ["pull", "uncover", "r", "right"],
    ];
    for (const [element, type, dir, direction] of cases) {
      const t = readPptxTransition(transitionEl(`spd="fast"`, `<p:${element} dir="${dir}"/>`));
      expect(t).toEqual({ type, direction, durationMs: 300 });
    }
  });

  it("parses split with any orient/dir", () => {
    expect(readPptxTransition(transitionEl(`spd="slow"`, `<p:split orient="vert" dir="in"/>`))).toEqual({
      type: "split",
      durationMs: 1000,
    });
  });

  it("maps p:cut and an empty transition to explicit none", () => {
    expect(readPptxTransition(transitionEl(`spd="med"`, `<p:cut/>`))?.type).toBe("none");
    expect(readPptxTransition(transitionEl(`advClick="1"`))?.type).toBe("none");
  });

  it("prefers a p14:dur extension over spd", () => {
    const inner = `<p:fade/><p:extLst><p:ext uri="x"><p14:dur val="2000"/></p:ext></p:extLst>`;
    expect(readPptxTransition(transitionEl(`spd="fast"`, inner))).toEqual({ type: "fade", durationMs: 2000 });
  });

  it("reads advClick and advTm", () => {
    const t = readPptxTransition(transitionEl(`spd="med" advClick="0" advTm="5000"`, `<p:fade/>`));
    expect(t).toEqual({ type: "fade", durationMs: 600, advanceOnClick: false, advanceAfterMs: 5000 });
  });

  it("returns undefined for an out-of-subset effect", () => {
    expect(readPptxTransition(transitionEl(`spd="med"`, `<p:morph/>`))).toBeUndefined();
    expect(readPptxTransition(transitionEl(`spd="med"`, `<p:zoom/>`))).toBeUndefined();
  });

  it("ignores a sound cue when locating the effect", () => {
    const inner = `<p:sndAc><p:stSnd><p:snd r:embed="rId9"/></p:stSnd></p:sndAc><p:fade/>`;
    // r namespace unused by the reader; declare to keep the fragment well-formed.
    const doc = parseXml(
      `<p:transition xmlns:p="${P_NS}" xmlns:r="http://x" spd="med">${inner}</p:transition>`,
    );
    expect(readPptxTransition(doc.documentElement)?.type).toBe("fade");
  });
});

describe("pptx write", () => {
  it("writes fade, split and directional effects", () => {
    expect(writePptxTransition({ type: "fade", durationMs: 600 })).toBe(`<p:transition spd="med"><p:fade/></p:transition>`);
    expect(writePptxTransition({ type: "split", durationMs: 300 })).toBe(
      `<p:transition spd="fast"><p:split orient="horz" dir="out"/></p:transition>`,
    );
    expect(writePptxTransition({ type: "push", direction: "up", durationMs: 1000 })).toBe(
      `<p:transition spd="slow"><p:push dir="u"/></p:transition>`,
    );
    expect(writePptxTransition({ type: "uncover", direction: "right", durationMs: 600 })).toBe(
      `<p:transition spd="med"><p:pull dir="r"/></p:transition>`,
    );
  });

  it("defaults a directional effect to left", () => {
    expect(writePptxTransition({ type: "cover" })).toBe(`<p:transition spd="med"><p:cover dir="l"/></p:transition>`);
  });

  it("writes advClick and advTm only when set", () => {
    expect(writePptxTransition({ type: "fade", advanceOnClick: false, advanceAfterMs: 5000 })).toBe(
      `<p:transition spd="med" advClick="0" advTm="5000"><p:fade/></p:transition>`,
    );
  });

  it("writes explicit none with no effect child", () => {
    expect(writePptxTransition({ type: "none" })).toBe(`<p:transition spd="med"></p:transition>`);
  });

  it("round-trips the six with quantisation as the only delta", () => {
    const originals: SlideTransition[] = [
      { type: "fade", durationMs: 2000 },
      { type: "push", direction: "right", durationMs: 2000 },
      { type: "wipe", direction: "up", durationMs: 2000 },
      { type: "cover", direction: "down", durationMs: 2000 },
      { type: "uncover", direction: "left", durationMs: 2000 },
      { type: "split", durationMs: 2000 },
    ];
    for (const original of originals) {
      // Re-parse the written element back to the model; duration quantises
      // 2000 -> slow (1000) — the single sanctioned round-trip delta.
      const written = writePptxTransition(original);
      const parsed = readPptxTransition(parseXml(written.replace("<p:transition ", `<p:transition xmlns:p="${P_NS}" `)).documentElement);
      expect(parsed?.type).toBe(original.type);
      expect(parsed?.direction).toBe(original.direction);
      expect(parsed?.durationMs).toBe(1000);
    }
  });
});

describe("odf duration", () => {
  it("parses and formats PTxS", () => {
    expect(parseOdfDurationToMs("PT0.6S")).toBe(600);
    expect(parseOdfDurationToMs("PT5S")).toBe(5000);
    expect(parseOdfDurationToMs("PT1.5S")).toBe(1500);
    expect(parseOdfDurationToMs("bogus")).toBeUndefined();
    expect(msToOdfDuration(600)).toBe("PT0.6S");
    expect(msToOdfDuration(5000)).toBe("PT5S");
    expect(msToOdfDuration(1500)).toBe("PT1.5S");
  });
});

describe("odp read", () => {
  const read = (a: OdpTransitionAttrs) => readOdpTransition(a);

  it("parses fade", () => {
    expect(read({ smilType: "fade", smilSubtype: "crossfade", smilDur: "PT0.6S" })).toEqual({
      type: "fade",
      durationMs: 600,
    });
  });

  it("parses push from the opposite edge for every direction", () => {
    expect(read({ smilType: "pushWipe", smilSubtype: "fromRight" })?.direction).toBe("left");
    expect(read({ smilType: "pushWipe", smilSubtype: "fromLeft" })?.direction).toBe("right");
    expect(read({ smilType: "pushWipe", smilSubtype: "fromBottom" })?.direction).toBe("up");
    expect(read({ smilType: "pushWipe", smilSubtype: "fromTop" })?.direction).toBe("down");
    expect(read({ smilType: "pushWipe", smilSubtype: "fromRight" })?.type).toBe("push");
  });

  it("parses wipe with the reverse flag selecting left/up", () => {
    expect(read({ smilType: "barWipe", smilSubtype: "leftToRight" })).toMatchObject({ type: "wipe", direction: "right" });
    expect(read({ smilType: "barWipe", smilSubtype: "leftToRight", smilDirection: "reverse" })).toMatchObject({
      type: "wipe",
      direction: "left",
    });
    expect(read({ smilType: "barWipe", smilSubtype: "topToBottom" })).toMatchObject({ type: "wipe", direction: "down" });
    expect(read({ smilType: "barWipe", smilSubtype: "topToBottom", smilDirection: "reverse" })).toMatchObject({
      type: "wipe",
      direction: "up",
    });
  });

  it("distinguishes cover and uncover by the reverse flag", () => {
    expect(read({ smilType: "slideWipe", smilSubtype: "fromRight" })).toMatchObject({ type: "cover", direction: "left" });
    expect(read({ smilType: "slideWipe", smilSubtype: "fromRight", smilDirection: "reverse" })).toMatchObject({
      type: "uncover",
      direction: "left",
    });
  });

  it("parses split", () => {
    expect(read({ smilType: "barnDoorWipe", smilSubtype: "horizontal" })?.type).toBe("split");
  });

  it("falls back to transition-speed for duration", () => {
    expect(read({ smilType: "fade", transitionSpeed: "slow" })?.durationMs).toBe(1000);
    expect(read({ smilType: "fade", transitionSpeed: "fast" })?.durationMs).toBe(300);
  });

  it("reads an automatic advance", () => {
    expect(
      read({ smilType: "fade", transitionType: "automatic", presentationDuration: "PT5S" })?.advanceAfterMs,
    ).toBe(5000);
    expect(read({ smilType: "fade", transitionType: "manual" })?.advanceAfterMs).toBeUndefined();
  });

  it("returns undefined for an unrecognised combination", () => {
    expect(read({ smilType: "spiralWipe", smilSubtype: "clockwiseTop" })).toBeUndefined();
    expect(read({ smilType: "checkerBoardWipe", smilSubtype: "down" })).toBeUndefined();
  });
});

describe("odp write", () => {
  it("writes the smil attributes for each effect", () => {
    expect(odpTransitionProps({ type: "fade", durationMs: 600 })).toBe(
      `smil:type="fade" smil:subtype="crossfade" smil:dur="PT0.6S" presentation:transition-type="manual"`,
    );
    expect(odpTransitionProps({ type: "push", direction: "left", durationMs: 600 })).toContain(
      `smil:type="pushWipe" smil:subtype="fromRight"`,
    );
    expect(odpTransitionProps({ type: "wipe", direction: "left", durationMs: 600 })).toContain(
      `smil:type="barWipe" smil:subtype="leftToRight" smil:direction="reverse"`,
    );
    expect(odpTransitionProps({ type: "uncover", direction: "left", durationMs: 600 })).toContain(
      `smil:type="slideWipe" smil:subtype="fromRight" smil:direction="reverse"`,
    );
    expect(odpTransitionProps({ type: "split", durationMs: 600 })).toContain(
      `smil:type="barnDoorWipe" smil:subtype="horizontal"`,
    );
  });

  it("writes an automatic advance", () => {
    expect(odpTransitionProps({ type: "fade", durationMs: 600, advanceAfterMs: 5000 })).toContain(
      `presentation:transition-type="automatic" presentation:duration="PT5S"`,
    );
  });

  it("writes only timing for explicit none", () => {
    expect(odpTransitionProps({ type: "none" })).toBe(`presentation:transition-type="manual"`);
  });

  it("round-trips the six through odp", () => {
    const originals: SlideTransition[] = [
      { type: "fade", durationMs: 800 },
      { type: "push", direction: "right", durationMs: 800 },
      { type: "wipe", direction: "up", durationMs: 800 },
      { type: "cover", direction: "down", durationMs: 800 },
      { type: "uncover", direction: "left", durationMs: 800 },
      { type: "split", durationMs: 800 },
    ];
    for (const original of originals) {
      const props = odpTransitionProps(original);
      const attrs: OdpTransitionAttrs = {
        smilType: /smil:type="([^"]*)"/.exec(props)?.[1],
        smilSubtype: /smil:subtype="([^"]*)"/.exec(props)?.[1],
        smilDirection: /smil:direction="([^"]*)"/.exec(props)?.[1],
        smilDur: /smil:dur="([^"]*)"/.exec(props)?.[1],
      };
      const parsed = readOdpTransition(attrs);
      expect(parsed?.type).toBe(original.type);
      expect(parsed?.direction).toBe(original.direction);
      expect(parsed?.durationMs).toBe(800);
    }
  });
});
