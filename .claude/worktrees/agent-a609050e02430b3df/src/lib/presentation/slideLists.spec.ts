import { describe, expect, it } from "vitest";
import type { SlideParagraph } from "./model";
import { effectiveHangingCm, effectiveIndentCm, paragraphMarkers } from "./slideLists";

const p = (extra: Partial<SlideParagraph> = {}): SlideParagraph => ({ runs: [{ text: "x" }], ...extra });
const texts = (paras: SlideParagraph[]): string[] => paragraphMarkers(paras).map((m) => m.text);

describe("paragraphMarkers (derived numbering, SPEC-097)", () => {
  it("numbers adjacent items and seeds the run from startAt", () => {
    expect(texts([p({ bullet: "number", startAt: 3 }), p({ bullet: "number", startAt: 3 }), p({ bullet: "number" })])).toEqual([
      "3.",
      "4.",
      "5.",
    ]);
  });

  it("a non-list paragraph restarts the numbering", () => {
    expect(texts([p({ bullet: "number" }), p(), p({ bullet: "number" })])).toEqual(["1.", "", "1."]);
  });

  it("a shallower item restarts the deeper level; a deeper run does not break the outer one", () => {
    const paras = [
      p({ bullet: "number" }),
      p({ bullet: "number", level: 1, bulletStyle: "lowerAlpha" }),
      p({ bullet: "number", level: 1, bulletStyle: "lowerAlpha" }),
      p({ bullet: "number" }),
      p({ bullet: "number", level: 1, bulletStyle: "lowerAlpha" }),
    ];
    expect(texts(paras)).toEqual(["1.", "a.", "b.", "2.", "a."]);
  });

  it("renders bullet glyphs, custom chars and picture bullets", () => {
    const markers = paragraphMarkers([
      p({ bullet: "disc" }),
      p({ bullet: "disc", bulletStyle: "square" }),
      p({ bullet: "disc", bulletChar: "★" }),
      p({ bullet: "disc", bulletImageId: "ppt/media/image2.png" }),
    ]);
    expect(markers).toEqual([{ text: "•" }, { text: "▪" }, { text: "★" }, { text: "", imageId: "ppt/media/image2.png" }]);
  });

  it("a bullet or a different number style at the same level restarts that level", () => {
    expect(
      texts([p({ bullet: "number" }), p({ bullet: "disc" }), p({ bullet: "number" }), p({ bullet: "number", bulletStyle: "upperRoman" })]),
    ).toEqual(["1.", "•", "1.", "I."]);
  });
});

describe("effective indents", () => {
  it("a list item's text edge is (level + 1) × 1.27 with a 0.635 hanging marker box", () => {
    expect(effectiveIndentCm(p({ bullet: "disc" }))).toBeCloseTo(1.27, 6);
    expect(effectiveIndentCm(p({ bullet: "disc", level: 2 }))).toBeCloseTo(3.81, 6);
    expect(effectiveHangingCm(p({ bullet: "disc" }))).toBe(0.635);
  });

  it("plain text stays flush (level × 1.27, no hanging) and explicit values win", () => {
    expect(effectiveIndentCm(p())).toBe(0);
    expect(effectiveIndentCm(p({ level: 1 }))).toBeCloseTo(1.27, 6);
    expect(effectiveHangingCm(p())).toBe(0);
    expect(effectiveIndentCm(p({ bullet: "disc", indentCm: 2 }))).toBe(2);
    expect(effectiveHangingCm(p({ bullet: "disc", hangingCm: 1 }))).toBe(1);
  });
});
