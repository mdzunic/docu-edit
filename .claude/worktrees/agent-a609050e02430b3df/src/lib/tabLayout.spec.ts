import { describe, expect, it } from "vitest";
import { layoutParagraphTabs, stopsToPx } from "./tabLayout";
import { PX_PER_CM } from "./core/units";
import type { TabStop } from "./documentModel";

// A fixed-width fake measurer: every character is 10 px wide, so widths are
// easy to reason about independent of any real font.
const CHAR = 10;
const measure = (from: number, to: number) => (to - from) * CHAR;

function px(cm: number): number {
  return cm * PX_PER_CM;
}

describe("SPEC-054 — layoutParagraphTabs", () => {
  it("advances a left tab so the following text starts at the stop (AC-1, AC-17)", () => {
    const stops = stopsToPx([{ positionCm: 3, alignment: "left" }] as TabStop[]);
    // "a\tb": pen after "a" is 10 px; the tab fills to the 3 cm stop.
    const { advances } = layoutParagraphTabs("a\tb", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(3) - CHAR, 4);
  });

  it("resumes the default grid past the last explicit stop, anchored at the start (AC-15, AC-18)", () => {
    // No explicit stops, default 1.27 cm (≈48 px). Pen after "a" = 10 px → the
    // first grid multiple strictly past it is 1×48 px.
    const { advances } = layoutParagraphTabs("a\tb", [], px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(1.27) - CHAR, 4);
  });

  it("ends the following text on a right stop (AC-17)", () => {
    const stops = stopsToPx([{ positionCm: 4, alignment: "right" }] as TabStop[]);
    // "a\tbc": following text "bc" is 20 px, ending on the 4 cm stop.
    const { advances } = layoutParagraphTabs("a\tbc", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(4) - CHAR - 20, 4);
  });

  it("centres the following text on a centre stop", () => {
    const stops = stopsToPx([{ positionCm: 4, alignment: "center" }] as TabStop[]);
    const { advances } = layoutParagraphTabs("a\tbc", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(4) - CHAR - 10, 4);
  });

  it("aligns a decimal stop on the document-language separator (AC-19)", () => {
    const stops = stopsToPx([{ positionCm: 5, alignment: "decimal" }] as TabStop[]);
    // "a\t12.34": prefix before "." is "12" = 20 px.
    const { advances } = layoutParagraphTabs("a\t12.34", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(5) - CHAR - 20, 4);
  });

  it("falls back to the first separator character when the language separator is absent (AC-20)", () => {
    const stops = stopsToPx([{ positionCm: 5, alignment: "decimal" }] as TabStop[]);
    // German separator "," but the text uses ".": align on the "." anyway.
    const { advances } = layoutParagraphTabs("a\t12.34", stops, px(1.27), ",", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(5) - CHAR - 20, 4);
  });

  it("falls back to right alignment when the text has no separator (AC-21)", () => {
    const stops = stopsToPx([{ positionCm: 5, alignment: "decimal" }] as TabStop[]);
    // "a\t1234": no separator → the whole "1234" (40 px) ends on the stop.
    const { advances } = layoutParagraphTabs("a\t1234", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(5) - CHAR - 40, 4);
  });

  it("draws a bar rule at each bar stop and never targets one with a tab (AC-22, AC-23)", () => {
    const stops = stopsToPx([
      { positionCm: 2, alignment: "bar" },
      { positionCm: 4, alignment: "left" },
    ] as TabStop[]);
    const { advances, bars } = layoutParagraphTabs("a\tb", stops, px(1.27), ".", measure);
    expect(bars).toHaveLength(1);
    expect(bars[0]).toBeCloseTo(px(2), 4);
    // The tab skips the 2 cm bar and lands on the 4 cm left stop, not the bar.
    expect(advances.get(1)?.widthPx).toBeCloseTo(px(4) - CHAR, 4);
  });

  it("carries the target stop's leader onto the tab (AC-24)", () => {
    const stops = stopsToPx([{ positionCm: 4, alignment: "right", leader: "dot" }] as TabStop[]);
    const { advances } = layoutParagraphTabs("a\tb", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.leader).toBe("dot");
  });

  it("resets the pen on a hard line break", () => {
    const stops = stopsToPx([{ positionCm: 3, alignment: "left" }] as TabStop[]);
    // Second line "\tb" starts at pen 0, so its tab fills the whole 3 cm.
    const { advances } = layoutParagraphTabs("aaa\n\tb", stops, px(1.27), ".", measure);
    expect(advances.get(4)?.widthPx).toBeCloseTo(px(3), 4);
  });

  it("clamps a tab that would overshoot to a zero-width box", () => {
    const stops = stopsToPx([{ positionCm: 1, alignment: "right" }] as TabStop[]);
    // Following text (40 px) is wider than the 1 cm stop allows → width 0.
    const { advances } = layoutParagraphTabs("a\t1234", stops, px(1.27), ".", measure);
    expect(advances.get(1)?.widthPx).toBe(0);
  });
});
