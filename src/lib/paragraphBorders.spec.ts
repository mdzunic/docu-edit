import { describe, expect, it } from "vitest";
import {
  ALL_EDGES,
  applyBorderPreset,
  borderRuns,
  DEFAULT_MARK_COLOR,
  detectBorderPreset,
  fragmentEdges,
  hasBorderPaint,
  isDarkFill,
  normaliseParagraphBorders,
  paragraphBorderStyle,
  paragraphBordersEqual,
  paragraphMarkColorStyle,
  paragraphPaddingPt,
} from "./paragraphBorders";
import type { BorderSide } from "./core/borders";
import type { Block, Paragraph, ParagraphBorders, ParagraphFormatting } from "./documentModel";

const LINE: BorderSide = { style: "single", widthPt: 1, color: "#000000" };
const THICK: BorderSide = { style: "thick", widthPt: 3, color: "#ff0000" };

function para(formatting?: ParagraphFormatting, text = "x"): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], formatting };
}

function boxed(extra: Partial<ParagraphFormatting> = {}): ParagraphFormatting {
  return { borders: { top: LINE, bottom: LINE, left: LINE, right: LINE }, ...extra };
}

describe("paragraphPaddingPt (AC-9)", () => {
  it("defaults every edge to 4pt", () => {
    expect(paragraphPaddingPt(undefined, "top")).toBe(4);
    expect(paragraphPaddingPt({ top: LINE }, "left")).toBe(4);
  });

  it("uses the paragraph's own value where it sets one", () => {
    const borders: ParagraphBorders = { top: LINE, paddingPt: { top: 12 } };
    expect(paragraphPaddingPt(borders, "top")).toBe(12);
    expect(paragraphPaddingPt(borders, "bottom")).toBe(4);
  });

  it("honours an explicit zero rather than falling back to the default", () => {
    expect(paragraphPaddingPt({ top: LINE, paddingPt: { top: 0 } }, "top")).toBe(0);
  });
});

describe("hasBorderPaint", () => {
  it("is false for no formatting, empty formatting and all-none borders", () => {
    expect(hasBorderPaint(undefined)).toBe(false);
    expect(hasBorderPaint({})).toBe(false);
    expect(hasBorderPaint({ borders: { top: { style: "none", widthPt: 0, color: "auto" } } })).toBe(false);
  });

  it("is true for a visible edge or a shading fill", () => {
    expect(hasBorderPaint({ borders: { bottom: LINE } })).toBe(true);
    expect(hasBorderPaint({ shadingFill: "#eeeeee" })).toBe(true);
  });
});

describe("paragraphBordersEqual (AC-12's 'identical')", () => {
  it("normalises the padding default, so absent equals an explicit 4pt on all edges", () => {
    const absent: ParagraphBorders = { top: LINE };
    const explicit: ParagraphBorders = { top: LINE, paddingPt: { top: 4, bottom: 4, left: 4, right: 4 } };
    expect(paragraphBordersEqual(absent, explicit)).toBe(true);
  });

  it("separates paragraphs whose padding genuinely differs", () => {
    expect(paragraphBordersEqual({ top: LINE }, { top: LINE, paddingPt: { top: 10 } })).toBe(false);
  });

  it("compares every edge including between", () => {
    expect(paragraphBordersEqual({ top: LINE, between: LINE }, { top: LINE })).toBe(false);
    expect(paragraphBordersEqual({ top: LINE }, { top: THICK })).toBe(false);
  });
});

describe("borderRuns (AC-12)", () => {
  it("gives a lone bordered paragraph every edge it sets", () => {
    const blocks: Block[] = [para(boxed())];
    expect(borderRuns(blocks)).toEqual([{ top: true, bottom: true, sides: true, between: false, continuesBelow: false }]);
  });

  it("merges three consecutive identical paragraphs into one block", () => {
    const formatting = boxed();
    const blocks: Block[] = [para({ ...formatting }), para({ ...formatting }), para({ ...formatting })];
    expect(borderRuns(blocks)).toEqual([
      { top: true, bottom: false, sides: true, between: false, continuesBelow: true },
      { top: false, bottom: false, sides: true, between: false, continuesBelow: true },
      { top: false, bottom: true, sides: true, between: false, continuesBelow: false },
    ]);
  });

  it("draws the between edge at each internal boundary when it is set", () => {
    const formatting: ParagraphFormatting = {
      borders: { top: LINE, bottom: LINE, left: LINE, right: LINE, between: THICK },
    };
    const blocks: Block[] = [para({ ...formatting }), para({ ...formatting }), para({ ...formatting })];
    const runs = borderRuns(blocks);
    expect(runs.map((r) => r.between)).toEqual([true, true, false]);
    expect(runs.map((r) => r.top)).toEqual([true, false, false]);
    expect(runs.map((r) => r.bottom)).toEqual([false, false, true]);
  });

  it("draws nothing at an internal boundary where between is absent", () => {
    const formatting = boxed();
    const blocks: Block[] = [para({ ...formatting }), para({ ...formatting })];
    expect(borderRuns(blocks).map((r) => r.between)).toEqual([false, false]);
  });

  it("breaks the run at a paragraph whose borders differ", () => {
    const blocks: Block[] = [para(boxed()), para({ borders: { top: THICK, bottom: THICK } }), para(boxed())];
    const runs = borderRuns(blocks);
    expect(runs[0]).toEqual({ top: true, bottom: true, sides: true, between: false, continuesBelow: false });
    expect(runs[1]).toEqual({ top: true, bottom: true, sides: true, between: false, continuesBelow: false });
    expect(runs[2]).toEqual({ top: true, bottom: true, sides: true, between: false, continuesBelow: false });
  });

  it("breaks the run at a different shading fill", () => {
    const blocks: Block[] = [
      para(boxed({ shadingFill: "#eeeeee" })),
      para(boxed({ shadingFill: "#ffffff" })),
    ];
    expect(borderRuns(blocks).map((r) => r.bottom)).toEqual([true, true]);
  });

  it("breaks the run at a different indent", () => {
    const blocks: Block[] = [para(boxed()), para(boxed({ indentStartCm: 2 }))];
    expect(borderRuns(blocks).map((r) => r.top)).toEqual([true, true]);
  });

  it("breaks the run at a non-paragraph block", () => {
    const table = { type: "table", rows: [], columnWidthsCm: [] } as unknown as Block;
    const blocks: Block[] = [para(boxed()), table, para(boxed())];
    const runs = borderRuns(blocks);
    expect(runs[0]).toEqual({ top: true, bottom: true, sides: true, between: false, continuesBelow: false });
    expect(runs[1]).toEqual({ top: false, bottom: false, sides: false, between: false, continuesBelow: false });
    expect(runs[2]).toEqual({ top: true, bottom: true, sides: true, between: false, continuesBelow: false });
  });

  it("gives an unbordered paragraph no edges at all", () => {
    expect(borderRuns([para({}), para(undefined)])).toEqual([
      { top: false, bottom: false, sides: false, between: false, continuesBelow: false },
      { top: false, bottom: false, sides: false, between: false, continuesBelow: false },
    ]);
  });

  it("merges paragraphs shaded alike with no borders at all", () => {
    const shaded: ParagraphFormatting = { shadingFill: "#f5f5f5" };
    const runs = borderRuns([para({ ...shaded }), para({ ...shaded })]);
    expect(runs.map((r) => r.sides)).toEqual([true, true]);
  });
});

describe("fragmentEdges (AC-14..16)", () => {
  const runs = borderRuns([para(boxed())]);

  it("draws the side borders on both fragments of a split paragraph", () => {
    expect(fragmentEdges(runs, 0, true, false).sides).toBe(true);
    expect(fragmentEdges(runs, 0, false, true).sides).toBe(true);
  });

  it("draws the top border only on the first fragment", () => {
    expect(fragmentEdges(runs, 0, true, false).top).toBe(true);
    expect(fragmentEdges(runs, 0, false, true).top).toBe(false);
    expect(fragmentEdges(runs, 0, false, false).top).toBe(false);
  });

  it("draws the bottom border only on the last fragment", () => {
    expect(fragmentEdges(runs, 0, false, true).bottom).toBe(true);
    expect(fragmentEdges(runs, 0, true, false).bottom).toBe(false);
  });

  it("marks a paragraph sitting under a run's between rule", () => {
    const withBetween = borderRuns([
      para({ borders: { top: LINE, bottom: LINE, left: LINE, right: LINE, between: LINE } }),
      para({ borders: { top: LINE, bottom: LINE, left: LINE, right: LINE, between: LINE } }),
    ]);
    expect(fragmentEdges(withBetween, 0, true, true).belowBetween).toBe(false);
    expect(fragmentEdges(withBetween, 1, true, true).belowBetween).toBe(true);
    // A continuation fragment of that second paragraph is below a page
    // break, not below the rule.
    expect(fragmentEdges(withBetween, 1, false, true).belowBetween).toBe(false);
  });
});

describe("paragraphBorderStyle (AC-13..16)", () => {
  it("returns nothing for a paragraph that paints nothing", () => {
    expect(paragraphBorderStyle(undefined)).toEqual({});
    expect(paragraphBorderStyle({ alignment: "center" })).toEqual({});
  });

  it("paints every set edge with its padding for a standalone paragraph", () => {
    const style = paragraphBorderStyle(boxed(), ALL_EDGES);
    expect(style.borderTop).toBe("1pt solid #000000");
    expect(style.borderBottom).toBe("1pt solid #000000");
    expect(style.borderLeft).toBe("1pt solid #000000");
    expect(style.borderRight).toBe("1pt solid #000000");
    expect(style.paddingTop).toBe("4pt");
    expect(style.paddingLeft).toBe("4pt");
  });

  it("keeps a bottom-rule heading free of padding on the three edges it has no rule on", () => {
    const style = paragraphBorderStyle({ borders: { bottom: LINE } }, ALL_EDGES);
    expect(style.borderBottom).toBe("1pt solid #000000");
    expect(style.borderTop).toBe("none");
    expect(style.paddingTop).toBe("0");
    expect(style.paddingLeft).toBe("0");
    expect(style.paddingBottom).toBe("4pt");
  });

  it("fills the padding box with the shading, so it spans the content width (AC-13)", () => {
    const style = paragraphBorderStyle(boxed({ shadingFill: "#eeeeee" }), ALL_EDGES);
    expect(style.backgroundColor).toBe("#eeeeee");
    expect(style.paddingLeft).toBe("4pt");
  });

  it("draws the between edge as the fragment's bottom rule and closes the gap", () => {
    const formatting: ParagraphFormatting = {
      borders: { top: LINE, bottom: LINE, left: LINE, right: LINE, between: THICK },
    };
    const style = paragraphBorderStyle(formatting, {
      top: true,
      bottom: false,
      sides: true,
      between: true,
      continuesBelow: true,
      belowBetween: false,
      continuesAbove: false,
    });
    expect(style.borderBottom).toBe("3pt solid #ff0000");
    expect(style.marginBottom).toBe("0");
  });

  it("gives the paragraph under a between rule its top padding but no top rule or margin", () => {
    const formatting: ParagraphFormatting = {
      borders: { top: LINE, bottom: LINE, left: LINE, right: LINE, between: LINE },
    };
    const style = paragraphBorderStyle(formatting, {
      top: false,
      bottom: true,
      sides: true,
      between: false,
      continuesBelow: false,
      belowBetween: true,
      continuesAbove: true,
    });
    expect(style.borderTop).toBe("none");
    expect(style.paddingTop).toBe("4pt");
    expect(style.marginTop).toBe("0");
  });

  it("leaves a mid-paragraph page fragment with sides only and no vertical padding", () => {
    const style = paragraphBorderStyle(boxed(), {
      top: false,
      bottom: false,
      sides: true,
      between: false,
      continuesBelow: false,
      belowBetween: false,
      continuesAbove: false,
    });
    expect(style.borderLeft).toBe("1pt solid #000000");
    expect(style.borderTop).toBe("none");
    expect(style.borderBottom).toBe("none");
    expect(style.paddingTop).toBe("0");
    expect(style.paddingBottom).toBe("0");
  });

  // The box has to survive paragraph spacing. A "Box" run draws no internal
  // rule, so nothing tells the members apart visually — but each still
  // carries its own space-before/space-after, and any margin left standing
  // between two members breaks the left and right rules there and lets the
  // page show through the side of the box (AC-12).
  it("closes the internal boundaries of a run that has no between edge", () => {
    const spaced = boxed({ spaceBeforePt: 8, spaceAfterPt: 8 });
    const runs = borderRuns([para({ ...spaced }), para({ ...spaced }), para({ ...spaced })]);

    const first = paragraphBorderStyle(spaced, fragmentEdges(runs, 0, true, true));
    const middle = paragraphBorderStyle(spaced, fragmentEdges(runs, 1, true, true));
    const last = paragraphBorderStyle(spaced, fragmentEdges(runs, 2, true, true));

    // No internal rule anywhere — this is a plain box, not "All".
    expect([first, middle, last].map((s) => s.borderBottom)).toEqual([
      "none",
      "none",
      "1pt solid #000000",
    ]);
    // ...and no margin at either internal boundary.
    expect(first.marginBottom).toBe("0");
    expect(middle.marginTop).toBe("0");
    expect(middle.marginBottom).toBe("0");
    expect(last.marginTop).toBe("0");
    // The run's outer edges keep the paragraph's own spacing: the box as a
    // whole still sits 8pt from what surrounds it.
    expect(first.marginTop).toBeUndefined();
    expect(last.marginBottom).toBeUndefined();
  });

  it("keeps a lone bordered paragraph's own space before and after", () => {
    const spaced = boxed({ spaceBeforePt: 8, spaceAfterPt: 8 });
    const runs = borderRuns([para(spaced)]);
    const style = paragraphBorderStyle(spaced, fragmentEdges(runs, 0, true, true));
    expect(style.marginTop).toBeUndefined();
    expect(style.marginBottom).toBeUndefined();
  });

  it("keeps the margin between two runs that did not merge", () => {
    const a = boxed({ spaceAfterPt: 8 });
    const b = boxed({ spaceBeforePt: 8, shadingFill: "#eeeeee" });
    const runs = borderRuns([para(a), para(b)]);
    expect(paragraphBorderStyle(a, fragmentEdges(runs, 0, true, true)).marginBottom).toBeUndefined();
    expect(paragraphBorderStyle(b, fragmentEdges(runs, 1, true, true)).marginTop).toBeUndefined();
  });

  it("leaves a page-split seam alone — only the run's own boundaries collapse", () => {
    const spaced = boxed({ spaceBeforePt: 8, spaceAfterPt: 8 });
    const runs = borderRuns([para({ ...spaced }), para({ ...spaced })]);
    // The first paragraph continues onto the next page: its opening fragment
    // is not at the run's internal boundary, its closing fragment is.
    expect(paragraphBorderStyle(spaced, fragmentEdges(runs, 0, true, false)).marginBottom).toBeUndefined();
    expect(paragraphBorderStyle(spaced, fragmentEdges(runs, 0, false, true)).marginBottom).toBe("0");
    // ...and the second paragraph's continuation fragment is below a page
    // break, not below the paragraph it merged with.
    expect(paragraphBorderStyle(spaced, fragmentEdges(runs, 1, true, true)).marginTop).toBe("0");
    expect(paragraphBorderStyle(spaced, fragmentEdges(runs, 1, false, true)).marginTop).toBeUndefined();
  });
});

describe("presets (AC-8)", () => {
  it("detects each named preset", () => {
    expect(detectBorderPreset(undefined)).toBe("none");
    expect(detectBorderPreset({})).toBe("none");
    expect(detectBorderPreset({ top: LINE, bottom: LINE, left: LINE, right: LINE })).toBe("box");
    expect(detectBorderPreset({ top: LINE, bottom: LINE, left: LINE, right: LINE, between: LINE })).toBe("all");
    expect(detectBorderPreset({ top: LINE, bottom: LINE })).toBe("topAndBottom");
  });

  it("reports a hand-built edge set as custom", () => {
    expect(detectBorderPreset({ left: LINE })).toBe("custom");
    expect(detectBorderPreset({ top: LINE, bottom: THICK })).toBe("custom");
    expect(detectBorderPreset({ top: LINE, bottom: LINE, left: LINE, right: THICK })).toBe("custom");
  });

  it("round-trips every preset through applyBorderPreset", () => {
    for (const preset of ["box", "all", "topAndBottom"] as const) {
      expect(detectBorderPreset(applyBorderPreset(preset, LINE))).toBe(preset);
    }
    expect(applyBorderPreset("none", LINE)).toBeUndefined();
  });

  it("keeps the existing padding when a preset is applied", () => {
    const applied = applyBorderPreset("box", LINE, { paddingPt: { top: 10 } });
    expect(applied?.paddingPt).toEqual({ top: 10 });
  });

  it("leaves a custom edge set alone", () => {
    const existing: ParagraphBorders = { left: THICK };
    expect(applyBorderPreset("custom", LINE, existing)).toBe(existing);
  });
});

describe("normaliseParagraphBorders", () => {
  it("drops an object with nothing left to paint", () => {
    expect(normaliseParagraphBorders(undefined)).toBeUndefined();
    expect(normaliseParagraphBorders({ paddingPt: { top: 8 } })).toBeUndefined();
    expect(normaliseParagraphBorders({ top: { style: "none", widthPt: 0, color: "auto" } })).toBeUndefined();
  });

  it("keeps an object with a visible edge", () => {
    const borders: ParagraphBorders = { top: LINE };
    expect(normaliseParagraphBorders(borders)).toBe(borders);
  });
});

describe("formatting marks on a shaded paragraph (AC-26)", () => {
  it("says nothing about a paragraph with no shading", () => {
    expect(paragraphMarkColorStyle(undefined)).toEqual({});
    expect(paragraphMarkColorStyle({ borders: { top: LINE } })).toEqual({});
  });

  it("keeps the default grey on a light fill", () => {
    expect(paragraphMarkColorStyle({ shadingFill: "#f5f5f5" })["--mark-color"]).toBe(DEFAULT_MARK_COLOR);
    expect(paragraphMarkColorStyle({ shadingFill: "#fff2cc" })["--mark-color"]).toBe(DEFAULT_MARK_COLOR);
  });

  it("flips to a light grey on a dark fill, so the pilcrow stays visible", () => {
    expect(paragraphMarkColorStyle({ shadingFill: "#1f1f1f" })["--mark-color"]).toBe("#d0d3d6");
    expect(paragraphMarkColorStyle({ shadingFill: "#20365c" })["--mark-color"]).toBe("#d0d3d6");
  });

  it("rides on the paragraph's own style object, beside the shading", () => {
    const style = paragraphBorderStyle({ shadingFill: "#1f1f1f" }, ALL_EDGES);
    expect(style.backgroundColor).toBe("#1f1f1f");
    expect(style["--mark-color"]).toBe("#d0d3d6");
  });

  it("reads luminance, not brightness of one channel", () => {
    expect(isDarkFill("#ffff00")).toBe(false);
    expect(isDarkFill("#0000ff")).toBe(true);
    expect(isDarkFill("not-a-colour")).toBe(false);
  });
});
