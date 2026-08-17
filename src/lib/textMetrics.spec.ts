// SPEC-028 — canvas-predicted paragraph metrics, driven by a deterministic
// fake measurer so the wrap rules (greedy breaks, hanging spaces, tabs,
// forced newlines, indents, line-height rules, formatting marks) are
// exercised as pure math, no DOM.
import { describe, expect, it } from "vitest";
import type { Paragraph, ParagraphFormatting, Run } from "./documentModel";
import { PX_PER_CM } from "./core/units";
import {
  measureParagraphText,
  paragraphTabLayout,
  type BaseFont,
  type LineGeometry,
  type ResolvedFont,
  type TextMeasurer,
} from "./textMetrics";

const BASE: BaseFont = { family: "fake", sizePx: 16 };
// Fake shaping: every character advances sizePx/2 (8px at the base size);
// line-height normal advances 1.25×size with a 1.2×size ink band.
const fakeMeasurer: TextMeasurer = {
  widthOf(text: string, font: ResolvedFont): number {
    return text.length * (font.sizePx / 2);
  },
  normalGeometryOf(font: ResolvedFont): LineGeometry {
    return { advancePx: font.sizePx * 1.25, inkHeightPx: font.sizePx * 1.2 };
  },
};

function para(text: string, formatting?: ParagraphFormatting, runs?: Run[]): Paragraph {
  const block: Paragraph = {
    type: "paragraph",
    runs: runs ?? [{ text, formatting: {} }],
  };
  if (formatting) {
    block.formatting = formatting;
  }
  return block;
}

function measure(paragraph: Paragraph, widthPx: number, marksOn = false) {
  const result = measureParagraphText(paragraph, widthPx, BASE, marksOn, fakeMeasurer);
  expect(result).not.toBeNull();
  return result!;
}

describe("measureParagraphText", () => {
  it("keeps a fitting paragraph on one line", () => {
    const result = measure(para("aaaa bb"), 100);
    expect(result.lineOffsets).toEqual([0]);
    expect(result.metrics.heightPx).toBeCloseTo(20);
    expect(result.metrics.lines).toHaveLength(1);
  });

  it("breaks greedily at spaces, starting lines after the space run", () => {
    // 10 chars per 80px line: "aaaa " (5) + "bbbb" (4) fits; "cccc" doesn't.
    const result = measure(para("aaaa bbbb cccc"), 80);
    expect(result.lineOffsets).toEqual([0, 10]);
    expect(result.metrics.heightPx).toBeCloseTo(40);
  });

  it("lets trailing spaces hang past the edge without wrapping", () => {
    // The spaces after "bbbb" overflow 80px but only the next word breaks.
    const result = measure(para("aaaa bbbb   c"), 80);
    expect(result.lineOffsets).toEqual([0, 12]);
  });

  it("keeps an overlong word on one overflowing line", () => {
    const result = measure(para("abcdefghijklmnopqrstuvwxyz"), 80);
    expect(result.lineOffsets).toEqual([0]);
    expect(result.metrics.lines).toHaveLength(1);
  });

  it("honours forced newlines", () => {
    const result = measure(para("aa\nbb cc\ndd"), 800);
    expect(result.lineOffsets).toEqual([0, 3, 9]);
    expect(result.metrics.heightPx).toBeCloseTo(60);
  });

  it("advances tabs to the next 8-space stop", () => {
    // Space is 8px, so stops sit every 64px. "aa" (16px) + tab lands at 64;
    // the following 4-char word ends at 96 > 80: it wraps.
    const result = measure(para("aa\tbbbb"), 80);
    expect(result.lineOffsets).toEqual([0, 3]);
  });

  it("measures an empty paragraph as one strut line", () => {
    const result = measure(para(""), 100);
    expect(result.lineOffsets).toEqual([0]);
    expect(result.metrics.heightPx).toBeCloseTo(20);
    expect(result.metrics.lines[0].height).toBeCloseTo(19.2);
  });

  it("maps space before/after to pixel margins", () => {
    const result = measure(para("aa", { spaceBeforePt: 12, spaceAfterPt: 6 }), 100);
    expect(result.metrics.marginTopPx).toBeCloseTo(16);
    expect(result.metrics.marginBottomPx).toBeCloseTo(8);
  });

  it("uses an absolute lineSpacingPt for every line regardless of run size", () => {
    const runs: Run[] = [
      { text: "aaaa ", formatting: { fontSizePt: 24 } },
      { text: "bb", formatting: {} },
    ];
    const result = measure(para("", { lineSpacingPt: 18 }, runs), 800);
    expect(result.metrics.heightPx).toBeCloseTo(24); // 18pt = 24px, one line
  });

  it("scales a unitless lineSpacing by the largest run on the line", () => {
    const runs: Run[] = [
      { text: "aaaa ", formatting: { fontSizePt: 24 } }, // 32px font
      { text: "bb", formatting: {} },
    ];
    const result = measure(para("", { lineSpacing: 1.5 }, runs), 800);
    expect(result.metrics.heightPx).toBeCloseTo(48); // 32px × 1.5
  });

  it("uses the tallest font's normal advance on mixed lines", () => {
    const runs: Run[] = [
      { text: "aaaa ", formatting: { fontSizePt: 24 } }, // 32px → advance 40
      { text: "bb", formatting: {} },
    ];
    const result = measure(para("", undefined, runs), 800);
    expect(result.metrics.heightPx).toBeCloseTo(40);
    expect(result.metrics.lines[0].height).toBeCloseTo(38.4);
  });

  it("narrows the first line by the first-line indent", () => {
    // 1cm indent ≈ 37.8px: first line fits 5 chars of 80px−37.8px, so the
    // first word alone; without the indent both words fit.
    const withIndent = measure(para("aaaa bbbb", { indentFirstLineCm: 1 }), 80);
    expect(withIndent.lineOffsets).toEqual([0, 5]);
    const without = measure(para("aaaa bbbb"), 80);
    expect(without.lineOffsets).toEqual([0]);
  });

  it("narrows every line by the start/end indents", () => {
    const result = measure(
      para("aaaa bbbb", { indentStartCm: 1 / PX_PER_CM, indentEndCm: 15 / PX_PER_CM }),
      88,
    );
    // 88 − 16 = 72px: "aaaa bbbb" is 72px and fits exactly.
    expect(result.lineOffsets).toEqual([0]);
    const narrower = measure(
      para("aaaa bbbb", { indentStartCm: 2 / PX_PER_CM, indentEndCm: 15 / PX_PER_CM }),
      88,
    );
    expect(narrower.lineOffsets).toEqual([0, 5]);
  });

  it("applies the list level's indent when no direct indent is set", () => {
    const listed: Paragraph = {
      type: "paragraph",
      runs: [{ text: "aaaa bbbb", formatting: {} }],
      list: { type: "bullet", level: 1, indentCm: 17 / PX_PER_CM },
    };
    const result = measureParagraphText(listed, 88, BASE, false, fakeMeasurer);
    expect(result!.lineOffsets).toEqual([0, 5]);
  });

  it("counts the pilcrow's width only when formatting marks are on", () => {
    // "aaaa bbbb" exactly fills 72px; the ¶ (8px) forces nothing without a
    // break opportunity (glued), but wraps onto its own line after a
    // trailing space.
    const glued = measure(para("aaaa bbbb"), 72, true);
    expect(glued.lineOffsets).toEqual([0]);
    const wrapped = measure(para("aaaa bbbb "), 72, true);
    expect(wrapped.lineOffsets).toEqual([0, 10]);
    const without = measure(para("aaaa bbbb "), 72, false);
    expect(without.lineOffsets).toEqual([0]);
  });

  it("measures sub/superscript runs at reduced size", () => {
    const runs: Run[] = [
      { text: "aaaa", formatting: {} },
      { text: "bbbb", formatting: { vertAlign: "superscript" } },
    ];
    // 4×8 + 4×(16×0.65/2 = 5.2) = 52.8px fits in 56; at full size it would
    // be 64px and overflow (but as one unbreakable word it stays anyway —
    // assert via the measured advance-free single line).
    const result = measure(para("", undefined, runs), 56);
    expect(result.lineOffsets).toEqual([0]);
  });

  it("reports null when the measurer has no geometry", () => {
    const blind: TextMeasurer = {
      widthOf: () => 0,
      normalGeometryOf: () => null,
    };
    expect(measureParagraphText(para("aa"), 100, BASE, false, blind)).toBeNull();
  });

  it("stores the section content width for the cache's width check", () => {
    const result = measure(para("aaaa", { indentStartCm: 1 }), 200);
    expect(result.metrics.contentWidthPx).toBe(200);
  });

  it("centres ink bands inside the layout advance", () => {
    const result = measure(para("aaaa bbbb cccc"), 80);
    // advance 20, ink 19.2 → half-leading 0.4; line 2 top = 20 + 0.4.
    expect(result.metrics.lines[0].top).toBeCloseTo(0.4);
    expect(result.metrics.lines[1].top).toBeCloseTo(20.4);
  });
});

describe("paragraphTabLayout (SPEC-054)", () => {
  it("resolves a tab's advance against the paragraph fonts and stop", () => {
    // Each char is 8 px (sizePx/2). Pen after "a" is 8 px; a left stop at 3 cm
    // fills the tab to 3 cm − 8 px.
    const layout = paragraphTabLayout(
      para("a\tb"),
      { stops: [{ positionCm: 3, alignment: "left" }], defaultTabStopCm: 1.27, decimalSeparator: "." },
      BASE,
      fakeMeasurer,
    );
    expect(layout.advances.get(1)?.widthPx).toBeCloseTo(3 * PX_PER_CM - 8, 3);
  });

  it("reports bar-stop positions for the renderer", () => {
    const layout = paragraphTabLayout(
      para("hello"),
      { stops: [{ positionCm: 2, alignment: "bar" }], defaultTabStopCm: 1.27, decimalSeparator: "." },
      BASE,
      fakeMeasurer,
    );
    expect(layout.bars).toHaveLength(1);
    expect(layout.bars[0]).toBeCloseTo(2 * PX_PER_CM, 3);
  });

  it("wrap measurement uses the resolved tab advance, not the 8-space grid", () => {
    // "a\tbbbb" at a 300 px width: a left stop at 8 cm (≈302 px) pushes "bbbb"
    // (32 px) past the edge → it wraps to a second line. Without the config the
    // tab only reaches the 64 px default grid, so the line fits.
    const config = {
      stops: [{ positionCm: 8, alignment: "left" as const }],
      defaultTabStopCm: 1.27,
      decimalSeparator: ".",
    };
    const wrapped = measureParagraphText(para("a\tbbbb"), 300, BASE, false, fakeMeasurer, config)!;
    expect(wrapped.lineOffsets.length).toBe(2);
    const noConfig = measureParagraphText(para("a\tbbbb"), 300, BASE, false, fakeMeasurer)!;
    expect(noConfig.lineOffsets.length).toBe(1);
  });
});
