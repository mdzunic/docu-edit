import { describe, expect, it } from "vitest";
import type { ShapeFrame, TextFrame } from "./model";
import {
  AUTOFIT_STEPS,
  DEFAULT_INSETS_CM,
  applyTextBody,
  availableTextHeightCm,
  effectiveAnchor,
  effectiveInsetsCm,
  fitScale,
  lineScaleFor,
  quantizeScalePct,
  resizedHeightCm,
  storedScale,
} from "./textBody";

function textFrame(extra: Partial<TextFrame> = {}): TextFrame {
  return { id: "t", kind: "text", xCm: 0, yCm: 0, widthCm: 10, heightCm: 4, paragraphs: [], ...extra };
}
function shapeFrame(extra: Partial<ShapeFrame> = {}): ShapeFrame {
  return { id: "s", kind: "shape", geometry: "rect", xCm: 0, yCm: 0, widthCm: 10, heightCm: 4, ...extra };
}

/** A synthetic measurer: content of `natural` cm at 100% scaling linearly
 * with the font scale and the line scale. */
const linear = (natural: number) => (font: number, line: number): number => (natural * font * line) / 10000;

describe("fitScale (the deterministic autofit ladder)", () => {
  it("keeps 100/100 when the content already fits", () => {
    expect(fitScale(linear(3), 3.5)).toEqual({ fontScalePct: 100, lineScalePct: 100 });
  });

  it("takes the first step that fits, pairing 90% line spacing at ≤70 and 80% at ≤50", () => {
    expect(fitScale(linear(4), 3.5)).toEqual({ fontScalePct: 85, lineScalePct: 100 });
    expect(fitScale(linear(6), 3.5)).toEqual({ fontScalePct: 60, lineScalePct: 90 });
    expect(fitScale(linear(10), 3.5)).toEqual({ fontScalePct: 40, lineScalePct: 80 });
  });

  it("falls back to the last step when nothing fits", () => {
    expect(fitScale(() => 100, 1)).toEqual({ fontScalePct: 25, lineScalePct: 80 });
  });

  it("walks the ladder in order and stops at the first fit", () => {
    const calls: number[] = [];
    fitScale((font) => {
      calls.push(font);
      return font <= 77.5 ? 1 : 10;
    }, 2);
    expect(calls).toEqual([100, 92.5, 85, 77.5]);
    expect(AUTOFIT_STEPS[AUTOFIT_STEPS.length - 1]).toBe(25);
    expect([lineScaleFor(100), lineScaleFor(70), lineScaleFor(50)]).toEqual([100, 90, 80]);
  });
});

describe("inset / anchor defaults", () => {
  it("defaults to the pptx EMU insets and the kind's anchor", () => {
    expect(effectiveInsetsCm(textFrame())).toEqual(DEFAULT_INSETS_CM);
    expect(effectiveAnchor(textFrame())).toBe("top");
    expect(effectiveAnchor(shapeFrame())).toBe("middle");
    expect(availableTextHeightCm(textFrame())).toBeCloseTo(4 - 0.254, 6);
  });

  it("applyTextBody drops values equal to the defaults and removes an empty body", () => {
    const frame = textFrame();
    applyTextBody(frame, { anchor: "middle", insetsCm: { left: 1, top: 0.5, right: 1, bottom: 0.5 }, columns: { count: 2, gapCm: 0.5 } });
    expect(frame.body).toEqual({ anchor: "middle", insetsCm: { left: 1, top: 0.5, right: 1, bottom: 0.5 }, columns: { count: 2, gapCm: 0.5 } });
    applyTextBody(frame, { anchor: "top", insetsCm: { ...DEFAULT_INSETS_CM }, columns: { count: 1, gapCm: 0 } });
    expect(frame.body).toBeUndefined();
    const shape = shapeFrame();
    applyTextBody(shape, { anchor: "middle" });
    expect(shape.body).toBeUndefined();
  });

  it("clamps columns to 1–16 and keeps the scale only with shrink", () => {
    const frame = textFrame();
    applyTextBody(frame, { columns: { count: 40, gapCm: 1 } });
    expect(frame.body?.columns?.count).toBe(16);
    applyTextBody(frame, { autofit: "shrink", fontScalePct: 62.5, lineScalePct: 80 });
    expect(storedScale(frame)).toEqual({ fontScalePct: 62.5, lineScalePct: 80 });
    applyTextBody(frame, { autofit: "none" });
    expect(frame.body).toEqual({ columns: { count: 16, gapCm: 1 } });
    expect(storedScale(frame)).toEqual({ fontScalePct: 100, lineScalePct: 100 });
  });

  it("editing autofit or anchor clears the preserved bodyPr, insets/columns do not", () => {
    const frame = textFrame({ pptxBodyPr: '<a:bodyPr vert="vert270"/>', odpBodyProps: 'style:writing-mode="tb-rl"' });
    applyTextBody(frame, { insetsCm: { left: 0, top: 0, right: 0, bottom: 0 } });
    applyTextBody(frame, { columns: { count: 2, gapCm: 1 } });
    expect(frame.pptxBodyPr).toBe('<a:bodyPr vert="vert270"/>');
    applyTextBody(frame, { autofit: "shrink" });
    expect(frame.pptxBodyPr).toBeUndefined();
    expect(frame.odpBodyProps).toBeUndefined();
  });
});

describe("resizeShape and quantisation", () => {
  it("resizes to the content height plus insets, floored at 0.5cm and rounded r3", () => {
    expect(resizedHeightCm(textFrame(), 2.12345)).toBe(2.377);
    expect(resizedHeightCm(textFrame({ body: { insetsCm: { left: 0, top: 0, right: 0, bottom: 0 } } }), 0.1)).toBe(0.5);
  });

  it("quantises a scale to 1000ths of a percent (pptx's sanctioned delta)", () => {
    expect(quantizeScalePct(62.50049)).toBe(62.5);
    expect(quantizeScalePct(33.33333)).toBe(33.333);
  });
});
