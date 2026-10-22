import { describe, it, expect } from "vitest";
import { markerPositions, valueFromMarkerPosition, type RulerContext } from "./ruler";

// A4 with standard margins: printable width 21 - 2.54 - 2.54 = 15.92 cm.
const A4: RulerContext = { pageWidthCm: 21, marginLeftCm: 2.54, marginRightCm: 2.54 };

describe("markerPositions (SPEC-014 AC-22, AC-25)", () => {
  it("places LTR markers from the left margin rightward", () => {
    const positions = markerPositions(A4, {
      indentStartCm: 1.27,
      indentEndCm: 0.5,
      indentFirstLineCm: 0.5,
    });
    expect(positions.marginLeft).toBeCloseTo(2.54, 10);
    expect(positions.marginRight).toBeCloseTo(18.46, 10);
    expect(positions.start).toBeCloseTo(3.81, 10);
    expect(positions.firstLine).toBeCloseTo(4.31, 10);
    expect(positions.end).toBeCloseTo(17.96, 10);
  });

  it("mirrors for RTL: the start indent hangs off the right margin (AC-25)", () => {
    const positions = markerPositions(A4, {
      direction: "rtl",
      indentStartCm: 1.27,
      indentEndCm: 0.5,
      indentFirstLineCm: 0.5,
    });
    expect(positions.start).toBeCloseTo(18.46 - 1.27, 10);
    // First-line grows in the text direction — leftward for RTL.
    expect(positions.firstLine).toBeCloseTo(18.46 - 1.27 - 0.5, 10);
    expect(positions.end).toBeCloseTo(2.54 + 0.5, 10);
  });

  it("treats absent indents as zero — markers sit on the margins", () => {
    const positions = markerPositions(A4, {});
    expect(positions.start).toBeCloseTo(2.54, 10);
    expect(positions.firstLine).toBeCloseTo(2.54, 10);
    expect(positions.end).toBeCloseTo(18.46, 10);
  });
});

describe("valueFromMarkerPosition (SPEC-014 AC-23, AC-25)", () => {
  it("is the inverse of markerPositions for every LTR marker", () => {
    const formatting = { indentStartCm: 1.25, indentEndCm: 0.5, indentFirstLineCm: 0.75 };
    const positions = markerPositions(A4, formatting);
    expect(valueFromMarkerPosition("start", positions.start, A4, formatting)).toEqual({
      kind: "indentStartCm",
      valueCm: 1.25,
    });
    expect(valueFromMarkerPosition("end", positions.end, A4, formatting)).toEqual({
      kind: "indentEndCm",
      valueCm: 0.5,
    });
    expect(valueFromMarkerPosition("first-line", positions.firstLine, A4, formatting)).toEqual({
      kind: "indentFirstLineCm",
      valueCm: 0.75,
    });
  });

  it("is the inverse of markerPositions for every RTL marker", () => {
    const formatting = {
      direction: "rtl" as const,
      indentStartCm: 1.25,
      indentEndCm: 0.5,
      indentFirstLineCm: -0.75,
    };
    const positions = markerPositions(A4, formatting);
    expect(valueFromMarkerPosition("start", positions.start, A4, formatting).valueCm).toBeCloseTo(1.25, 10);
    expect(valueFromMarkerPosition("end", positions.end, A4, formatting).valueCm).toBeCloseTo(0.5, 10);
    expect(
      valueFromMarkerPosition("first-line", positions.firstLine, A4, formatting).valueCm,
    ).toBeCloseTo(-0.75, 10);
  });

  it("dragging the LTR start marker right increases the start indent", () => {
    const update = valueFromMarkerPosition("start", 5.04, A4, {});
    expect(update).toEqual({ kind: "indentStartCm", valueCm: 2.5 });
  });

  it("dragging the RTL start marker LEFT increases the start indent", () => {
    const update = valueFromMarkerPosition("start", 16.46, A4, { direction: "rtl" });
    expect(update.kind).toBe("indentStartCm");
    expect(update.valueCm).toBeCloseTo(2, 10);
  });

  it("clamps indents to non-negative and inside the printable width", () => {
    expect(valueFromMarkerPosition("start", 0, A4, {}).valueCm).toBe(0);
    expect(valueFromMarkerPosition("start", 21, A4, {}).valueCm).toBeCloseTo(15.92, 10);
    expect(valueFromMarkerPosition("end", 21, A4, {}).valueCm).toBe(0);
  });

  it("allows a negative first-line value (hanging indent) down to the page edge", () => {
    const update = valueFromMarkerPosition("first-line", 0, A4, { indentStartCm: 1 });
    expect(update.kind).toBe("indentFirstLineCm");
    expect(update.valueCm).toBeCloseTo(-3.54, 10);
  });

  it("snaps drag positions to 0.05 cm steps", () => {
    const update = valueFromMarkerPosition("start", 3.831, A4, {});
    expect(update.valueCm).toBeCloseTo(1.3, 10);
  });

  it("maps margin markers to the page margins", () => {
    expect(valueFromMarkerPosition("margin-left", 3, A4, {})).toEqual({
      kind: "marginLeftCm",
      valueCm: 3,
    });
    expect(valueFromMarkerPosition("margin-right", 19, A4, {})).toEqual({
      kind: "marginRightCm",
      valueCm: 2,
    });
  });
});
