import { describe, it, expect } from "vitest";
import { createDocument } from "./documentModel";
import {
  documentFontFamilies,
  FONT_SIZE_RANGE_MESSAGE,
  MAX_FONT_SIZE_PT,
  MIN_FONT_SIZE_PT,
  parseFontSizeInput,
  stepFontSize,
} from "./fonts";
import { STANDARD_FONT_FAMILIES } from "./core/fonts";

describe("STANDARD_FONT_FAMILIES", () => {
  it("includes the web-safe families and the common document fonts (AC)", () => {
    for (const family of [
      "Calibri",
      "Cambria",
      "Times New Roman",
      "Arial",
      "Georgia",
      "Verdana",
      "Courier New",
      "Liberation Serif",
      "Liberation Sans",
      "Carlito",
    ]) {
      expect(STANDARD_FONT_FAMILIES).toContain(family);
    }
  });
});

describe("documentFontFamilies", () => {
  it("lists the document default plus every family a run names, sorted", () => {
    const model = createDocument("fonts-test");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [
              { text: "a", formatting: { fontFamily: "Zapfino" } },
              { text: "b", formatting: {} },
              { text: "c", formatting: { fontFamily: "Aptos" } },
            ],
          },
        ],
      },
    ];
    // Calibri is the model default; Zapfino/Aptos need not be installed —
    // the list comes from the model (AC "not installed locally").
    expect(documentFontFamilies(model)).toEqual(["Aptos", "Calibri", "Zapfino"]);
  });

  it("de-duplicates case-insensitively, first spelling wins", () => {
    const model = createDocument("fonts-dedupe");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [
              { text: "a", formatting: { fontFamily: "GEORGIA" } },
              { text: "b", formatting: { fontFamily: "georgia" } },
            ],
          },
        ],
      },
    ];
    expect(documentFontFamilies(model)).toEqual(["Calibri", "GEORGIA"]);
  });
});

describe("parseFontSizeInput", () => {
  it("accepts in-range values and decimal commas", () => {
    expect(parseFontSizeInput("11.5")).toEqual({ value: 11.5, clamped: false });
    expect(parseFontSizeInput(" 12 ")).toEqual({ value: 12, clamped: false });
    expect(parseFontSizeInput("11,5")).toEqual({ value: 11.5, clamped: false });
  });

  it("snaps to the 0.5 pt grid (AC)", () => {
    expect(parseFontSizeInput("11.3")).toEqual({ value: 11.5, clamped: false });
    expect(parseFontSizeInput("11.2")).toEqual({ value: 11, clamped: false });
  });

  it("clamps out-of-range input and flags it for the tooltip (AC)", () => {
    expect(parseFontSizeInput("500")).toEqual({ value: MAX_FONT_SIZE_PT, clamped: true });
    expect(parseFontSizeInput("0.4")).toEqual({ value: MIN_FONT_SIZE_PT, clamped: true });
    expect(parseFontSizeInput("-3")).toEqual({ value: MIN_FONT_SIZE_PT, clamped: true });
  });

  it("rejects non-numeric input without clamping", () => {
    expect(parseFontSizeInput("abc")).toEqual({ clamped: false });
    expect(parseFontSizeInput("")).toEqual({ clamped: false });
  });

  it("has a range message naming the limits", () => {
    expect(FONT_SIZE_RANGE_MESSAGE).toContain("1");
    expect(FONT_SIZE_RANGE_MESSAGE).toContain("409");
  });
});

describe("stepFontSize", () => {
  it("steps by half a point (AC)", () => {
    expect(stepFontSize(11, 1)).toBe(11.5);
    expect(stepFontSize(11.5, 1)).toBe(12);
    expect(stepFontSize(11, -1)).toBe(10.5);
  });

  it("snaps an off-grid size to the next grid point", () => {
    expect(stepFontSize(11.3, 1)).toBe(11.5);
    expect(stepFontSize(11.3, -1)).toBe(11);
  });

  it("clamps at the range ends (AC)", () => {
    expect(stepFontSize(MAX_FONT_SIZE_PT, 1)).toBe(MAX_FONT_SIZE_PT);
    expect(stepFontSize(MIN_FONT_SIZE_PT, -1)).toBe(MIN_FONT_SIZE_PT);
  });
});
