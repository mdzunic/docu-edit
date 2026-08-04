import { describe, it, expect } from "vitest";
import { headingTag, paragraphStyle, runKey, runStyle } from "./renderFormatting";

describe("headingTag", () => {
  it("returns 'p' for no heading level", () => {
    expect(headingTag(undefined)).toBe("p");
  });

  it("returns hN for a heading level", () => {
    expect(headingTag(1)).toBe("h1");
    expect(headingTag(3)).toBe("h3");
  });

  it("clamps levels above 6 to h6", () => {
    expect(headingTag(9)).toBe("h6");
  });
});

describe("runStyle", () => {
  it("maps bold, italic, underline and strikethrough", () => {
    const style = runStyle({ bold: true, italic: true, underline: true, strikethrough: true });
    expect(style.fontWeight).toBe("bold");
    expect(style.fontStyle).toBe("italic");
    expect(style.textDecorationLine).toBe("underline line-through");
  });

  it("maps font family, size and colour", () => {
    const style = runStyle({ fontFamily: "Arial", fontSizePt: 14, color: "#1155cc" });
    expect(style.fontFamily).toBe("Arial");
    expect(style.fontSize).toBe("14pt");
    expect(style.color).toBe("#1155cc");
  });

  it("maps a known highlight name to its CSS colour", () => {
    expect(runStyle({ highlight: "yellow" }).backgroundColor).toBe("#ffff00");
  });

  it("falls back to the raw value for an unknown highlight name", () => {
    expect(runStyle({ highlight: "mystery" }).backgroundColor).toBe("mystery");
  });

  it("returns an empty style object for empty formatting", () => {
    expect(runStyle({})).toEqual({});
  });

  it("paints nothing for an explicit 'No colour' highlight (SPEC-013)", () => {
    expect(runStyle({ highlight: false }).backgroundColor).toBeUndefined();
  });

  it("paints an out-of-palette highlight hex as-is (SPEC-013)", () => {
    expect(runStyle({ highlight: "#ff8800" }).backgroundColor).toBe("#ff8800");
  });

  it("maps the white palette entry, which the picker hides (SPEC-013)", () => {
    expect(runStyle({ highlight: "white" }).backgroundColor).toBe("#ffffff");
  });

  it("renders sub/superscript shifted and reduced (SPEC-013)", () => {
    expect(runStyle({ vertAlign: "subscript" }).verticalAlign).toBe("sub");
    expect(runStyle({ vertAlign: "subscript" }).fontSize).toBe("0.65em");
    const sized = runStyle({ vertAlign: "superscript", fontSizePt: 20 });
    expect(sized.verticalAlign).toBe("super");
    expect(sized.fontSize).toBe("13pt");
  });

  it("never paints the shadingFill preservation slot (SPEC-013)", () => {
    expect(runStyle({ highlight: "yellow", shadingFill: "#ff8800" }).backgroundColor).toBe(
      "#ffff00",
    );
  });
});

describe("paragraphStyle", () => {
  it("always sets white-space: pre-wrap so tabs/line breaks render", () => {
    expect(paragraphStyle({ type: "paragraph", runs: [] }).whiteSpace).toBe("pre-wrap");
  });

  it("maps alignment, indents and spacing", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: {
        alignment: "center",
        indentStartCm: 1.27,
        indentEndCm: 1.27,
        indentFirstLineCm: 0.5,
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        lineSpacing: 1.5,
      },
    });
    expect(style.textAlign).toBe("center");
    expect(style.marginInlineStart).toBe("1.27cm");
    expect(style.marginInlineEnd).toBe("1.27cm");
    expect(style.textIndent).toBe("0.5cm");
    expect(style.marginTop).toBe("12pt");
    expect(style.marginBottom).toBe("6pt");
    expect(style.lineHeight).toBe("1.5");
  });

  it("prefers an absolute lineSpacingPt over a lineSpacing multiplier", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { lineSpacing: 1.5, lineSpacingPt: 14 },
    });
    expect(style.lineHeight).toBe("14pt");
  });

  it("sets direction: rtl so logical indents mirror physically (SPEC-014)", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { direction: "rtl", indentStartCm: 2, indentEndCm: 1 },
    });
    expect(style.direction).toBe("rtl");
    // The values stay logical — the browser's inline axis does the mirroring.
    expect(style.marginInlineStart).toBe("2cm");
    expect(style.marginInlineEnd).toBe("1cm");
  });

  it("leaves direction unset for LTR paragraphs", () => {
    const style = paragraphStyle({
      type: "paragraph",
      runs: [],
      formatting: { indentStartCm: 2 },
    });
    expect(style.direction).toBeUndefined();
  });
});

describe("runKey", () => {
  it("produces a stable key from index and text length", () => {
    expect(runKey({ text: "hello", formatting: {} }, 2)).toBe("2-5");
  });
});
