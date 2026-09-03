import { describe, expect, it } from "vitest";
import {
  BORDER_LINE_STYLES,
  borderColorCss,
  borderSideCss,
  borderSideVisible,
  borderSidesEqual,
  borderStyleCss,
  type BorderLineStyle,
  type BorderSide,
} from "./borders";
import type { BorderLineStyle as ReExportedStyle, BorderSide as ReExportedSide } from "./tables";

function side(style: BorderLineStyle, widthPt = 1, color = "#000000"): BorderSide {
  return { style, widthPt, color };
}

describe("border line styles (SPEC-074 AC-3)", () => {
  it("includes the three styles SPEC-074 adds to SPEC-023's five", () => {
    expect(BORDER_LINE_STYLES).toContain("thick");
    expect(BORDER_LINE_STYLES).toContain("dashDot");
    expect(BORDER_LINE_STYLES).toContain("dashDotDot");
  });

  it("keeps every style SPEC-023 shipped", () => {
    for (const style of ["none", "single", "double", "dashed", "dotted"] as const) {
      expect(BORDER_LINE_STYLES).toContain(style);
    }
  });

  it("re-exports the types through core/tables so SPEC-023's imports still compile (AC-2)", () => {
    // Type-level assertion: assigning across the two import paths only
    // compiles if they are the same type.
    const viaTables: ReExportedSide = { style: "dashDot" as ReExportedStyle, widthPt: 1, color: "auto" };
    const viaBorders: BorderSide = viaTables;
    expect(viaBorders.style).toBe("dashDot");
  });
});

describe("borderStyleCss", () => {
  it("maps the dash-dot family to CSS's nearest equivalent, dashed", () => {
    expect(borderStyleCss("dashDot")).toBe("dashed");
    expect(borderStyleCss("dashDotDot")).toBe("dashed");
    expect(borderStyleCss("dashed")).toBe("dashed");
  });

  it("renders thick as a solid line — its weight is its width", () => {
    expect(borderStyleCss("thick")).toBe("solid");
    expect(borderStyleCss("single")).toBe("solid");
  });

  it("keeps double and dotted", () => {
    expect(borderStyleCss("double")).toBe("double");
    expect(borderStyleCss("dotted")).toBe("dotted");
  });
});

describe("borderSideCss", () => {
  it("paints an unset, none or zero-width edge as none", () => {
    expect(borderSideCss(undefined)).toBe("none");
    expect(borderSideCss(side("none", 1))).toBe("none");
    expect(borderSideCss(side("single", 0))).toBe("none");
  });

  it("emits a CSS shorthand with the width in points", () => {
    expect(borderSideCss(side("single", 1.5, "#ff0000"))).toBe("1.5pt solid #ff0000");
  });

  it("resolves the automatic colour to black", () => {
    expect(borderSideCss(side("single", 1, "auto"))).toBe("1pt solid #000000");
    expect(borderColorCss("auto")).toBe("#000000");
    expect(borderColorCss("#123456")).toBe("#123456");
  });

  it("emits the model's own width unchanged — SPEC-023's cell rules depend on it", () => {
    expect(borderSideCss(side("double", 1.5, "#00ff00"))).toBe("1.5pt double #00ff00");
    expect(borderSideCss(side("dashDot", 2))).toBe("2pt dashed #000000");
  });
});

describe("borderSidesEqual", () => {
  it("treats two absent sides as equal and absent-vs-present as different", () => {
    expect(borderSidesEqual(undefined, undefined)).toBe(true);
    expect(borderSidesEqual(undefined, side("none", 0))).toBe(false);
  });

  it("compares style, width and colour structurally", () => {
    expect(borderSidesEqual(side("single", 1, "#000000"), side("single", 1, "#000000"))).toBe(true);
    expect(borderSidesEqual(side("single", 1), side("dashed", 1))).toBe(false);
    expect(borderSidesEqual(side("single", 1), side("single", 2))).toBe(false);
    expect(borderSidesEqual(side("single", 1, "#000000"), side("single", 1, "auto"))).toBe(false);
  });

  it("ignores float noise below a thousandth of a point", () => {
    expect(borderSidesEqual(side("single", 0.5), side("single", 0.5000001))).toBe(true);
  });
});

describe("borderSideVisible", () => {
  it("is false for unset, none and zero-width edges", () => {
    expect(borderSideVisible(undefined)).toBe(false);
    expect(borderSideVisible(side("none", 2))).toBe(false);
    expect(borderSideVisible(side("single", 0))).toBe(false);
    expect(borderSideVisible(side("single", 0.5))).toBe(true);
  });
});
