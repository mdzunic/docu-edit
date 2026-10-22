import { describe, it, expect } from "vitest";
import { isCodeStyledFont } from "./skip";

describe("isCodeStyledFont (SPEC-035 AC-12.1, AC-12.2)", () => {
  it("recognises common monospace font families", () => {
    for (const font of ["Courier New", "Consolas", "Monaco", "Menlo", "monospace", "Source Code Pro", "Fira Code"]) {
      expect(isCodeStyledFont(font)).toBe(true);
    }
  });

  it("does not flag ordinary body fonts", () => {
    for (const font of ["Calibri", "Georgia", "Times New Roman", "Arial"]) {
      expect(isCodeStyledFont(font)).toBe(false);
    }
  });

  it("returns false when no font family is set", () => {
    expect(isCodeStyledFont(undefined)).toBe(false);
  });
});
