import { describe, expect, it } from "vitest";
import {
  BUNDLED_FAMILIES,
  DOCUMENTED_DEFAULT_FAMILY,
  fontStack,
  isBundledFamily,
  substituteFor,
} from "./substitution";

describe("substituteFor", () => {
  it("maps the common Microsoft fonts to their bundled metric twins (AC-1)", () => {
    expect(substituteFor("Calibri")?.family).toBe("Carlito");
    expect(substituteFor("Cambria")?.family).toBe("Caladea");
    expect(substituteFor("Arial")?.family).toBe("Liberation Sans");
    expect(substituteFor("Times New Roman")?.family).toBe("Liberation Serif");
    expect(substituteFor("Courier New")?.family).toBe("Liberation Mono");
  });

  it("matches case-insensitively and ignores a trailing CSS stack", () => {
    expect(substituteFor("calibri")?.family).toBe("Carlito");
    expect(substituteFor('"Cambria", serif')?.family).toBe("Caladea");
  });

  it("has no twin for an unknown family", () => {
    expect(substituteFor("Comic Sans MS")).toBeUndefined();
    expect(substituteFor("")).toBeUndefined();
  });
});

describe("isBundledFamily", () => {
  it("recognises the fonts we ship ourselves", () => {
    for (const family of BUNDLED_FAMILIES) {
      expect(isBundledFamily(family)).toBe(true);
    }
    expect(isBundledFamily("carlito")).toBe(true);
    expect(isBundledFamily("Calibri")).toBe(false);
  });
});

describe("fontStack", () => {
  it("puts the original first so an installed original still wins (AC-2)", () => {
    expect(fontStack("Calibri")).toBe("Calibri, Carlito, sans-serif");
    expect(fontStack("Cambria")).toBe("Cambria, Caladea, serif");
    expect(fontStack("Arial")).toBe("Arial, Liberation Sans, sans-serif");
  });

  it("quotes multi-word original names", () => {
    expect(fontStack("Times New Roman")).toBe('"Times New Roman", Liberation Serif, serif');
    expect(fontStack("Courier New")).toBe('"Courier New", Liberation Mono, monospace');
  });

  it("falls back a genuinely missing font to the documented default (AC-9)", () => {
    // No twin exists: the documented default backs it, then the family's own
    // generic. A missing serif still reads serif after the default.
    expect(fontStack("Garamond")).toBe(`Garamond, ${DOCUMENTED_DEFAULT_FAMILY}, serif`);
    expect(fontStack("Papyrus")).toBe(`Papyrus, ${DOCUMENTED_DEFAULT_FAMILY}, sans-serif`);
    expect(fontStack("Consolas")).toBe(`Consolas, ${DOCUMENTED_DEFAULT_FAMILY}, monospace`);
  });

  it("renders a bundled family directly, never substituting it", () => {
    expect(fontStack("Carlito")).toBe("Carlito, sans-serif");
    expect(fontStack("Liberation Mono")).toBe('"Liberation Mono", monospace');
  });

  it("leaves an unset family to the inherited base", () => {
    expect(fontStack(undefined)).toBe("");
    expect(fontStack("")).toBe("");
    expect(fontStack("   ")).toBe("");
  });
});
