import { describe, expect, it } from "vitest";
import { parseCssColor } from "./domExtract";
import { allFaceKeys, classifyFamily, faceFileName, faceKeyFor, firstFamilyName } from "./fontMap";

describe("firstFamilyName", () => {
  it("takes the first family, unquoted and lowercased", () => {
    expect(firstFamilyName('"Times New Roman", serif')).toBe("times new roman");
    expect(firstFamilyName("Calibri, Carlito, sans-serif")).toBe("calibri");
    expect(firstFamilyName("'Courier New'")).toBe("courier new");
  });
});

describe("classifyFamily", () => {
  it("maps common document fonts to their Liberation family", () => {
    expect(classifyFamily("Calibri, sans-serif")).toBe("sans");
    expect(classifyFamily("Arial")).toBe("sans");
    expect(classifyFamily("Helvetica, Arial, sans-serif")).toBe("sans");
    expect(classifyFamily('"Times New Roman", Times, serif')).toBe("serif");
    expect(classifyFamily("Georgia")).toBe("serif");
    expect(classifyFamily("Cambria")).toBe("serif");
    expect(classifyFamily('"Courier New", monospace')).toBe("mono");
    expect(classifyFamily("Consolas")).toBe("mono");
    expect(classifyFamily("Liberation Serif")).toBe("serif");
    expect(classifyFamily("Liberation Mono")).toBe("mono");
  });

  it("maps generic keywords", () => {
    expect(classifyFamily("serif")).toBe("serif");
    expect(classifyFamily("monospace")).toBe("mono");
    expect(classifyFamily("sans-serif")).toBe("sans");
    expect(classifyFamily("system-ui, sans-serif")).toBe("sans");
  });
});

describe("faceKeyFor", () => {
  it("combines family classification with weight and slant", () => {
    expect(faceKeyFor("Arial", false, false)).toBe("sans-regular");
    expect(faceKeyFor("Arial", true, false)).toBe("sans-bold");
    expect(faceKeyFor("Georgia", false, true)).toBe("serif-italic");
    expect(faceKeyFor("Courier New", true, true)).toBe("mono-bolditalic");
  });
});

describe("faceFileName", () => {
  it("names a bundled Liberation file for every face key", () => {
    for (const key of allFaceKeys()) {
      expect(faceFileName(key)).toMatch(/^Liberation(Sans|Serif|Mono)-\w+\.ttf$/);
    }
    expect(faceFileName("serif-bolditalic")).toBe("LiberationSerif-BoldItalic.ttf");
  });
});

describe("parseCssColor", () => {
  it("parses computed rgb()/rgba() strings into 0..1 channels", () => {
    expect(parseCssColor("rgb(255, 0, 0)")).toEqual([1, 0, 0]);
    expect(parseCssColor("rgba(0, 128, 255, 0.5)")).toEqual([0, 128 / 255, 1]);
    expect(parseCssColor("rgb(32 33 36)")).toEqual([32 / 255, 33 / 255, 36 / 255]);
  });

  it("falls back to black for unparseable colors", () => {
    expect(parseCssColor("currentcolor")).toEqual([0, 0, 0]);
    expect(parseCssColor("")).toEqual([0, 0, 0]);
  });
});
