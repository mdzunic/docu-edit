import { describe, expect, it } from "vitest";
import { isImageFormat, libraryRowMenuItems } from "./libraryMenu";

describe("libraryRowMenuItems", () => {
  it("offers Convert to PDF… on image entries only (SPEC-108)", () => {
    for (const format of ["png", "jpeg", "gif", "svg", "webp"]) {
      expect(libraryRowMenuItems(false, format).map((item) => item.testid), format).toEqual([
        "menu-star",
        "menu-rename",
        "menu-duplicate",
        "menu-convert-pdf",
        "menu-delete",
      ]);
    }
    for (const format of ["pdf", "docx", "pptx", "xlsx", undefined]) {
      expect(libraryRowMenuItems(true, format).map((item) => item.id), String(format)).not.toContain("convert-pdf");
    }
    expect(libraryRowMenuItems(false, "png")[3].label).toBe("Convert to PDF…");
  });

  it("isImageFormat", () => {
    expect(isImageFormat("webp")).toBe(true);
    expect(isImageFormat("pdf")).toBe(false);
    expect(isImageFormat(undefined)).toBe(false);
  });
});
