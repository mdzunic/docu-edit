// The content-type/extension mapping tests, moved verbatim from
// docx/drawing.spec.ts when the helpers moved to core/media.ts (both format
// packages share them; the AC references are SPEC-024's).
import { describe, it, expect } from "vitest";
import { CONTENT_TYPE_BY_EXTENSION, contentTypeForPartName, extensionForContentType } from "./media";

describe("contentTypeForPartName / extensionForContentType (AC-1, AC-2)", () => {
  it("maps every supported extension to its content type", () => {
    expect(contentTypeForPartName("word/media/image1.png")).toBe("image/png");
    expect(contentTypeForPartName("word/media/image2.jpg")).toBe("image/jpeg");
    expect(contentTypeForPartName("word/media/image3.jpeg")).toBe("image/jpeg");
    expect(contentTypeForPartName("word/media/image4.gif")).toBe("image/gif");
    expect(contentTypeForPartName("word/media/image5.svg")).toBe("image/svg+xml");
  });

  it("is case-insensitive on the extension", () => {
    expect(contentTypeForPartName("word/media/image1.PNG")).toBe("image/png");
  });

  it("returns undefined for an unknown extension or no extension", () => {
    expect(contentTypeForPartName("word/media/image1.bmp")).toBeUndefined();
    expect(contentTypeForPartName("word/media/image1")).toBeUndefined();
  });

  it("maps every content type CONTENT_TYPE_BY_EXTENSION covers back to its canonical extension", () => {
    expect(extensionForContentType("image/png")).toBe("png");
    expect(extensionForContentType("image/jpeg")).toBe("jpeg");
    expect(extensionForContentType("image/gif")).toBe("gif");
    expect(extensionForContentType("image/svg+xml")).toBe("svg");
  });

  it("falls back to png for an unknown content type", () => {
    expect(extensionForContentType("image/bmp")).toBe("png");
  });

  it("CONTENT_TYPE_BY_EXTENSION covers the four accepted formats", () => {
    expect(Object.keys(CONTENT_TYPE_BY_EXTENSION).sort()).toEqual(["gif", "jpeg", "jpg", "png", "svg"]);
  });
});
