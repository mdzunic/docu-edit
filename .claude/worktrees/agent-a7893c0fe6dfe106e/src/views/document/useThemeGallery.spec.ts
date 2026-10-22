// SPEC-077 AC-11: the gallery card miniature's source — the first page's
// leading paragraphs as plain text, or [] (→ specimen fallback) for a
// document with nothing laid out yet.
import { describe, expect, it } from "vitest";
import { createDocument, type Paragraph } from "../../lib/documentModel";
import { themePreviewParagraphs } from "./useThemeGallery";

function paragraph(text: string, headingLevel?: number): Paragraph {
  const block: Paragraph = { type: "paragraph", runs: [{ text, formatting: {} }] };
  if (headingLevel !== undefined) {
    block.headingLevel = headingLevel;
  }
  return block;
}

describe("themePreviewParagraphs (SPEC-077 AC-11)", () => {
  it("returns [] for a null or blank document (specimen fallback)", () => {
    expect(themePreviewParagraphs(null)).toEqual([]);
    // A fresh document holds one empty paragraph — nothing to miniaturise.
    expect(themePreviewParagraphs(createDocument("blank"))).toEqual([]);
  });

  it("returns the leading paragraphs with their heading flag", () => {
    const model = createDocument("doc");
    model.sections = [
      { blocks: [paragraph("Annual report", 1), paragraph(""), paragraph("Body line one"), paragraph("Body line two")] },
    ];
    expect(themePreviewParagraphs(model)).toEqual([
      { text: "Annual report", heading: true },
      { text: "Body line one", heading: false },
      { text: "Body line two", heading: false },
    ]);
  });

  it("caps the miniature at six lines and 120 characters per line", () => {
    const model = createDocument("doc");
    const long = "x".repeat(500);
    model.sections = [{ blocks: Array.from({ length: 10 }, (_, i) => paragraph(`${i} ${long}`)) }];
    const preview = themePreviewParagraphs(model);
    expect(preview).toHaveLength(6);
    for (const line of preview) {
      expect(line.text.length).toBeLessThanOrEqual(120);
    }
  });
});
