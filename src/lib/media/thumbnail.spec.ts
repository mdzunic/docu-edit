// The PDF card placeholder's figures (SPEC-107 AC-29): the paper is drawn at
// the first page's real aspect ratio — rotation applied — and the card sizes
// the paper from that ratio, so a degenerate page must still yield a usable
// one. The rendered geometry itself is e2e's (e2e/SPEC-107.spec.ts).
import { describe, expect, it } from "vitest";
import { pdfCardPlaceholder, pdfPageLines, pdfPageSizesLabel } from "./thumbnail";
import type { PdfSummary } from "./model";

function summary(pages: PdfSummary["pages"]): PdfSummary {
  return { version: "1.7", pageCount: pages.length, pages, linearised: false, reconstructed: false };
}

describe("pdfCardPlaceholder", () => {
  it("is the first page's own aspect ratio, page count and size", () => {
    const placeholder = pdfCardPlaceholder(
      summary([
        { widthPt: 612, heightPt: 792, rotate: 0 },
        { widthPt: 842, heightPt: 595, rotate: 0 },
      ]),
    );
    expect(placeholder.aspectRatio).toBeCloseTo(612 / 792);
    expect(placeholder.pageCount).toBe(2);
    expect(placeholder.pageSize).toBe("612 × 792 pt");
  });

  it("applies the first page's rotation, so a quarter-turned portrait page reads landscape", () => {
    expect(pdfCardPlaceholder(summary([{ widthPt: 595, heightPt: 842, rotate: 90 }])).aspectRatio).toBeCloseTo(842 / 595);
    expect(pdfCardPlaceholder(summary([{ widthPt: 595, heightPt: 842, rotate: 270 }])).aspectRatio).toBeCloseTo(842 / 595);
    expect(pdfCardPlaceholder(summary([{ widthPt: 595, heightPt: 842, rotate: 180 }])).aspectRatio).toBeCloseTo(595 / 842);
  });

  it("draws a zero-area first page square rather than as a zero-width paper", () => {
    expect(pdfCardPlaceholder(summary([{ widthPt: 0, heightPt: 792, rotate: 0 }])).aspectRatio).toBe(1);
    expect(pdfCardPlaceholder(summary([{ widthPt: 612, heightPt: 0, rotate: 0 }])).aspectRatio).toBe(1);
  });

  it("falls back to US Letter when the summary lists no page", () => {
    const placeholder = pdfCardPlaceholder(summary([]));
    expect(placeholder.aspectRatio).toBeCloseTo(612 / 792);
    expect(placeholder.pageSize).toBe("612 × 792 pt");
  });
});

describe("pdfPageSizesLabel (SPEC-108: the size in the user's unit too)", () => {
  const a4 = { widthPt: 595.28, heightPt: 841.89, rotate: 0 as const };
  const letter = { widthPt: 612, heightPt: 792, rotate: 0 as const };

  it("reads in points alone without a unit, as before", () => {
    expect(pdfPageSizesLabel(summary([a4]))).toBe("595 × 842 pt");
  });

  it("adds the size in the measurement unit", () => {
    expect(pdfPageSizesLabel(summary([a4]), "cm")).toBe("595 × 842 pt · 21 × 29.7 cm");
    expect(pdfPageSizesLabel(summary([letter]), "in")).toBe("612 × 792 pt · 8.5 × 11 in");
    expect(pdfPageSizesLabel(summary([letter, a4]), "mm")).toBe("mixed (first 612 × 792 pt · 215.9 × 279.4 mm)");
    expect(pdfPageSizesLabel(summary([a4]), "pt")).toBe("595 × 842 pt");
  });
});

describe("pdfPageLines (SPEC-111: each page's size and rotation)", () => {
  it("lists every page, with its size in the unit beside points and its rotation when it has one", () => {
    const pages = summary([
      { widthPt: 720, heightPt: 540, rotate: 0 },
      { widthPt: 595.28, heightPt: 841.89, rotate: 90 },
    ]);
    expect(pdfPageLines(pages, "cm")).toEqual(["Page 1 · 720 × 540 pt · 25.4 × 19.05 cm", "Page 2 · 595 × 842 pt · 21 × 29.7 cm · rotate 90"]);
    expect(pdfPageLines(pages)).toEqual(["Page 1 · 720 × 540 pt", "Page 2 · 595 × 842 pt · rotate 90"]);
    expect(pdfPageLines(pages, "pt")).toEqual(["Page 1 · 720 × 540 pt", "Page 2 · 595 × 842 pt · rotate 90"]);
  });
});
