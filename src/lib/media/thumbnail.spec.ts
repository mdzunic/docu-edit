// The PDF card placeholder's figures (SPEC-107 AC-29): the paper is drawn at
// the first page's real aspect ratio — rotation applied — and the card sizes
// the paper from that ratio, so a degenerate page must still yield a usable
// one. The rendered geometry itself is e2e's (e2e/SPEC-107.spec.ts).
import { describe, expect, it } from "vitest";
import { pdfCardPlaceholder } from "./thumbnail";
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
