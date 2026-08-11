// SPEC-045 AC-23, AC-24: generating a 50-page document (50 page-break-
// separated sections, each with one heading and 20 paragraphs of 5
// formatted runs) completes in under 2 seconds in Node.
import { describe, expect, it } from "vitest";
import { DocumentBuilder } from "./builder";

describe("performance (AC-23, AC-24)", () => {
  it("generates a 50-section, 1000-paragraph .docx in under 2 seconds", async () => {
    const builder = new DocumentBuilder({ title: "Performance" });
    for (let section = 0; section < 50; section++) {
      builder.addHeading(`Section ${section + 1}`, 1);
      for (let p = 0; p < 20; p++) {
        builder.addParagraph(
          Array.from({ length: 5 }, (_, i) => ({
            text: `run ${i} of paragraph ${p} `,
            bold: i % 2 === 0,
            italic: i % 3 === 0,
          })),
        );
      }
      if (section < 49) {
        builder.addPageBreak();
      }
    }

    const start = performance.now();
    const bytes = await builder.toBytes("docx");
    const elapsedMs = performance.now() - start;

    expect(bytes.length).toBeGreaterThan(0);
    expect(elapsedMs).toBeLessThan(2000);
  });
});
