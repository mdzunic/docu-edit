// @vitest-environment node
// SPEC-106 — generation is DOM-free and fast. This file runs in vitest's
// plain Node environment (no jsdom): `document`, `window`, `DOMParser` and
// any canvas are simply absent, so a module on the deck path that touched
// one would throw here. presentation.ts and slideMarkdown.ts are imported
// and run with no polyfill — `ensureNodeXmlGlobals` belongs to the READ
// path (readDeck.ts) alone, which this file never calls. And SPEC-045's
// performance bar, the shape performance.spec.ts has: a 100-slide deck of
// title + six bullets builds and writes in under 2 seconds, one spec per
// format.
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import { PresentationBuilder } from "./presentation";
import { parseSlideMarkdownWithReport, writeSlideMarkdown } from "./slideMarkdown";

const PNG_4X2 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";

describe("the deck path is DOM-free (AC-22)", () => {
  it("runs with no document, window, DOMParser or canvas in scope", () => {
    expect(typeof document).toBe("undefined");
    expect(typeof window).toBe("undefined");
    expect(typeof DOMParser).toBe("undefined");
    expect(typeof HTMLCanvasElement).toBe("undefined");
    expect(typeof OffscreenCanvas).toBe("undefined");
  });

  it("builds, writes both formats and exports HTML through presentation.ts with no polyfill", async () => {
    const builder = new PresentationBuilder({ title: "Node" })
      .addTitleSlide("Node", "No browser")
      .addBulletSlide("Rows", ["one", ["two"]])
      .setNotes(["A note."])
      .addImageSlide("Picture", { data: PNG_4X2, contentType: "image/png" })
      .setBackground("#202020");
    for (const format of ["pptx", "odp"] as const) {
      const parts = await unzip(await builder.toBytes(format));
      expect(parts.size).toBeGreaterThan(4);
    }
    expect(Buffer.isBuffer(await builder.toBuffer())).toBe(true);
    expect(builder.toHtml({ includeNotes: true })).toContain("A note.");
    expect(typeof DOMParser).toBe("undefined");
  });

  it("parses and writes slide Markdown through slideMarkdown.ts with no polyfill", async () => {
    const { deck, dropped } = parseSlideMarkdownWithReport(
      `# Node\nNo browser\n\n---\n\n## Rows\n- **one**\n  - two\n\n![](data:image/png;base64,${PNG_4X2})\n\n> dropped\n\nNotes: Say it.`,
    );
    expect(deck.slides).toHaveLength(2);
    expect(dropped.map((record) => record.kind)).toEqual(["blockQuote"]);
    expect(writeSlideMarkdown(deck).text).toContain("- **one**\n  - two");
    expect(typeof DOMParser).toBe("undefined");
  });
});

/** 100 slides, each a title + six bullets. */
function hundredSlides(format: "pptx" | "odp"): PresentationBuilder {
  const builder = new PresentationBuilder({ title: "Performance", format });
  for (let slide = 0; slide < 100; slide++) {
    builder.addBulletSlide(
      `Slide ${slide + 1}`,
      Array.from({ length: 6 }, (_, row) => `Bullet ${row + 1} of slide ${slide + 1}`),
    );
  }
  return builder;
}

describe("performance (AC-23)", () => {
  for (const format of ["pptx", "odp"] as const) {
    it(`builds and writes a 100-slide, six-bullet .${format} in under 2 seconds`, async () => {
      const start = performance.now();
      const bytes = await hundredSlides(format).toBytes();
      const elapsedMs = performance.now() - start;

      expect(bytes.length).toBeGreaterThan(0);
      const parts = await unzip(bytes);
      if (format === "pptx") {
        expect(parts.has("ppt/slides/slide100.xml")).toBe(true);
      } else {
        expect(new TextDecoder().decode(parts.get("content.xml")).match(/<draw:page /g)).toHaveLength(100);
      }
      expect(elapsedMs).toBeLessThan(2000);
    });
  }
});
