import { describe, it, expect, beforeEach } from "vitest";
import { createDocument, type Paragraph } from "./documentModel";
import {
  buildThumbnail,
  clearThumbnailCache,
  getCachedThumbnail,
  renderThumbnail,
  THUMBNAIL_MAX_LINES,
} from "./libraryThumbnail";

function paragraph(text: string, headingLevel?: number): Paragraph {
  return {
    type: "paragraph",
    runs: [{ text, formatting: {} }],
    ...(headingLevel === undefined ? {} : { headingLevel }),
  };
}

beforeEach(() => {
  clearThumbnailCache();
});

describe("buildThumbnail (SPEC-067)", () => {
  it("extracts the leading paragraphs with their heading levels", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [
      paragraph("The Title", 1),
      paragraph("Body one"),
      paragraph("Body two"),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.failed).toBe(false);
    expect(thumb.lines).toEqual([
      { text: "The Title", level: 1 },
      { text: "Body one", level: 0 },
      { text: "Body two", level: 0 },
    ]);
  });

  it("skips empty paragraphs and caps the line count", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [
      paragraph("   "),
      ...Array.from({ length: THUMBNAIL_MAX_LINES + 5 }, (_unused, i) =>
        paragraph(`line ${i}`),
      ),
    ];
    const thumb = buildThumbnail(model);
    expect(thumb.lines).toHaveLength(THUMBNAIL_MAX_LINES);
    expect(thumb.lines[0]).toEqual({ text: "line 0", level: 0 });
  });

  it("never throws on a malformed model — it reports failure instead", () => {
    const broken = { sections: [{ blocks: [{ type: "paragraph" }] }] } as never;
    expect(() => buildThumbnail(broken)).not.toThrow();
  });
});

describe("thumbnail cache (SPEC-067)", () => {
  it("caches per id + modified time and re-renders when the timestamp changes", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [paragraph("first")];

    const first = renderThumbnail("doc", 1000, model);
    expect(getCachedThumbnail("doc", 1000)).toBe(first);

    // Same id + timestamp returns the cached instance.
    expect(renderThumbnail("doc", 1000, model)).toBe(first);

    // A newer edit invalidates the old cache entry.
    expect(getCachedThumbnail("doc", 2000)).toBeUndefined();
    model.sections[0].blocks = [paragraph("edited")];
    const second = renderThumbnail("doc", 2000, model);
    expect(second).not.toBe(first);
    expect(second.lines[0]?.text).toBe("edited");
    expect(getCachedThumbnail("doc", 1000)).toBeUndefined();
  });
});
