import { describe, it, expect } from "vitest";
import { createDocument, DEFAULT_TITLE } from "./documentModel";

describe("createDocument", () => {
  it("creates a document with exactly one empty paragraph", () => {
    const doc = createDocument();
    expect(doc.sections).toHaveLength(1);
    expect(doc.sections[0].blocks).toHaveLength(1);
    expect(doc.sections[0].blocks[0]).toEqual({ type: "paragraph", runs: [] });
  });

  it("defaults the title to 'Untitled document'", () => {
    const doc = createDocument();
    expect(doc.title).toBe(DEFAULT_TITLE);
  });

  it("initialises defaultStyles to A4, 2.54cm margins and a body font", () => {
    const doc = createDocument();
    expect(doc.defaultStyles.pageSize).toEqual({ widthCm: 21, heightCm: 29.7 });
    expect(doc.defaultStyles.margins).toEqual({
      topCm: 2.54,
      rightCm: 2.54,
      bottomCm: 2.54,
      leftCm: 2.54,
    });
    expect(doc.defaultStyles.fontFamily).toBeTruthy();
    expect(doc.defaultStyles.fontSizePt).toBeGreaterThan(0);
  });

  it("generates a distinct id per document when none is given", () => {
    const a = createDocument();
    const b = createDocument();
    expect(a.id).not.toBe(b.id);
  });

  it("uses a given id instead of generating one", () => {
    const doc = createDocument("fixed-id");
    expect(doc.id).toBe("fixed-id");
  });

  it("reads and writes the title", () => {
    const doc = createDocument();
    doc.title = "Quarterly Report";
    expect(doc.title).toBe("Quarterly Report");
  });

  it("serialises to JSON and back without loss", () => {
    const doc = createDocument("round-trip-id");
    doc.title = "My Document";
    doc.sections[0].blocks[0] = {
      type: "paragraph",
      runs: [{ text: "hello", formatting: { bold: true } }],
    };

    const restored = JSON.parse(JSON.stringify(doc));

    expect(restored).toEqual(doc);
  });
});
