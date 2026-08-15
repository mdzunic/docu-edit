// SPEC-024 — images.odt carries one inline picture (`text:anchor-type=
// "as-char"`) and one floating picture (`text:anchor-type="paragraph"` +
// a graphic style with `style:wrap="parallel"`), hand-built the way
// LibreOffice itself emits them (see scripts/build-spec024-fixtures.mjs) —
// a genuine test of the reader, not a round trip through this project's
// own writer.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../core/zip";
import type { DocumentModel, Image } from "../documentModel";
import { asParagraph } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

describe("parseOdt — images (AC-4..AC-9, AC-24, AC-25, AC-29)", () => {
  const fixture = loadFixture("images.odt");

  it("models the inline picture as an Image block between its surrounding text", async () => {
    const parsed = await parseOdt(fixture);
    const blocks = parsed.sections[0].blocks;
    // Paragraph 1 ("Inline picture: <img> end of paragraph.") splits into
    // three blocks around the inline image; paragraph 2 opens directly with
    // the floating image, so it contributes no empty leading paragraph.
    expect(blocks.map((b) => b.type)).toEqual(["paragraph", "image", "paragraph", "image", "paragraph"]);
    expect(asParagraph(blocks[0]).runs.map((r) => r.text)).toEqual(["Inline picture: "]);
    expect(asParagraph(blocks[2]).runs.map((r) => r.text)).toEqual(["end of paragraph."]);
    expect(asParagraph(blocks[4]).runs.map((r) => r.text)).toEqual(["Text beside the floating picture."]);
  });

  it("reads the inline picture's wrap mode, size and alt text (AC-13, AC-24, AC-25, AC-29)", async () => {
    const parsed = await parseOdt(fixture);
    const image = parsed.sections[0].blocks[1] as Image;
    expect(image.type).toBe("image");
    expect(image.wrap).toBe("inline");
    expect(image.widthCm).toBeCloseTo(7.9375, 3);
    expect(image.heightCm).toBeCloseTo(5.2917, 3);
    expect(image.naturalWidthPx).toBe(300);
    expect(image.naturalHeightPx).toBe(200);
    expect(image.altText).toBe("A blue landscape photo");
  });

  it("reads the floating picture's square wrap mode and size (AC-13, AC-14, AC-29)", async () => {
    const parsed = await parseOdt(fixture);
    const image = parsed.sections[0].blocks[3] as Image;
    expect(image.type).toBe("image");
    expect(image.wrap).toBe("square");
    expect(image.widthCm).toBeCloseTo(3.175, 3);
    expect(image.heightCm).toBeCloseTo(3.175, 3);
    expect(image.naturalWidthPx).toBe(120);
    expect(image.naturalHeightPx).toBe(120);
    expect(image.altText).toBe("A red square logo");
  });

  it("models the two pictures' media as distinct PNG parts (AC-26)", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.media).toHaveLength(2);
    for (const part of parsed.media) {
      expect(part.contentType).toBe("image/png");
      expect(part.data.length).toBeGreaterThan(0);
    }
    const inline = parsed.sections[0].blocks[1] as Image;
    const floating = parsed.sections[0].blocks[3] as Image;
    expect(inline.mediaId).not.toBe(floating.mediaId);
    expect(parsed.media.map((m) => m.mediaId)).toEqual(
      expect.arrayContaining([inline.mediaId, floating.mediaId]),
    );
  });

  it("does not retain the media parts as opaque (SPEC-008 excludes modelled images)", async () => {
    const parsed = await parseOdt(fixture);
    const names = parsed.opaque.map((part) => part.name);
    expect(names).not.toContain("Pictures/image1.png");
    expect(names).not.toContain("Pictures/image2.png");
  });
});

describe("writeOdt — images same-format round trip", () => {
  const fixture = loadFixture("images.odt");

  it("re-opens with the same images after an unedited save", async () => {
    const parsed = await parseOdt(fixture);
    const model: DocumentModel = {
      id: "images-test",
      title: "Images",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "odt",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
      media: parsed.media,
    };
    const { bytes, dropped } = await writeOdt(model);
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    const inline = reopened.sections[0].blocks[1] as Image;
    const floating = reopened.sections[0].blocks[3] as Image;
    expect(inline.wrap).toBe("inline");
    expect(inline.widthCm).toBeCloseTo(7.9375, 3);
    expect(inline.altText).toBe("A blue landscape photo");
    expect(floating.wrap).toBe("square");
    expect(floating.widthCm).toBeCloseTo(3.175, 3);
    expect(floating.altText).toBe("A red square logo");
    expect(reopened.media).toHaveLength(2);

    const savedParts = await unzip(bytes);
    expect(Array.from(savedParts.keys()).some((name) => /Pictures\/image\d+\.png/.test(name))).toBe(true);
  });
});
