// SPEC-024 — images.docx carries one inline picture (`wp:inline`) and one
// floating picture (`wp:anchor`/`wp:wrapSquare`), hand-built the way Word
// itself emits them (see scripts/build-spec024-fixtures.mjs) — a genuine
// test of the reader, not a round trip through this project's own writer.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "../core/zip";
import type { DocumentModel, Image } from "../documentModel";
import { asParagraph } from "../documentModel";
import { ensureMediaEncoded, mediaDataUri } from "../core/media";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

describe("parseDocx — images (AC-4..AC-9, AC-23, AC-25, AC-28)", () => {
  const fixture = loadFixture("images.docx");

  it("models the inline picture as an Image block between its surrounding text", async () => {
    const parsed = await parseDocx(fixture);
    const blocks = parsed.sections[0].blocks;
    // Paragraph 1 ("Inline picture: <img> end of paragraph.") splits into
    // three blocks around the inline image; paragraph 2 opens directly with
    // the floating image (no leading text run), so it contributes no empty
    // leading paragraph — just the image and the trailing text.
    expect(blocks.map((b) => b.type)).toEqual(["paragraph", "image", "paragraph", "image", "paragraph"]);
    expect(asParagraph(blocks[0]).runs.map((r) => r.text)).toEqual(["Inline picture: "]);
    expect(asParagraph(blocks[2]).runs.map((r) => r.text)).toEqual([" end of paragraph."]);
    expect(asParagraph(blocks[4]).runs.map((r) => r.text)).toEqual(["Text beside the floating picture."]);
  });

  it("reads the inline picture's wrap mode, size and alt text (AC-13, AC-23, AC-25, AC-28)", async () => {
    const parsed = await parseDocx(fixture);
    const image = parsed.sections[0].blocks[1] as Image;
    expect(image.type).toBe("image");
    expect(image.wrap).toBe("inline");
    // 300x200px at 96 DPI = 7.9375cm x 5.2917cm — the size Word itself shows.
    expect(image.widthCm).toBeCloseTo(7.9375, 3);
    expect(image.heightCm).toBeCloseTo(5.2917, 3);
    expect(image.naturalWidthPx).toBe(300);
    expect(image.naturalHeightPx).toBe(200);
    expect(image.altText).toBe("A blue landscape photo");
  });

  it("reads the floating picture's square wrap mode and size (AC-13, AC-14, AC-28)", async () => {
    const parsed = await parseDocx(fixture);
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
    const parsed = await parseDocx(fixture);
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
    const parsed = await parseDocx(fixture);
    const names = parsed.opaque.map((part) => part.name);
    expect(names).not.toContain("word/media/image1.png");
    expect(names).not.toContain("word/media/image2.png");
  });
});

describe("writeDocx — images same-format round trip", () => {
  const fixture = loadFixture("images.docx");

  it("re-opens with the same images after an unedited save", async () => {
    const parsed = await parseDocx(fixture);
    const model: DocumentModel = {
      id: "images-test",
      title: "Images",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "docx",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
      media: parsed.media,
    };
    const { bytes, dropped } = await writeDocx(model);
    expect(dropped).toEqual([]);
    const reopened = await parseDocx(bytes);
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
    expect(Array.from(savedParts.keys()).some((name) => /word\/media\/image\d+\.png/.test(name))).toBe(true);
  });
});

// SPEC-047: the interactive open path parses with { deferMedia: true } so the
// image base64 pass does not sit between choosing the file and first paint.
describe("parseDocx — deferred media (SPEC-047)", () => {
  const fixture = loadFixture("images.docx");

  it("leaves media parts unencoded but force-encodes on demand to the same bytes", async () => {
    const eager = await parseDocx(fixture);
    const deferred = await parseDocx(fixture, undefined, { deferMedia: true });
    expect(deferred.media).toHaveLength(2);
    // Deferred: no data yet.
    expect(deferred.media.every((part) => part.data === "")).toBe(true);
    // Force-encoding yields exactly what the eager parse produced.
    for (const part of deferred.media) {
      ensureMediaEncoded(part);
    }
    const eagerById = new Map(eager.media.map((p) => [p.mediaId, p.data]));
    for (const part of deferred.media) {
      expect(part.data).toBe(eagerById.get(part.mediaId));
    }
  });

  it("mediaDataUri renders a page-1 image without waiting for the background pass", async () => {
    const deferred = await parseDocx(fixture, undefined, { deferMedia: true });
    const inline = deferred.sections[0].blocks[1] as Image;
    const part = deferred.media.find((p) => p.mediaId === inline.mediaId)!;
    const uri = mediaDataUri(part);
    expect(uri.startsWith("data:image/png;base64,")).toBe(true);
    expect(uri.length).toBeGreaterThan("data:image/png;base64,".length);
  });

  it("a save before the background encode has run still writes every image (round-trip fidelity)", async () => {
    // Parse deferred and DO NOT flush — the writer must force-encode itself.
    const parsed = await parseDocx(fixture, undefined, { deferMedia: true });
    expect(parsed.media.every((part) => part.data === "")).toBe(true);
    const model: DocumentModel = {
      id: "images-deferred",
      title: "Images",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "docx",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
      media: parsed.media,
    };
    const { bytes, dropped } = await writeDocx(model);
    expect(dropped).toEqual([]);
    const savedParts = await unzip(bytes);
    const mediaNames = Array.from(savedParts.keys()).filter((name) => /word\/media\/image\d+\.png/.test(name));
    expect(mediaNames).toHaveLength(2);
    for (const name of mediaNames) {
      expect(savedParts.get(name)!.length).toBeGreaterThan(0);
    }
    const reopened = await parseDocx(bytes);
    expect(reopened.media).toHaveLength(2);
  });
});
