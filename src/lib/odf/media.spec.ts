import { describe, it, expect } from "vitest";
import { planOdtMediaWrite } from "./media";
import { base64ToBytes } from "../opaque";
import { createDocument, type DocumentModel, type Image, type Paragraph } from "../documentModel";

function paragraph(text = ""): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function image(mediaId: string): Image {
  return {
    type: "image",
    mediaId,
    naturalWidthPx: 100,
    naturalHeightPx: 100,
    widthCm: 2,
    heightCm: 2,
    wrap: "inline",
  };
}

function docWithImages(blocks: Array<Paragraph | Image>, media: DocumentModel["media"]): DocumentModel {
  const model = createDocument("odt-media-fixture");
  model.sections = [{ blocks }];
  model.media = media;
  return model;
}

describe("planOdtMediaWrite — no images", () => {
  it("returns an empty plan when no block references an image", () => {
    const model = docWithImages([paragraph("just text")], []);
    const plan = planOdtMediaWrite(model);
    expect(plan.zipEntries).toEqual([]);
    expect(plan.manifestEntries).toEqual([]);
    expect(plan.hrefFor("anything")).toBe("");
  });
});

describe("planOdtMediaWrite — referenced images", () => {
  it("plans one image: Pictures/imageN.ext path and manifest entry", () => {
    const model = docWithImages([image("m1"), paragraph()], [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }]);
    const plan = planOdtMediaWrite(model);
    expect(plan.zipEntries).toHaveLength(1);
    expect(plan.zipEntries[0].name).toBe("Pictures/image1.png");
    expect(plan.zipEntries[0].data).toEqual(base64ToBytes("AAAA"));
    expect(plan.manifestEntries).toEqual([{ path: "Pictures/image1.png", mediaType: "image/png" }]);
    expect(plan.hrefFor("m1")).toBe("Pictures/image1.png");
  });

  it("plans two images with sequential Pictures/ paths", () => {
    const model = docWithImages(
      [image("m1"), image("m2"), paragraph()],
      [
        { mediaId: "m1", contentType: "image/png", data: "AAAA" },
        { mediaId: "m2", contentType: "image/jpeg", data: "BBBB" },
      ],
    );
    const plan = planOdtMediaWrite(model);
    expect(plan.zipEntries.map((e) => e.name)).toEqual(["Pictures/image1.png", "Pictures/image2.jpeg"]);
    expect(plan.hrefFor("m1")).toBe("Pictures/image1.png");
    expect(plan.hrefFor("m2")).toBe("Pictures/image2.jpeg");
    expect(plan.manifestEntries).toEqual([
      { path: "Pictures/image1.png", mediaType: "image/png" },
      { path: "Pictures/image2.jpeg", mediaType: "image/jpeg" },
    ]);
  });

  it("hrefFor returns empty string for an unknown mediaId", () => {
    const model = docWithImages([image("m1"), paragraph()], [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }]);
    const plan = planOdtMediaWrite(model);
    expect(plan.hrefFor("does-not-exist")).toBe("");
  });

  it("includes an image/svg+xml media part normally — ODF has no SVG restriction", () => {
    const model = docWithImages(
      [image("m-svg"), paragraph()],
      [{ mediaId: "m-svg", contentType: "image/svg+xml", data: "AAAA" }],
    );
    const plan = planOdtMediaWrite(model);
    expect(plan.zipEntries).toHaveLength(1);
    expect(plan.zipEntries[0].name).toBe("Pictures/image1.svg");
    expect(plan.manifestEntries).toEqual([{ path: "Pictures/image1.svg", mediaType: "image/svg+xml" }]);
    expect(plan.hrefFor("m-svg")).toBe("Pictures/image1.svg");
  });

  it("seeds part naming from existingPartNames to avoid colliding with an already-used Pictures/ path", () => {
    const model = docWithImages([image("m1"), paragraph()], [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }]);
    const plan = planOdtMediaWrite(model, new Set(["Pictures/image1.png"]));
    expect(plan.zipEntries[0].name).toBe("Pictures/image2.png");
  });

  it("dedups a mediaId referenced by two Image blocks into a single zip entry/manifest entry", () => {
    const model = docWithImages(
      [image("shared"), image("shared"), paragraph()],
      [{ mediaId: "shared", contentType: "image/png", data: "AAAA" }],
    );
    const plan = planOdtMediaWrite(model);
    expect(plan.zipEntries).toHaveLength(1);
    expect(plan.manifestEntries).toHaveLength(1);
    expect(plan.hrefFor("shared")).toBe("Pictures/image1.png");
  });

  it("skips a referenced mediaId that has no matching model.media entry", () => {
    const model = docWithImages([image("missing"), paragraph()], []);
    const plan = planOdtMediaWrite(model);
    expect(plan.zipEntries).toEqual([]);
    expect(plan.hrefFor("missing")).toBe("");
  });
});
