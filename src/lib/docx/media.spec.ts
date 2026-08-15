import { describe, it, expect } from "vitest";
import { nextFreeRelId, planMediaWrite, takenRelIdsOf, IMAGE_RELATIONSHIP } from "./media";
import { base64ToBytes } from "../opaque";
import { createDocument, type DocumentModel, type Image, type Paragraph } from "../documentModel";
import type { Relationship } from "../core/xml";

function paragraph(text = ""): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function image(mediaId: string, wrap: Image["wrap"] = "inline"): Image {
  return {
    type: "image",
    mediaId,
    naturalWidthPx: 100,
    naturalHeightPx: 100,
    widthCm: 2,
    heightCm: 2,
    wrap,
  };
}

function docWithImages(blocks: Array<Paragraph | Image>, media: DocumentModel["media"]): DocumentModel {
  const model = createDocument("media-fixture");
  model.sections = [{ blocks }];
  model.media = media;
  return model;
}

describe("nextFreeRelId", () => {
  it("starts at rId1 by default", () => {
    expect(nextFreeRelId(new Set())).toBe("rId1");
  });

  it("skips ids already taken", () => {
    expect(nextFreeRelId(new Set(["rId1", "rId2"]))).toBe("rId3");
  });

  it("finds a gap even if it isn't contiguous", () => {
    expect(nextFreeRelId(new Set(["rId1", "rId3"]))).toBe("rId2");
  });

  it("honours a custom start", () => {
    expect(nextFreeRelId(new Set(), 5)).toBe("rId5");
    expect(nextFreeRelId(new Set(["rId5"]), 5)).toBe("rId6");
  });
});

describe("takenRelIdsOf", () => {
  it("collects relationship ids into a set", () => {
    const rels: Relationship[] = [
      { id: "rId1", type: "t", target: "a" },
      { id: "rId2", type: "t", target: "b" },
    ];
    expect(takenRelIdsOf(rels)).toEqual(new Set(["rId1", "rId2"]));
  });

  it("is empty for no relationships", () => {
    expect(takenRelIdsOf([])).toEqual(new Set());
  });
});

describe("planMediaWrite — no images", () => {
  it("returns an empty plan when no block references an image", () => {
    const model = docWithImages([paragraph("just text")], []);
    const plan = planMediaWrite(model);
    expect(plan.zipEntries).toEqual([]);
    expect(plan.relationshipsXml).toBe("");
    expect(plan.extensions).toEqual([]);
    expect(plan.relIdFor("anything")).toBe("");
  });
});

describe("planMediaWrite — referenced images", () => {
  it("plans one image: sequential part name, minted rId and extension", () => {
    const model = docWithImages(
      [image("m1"), paragraph()],
      [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }],
    );
    const plan = planMediaWrite(model);
    expect(plan.zipEntries).toHaveLength(1);
    expect(plan.zipEntries[0].name).toBe("word/media/image1.png");
    expect(plan.zipEntries[0].data).toEqual(base64ToBytes("AAAA"));
    expect(plan.extensions).toEqual(["png"]);
    const relId = plan.relIdFor("m1");
    expect(relId).toBe("rId1");
    expect(plan.relationshipsXml).toBe(
      `<Relationship Id="rId1" Type="${IMAGE_RELATIONSHIP}" Target="media/image1.png"/>`,
    );
  });

  it("plans two images with sequential part names and rIds", () => {
    const model = docWithImages(
      [image("m1"), image("m2"), paragraph()],
      [
        { mediaId: "m1", contentType: "image/png", data: "AAAA" },
        { mediaId: "m2", contentType: "image/jpeg", data: "BBBB" },
      ],
    );
    const plan = planMediaWrite(model);
    expect(plan.zipEntries.map((e) => e.name)).toEqual(["word/media/image1.png", "word/media/image2.jpeg"]);
    expect(plan.relIdFor("m1")).toBe("rId1");
    expect(plan.relIdFor("m2")).toBe("rId2");
    expect(plan.extensions.sort()).toEqual(["jpeg", "png"]);
  });

  it("relIdFor returns empty string for an unknown mediaId", () => {
    const model = docWithImages([image("m1"), paragraph()], [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }]);
    const plan = planMediaWrite(model);
    expect(plan.relIdFor("does-not-exist")).toBe("");
  });

  it("skips an image/svg+xml media part (.docx cannot store SVG) — no zip entry, no relationship, relIdFor empty", () => {
    const model = docWithImages(
      [image("m-svg"), paragraph()],
      [{ mediaId: "m-svg", contentType: "image/svg+xml", data: "AAAA" }],
    );
    const plan = planMediaWrite(model);
    expect(plan.zipEntries).toEqual([]);
    expect(plan.relationshipsXml).toBe("");
    expect(plan.relIdFor("m-svg")).toBe("");
  });

  it("seeds the rId allocator from takenRelIds to avoid collisions on a same-format save", () => {
    const model = docWithImages([image("m1"), paragraph()], [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }]);
    const plan = planMediaWrite(model, new Set(["rId1", "rId2"]));
    expect(plan.relIdFor("m1")).toBe("rId3");
  });

  it("seeds part naming from existingPartNames to avoid colliding with an unmodelled word/media part", () => {
    const model = docWithImages([image("m1"), paragraph()], [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }]);
    const plan = planMediaWrite(model, new Set(), new Set(["word/media/image1.png"]));
    expect(plan.zipEntries[0].name).toBe("word/media/image2.png");
  });

  it("dedups a mediaId referenced by two Image blocks into a single zip entry/relationship", () => {
    const model = docWithImages(
      [image("shared"), image("shared"), paragraph()],
      [{ mediaId: "shared", contentType: "image/png", data: "AAAA" }],
    );
    const plan = planMediaWrite(model);
    expect(plan.zipEntries).toHaveLength(1);
    expect(plan.relIdFor("shared")).toBe("rId1");
  });

  it("skips a referenced mediaId that has no matching model.media entry", () => {
    const model = docWithImages([image("missing"), paragraph()], []);
    const plan = planMediaWrite(model);
    expect(plan.zipEntries).toEqual([]);
    expect(plan.relIdFor("missing")).toBe("");
  });
});
