// SPEC-057 native `draw:g` image members (.odt), AC-62/-63. A container holding
// a `draw:frame` picture member is written as a positioned `draw:frame` inside
// the `draw:g` (drawFrameGroupChildXml) and reads back as an Image member
// (groupReader's frame branch), with its media surviving the round trip.
import { describe, it, expect } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import { bytesToBase64 } from "../opaque";
import { createDocument, isDrawingGroup, isImage, isShape, type DocumentModel, type DrawingGroup } from "../documentModel";

function buildPngBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12);
  bytes[16] = (width >>> 24) & 0xff;
  bytes[17] = (width >>> 16) & 0xff;
  bytes[18] = (width >>> 8) & 0xff;
  bytes[19] = width & 0xff;
  bytes[20] = (height >>> 24) & 0xff;
  bytes[21] = (height >>> 16) & 0xff;
  bytes[22] = (height >>> 8) & 0xff;
  bytes[23] = height & 0xff;
  return bytes;
}

function modelWithImageGroup(): DocumentModel {
  const group: DrawingGroup = {
    type: "drawingGroup",
    widthCm: 6,
    heightCm: 3,
    wrap: "square",
    offsetXCm: 1,
    offsetYCm: 2,
    relativeToH: "column",
    relativeToV: "paragraph",
    members: [
      { type: "shape", geometry: "rect", widthCm: 3, heightCm: 3, wrap: "square", offsetXCm: 0, offsetYCm: 0 },
      { type: "image", mediaId: "Pictures/photo.png", naturalWidthPx: 400, naturalHeightPx: 200, widthCm: 3, heightCm: 2, wrap: "square", offsetXCm: 3, offsetYCm: 0, altText: "a photo" },
    ],
  };
  const model = createDocument("odt-group-image-test");
  model.format = "odt";
  model.sections[0].blocks = [model.sections[0].blocks[0], group];
  model.media = [{ mediaId: "Pictures/photo.png", contentType: "image/png", data: bytesToBase64(buildPngBytes(400, 200)) }];
  return model;
}

describe("SPEC-057 native draw:g image members round-trip (.odt) (AC-62/-63)", () => {
  it("writes an image member as a positioned draw:frame and reads it back", async () => {
    const { bytes } = await writeOdt(modelWithImageGroup());
    const reopened = await parseOdt(bytes);
    const blocks = reopened.sections.flatMap((s) => s.blocks);
    const group = blocks.find(isDrawingGroup) as DrawingGroup;
    expect(group).toBeDefined();
    expect(group.members).toHaveLength(2);
    const image = group.members.find(isImage);
    expect(image).toBeDefined();
    expect(image!.mediaId).toBeDefined();
    expect(image!.altText).toBe("a photo");
    // Members renormalize to a 0-origin box: the shape sits at x=0, the image
    // 3cm to its right.
    expect(group.members.find(isShape)!.offsetXCm).toBeCloseTo(0, 4);
    expect(image!.offsetXCm).toBeCloseTo(3, 4);
    expect(image!.widthCm).toBeCloseTo(3, 4);
  });
});
