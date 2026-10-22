// SPEC-112: the edited media model — which part is the current image, its
// size, and the save / copy / flatten transitions, every one asserted
// against a DEEP-FROZEN model (a mutation would throw), with the source
// part carried over by reference and its bytes untouched.
import { describe, expect, it } from "vitest";
import { downloadMedia, currentImageBytes, currentImageFileName } from "./download";
import {
  copiedWithStack,
  currentImagePart,
  currentImageSize,
  editBasePart,
  formatOfContentType,
  hasEdits,
  renderContentType,
  withFlattened,
  withSavedStack,
  type RenderedImage,
} from "./editModel";
import { mediaImageDataUri } from "./thumbnail";
import type { MediaModel } from "./model";
import type { ImageEdit } from "../core/imageEdits";

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) {
      deepFreeze(child);
    }
  }
  return value;
}

function photo(): MediaModel {
  return deepFreeze({
    kind: "media",
    id: "doc-1",
    title: "Photo",
    format: "jpeg",
    contentType: "image/jpeg",
    sourceMediaId: "src",
    media: [{ mediaId: "src", contentType: "image/jpeg", data: "AAEC" }],
    byteLength: 3,
    image: { widthPx: 160, heightPx: 120 },
  } as MediaModel);
}

const STACK: ImageEdit[] = [
  { kind: "crop", x: 0, y: 15, width: 160, height: 90 },
  { kind: "rotate", quarterTurns: 1 },
];

const RENDER: RenderedImage = { bytes: new Uint8Array([7, 8, 9]), contentType: "image/jpeg", widthPx: 90, heightPx: 160 };

function ids(...values: string[]) {
  let i = 0;
  return () => values[i++] ?? `id-${i}`;
}

describe("the current image", () => {
  it("is the source of an unedited document, verbatim", () => {
    const model = photo();
    expect(hasEdits(model)).toBe(false);
    expect(currentImagePart(model)).toBe(model.media[0]);
    expect(currentImageSize(model)).toEqual({ widthPx: 160, heightPx: 120 });
    expect(currentImageBytes(model)).toEqual({ bytes: new Uint8Array([0, 1, 2]), contentType: "image/jpeg" });
    expect(currentImageFileName(model)).toBe("Photo.jpeg");
    expect(mediaImageDataUri(model)).toBe("data:image/jpeg;base64,AAEC");
  });

  it("is the render of an edited one, at the stack's output size", () => {
    const saved = withSavedStack(photo(), STACK, RENDER, null, ids("render-1"));
    expect(hasEdits(saved)).toBe(true);
    expect(currentImagePart(saved)?.mediaId).toBe("render-1");
    expect(currentImageSize(saved)).toEqual({ widthPx: 90, heightPx: 160 });
    expect(currentImageBytes(saved).bytes).toEqual(new Uint8Array([7, 8, 9]));
    expect(mediaImageDataUri(saved)).toBe("data:image/jpeg;base64,BwgJ");
    // The stack's base is still the source.
    expect(editBasePart(saved)?.mediaId).toBe("src");
  });

  it("refuses to fall back to the original when an edited document lost its render", () => {
    const saved = withSavedStack(photo(), [{ kind: "redact", x: 0, y: 0, width: 5, height: 5, color: "#000000" }], RENDER);
    const missing = { ...saved, media: saved.media.filter((part) => part.mediaId === saved.sourceMediaId) };
    expect(currentImagePart(missing)).toBeUndefined();
    expect(() => downloadMedia(missing)).toThrow(/no saved render/);
    expect(mediaImageDataUri(missing)).toBe("");
  });

  it("renders as JPEG only for a JPEG source", () => {
    expect(renderContentType(photo())).toBe("image/jpeg");
    expect(renderContentType({ ...photo(), contentType: "image/png" })).toBe("image/png");
    expect(renderContentType({ ...photo(), contentType: "image/svg+xml" })).toBe("image/png");
    expect(formatOfContentType("image/png")).toBe("png");
    expect(formatOfContentType("image/jpeg")).toBe("jpeg");
  });
});

describe("Save", () => {
  it("stores the normalised stack and a fresh render beside the untouched source", () => {
    const model = photo();
    const saved = withSavedStack(model, STACK.toReversed(), RENDER, null, ids("render-1"));
    expect(saved).not.toBe(model);
    expect(saved.sourceMediaId).toBe("src");
    expect(saved.media[0]).toBe(model.media[0]);
    expect(saved.edits).toEqual(STACK);
    expect(saved.renderMediaId).toBe("render-1");
    expect(saved.media.map((part) => part.mediaId)).toEqual(["src", "render-1"]);
    // A second save regenerates the render under a new id, keeping one.
    const again = withSavedStack(saved, STACK, RENDER, null, ids("render-2"));
    expect(again.media.map((part) => part.mediaId)).toEqual(["src", "render-2"]);
  });

  it("drops the render and the base when the stack is empty (a reverted image)", () => {
    const saved = withSavedStack(photo(), STACK, RENDER);
    const reverted = withSavedStack(saved, [], null);
    expect(reverted.edits).toBeUndefined();
    expect(reverted.renderMediaId).toBeUndefined();
    expect(reverted.media).toEqual([photo().media[0]]);
    expect(currentImagePart(reverted)?.mediaId).toBe("src");
  });

  it("keeps an SVG as the source and stores its raster as the stack's base", () => {
    const svg = deepFreeze({
      ...photo(),
      format: "svg",
      contentType: "image/svg+xml",
      media: [{ mediaId: "src", contentType: "image/svg+xml", data: "PHN2Zz4=" }],
      image: undefined,
    } as MediaModel);
    const raster: RenderedImage = { bytes: new Uint8Array([1]), contentType: "image/png", widthPx: 300, heightPx: 150 };
    const saved = withSavedStack(svg, [{ kind: "rotate", quarterTurns: 2 }], { ...RENDER, contentType: "image/png" }, raster, ids("base", "render"));
    expect(saved.sourceMediaId).toBe("src");
    expect(saved.baseMediaId).toBe("base");
    expect(saved.image).toEqual({ widthPx: 300, heightPx: 150 });
    expect(editBasePart(saved)?.contentType).toBe("image/png");
    expect(currentImageFileName(saved)).toBe("Photo.png");
    // A later save keeps the stored base rather than rasterising again.
    const again = withSavedStack(saved, [{ kind: "rotate", quarterTurns: 1 }], RENDER, null, ids("render-2"));
    expect(again.baseMediaId).toBe("base");
    expect(again.media.map((part) => part.mediaId)).toEqual(["src", "base", "render-2"]);
  });
});

describe("Save as a copy", () => {
  it("makes a new document with the same source bytes and stack, leaving the original alone", () => {
    const model = photo();
    const copy = copiedWithStack(model, STACK, RENDER, null, "Photo (copy)", ids("doc-2", "render-c"));
    expect(copy.id).toBe("doc-2");
    expect(copy.title).toBe("Photo (copy)");
    expect(copy.sourceMediaId).toBe("src");
    expect(copy.media[0]).not.toBe(model.media[0]);
    expect(copy.media[0]).toEqual(model.media[0]);
    expect(copy.edits).toEqual(STACK);
    expect(copy.renderMediaId).toBe("render-c");
  });
});

describe("Flatten and discard original", () => {
  it("replaces the source part's bytes under the same id and clears the stack", () => {
    const saved = withSavedStack(photo(), STACK, RENDER);
    const flattened = withFlattened(saved, RENDER);
    expect(flattened.sourceMediaId).toBe("src");
    expect(flattened.media).toEqual([{ mediaId: "src", contentType: "image/jpeg", data: "BwgJ" }]);
    expect(flattened.edits).toBeUndefined();
    expect(flattened.renderMediaId).toBeUndefined();
    expect(flattened.image).toEqual({ widthPx: 90, heightPx: 160 });
    expect(flattened.byteLength).toBe(3);
    expect(JSON.stringify(flattened)).not.toContain("AAEC");
  });

  it("turns a flattened SVG into a PNG document", () => {
    const svg = deepFreeze({ ...photo(), format: "svg", contentType: "image/svg+xml" } as MediaModel);
    const flattened = withFlattened(svg, { ...RENDER, contentType: "image/png" });
    expect(flattened.format).toBe("png");
    expect(flattened.contentType).toBe("image/png");
    expect(currentImageFileName(flattened)).toBe("Photo.png");
  });
});
