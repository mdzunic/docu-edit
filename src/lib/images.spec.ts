import { describe, it, expect } from "vitest";
import {
  ACCEPTED_IMAGE_MIME_TYPES,
  IMAGE_TOO_LARGE_MESSAGE,
  MAX_IMAGE_SIZE_BYTES,
  MIN_IMAGE_SIZE_CM,
  cmToPx,
  deleteImage,
  fitImageSizeCm,
  imageAt,
  insertImage,
  isAcceptedImageType,
  isImageTooLarge,
  pxToCm,
  resetImageSize,
  setImageAltText,
  setImageSize,
  setImageWrap,
  textAreaWidthCm,
  type ImageLocation,
  type NewImageMedia,
} from "./images";
import { caretSelection, type ModelSelection } from "./commands";
import { equalStructural } from "./equality";
import { blockText, type CaretPosition } from "./textEditing";
import { createDocument, isImage, type DocumentModel, type Image, type Paragraph, type Table } from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function doc(...blocks: Array<Paragraph | Table | Image>): DocumentModel {
  const model = createDocument("images-fixture");
  model.sections = [{ blocks }];
  return model;
}

function cell(text: string) {
  return { blocks: [paragraph(text)] };
}

function grid(rows: string[][]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({ cells: cells.map(cell) })),
    columnWidthsCm: rows[0].map(() => 5),
  };
}

function caret(block: number, offset: number, cellPath?: { row: number; col: number; block: number }): CaretPosition {
  const position: CaretPosition = { section: 0, block, offset };
  if (cellPath) {
    position.cell = cellPath;
  }
  return position;
}

function selectionAt(block: number, offset: number): ModelSelection {
  return caretSelection(caret(block, offset));
}

function image(overrides: Partial<Image> = {}): Image {
  return {
    type: "image",
    mediaId: "media-1",
    naturalWidthPx: 400,
    naturalHeightPx: 200,
    widthCm: 10,
    heightCm: 5,
    wrap: "inline",
    ...overrides,
  };
}

function media(mediaId = "media-1"): NewImageMedia {
  return { mediaId, contentType: "image/png", data: "AAAA" };
}

// --- pure geometry helpers ---------------------------------------------------

describe("pxToCm / cmToPx (AC-21, 96 DPI)", () => {
  it("converts 96px to exactly 2.54cm", () => {
    expect(pxToCm(96)).toBeCloseTo(2.54, 10);
  });

  it("cmToPx inverts pxToCm", () => {
    for (const px of [0, 48, 96, 300, 1024]) {
      expect(cmToPx(pxToCm(px))).toBeCloseTo(px, 8);
    }
  });

  it("cmToPx(2.54) is 96", () => {
    expect(cmToPx(2.54)).toBeCloseTo(96, 10);
  });
});

describe("textAreaWidthCm (AC-5, AC-6)", () => {
  it("is the page width minus left/right margins for a blank document", () => {
    const model = createDocument("area");
    // A4 (21cm) minus 2 x 2.54cm standard margins.
    expect(textAreaWidthCm(model)).toBeCloseTo(15.92, 5);
  });

  it("never goes below 1cm even for degenerate margins", () => {
    const model = createDocument("area");
    model.defaultStyles.pageSize.widthCm = 3;
    model.defaultStyles.margins.leftCm = 5;
    model.defaultStyles.margins.rightCm = 5;
    expect(textAreaWidthCm(model)).toBe(1);
  });
});

describe("fitImageSizeCm (AC-5, AC-6)", () => {
  it("keeps the natural size (converted to cm) when it fits the text area", () => {
    // 100px at 96 DPI is ~2.65cm, well under any reasonable text area.
    const result = fitImageSizeCm(100, 50, 15.92);
    expect(result.widthCm).toBeCloseTo(pxToCm(100), 8);
    expect(result.heightCm).toBeCloseTo(pxToCm(50), 8);
  });

  it("scales down to the text area width, preserving aspect ratio", () => {
    // 4000px wide is far beyond the text area; height should scale by the
    // same factor as width.
    const naturalWidthCm = pxToCm(4000);
    const naturalHeightCm = pxToCm(2000);
    const textAreaCm = 15.92;
    const result = fitImageSizeCm(4000, 2000, textAreaCm);
    expect(result.widthCm).toBeCloseTo(textAreaCm, 8);
    const expectedHeight = naturalHeightCm * (textAreaCm / naturalWidthCm);
    expect(result.heightCm).toBeCloseTo(expectedHeight, 8);
    // Aspect ratio preserved.
    expect(result.widthCm / result.heightCm).toBeCloseTo(naturalWidthCm / naturalHeightCm, 8);
  });

  it("scales down only when strictly wider than the text area (equal fits as-is)", () => {
    const textAreaCm = pxToCm(1000);
    const result = fitImageSizeCm(1000, 500, textAreaCm);
    expect(result.widthCm).toBeCloseTo(textAreaCm, 8);
    expect(result.heightCm).toBeCloseTo(pxToCm(500), 8);
  });
});

describe("isImageTooLarge / isAcceptedImageType (AC-33, AC-1)", () => {
  it("accepts a file exactly at the size limit", () => {
    expect(isImageTooLarge(MAX_IMAGE_SIZE_BYTES)).toBe(false);
  });

  it("rejects a file one byte over the limit", () => {
    expect(isImageTooLarge(MAX_IMAGE_SIZE_BYTES + 1)).toBe(true);
  });

  it("has a human-readable too-large message", () => {
    expect(IMAGE_TOO_LARGE_MESSAGE.length).toBeGreaterThan(0);
  });

  it("accepts every declared mime type", () => {
    for (const type of ACCEPTED_IMAGE_MIME_TYPES) {
      expect(isAcceptedImageType(type)).toBe(true);
    }
  });

  it("rejects an unsupported mime type", () => {
    expect(isAcceptedImageType("image/bmp")).toBe(false);
    expect(isAcceptedImageType("application/pdf")).toBe(false);
  });
});

// --- insertImage --------------------------------------------------------------

describe("insertImage (AC-4)", () => {
  it("inserts before an empty caret paragraph, which becomes the trailing paragraph", () => {
    const model = doc(paragraph(""));
    const { model: next, selection } = insertImage(model, selectionAt(0, 0), media(), 400, 200);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(2);
    expect(isImage(blocks[0])).toBe(true);
    expect(blocks[1].type).toBe("paragraph");
    // Caret lands in the trailing paragraph, right after the image.
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("splits a paragraph at a middle offset, image between the two halves", () => {
    const model = doc(paragraph("hello world"));
    const { model: next, selection } = insertImage(model, selectionAt(0, 5), media(), 400, 200);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(blocks[0])).toBe("hello");
    expect(isImage(blocks[1])).toBe(true);
    expect(blockText(blocks[2])).toBe(" world");
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("inserts at the end of a paragraph, followed by a fresh trailing paragraph", () => {
    const model = doc(paragraph("above"));
    const { model: next, selection } = insertImage(model, selectionAt(0, 5), media(), 400, 200);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(blocks[0])).toBe("above");
    expect(isImage(blocks[1])).toBe(true);
    expect(blockText(blocks[2])).toBe("");
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("adds the media payload to model.media", () => {
    const model = doc(paragraph(""));
    const payload = media("m-42");
    const { model: next } = insertImage(model, selectionAt(0, 0), payload, 400, 200);
    expect(next.media).toEqual([{ mediaId: "m-42", contentType: "image/png", data: "AAAA" }]);
  });

  it("sizes the inserted image via fitImageSizeCm (fits the text area for a small natural size)", () => {
    const model = doc(paragraph(""));
    const { model: next } = insertImage(model, selectionAt(0, 0), media(), 400, 200);
    const inserted = next.sections[0].blocks[0] as Image;
    expect(inserted.widthCm).toBeCloseTo(pxToCm(400), 8);
    expect(inserted.heightCm).toBeCloseTo(pxToCm(200), 8);
    expect(inserted.naturalWidthPx).toBe(400);
    expect(inserted.naturalHeightPx).toBe(200);
  });

  it("defaults to inline wrap, honours an explicit wrap argument", () => {
    const model = doc(paragraph(""));
    const { model: inlineNext } = insertImage(model, selectionAt(0, 0), media(), 400, 200);
    expect((inlineNext.sections[0].blocks[0] as Image).wrap).toBe("inline");

    const { model: squareNext } = insertImage(model, selectionAt(0, 0), media(), 400, 200, "square");
    expect((squareNext.sections[0].blocks[0] as Image).wrap).toBe("square");
  });

  it("is a no-op for a caret inside a table cell (images are top-level blocks only)", () => {
    const model = doc(grid([["a"]]), paragraph(""));
    const inCell = caretSelection(caret(0, 0, { row: 0, col: 0, block: 0 }));
    const { model: next } = insertImage(model, inCell, media(), 400, 200);
    expect(next.sections[0].blocks).toHaveLength(2);
    expect(next.sections[0].blocks[0]).toBe(model.sections[0].blocks[0]);
    expect(next.media ?? []).toHaveLength(0);
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph("keep"));
    const snapshot = structuredClone(model);
    insertImage(model, selectionAt(0, 2), media(), 400, 200);
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

// --- setImageSize / resetImageSize -------------------------------------------

describe("setImageSize (AC-16, AC-19, AC-20)", () => {
  const location: ImageLocation = { section: 0, block: 0 };

  it("sets width/height in cm", () => {
    const model = doc(image(), paragraph(""));
    const { model: next } = setImageSize(model, selectionAt(0, 0), location, 8, 4);
    expect(imageAt(next, location)).toMatchObject({ widthCm: 8, heightCm: 4 });
  });

  it("clamps below MIN_IMAGE_SIZE_CM", () => {
    const model = doc(image(), paragraph(""));
    const { model: next } = setImageSize(model, selectionAt(0, 0), location, 0.01, 0.05);
    expect(imageAt(next, location)).toMatchObject({
      widthCm: MIN_IMAGE_SIZE_CM,
      heightCm: MIN_IMAGE_SIZE_CM,
    });
  });

  it("leaves a value at exactly the minimum unclamped", () => {
    const model = doc(image(), paragraph(""));
    const { model: next } = setImageSize(model, selectionAt(0, 0), location, MIN_IMAGE_SIZE_CM, MIN_IMAGE_SIZE_CM);
    expect(imageAt(next, location)).toMatchObject({
      widthCm: MIN_IMAGE_SIZE_CM,
      heightCm: MIN_IMAGE_SIZE_CM,
    });
  });

  it("is a no-op when the location doesn't address an image", () => {
    const model = doc(paragraph("text"));
    const { model: next } = setImageSize(model, selectionAt(0, 0), location, 8, 4);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });
});

describe("resetImageSize (AC-21)", () => {
  it("restores the natural pixel size converted to cm at 96 DPI", () => {
    const location: ImageLocation = { section: 0, block: 0 };
    const model = doc(image({ naturalWidthPx: 960, naturalHeightPx: 480, widthCm: 2, heightCm: 1 }), paragraph(""));
    const { model: next } = resetImageSize(model, selectionAt(0, 0), location);
    const result = imageAt(next, location)!;
    expect(result.widthCm).toBeCloseTo(pxToCm(960), 8);
    expect(result.heightCm).toBeCloseTo(pxToCm(480), 8);
  });
});

// --- setImageWrap / setImageAltText ------------------------------------------

describe("setImageWrap (AC-13, AC-14)", () => {
  it("updates the wrap mode", () => {
    const location: ImageLocation = { section: 0, block: 0 };
    const model = doc(image({ wrap: "inline" }), paragraph(""));
    const { model: next } = setImageWrap(model, selectionAt(0, 0), location, "behind");
    expect(imageAt(next, location)?.wrap).toBe("behind");
  });
});

describe("setImageAltText (AC-7, AC-15)", () => {
  const location: ImageLocation = { section: 0, block: 0 };

  it("sets alt text", () => {
    const model = doc(image(), paragraph(""));
    const { model: next } = setImageAltText(model, selectionAt(0, 0), location, "A red circle");
    expect(imageAt(next, location)?.altText).toBe("A red circle");
  });

  it("an empty string clears alt text", () => {
    const model = doc(image({ altText: "old description" }), paragraph(""));
    const { model: next } = setImageAltText(model, selectionAt(0, 0), location, "");
    expect(imageAt(next, location)?.altText).toBeUndefined();
    expect("altText" in imageAt(next, location)!).toBe(false);
  });
});

// --- deleteImage --------------------------------------------------------------

describe("deleteImage (AC-30, AC-31, AC-32)", () => {
  it("removes the image block and lands the caret where it was", () => {
    const model = doc(paragraph("before"), image(), paragraph("after"));
    const location: ImageLocation = { section: 0, block: 1 };
    const { model: next, selection } = deleteImage(model, selectionAt(1, 0), location);
    expect(next.sections[0].blocks.map(blockText)).toEqual(["before", "after"]);
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("leaves one empty paragraph when deleting the only block", () => {
    const model = doc(image());
    const location: ImageLocation = { section: 0, block: 0 };
    const { model: next, selection } = deleteImage(model, selectionAt(0, 0), location);
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("");
    expect(selection.focus).toEqual(caret(0, 0));
  });

  it("prunes model.media when no other image block references the mediaId", () => {
    const model = doc(image({ mediaId: "solo" }), paragraph(""));
    model.media = [{ mediaId: "solo", contentType: "image/png", data: "AAAA" }];
    const location: ImageLocation = { section: 0, block: 0 };
    const { model: next } = deleteImage(model, selectionAt(0, 0), location);
    expect(next.media).toEqual([]);
  });

  it("keeps a shared media entry alive while another image still references it, prunes it once the last use is gone", () => {
    const model = doc(image({ mediaId: "shared" }), image({ mediaId: "shared" }), paragraph(""));
    model.media = [{ mediaId: "shared", contentType: "image/png", data: "AAAA" }];
    const firstLocation: ImageLocation = { section: 0, block: 0 };
    const { model: afterFirstDelete } = deleteImage(model, selectionAt(0, 0), firstLocation);
    // One use remains (the second Image block, now at index 0) — the media
    // entry must survive (AC-32).
    expect(afterFirstDelete.media).toEqual([
      { mediaId: "shared", contentType: "image/png", data: "AAAA" },
    ]);
    expect(afterFirstDelete.sections[0].blocks.filter(isImage)).toHaveLength(1);

    const secondLocation: ImageLocation = { section: 0, block: 0 };
    const { model: afterSecondDelete } = deleteImage(afterFirstDelete, selectionAt(0, 0), secondLocation);
    // No use left — now it's pruned.
    expect(afterSecondDelete.media).toEqual([]);
  });

  it("is a no-op when the location doesn't address an image", () => {
    const model = doc(paragraph("text"));
    const location: ImageLocation = { section: 0, block: 0 };
    const { model: next } = deleteImage(model, selectionAt(0, 0), location);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });

  it("never mutates the input model", () => {
    const model = doc(image({ mediaId: "shared" }), image({ mediaId: "shared" }), paragraph(""));
    model.media = [{ mediaId: "shared", contentType: "image/png", data: "AAAA" }];
    const snapshot = structuredClone(model);
    deleteImage(model, selectionAt(0, 0), { section: 0, block: 0 });
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});
