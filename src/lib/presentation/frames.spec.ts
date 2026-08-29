// SPEC-082: the direct-manipulation slide-object commands. Every command
// against a from-scratch deck (createBlankPresentation) — clamping, aspect,
// z-order maths, ungroup promotion (incl. foreign children), media minting
// against colliding names, fill/stroke — plus the end-to-end round trip
// (insert text box + shape + image → save → reopen) in both formats.
import { describe, expect, it } from "vitest";
import { createBlankPresentation } from "./slides";
import {
  DECK_DEFAULT_SHAPE_FILL,
  DECK_DEFAULT_SHAPE_STROKE,
  DEFAULT_TEXT_BOX_SIZE_CM,
  DUPLICATE_OFFSET_CM,
  FRAME_MIN_ON_SLIDE_CM,
  MIN_FRAME_SIZE_CM,
  deleteFrame,
  duplicateFrame,
  fitImageWithinCm,
  insertImageFrame,
  insertImageRect,
  insertShapeFrame,
  insertTableFrame,
  insertTextBoxFrame,
  mediaExtension,
  moveFrame,
  reorderFrame,
  resizeFrame,
  resizeTableColumnBoundary,
  resizeTableRowBoundary,
  setFrameFill,
  setFrameStroke,
  ungroupFrame,
} from "./frames";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { buildOdpFixtureBytes, buildPptxFixtureBytes } from "./fixtures/build";
import { unzip } from "../core/zip";
import { equalBytes } from "../core/equality";
import { MIN_COLUMN_WIDTH_CM, MIN_ROW_HEIGHT_CM } from "../core/tables";
import type {
  GroupFrame,
  ImageFrame,
  PresentationModel,
  ShapeFrame,
  Slide,
  SlideFrame,
  TextFrame,
} from "./model";
import type { MediaPart } from "../core/media";

// A 2×3 px PNG (real IHDR so decodeImageDimensions reads its natural size).
const PNG_2x3 = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAYAAADkc372AAAAEUlEQVR42mNk+M9QzwAFjDAaAF1eBFsPTQEwAAAAAElFTkSuQmCC";
function pngMedia(): MediaPart {
  return { mediaId: "", contentType: "image/png", data: PNG_2x3 };
}

function blank(): PresentationModel {
  return createBlankPresentation();
}

function only(slide: Slide): SlideFrame {
  return slide.frames[slide.frames.length - 1];
}

describe("insert commands", () => {
  it("insertTextBoxFrame appends an empty free text box and dirties the slide", () => {
    const model = blank();
    const id = insertTextBoxFrame(model, 0, { xCm: 2, yCm: 2, widthCm: 10, heightCm: 2.5 });
    const frame = only(model.slides[0]) as TextFrame;
    expect(frame.id).toBe(id);
    expect(frame.kind).toBe("text");
    expect(frame.role).toBeUndefined();
    expect(frame.paragraphs).toEqual([{ runs: [] }]);
    expect(model.slides[0].dirty).toBe(true);
  });

  it("insertShapeFrame lands the deck-neutral default style (AC-14)", () => {
    const model = blank();
    insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 });
    const frame = only(model.slides[0]) as ShapeFrame;
    expect(frame.kind).toBe("shape");
    expect(frame.fill).toEqual({ type: "solid", color: DECK_DEFAULT_SHAPE_FILL });
    expect(frame.stroke).toEqual(DECK_DEFAULT_SHAPE_STROKE);
  });

  it("insertTableFrame lands a centred 80%-width table with the caret cell (AC-5c)", () => {
    const model = blank();
    const id = insertTableFrame(model, 0, 3, 4)!;
    const frame = only(model.slides[0]);
    expect(frame.kind).toBe("table");
    expect(id).toBe(frame.id);
    if (frame.kind === "table") {
      expect(frame.widthCm).toBeCloseTo(model.slideWidthCm * 0.8, 2);
      expect(frame.columnWidthsCm).toHaveLength(4);
      expect(frame.rows).toHaveLength(3);
      expect(frame.firstRow).toBe(true);
    }
  });

  it("resizeTableColumnBoundary moves an internal boundary, keeping table width (AC-8a)", () => {
    const model = blank();
    const id = insertTableFrame(model, 0, 2, 2)!;
    const frame = only(model.slides[0]);
    if (frame.kind !== "table") throw new Error("expected table");
    const [w0, w1] = frame.columnWidthsCm;
    const totalBefore = frame.widthCm;
    resizeTableColumnBoundary(model, 0, id, 0, [...frame.columnWidthsCm], 1);
    expect(frame.columnWidthsCm[0]).toBeCloseTo(w0 + 1, 3);
    expect(frame.columnWidthsCm[1]).toBeCloseTo(w1 - 1, 3);
    // The right neighbour absorbs exactly what column 0 gains: width unchanged.
    expect(frame.widthCm).toBeCloseTo(totalBefore, 3);
  });

  it("resizeTableColumnBoundary floors both columns at MIN_COLUMN_WIDTH_CM (AC-8a)", () => {
    const model = blank();
    const id = insertTableFrame(model, 0, 1, 2)!;
    const frame = only(model.slides[0]);
    if (frame.kind !== "table") throw new Error("expected table");
    const origin = [...frame.columnWidthsCm];
    // Drag far past the neighbour: it clamps rather than going below the floor.
    resizeTableColumnBoundary(model, 0, id, 0, origin, 1000);
    expect(frame.columnWidthsCm[1]).toBeCloseTo(MIN_COLUMN_WIDTH_CM, 3);
    expect(frame.columnWidthsCm[0]).toBeCloseTo(origin[0] + origin[1] - MIN_COLUMN_WIDTH_CM, 3);
  });

  it("resizeTableRowBoundary sets the row height, floored at MIN_ROW_HEIGHT_CM (AC-8b)", () => {
    const model = blank();
    const id = insertTableFrame(model, 0, 2, 1)!;
    const frame = only(model.slides[0]);
    if (frame.kind !== "table") throw new Error("expected table");
    const origin = frame.rows[0].heightCm ?? 0;
    resizeTableRowBoundary(model, 0, id, 0, origin, 1.5);
    expect(frame.rows[0].heightCm).toBeCloseTo(origin + 1.5, 3);
    // A large negative delta cannot shrink past the floor.
    resizeTableRowBoundary(model, 0, id, 0, origin, -1000);
    expect(frame.rows[0].heightCm).toBeCloseTo(MIN_ROW_HEIGHT_CM, 3);
  });

  it("resizeTable* is a no-op on a non-table frame", () => {
    const model = blank();
    const id = insertTextBoxFrame(model, 0, { xCm: 1, yCm: 1, widthCm: 4, heightCm: 2 })!;
    expect(resizeTableColumnBoundary(model, 0, id, 0, [1, 1], 1)).toBeNull();
    expect(resizeTableRowBoundary(model, 0, id, 0, 1, 1)).toBeNull();
  });

  it("insertShapeFrame gives a line geometry a stroke but no fill", () => {
    const model = blank();
    insertShapeFrame(model, 0, "line", { xCm: 1, yCm: 1, widthCm: 3, heightCm: 0 });
    const frame = only(model.slides[0]) as ShapeFrame;
    expect(frame.fill).toBeUndefined();
    expect(frame.stroke).toEqual(DECK_DEFAULT_SHAPE_STROKE);
  });

  it("insertImageFrame adds a MediaPart + ImageFrame, selects it (AC-6/7)", () => {
    const model = blank();
    const id = insertImageFrame(model, 0, pngMedia(), { xCm: 1, yCm: 1, widthCm: 4, heightCm: 6 });
    const frame = only(model.slides[0]) as ImageFrame;
    expect(frame.id).toBe(id);
    expect(frame.kind).toBe("image");
    expect(model.media).toHaveLength(1);
    expect(model.media![0].mediaId).toBe("ppt/media/image1.png");
    expect(frame.mediaId).toBe("ppt/media/image1.png");
  });
});

describe("media minting", () => {
  it("mints past the highest existing image suffix across opaque + media", () => {
    const model = blank();
    model.opaque = [
      { name: "ppt/media/image3.png", data: "" },
      { name: "ppt/slides/slide1.xml", data: "" },
    ];
    model.media = [{ mediaId: "ppt/media/image7.jpeg", contentType: "image/jpeg", data: "" }];
    insertImageFrame(model, 0, pngMedia(), { xCm: 0, yCm: 0, widthCm: 2, heightCm: 3 });
    expect(model.media![model.media!.length - 1].mediaId).toBe("ppt/media/image8.png");
  });

  it("uses the odp Pictures/ directory for an odp deck", () => {
    const model = blank();
    model.format = "odp";
    insertImageFrame(model, 0, pngMedia(), { xCm: 0, yCm: 0, widthCm: 2, heightCm: 3 });
    expect(model.media![0].mediaId).toBe("Pictures/image1.png");
  });

  it("maps webp to a webp extension", () => {
    expect(mediaExtension("image/webp")).toBe("webp");
    expect(mediaExtension("image/png")).toBe("png");
  });
});

describe("image fit", () => {
  it("scales down to fit within a box, aspect preserved (AC-5)", () => {
    // 400×200px ≈ 10.58×5.29cm; fit box 5×5 → width-bound to 5cm, 2.5cm tall.
    const size = fitImageWithinCm({ widthPx: 400, heightPx: 200 }, 5, 5);
    expect(size.widthCm).toBeCloseTo(5, 3);
    expect(size.heightCm).toBeCloseTo(2.5, 3);
  });

  it("keeps a small image at natural size", () => {
    const size = fitImageWithinCm({ widthPx: 10, heightPx: 10 }, 50, 50);
    expect(size.widthCm).toBeLessThan(1);
    expect(size.widthCm).toBeCloseTo(size.heightCm, 3);
  });

  it("centres the inserted image on the slide", () => {
    const model = blank();
    const rect = insertImageRect(model, { widthPx: 100, heightPx: 100 });
    expect(rect.xCm).toBeCloseTo((model.slideWidthCm - rect.widthCm) / 2, 2);
    expect(rect.yCm).toBeCloseTo((model.slideHeightCm - rect.heightCm) / 2, 2);
  });
});

describe("move / clamp", () => {
  it("nudges by a delta and clamps at least 1cm on the slide (AC-18)", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 5, yCm: 5, widthCm: 3, heightCm: 2 })!;
    moveFrame(model, 0, id, -100, -100);
    const frame = model.slides[0].frames.find((f) => f.id === id)!;
    // xCm clamps to FRAME_MIN_ON_SLIDE - width; yCm to FRAME_MIN_ON_SLIDE - height.
    expect(frame.xCm).toBeCloseTo(FRAME_MIN_ON_SLIDE_CM - 3, 3);
    expect(frame.yCm).toBeCloseTo(FRAME_MIN_ON_SLIDE_CM - 2, 3);
  });

  it("clamps the far edge to slideWidth - 1cm", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 5, yCm: 5, widthCm: 3, heightCm: 2 })!;
    moveFrame(model, 0, id, 1000, 0);
    const frame = model.slides[0].frames.find((f) => f.id === id)!;
    expect(frame.xCm).toBeCloseTo(model.slideWidthCm - FRAME_MIN_ON_SLIDE_CM, 3);
  });
});

describe("resize", () => {
  it("enforces the 0.5cm minimum (AC-24)", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 })!;
    resizeFrame(model, 0, id, { xCm: 1, yCm: 1, widthCm: 0.1, heightCm: 0.1 });
    const frame = model.slides[0].frames.find((f) => f.id === id)!;
    expect(frame.widthCm).toBe(MIN_FRAME_SIZE_CM);
    expect(frame.heightCm).toBe(MIN_FRAME_SIZE_CM);
  });
});

describe("duplicate / delete", () => {
  it("duplicates offset by +0.5/+0.5 with a fresh id (AC-28)", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 2, yCm: 2, widthCm: 3, heightCm: 2 })!;
    const copyId = duplicateFrame(model, 0, id)!;
    const copy = model.slides[0].frames.find((f) => f.id === copyId) as ShapeFrame;
    expect(copyId).not.toBe(id);
    expect(copy.xCm).toBeCloseTo(2 + DUPLICATE_OFFSET_CM, 3);
    expect(copy.yCm).toBeCloseTo(2 + DUPLICATE_OFFSET_CM, 3);
  });

  it("delete removes the frame and returns the neighbour to select", () => {
    const model = blank();
    const a = insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 })!;
    const b = insertShapeFrame(model, 0, "ellipse", { xCm: 2, yCm: 2, widthCm: 3, heightCm: 2 })!;
    const next = deleteFrame(model, 0, b);
    expect(next).toBe(a);
    expect(model.slides[0].frames.some((f) => f.id === b)).toBe(false);
  });
});

function order(model: PresentationModel): string[] {
  return model.slides[0].frames.map((f) => f.id);
}

describe("z-order", () => {
  it("reorders to front / back / forward / backward (AC-30)", () => {
    const model = blank();
    // start with two placeholders; add three shapes a,b,c on top.
    const a = insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 1, heightCm: 1 })!;
    const b = insertShapeFrame(model, 0, "rect", { xCm: 2, yCm: 2, widthCm: 1, heightCm: 1 })!;
    const c = insertShapeFrame(model, 0, "rect", { xCm: 3, yCm: 3, widthCm: 1, heightCm: 1 })!;
    reorderFrame(model, 0, a, "back");
    expect(order(model)[0]).toBe(a);
    reorderFrame(model, 0, c, "front");
    expect(order(model)[order(model).length - 1]).toBe(c);
    const before = order(model).indexOf(b);
    reorderFrame(model, 0, b, "backward");
    expect(order(model).indexOf(b)).toBe(before - 1);
    reorderFrame(model, 0, b, "forward");
    expect(order(model).indexOf(b)).toBe(before);
  });
});

describe("ungroup", () => {
  function groupSlide(): { model: PresentationModel; groupId: string; childIds: string[] } {
    const model = blank();
    const child1: ShapeFrame = {
      id: "c1",
      xCm: 1,
      yCm: 1,
      widthCm: 2,
      heightCm: 2,
      kind: "shape",
      geometry: "rect",
      fill: { type: "solid", color: "#ff0000" },
    };
    const child2: ForeignChild = {
      id: "c2",
      xCm: 3,
      yCm: 3,
      widthCm: 2,
      heightCm: 2,
      kind: "foreign",
      xml: "<p:sp/>",
      label: "",
    };
    const group: GroupFrame = {
      id: "g1",
      xCm: 0,
      yCm: 0,
      widthCm: 6,
      heightCm: 6,
      kind: "group",
      children: [child1, child2],
      xml: "<p:grpSp>verbatim</p:grpSp>",
    };
    model.slides[0].frames.push(group);
    return { model, groupId: "g1", childIds: ["c1", "c2"] };
  }

  it("promotes children in place, drops the xml, dirties (AC-32/33/34)", () => {
    const { model, groupId } = groupSlide();
    const promoted = ungroupFrame(model, 0, groupId)!;
    expect(model.slides[0].frames.some((f) => f.kind === "group")).toBe(false);
    expect(promoted).toHaveLength(2);
    const frames = model.slides[0].frames;
    expect(frames.some((f) => f.id === promoted[0])).toBe(true);
    expect(model.slides[0].dirty).toBe(true);
  });

  it("keeps foreign children as foreign frames with a stable label (AC-36/52)", () => {
    const { model, groupId } = groupSlide();
    ungroupFrame(model, 0, groupId);
    const foreign = model.slides[0].frames.find((f) => f.kind === "foreign") as ForeignChild;
    expect(foreign).toBeTruthy();
    expect(foreign.label).toBe("Object");
  });

  it("is a no-op for a non-group frame", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 1, heightCm: 1 })!;
    expect(ungroupFrame(model, 0, id)).toBeNull();
  });
});

type ForeignChild = Extract<SlideFrame, { kind: "foreign" }>;

describe("fill / stroke", () => {
  it("sets a shape fill including No fill (AC-37)", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 })!;
    setFrameFill(model, 0, id, { type: "none" });
    expect((model.slides[0].frames.find((f) => f.id === id) as ShapeFrame).fill).toEqual({ type: "none" });
  });

  it("clamps stroke width to 0.25–8pt (AC-39)", () => {
    const model = blank();
    const id = insertShapeFrame(model, 0, "rect", { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 })!;
    setFrameStroke(model, 0, id, { color: "#000000", widthPt: 99 });
    expect((model.slides[0].frames.find((f) => f.id === id) as ShapeFrame).stroke!.widthPt).toBe(8);
    setFrameStroke(model, 0, id, { color: "#000000", widthPt: 0 });
    expect((model.slides[0].frames.find((f) => f.id === id) as ShapeFrame).stroke!.widthPt).toBe(0.25);
  });

  it("edits stroke on images and clears it with null (AC-40)", () => {
    const model = blank();
    const id = insertImageFrame(model, 0, pngMedia(), { xCm: 1, yCm: 1, widthCm: 2, heightCm: 3 })!;
    setFrameStroke(model, 0, id, { color: "#123456", widthPt: 2 });
    expect((model.slides[0].frames.find((f) => f.id === id) as ImageFrame).stroke).toEqual({ color: "#123456", widthPt: 2 });
    setFrameStroke(model, 0, id, null);
    expect((model.slides[0].frames.find((f) => f.id === id) as ImageFrame).stroke).toBeUndefined();
  });
});

describe("round trip (insert text box + shape + image)", () => {
  async function insertThree(model: PresentationModel): Promise<void> {
    insertTextBoxFrame(model, 0, { xCm: 1, yCm: 1, ...DEFAULT_TEXT_BOX_SIZE_CM });
    insertShapeFrame(model, 0, "ellipse", { xCm: 4, yCm: 4, widthCm: 5, heightCm: 3 });
    insertImageFrame(model, 0, pngMedia(), insertImageRect(model, { widthPx: 2, heightPx: 3 }));
  }

  for (const format of ["pptx", "odp"] as const) {
    it(`preserves geometry, z-order, style and media in ${format}`, async () => {
      const model = blank();
      model.format = format;
      await insertThree(model);
      const bytes = format === "pptx" ? (await writePptx(model)).bytes : (await writeOdp(model)).bytes;
      const reopened = await readPresentationBytes(bytes, format);

      const kinds = reopened.slides[0].frames.map((f) => f.kind);
      expect(kinds).toContain("text");
      expect(kinds).toContain("shape");
      expect(kinds).toContain("image");
      // Image last (inserted on top).
      expect(kinds[kinds.length - 1]).toBe("image");

      const shape = reopened.slides[0].frames.find((f) => f.kind === "shape") as ShapeFrame;
      expect(shape.geometry).toBe("ellipse");
      expect(shape.fill).toEqual({ type: "solid", color: DECK_DEFAULT_SHAPE_FILL });
      expect(shape.widthCm).toBeCloseTo(5, 1);
      expect(shape.heightCm).toBeCloseTo(3, 1);

      const image = reopened.slides[0].frames.find((f) => f.kind === "image") as ImageFrame;
      expect(reopened.media?.some((m) => m.mediaId === image.mediaId)).toBe(true);
    });
  }
});

describe("inserting an image on an opened deck", () => {
  it("pptx: adds media, Content-Types default and slide .rels image rel; untouched parts byte-verbatim (AC-47/48/50)", async () => {
    const original = await buildPptxFixtureBytes();
    const model = await readPresentationBytes(original, "pptx");
    const before = await unzip((await writePptx(model)).bytes);

    insertImageFrame(model, 0, pngMedia(), { xCm: 1, yCm: 1, widthCm: 2, heightCm: 3 });
    const after = await unzip((await writePptx(model)).bytes);

    const slidePart = model.slides[0].partName!;
    const slideRelsPart = slidePart.replace(/slides\/(slide\d+)\.xml$/, "slides/_rels/$1.xml.rels");

    // New media part present.
    const mediaId = model.media![model.media!.length - 1].mediaId;
    expect(after.has(mediaId)).toBe(true);
    // Content-Types gained a png default.
    expect(new TextDecoder().decode(after.get("[Content_Types].xml")!)).toMatch(/Extension="png"/i);
    // The slide .rels gained an image relationship.
    expect(new TextDecoder().decode(after.get(slideRelsPart)!)).toMatch(/relationships\/image/);

    // Every part untouched by the insert re-emits byte-verbatim.
    const touched = new Set([slidePart, slideRelsPart, "[Content_Types].xml", mediaId]);
    for (const [name, data] of before) {
      if (touched.has(name)) {
        continue;
      }
      expect(after.has(name), `missing ${name}`).toBe(true);
      expect(equalBytes(after.get(name)!, data), `changed ${name}`).toBe(true);
    }
  });

  it("odp: adds Pictures/ media and a manifest entry; untouched parts byte-verbatim (AC-49/50)", async () => {
    const original = await buildOdpFixtureBytes();
    const model = await readPresentationBytes(original, "odp");
    const before = await unzip((await writeOdp(model)).bytes);

    insertImageFrame(model, 0, pngMedia(), { xCm: 1, yCm: 1, widthCm: 2, heightCm: 3 });
    const after = await unzip((await writeOdp(model)).bytes);

    const mediaId = model.media![model.media!.length - 1].mediaId;
    expect(mediaId.startsWith("Pictures/")).toBe(true);
    expect(after.has(mediaId)).toBe(true);
    expect(new TextDecoder().decode(after.get("META-INF/manifest.xml")!)).toContain(mediaId);

    const touched = new Set(["content.xml", "META-INF/manifest.xml", mediaId]);
    for (const [name, data] of before) {
      if (touched.has(name)) {
        continue;
      }
      expect(after.has(name), `missing ${name}`).toBe(true);
      expect(equalBytes(after.get(name)!, data), `changed ${name}`).toBe(true);
    }
  });
});
