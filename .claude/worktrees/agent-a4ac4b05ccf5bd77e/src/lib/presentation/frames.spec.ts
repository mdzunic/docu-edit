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
  groupFrames,
  insertMediaFrame,
  insertTableFrame,
  insertTextBoxFrame,
  markFrameEdited,
  MEDIA_AUDIO_SIZE_CM,
  MEDIA_VIDEO_SIZE_CM,
  moveFrame,
  moveFrames,
  reorderFrame,
  reorderFrames,
  resizeFrame,
  setFrameAdjustments,
  setFrameRotation,
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
  MediaFrame,
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

  // (webp/mp4/… extension mapping folded into core/media.ts's
  // extensionForContentType — covered by core/media.spec.ts.)
});

function mp4(): MediaPart {
  return { mediaId: "", contentType: "video/mp4", data: "AAAA" };
}

function wav(): MediaPart {
  return { mediaId: "", contentType: "audio/wav", data: "AAAA" };
}

describe("insertMediaFrame (SPEC-095)", () => {
  it("lands a video at a 16×9 box, centred, selected, and mints a media part", () => {
    const model = blank();
    const id = insertMediaFrame(model, 0, mp4(), "video");
    const frame = only(model.slides[0]) as MediaFrame;
    expect(frame.id).toBe(id);
    expect(frame.kind).toBe("media");
    expect(frame.media).toBe("video");
    expect(frame.widthCm).toBe(MEDIA_VIDEO_SIZE_CM.widthCm);
    expect(frame.heightCm).toBe(MEDIA_VIDEO_SIZE_CM.heightCm);
    expect(frame.xCm).toBeCloseTo((model.slideWidthCm - 16) / 2, 3);
    expect(frame.yCm).toBeCloseTo((model.slideHeightCm - 9) / 2, 3);
    expect(model.media).toHaveLength(1);
    expect(model.media![0].mediaId).toBe("ppt/media/image1.mp4");
    expect(frame.mediaId).toBe("ppt/media/image1.mp4");
    expect(frame.contentType).toBe("video/mp4");
    expect(model.slides[0].dirty).toBe(true);
  });

  it("lands audio at an 8×1.6 strip, minted with a .wav extension", () => {
    const model = blank();
    insertMediaFrame(model, 0, wav(), "audio");
    const frame = only(model.slides[0]) as MediaFrame;
    expect(frame.media).toBe("audio");
    expect(frame.widthCm).toBe(MEDIA_AUDIO_SIZE_CM.widthCm);
    expect(frame.heightCm).toBe(MEDIA_AUDIO_SIZE_CM.heightCm);
    expect(model.media![0].mediaId).toBe("ppt/media/image1.wav");
  });

  it("clamps an oversized box into a small slide so at least a corner stays on", () => {
    const model = blank();
    model.slideWidthCm = 10;
    model.slideHeightCm = 6;
    insertMediaFrame(model, 0, mp4(), "video");
    const frame = only(model.slides[0]) as MediaFrame;
    // Box shrinks to the slide when it would overflow.
    expect(frame.widthCm).toBeLessThanOrEqual(10);
    expect(frame.heightCm).toBeLessThanOrEqual(6);
    expect(frame.xCm).toBeGreaterThanOrEqual(FRAME_MIN_ON_SLIDE_CM - frame.widthCm);
  });

  it("markFrameEdited drops the frame's sourceXml and dirties the slide", () => {
    const model = blank();
    insertMediaFrame(model, 0, mp4(), "video");
    const slide = model.slides[0];
    const frame = only(slide) as MediaFrame;
    frame.sourceXml = "<p:pic/>";
    delete slide.dirty;
    markFrameEdited(slide, frame, model.format);
    expect(frame.sourceXml).toBeUndefined();
    expect(slide.dirty).toBe(true);
  });

  it("moving a media frame drops its sourceXml (geometry is a modelled field)", () => {
    const model = blank();
    const id = insertMediaFrame(model, 0, mp4(), "video");
    const frame = only(model.slides[0]) as MediaFrame;
    frame.sourceXml = "<p:pic/>";
    moveFrame(model, 0, id!, 1, 1);
    expect(frame.sourceXml).toBeUndefined();
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

// --- SPEC-096: multi-selection commands, rotation, groups -------------------

/** A blank deck with the placeholder frames cleared, so index maths below is
 * over exactly the frames each test inserts. */
function bare(): PresentationModel {
  const model = blank();
  model.slides[0].frames = [];
  return model;
}

function shapeAt(model: PresentationModel, x: number, y: number, w = 2, h = 2): string {
  return insertShapeFrame(model, 0, "rect", { xCm: x, yCm: y, widthCm: w, heightCm: h })!;
}

describe("moveFrames (SPEC-096)", () => {
  it("moves every frame by the same delta in slide space", () => {
    const model = bare();
    const a = shapeAt(model, 1, 1);
    const b = shapeAt(model, 5, 5);
    moveFrames(model, 0, [a, b], 2, 1);
    const [fa, fb] = model.slides[0].frames;
    expect([fa.xCm, fa.yCm]).toEqual([3, 2]);
    expect([fb.xCm, fb.yCm]).toEqual([7, 6]);
  });

  it("clamps on the UNION box so relative positions never shift (AC-13)", () => {
    const model = bare();
    const a = shapeAt(model, 1, 1);
    const b = shapeAt(model, 5, 5);
    moveFrames(model, 0, [a, b], -100, 0);
    const [fa, fb] = model.slides[0].frames;
    // Union spans x 1..7 → clamped so its right edge keeps 1cm on slide:
    // minX = 1 - 6 + ... union width 6, minX clamped to 1 − 6 = −5.
    expect(fb.xCm - fa.xCm).toBe(4); // relative gap intact
    expect(fb.xCm + fb.widthCm).toBe(FRAME_MIN_ON_SLIDE_CM);
  });

  it("uses the rotated box for the clamp (AC-30)", () => {
    const model = bare();
    const a = shapeAt(model, 2, 2, 6, 2);
    setFrameRotation(model, 0, a, 90);
    // Rotated 90° the 6×2 frame's axis-aligned box is 2×6 (padX −2), so the
    // clamp must keep 1cm of the ROTATED box on the slide.
    moveFrames(model, 0, [a], -100, 0);
    const frame = model.slides[0].frames[0];
    const padX = (2 - 6) / 2;
    const boxRight = frame.xCm - padX + 2;
    expect(boxRight).toBeCloseTo(FRAME_MIN_ON_SLIDE_CM, 3);
  });
});

describe("setFrameRotation (SPEC-096)", () => {
  it("stores whole degrees 0–359 and deletes zero", () => {
    const model = bare();
    const id = shapeAt(model, 2, 2);
    setFrameRotation(model, 0, id, 37.4);
    expect(model.slides[0].frames[0].rotationDeg).toBe(37);
    setFrameRotation(model, 0, id, 360);
    expect(model.slides[0].frames[0].rotationDeg).toBeUndefined();
    setFrameRotation(model, 0, id, -90);
    expect(model.slides[0].frames[0].rotationDeg).toBe(270);
  });

  it("composes a group rotation onto the children and deletes the group's own (AC-45)", () => {
    const model = bare();
    const a = shapeAt(model, 2, 2, 2, 2); // centre (3,3)
    const b = shapeAt(model, 6, 2, 2, 2); // centre (7,3)
    const groupId = groupFrames(model, 0, [a, b])!;
    setFrameRotation(model, 0, groupId, 90);
    const group = model.slides[0].frames[0] as GroupFrame;
    expect(group.rotationDeg).toBeUndefined();
    expect(group.xml).toBeUndefined();
    const [ca, cb] = group.children;
    // Group centre (5,3); rotating 90° cw sends (3,3)→(5,1) and (7,3)→(5,5).
    expect([ca.xCm + 1, ca.yCm + 1]).toEqual([5, 1]);
    expect([cb.xCm + 1, cb.yCm + 1]).toEqual([5, 5]);
    expect(ca.rotationDeg).toBe(90);
    expect(cb.rotationDeg).toBe(90);
  });
});

describe("setFrameAdjustments (SPEC-096)", () => {
  it("stores clamped fractions on a shape and is a no-op elsewhere", () => {
    const model = bare();
    const id = insertShapeFrame(model, 0, "roundRect", { xCm: 1, yCm: 1, widthCm: 4, heightCm: 2 })!;
    setFrameAdjustments(model, 0, id, [0.3]);
    expect((model.slides[0].frames[0] as ShapeFrame).adjustments).toEqual([0.3]);
    setFrameAdjustments(model, 0, id, [99]);
    const clamped = (model.slides[0].frames[0] as ShapeFrame).adjustments![0];
    expect(clamped).toBeLessThanOrEqual(0.5);
    const text = insertTextBoxFrame(model, 0, { xCm: 1, yCm: 5, widthCm: 2, heightCm: 1 })!;
    expect(setFrameAdjustments(model, 0, text, [0.5])).toBeNull();
  });
});

describe("groupFrames (SPEC-096)", () => {
  it("replaces the frames with one authored group at their union box, order kept (AC-40)", () => {
    const model = bare();
    const a = shapeAt(model, 2, 2, 2, 2);
    const b = shapeAt(model, 6, 4, 2, 3);
    const c = shapeAt(model, 10, 1, 1, 1); // stays outside the group
    const groupId = groupFrames(model, 0, [a, b])!;
    const frames = model.slides[0].frames;
    expect(frames).toHaveLength(2);
    const group = frames.find((f) => f.id === groupId) as GroupFrame;
    expect(group.kind).toBe("group");
    expect(group.xml).toBeUndefined();
    expect({ x: group.xCm, y: group.yCm, w: group.widthCm, h: group.heightCm }).toEqual({
      x: 2,
      y: 2,
      w: 6,
      h: 5,
    });
    expect(group.children.map((f) => f.id)).toEqual([a, b]);
    expect(frames.map((f) => f.id)).toEqual([groupId, c]);
    expect(model.slides[0].dirty).toBe(true);
  });

  it("needs at least two matching top-level frames", () => {
    const model = bare();
    const a = shapeAt(model, 2, 2);
    expect(groupFrames(model, 0, [a])).toBeNull();
    expect(groupFrames(model, 0, ["nope", "nada"])).toBeNull();
  });

  it("drops animations targeting the grouped frames (out of scope: animating groups)", () => {
    const model = bare();
    const a = shapeAt(model, 2, 2);
    const b = shapeAt(model, 6, 2);
    const c = shapeAt(model, 10, 2);
    model.slides[0].animations = [
      { frameId: a, effect: "appear", trigger: "click" },
      { frameId: c, effect: "fadeIn", trigger: "click" },
    ];
    groupFrames(model, 0, [a, b]);
    expect(model.slides[0].animations).toEqual([{ frameId: c, effect: "fadeIn", trigger: "click" }]);
  });
});

function verbatimGroup(model: PresentationModel): GroupFrame {
  const child: ShapeFrame = { id: "vc", xCm: 2, yCm: 2, widthCm: 2, heightCm: 2, kind: "shape", geometry: "rect" };
  const group: GroupFrame = {
    id: "vg",
    xCm: 2,
    yCm: 2,
    widthCm: 2,
    heightCm: 2,
    kind: "group",
    children: [child],
    xml: "<p:grpSp>verbatim</p:grpSp>",
  };
  model.slides[0].frames.push(group);
  return group;
}

describe("group xml deletion on every mutating path (SPEC-096 AC-41)", () => {
  it("move shifts the children with the box and deletes xml", () => {
    const model = bare();
    const group = verbatimGroup(model);
    moveFrame(model, 0, "vg", 1, 1);
    expect(group.xml).toBeUndefined();
    expect([group.children[0].xCm, group.children[0].yCm]).toEqual([3, 3]);
  });

  it("resize scales the children into the new box and deletes xml", () => {
    const model = bare();
    const group = verbatimGroup(model);
    resizeFrame(model, 0, "vg", { xCm: 2, yCm: 2, widthCm: 4, heightCm: 6 });
    expect(group.xml).toBeUndefined();
    expect(group.children[0].widthCm).toBe(4);
    expect(group.children[0].heightCm).toBe(6);
  });

  it("rotation composes onto children and deletes xml", () => {
    const model = bare();
    const group = verbatimGroup(model);
    setFrameRotation(model, 0, "vg", 90);
    expect(group.xml).toBeUndefined();
  });

  it("duplicate's copy loses the xml (its geometry moved); the original keeps it", () => {
    const model = bare();
    const group = verbatimGroup(model);
    const copyId = duplicateFrame(model, 0, "vg")!;
    const copy = model.slides[0].frames.find((f) => f.id === copyId) as GroupFrame;
    expect(copy.xml).toBeUndefined();
    expect(copy.children[0].xCm).toBe(2.5);
    expect(group.xml).toBe("<p:grpSp>verbatim</p:grpSp>");
  });
});

describe("reorderFrames (SPEC-096)", () => {
  function fourShapes(model: PresentationModel): string[] {
    return [shapeAt(model, 1, 1), shapeAt(model, 4, 1), shapeAt(model, 7, 1), shapeAt(model, 10, 1)];
  }

  it("moves a block to front preserving relative order (AC-17)", () => {
    const model = bare();
    const [a, b, c, d] = fourShapes(model);
    reorderFrames(model, 0, [a, c], "front");
    expect(model.slides[0].frames.map((f) => f.id)).toEqual([b, d, a, c]);
  });

  it("moves a block to back preserving relative order", () => {
    const model = bare();
    const [a, b, c, d] = fourShapes(model);
    reorderFrames(model, 0, [b, d], "back");
    expect(model.slides[0].frames.map((f) => f.id)).toEqual([b, d, a, c]);
  });

  it("moves a block one step forward / backward contiguously", () => {
    const model = bare();
    const [a, b, c, d] = fourShapes(model);
    reorderFrames(model, 0, [a, b], "forward");
    expect(model.slides[0].frames.map((f) => f.id)).toEqual([c, a, b, d]);
    reorderFrames(model, 0, [a, b], "backward");
    expect(model.slides[0].frames.map((f) => f.id)).toEqual([a, b, c, d]);
  });

  it("is a no-op (slide untouched) when the block already sits at that end (AC-18)", () => {
    const model = bare();
    const [a, b, c, d] = fourShapes(model);
    delete model.slides[0].dirty;
    reorderFrames(model, 0, [c, d], "front");
    expect(model.slides[0].frames.map((f) => f.id)).toEqual([a, b, c, d]);
    expect(model.slides[0].dirty).toBeUndefined();
    reorderFrames(model, 0, [a, b], "back");
    expect(model.slides[0].dirty).toBeUndefined();
  });
});

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
