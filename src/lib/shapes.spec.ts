// SPEC-051: shape/text-box structure commands. Mirrors images.spec.ts's
// fixture/testing conventions exactly — see that file's own header comment.
import { describe, it, expect } from "vitest";
import {
  DEFAULT_DRAWING_POSITION,
  deleteShape,
  formatDrawings,
  duplicateShape,
  ensureShapeTextFrame,
  freshGroupOrdinal,
  groupBoundingBoxCm,
  groupMembers,
  groupShapes,
  insertShape,
  insertTextBox,
  moveGroup,
  moveShape,
  nearestParagraph,
  nudgeShape,
  DUPLICATE_OFFSET_CM,
  reanchorShape,
  reorderShapeZ,
  resizeGroup,
  rotateGroup,
  setShapeAltText,
  setShapeFill,
  setShapeOutline,
  setShapeShadow,
  arrangeDrawings,
  setShapePosition,
  setShapeRotation,
  setShapePath,
  setShapeSize,
  setShapeWrap,
  shapeHasVisibleText,
  shapesInPaintOrder,
  shapeAt,
  ungroupShapes,
  type NewShape,
  type ParagraphSpan,
  type ShapeLocation,
} from "./shapes";
import {
  constrainedDragSizeCm,
  DEFAULT_AREA_SIZE_CM,
  DEFAULT_LINE_SIZE_CM,
  defaultClickSizeCm,
  DRAW_DRAG_THRESHOLD_CM,
  drawGestureBox,
  MIN_SHAPE_SIZE_CM,
  NUDGE_FINE_PT_CM,
  NUDGE_PT_CM,
  normalizeRotationDeg,
} from "./core/shapes";
import { caretSelection, commitShapeModel, deleteRange, insertText, splitBlock, withShapeModel, type ModelSelection } from "./commands";
import { equalStructural } from "./equality";
import { resolveAnchorBlockIndex } from "./drawingAnchors";
import { blockText, type CaretPosition } from "./textEditing";
import { createDocument, isImage, isShape, splitParagraphAt, type DocumentModel, type Image, type Paragraph, type Shape, type Table } from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function doc(...blocks: Array<Paragraph | Table | Image | Shape>): DocumentModel {
  const model = createDocument("shapes-fixture");
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

function shape(overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 3,
    heightCm: 2,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 0,
    offsetYCm: 0,
    ...overrides,
  };
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

// --- insertShape ---------------------------------------------------------------

describe("insertShape (AC-3..AC-9)", () => {
  it("inserts before an empty caret paragraph, which becomes the trailing paragraph", () => {
    const model = doc(paragraph(""));
    const spec: NewShape = { geometry: "rect", widthCm: 3, heightCm: 2 };
    const { model: next, selection } = insertShape(model, selectionAt(0, 0), spec);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(2);
    expect(isShape(blocks[0])).toBe(true);
    expect(blocks[1].type).toBe("paragraph");
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("splits a paragraph at a middle offset, shape between the two halves", () => {
    const model = doc(paragraph("hello world"));
    const spec: NewShape = { geometry: "ellipse", widthCm: 4, heightCm: 4 };
    const { model: next, selection } = insertShape(model, selectionAt(0, 5), spec);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(blocks[0])).toBe("hello");
    expect(isShape(blocks[1])).toBe(true);
    expect(blockText(blocks[2])).toBe(" world");
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("inserts at the end of a paragraph, followed by a fresh trailing paragraph", () => {
    const model = doc(paragraph("above"));
    const spec: NewShape = { geometry: "rect", widthCm: 3, heightCm: 2 };
    const { model: next, selection } = insertShape(model, selectionAt(0, 5), spec);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(blocks[0])).toBe("above");
    expect(isShape(blocks[1])).toBe(true);
    expect(blockText(blocks[2])).toBe("");
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("floats at the shared anchored default position (wrap square, column/paragraph, 0/0)", () => {
    const model = doc(paragraph(""));
    const spec: NewShape = { geometry: "rect", widthCm: 3, heightCm: 2 };
    const { model: next } = insertShape(model, selectionAt(0, 0), spec);
    const inserted = next.sections[0].blocks[0] as Shape;
    expect(inserted.wrap).toBe("square");
    expect(inserted.relativeToH).toBe(DEFAULT_DRAWING_POSITION.relativeToH);
    expect(inserted.relativeToV).toBe(DEFAULT_DRAWING_POSITION.relativeToV);
    expect(inserted.offsetXCm).toBe(0);
    expect(inserted.offsetYCm).toBe(0);
  });

  it("honours an explicit offset instead of the shared default", () => {
    const model = doc(paragraph(""));
    const spec: NewShape = { geometry: "rect", widthCm: 3, heightCm: 2, offsetXCm: 2.5, offsetYCm: 1.25 };
    const { model: next } = insertShape(model, selectionAt(0, 0), spec);
    const inserted = next.sections[0].blocks[0] as Shape;
    expect(inserted.offsetXCm).toBe(2.5);
    expect(inserted.offsetYCm).toBe(1.25);
  });

  it("carries optional fill/outline/text through when supplied", () => {
    const model = doc(paragraph(""));
    const spec: NewShape = {
      geometry: "rect",
      widthCm: 3,
      heightCm: 2,
      fill: { kind: "solid", color: "#ff0000" },
      outline: { color: "#000000", widthPt: 1, dash: "solid" },
      text: { blocks: [paragraph("hi")], vAlign: "top" },
    };
    const { model: next } = insertShape(model, selectionAt(0, 0), spec);
    const inserted = next.sections[0].blocks[0] as Shape;
    expect(inserted.fill).toEqual({ kind: "solid", color: "#ff0000" });
    expect(inserted.outline).toEqual({ color: "#000000", widthPt: 1, dash: "solid" });
    expect(inserted.text?.blocks).toHaveLength(1);
  });

  it("omits fill/outline/text entirely when not supplied", () => {
    const model = doc(paragraph(""));
    const spec: NewShape = { geometry: "rect", widthCm: 3, heightCm: 2 };
    const { model: next } = insertShape(model, selectionAt(0, 0), spec);
    const inserted = next.sections[0].blocks[0] as Shape;
    expect("fill" in inserted).toBe(false);
    expect("outline" in inserted).toBe(false);
    expect("text" in inserted).toBe(false);
  });

  it("is a no-op for a caret inside a table cell (shapes are top-level blocks only)", () => {
    const model = doc(grid([["a"]]), paragraph(""));
    const inCell = caretSelection(caret(0, 0, { row: 0, col: 0, block: 0 }));
    const spec: NewShape = { geometry: "rect", widthCm: 3, heightCm: 2 };
    const { model: next } = insertShape(model, inCell, spec);
    expect(next.sections[0].blocks).toHaveLength(2);
    expect(next.sections[0].blocks[0]).toBe(model.sections[0].blocks[0]);
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph("keep"));
    const snapshot = structuredClone(model);
    insertShape(model, selectionAt(0, 2), { geometry: "rect", widthCm: 3, heightCm: 2 });
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

describe("insertTextBox (AC-9)", () => {
  it("inserts a rect with no fill/outline and an empty text frame", () => {
    const model = doc(paragraph(""));
    const { model: next } = insertTextBox(model, selectionAt(0, 0));
    const inserted = next.sections[0].blocks[0] as Shape;
    expect(inserted.geometry).toBe("rect");
    expect("fill" in inserted).toBe(false);
    expect("outline" in inserted).toBe(false);
    expect(inserted.widthCm).toBe(3);
    expect(inserted.heightCm).toBe(2);
    expect(inserted.text).toEqual({ blocks: [{ type: "paragraph", runs: [] }], vAlign: "top" });
  });
});

// --- defaultClickSizeCm / constrainedDragSizeCm ---------------------------------

describe("defaultClickSizeCm (AC-5, AC-6)", () => {
  it("is 3cm x 2cm for an area geometry", () => {
    expect(defaultClickSizeCm("rect")).toEqual(DEFAULT_AREA_SIZE_CM);
    expect(defaultClickSizeCm("ellipse")).toEqual({ widthCm: 3, heightCm: 2 });
  });

  it("is a horizontal 3cm segment for Line/Arrow", () => {
    expect(defaultClickSizeCm("line")).toEqual(DEFAULT_LINE_SIZE_CM);
    expect(defaultClickSizeCm("arrow")).toEqual({ widthCm: 3, heightCm: 0 });
  });
});

describe("constrainedDragSizeCm (AC-7, AC-8)", () => {
  it("snaps an area geometry to a square on its longer side, regardless of ratio", () => {
    expect(constrainedDragSizeCm("rect", 8, 3)).toEqual({ widthCm: 8, heightCm: 8 });
    expect(constrainedDragSizeCm("ellipse", 2, 6)).toEqual({ widthCm: 6, heightCm: 6 });
  });

  it("uses the absolute value, so a negative drag still produces a positive square", () => {
    expect(constrainedDragSizeCm("rect", -8, 3)).toEqual({ widthCm: 8, heightCm: 8 });
  });

  it("snaps a near-horizontal line drag to purely horizontal", () => {
    expect(constrainedDragSizeCm("line", 10, 1)).toEqual({ widthCm: 10, heightCm: 0 });
  });

  it("snaps a near-vertical line drag to purely vertical", () => {
    expect(constrainedDragSizeCm("arrow", 1, 10)).toEqual({ widthCm: 0, heightCm: 10 });
  });

  it("snaps an exactly-45-degree line drag to a square (diagonal) bounding box", () => {
    expect(constrainedDragSizeCm("line", 6, 6)).toEqual({ widthCm: 6, heightCm: 6 });
  });

  it("snaps a ~30-degree line drag to the nearest 45-degree increment (horizontal)", () => {
    // atan2(3, 10) ~ 16.7deg, closer to 0 than to 45.
    expect(constrainedDragSizeCm("line", 10, 3)).toEqual({ widthCm: 10, heightCm: 0 });
  });

  it("snaps a ~60-degree line drag to the nearest 45-degree increment (diagonal)", () => {
    // atan2(8, 5) ~ 58deg, closer to 45 than to 90.
    expect(constrainedDragSizeCm("line", 5, 8)).toEqual({ widthCm: 8, heightCm: 8 });
  });
});

// --- drawGestureBox ----------------------------------------------------------

describe("drawGestureBox (AC-3..AC-8)", () => {
  const start = { xCm: 4, yCm: 3 };

  it("places the default click size at the pointer while it has barely moved", () => {
    const nudged = { xCm: 4 + DRAW_DRAG_THRESHOLD_CM, yCm: 3 };
    expect(drawGestureBox("rect", start, nudged)).toEqual({
      leftCm: nudged.xCm,
      topCm: 3,
      ...DEFAULT_AREA_SIZE_CM,
    });
    expect(drawGestureBox("line", start, start)).toMatchObject(DEFAULT_LINE_SIZE_CM);
  });

  it("draws the dragged rectangle once past the threshold", () => {
    expect(drawGestureBox("rect", start, { xCm: 9, yCm: 7 })).toEqual({
      leftCm: 4,
      topCm: 3,
      widthCm: 5,
      heightCm: 4,
    });
  });

  it("puts the box's own corner at the pointer when dragging up and left", () => {
    expect(drawGestureBox("rect", start, { xCm: 1, yCm: 1 })).toEqual({
      leftCm: 1,
      topCm: 1,
      widthCm: 3,
      heightCm: 2,
    });
  });

  it("hands Shift-constrained drags to constrainedDragSizeCm", () => {
    const dragged = { xCm: 12, yCm: 6 };
    const box = drawGestureBox("rect", start, dragged, { shiftKey: true });
    expect(box).toMatchObject(constrainedDragSizeCm("rect", dragged.xCm - start.xCm, dragged.yCm - start.yCm));
  });

  it("keeps a text box at its default size however far the pointer is dragged", () => {
    const box = drawGestureBox("rect", start, { xCm: 20, yCm: 20 }, { forceDefaultSize: true });
    expect(box).toMatchObject(DEFAULT_AREA_SIZE_CM);
  });

  it("is the same box the live preview and the insertion each read", () => {
    // Not a behaviour of its own — the point of the function. A drag that
    // ends where the pointer currently is must describe the same box the
    // preview is showing for that pointer, or the shape lands somewhere
    // other than where it was dragged out.
    const pointer = { xCm: 11, yCm: 9 };
    expect(drawGestureBox("ellipse", start, pointer)).toEqual(drawGestureBox("ellipse", start, pointer));
  });
});

// --- setShapeSize ------------------------------------------------------------

describe("setShapeSize (AC-4..AC-8's resize handles / toolbar)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("sets width/height in cm", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapeSize(model, selectionAt(0, 0), location, 8, 4);
    expect(shapeAt(next, location)).toMatchObject({ widthCm: 8, heightCm: 4 });
  });

  it("clamps width below MIN_SHAPE_SIZE_CM", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapeSize(model, selectionAt(0, 0), location, 0.01, 4);
    expect(shapeAt(next, location)?.widthCm).toBe(MIN_SHAPE_SIZE_CM);
  });

  it("clamps height below MIN_SHAPE_SIZE_CM for an area geometry", () => {
    const model = doc(shape({ geometry: "rect" }), paragraph(""));
    const { model: next } = setShapeSize(model, selectionAt(0, 0), location, 4, 0.01);
    expect(shapeAt(next, location)?.heightCm).toBe(MIN_SHAPE_SIZE_CM);
  });

  it("allows a zero (or negative, clamped to zero) height for a Line — a real horizontal line, not a collapsed shape", () => {
    const model = doc(shape({ geometry: "line", heightCm: 0 }), paragraph(""));
    const { model: next } = setShapeSize(model, selectionAt(0, 0), location, 5, -3);
    expect(shapeAt(next, location)?.heightCm).toBe(0);
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"));
    const { model: next } = setShapeSize(model, selectionAt(0, 0), location, 8, 4);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });
});

describe("setShapePath (SPEC-071 AC-22/AC-26)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };
  const path: Shape["path"] = {
    subpaths: [{ steps: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 1, y: 1 }], closed: false }],
  };

  it("writes the edited path and its renormalised box onto the shape", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapePath(model, selectionAt(0, 0), location, path, { widthCm: 5, heightCm: 4 });
    expect(shapeAt(next, location)).toMatchObject({ path, widthCm: 5, heightCm: 4, geometry: "rect" });
  });

  it("drops a stale pathSource so a same-format save serialises the edited path, not the original geometry", () => {
    const model = doc(
      shape({ pathSource: { format: "docx", xml: "<a:custGeom/>" } }),
      paragraph(""),
    );
    const { model: next } = setShapePath(model, selectionAt(0, 0), location, path, { widthCm: 5, heightCm: 4 });
    expect(shapeAt(next, location)?.pathSource).toBeUndefined();
  });
});

// --- setShapePosition (AC-42's numeric position dialog) ----------------------

describe("setShapePosition (SPEC-068 AC-42)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("sets an absolute offset for a floating shape", () => {
    const model = doc(shape({ wrap: "square", offsetXCm: 1, offsetYCm: 2 }), paragraph(""));
    const { model: next } = setShapePosition(model, selectionAt(0, 0), location, 5.5, 7.25);
    expect(shapeAt(next, location)).toMatchObject({ offsetXCm: 5.5, offsetYCm: 7.25 });
  });

  it("replaces (does not add to) the current offset", () => {
    const model = doc(shape({ wrap: "square", offsetXCm: 10, offsetYCm: 10 }), paragraph(""));
    const { model: next } = setShapePosition(model, selectionAt(0, 0), location, 0, 0);
    expect(shapeAt(next, location)).toMatchObject({ offsetXCm: 0, offsetYCm: 0 });
  });

  it("is a no-op for an inline shape (the wrap mode ignores offsets)", () => {
    const model = doc(shape({ wrap: "inline", offsetXCm: 1, offsetYCm: 2 }), paragraph(""));
    const { model: next } = setShapePosition(model, selectionAt(0, 0), location, 3, 3);
    // The offset is left exactly as it was — the dialog's position fields are
    // disabled for an inline shape, and the setter refuses to write them.
    expect(shapeAt(next, location)).toMatchObject({ offsetXCm: 1, offsetYCm: 2 });
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"));
    const { model: next } = setShapePosition(model, selectionAt(0, 0), location, 1, 1);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });
});

// --- moveShape / nudgeShape ---------------------------------------------------

describe("moveShape", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("adds a page-space delta to the shape's own offset", () => {
    const model = doc(shape({ offsetXCm: 1, offsetYCm: 2 }), paragraph(""));
    const { model: next } = moveShape(model, selectionAt(0, 0), location, 0.5, -0.25);
    expect(shapeAt(next, location)).toMatchObject({ offsetXCm: 1.5, offsetYCm: 1.75 });
  });

  it("defaults absent offsets to the shared anchored default before adding the delta", () => {
    const model = doc(shape({ offsetXCm: undefined, offsetYCm: undefined }), paragraph(""));
    const { model: next } = moveShape(model, selectionAt(0, 0), location, 1, 1);
    expect(shapeAt(next, location)).toMatchObject({ offsetXCm: 1, offsetYCm: 1 });
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"));
    const { model: next } = moveShape(model, selectionAt(0, 0), location, 1, 1);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });
});

describe("nudgeShape (AC-40, AC-41)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("moves by exactly NUDGE_PT_CM (1pt) for a plain arrow key", () => {
    const model = doc(shape({ offsetXCm: 0, offsetYCm: 0 }), paragraph(""));
    const { model: next } = nudgeShape(model, selectionAt(0, 0), location, "right", false);
    expect(shapeAt(next, location)?.offsetXCm).toBeCloseTo(NUDGE_PT_CM, 10);
  });

  it("moves by exactly NUDGE_FINE_PT_CM (0.1pt) for Ctrl/Cmd+arrow", () => {
    const model = doc(shape({ offsetXCm: 0, offsetYCm: 0 }), paragraph(""));
    const { model: next } = nudgeShape(model, selectionAt(0, 0), location, "right", true);
    expect(shapeAt(next, location)?.offsetXCm).toBeCloseTo(NUDGE_FINE_PT_CM, 10);
  });

  it("NUDGE_FINE_PT_CM is a tenth of NUDGE_PT_CM", () => {
    expect(NUDGE_FINE_PT_CM).toBeCloseTo(NUDGE_PT_CM / 10, 12);
  });

  it("moves in the correct sign per direction", () => {
    const model = doc(shape({ offsetXCm: 5, offsetYCm: 5 }), paragraph(""));
    expect(shapeAt(nudgeShape(model, selectionAt(0, 0), location, "up", false).model, location)?.offsetYCm).toBeLessThan(5);
    expect(shapeAt(nudgeShape(model, selectionAt(0, 0), location, "down", false).model, location)?.offsetYCm).toBeGreaterThan(5);
    expect(shapeAt(nudgeShape(model, selectionAt(0, 0), location, "left", false).model, location)?.offsetXCm).toBeLessThan(5);
    expect(shapeAt(nudgeShape(model, selectionAt(0, 0), location, "right", false).model, location)?.offsetXCm).toBeGreaterThan(5);
  });
});

// --- setShapeFill / setShapeOutline ---------------------------------------------

describe("setShapeFill", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("sets a solid fill colour", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapeFill(model, selectionAt(0, 0), location, { kind: "solid", color: "#00ff00" });
    expect(shapeAt(next, location)?.fill).toEqual({ kind: "solid", color: "#00ff00" });
  });

  it("undefined clears the fill (no key at all)", () => {
    const model = doc(shape({ fill: { kind: "solid", color: "#00ff00" } }), paragraph(""));
    const { model: next } = setShapeFill(model, selectionAt(0, 0), location, undefined);
    expect(shapeAt(next, location)?.fill).toBeUndefined();
    expect("fill" in shapeAt(next, location)!).toBe(false);
  });
});

describe("setShapeOutline", () => {
  const location: ShapeLocation = { section: 0, block: 0 };
  const outline = { color: "#000000", widthPt: 1.5, dash: "dash" as const };

  it("sets an outline", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapeOutline(model, selectionAt(0, 0), location, outline);
    expect(shapeAt(next, location)?.outline).toEqual(outline);
  });

  it("undefined clears the outline (no key at all)", () => {
    const model = doc(shape({ outline }), paragraph(""));
    const { model: next } = setShapeOutline(model, selectionAt(0, 0), location, undefined);
    expect("outline" in shapeAt(next, location)!).toBe(false);
  });
});

// --- SPEC-073 AC-36: style edits clear the preserved slot they replace ----------

describe("style edits over preserved unparsed styling (AC-36 one-source)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };
  const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
  // A .docx-flavour preserved set: a scheme-colour fill, a gradient-stroke
  // a:ln, an unmodelled glow effect and an unparseable outerShdw.
  const preserved =
    `<a:solidFill xmlns:a="${A_NS}"><a:schemeClr val="accent1"/></a:solidFill>` +
    `<a:ln xmlns:a="${A_NS}" w="25400"><a:gradFill><a:gsLst/></a:gradFill></a:ln>` +
    `<a:glow xmlns:a="${A_NS}" rad="63500"><a:srgbClr val="00FF00"/></a:glow>` +
    `<a:outerShdw xmlns:a="${A_NS}" blurRad="9"><a:schemeClr val="accent2"/></a:outerShdw>`;

  it("setShapeFill clears the preserved fill slot so the edit cannot be resurrected on save", () => {
    const model = doc(shape({ unsupportedStyleXml: preserved }), paragraph(""));
    const { model: next } = setShapeFill(model, selectionAt(0, 0), location, { kind: "solid", color: "#ff0000" });
    const kept = shapeAt(next, location)!.unsupportedStyleXml!;
    expect(kept).not.toContain("solidFill");
    expect(kept).toContain("a:ln");
    expect(kept).toContain("a:glow");
  });

  it("setShapeFill to NONE also clears the preserved fill slot (undefined fill must not resurrect it)", () => {
    const model = doc(shape({ unsupportedStyleXml: preserved }), paragraph(""));
    const { model: next } = setShapeFill(model, selectionAt(0, 0), location, undefined);
    expect(shapeAt(next, location)!.unsupportedStyleXml).not.toContain("solidFill");
  });

  it("setShapeOutline clears the preserved a:ln slot only", () => {
    const model = doc(shape({ unsupportedStyleXml: preserved }), paragraph(""));
    const { model: next } = setShapeOutline(model, selectionAt(0, 0), location, { color: "#000000", widthPt: 1, dash: "solid" });
    const kept = shapeAt(next, location)!.unsupportedStyleXml!;
    expect(kept).not.toContain("<a:ln");
    expect(kept).toContain("solidFill");
    expect(kept).toContain("a:glow");
  });

  it("setShapeShadow clears a preserved outerShdw but keeps the glow", () => {
    const model = doc(shape({ unsupportedStyleXml: preserved }), paragraph(""));
    const { model: next } = setShapeShadow(model, selectionAt(0, 0), location, { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 });
    const kept = shapeAt(next, location)!.unsupportedStyleXml!;
    expect(kept).not.toContain("outerShdw");
    expect(kept).toContain("a:glow");
  });

  it("formatDrawings clears exactly the slots its patch applies, per target", () => {
    const model = doc(shape({ unsupportedStyleXml: preserved }), paragraph(""));
    const { model: next } = formatDrawings(model, selectionAt(1, 0), [location], { fill: { kind: "solid", color: "#ff0000" } });
    const kept = shapeAt(next, location)!.unsupportedStyleXml!;
    expect(kept).not.toContain("solidFill");
    // Slots the patch does not touch survive untouched.
    expect(kept).toContain("<a:ln");
    expect(kept).toContain("a:glow");
    expect(kept).toContain("outerShdw");
  });

  it("the whole preserved field is deleted when every element sat in a cleared slot", () => {
    const fillOnly = `<a:solidFill xmlns:a="${A_NS}"><a:schemeClr val="accent1"/></a:solidFill>`;
    const model = doc(shape({ unsupportedStyleXml: fillOnly }), paragraph(""));
    const { model: next } = setShapeFill(model, selectionAt(0, 0), location, { kind: "solid", color: "#ff0000" });
    expect("unsupportedStyleXml" in shapeAt(next, location)!).toBe(false);
  });
});

// --- normalizeRotationDeg / setShapeRotation ------------------------------------

describe("normalizeRotationDeg (AC-16)", () => {
  it("rounds to the nearest whole degree", () => {
    expect(normalizeRotationDeg(45.4)).toBe(45);
    expect(normalizeRotationDeg(45.6)).toBe(46);
  });

  it("wraps a negative value into 0..359", () => {
    expect(normalizeRotationDeg(-10)).toBe(350);
    expect(normalizeRotationDeg(-370)).toBe(350);
  });

  it("wraps a value >= 360 back into range", () => {
    expect(normalizeRotationDeg(370)).toBe(10);
    expect(normalizeRotationDeg(360)).toBe(0);
    expect(normalizeRotationDeg(720)).toBe(0);
  });

  it("leaves an in-range value untouched", () => {
    expect(normalizeRotationDeg(0)).toBe(0);
    expect(normalizeRotationDeg(90)).toBe(90);
    expect(normalizeRotationDeg(359)).toBe(359);
  });
});

describe("setShapeRotation (AC-16)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("sets a normalized whole-degree rotation", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapeRotation(model, selectionAt(0, 0), location, 45.6);
    expect(shapeAt(next, location)?.rotationDeg).toBe(46);
  });

  it("a rotation that normalizes to 0 clears the field entirely, rather than storing 0", () => {
    const model = doc(shape({ rotationDeg: 45 }), paragraph(""));
    const { model: next } = setShapeRotation(model, selectionAt(0, 0), location, 360);
    expect("rotationDeg" in shapeAt(next, location)!).toBe(false);
  });

  it("wraps a negative or overshooting input rather than clamping", () => {
    const model = doc(shape(), paragraph(""));
    expect(setShapeRotation(model, selectionAt(0, 0), location, -30).model.sections[0].blocks[0]).toMatchObject({ rotationDeg: 330 });
    expect(setShapeRotation(model, selectionAt(0, 0), location, 400).model.sections[0].blocks[0]).toMatchObject({ rotationDeg: 40 });
  });
});

// --- setShapeWrap --------------------------------------------------------------

describe("setShapeWrap (AC-25..AC-28)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("switching to inline strips offsets and relativeTo fields entirely", () => {
    const model = doc(shape({ wrap: "square", offsetXCm: 3, offsetYCm: 4 }), paragraph(""));
    const { model: next } = setShapeWrap(model, selectionAt(0, 0), location, "inline");
    const updated = shapeAt(next, location)!;
    expect(updated.wrap).toBe("inline");
    expect("offsetXCm" in updated).toBe(false);
    expect("offsetYCm" in updated).toBe(false);
    expect("relativeToH" in updated).toBe(false);
    expect("relativeToV" in updated).toBe(false);
  });

  it("switching from inline (no offsets) to a floating mode fills in the shared default position", () => {
    const model = doc(
      { ...shape({ wrap: "inline" }), offsetXCm: undefined, offsetYCm: undefined, relativeToH: undefined, relativeToV: undefined },
      paragraph(""),
    );
    const { model: next } = setShapeWrap(model, selectionAt(0, 0), location, "square");
    const updated = shapeAt(next, location)!;
    expect(updated.offsetXCm).toBe(DEFAULT_DRAWING_POSITION.offsetXCm);
    expect(updated.offsetYCm).toBe(DEFAULT_DRAWING_POSITION.offsetYCm);
    expect(updated.relativeToH).toBe(DEFAULT_DRAWING_POSITION.relativeToH);
    expect(updated.relativeToV).toBe(DEFAULT_DRAWING_POSITION.relativeToV);
  });

  it("switching between two floating modes keeps an existing offset rather than re-defaulting it", () => {
    const model = doc(shape({ wrap: "square", offsetXCm: 7, offsetYCm: 9 }), paragraph(""));
    const { model: next } = setShapeWrap(model, selectionAt(0, 0), location, "behind");
    const updated = shapeAt(next, location)!;
    expect(updated.wrap).toBe("behind");
    expect(updated.offsetXCm).toBe(7);
    expect(updated.offsetYCm).toBe(9);
  });
});

// --- setShapeAltText -----------------------------------------------------------

describe("setShapeAltText (AC-46, AC-47)", () => {
  const location: ShapeLocation = { section: 0, block: 0 };

  it("sets alt text", () => {
    const model = doc(shape(), paragraph(""));
    const { model: next } = setShapeAltText(model, selectionAt(0, 0), location, "A callout box");
    expect(shapeAt(next, location)?.altText).toBe("A callout box");
  });

  it("an empty string clears alt text", () => {
    const model = doc(shape({ altText: "old" }), paragraph(""));
    const { model: next } = setShapeAltText(model, selectionAt(0, 0), location, "");
    expect("altText" in shapeAt(next, location)!).toBe(false);
  });
});

// --- deleteShape ---------------------------------------------------------------

describe("deleteShape (AC-42)", () => {
  it("removes the shape block and lands the caret where it was", () => {
    const model = doc(paragraph("before"), shape(), paragraph("after"));
    const location: ShapeLocation = { section: 0, block: 1 };
    const { model: next, selection } = deleteShape(model, selectionAt(1, 0), location);
    expect(next.sections[0].blocks.map(blockText)).toEqual(["before", "after"]);
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("leaves one empty paragraph when deleting the only block", () => {
    const model = doc(shape());
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next, selection } = deleteShape(model, selectionAt(0, 0), location);
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("");
    expect(selection.focus).toEqual(caret(0, 0));
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = deleteShape(model, selectionAt(0, 0), location);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph("before"), shape(), paragraph("after"));
    const snapshot = structuredClone(model);
    deleteShape(model, selectionAt(1, 0), { section: 0, block: 1 });
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

// --- duplicateShape --------------------------------------------------------------

describe("duplicateShape (AC-43)", () => {
  it("inserts a copy offset 0.5cm right and down, immediately after the source", () => {
    const model = doc(shape({ offsetXCm: 1, offsetYCm: 1 }), paragraph(""));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next, selection } = duplicateShape(model, selectionAt(0, 0), location);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(isShape(blocks[0])).toBe(true);
    const copy = blocks[1] as Shape;
    expect(copy.offsetXCm).toBeCloseTo(1 + DUPLICATE_OFFSET_CM, 10);
    expect(copy.offsetYCm).toBeCloseTo(1 + DUPLICATE_OFFSET_CM, 10);
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("bumps the copy's zOrder one past the highest zOrder among every drawing in the document", () => {
    const model = doc(shape({ zOrder: 2 }), image({ zOrder: 5 }), paragraph(""));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = duplicateShape(model, selectionAt(0, 0), location);
    const copy = next.sections[0].blocks[1] as Shape;
    expect(copy.zOrder).toBe(6);
  });

  it("clears groupOrdinal on the copy even when the source was grouped", () => {
    const model = doc(shape({ groupOrdinal: 3 }), paragraph(""));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = duplicateShape(model, selectionAt(0, 0), location);
    const copy = next.sections[0].blocks[1] as Shape;
    expect("groupOrdinal" in copy).toBe(false);
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = duplicateShape(model, selectionAt(0, 0), location);
    expect(next.sections[0].blocks).toHaveLength(1);
  });
});

// --- reorderShapeZ (AC-29..AC-33) -----------------------------------------------

describe("reorderShapeZ", () => {
  function threeShapes(): DocumentModel {
    // zOrder 0, 1, 2 at blocks 0, 1, 2.
    return doc(shape({ zOrder: 0 }), shape({ zOrder: 1 }), shape({ zOrder: 2 }), paragraph(""));
  }

  function zOrders(model: DocumentModel): number[] {
    return model.sections[0].blocks.filter(isShape).map((s) => s.zOrder!);
  }

  it("bring to front: moves the target to the highest zOrder, others shift down", () => {
    const model = threeShapes();
    const location: ShapeLocation = { section: 0, block: 0 }; // zOrder 0
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "front");
    const blocks = next.sections[0].blocks;
    // The originally-lowest shape now paints last (topmost).
    expect((blocks[0] as Shape).zOrder).toBe(2);
    expect((blocks[1] as Shape).zOrder).toBe(0);
    expect((blocks[2] as Shape).zOrder).toBe(1);
    expect(zOrders(next).sort()).toEqual([0, 1, 2]);
  });

  it("send to back: moves the target to the lowest zOrder", () => {
    const model = threeShapes();
    const location: ShapeLocation = { section: 0, block: 2 }; // zOrder 2
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "back");
    const blocks = next.sections[0].blocks;
    expect((blocks[2] as Shape).zOrder).toBe(0);
    expect((blocks[0] as Shape).zOrder).toBe(1);
    expect((blocks[1] as Shape).zOrder).toBe(2);
  });

  it("bring forward: swaps with the next-higher neighbour by one step", () => {
    const model = threeShapes();
    const location: ShapeLocation = { section: 0, block: 0 }; // zOrder 0
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "forward");
    const blocks = next.sections[0].blocks;
    expect((blocks[0] as Shape).zOrder).toBe(1);
    expect((blocks[1] as Shape).zOrder).toBe(0);
    expect((blocks[2] as Shape).zOrder).toBe(2);
  });

  it("send backward: swaps with the next-lower neighbour by one step", () => {
    const model = threeShapes();
    const location: ShapeLocation = { section: 0, block: 2 }; // zOrder 2
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "backward");
    const blocks = next.sections[0].blocks;
    expect((blocks[2] as Shape).zOrder).toBe(1);
    expect((blocks[1] as Shape).zOrder).toBe(2);
  });

  it("bring forward at the top and send backward at the bottom are no-ops on order", () => {
    const model = threeShapes();
    const topLocation: ShapeLocation = { section: 0, block: 2 };
    const { model: forwardResult } = reorderShapeZ(model, selectionAt(0, 0), topLocation, "forward");
    expect(zOrders(forwardResult)).toEqual([0, 1, 2]);

    const bottomLocation: ShapeLocation = { section: 0, block: 0 };
    const { model: backwardResult } = reorderShapeZ(model, selectionAt(0, 0), bottomLocation, "backward");
    expect(zOrders(backwardResult)).toEqual([0, 1, 2]);
  });

  it("reorders against Images too — every drawing shares one z-order space", () => {
    const model = doc(shape({ zOrder: 0 }), image({ zOrder: 1 }), paragraph(""));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "front");
    expect((next.sections[0].blocks[0] as Shape).zOrder).toBe(1);
    expect((next.sections[0].blocks[1] as Image).zOrder).toBe(0);
  });

  it("every drawing gets an explicit zOrder afterward (no more absent-means-0 ties)", () => {
    // Neither shape carries an explicit zOrder up front.
    const model = doc(shape(), shape(), paragraph(""));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "front");
    for (const block of next.sections[0].blocks.filter(isShape)) {
      expect(block.zOrder).toBeDefined();
    }
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"));
    const location: ShapeLocation = { section: 0, block: 0 };
    const { model: next } = reorderShapeZ(model, selectionAt(0, 0), location, "front");
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
  });
});

// --- grouping (AC-34..AC-39) -------------------------------------------------

describe("freshGroupOrdinal (AC-35)", () => {
  it("is 1 for a document with no grouped drawings", () => {
    expect(freshGroupOrdinal(doc(shape(), paragraph("")))).toBe(1);
  });

  it("is one past the highest existing groupOrdinal", () => {
    const model = doc(shape({ groupOrdinal: 3 }), image({ groupOrdinal: 7 }), shape({ groupOrdinal: 2 }));
    expect(freshGroupOrdinal(model)).toBe(8);
  });

  it("ignores drawings with no groupOrdinal", () => {
    const model = doc(shape(), shape());
    expect(freshGroupOrdinal(model)).toBe(1);
  });
});

describe("groupShapes (AC-34..AC-36)", () => {
  it("assigns every named location the same fresh groupOrdinal", () => {
    const model = doc(shape(), shape(), paragraph(""));
    const locations: ShapeLocation[] = [{ section: 0, block: 0 }, { section: 0, block: 1 }];
    const { model: next } = groupShapes(model, selectionAt(0, 0), locations);
    const a = shapeAt(next, locations[0])!;
    const b = shapeAt(next, locations[1])!;
    expect(a.groupOrdinal).toBeDefined();
    expect(a.groupOrdinal).toBe(b.groupOrdinal);
  });

  it("converts an inline member to square at the shared default position", () => {
    const model = doc(shape({ wrap: "inline", offsetXCm: undefined, offsetYCm: undefined }), shape(), paragraph(""));
    const locations: ShapeLocation[] = [{ section: 0, block: 0 }, { section: 0, block: 1 }];
    const { model: next } = groupShapes(model, selectionAt(0, 0), locations);
    const converted = shapeAt(next, locations[0])!;
    expect(converted.wrap).toBe("square");
    expect(converted.offsetXCm).toBe(DEFAULT_DRAWING_POSITION.offsetXCm);
    expect(converted.offsetYCm).toBe(DEFAULT_DRAWING_POSITION.offsetYCm);
    expect(converted.relativeToH).toBe(DEFAULT_DRAWING_POSITION.relativeToH);
    expect(converted.relativeToV).toBe(DEFAULT_DRAWING_POSITION.relativeToV);
  });

  it("skips a location that isn't a shape or image", () => {
    const model = doc(paragraph("text"), shape());
    const locations: ShapeLocation[] = [{ section: 0, block: 0 }, { section: 0, block: 1 }];
    const { model: next } = groupShapes(model, selectionAt(0, 0), locations);
    expect(blockText(next.sections[0].blocks[0])).toBe("text");
    expect((next.sections[0].blocks[1] as Shape).groupOrdinal).toBeDefined();
  });

  it("groups a Shape and an Image together", () => {
    const model = doc(shape(), image());
    const locations: ShapeLocation[] = [{ section: 0, block: 0 }, { section: 0, block: 1 }];
    const { model: next } = groupShapes(model, selectionAt(0, 0), locations);
    const a = (next.sections[0].blocks[0] as Shape).groupOrdinal;
    const b = (next.sections[0].blocks[1] as Image).groupOrdinal;
    expect(a).toBe(b);
  });
});

describe("ungroupShapes (AC-39)", () => {
  it("clears groupOrdinal on every member sharing it, document-wide, keeping position/size/rotation", () => {
    const model = doc(
      shape({ groupOrdinal: 5, widthCm: 4, heightCm: 2, rotationDeg: 30 }),
      shape({ groupOrdinal: 5, widthCm: 6, heightCm: 3 }),
      shape({ groupOrdinal: 9 }),
    );
    const { model: next } = ungroupShapes(model, selectionAt(0, 0), 5);
    const blocks = next.sections[0].blocks as Shape[];
    expect("groupOrdinal" in blocks[0]).toBe(false);
    expect("groupOrdinal" in blocks[1]).toBe(false);
    expect(blocks[2].groupOrdinal).toBe(9); // untouched — different group
    expect(blocks[0].widthCm).toBe(4);
    expect(blocks[0].rotationDeg).toBe(30);
  });
});

describe("groupMembers", () => {
  it("returns every drawing sharing the ordinal, in document order", () => {
    const model = doc(shape({ groupOrdinal: 1 }), paragraph(""), image({ groupOrdinal: 1 }), shape({ groupOrdinal: 2 }));
    const members = groupMembers(model, 1);
    expect(members).toHaveLength(2);
    expect(members[0].location).toEqual({ section: 0, block: 0 });
    expect(members[1].location).toEqual({ section: 0, block: 2 });
  });

  it("is empty for an ordinal no drawing carries", () => {
    expect(groupMembers(doc(shape({ groupOrdinal: 1 })), 42)).toEqual([]);
  });
});

// --- Grouping — move, resize, rotate --------------------------------------------

describe("moveGroup (AC-36's 'a group moves as one unit')", () => {
  it("adds the same delta to every member's offset; sizes untouched", () => {
    const model = doc(
      shape({ groupOrdinal: 1, offsetXCm: 1, offsetYCm: 1, widthCm: 3, heightCm: 2 }),
      shape({ groupOrdinal: 1, offsetXCm: 5, offsetYCm: 2, widthCm: 4, heightCm: 4 }),
    );
    const { model: next } = moveGroup(model, selectionAt(0, 0), 1, 2, -1);
    const blocks = next.sections[0].blocks as Shape[];
    expect(blocks[0].offsetXCm).toBeCloseTo(3, 10);
    expect(blocks[0].offsetYCm).toBeCloseTo(0, 10);
    expect(blocks[1].offsetXCm).toBeCloseTo(7, 10);
    expect(blocks[1].offsetYCm).toBeCloseTo(1, 10);
    expect(blocks[0].widthCm).toBe(3);
    expect(blocks[1].widthCm).toBe(4);
  });
});

describe("groupBoundingBoxCm (AC-37's resize pivot)", () => {
  it("spans the union of every member's box", () => {
    const model = doc(
      shape({ groupOrdinal: 1, offsetXCm: 1, offsetYCm: 1, widthCm: 3, heightCm: 2 }),
      shape({ groupOrdinal: 1, offsetXCm: 5, offsetYCm: 0, widthCm: 2, heightCm: 4 }),
    );
    expect(groupBoundingBoxCm(model, 1)).toEqual({ leftCm: 1, topCm: 0, widthCm: 6, heightCm: 4 });
  });

  it("is null for an ordinal with no members", () => {
    expect(groupBoundingBoxCm(doc(shape({ groupOrdinal: 1 })), 42)).toBeNull();
  });
});

describe("resizeGroup (AC-37 — 'Grouping — resize' formulas)", () => {
  it("scales every member's width/height/offsets by the per-axis factor", () => {
    const model = doc(
      shape({ groupOrdinal: 1, widthCm: 4, heightCm: 2, offsetXCm: 1, offsetYCm: 1 }),
      shape({ groupOrdinal: 1, widthCm: 2, heightCm: 3, offsetXCm: 5, offsetYCm: 2 }),
    );
    const { model: next } = resizeGroup(model, selectionAt(0, 0), 1, 2, 0.5);
    const blocks = next.sections[0].blocks as Shape[];
    expect(blocks[0]).toMatchObject({ widthCm: 8, heightCm: 1, offsetXCm: 2, offsetYCm: 0.5 });
    expect(blocks[1]).toMatchObject({ widthCm: 4, heightCm: 1.5, offsetXCm: 10, offsetYCm: 1 });
  });

  it("leaves member rotationDeg unchanged", () => {
    const model = doc(shape({ groupOrdinal: 1, rotationDeg: 40 }));
    const { model: next } = resizeGroup(model, selectionAt(0, 0), 1, 2, 2);
    expect((next.sections[0].blocks[0] as Shape).rotationDeg).toBe(40);
  });

  it("uniform (Shift) resize is just sx === sy", () => {
    const model = doc(shape({ groupOrdinal: 1, widthCm: 4, heightCm: 2, offsetXCm: 2, offsetYCm: 2 }));
    const { model: next } = resizeGroup(model, selectionAt(0, 0), 1, 1.5, 1.5);
    expect(next.sections[0].blocks[0]).toMatchObject({ widthCm: 6, heightCm: 3, offsetXCm: 3, offsetYCm: 3 });
  });

  it("clamps width to MIN_SHAPE_SIZE_CM but only floors height at 0", () => {
    const model = doc(shape({ groupOrdinal: 1, widthCm: 4, heightCm: 2 }));
    const { model: next } = resizeGroup(model, selectionAt(0, 0), 1, 0.001, -1);
    const updated = next.sections[0].blocks[0] as Shape;
    expect(updated.widthCm).toBe(MIN_SHAPE_SIZE_CM);
    expect(updated.heightCm).toBe(0);
  });
});

describe("rotateGroup (AC-38)", () => {
  it("adds delta to every member's own rotationDeg", () => {
    const model = doc(shape({ groupOrdinal: 1, rotationDeg: 10 }), shape({ groupOrdinal: 1 }));
    const { model: next } = rotateGroup(model, selectionAt(0, 0), 1, 20);
    const blocks = next.sections[0].blocks as Shape[];
    expect(blocks[0].rotationDeg).toBe(30);
    expect(blocks[1].rotationDeg).toBe(20);
  });

  it("normalizes and wraps around 360", () => {
    const model = doc(shape({ groupOrdinal: 1, rotationDeg: 350 }));
    const { model: next } = rotateGroup(model, selectionAt(0, 0), 1, 20);
    expect((next.sections[0].blocks[0] as Shape).rotationDeg).toBe(10);
  });

  it("a delta that normalizes to exactly 0 clears rotationDeg rather than storing 0", () => {
    const model = doc(shape({ groupOrdinal: 1, rotationDeg: 40 }));
    const { model: next } = rotateGroup(model, selectionAt(0, 0), 1, -40);
    expect("rotationDeg" in (next.sections[0].blocks[0] as Shape)).toBe(false);
  });
});

// --- Re-anchoring: "the nearest paragraph" --------------------------------------

function span(section: number, block: number, topCm: number, bottomCm: number): ParagraphSpan {
  return { section, block, topCm, bottomCm };
}

describe("nearestParagraph", () => {
  it("scores 0 (and wins) when the centre falls inside the candidate's span", () => {
    const candidates = [span(0, 0, 0, 5), span(0, 1, 5, 10)];
    expect(nearestParagraph(7, candidates)).toEqual(candidates[1]);
  });

  it("picks the candidate whose nearer edge is closest when cy falls in a gap", () => {
    const candidates = [span(0, 0, 0, 4), span(0, 1, 10, 14)];
    // cy=6: distance to first candidate's bottom (4) is 2; distance to
    // second's top (10) is 4 — the first candidate wins.
    expect(nearestParagraph(6, candidates)).toEqual(candidates[0]);
  });

  it("ties go to the earlier paragraph in document order", () => {
    // cy=5 is equidistant (score 0? no — exactly between two disjoint spans)
    // from both candidates' nearer edges: distance 1 to each.
    const candidates = [span(0, 0, 0, 4), span(0, 1, 6, 10)];
    expect(nearestParagraph(5, candidates)).toEqual(candidates[0]);
  });

  it("returns null for an empty candidate list", () => {
    expect(nearestParagraph(5, [])).toBeNull();
  });
});

describe("reanchorShape (AC-28)", () => {
  it("is a no-op with no candidates", () => {
    const model = doc(paragraph("p0"), shape(), paragraph("p1"));
    const location: ShapeLocation = { section: 0, block: 1 };
    const { model: next } = reanchorShape(model, selectionAt(1, 0), location, [], 0, 0);
    expect(equalStructural(model, next)).toBe(true);
  });

  it("is a no-op when the nearest candidate is the shape's current anchor location", () => {
    const model = doc(paragraph("p0"), shape(), paragraph("p1"));
    const location: ShapeLocation = { section: 0, block: 1 };
    const candidates: ParagraphSpan[] = [{ section: 0, block: 1, topCm: 0, bottomCm: 5 }];
    const { model: next } = reanchorShape(model, selectionAt(1, 0), location, candidates, 2, 0);
    expect(equalStructural(model, next)).toBe(true);
  });

  it("moves the shape to sit after the nearest candidate paragraph and adjusts offsetYCm by the anchor-top delta", () => {
    // [P0, Shape, P1, P2] — shape currently anchored near P0 (block 1).
    const model = doc(paragraph("P0"), shape({ offsetXCm: 1, offsetYCm: 2 }), paragraph("P1"), paragraph("P2"));
    const location: ShapeLocation = { section: 0, block: 1 };
    const candidates: ParagraphSpan[] = [
      { section: 0, block: 0, topCm: 0, bottomCm: 1 },
      { section: 0, block: 2, topCm: 10, bottomCm: 11 },
      { section: 0, block: 3, topCm: 20, bottomCm: 21 },
    ];
    // Dragged so its centre (cy=20.5) is nearest P2 (block 3, top 20).
    const oldAnchorTopCm = 0; // P0's own top, the shape's old anchor.
    const { model: next, selection } = reanchorShape(model, selectionAt(1, 0), location, candidates, 20.5, oldAnchorTopCm);
    const blocks = next.sections[0].blocks;
    // Shape removed from index 1, so P1/P2 shift down by one; it lands
    // immediately after P2 (now at index 2) — end of the list.
    expect(blocks.map((b) => (isShape(b) ? "SHAPE" : blockText(b)))).toEqual(["P0", "P1", "P2", "SHAPE"]);
    const moved = blocks[3] as Shape;
    // offsetYCm adjusted by (new top - old top) = 20 - 0 = 20, subtracted.
    expect(moved.offsetYCm).toBeCloseTo(2 - 20, 10);
    expect(moved.offsetXCm).toBe(1); // horizontal position is untouched.
    expect(selection.focus).toEqual(caret(4, 0));
  });

  it("relativeToH/relativeToV are unchanged by re-anchoring", () => {
    const model = doc(paragraph("P0"), shape({ relativeToH: "margin", relativeToV: "page" }), paragraph("P1"));
    const location: ShapeLocation = { section: 0, block: 1 };
    const candidates: ParagraphSpan[] = [{ section: 0, block: 2, topCm: 5, bottomCm: 6 }];
    const { model: next } = reanchorShape(model, selectionAt(1, 0), location, candidates, 5.5, 0);
    const moved = next.sections[0].blocks.find(isShape)!;
    expect(moved.relativeToH).toBe("margin");
    expect(moved.relativeToV).toBe("page");
  });

  it("is a no-op when the location doesn't address a shape", () => {
    const model = doc(paragraph("text"), paragraph("p1"));
    const location: ShapeLocation = { section: 0, block: 0 };
    const candidates: ParagraphSpan[] = [{ section: 0, block: 1, topCm: 5, bottomCm: 6 }];
    const { model: next } = reanchorShape(model, selectionAt(0, 0), location, candidates, 5.5, 0);
    expect(equalStructural(model, next)).toBe(true);
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph("P0"), shape(), paragraph("P1"), paragraph("P2"));
    const snapshot = structuredClone(model);
    const candidates: ParagraphSpan[] = [{ section: 0, block: 3, topCm: 20, bottomCm: 21 }];
    reanchorShape(model, selectionAt(1, 0), { section: 0, block: 1 }, candidates, 20.5, 0);
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

// --- shapeAt / isImage sanity ----------------------------------------------------

describe("shapeAt", () => {
  it("returns the Shape at a location, or null for anything else", () => {
    const model = doc(shape(), image(), paragraph(""));
    expect(shapeAt(model, { section: 0, block: 0 })).not.toBeNull();
    expect(shapeAt(model, { section: 0, block: 1 })).toBeNull();
    expect(shapeAt(model, { section: 0, block: 2 })).toBeNull();
    expect(isImage(model.sections[0].blocks[1])).toBe(true);
  });
});

// --- shape text editing (SPEC-051 AC-17, AC-18) ----------------------------

describe("ensureShapeTextFrame", () => {
  it("gives a textless shape one empty paragraph", () => {
    const model = doc(shape());
    const { model: next } = ensureShapeTextFrame(model, selectionAt(0, 0), { section: 0, block: 0 });
    expect(shapeAt(next, { section: 0, block: 0 })?.text).toEqual({
      blocks: [{ type: "paragraph", runs: [] }],
      vAlign: "top",
    });
  });

  it("is a no-op for a shape that already has a text frame", () => {
    const withText = shape({ text: { blocks: [paragraph("hi")], vAlign: "middle" } });
    const model = doc(withText);
    const { model: next } = ensureShapeTextFrame(model, selectionAt(0, 0), { section: 0, block: 0 });
    expect(shapeAt(next, { section: 0, block: 0 })?.text).toEqual(withText.text);
  });

  it("never mutates the input model", () => {
    const model = doc(shape());
    const snapshot = structuredClone(model);
    ensureShapeTextFrame(model, selectionAt(0, 0), { section: 0, block: 0 });
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

describe("withShapeModel/commitShapeModel (SPEC-051 AC-17, AC-18)", () => {
  it("edits a shape's text through the body-editing primitives — paragraphs, lists and tables all flow through unmodified", () => {
    const model = doc(paragraph("before"), shape({ text: { blocks: [paragraph("")], vAlign: "top" } }));
    const location = { section: 0, block: 1 };
    const synthetic = withShapeModel(model, location);
    expect(synthetic.sections).toHaveLength(1);
    const typed = insertText(synthetic, selectionAt(0, 0), "Shape text");
    const committed = commitShapeModel(model, location, typed.model);
    expect((shapeAt(committed, location)!.text!.blocks[0] as Paragraph).runs[0].text).toBe("Shape text");
    // The rest of the document (and the shape's own vAlign) is untouched.
    expect((committed.sections[0].blocks[0] as Paragraph).runs[0].text).toBe("before");
    expect(shapeAt(committed, location)!.text!.vAlign).toBe("top");
  });

  it("starts a shape with no text frame from one empty paragraph", () => {
    const model = doc(shape());
    const synthetic = withShapeModel(model, { section: 0, block: 0 });
    expect(synthetic.sections[0].blocks).toEqual([{ type: "paragraph", runs: [] }]);
  });

  it("never mutates the input model", () => {
    const model = doc(shape({ text: { blocks: [paragraph("x")], vAlign: "top" } }));
    const snapshot = structuredClone(model);
    const synthetic = withShapeModel(model, { section: 0, block: 0 });
    insertText(synthetic, selectionAt(0, 0), "y");
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

// --- hit testing / paint order (SPEC-051 AC-11) ----------------------------

describe("shapeHasVisibleText", () => {
  it("is false for a shape with no text frame", () => {
    expect(shapeHasVisibleText(shape())).toBe(false);
  });

  it("is false for a text frame holding only an empty paragraph", () => {
    expect(shapeHasVisibleText(shape({ text: { blocks: [paragraph("")], vAlign: "top" } }))).toBe(false);
  });

  it("is true once a paragraph carries a run with text", () => {
    expect(shapeHasVisibleText(shape({ text: { blocks: [paragraph("hi")], vAlign: "top" } }))).toBe(true);
  });

  it("is true for any non-paragraph block (e.g. a table) without inspecting it", () => {
    const withTable = shape({ text: { blocks: [grid([["a", "b"]])], vAlign: "top" } });
    expect(shapeHasVisibleText(withTable)).toBe(true);
  });
});

describe("shapesInPaintOrder", () => {
  it("orders by descending zOrder — the highest zOrder (nearest the reader) first", () => {
    const low = shape({ zOrder: 0 });
    const high = shape({ zOrder: 5 });
    const model = doc(low, high);
    const order = shapesInPaintOrder(model).map((d) => d.location.block);
    expect(order).toEqual([1, 0]);
  });

  it("ties on zOrder break to later document order first (paints on top)", () => {
    const a = shape();
    const b = shape();
    const model = doc(a, b);
    const order = shapesInPaintOrder(model).map((d) => d.location.block);
    expect(order).toEqual([1, 0]);
  });

  it("excludes images — Shape-only", () => {
    const model = doc(shape(), image());
    expect(shapesInPaintOrder(model)).toHaveLength(1);
  });
});

// --- anchor pairing (SPEC-078) ----------------------------------------------

describe("anchor pairing (SPEC-078)", () => {
  const shapeIn = (model: DocumentModel, block: number): Shape => {
    const found = shapeAt(model, { section: 0, block });
    if (!found) {
      throw new Error(`no shape at block ${block}`);
    }
    return found;
  };
  const paragraphIn = (model: DocumentModel, block: number): Paragraph => {
    const found = model.sections[0].blocks[block];
    if (found.type !== "paragraph") {
      throw new Error(`no paragraph at block ${block}`);
    }
    return found;
  };

  it("insertShape pairs the shape with the caret paragraph (end-of-paragraph insert)", () => {
    const model = doc(paragraph("Anchor me"));
    const { model: next } = insertShape(model, selectionAt(0, 9), { geometry: "rect", widthCm: 3, heightCm: 2 });
    const blocks = next.sections[0].blocks;
    expect(blocks.map((b) => b.type)).toEqual(["paragraph", "shape", "paragraph"]);
    const ordinal = paragraphIn(next, 0).anchorOrdinal;
    expect(ordinal).toBeDefined();
    expect(shapeIn(next, 1).anchorOrdinal).toBe(ordinal);
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
  });

  it("insertShape at offset 0 pairs with the paragraph it lands before", () => {
    const model = doc(paragraph("Anchor me"));
    const { model: next } = insertShape(model, selectionAt(0, 0), { geometry: "rect", widthCm: 3, heightCm: 2 });
    const blocks = next.sections[0].blocks;
    expect(blocks.map((b) => b.type)).toEqual(["shape", "paragraph"]);
    expect(shapeIn(next, 0).anchorOrdinal).toBe(paragraphIn(next, 1).anchorOrdinal);
    expect(resolveAnchorBlockIndex(blocks, 0)).toBe(1);
  });

  it("a mid-paragraph insert pairs with the FIRST half — the piece keeping the paragraph's top", () => {
    const model = doc(paragraph("Split here"));
    const { model: next } = insertShape(model, selectionAt(0, 5), { geometry: "rect", widthCm: 3, heightCm: 2 });
    const blocks = next.sections[0].blocks;
    expect(blocks.map((b) => b.type)).toEqual(["paragraph", "shape", "paragraph"]);
    expect(paragraphIn(next, 0).anchorOrdinal).toBeDefined();
    expect(paragraphIn(next, 2).anchorOrdinal).toBeUndefined();
    expect(shapeIn(next, 1).anchorOrdinal).toBe(paragraphIn(next, 0).anchorOrdinal);
  });

  it("two inserts onto one paragraph share a single ordinal", () => {
    const first = insertShape(doc(paragraph("Anchor")), selectionAt(0, 6), { geometry: "rect", widthCm: 3, heightCm: 2 });
    // Caret back into the anchor paragraph's end for the second insert.
    const second = insertShape(first.model, selectionAt(0, 6), { geometry: "ellipse", widthCm: 2, heightCm: 2 });
    const blocks = second.model.sections[0].blocks;
    const ordinal = paragraphIn(second.model, 0).anchorOrdinal;
    const shapes = blocks
      .map((block, index) => ({ block, index }))
      .filter(({ block }) => isShape(block));
    expect(shapes).toHaveLength(2);
    for (const { block, index } of shapes) {
      expect((block as Shape).anchorOrdinal).toBe(ordinal);
      expect(resolveAnchorBlockIndex(blocks, index)).toBe(0);
    }
  });

  it("duplicateShape's copy keeps the source's pairing — two drawings, one anchor", () => {
    const inserted = insertShape(doc(paragraph("Anchor")), selectionAt(0, 6), { geometry: "rect", widthCm: 3, heightCm: 2 });
    const { model: next } = duplicateShape(inserted.model, selectionAt(2, 0), { section: 0, block: 1 });
    const ordinal = paragraphIn(next, 0).anchorOrdinal;
    expect(shapeIn(next, 1).anchorOrdinal).toBe(ordinal);
    expect(shapeIn(next, 2).anchorOrdinal).toBe(ordinal);
    expect(resolveAnchorBlockIndex(next.sections[0].blocks, 2)).toBe(0);
  });

  it("re-anchoring one of two shared-anchor drawings leaves the other's pairing, offsets and resolution untouched", () => {
    // [P0(1), d1(->1), d2(->1), P1] — drag d2 down to P1.
    const model = doc(
      paragraph("P0"),
      shape({ offsetXCm: 1, offsetYCm: 2 }),
      shape({ offsetXCm: 4, offsetYCm: 5 }),
      paragraph("P1"),
    );
    paragraphIn(model, 0).anchorOrdinal = 1;
    shapeIn(model, 1).anchorOrdinal = 1;
    shapeIn(model, 2).anchorOrdinal = 1;
    const candidates: ParagraphSpan[] = [
      { section: 0, block: 0, topCm: 0, bottomCm: 1 },
      { section: 0, block: 3, topCm: 10, bottomCm: 11 },
    ];
    const { model: next } = reanchorShape(model, selectionAt(2, 0), { section: 0, block: 2 }, candidates, 10.5, 0);
    const blocks = next.sections[0].blocks;
    // d2 removed from index 2, lands after P1: [P0, d1, P1, d2].
    expect(blocks.map((b) => b.type)).toEqual(["paragraph", "shape", "paragraph", "shape"]);
    // d1 untouched: same pairing, same offsets, still resolving to P0.
    expect(shapeIn(next, 1).anchorOrdinal).toBe(1);
    expect(shapeIn(next, 1).offsetYCm).toBe(2);
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
    expect(paragraphIn(next, 0).anchorOrdinal).toBe(1);
    // d2 re-paired with P1 (fresh ordinal), offset re-expressed from P1's top.
    const newOrdinal = paragraphIn(next, 2).anchorOrdinal;
    expect(newOrdinal).toBeDefined();
    expect(newOrdinal).not.toBe(1);
    expect(shapeIn(next, 3).anchorOrdinal).toBe(newOrdinal);
    expect(shapeIn(next, 3).offsetYCm).toBe(5 - 10);
    expect(resolveAnchorBlockIndex(blocks, 3)).toBe(2);
  });

  it("re-anchoring the LAST referrer releases the old paragraph's ordinal", () => {
    const model = doc(paragraph("P0"), shape({ offsetYCm: 0 }), paragraph("P1"));
    paragraphIn(model, 0).anchorOrdinal = 1;
    shapeIn(model, 1).anchorOrdinal = 1;
    const candidates: ParagraphSpan[] = [
      { section: 0, block: 0, topCm: 0, bottomCm: 1 },
      { section: 0, block: 2, topCm: 10, bottomCm: 11 },
    ];
    const { model: next } = reanchorShape(model, selectionAt(1, 0), { section: 0, block: 1 }, candidates, 10.5, 0);
    expect(paragraphIn(next, 0).anchorOrdinal).toBeUndefined();
    expect(paragraphIn(next, 1).anchorOrdinal).toBeDefined();
  });

  it("a drawing spliced between a shared anchor and its other drawing does not disturb that drawing", () => {
    // [P0(1), d1(->1), P1(2), d2(->2)] — drag d2 up to P0; it lands right
    // after P0, BETWEEN P0 and d1 — the exact splice that used to flip d1's
    // inferred anchor.
    const model = doc(
      paragraph("P0"),
      shape({ offsetYCm: 2 }),
      paragraph("P1"),
      shape({ offsetYCm: 3 }),
    );
    paragraphIn(model, 0).anchorOrdinal = 1;
    shapeIn(model, 1).anchorOrdinal = 1;
    paragraphIn(model, 2).anchorOrdinal = 2;
    shapeIn(model, 3).anchorOrdinal = 2;
    const candidates: ParagraphSpan[] = [
      { section: 0, block: 0, topCm: 0, bottomCm: 1 },
      { section: 0, block: 2, topCm: 10, bottomCm: 11 },
    ];
    const { model: next } = reanchorShape(model, selectionAt(3, 0), { section: 0, block: 3 }, candidates, 0.5, 10);
    const blocks = next.sections[0].blocks;
    // d2 lands immediately after P0: [P0, d2, d1, P1].
    expect(blocks.map((b) => b.type)).toEqual(["paragraph", "shape", "shape", "paragraph"]);
    expect(shapeIn(next, 1).offsetYCm).toBe(3 + 10);
    expect(shapeIn(next, 1).anchorOrdinal).toBe(1);
    // d1 still resolves to P0 despite the newcomer between them.
    expect(shapeIn(next, 2).anchorOrdinal).toBe(1);
    expect(shapeIn(next, 2).offsetYCm).toBe(2);
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
    // P1's ordinal is gone — d2 was its last referrer.
    expect(paragraphIn(next, 3).anchorOrdinal).toBeUndefined();
  });

  it("deleteShape releases the pairing only when it was the last referrer", () => {
    const model = doc(paragraph("P0"), shape(), shape());
    paragraphIn(model, 0).anchorOrdinal = 1;
    shapeIn(model, 1).anchorOrdinal = 1;
    shapeIn(model, 2).anchorOrdinal = 1;
    const { model: afterFirst } = deleteShape(model, selectionAt(1, 0), { section: 0, block: 1 });
    expect(paragraphIn(afterFirst, 0).anchorOrdinal).toBe(1);
    const { model: afterBoth } = deleteShape(afterFirst, selectionAt(1, 0), { section: 0, block: 1 });
    expect(paragraphIn(afterBoth, 0).anchorOrdinal).toBeUndefined();
  });

  it("setShapeWrap to inline drops and releases; back to floating re-pairs by adjacency", () => {
    const model = doc(paragraph("P0"), shape());
    paragraphIn(model, 0).anchorOrdinal = 1;
    shapeIn(model, 1).anchorOrdinal = 1;
    const location: ShapeLocation = { section: 0, block: 1 };
    const { model: inline } = setShapeWrap(model, selectionAt(1, 0), location, "inline");
    expect(shapeIn(inline, 1).anchorOrdinal).toBeUndefined();
    expect(paragraphIn(inline, 0).anchorOrdinal).toBeUndefined();
    const { model: floating } = setShapeWrap(inline, selectionAt(1, 0), location, "square");
    const ordinal = paragraphIn(floating, 0).anchorOrdinal;
    expect(ordinal).toBeDefined();
    expect(shapeIn(floating, 1).anchorOrdinal).toBe(ordinal);
  });

  it("deleting the shape atom through deleteRange (the keyboard path) also releases the pairing", () => {
    const model = doc(paragraph("P0"), shape());
    paragraphIn(model, 0).anchorOrdinal = 1;
    shapeIn(model, 1).anchorOrdinal = 1;
    const { model: next } = deleteRange(model, {
      anchor: caret(1, 0),
      focus: caret(1, 1),
    });
    expect(next.sections[0].blocks.map((b) => b.type)).toEqual(["paragraph"]);
    expect(paragraphIn(next, 0).anchorOrdinal).toBeUndefined();
  });

  it("splitParagraphAt keeps the pairing on the first half only", () => {
    const source = paragraph("Split me");
    source.anchorOrdinal = 4;
    const [before, after] = splitParagraphAt(source, 5);
    expect(before.anchorOrdinal).toBe(4);
    expect(after.anchorOrdinal).toBeUndefined();
  });

  it("the Enter key's splitBlock keeps the pairing on the first half only", () => {
    const model = doc(paragraph("Split me"), shape());
    paragraphIn(model, 0).anchorOrdinal = 2;
    shapeIn(model, 1).anchorOrdinal = 2;
    const { model: next } = splitBlock(model, selectionAt(0, 5));
    expect(paragraphIn(next, 0).anchorOrdinal).toBe(2);
    expect(paragraphIn(next, 1).anchorOrdinal).toBeUndefined();
    // The shape still resolves to the top half.
    expect(resolveAnchorBlockIndex(next.sections[0].blocks, 2)).toBe(0);
  });
});

// --- arrangeDrawings (SPEC-072) ------------------------------------------------

describe("arrangeDrawings (SPEC-072)", () => {
  it("sets several drawings' offsets in one command", () => {
    const model = doc(
      shape({ offsetXCm: 1, offsetYCm: 2 }),
      shape({ offsetXCm: 5, offsetYCm: 6 }),
      paragraph("body"),
    );
    const { model: next } = arrangeDrawings(model, selectionAt(2, 0), [
      { location: { section: 0, block: 0 }, offsetXCm: 3, offsetYCm: 3 },
      { location: { section: 0, block: 1 }, offsetXCm: 3, offsetYCm: 9 },
    ]);
    expect(shapeAt(next, { section: 0, block: 0 })).toMatchObject({ offsetXCm: 3, offsetYCm: 3 });
    expect(shapeAt(next, { section: 0, block: 1 })).toMatchObject({ offsetXCm: 3, offsetYCm: 9 });
  });

  it("applies a single-axis reference change (single-drawing align)", () => {
    const model = doc(shape({ relativeToH: "column", offsetXCm: 4 }), paragraph("body"));
    const { model: next } = arrangeDrawings(model, selectionAt(1, 0), [
      { location: { section: 0, block: 0 }, relativeToH: "margin", offsetXCm: 0 },
    ]);
    const updated = shapeAt(next, { section: 0, block: 0 });
    expect(updated).toMatchObject({ relativeToH: "margin", offsetXCm: 0 });
    // The untouched vertical axis is preserved.
    expect(updated?.relativeToV).toBe("paragraph");
  });

  it("skips an inline shape, whose offset the wrap mode ignores", () => {
    const model = doc(shape({ wrap: "inline", offsetXCm: 0 }), paragraph("body"));
    const { model: next } = arrangeDrawings(model, selectionAt(1, 0), [
      { location: { section: 0, block: 0 }, offsetXCm: 7, offsetYCm: 7 },
    ]);
    expect(shapeAt(next, { section: 0, block: 0 })?.offsetXCm).toBe(0);
  });

  it("does not mutate the input model", () => {
    const model = doc(shape({ offsetXCm: 1 }), paragraph("body"));
    arrangeDrawings(model, selectionAt(1, 0), [
      { location: { section: 0, block: 0 }, offsetXCm: 9, offsetYCm: 9 },
    ]);
    expect(shapeAt(model, { section: 0, block: 0 })?.offsetXCm).toBe(1);
  });
});

// --- SPEC-073: formatDrawings fan-out (AC-26, AC-27) -------------------------

describe("formatDrawings — one-command styling fan-out (SPEC-073)", () => {
  const patch = {
    fill: { kind: "solid", color: "#dae3f3" } as const,
    outline: { color: "#4472c4", widthPt: 1, dash: "solid" } as const,
    shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 } as const,
  };

  it("applies fill/outline/shadow to every member of a multi-selection in one CommandResult", () => {
    const model = doc(shape(), shape({ offsetXCm: 5 }), paragraph(""));
    const { model: next } = formatDrawings(model, selectionAt(2, 0), [
      { section: 0, block: 0 },
      { section: 0, block: 1 },
    ], patch);
    for (const block of [0, 1]) {
      const s = shapeAt(next, { section: 0, block });
      expect(s?.fill).toEqual(patch.fill);
      expect(s?.outline).toEqual(patch.outline);
      expect(s?.shadow).toEqual(patch.shadow);
    }
  });

  it("a container location fans out to every group member (AC-26)", () => {
    const model = doc(shape({ groupOrdinal: 1 }), shape({ groupOrdinal: 1, offsetXCm: 5 }), paragraph(""));
    const { model: next } = formatDrawings(model, selectionAt(2, 0), [{ section: 0, block: 0 }], patch);
    expect(shapeAt(next, { section: 0, block: 0 })?.fill).toEqual(patch.fill);
    expect(shapeAt(next, { section: 0, block: 1 })?.fill).toEqual(patch.fill);
  });

  it("a connector member takes outline only (AC-26)", () => {
    const model = doc(shape({ groupOrdinal: 1 }), paragraph(""));
    const connector = {
      type: "connector",
      kind: "straight",
      start: { xCm: 0, yCm: 0 },
      end: { xCm: 3, yCm: 0 },
      widthCm: 3,
      heightCm: 0,
      wrap: "infront",
      groupOrdinal: 1,
    };
    model.sections[0].blocks.splice(1, 0, connector as never);
    const { model: next } = formatDrawings(model, selectionAt(2, 0), [{ section: 0, block: 0 }], patch);
    const member = next.sections[0].blocks[1] as { outline?: unknown; fill?: unknown; shadow?: unknown };
    expect(member.outline).toEqual(patch.outline);
    expect(member.fill).toBeUndefined();
    expect(member.shadow).toBeUndefined();
  });

  it("a line geometry member skips the fill (no area to fill)", () => {
    const model = doc(shape({ geometry: "line", heightCm: 0, groupOrdinal: 1 }), shape({ groupOrdinal: 1 }), paragraph(""));
    const { model: next } = formatDrawings(model, selectionAt(2, 0), [{ section: 0, block: 1 }], patch);
    expect(shapeAt(next, { section: 0, block: 0 })?.fill).toBeUndefined();
    expect(shapeAt(next, { section: 0, block: 0 })?.outline).toEqual(patch.outline);
  });

  it("an image member is skipped entirely", () => {
    const model = doc(shape({ groupOrdinal: 1 }), image({ groupOrdinal: 1 }), paragraph(""));
    const { model: next } = formatDrawings(model, selectionAt(2, 0), [{ section: 0, block: 0 }], patch);
    const member = next.sections[0].blocks[1] as Image & { fill?: unknown };
    expect(member.fill).toBeUndefined();
  });

  it("a sealed container is excluded — nothing in it is rewritten (AC-27)", () => {
    const preserved = shape({ groupOrdinal: 1, unsupportedGeometryXml: "<a:custGeom/>" });
    const model = doc(shape({ groupOrdinal: 1, fill: { kind: "solid", color: "#ff0000" } }), preserved, paragraph(""));
    const { model: next } = formatDrawings(model, selectionAt(2, 0), [{ section: 0, block: 0 }], patch);
    expect(shapeAt(next, { section: 0, block: 0 })?.fill).toEqual({ kind: "solid", color: "#ff0000" });
    expect(shapeAt(next, { section: 0, block: 1 })?.fill).toBeUndefined();
  });

  it("null clears a property; an absent key leaves it untouched", () => {
    const model = doc(
      shape({
        fill: { kind: "solid", color: "#ff0000" },
        outline: { color: "#000000", widthPt: 1, dash: "solid" },
        shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 },
      }),
      paragraph(""),
    );
    const { model: next } = formatDrawings(model, selectionAt(1, 0), [{ section: 0, block: 0 }], { shadow: null });
    const s = shapeAt(next, { section: 0, block: 0 });
    expect(s?.shadow).toBeUndefined();
    expect(s?.fill).toEqual({ kind: "solid", color: "#ff0000" });
    expect(s?.outline).toEqual({ color: "#000000", widthPt: 1, dash: "solid" });
  });
});
