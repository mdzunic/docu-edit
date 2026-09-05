// SPEC-078: the explicit-anchor resolver and its bookkeeping — the single
// derivation every consumer (anchor marker, re-anchor math, renderer
// hoisting, both writers) shares.
import { describe, expect, it } from "vitest";
import {
  anchorOrdinalReferenced,
  anchoredParagraph,
  isFloatingDrawing,
  mintAnchorOrdinal,
  ordinalForParagraph,
  paragraphIndexForOrdinal,
  planAnchoredDrawings,
  releaseAnchorOrdinal,
  resolveAnchorBlockIndex,
} from "./drawingAnchors";
import { cloneContainers, createDocument, type Block, type DocumentModel, type Image, type Paragraph, type Shape } from "./documentModel";

function paragraph(text: string, anchorOrdinal?: number): Paragraph {
  const block: Paragraph = { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
  if (anchorOrdinal !== undefined) {
    block.anchorOrdinal = anchorOrdinal;
  }
  return block;
}

function shape(overrides: Partial<Shape> = {}): Shape {
  return { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square", ...overrides };
}

function image(overrides: Partial<Image> = {}): Image {
  return {
    type: "image",
    mediaId: "m1",
    naturalWidthPx: 10,
    naturalHeightPx: 10,
    widthCm: 2,
    heightCm: 2,
    wrap: "square",
    ...overrides,
  };
}

function doc(...blocks: Block[]): DocumentModel {
  const model = createDocument("anchors-fixture");
  model.sections = [{ blocks }];
  return model;
}

describe("isFloatingDrawing", () => {
  it("is true for any non-inline drawing and false for inline ones and paragraphs", () => {
    expect(isFloatingDrawing(shape())).toBe(true);
    expect(isFloatingDrawing(shape({ wrap: "inline" }))).toBe(false);
    // Image joined this list when it gained real placement: its offsets
    // measure from the anchor paragraph's own origin, which only holds if it
    // pairs and hoists the same way every other floating drawing does.
    expect(isFloatingDrawing(image())).toBe(true);
    expect(isFloatingDrawing(image({ wrap: "inline" }))).toBe(false);
    expect(isFloatingDrawing(paragraph("p"))).toBe(false);
    expect(isFloatingDrawing(undefined)).toBe(false);
  });
});

describe("resolveAnchorBlockIndex — explicit ordinal", () => {
  it("resolves to the paragraph sharing the drawing's ordinal, wherever it sits", () => {
    const blocks: Block[] = [paragraph("a", 7), paragraph("b"), shape({ anchorOrdinal: 7 })];
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
  });

  it("two drawings anchored to one paragraph both resolve to it — the SPEC-078 headline case", () => {
    const blocks: Block[] = [paragraph("anchor", 3), shape({ anchorOrdinal: 3 }), shape({ anchorOrdinal: 3 }), paragraph("after")];
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
  });

  it("a drawing spliced between another drawing and their shared anchor changes neither resolution", () => {
    const blocks: Block[] = [paragraph("anchor", 3), shape({ anchorOrdinal: 3 }), shape({ anchorOrdinal: 3 })];
    // Simulates reanchorShape landing a third drawing right after the anchor.
    blocks.splice(1, 0, shape({ anchorOrdinal: 3 }));
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
    expect(resolveAnchorBlockIndex(blocks, 3)).toBe(0);
  });

  it("picks the FIRST paragraph carrying the ordinal (a split's top half)", () => {
    const blocks: Block[] = [paragraph("top", 5), paragraph("bottom", 5), shape({ anchorOrdinal: 5 })];
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
  });

  it("a dangling ordinal falls back to adjacency", () => {
    const blocks: Block[] = [paragraph("a"), shape({ anchorOrdinal: 99 })];
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
  });
});

describe("resolveAnchorBlockIndex — adjacency fallback", () => {
  it("walks backward over consecutive drawings to the nearest paragraph", () => {
    const blocks: Block[] = [paragraph("a"), shape(), shape(), shape()];
    expect(resolveAnchorBlockIndex(blocks, 3)).toBe(0);
  });

  it("walks over an image too — any drawing block, not only shapes", () => {
    const blocks: Block[] = [paragraph("a"), image(), shape()];
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
  });

  it("walks forward when nothing paragraph-shaped precedes", () => {
    const blocks: Block[] = [shape(), shape(), paragraph("after")];
    expect(resolveAnchorBlockIndex(blocks, 0)).toBe(2);
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(2);
  });

  it("stops at a non-drawing block in both directions", () => {
    const table = { type: "table", rows: [], columnWidthsCm: [] } as unknown as Block;
    expect(resolveAnchorBlockIndex([paragraph("a"), table, shape()], 2)).toBe(null);
    expect(resolveAnchorBlockIndex([shape(), table, paragraph("b")], 0)).toBe(null);
  });

  it("is null for a non-drawing index", () => {
    expect(resolveAnchorBlockIndex([paragraph("a")], 0)).toBe(null);
  });
});

describe("paragraphIndexForOrdinal / mintAnchorOrdinal / anchorOrdinalReferenced", () => {
  it("finds the first paragraph carrying the ordinal, or null", () => {
    const blocks: Block[] = [paragraph("a"), paragraph("b", 2)];
    expect(paragraphIndexForOrdinal(blocks, 2)).toBe(1);
    expect(paragraphIndexForOrdinal(blocks, 9)).toBe(null);
  });

  it("mints past every paragraph and drawing ordinal — dangling references included", () => {
    expect(mintAnchorOrdinal(doc(paragraph("a", 2), shape({ anchorOrdinal: 6 })))).toBe(7);
    expect(mintAnchorOrdinal(doc(paragraph("a")))).toBe(1);
  });

  it("reports whether any drawing besides the excluded one references an ordinal", () => {
    const model = doc(paragraph("a", 2), shape({ anchorOrdinal: 2 }), shape({ anchorOrdinal: 2 }));
    expect(anchorOrdinalReferenced(model, 2)).toBe(true);
    expect(anchorOrdinalReferenced(model, 2, { section: 0, block: 1 })).toBe(true);
    const single = doc(paragraph("a", 2), shape({ anchorOrdinal: 2 }));
    expect(anchorOrdinalReferenced(single, 2, { section: 0, block: 1 })).toBe(false);
  });

  it("ordinalForParagraph reuses the paragraph's own ordinal and mints otherwise", () => {
    const model = doc(paragraph("a", 4), paragraph("b"), shape({ anchorOrdinal: 4 }));
    expect(ordinalForParagraph(model, model.sections[0].blocks[0] as Paragraph)).toBe(4);
    expect(ordinalForParagraph(model, model.sections[0].blocks[1] as Paragraph)).toBe(5);
  });
});

describe("releaseAnchorOrdinal", () => {
  it("strips an unreferenced ordinal from its paragraph without mutating the shared block", () => {
    const source = doc(paragraph("a", 2));
    const model = cloneContainers(source);
    releaseAnchorOrdinal(model, 2);
    expect((model.sections[0].blocks[0] as Paragraph).anchorOrdinal).toBeUndefined();
    expect((source.sections[0].blocks[0] as Paragraph).anchorOrdinal).toBe(2);
  });

  it("keeps an ordinal another drawing still references", () => {
    const model = cloneContainers(doc(paragraph("a", 2), shape({ anchorOrdinal: 2 })));
    releaseAnchorOrdinal(model, 2);
    expect((model.sections[0].blocks[0] as Paragraph).anchorOrdinal).toBe(2);
  });

  it("anchoredParagraph returns a carrying copy, leaving the original alone", () => {
    const original = paragraph("a");
    const paired = anchoredParagraph(original, 9);
    expect(paired.anchorOrdinal).toBe(9);
    expect(original.anchorOrdinal).toBeUndefined();
  });
});

describe("planAnchoredDrawings", () => {
  it("absorbs only explicitly paired drawings that resolve, split by side", () => {
    const blocks: Block[] = [
      shape({ anchorOrdinal: 1 }), // leading (precedes its anchor)
      paragraph("anchor", 1),
      shape({ anchorOrdinal: 1 }), // trailing
      shape(), // no ordinal — legacy, stays at its slot
      shape({ anchorOrdinal: 99 }), // dangling — stays at its slot
    ];
    const plan = planAnchoredDrawings(blocks);
    expect([...plan.absorbed].toSorted()).toEqual([0, 2]);
    expect(plan.byParagraph.get(1)).toEqual({ leading: [0], trailing: [2] });
  });

  it("an inline drawing never absorbs even with a stray ordinal", () => {
    const blocks: Block[] = [paragraph("a", 1), shape({ wrap: "inline", anchorOrdinal: 1 })];
    expect(planAnchoredDrawings(blocks).absorbed.size).toBe(0);
  });
});
