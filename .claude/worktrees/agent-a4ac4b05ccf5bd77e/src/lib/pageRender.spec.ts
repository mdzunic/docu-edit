// SPEC-051 AC-17..21: a shape's text frame rendered as DocBlockEntries
// entries — DocShapeBlock.vue's only source of what a shape's text looks
// like, so it is worth testing independent of mounting a component.
import { describe, expect, it } from "vitest";
import { hoistFloatingDrawingEntries, shapeTextRenderEntries, type RenderEntry } from "./pageRender";
import { createDocument, type Block, type DocumentModel, type Paragraph, type Shape, type Table } from "./documentModel";
import type { ListDefinition } from "./core/lists";

function paragraph(text: string, overrides: Partial<Paragraph> = {}): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [], ...overrides };
}

function textShape(blocks: Array<Paragraph | Table>): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 3,
    heightCm: 2,
    wrap: "square",
    text: { blocks, vAlign: "top" },
  };
}

describe("shapeTextRenderEntries", () => {
  it("is empty for a shape with no text frame", () => {
    const shape: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square" };
    expect(shapeTextRenderEntries(shape)).toEqual([]);
  });

  it("renders each paragraph with its own run text and section/block addressing at section 0", () => {
    const shape = textShape([paragraph("First"), paragraph("Second")]);
    const entries = shapeTextRenderEntries(shape);
    expect(entries).toHaveLength(2);
    expect(entries[0].kind).toBe("paragraph");
    expect(entries[0].dataSection).toBe(0);
    expect(entries[0].dataBlock).toBe(0);
    expect(entries[0].runs?.map((r) => r.text)).toEqual(["First"]);
    expect(entries[1].dataBlock).toBe(1);
    expect(entries[1].runs?.map((r) => r.text)).toEqual(["Second"]);
  });

  it("carries character formatting through runStyle", () => {
    const shape = textShape([paragraph("bold", { runs: [{ text: "bold", formatting: { bold: true } }] })]);
    const entries = shapeTextRenderEntries(shape);
    expect(entries[0].runs?.[0].styleObj.fontWeight).toBe("bold");
  });

  it("computes list markers scoped to the shape's own text frame", () => {
    const item = paragraph("one", { list: { type: "bullet", level: 1 } });
    const shape = textShape([item]);
    const entries = shapeTextRenderEntries(shape);
    expect(entries[0].listMarker).toBeTruthy();
    expect(entries[0].listLevel).toBe(1);
  });

  it("resolves a defined list marker against the passed-in listDefinitions", () => {
    const defs: ListDefinition[] = [
      {
        defId: "D1",
        levels: Array.from({ length: 9 }, () => ({ style: "decimal" as const })),
      },
    ];
    const item = paragraph("one", { list: { type: "number", level: 1, defId: "D1" } });
    const shape = textShape([item]);
    const entries = shapeTextRenderEntries(shape, defs);
    expect(entries[0].listMarker).toBe("1");
  });

  it("renders a table block as a table entry addressed at section 0", () => {
    const table: Table = {
      type: "table",
      rows: [{ cells: [{ blocks: [paragraph("cell")] }] }],
      columnWidthsCm: [5],
    };
    const shape = textShape([table]);
    const entries = shapeTextRenderEntries(shape);
    expect(entries).toHaveLength(1);
    expect(entries[0].kind).toBe("table");
    expect(entries[0].sectionProp).toBe(0);
    expect(entries[0].blockIndexProp).toBe(0);
  });
});

// --- hoistFloatingDrawingEntries (SPEC-078) ---------------------------------

function model(...blocks: Block[]): DocumentModel {
  const base = createDocument("hoist-fixture");
  base.sections = [{ blocks }];
  return base;
}

function paragraphEntry(block: number, column?: number): RenderEntry {
  return { key: `p:${block}`, memo: [block], kind: "paragraph", dataSection: 0, dataBlock: block, column };
}

function shapeEntry(block: number, blockProp: Shape, column?: number): RenderEntry {
  return {
    key: `s:${block}`,
    memo: [block, column],
    kind: "shape",
    sectionProp: 0,
    blockIndexProp: block,
    blockProp,
    column,
  };
}

const floating = (anchorOrdinal?: number): Shape => {
  const block: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square" };
  if (anchorOrdinal !== undefined) {
    block.anchorOrdinal = anchorOrdinal;
  }
  return block;
};

describe("hoistFloatingDrawingEntries", () => {

  it("moves a floating shape entry to sit immediately before its anchor paragraph's entry", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const drawing = floating(1);
    const m = model(anchor, drawing);
    const entries = [paragraphEntry(0), shapeEntry(1, drawing)];
    const hoisted = hoistFloatingDrawingEntries(entries, m);
    expect(hoisted.map((entry) => entry.key)).toEqual(["s:1", "p:0"]);
  });

  it("keeps the relative order of drawings sharing an anchor", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const d1 = floating(1);
    const d2 = floating(1);
    const m = model(anchor, d1, d2);
    const entries = [paragraphEntry(0), shapeEntry(1, d1), shapeEntry(2, d2)];
    const hoisted = hoistFloatingDrawingEntries(entries, m);
    expect(hoisted.map((entry) => entry.key)).toEqual(["s:1", "s:2", "p:0"]);
  });

  it("hoists by the adjacency fallback too — legacy documents with no ordinals", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [] };
    const drawing = floating();
    const m = model(anchor, drawing);
    const entries = [paragraphEntry(0), shapeEntry(1, drawing)];
    expect(hoistFloatingDrawingEntries(entries, m).map((entry) => entry.key)).toEqual(["s:1", "p:0"]);
  });

  it("adopts the anchor's column and extends the memo so v-memo sees the change", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const drawing = floating(1);
    const m = model(anchor, drawing);
    const entries = [paragraphEntry(0, 0), shapeEntry(1, drawing, 1)];
    const hoisted = hoistFloatingDrawingEntries(entries, m);
    expect(hoisted[0].column).toBe(0);
    expect(hoisted[0].memo[hoisted[0].memo.length - 1]).toBe(0);
  });

  it("targets the anchor paragraph's FIRST entry when it is split into slices", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const drawing = floating(1);
    const m = model(anchor, drawing);
    const secondSlice: RenderEntry = { ...paragraphEntry(0), key: "p:0:slice2" };
    const entries = [paragraphEntry(0), secondSlice, shapeEntry(1, drawing)];
    expect(hoistFloatingDrawingEntries(entries, m).map((entry) => entry.key)).toEqual([
      "s:1",
      "p:0",
      "p:0:slice2",
    ]);
  });

  it("leaves a drawing alone when its anchor has no entry on the page, and inline shapes always", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const drawing = floating(1);
    const inline: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "inline" };
    const m = model(anchor, drawing, inline);
    const offPage = [paragraphEntry(5), shapeEntry(1, drawing)];
    expect(hoistFloatingDrawingEntries(offPage, m)).toBe(offPage);
    const inlineEntries = [paragraphEntry(0), shapeEntry(2, inline)];
    expect(hoistFloatingDrawingEntries(inlineEntries, m)).toBe(inlineEntries);
  });

  it("returns the input array untouched when everything already sits in place", () => {
    const anchor: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const drawing = floating(1);
    const m = model(anchor, drawing);
    const entries = [shapeEntry(1, drawing), paragraphEntry(0)];
    expect(hoistFloatingDrawingEntries(entries, m)).toBe(entries);
  });
});

// Main's original hoist suite (ed41260) ported to the SPEC-078 resolver
// signature. Two of its pinned outcomes changed with the explicit pairing —
// each notes the old expectation it supersedes.
describe("hoistFloatingDrawingEntries — ported adjacency-era cases", () => {
  it("moves a floating shape (and a chart) before its anchor paragraph", () => {
    const p0: Paragraph = { type: "paragraph", runs: [] };
    const d1 = floating();
    const p2: Paragraph = { type: "paragraph", runs: [] };
    const chart = { type: "chart", chartType: "columnClustered", categories: [], series: [], options: {}, widthCm: 12, heightCm: 7.5, wrap: "square" } as unknown as Block;
    const m = model(p0, d1, p2, chart);
    const chartEntry: RenderEntry = { key: "c:3", memo: [], kind: "chart", sectionProp: 0, blockIndexProp: 3, blockProp: chart };
    const out = hoistFloatingDrawingEntries([paragraphEntry(0), shapeEntry(1, d1), paragraphEntry(2), chartEntry], m);
    expect(out.map((e) => e.key)).toEqual(["s:1", "p:0", "c:3", "p:2"]);
  });

  it("a stacked follower with NO explicit pairing hoists too — the fallback walks over its sibling (SPEC-078; the adjacency-only hoist stranded it after the paragraph)", () => {
    const p0: Paragraph = { type: "paragraph", runs: [] };
    const d1 = floating();
    const d2 = floating();
    const m = model(p0, d1, d2);
    const out = hoistFloatingDrawingEntries([paragraphEntry(0), shapeEntry(1, d1), shapeEntry(2, d2)], m);
    expect(out.map((e) => e.key)).toEqual(["s:1", "s:2", "p:0"]);
  });

  it("leaves a lands-before drawing (paired with the FOLLOWING paragraph) at its anchor's top", () => {
    const p0: Paragraph = { type: "paragraph", runs: [] };
    const p1: Paragraph = { type: "paragraph", runs: [] };
    const d2 = floating(1);
    const p3: Paragraph = { type: "paragraph", runs: [], anchorOrdinal: 1 };
    const m = model(p0, p1, d2, p3);
    const entries = [paragraphEntry(0), shapeEntry(2, d2), paragraphEntry(3)];
    expect(hoistFloatingDrawingEntries(entries, m)).toBe(entries);
  });

  it("leaves a drawing whose anchor paragraph's entry is missing from the page (a page-boundary split)", () => {
    const p0: Paragraph = { type: "paragraph", runs: [] };
    const d1 = floating();
    const p2: Paragraph = { type: "paragraph", runs: [] };
    const m = model(p0, d1, p2);
    const entries = [shapeEntry(1, d1), paragraphEntry(2)];
    expect(hoistFloatingDrawingEntries(entries, m)).toBe(entries);
  });

  it("a drawing separated from its resolved anchor by off-page paragraphs stays put", () => {
    const blocks: Block[] = [
      { type: "paragraph", runs: [] },
      { type: "paragraph", runs: [] },
      { type: "paragraph", runs: [] },
      { type: "paragraph", runs: [] },
      { type: "paragraph", runs: [] },
      floating(),
    ];
    const m = model(...blocks);
    // Fallback resolves block 5's anchor to block 4, which has no entry
    // here — only block 0 does.
    const entries = [paragraphEntry(0), shapeEntry(5, blocks[5] as Shape)];
    expect(hoistFloatingDrawingEntries(entries, m)).toBe(entries);
  });

  it("hoists across a column boundary by adopting the anchor's column (SPEC-078; the adjacency-only hoist refused cross-column moves)", () => {
    const p0: Paragraph = { type: "paragraph", runs: [] };
    const d1 = floating();
    const m = model(p0, d1);
    const out = hoistFloatingDrawingEntries([paragraphEntry(0, 0), shapeEntry(1, d1, 1)], m);
    expect(out.map((e) => e.key)).toEqual(["s:1", "p:0"]);
    expect(out[0].column).toBe(0);
  });
});
