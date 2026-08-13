// SPEC-051 AC-17..21: a shape's text frame rendered as DocBlockEntries
// entries — DocShapeBlock.vue's only source of what a shape's text looks
// like, so it is worth testing independent of mounting a component.
import { describe, expect, it } from "vitest";
import { shapeTextRenderEntries } from "./pageRender";
import type { ListDefinition, Paragraph, Shape, Table } from "./documentModel";

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
