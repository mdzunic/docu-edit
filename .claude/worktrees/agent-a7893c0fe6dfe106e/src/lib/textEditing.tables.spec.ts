// Cell-aware behaviour of the SPEC-010/011 editing primitives (SPEC-022):
// carets that descend into table cells, table atoms in ranges, and the
// boundary guarantees (no merge across a cell edge, no keystroke that
// swallows a table).
import { describe, expect, it } from "vitest";
import {
  blockLength,
  blockText,
  clampCaret,
  compareCaret,
  deleteBackward,
  deleteForward,
  deleteRange,
  insertText,
  normalizeRange,
  paragraphAt,
  paragraphSlicesInRange,
  patchFormatting,
  splitParagraph,
  type CaretPosition,
} from "./textEditing";
import {
  createDocument,
  isTable,
  type DocumentModel,
  type Paragraph,
  type Table,
} from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function grid(rows: string[][]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({
      cells: cells.map((text) => ({ blocks: [paragraph(text)] })),
    })),
    columnWidthsCm: rows[0].map(() => 5),
  };
}

function doc(...blocks: Array<Paragraph | Table>): DocumentModel {
  const model = createDocument("cells-fixture");
  model.sections = [{ blocks }];
  return model;
}

function cellCaret(
  block: number,
  row: number,
  col: number,
  cellBlock: number,
  offset: number,
): CaretPosition {
  return { section: 0, block, offset, cell: { row, col, block: cellBlock } };
}

function tableOf(model: DocumentModel, block: number): Table {
  const target = model.sections[0].blocks[block];
  if (!isTable(target)) {
    throw new Error("expected a table block");
  }
  return target;
}

function cellText(model: DocumentModel, block: number, row: number, col: number): string {
  return tableOf(model, block)
    .rows[row].cells[col].blocks.map(blockText)
    .join("|");
}

describe("blockText / blockLength over tables", () => {
  it("concatenates every cell paragraph's text, matching the DOM's textContent", () => {
    expect(blockText(grid([["a", "b"], ["c", "d"]]))).toBe("abcd");
  });

  it("gives a table the atom length 1", () => {
    expect(blockLength(grid([["a"]]))).toBe(1);
  });
});

describe("compareCaret with cell paths", () => {
  const before: CaretPosition = { section: 0, block: 1, offset: 0 };
  const after: CaretPosition = { section: 0, block: 1, offset: 1 };
  const first = cellCaret(1, 0, 0, 0, 0);
  const second = cellCaret(1, 0, 1, 0, 0);
  const below = cellCaret(1, 1, 0, 0, 0);

  it("orders cells by row, then column, then paragraph and offset", () => {
    expect(compareCaret(first, second)).toBeLessThan(0);
    expect(compareCaret(second, below)).toBeLessThan(0);
    expect(compareCaret(first, first)).toBe(0);
  });

  it("places the table atom's offset 0 before and offset 1 after every cell", () => {
    expect(compareCaret(before, first)).toBeLessThan(0);
    expect(compareCaret(after, below)).toBeGreaterThan(0);
  });
});

describe("clampCaret with tables", () => {
  const model = doc(paragraph("intro"), grid([["ab", "cd"]]));

  it("clamps a cell path into the table's grid and offset space", () => {
    expect(clampCaret(model, cellCaret(1, 5, 9, 4, 99))).toEqual(cellCaret(1, 0, 1, 0, 2));
  });

  it("clamps a cell-less table address to the 0..1 atom range", () => {
    expect(clampCaret(model, { section: 0, block: 1, offset: 7 })).toEqual({
      section: 0,
      block: 1,
      offset: 1,
    });
  });

  it("drops a stray cell path on a paragraph address", () => {
    expect(clampCaret(model, cellCaret(0, 0, 0, 0, 3))).toEqual({ section: 0, block: 0, offset: 3 });
  });
});

describe("insertText into a cell (AC-5)", () => {
  it("types into the addressed cell only, replacing the table block immutably", () => {
    const model = doc(paragraph("intro"), grid([["ab", "cd"]]), paragraph("outro"));
    const original = tableOf(model, 1);
    const caret = insertText(model, cellCaret(1, 0, 0, 0, 1), "X");
    expect(cellText(model, 1, 0, 0)).toBe("aXb");
    expect(cellText(model, 1, 0, 1)).toBe("cd");
    expect(caret).toEqual(cellCaret(1, 0, 0, 0, 2));
    // Replace-don't-mutate: a new table object, the untouched cell shared.
    const replaced = tableOf(model, 1);
    expect(replaced).not.toBe(original);
    expect(replaced.rows[0].cells[1]).toBe(original.rows[0].cells[1]);
  });

  it("is inert at a table atom position", () => {
    const model = doc(grid([["a"]]), paragraph(""));
    const caret = insertText(model, { section: 0, block: 0, offset: 0 }, "X");
    expect(cellText(model, 0, 0, 0)).toBe("a");
    expect(caret).toEqual({ section: 0, block: 0, offset: 0 });
  });
});

describe("splitParagraph inside a cell (Enter — multi-paragraph cells)", () => {
  it("grows the cell's own block list", () => {
    const model = doc(grid([["hello", "x"]]), paragraph(""));
    const caret = splitParagraph(model, cellCaret(0, 0, 0, 0, 2));
    expect(cellText(model, 0, 0, 0)).toBe("he|llo");
    expect(cellText(model, 0, 0, 1)).toBe("x");
    expect(caret).toEqual(cellCaret(0, 0, 0, 1, 0));
    // The section still holds two blocks — the split stayed inside the cell.
    expect(model.sections[0].blocks).toHaveLength(2);
  });
});

describe("deleteBackward / deleteForward at table boundaries (AC-14 guards)", () => {
  it("merges paragraphs within a cell", () => {
    const model = doc(grid([["ab"]]), paragraph(""));
    splitParagraph(model, cellCaret(0, 0, 0, 0, 1));
    expect(cellText(model, 0, 0, 0)).toBe("a|b");
    const caret = deleteBackward(model, cellCaret(0, 0, 0, 1, 0));
    expect(cellText(model, 0, 0, 0)).toBe("ab");
    expect(caret).toEqual(cellCaret(0, 0, 0, 0, 1));
  });

  it("is a no-op at the start of a cell's first paragraph", () => {
    const model = doc(paragraph("before"), grid([["ab"]]), paragraph(""));
    const at = cellCaret(1, 0, 0, 0, 0);
    expect(deleteBackward(model, at)).toEqual(at);
    expect(model.sections[0].blocks.map(blockText)).toEqual(["before", "ab", ""]);
  });

  it("never merges a paragraph after the table into it", () => {
    const model = doc(grid([["ab"]]), paragraph("after"));
    const at: CaretPosition = { section: 0, block: 1, offset: 0 };
    expect(deleteBackward(model, at)).toEqual(at);
    expect(blockText(model.sections[0].blocks[1])).toBe("after");
  });

  it("deleteForward stops at the end of a cell's last paragraph and before a table", () => {
    const model = doc(paragraph("before"), grid([["ab"]]), paragraph(""));
    const cellEnd = cellCaret(1, 0, 0, 0, 2);
    expect(deleteForward(model, cellEnd)).toEqual(cellEnd);
    const beforeTable: CaretPosition = { section: 0, block: 0, offset: 6 };
    expect(deleteForward(model, beforeTable)).toEqual(beforeTable);
    expect(model.sections[0].blocks.map(blockText)).toEqual(["before", "ab", ""]);
  });
});

describe("normalizeRange and deleteRange across tables", () => {
  it("keeps a same-cell range untouched", () => {
    const model = doc(grid([["hello"]]), paragraph(""));
    const { start, end } = normalizeRange(model, cellCaret(0, 0, 0, 0, 1), cellCaret(0, 0, 0, 0, 4));
    expect(start.cell).toBeDefined();
    expect(end.cell).toBeDefined();
    const caret = deleteRange(model, start, end);
    expect(cellText(model, 0, 0, 0)).toBe("ho");
    expect(caret).toEqual(cellCaret(0, 0, 0, 0, 1));
  });

  it("deletes across paragraphs within one cell", () => {
    const model = doc(grid([["ab"]]), paragraph(""));
    splitParagraph(model, cellCaret(0, 0, 0, 0, 1));
    deleteRange(model, cellCaret(0, 0, 0, 0, 0), cellCaret(0, 0, 0, 1, 1));
    expect(cellText(model, 0, 0, 0)).toBe("");
  });

  it("snaps a cross-cell range to the whole table and deletes it whole", () => {
    const model = doc(paragraph("before"), grid([["a", "b"], ["c", "d"]]), paragraph("after"));
    const caret = deleteRange(model, cellCaret(1, 0, 0, 0, 0), cellCaret(1, 1, 1, 0, 1));
    expect(model.sections[0].blocks.map(blockText)).toEqual(["before", "after"]);
    expect(caret).toEqual({ section: 0, block: 1, offset: 0 });
  });

  it("a range from outside into a table covers the table whole", () => {
    const model = doc(paragraph("before"), grid([["a"]]), paragraph("after"));
    deleteRange(model, { section: 0, block: 0, offset: 3 }, cellCaret(1, 0, 0, 0, 1));
    expect(model.sections[0].blocks.map(blockText)).toEqual(["bef", "after"]);
  });

  it("a range spanning past a table removes it and merges the boundary paragraphs", () => {
    const model = doc(paragraph("head"), grid([["x"]]), paragraph("tail"));
    deleteRange(model, { section: 0, block: 0, offset: 2 }, { section: 0, block: 2, offset: 2 });
    expect(model.sections[0].blocks.map(blockText)).toEqual(["heil"]);
  });
});

describe("paragraphSlicesInRange through tables", () => {
  it("covers every cell paragraph of a table inside a top-level range", () => {
    const model = doc(paragraph("head"), grid([["a", "b"]]), paragraph("tail"));
    const slices = paragraphSlicesInRange(
      model,
      { section: 0, block: 0, offset: 0 },
      { section: 0, block: 2, offset: 4 },
    );
    expect(slices.map((slice) => blockText(paragraphAt(model, slice.address)!))).toEqual([
      "head",
      "a",
      "b",
      "tail",
    ]);
  });

  it("lets select-all formatting reach cell text (AC-6 via ranges)", () => {
    const model = doc(paragraph("head"), grid([["a", "b"]]), paragraph("tail"));
    patchFormatting(
      model,
      { section: 0, block: 0, offset: 0 },
      { section: 0, block: 2, offset: 4 },
      { bold: true },
    );
    const table = tableOf(model, 1);
    expect(table.rows[0].cells[0].blocks[0].runs[0].formatting.bold).toBe(true);
    expect(table.rows[0].cells[1].blocks[0].runs[0].formatting.bold).toBe(true);
  });

  it("scopes a same-cell range to that cell's paragraphs", () => {
    const model = doc(grid([["ab", "cd"]]), paragraph(""));
    const slices = paragraphSlicesInRange(model, cellCaret(0, 0, 0, 0, 0), cellCaret(0, 0, 0, 0, 2));
    expect(slices).toHaveLength(1);
    expect(slices[0].address.cell).toEqual({ row: 0, col: 0, block: 0 });
  });
});
