// SPEC-023: cell merge/split, column/row resize, border presets, shading,
// padding, vertical alignment, table alignment/width. Framework-free
// commands over the SPEC-022 table model — see tables.spec.ts for the
// insert/delete/navigation coverage this file doesn't repeat.
import { describe, expect, it } from "vitest";
import {
  applyBorderPreset,
  fullTableRange,
  mergeCells,
  setCellPadding,
  setCellShading,
  setCellVerticalAlign,
  setColumnWidths,
  setRowIsHeader,
  setRowMinHeight,
  setTableAlignment,
  setTableWidthMode,
  splitCell,
  tableAt,
  type TableLocation,
} from "./tables";
import { MIN_COLUMN_WIDTH_CM, MIN_ROW_HEIGHT_CM, type CellRange } from "./core/tables";
import { caretSelection } from "./commands";
import { blockText } from "./textEditing";
import { createDocument, isTable, type DocumentModel, type Paragraph, type Table } from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

// These commands don't use the incoming selection for anything but passing
// it through unchanged (the caller restores the caret itself), so an
// arbitrary collapsed one stands in for it in every test below.
function someSelection() {
  return caretSelection({ section: 0, block: 0, offset: 0 });
}

function cell(text: string) {
  return { blocks: [paragraph(text)] };
}

function grid(rows: string[][], widths?: number[]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({ cells: cells.map(cell) })),
    columnWidthsCm: widths ?? rows[0].map(() => 5),
  };
}

function doc(...blocks: Array<Paragraph | Table>): DocumentModel {
  const model = createDocument("tables-formatting-fixture");
  model.sections = [{ blocks }];
  return model;
}

function loc(row: number, col: number): TableLocation {
  return { section: 0, block: 0, row, col };
}

function range(anchor: { row: number; col: number }, focus: { row: number; col: number }): CellRange {
  return { section: 0, block: 0, anchor, focus };
}

function table(model: DocumentModel): Table {
  const t = tableAt(model, 0, 0);
  if (!t) {
    throw new Error("expected a table");
  }
  return t;
}

describe("mergeCells (AC-6, AC-7)", () => {
  it("collapses a rectangular range into one cell with content joined as paragraphs", () => {
    const model = doc(grid([["A1", "B1"], ["A2", "B2"]]));
    const { model: next } = mergeCells(model, someSelection(), range({ row: 0, col: 0 }, { row: 1, col: 1 }));
    const t = table(next);
    expect(t.rows[0].cells[0].colSpan).toBe(2);
    expect(t.rows[0].cells[0].rowSpan).toBe(2);
    expect(t.rows[0].cells[0].blocks.map(blockText)).toEqual(["A1", "B1", "A2", "B2"]);
    expect(t.rows[0].cells[1].covered).toBe(true);
    expect(t.rows[1].cells[0].covered).toBe(true);
    expect(t.rows[1].cells[1].covered).toBe(true);
    // Every row still has one cell per grid column.
    expect(t.rows.every((row) => row.cells.length === 2)).toBe(true);
  });

  it("is a no-op for a range that resolves to a single cell", () => {
    const model = doc(grid([["A1", "B1"]]));
    const { model: next } = mergeCells(model, someSelection(), range({ row: 0, col: 0 }, { row: 0, col: 0 }));
    expect(table(next)).toEqual(table(model));
  });

  it("expands a range that only partly overlaps an already-merged cell", () => {
    const model = doc(grid([["A1", "B1", "C1"], ["A2", "B2", "C2"]]));
    // A vertical merge spanning column 0, rows 0..1.
    const merged = mergeCells(model, someSelection(), range({ row: 0, col: 0 }, { row: 1, col: 0 })).model;
    // Selecting row 0, columns 0..1 only touches the top half of the merged
    // cell; the merge's own rowSpan pulls row 1 into the final rectangle too.
    const { model: next } = mergeCells(merged, someSelection(), range({ row: 0, col: 0 }, { row: 0, col: 1 }));
    const t = table(next);
    expect(t.rows[0].cells[0].rowSpan).toBe(2);
    expect(t.rows[0].cells[0].colSpan).toBe(2);
  });

  it("moves the caret to the merged cell's top-left corner", () => {
    const model = doc(grid([["A1", "B1"], ["A2", "B2"]]));
    const { selection } = mergeCells(model, someSelection(), range({ row: 0, col: 0 }, { row: 1, col: 1 }));
    expect(selection.focus.cell).toEqual({ row: 0, col: 0, block: 0 });
  });
});

describe("splitCell (AC-8, AC-9)", () => {
  it("splits a plain cell into more columns, extending sibling rows' colSpan", () => {
    const model = doc(grid([["A1", "B1"], ["A2", "B2"]]));
    const { model: next } = splitCell(model, someSelection(), loc(0, 0), 1, 3);
    const t = table(next);
    expect(t.columnWidthsCm).toHaveLength(4);
    expect(t.rows[0].cells).toHaveLength(4);
    // 1 row × 3 columns is 3 independent cells, not a span: the original
    // content lands in the first, the rest start empty.
    expect(t.rows[0].cells.map((c) => c.blocks.map(blockText))).toEqual([["A1"], [""], [""], ["B1"]]);
    expect(t.rows[0].cells.every((c) => !c.covered)).toBe(true);
    // The other row's original cell now spans the 3 new grid columns so the
    // table stays rectangular.
    expect(t.rows[1].cells[0].colSpan).toBe(3);
    expect(t.rows[1].cells[1].covered).toBe(true);
    expect(t.rows[1].cells[2].covered).toBe(true);
    expect(t.rows[1].cells[3].blocks.map(blockText)).toEqual(["B2"]);
  });

  it("splits a plain cell into more rows, extending sibling columns' rowSpan", () => {
    const model = doc(grid([["A1", "B1"]]));
    const { model: next } = splitCell(model, someSelection(), loc(0, 0), 3, 1);
    const t = table(next);
    expect(t.rows).toHaveLength(3);
    expect(t.rows.map((row) => row.cells[0].blocks.map(blockText))).toEqual([["A1"], [""], [""]]);
    expect(t.rows.every((row) => !row.cells[0].covered)).toBe(true);
    expect(t.rows[0].cells[1].rowSpan).toBe(3);
    expect(t.rows[1].cells[1].covered).toBe(true);
    expect(t.rows[2].cells[1].covered).toBe(true);
  });

  it("splits a merged cell back down within its existing span", () => {
    const model = doc(grid([["A1", "B1", "C1"], ["A2", "B2", "C2"]]));
    const merged = mergeCells(model, someSelection(), range({ row: 0, col: 0 }, { row: 1, col: 1 })).model;
    const { model: next } = splitCell(merged, someSelection(), loc(0, 0), 2, 2);
    const t = table(next);
    expect(t.rows).toHaveLength(2);
    expect(t.columnWidthsCm).toHaveLength(3);
    expect(t.rows[0].cells[0].colSpan).toBeUndefined();
    expect(t.rows[0].cells[0].rowSpan).toBeUndefined();
    expect(t.rows[0].cells[1].colSpan).toBeUndefined();
    expect(t.rows[1].cells[0].covered).toBeUndefined();
    // Original merged content landed in the top-left sub-cell only.
    expect(t.rows[0].cells[0].blocks.map(blockText)).toEqual(["A1", "B1", "A2", "B2"]);
    expect(t.rows[0].cells[1].blocks.map(blockText)).toEqual([""]);
  });

  it("moves the caret to the split cell's origin", () => {
    const model = doc(grid([["A1", "B1"]]));
    const { selection } = splitCell(model, someSelection(), loc(0, 0), 1, 2);
    expect(selection.focus.cell).toEqual({ row: 0, col: 0, block: 0 });
  });
});

describe("setColumnWidths (AC-1..3)", () => {
  it("applies the given widths, clamped to the minimum", () => {
    const model = doc(grid([["A1", "B1"]], [5, 5]));
    const { model: next } = setColumnWidths(model, someSelection(), loc(0, 0), [8, 0.01]);
    const t = table(next);
    expect(t.columnWidthsCm[0]).toBeCloseTo(8, 2);
    expect(t.columnWidthsCm[1]).toBeCloseTo(MIN_COLUMN_WIDTH_CM, 2);
  });

  it("is a no-op when the widths array length does not match the column count", () => {
    const model = doc(grid([["A1", "B1"]], [5, 5]));
    const { model: next } = setColumnWidths(model, someSelection(), loc(0, 0), [8]);
    expect(table(next).columnWidthsCm).toEqual([5, 5]);
  });
});

describe("setRowMinHeight (AC-4, AC-5)", () => {
  it("sets the row's minimum height, clamped to the floor", () => {
    const model = doc(grid([["A1"], ["A2"]]));
    const { model: next } = setRowMinHeight(model, someSelection(), loc(0, 0), 0, 2);
    expect(table(next).rows[0].minHeightCm).toBeCloseTo(2, 2);
    const { model: tooSmall } = setRowMinHeight(model, someSelection(), loc(0, 0), 1, 0.01);
    expect(table(tooSmall).rows[1].minHeightCm).toBeCloseTo(MIN_ROW_HEIGHT_CM, 2);
  });
});

describe("setRowIsHeader (SPEC-043 mark-header-row fix)", () => {
  it("marks a row as header and clears its preserved source markup", () => {
    const model = doc(grid([["A1"], ["A2"]]));
    const withStale: DocumentModel = {
      ...model,
      sections: [
        {
          blocks: [
            {
              ...(table(model) as Table),
              rows: [
                { ...table(model).rows[0], docxTrPr: "<w:trPr/>", odtStyleName: "Stale" },
                table(model).rows[1],
              ],
            },
          ],
        },
      ],
    };
    const { model: next } = setRowIsHeader(withStale, someSelection(), loc(0, 0), 0, true);
    expect(table(next).rows[0].isHeader).toBe(true);
    expect(table(next).rows[0].docxTrPr).toBeUndefined();
    expect(table(next).rows[0].odtStyleName).toBeUndefined();
    expect(table(next).rows[1].isHeader).toBeUndefined();
  });

  it("unmarks a header row", () => {
    const model = doc(grid([["A1"], ["A2"]]));
    const { model: marked } = setRowIsHeader(model, someSelection(), loc(0, 0), 0, true);
    const { model: unmarked } = setRowIsHeader(marked, someSelection(), loc(0, 0), 0, false);
    expect(table(unmarked).rows[0].isHeader).toBeUndefined();
  });
});

describe("applyBorderPreset (AC-10..14, AC-18..21)", () => {
  const model = doc(grid([["A1", "B1"], ["A2", "B2"]]));
  const side = { style: "single" as const, widthPt: 1, color: "#000000" };

  it("'all' paints every edge of every cell in the range", () => {
    const { model: next } = applyBorderPreset(model, someSelection(), fullTableRange(model, loc(0, 0))!, "all", side);
    const t = table(next);
    for (const row of t.rows) {
      for (const c of row.cells) {
        expect(c.borders?.top?.style).toBe("single");
        expect(c.borders?.bottom?.style).toBe("single");
        expect(c.borders?.left?.style).toBe("single");
        expect(c.borders?.right?.style).toBe("single");
      }
    }
  });

  it("'outer' paints only the range's own boundary and clears the rest", () => {
    const { model: next } = applyBorderPreset(model, someSelection(), fullTableRange(model, loc(0, 0))!, "outer", side);
    const t = table(next);
    expect(t.rows[0].cells[0].borders?.top?.style).toBe("single");
    expect(t.rows[0].cells[0].borders?.left?.style).toBe("single");
    expect(t.rows[0].cells[0].borders?.right?.style).toBe("none");
    expect(t.rows[0].cells[0].borders?.bottom?.style).toBe("none");
    expect(t.rows[1].cells[1].borders?.bottom?.style).toBe("single");
    expect(t.rows[1].cells[1].borders?.right?.style).toBe("single");
  });

  it("'inner' paints only the shared edges between cells", () => {
    const { model: next } = applyBorderPreset(model, someSelection(), fullTableRange(model, loc(0, 0))!, "inner", side);
    const t = table(next);
    expect(t.rows[0].cells[0].borders?.top?.style).toBe("none");
    expect(t.rows[0].cells[0].borders?.left?.style).toBe("none");
    expect(t.rows[0].cells[0].borders?.right?.style).toBe("single");
    expect(t.rows[0].cells[0].borders?.bottom?.style).toBe("single");
  });

  it("'none' clears every edge", () => {
    const withBorders = applyBorderPreset(model, someSelection(), fullTableRange(model, loc(0, 0))!, "all", side).model;
    const { model: next } = applyBorderPreset(withBorders, someSelection(), fullTableRange(withBorders, loc(0, 0))!, "none", side);
    const t = table(next);
    expect(t.rows[0].cells[0].borders?.top?.style).toBe("none");
    expect(t.rows[1].cells[1].borders?.right?.style).toBe("none");
  });

  it("applies to just a selected range, leaving other cells untouched (AC-14)", () => {
    const { model: next } = applyBorderPreset(model, someSelection(), range({ row: 0, col: 0 }, { row: 0, col: 0 }), "all", side);
    const t = table(next);
    expect(t.rows[0].cells[0].borders?.top?.style).toBe("single");
    expect(t.rows[1].cells[1].borders).toBeUndefined();
  });
});

describe("setCellShading, setCellPadding, setCellVerticalAlign (AC-15..17)", () => {
  it("sets shading on the whole table via fullTableRange", () => {
    const model = doc(grid([["A1", "B1"]]));
    const { model: next } = setCellShading(model, someSelection(), fullTableRange(model, loc(0, 0))!, "#ffcc00");
    const t = table(next);
    expect(t.rows[0].cells[0].shadingFill).toBe("#ffcc00");
    expect(t.rows[0].cells[1].shadingFill).toBe("#ffcc00");
  });

  it("clears shading when given undefined", () => {
    const model = doc(grid([["A1"]]));
    const shaded = setCellShading(model, someSelection(), range({ row: 0, col: 0 }, { row: 0, col: 0 }), "#ffcc00").model;
    const { model: next } = setCellShading(shaded, someSelection(), range({ row: 0, col: 0 }, { row: 0, col: 0 }), undefined);
    expect(table(next).rows[0].cells[0].shadingFill).toBeUndefined();
  });

  it("sets padding on selected cells only", () => {
    const model = doc(grid([["A1", "B1"]]));
    const { model: next } = setCellPadding(model, someSelection(), range({ row: 0, col: 0 }, { row: 0, col: 0 }), 0.5);
    const t = table(next);
    expect(t.rows[0].cells[0].paddingCm).toBeCloseTo(0.5, 2);
    expect(t.rows[0].cells[1].paddingCm).toBeUndefined();
  });

  it("sets vertical alignment on selected cells", () => {
    const model = doc(grid([["A1", "B1"]]));
    const { model: next } = setCellVerticalAlign(model, someSelection(), range({ row: 0, col: 1 }, { row: 0, col: 1 }), "bottom");
    expect(table(next).rows[0].cells[1].verticalAlign).toBe("bottom");
  });
});

describe("setTableAlignment, setTableWidthMode (AC-22, AC-23)", () => {
  it("sets the table's page alignment", () => {
    const model = doc(grid([["A1"]]));
    const { model: next } = setTableAlignment(model, someSelection(), loc(0, 0), "right");
    expect(table(next).alignment).toBe("right");
  });

  it("sets a fixed/auto/percent width mode", () => {
    const model = doc(grid([["A1"]]));
    const percent = setTableWidthMode(model, someSelection(), loc(0, 0), "percent", 45).model;
    expect(table(percent).widthMode).toBe("percent");
    expect(table(percent).widthPercent).toBe(45);
    const auto = setTableWidthMode(percent, someSelection(), loc(0, 0), "auto").model;
    expect(table(auto).widthMode).toBe("auto");
    expect(table(auto).widthPercent).toBeUndefined();
  });
});

describe("mergeCells and splitCell are single edits (AC-34)", () => {
  it("mergeCells returns one CommandResult for the whole action", () => {
    const model = doc(grid([["A1", "B1"], ["A2", "B2"]]));
    const result = mergeCells(model, caretSelection({ section: 0, block: 0, offset: 0 }), range({ row: 0, col: 0 }, { row: 1, col: 1 }));
    expect(result.model).not.toBe(model);
    expect(isTable(table(result.model))).toBe(true);
  });
});
