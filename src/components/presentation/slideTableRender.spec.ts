import { describe, expect, it } from "vitest";
import type { TableFrame } from "../../lib/presentation/model";
import {
  tableColumnBoundaries,
  tableColumnWidths,
  tableRenderRows,
  tableRowBoundaries,
  tableRowHeights,
} from "./slideRender";

function table(): TableFrame {
  return {
    id: "t",
    kind: "table",
    xCm: 0,
    yCm: 0,
    widthCm: 8,
    heightCm: 2,
    columnWidthsCm: [4, 4],
    firstRow: true,
    rows: [
      { cells: [{ paragraphs: [{ runs: [{ text: "H1" }] }] }, { paragraphs: [{ runs: [{ text: "H2" }] }] }] },
      { cells: [{ paragraphs: [{ runs: [{ text: "a" }] }] }, { paragraphs: [{ runs: [{ text: "b" }] }] }] },
    ],
  };
}

describe("table render helpers (AC-4)", () => {
  it("emits one <col> per column width", () => {
    expect(tableColumnWidths(table())).toHaveLength(2);
  });

  it("drops covered slots and carries the owner's colspan/rowspan", () => {
    const t = table();
    // Merge row 0 into one 2-wide header cell.
    t.rows[0].cells[0].colSpan = 2;
    t.rows[0].cells[1] = { paragraphs: [], covered: true };
    const rows = tableRenderRows(t);
    expect(rows[0]).toHaveLength(1);
    expect(rows[0][0].colSpan).toBe(2);
    expect(rows[1]).toHaveLength(2);
  });

  it("resolves the header look from firstRow and banding from bandRow (AC-4b)", () => {
    const t = table();
    t.bandRow = true;
    t.rows.push({ cells: [{ paragraphs: [] }, { paragraphs: [] }] });
    const rows = tableRenderRows(t);
    // Row 0 is the header (firstRow) → header fill, not transparent.
    expect(rows[0][0].style.background).not.toBe("transparent");
    // Body row index 1 is the first body row (even) → no band; index 2 → banded.
    expect(rows[1][0].style.background).toBe("transparent");
    expect(rows[2][0].style.background).not.toBe("transparent");
  });

  it("explicit cell fill wins over the look", () => {
    const t = table();
    t.rows[1].cells[0].fill = { type: "solid", color: "#123456" };
    const rows = tableRenderRows(t);
    expect(rows[1][0].style.background).toBe("#123456");
  });

  it("falls back to a neutral grid border and honours explicit borders", () => {
    const t = table();
    t.rows[1].cells[1].borders = { top: { style: "none", widthPt: 0, color: "#000000" } };
    const rows = tableRenderRows(t);
    expect(rows[0][0].style.borderTop).toContain("solid"); // neutral fallback
    expect(rows[1][1].style.borderTop).toBe("none"); // explicit none wins
  });
});

describe("table resize boundaries (AC-8a/b)", () => {
  it("places one internal column boundary per gap at its cumulative width %", () => {
    const t = table(); // two 4cm columns → one boundary at 50%
    const cols = tableColumnBoundaries(t);
    expect(cols).toHaveLength(1);
    expect(cols[0]).toEqual({ index: 0, offsetPct: 50 });
  });

  it("places row boundaries from each row's height (default when unset)", () => {
    const t = table(); // two rows, no explicit heights → one boundary at 50%
    const rows = tableRowBoundaries(t);
    expect(rows).toHaveLength(1);
    expect(rows[0].index).toBe(0);
    expect(rows[0].offsetPct).toBeCloseTo(50, 5);
  });

  it("weights boundaries by uneven column widths", () => {
    const t = table();
    t.columnWidthsCm = [3, 1]; // boundary at 75%
    expect(tableColumnBoundaries(t)[0].offsetPct).toBeCloseTo(75, 5);
  });

  it("renders each row at its own heightCm so a resized row grows alone", () => {
    const t = table();
    t.rows[0].heightCm = 2;
    t.rows[1].heightCm = 0.5;
    const [top, bottom] = tableRowHeights(t);
    // Both are px strings and the taller row's px value is strictly larger.
    expect(parseFloat(top)).toBeGreaterThan(parseFloat(bottom));
  });
});
