// SPEC-076: the pure multi-key table sort. Nothing here mounts a component
// (AC-42, AC-78).
import { describe, expect, it } from "vitest";
import {
  parseSortDate,
  parseSortNumber,
  sortRangeHasMergedRows,
  sortRangeRowCount,
  sortRangeStart,
  sortTableRows,
  type SortKey,
} from "./tableSort";
import { equalStructural } from "./equality";
import type { Table, TableCell } from "./documentModel";

function cell(text: string): TableCell {
  return { blocks: [{ type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] }] };
}

function grid(rows: string[][]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({ cells: cells.map(cell) })),
    columnWidthsCm: rows[0].map(() => 3),
  };
}

function column(table: Table, columnIndex: number): string[] {
  return table.rows.map((row) => row.cells[columnIndex].blocks[0].runs.map((run) => run.text).join(""));
}

function key(columnIndex: number, direction: "asc" | "desc" = "asc", type: SortKey["type"] = "text"): SortKey {
  return { columnIndex, direction, type };
}

describe("sortTableRows — the sort range (AC-41)", () => {
  it("skips the leading isHeader run when hasHeader is on", () => {
    const table = grid([["Name"], ["Charlie"], ["alpha"], ["Bravo"]]);
    table.rows[0].isHeader = true;
    const sorted = sortTableRows(table, [key(0)], true);
    expect(column(sorted, 0)).toEqual(["Name", "alpha", "Bravo", "Charlie"]);
  });

  it("sorts every row when hasHeader is off", () => {
    const table = grid([["Name"], ["Charlie"], ["alpha"]]);
    table.rows[0].isHeader = true;
    const sorted = sortTableRows(table, [key(0)], false);
    expect(column(sorted, 0)).toEqual(["alpha", "Charlie", "Name"]);
  });

  it("skips a multi-row header run", () => {
    const table = grid([["H1"], ["H2"], ["b"], ["a"]]);
    table.rows[0].isHeader = true;
    table.rows[1].isHeader = true;
    expect(sortRangeStart(table, true)).toBe(2);
    expect(sortRangeRowCount(table, true)).toBe(2);
    expect(column(sortTableRows(table, [key(0)], true), 0)).toEqual(["H1", "H2", "a", "b"]);
  });
});

describe("sortTableRows — comparison (AC-43..AC-46)", () => {
  it("a text key compares with the document language's collation", () => {
    const table = grid([["éclair"], ["zebra"], ["Apple"]]);
    expect(column(sortTableRows(table, [key(0)], false, "fr-FR"), 0)).toEqual(["Apple", "éclair", "zebra"]);
  });

  it("a number key parses leniently and sorts numerically, not lexically", () => {
    const table = grid([["1,234.50"], ["9"], ["(42)"], ["100"]]);
    expect(column(sortTableRows(table, [key(0, "asc", "number")], false), 0)).toEqual([
      "(42)",
      "9",
      "100",
      "1,234.50",
    ]);
  });

  it("a cell that does not parse as a number sorts before every cell that does", () => {
    const table = grid([["7"], ["n/a"], ["2"]]);
    expect(column(sortTableRows(table, [key(0, "asc", "number")], false), 0)).toEqual(["n/a", "2", "7"]);
  });

  it("a date key parses ISO 8601 and the language's common formats", () => {
    const table = grid([["2026-03-01"], ["01/02/2026"], ["12 January 2026"]]);
    expect(column(sortTableRows(table, [key(0, "asc", "date")], false, "en-GB"), 0)).toEqual([
      "12 January 2026",
      "01/02/2026",
      "2026-03-01",
    ]);
  });

  it("a cell that does not parse as a date sorts before every cell that does", () => {
    const table = grid([["2026-01-02"], ["soon"], ["2025-01-02"]]);
    expect(column(sortTableRows(table, [key(0, "asc", "date")], false), 0)).toEqual([
      "soon",
      "2025-01-02",
      "2026-01-02",
    ]);
  });

  it("desc reverses only its own key", () => {
    const table = grid([
      ["b", "1"],
      ["a", "2"],
      ["b", "2"],
      ["a", "1"],
    ]);
    const sorted = sortTableRows(table, [key(0, "desc"), key(1, "asc", "number")], false);
    expect(sorted.rows.map((row) => row.cells.map((c) => c.blocks[0].runs[0].text).join(""))).toEqual([
      "b1",
      "b2",
      "a1",
      "a2",
    ]);
  });
});

describe("sortTableRows — stability (AC-47)", () => {
  it("ties fall through to the next key, then to the original row order", () => {
    const table = grid([
      ["a", "z", "1"],
      ["a", "y", "2"],
      ["a", "y", "3"],
    ]);
    const sorted = sortTableRows(table, [key(0), key(1)], false);
    expect(column(sorted, 2)).toEqual(["2", "3", "1"]);
  });

  it("applying the same keys twice produces an equalStructural table", () => {
    const table = grid([["c", "2"], ["a", "1"], ["b", "2"], ["a", "3"]]);
    const keys = [key(1, "asc", "number"), key(0)];
    const once = sortTableRows(table, keys, false);
    const twice = sortTableRows(once, keys, false);
    expect(equalStructural(once, twice)).toBe(true);
  });
});

describe("sortTableRows — rows move whole (AC-48)", () => {
  it("carries cells, spans, shading, minHeightCm, preserved markup and tracked changes with the row", () => {
    const table = grid([["b", "x"], ["a", "y"]]);
    table.rows[0].minHeightCm = 1.5;
    table.rows[0].docxTrPr = "<w:trPr/>";
    table.rows[0].odtStyleName = "ro1";
    table.rows[0].changes = { insert: { changeId: "c1", author: "A", date: "2026-01-01T00:00:00Z" } };
    table.rows[0].cells[0].shadingFill = "#ff0000";
    table.rows[0].cells[0].colSpan = 1;
    table.rows[0].cells[1].borders = { top: { style: "single", widthPt: 1, color: "#000000" } };

    const sorted = sortTableRows(table, [key(0)], false);
    expect(column(sorted, 0)).toEqual(["a", "b"]);
    const moved = sorted.rows[1];
    expect(moved.minHeightCm).toBe(1.5);
    expect(moved.docxTrPr).toBe("<w:trPr/>");
    expect(moved.odtStyleName).toBe("ro1");
    expect(moved.changes?.insert?.author).toBe("A");
    expect(moved.cells[0].shadingFill).toBe("#ff0000");
    expect(moved.cells[1].borders?.top?.color).toBe("#000000");
  });

  it("never mutates the input table", () => {
    const table = grid([["b"], ["a"]]);
    const before = JSON.stringify(table);
    sortTableRows(table, [key(0)], false);
    expect(JSON.stringify(table)).toBe(before);
  });
});

describe("sort refusals (AC-49, AC-50)", () => {
  it("a rowSpan > 1 inside the range refuses the sort", () => {
    const table = grid([["h"], ["a"], ["b"], ["c"]]);
    table.rows[0].isHeader = true;
    table.rows[1].cells[0].rowSpan = 2;
    table.rows[2].cells[0] = { blocks: [], covered: true };
    expect(sortRangeHasMergedRows(table, true)).toBe(true);
  });

  it("a covered cell whose owner spans rows from above the range refuses the sort", () => {
    const table = grid([["h"], ["a"], ["b"]]);
    table.rows[0].isHeader = true;
    table.rows[0].cells[0].rowSpan = 2;
    table.rows[1].cells[0] = { blocks: [], covered: true };
    expect(sortRangeHasMergedRows(table, true)).toBe(true);
  });

  it("a colSpan-only merge does not refuse the sort", () => {
    const table = grid([
      ["a", "x"],
      ["b", "y"],
    ]);
    table.rows[0].cells[0].colSpan = 2;
    table.rows[0].cells[1] = { blocks: [], covered: true };
    expect(sortRangeHasMergedRows(table, false)).toBe(false);
  });

  it("a range of fewer than two rows has nothing to sort", () => {
    const table = grid([["h"], ["only"]]);
    table.rows[0].isHeader = true;
    expect(sortRangeRowCount(table, true)).toBe(1);
    expect(equalStructural(sortTableRows(table, [key(0)], true), table)).toBe(true);
  });
});

describe("lenient parsing", () => {
  it("parses numbers with either separator convention, signs and currency", () => {
    expect(parseSortNumber("1,234.5")).toBe(1234.5);
    expect(parseSortNumber("1.234,5")).toBe(1234.5);
    expect(parseSortNumber("1.234")).toBe(1234);
    expect(parseSortNumber("1,5")).toBe(1.5);
    expect(parseSortNumber(" £ 42 ")).toBe(42);
    expect(parseSortNumber("-3.5")).toBe(-3.5);
    expect(parseSortNumber("(7)")).toBe(-7);
    expect(parseSortNumber("12%")).toBe(12);
    expect(parseSortNumber("")).toBeNull();
    expect(parseSortNumber("abc")).toBeNull();
    expect(parseSortNumber("1-2-3")).toBeNull();
  });

  it("reads a numeric date day-first or month-first by the document language", () => {
    expect(parseSortDate("03/04/2026", "en-GB")).toBe(Date.UTC(2026, 3, 3));
    expect(parseSortDate("03/04/2026", "en-US")).toBe(Date.UTC(2026, 2, 4));
    expect(parseSortDate("2026-04-03")).toBe(Date.UTC(2026, 3, 3));
    expect(parseSortDate("2026-04-03T10:00:00Z")).toBe(Date.UTC(2026, 3, 3));
    expect(parseSortDate("3 Apr 2026")).toBe(Date.UTC(2026, 3, 3));
    expect(parseSortDate("April 3, 2026")).toBe(Date.UTC(2026, 3, 3));
    expect(parseSortDate("not a date")).toBeNull();
    expect(parseSortDate("2026-13-01")).toBeNull();
  });
});
