// SPEC-076: the four-level cell-formatting resolution, band counting, and
// the built-in gallery looks. Pure — nothing here mounts a component (AC-78).
import { describe, expect, it } from "vitest";
import {
  BUILT_IN_TABLE_STYLES,
  builtInTableStyle,
  leadingHeaderRowCount,
  resolveCellFormatting,
  styleHasBanding,
  tableStyleFor,
} from "./tableStyles";
import type { BorderSide, CellBorders } from "./core/tables";
import type { NamedStyle, Table } from "./documentModel";

function side(color: string): BorderSide {
  return { style: "single", widthPt: 0.5, color };
}

function allSides(color: string): CellBorders {
  return { top: side(color), bottom: side(color), left: side(color), right: side(color) };
}

function grid(rows: number, cols: number): Table {
  return {
    type: "table",
    rows: Array.from({ length: rows }, () => ({
      cells: Array.from({ length: cols }, () => ({ blocks: [{ type: "paragraph" as const, runs: [] }] })),
    })),
    columnWidthsCm: Array.from({ length: cols }, () => 3),
  };
}

describe("resolveCellFormatting — the four-level resolution order (AC-24, AC-25)", () => {
  const style: NamedStyle = {
    id: "S",
    type: "table",
    cellShadingFill: "#111111",
    cellBorders: allSides("#111111"),
    banding: {
      bandedRows: { size: 1, shading: "#222222" },
      firstRow: { shading: "#333333", bold: true, borders: { top: side("#333333") } },
    },
  };

  it("level 1: the style's base borders and shading apply with no look at all", () => {
    const table = grid(3, 2);
    table.styleId = "S";
    const resolved = resolveCellFormatting(table, style, 1, 0);
    expect(resolved.shadingFill).toBe("#111111");
    expect(resolved.borders?.top?.color).toBe("#111111");
  });

  it("level 2 overrides level 1: an enabled band shades over the style's base", () => {
    const table = grid(3, 2);
    table.styleId = "S";
    table.look = { bandedRows: true };
    // Rows 0 and 2 are band 1 and band 3 (odd); row 1 is band 2.
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBe("#222222");
    expect(resolveCellFormatting(table, style, 1, 0).shadingFill).toBe("#111111");
    expect(resolveCellFormatting(table, style, 2, 0).shadingFill).toBe("#222222");
  });

  it("level 3 overrides level 2: firstRow wins over the row band at row 0", () => {
    const table = grid(3, 2);
    table.styleId = "S";
    table.look = { bandedRows: true, firstRow: true };
    const resolved = resolveCellFormatting(table, style, 0, 0);
    expect(resolved.shadingFill).toBe("#333333");
    expect(resolved.bold).toBe(true);
    expect(resolved.borders?.top?.color).toBe("#333333");
    // The sides firstRow does not set still come from the style's base.
    expect(resolved.borders?.bottom?.color).toBe("#111111");
  });

  it("level 4 overrides everything: the cell's own shading and borders win", () => {
    const table = grid(3, 2);
    table.styleId = "S";
    table.look = { bandedRows: true, firstRow: true };
    table.rows[0].cells[0].shadingFill = "#444444";
    table.rows[0].cells[0].borders = allSides("#444444");
    const resolved = resolveCellFormatting(table, style, 0, 0);
    expect(resolved.shadingFill).toBe("#444444");
    expect(resolved.borders?.top?.color).toBe("#444444");
    expect(resolved.borders?.bottom?.color).toBe("#444444");
  });

  it("a table with no style resolves to the cell's own formatting only", () => {
    const table = grid(2, 2);
    table.rows[1].cells[1].shadingFill = "#abcdef";
    expect(resolveCellFormatting(table, undefined, 1, 1)).toEqual({ shadingFill: "#abcdef" });
    expect(resolveCellFormatting(table, undefined, 0, 0)).toEqual({});
  });

  it("a look whose flags are all false applies no banding (AC-23)", () => {
    const table = grid(3, 2);
    table.styleId = "S";
    table.look = { bandedRows: false, firstRow: false };
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBe("#111111");
  });
});

describe("resolveCellFormatting — banding is never written into cells (AC-26)", () => {
  it("toggling look.bandedRows on and off leaves every cell's shading and borders identical", () => {
    const table = grid(4, 3);
    table.styleId = "S";
    table.rows[2].cells[1].shadingFill = "#0f0f0f";
    table.rows[2].cells[1].borders = allSides("#0f0f0f");
    const style: NamedStyle = {
      id: "S",
      type: "table",
      banding: { bandedRows: { size: 1, shading: "#eeeeee" } },
    };
    const before = JSON.stringify(table.rows);

    table.look = { bandedRows: true };
    for (let r = 0; r < table.rows.length; r += 1) {
      for (let c = 0; c < 3; c += 1) {
        resolveCellFormatting(table, style, r, c);
      }
    }
    delete table.look;
    for (let r = 0; r < table.rows.length; r += 1) {
      for (let c = 0; c < 3; c += 1) {
        resolveCellFormatting(table, style, r, c);
      }
    }

    expect(JSON.stringify(table.rows)).toBe(before);
  });
});

describe("band counting (AC-27, AC-28)", () => {
  const style: NamedStyle = {
    id: "S",
    type: "table",
    banding: {
      bandedRows: { size: 1, shading: "#band" },
      bandedColumns: { size: 1, shading: "#col" },
      firstRow: { shading: "#first" },
      lastRow: { shading: "#last" },
      firstColumn: { shading: "#firstcol" },
      lastColumn: { shading: "#lastcol" },
    },
  };

  it("row bands start after the leading isHeader run", () => {
    const table = grid(5, 1);
    table.rows[0].isHeader = true;
    table.look = { bandedRows: true };
    // Counting starts at row 1: bands are rows 1, 3 (odd) and 2, 4 (even).
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBeUndefined();
    expect(resolveCellFormatting(table, style, 1, 0).shadingFill).toBe("#band");
    expect(resolveCellFormatting(table, style, 2, 0).shadingFill).toBeUndefined();
    expect(resolveCellFormatting(table, style, 3, 0).shadingFill).toBe("#band");
  });

  it("row bands start after an enabled firstRow", () => {
    const table = grid(4, 1);
    table.look = { bandedRows: true, firstRow: true };
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBe("#first");
    expect(resolveCellFormatting(table, style, 1, 0).shadingFill).toBe("#band");
    expect(resolveCellFormatting(table, style, 2, 0).shadingFill).toBeUndefined();
  });

  it("bandedRows.size groups rows into blocks of that size", () => {
    const table = grid(8, 1);
    table.look = { bandedRows: true };
    const style2: NamedStyle = { id: "S", type: "table", banding: { bandedRows: { size: 2, shading: "#band" } } };
    const shaded = table.rows.map((_, r) => resolveCellFormatting(table, style2, r, 0).shadingFill === "#band");
    expect(shaded).toEqual([true, true, false, false, true, true, false, false]);
  });

  it("column bands follow the same rule from the first column not consumed by firstColumn", () => {
    const table = grid(1, 5);
    table.look = { bandedColumns: true, firstColumn: true };
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBe("#firstcol");
    expect(resolveCellFormatting(table, style, 0, 1).shadingFill).toBe("#col");
    expect(resolveCellFormatting(table, style, 0, 2).shadingFill).toBeUndefined();
    expect(resolveCellFormatting(table, style, 0, 3).shadingFill).toBe("#col");
  });

  it("an enabled lastRow takes the final row and wins over banding there", () => {
    const table = grid(4, 1);
    table.look = { bandedRows: true, lastRow: true };
    // Band counting stops before the last row, and lastRow paints it.
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBe("#band");
    expect(resolveCellFormatting(table, style, 2, 0).shadingFill).toBe("#band");
    expect(resolveCellFormatting(table, style, 3, 0).shadingFill).toBe("#last");
  });

  it("an enabled lastColumn takes the final column and wins over banding there", () => {
    const table = grid(1, 4);
    table.look = { bandedColumns: true, lastColumn: true };
    expect(resolveCellFormatting(table, style, 0, 0).shadingFill).toBe("#col");
    expect(resolveCellFormatting(table, style, 0, 3).shadingFill).toBe("#lastcol");
  });
});

describe("leadingHeaderRowCount (AC-4, AC-5)", () => {
  it("counts the contiguous leading run", () => {
    const table = grid(4, 1);
    table.rows[0].isHeader = true;
    table.rows[1].isHeader = true;
    expect(leadingHeaderRowCount(table)).toBe(2);
  });

  it("is zero when the first row is not a header, even if a later row is", () => {
    const table = grid(4, 1);
    table.rows[2].isHeader = true;
    expect(leadingHeaderRowCount(table)).toBe(0);
  });
});

describe("the built-in gallery (AC-31)", () => {
  it("ships the five looks in the Reference, each materialisable", () => {
    expect(BUILT_IN_TABLE_STYLES.map((entry) => entry.name)).toEqual([
      "Plain",
      "Grid",
      "Banded rows",
      "Header emphasis",
      "Accent",
    ]);
    for (const entry of BUILT_IN_TABLE_STYLES) {
      const style = builtInTableStyle(entry.id);
      expect(style?.id).toBe(entry.id);
      expect(style?.type).toBe("table");
    }
  });

  it("Plain has no borders and no banding; Grid has borders and no banding", () => {
    expect(styleHasBanding(builtInTableStyle("TableStylePlain"))).toBe(false);
    expect(builtInTableStyle("TableStylePlain")?.cellBorders).toBeUndefined();
    expect(builtInTableStyle("TableStyleGrid")?.cellBorders?.top?.widthPt).toBe(0.5);
    expect(styleHasBanding(builtInTableStyle("TableStyleGrid"))).toBe(false);
  });

  it("Accent bands its rows and emphasises its first row", () => {
    const accent = builtInTableStyle("TableStyleAccent");
    expect(accent?.banding?.bandedRows).toEqual({ size: 1, shading: "#d9e2f3" });
    expect(accent?.banding?.firstRow).toEqual({ shading: "#2f5496", bold: true });
    expect(BUILT_IN_TABLE_STYLES.find((entry) => entry.id === "TableStyleAccent")?.look).toEqual({
      firstRow: true,
      bandedRows: true,
    });
  });

  it("two materialisations never alias the same border object", () => {
    const a = builtInTableStyle("TableStyleGrid");
    const b = builtInTableStyle("TableStyleGrid");
    expect(a?.cellBorders).not.toBe(b?.cellBorders);
    expect(a?.cellBorders?.top).not.toBe(b?.cellBorders?.top);
  });

  it("an unknown id materialises nothing", () => {
    expect(builtInTableStyle("NotAStyle")).toBeUndefined();
  });
});

describe("tableStyleFor", () => {
  it("finds the table's catalogue entry, or undefined without a styleId", () => {
    const styles: NamedStyle[] = [{ id: "S", type: "table" }];
    const table = grid(1, 1);
    expect(tableStyleFor(styles, table)).toBeUndefined();
    table.styleId = "S";
    expect(tableStyleFor(styles, table)?.id).toBe("S");
    table.styleId = "Missing";
    expect(tableStyleFor(styles, table)).toBeUndefined();
  });
});
