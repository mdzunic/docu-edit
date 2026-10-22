// SPEC-076: the pure text ↔ table transforms. Nothing here mounts a
// component (AC-78).
import { describe, expect, it } from "vitest";
import {
  CONVERT_SEPARATORS,
  isValidCustomSeparator,
  separatorCharacter,
  tableToText,
  textToTable,
  textToTableSize,
} from "./tableConvert";
import { equalStructural } from "./equality";
import type { Paragraph, Table } from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function cellTexts(table: Table): string[][] {
  return table.rows.map((row) =>
    row.cells.map((cell) => cell.blocks.map((block) => block.runs.map((run) => run.text).join("")).join("")),
  );
}

function paragraphTexts(paragraphs: Paragraph[]): string[] {
  return paragraphs.map((entry) => entry.runs.map((run) => run.text).join(""));
}

describe("separator choices (AC-53, AC-55)", () => {
  it("offers Tab, Comma, Semicolon and Custom", () => {
    expect(CONVERT_SEPARATORS.map((entry) => entry.label)).toEqual(["Tab", "Comma", "Semicolon", "Custom"]);
    expect(separatorCharacter("tab", "")).toBe("\t");
    expect(separatorCharacter("comma", "")).toBe(",");
    expect(separatorCharacter("semicolon", "")).toBe(";");
    expect(separatorCharacter("custom", "|")).toBe("|");
  });

  it("rejects an empty, whitespace-only or multi-character custom separator", () => {
    expect(isValidCustomSeparator("|")).toBe(true);
    expect(isValidCustomSeparator("")).toBe(false);
    expect(isValidCustomSeparator(" ")).toBe(false);
    expect(isValidCustomSeparator("::")).toBe(false);
  });
});

describe("textToTable (AC-56..AC-59)", () => {
  it("takes as many columns as the paragraph with the most fields, padding shorter rows", () => {
    const table = textToTable([paragraph("a\tb\tc"), paragraph("d\te"), paragraph("f")], "\t");
    expect(table.columnWidthsCm).toHaveLength(3);
    expect(cellTexts(table)).toEqual([
      ["a", "b", "c"],
      ["d", "e", ""],
      ["f", "", ""],
    ]);
  });

  it("a paragraph ending in the separator yields a trailing empty cell", () => {
    const table = textToTable([paragraph("a\tb\t")], "\t");
    expect(table.columnWidthsCm).toHaveLength(3);
    expect(cellTexts(table)).toEqual([["a", "b", ""]]);
  });

  it("each field keeps the run formatting of the text it came from", () => {
    const source: Paragraph = {
      type: "paragraph",
      runs: [
        { text: "bold", formatting: { bold: true } },
        { text: "\tplain", formatting: {} },
        { text: "\tital", formatting: { italic: true } },
      ],
    };
    const table = textToTable([source], "\t");
    expect(table.rows[0].cells[0].blocks[0].runs[0].formatting).toEqual({ bold: true });
    expect(table.rows[0].cells[1].blocks[0].runs[0].formatting).toEqual({});
    expect(table.rows[0].cells[2].blocks[0].runs[0].formatting).toEqual({ italic: true });
  });

  it("a field spanning two runs keeps both", () => {
    const source: Paragraph = {
      type: "paragraph",
      runs: [
        { text: "one\tt", formatting: { bold: true } },
        { text: "wo", formatting: { italic: true } },
      ],
    };
    const table = textToTable([source], "\t");
    expect(table.rows[0].cells[1].blocks[0].runs).toEqual([
      { text: "t", formatting: { bold: true } },
      { text: "wo", formatting: { italic: true } },
    ]);
  });

  it("column widths divide the section's content width equally", () => {
    const table = textToTable([paragraph("a,b,c,d")], ",", 16);
    expect(table.columnWidthsCm).toHaveLength(4);
    const total = table.columnWidthsCm.reduce((sum, width) => sum + width, 0);
    expect(total).toBeCloseTo(16, 1);
    expect(new Set(table.columnWidthsCm).size).toBe(1);
  });

  it("previews the row and column count the conversion would produce (AC-54)", () => {
    expect(textToTableSize([paragraph("a,b"), paragraph("c,d,e")], ",")).toEqual({ rows: 2, columns: 3 });
    expect(textToTableSize([paragraph("a,b"), paragraph("c,d,e")], ";")).toEqual({ rows: 2, columns: 1 });
    expect(textToTableSize([], ",")).toEqual({ rows: 0, columns: 0 });
  });
});

describe("tableToText (AC-62, AC-63)", () => {
  it("emits one paragraph per row with cells joined by the separator", () => {
    const table = textToTable([paragraph("a\tb"), paragraph("c\td")], "\t");
    expect(paragraphTexts(tableToText(table, "\t"))).toEqual(["a\tb", "c\td"]);
  });

  it("takes the first paragraph of each cell for formatting and appends the rest inline", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [3, 3],
      rows: [
        {
          cells: [
            {
              blocks: [
                { type: "paragraph", runs: [{ text: "one", formatting: {} }], formatting: { alignment: "center" } },
                paragraph("two"),
              ],
            },
            { blocks: [paragraph("x")] },
          ],
        },
      ],
    };
    const result = tableToText(table, "|");
    expect(paragraphTexts(result)).toEqual(["onetwo|x"]);
    expect(result[0].formatting).toEqual({ alignment: "center" });
  });

  it("a cell holding a nested table contributes that table's flattened text in document order", () => {
    const nested = textToTable([paragraph("n1\tn2"), paragraph("n3\tn4")], "\t");
    const table: Table = {
      type: "table",
      columnWidthsCm: [3, 3],
      rows: [
        {
          cells: [
            { blocks: [paragraph("first"), nested] as Table["rows"][0]["cells"][0]["blocks"] },
            { blocks: [paragraph("after")] },
          ],
        },
      ],
    };
    expect(paragraphTexts(tableToText(table, "\t"))).toEqual(["firstn1\tn2\nn3\tn4\tafter"]);
  });

  it("covered cells of a merge contribute nothing", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [3, 3],
      rows: [{ cells: [{ blocks: [paragraph("wide")], colSpan: 2 }, { blocks: [], covered: true }] }],
    };
    expect(paragraphTexts(tableToText(table, "\t"))).toEqual(["wide"]);
  });
});

describe("round-trip stability for the simple case (AC-65)", () => {
  // The simple case (D-13): every paragraph has the same field count, no
  // field's text contains the separator, and no paragraph is empty. The
  // paragraphs are in the model's normalised run form — adjacent runs with
  // identical formatting already coalesced, which is the only form the
  // editor ever produces (textEditing.normalizeRuns).
  it("tableToText(textToTable(p, s), s) is equalStructural to p", () => {
    const paragraphs: Paragraph[] = [
      { type: "paragraph", runs: [{ text: "Name\tRole", formatting: { bold: true } }] },
      {
        type: "paragraph",
        runs: [
          { text: "Ada\t", formatting: {} },
          { text: "Engineer", formatting: { italic: true } },
        ],
      },
    ];
    const round = tableToText(textToTable(paragraphs, "\t"), "\t");
    expect(equalStructural(round, paragraphs)).toBe(true);
  });

  it("holds for a comma-separated list too", () => {
    const paragraphs = [paragraph("a,b,c"), paragraph("d,e,f")];
    const round = tableToText(textToTable(paragraphs, ","), ",");
    expect(equalStructural(round, paragraphs)).toBe(true);
  });
});
