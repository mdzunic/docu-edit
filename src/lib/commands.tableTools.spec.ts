// SPEC-076 — the model-level Table commands: applying a gallery look,
// toggling look flags, sorting and both conversions. Each is a pure
// model→model function, so none of this needs a component.
import { describe, expect, it } from "vitest";
import {
  applyTableStyleLook,
  convertibleParagraphRange,
  convertTableToText,
  convertTextToTable,
  setTableLook,
  sortTable,
} from "./commands.tableTools";
import { caretSelection } from "./commands";
import { equalStructural } from "./equality";
import { createDocument, type DocumentModel, type Paragraph, type Table } from "./documentModel";
import type { TableLocation } from "./tables";

const AT: TableLocation = { section: 0, block: 0, row: 0, col: 0 };
const SELECTION = caretSelection({ section: 0, block: 0, offset: 0 });

function cell(text: string) {
  return { blocks: [{ type: "paragraph" as const, runs: [{ text, formatting: {} }] }] };
}

function table(rows: string[][], headerRow = true): Table {
  return {
    type: "table",
    columnWidthsCm: rows[0].map(() => 4),
    rows: rows.map((texts, index) => ({
      cells: texts.map(cell),
      ...(headerRow && index === 0 ? { isHeader: true as const } : {}),
    })),
  };
}

function docWith(...blocks: (Table | Paragraph)[]): DocumentModel {
  const model = createDocument("table-tools-fixture");
  model.sections = [{ blocks }];
  return model;
}

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function tableOf(model: DocumentModel, block = 0): Table {
  return model.sections[0].blocks[block] as Table;
}

describe("applyTableStyleLook (SPEC-076 AC-33, AC-34)", () => {
  it("sets styleId and look and materialises the built-in style in one action", () => {
    const model = docWith(table([["a"], ["b"]]));
    const next = applyTableStyleLook(model, SELECTION, AT, "TableStyleBandedRows", {
      bandedRows: true,
    }).model;
    expect(tableOf(next).styleId).toBe("TableStyleBandedRows");
    expect(tableOf(next).look).toEqual({ bandedRows: true });
    const materialised = next.styles?.find((style) => style.id === "TableStyleBandedRows");
    expect(materialised?.type).toBe("table");
    expect(materialised?.banding?.bandedRows).toEqual({ size: 1, shading: "#f2f2f2" });
  });

  it("does not duplicate a style the document already carries", () => {
    const model = docWith(table([["a"], ["b"]]));
    model.styles = [{ id: "TableStyleGrid", name: "Grid", type: "table" }];
    const next = applyTableStyleLook(model, SELECTION, AT, "TableStyleGrid", {}).model;
    expect(next.styles?.filter((style) => style.id === "TableStyleGrid")).toHaveLength(1);
  });

  it("the None entry clears both styleId and look (AC-34)", () => {
    const model = docWith(table([["a"], ["b"]]));
    const applied = applyTableStyleLook(model, SELECTION, AT, "TableStyleAccent", {
      firstRow: true,
    }).model;
    const cleared = applyTableStyleLook(applied, SELECTION, AT, null, undefined).model;
    expect(tableOf(cleared).styleId).toBeUndefined();
    expect(tableOf(cleared).look).toBeUndefined();
  });

  it("leaves the source model untouched — the command is a snapshot edit (AC-33)", () => {
    const model = docWith(table([["a"], ["b"]]));
    const before = structuredClone(model);
    applyTableStyleLook(model, SELECTION, AT, "TableStyleGrid", {});
    expect(equalStructural(model, before)).toBe(true);
  });
});

describe("setTableLook (SPEC-076 AC-26, AC-72)", () => {
  it("never writes banding into a cell: toggling bandedRows on and off restores every cell exactly", () => {
    const model = docWith(table([["a", "b"], ["c", "d"], ["e", "f"]]));
    model.styles = [
      {
        id: "S",
        name: "Banded",
        type: "table",
        banding: { bandedRows: { size: 1, shading: "#f2f2f2" } },
      },
    ];
    tableOf(model).styleId = "S";
    const cellsBefore = structuredClone(tableOf(model).rows);
    const on = setTableLook(model, SELECTION, AT, { bandedRows: true }).model;
    const off = setTableLook(on, SELECTION, AT, {}).model;
    expect(equalStructural(tableOf(off).rows, cellsBefore)).toBe(true);
    // …and the cells were never touched while it was on, either.
    expect(equalStructural(tableOf(on).rows, cellsBefore)).toBe(true);
  });

  it("clears the preserved w:tblPr, because w:tblLook lives inside it (AC-72)", () => {
    const model = docWith(table([["a"], ["b"]]));
    tableOf(model).docxTblPr = "<w:tblPr><w:tblLook w:val=\"04A0\"/></w:tblPr>";
    const next = setTableLook(model, SELECTION, AT, { firstRow: true }).model;
    expect(tableOf(next).docxTblPr).toBeUndefined();
    expect(tableOf(next).look).toEqual({ firstRow: true });
  });
});

describe("sortTable (SPEC-076 AC-51)", () => {
  it("sorts the rows after the header run and returns one new model", () => {
    const model = docWith(table([["Name"], ["c"], ["a"], ["b"]]));
    const before = structuredClone(model);
    const next = sortTable(
      model,
      SELECTION,
      AT,
      [{ columnIndex: 0, direction: "asc", type: "text" }],
      true,
    ).model;
    expect(tableOf(next).rows.map((row) => row.cells[0].blocks[0].runs[0].text)).toEqual([
      "Name",
      "a",
      "b",
      "c",
    ]);
    // Undo is the snapshot before the edit — so the command must not have
    // mutated it (AC-51's "undo restores the original row order exactly").
    expect(equalStructural(model, before)).toBe(true);
  });
});

describe("convertibleParagraphRange (SPEC-076 AC-52)", () => {
  it("accepts a run of paragraphs", () => {
    const model = docWith(paragraph("a\tb"), paragraph("c\td"));
    const range = convertibleParagraphRange(model, {
      anchor: { section: 0, block: 0, offset: 0 },
      focus: { section: 0, block: 1, offset: 3 },
    });
    expect(range).toEqual({ section: 0, first: 0, last: 1 });
  });

  it("refuses a range containing a table", () => {
    const model = docWith(paragraph("a"), table([["x"]]), paragraph("b"));
    expect(
      convertibleParagraphRange(model, {
        anchor: { section: 0, block: 0, offset: 0 },
        focus: { section: 0, block: 2, offset: 1 },
      }),
    ).toBeNull();
  });

  it("refuses a caret inside a cell", () => {
    const model = docWith(table([["x"]]));
    expect(
      convertibleParagraphRange(
        model,
        caretSelection({ section: 0, block: 0, offset: 0, cell: { row: 0, col: 0, block: 0 } }),
      ),
    ).toBeNull();
  });
});

describe("convertTextToTable / convertTableToText (SPEC-076 AC-60, AC-64)", () => {
  it("replaces the selected paragraphs with the table, at their position", () => {
    const model = docWith(paragraph("keep"), paragraph("a\tb"), paragraph("c\td"));
    const next = convertTextToTable(model, {
      anchor: { section: 0, block: 1, offset: 0 },
      focus: { section: 0, block: 2, offset: 3 },
    }, "\t").model;
    expect(next.sections[0].blocks).toHaveLength(2);
    expect(next.sections[0].blocks[0].type).toBe("paragraph");
    const built = next.sections[0].blocks[1] as Table;
    expect(built.type).toBe("table");
    expect(built.rows).toHaveLength(2);
    expect(built.columnWidthsCm).toHaveLength(2);
  });

  it("replaces the table with its paragraphs, at its position", () => {
    const model = docWith(table([["a", "b"], ["c", "d"]], false), paragraph("after"));
    const next = convertTableToText(model, SELECTION, AT, "\t").model;
    expect(next.sections[0].blocks.map((block) => block.type)).toEqual([
      "paragraph",
      "paragraph",
      "paragraph",
    ]);
    expect((next.sections[0].blocks[0] as Paragraph).runs.map((run) => run.text).join("")).toBe("a\tb");
  });

  it("round-trips a simple tab-separated list through both commands (AC-65)", () => {
    const model = docWith(paragraph("a\tb"), paragraph("c\td"));
    const before = structuredClone(model.sections[0].blocks);
    const asTable = convertTextToTable(model, {
      anchor: { section: 0, block: 0, offset: 0 },
      focus: { section: 0, block: 1, offset: 3 },
    }, "\t").model;
    const back = convertTableToText(asTable, SELECTION, AT, "\t").model;
    expect(equalStructural(back.sections[0].blocks, before)).toBe(true);
  });
});
