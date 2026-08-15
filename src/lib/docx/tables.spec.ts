// SPEC-022 AC-24, AC-26..28: tables through the .docx reader and writer —
// w:tbl with a grid and cell widths, preserving row/column counts and cell
// content, with empty and multi-paragraph cells; tables the model cannot
// express stay opaque (SPEC-008).
import { describe, expect, it } from "vitest";
import { parseDocx } from "./reader";
import { buildDocumentXml, writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { zip } from "../core/zipWriter";
import { firstStructuralDifference } from "../equality";
import { blockText } from "../textEditing";
import {
  createDocument,
  isTable,
  type DocumentModel,
  type Paragraph,
  type Table,
} from "../documentModel";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function tableModel(): DocumentModel {
  const model = createDocument("docx-table-fixture");
  const table: Table = {
    type: "table",
    rows: [
      { cells: [{ blocks: [paragraph("A1")] }, { blocks: [paragraph("B1")] }] },
      {
        cells: [
          // An empty cell (AC-27) and a multi-paragraph cell (AC-28).
          { blocks: [paragraph("")] },
          { blocks: [paragraph("first"), paragraph("second")] },
        ],
      },
    ],
    columnWidthsCm: [7.96, 7.96],
  };
  model.sections = [{ blocks: [paragraph("before"), table, paragraph("after")] }];
  return model;
}

const encoder = new TextEncoder();

async function makeDocx(documentXml: string): Promise<Uint8Array> {
  return zip([{ name: "word/document.xml", data: encoder.encode(documentXml) }]);
}

function bodyDocx(body: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`
  );
}

describe("writeDocx — w:tbl emission (AC-24)", () => {
  it("emits tblGrid column widths and per-cell tcW in twips", () => {
    const xml = buildDocumentXml(tableModel());
    expect(xml).toContain("<w:tbl>");
    // 7.96 cm = 4513 twips (rounded).
    expect(xml).toContain('<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>');
    expect(xml).toContain('<w:tcW w:w="4513" w:type="dxa"/>');
    // Rows and cells in order, the empty cell as an empty w:p.
    expect(xml).toContain("A1");
    expect(xml).toContain("first");
    expect(xml).toContain("second");
  });

  it("produces a package that passes the writer's own validator", async () => {
    const { bytes } = await writeDocx(tableModel());
    expect(await validateDocx(bytes)).toEqual([]);
  });
});

describe("docx table round trip (AC-24, AC-26..28)", () => {
  it("save + re-open reproduces rows, columns, widths and cell content", async () => {
    const model = tableModel();
    const { bytes, dropped } = await writeDocx(model);
    expect(dropped).toEqual([]);
    const reopened = await parseDocx(bytes);
    const table = reopened.sections[0].blocks[1];
    expect(isTable(table)).toBe(true);
    if (!isTable(table)) {
      return;
    }
    expect(table.rows).toHaveLength(2);
    expect(table.columnWidthsCm).toHaveLength(2);
    expect(table.columnWidthsCm[0]).toBeCloseTo(7.96, 2);
    expect(table.rows[0].cells.map((cell) => cell.blocks.map(blockText).join("|"))).toEqual([
      "A1",
      "B1",
    ]);
    expect(table.rows[1].cells.map((cell) => cell.blocks.map(blockText).join("|"))).toEqual([
      "",
      "first|second",
    ]);
    // A second save + re-open is exactly stable (the generated tblPr/tcPr
    // are preserved verbatim from the first save).
    const model2: DocumentModel = {
      ...createDocument("again"),
      sections: reopened.sections,
      defaultStyles: reopened.defaultStyles,
      format: "docx",
      opaque: reopened.opaque,
      sourceXml: reopened.sourceXml,
    };
    const again = await parseDocx((await writeDocx(model2)).bytes);
    expect(firstStructuralDifference(reopened.sections, again.sections)).toBeNull();
  });
});

describe("parseDocx — table shapes (AC-26..28, SPEC-008 fallback)", () => {
  const simpleTable =
    "<w:tbl>" +
    "<w:tblPr><w:tblW w:w=\"9026\" w:type=\"dxa\"/></w:tblPr>" +
    "<w:tblGrid><w:gridCol w:w=\"4513\"/><w:gridCol w:w=\"4513\"/></w:tblGrid>" +
    "<w:tr>" +
    "<w:tc><w:tcPr><w:tcW w:w=\"4513\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>" +
    "<w:tc><w:tcPr><w:tcW w:w=\"4513\" w:type=\"dxa\"/></w:tcPr><w:p/></w:tc>" +
    "</w:tr>" +
    "<w:tr>" +
    "<w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc>" +
    "<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>" +
    "</w:tr>" +
    "</w:tbl>";

  it("models a regular table, including empty and multi-paragraph cells", async () => {
    const parsed = await parseDocx(await makeDocx(bodyDocx(`<w:p/>${simpleTable}<w:p/>`)));
    const table = parsed.sections[0].blocks[1];
    expect(isTable(table)).toBe(true);
    if (!isTable(table)) {
      return;
    }
    expect(table.rows).toHaveLength(2);
    expect(table.columnWidthsCm).toHaveLength(2);
    expect(table.rows[0].cells[0].blocks.map(blockText)).toEqual(["A1"]);
    // The empty cell holds one empty paragraph so a caret can land in it.
    expect(table.rows[0].cells[1].blocks.map(blockText)).toEqual([""]);
    expect(table.rows[1].cells[0].blocks.map(blockText)).toEqual(["one", "two"]);
    // The source's own property bags survive verbatim.
    expect(table.docxTblPr).toContain("tblW");
    expect(table.rows[0].cells[0].docxTcPr).toContain("tcW");
    expect(parsed.sections[0].opaqueBlocks ?? []).toHaveLength(0);
  });

  it("models a horizontally merged cell's gridSpan instead of flattening it (SPEC-023 AC-24)", async () => {
    const merged =
      "<w:tbl>" +
      "<w:tblGrid><w:gridCol w:w=\"4513\"/><w:gridCol w:w=\"4513\"/></w:tblGrid>" +
      "<w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr>" +
      "</w:tbl>";
    const parsed = await parseDocx(await makeDocx(bodyDocx(`<w:p/>${merged}`)));
    expect(parsed.sections[0].opaqueBlocks ?? []).toHaveLength(0);
    const table = parsed.sections[0].blocks[1];
    expect(isTable(table)).toBe(true);
    if (!isTable(table)) {
      return;
    }
    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.rows[0].cells[0].colSpan).toBe(2);
    expect(table.rows[0].cells[0].blocks.map(blockText)).toEqual(["wide"]);
    expect(table.rows[0].cells[1].covered).toBe(true);
  });

  it("keeps a table with a nested table opaque", async () => {
    const nested =
      "<w:tbl>" +
      "<w:tblGrid><w:gridCol w:w=\"9026\"/></w:tblGrid>" +
      "<w:tr><w:tc>" +
      "<w:tbl><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>" +
      "<w:p/>" +
      "</w:tc></w:tr>" +
      "</w:tbl>";
    const parsed = await parseDocx(await makeDocx(bodyDocx(`<w:p/>${nested}`)));
    expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
  });

  it("keeps a table whose grid does not match its rows opaque", async () => {
    const ragged =
      "<w:tbl>" +
      "<w:tblGrid><w:gridCol w:w=\"4513\"/></w:tblGrid>" +
      "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>" +
      "</w:tbl>";
    const parsed = await parseDocx(await makeDocx(bodyDocx(`<w:p/>${ragged}`)));
    expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
  });

  it("resolves lists and styled text inside cells (AC-6..AC-9)", async () => {
    const styled =
      "<w:tbl>" +
      "<w:tblGrid><w:gridCol w:w=\"9026\"/></w:tblGrid>" +
      "<w:tr><w:tc><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
      "<w:r><w:rPr><w:b/></w:rPr><w:t>bold centred</w:t></w:r></w:p></w:tc></w:tr>" +
      "</w:tbl>";
    const parsed = await parseDocx(await makeDocx(bodyDocx(`<w:p/>${styled}`)));
    const table = parsed.sections[0].blocks[1];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    const cellParagraph = table.rows[0].cells[0].blocks[0];
    expect(cellParagraph.formatting?.alignment).toBe("center");
    expect(cellParagraph.runs[0].formatting.bold).toBe(true);
  });

  it("models a row's w:tblHeader as isHeader (SPEC-043)", async () => {
    const headered =
      "<w:tbl>" +
      "<w:tblGrid><w:gridCol w:w=\"9026\"/></w:tblGrid>" +
      "<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>" +
      "<w:tr><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>" +
      "</w:tbl>";
    const parsed = await parseDocx(await makeDocx(bodyDocx(`<w:p/>${headered}`)));
    const table = parsed.sections[0].blocks[1];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    expect(table.rows[0].isHeader).toBe(true);
    expect(table.rows[1].isHeader).toBeUndefined();
  });
});

describe("writeDocx — w:tblHeader emission (SPEC-043)", () => {
  it("emits w:tblHeader for a row marked isHeader and round-trips it", async () => {
    const model = createDocument("docx-header-row");
    const table: Table = {
      type: "table",
      rows: [
        { cells: [{ blocks: [paragraph("h")] }], isHeader: true },
        { cells: [{ blocks: [paragraph("b")] }] },
      ],
      columnWidthsCm: [9.0],
    };
    model.sections = [{ blocks: [paragraph("before"), table] }];
    const xml = buildDocumentXml(model);
    expect(xml).toContain("<w:tblHeader/>");
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const reopenedTable = reopened.sections[0].blocks[1];
    if (!isTable(reopenedTable)) {
      throw new Error("expected a table");
    }
    expect(reopenedTable.rows[0].isHeader).toBe(true);
    expect(reopenedTable.rows[1].isHeader).toBeUndefined();
  });
});

describe("writeDocx — lists inside cells reference a numbering plan (AC-8)", () => {
  it("gives a cell's list paragraphs their own numPr", async () => {
    const model = createDocument("cell-lists");
    const item: Paragraph = { ...paragraph("item"), list: { type: "number", level: 1 } };
    const table: Table = {
      type: "table",
      rows: [{ cells: [{ blocks: [item, { ...paragraph("two"), list: { type: "number", level: 1 } }] }] }],
      columnWidthsCm: [10],
    };
    model.sections = [{ blocks: [table, paragraph("")] }];
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const reopenedTable = reopened.sections[0].blocks[0];
    if (!isTable(reopenedTable)) {
      throw new Error("expected a table");
    }
    const items = reopenedTable.rows[0].cells[0].blocks;
    expect(items.map((p) => p.list?.type)).toEqual(["number", "number"]);
  });
});
