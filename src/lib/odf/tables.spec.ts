// SPEC-022 AC-25: tables through the .odt reader and writer — table:table
// with column definitions, preserving row/column counts and cell content;
// tables the model cannot express stay opaque (SPEC-008).
import { describe, expect, it } from "vitest";
import { parseOdt } from "./reader";
import { buildContentXml, writeOdt } from "./writer";
import { zip } from "../docx/zipWriter";
import { firstStructuralDifference } from "../equality";
import { blockText } from "../textEditing";
import {
  createDocument,
  isTable,
  type DocumentModel,
  type Paragraph,
  type Table,
} from "../documentModel";

const encoder = new TextEncoder();

async function makeOdt(contentXml: string): Promise<Uint8Array> {
  return zip([
    {
      name: "mimetype",
      data: encoder.encode("application/vnd.oasis.opendocument.text"),
      store: true,
    },
    { name: "content.xml", data: encoder.encode(contentXml) },
  ]);
}

function contentWith(body: string, automaticStyles = ""): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
    ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
    ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
    ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
    ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">' +
    `<office:automatic-styles>${automaticStyles}</office:automatic-styles>` +
    `<office:body><office:text>${body}</office:text></office:body>` +
    "</office:document-content>"
  );
}

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function tableModel(): DocumentModel {
  const model = createDocument("odt-table-fixture");
  const table: Table = {
    type: "table",
    rows: [
      { cells: [{ blocks: [paragraph("A1")] }, { blocks: [paragraph("B1")] }] },
      { cells: [{ blocks: [paragraph("")] }, { blocks: [paragraph("first"), paragraph("second")] }] },
    ],
    columnWidthsCm: [7.96, 7.96],
  };
  model.sections = [{ blocks: [paragraph("before"), table, paragraph("after")] }];
  return model;
}

describe("writeOdt — table:table emission (AC-25)", () => {
  it("emits column definitions whose styles carry the widths", () => {
    const xml = buildContentXml(tableModel());
    expect(xml).toContain("<table:table");
    expect(xml).toContain('table:style-name="Table1"');
    expect(xml).toContain('style:family="table-column"');
    expect(xml).toContain('style:column-width="7.96cm"');
    expect((xml.match(/<table:table-column /g) ?? []).length).toBe(2);
    expect((xml.match(/<table:table-row/g) ?? []).length).toBe(2);
    expect(xml).toContain("A1");
    expect(xml).toContain("second");
  });
});

describe("odt table round trip (AC-25)", () => {
  it("save + re-open reproduces rows, columns, widths and cell content", async () => {
    const { bytes, dropped } = await writeOdt(tableModel());
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    const table = reopened.sections[0].blocks[1];
    expect(isTable(table)).toBe(true);
    if (!isTable(table)) {
      return;
    }
    expect(table.rows).toHaveLength(2);
    expect(table.columnWidthsCm.map((cm) => Math.round(cm * 100) / 100)).toEqual([7.96, 7.96]);
    expect(table.rows[0].cells.map((cell) => cell.blocks.map(blockText).join("|"))).toEqual([
      "A1",
      "B1",
    ]);
    expect(table.rows[1].cells.map((cell) => cell.blocks.map(blockText).join("|"))).toEqual([
      "",
      "first|second",
    ]);
    // A second save + re-open is exactly stable (style names now preserved).
    const model2: DocumentModel = {
      ...createDocument("again"),
      sections: reopened.sections,
      defaultStyles: reopened.defaultStyles,
      format: "odt",
      opaque: reopened.opaque,
      sourceXml: reopened.sourceXml,
    };
    const again = await parseOdt((await writeOdt(model2)).bytes);
    expect(firstStructuralDifference(reopened.sections, again.sections)).toBeNull();
  });
});

describe("parseOdt — table shapes (AC-25, SPEC-008 fallback)", () => {
  const columnStyles =
    '<style:style style:name="T.A" style:family="table-column">' +
    '<style:table-column-properties style:column-width="4cm"/></style:style>' +
    '<style:style style:name="T.B" style:family="table-column">' +
    '<style:table-column-properties style:column-width="6cm"/></style:style>';

  it("reads column widths from the referenced column styles", async () => {
    const body =
      '<table:table table:name="T" table:style-name="T">' +
      '<table:table-column table:style-name="T.A"/>' +
      '<table:table-column table:style-name="T.B"/>' +
      "<table:table-row>" +
      '<table:table-cell office:value-type="string"><text:p>left</text:p></table:table-cell>' +
      '<table:table-cell office:value-type="string"><text:p>right</text:p></table:table-cell>' +
      "</table:table-row>" +
      "</table:table><text:p/>";
    const parsed = await parseOdt(await makeOdt(contentWith(body, columnStyles)));
    const table = parsed.sections[0].blocks[0];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    expect(table.columnWidthsCm).toEqual([4, 6]);
    expect(table.odtName).toBe("T");
    expect(table.odtColumnStyleNames).toEqual(["T.A", "T.B"]);
    expect(table.rows[0].cells.map((cell) => blockText(cell.blocks[0]))).toEqual(["left", "right"]);
  });

  it("expands number-columns-repeated on columns and cells", async () => {
    const body =
      "<table:table>" +
      '<table:table-column table:style-name="T.A" table:number-columns-repeated="3"/>' +
      "<table:table-row>" +
      '<table:table-cell table:number-columns-repeated="3"><text:p>x</text:p></table:table-cell>' +
      "</table:table-row>" +
      "</table:table><text:p/>";
    const parsed = await parseOdt(await makeOdt(contentWith(body, columnStyles)));
    const table = parsed.sections[0].blocks[0];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    expect(table.columnWidthsCm).toHaveLength(3);
    expect(table.rows[0].cells.map((cell) => blockText(cell.blocks[0]))).toEqual(["x", "x", "x"]);
  });

  it("models a horizontally merged cell's span instead of flattening it (SPEC-023 AC-28)", async () => {
    const body =
      "<table:table>" +
      "<table:table-column/><table:table-column/>" +
      "<table:table-row>" +
      '<table:table-cell table:number-columns-spanned="2"><text:p>wide</text:p></table:table-cell>' +
      "<table:covered-table-cell/>" +
      "</table:table-row>" +
      "</table:table><text:p/>";
    const parsed = await parseOdt(await makeOdt(contentWith(body)));
    expect(parsed.sections[0].opaqueBlocks ?? []).toHaveLength(0);
    const table = parsed.sections[0].blocks[0];
    expect(isTable(table)).toBe(true);
    if (!isTable(table)) {
      return;
    }
    expect(table.rows[0].cells[0].colSpan).toBe(2);
    expect(table.rows[0].cells[0].blocks.map(blockText)).toEqual(["wide"]);
    expect(table.rows[0].cells[1].covered).toBe(true);
  });

  it("keeps a header-rows table opaque (repetition is out of scope)", async () => {
    const body =
      "<table:table>" +
      "<table:table-column/>" +
      "<table:table-header-rows><table:table-row><table:table-cell><text:p>h</text:p></table:table-cell></table:table-row></table:table-header-rows>" +
      "<table:table-row><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row>" +
      "</table:table><text:p/>";
    const parsed = await parseOdt(await makeOdt(contentWith(body)));
    expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
  });

  it("falls back to equal columns over the text area when widths are undeclared", async () => {
    const body =
      "<table:table>" +
      "<table:table-column/><table:table-column/>" +
      "<table:table-row>" +
      "<table:table-cell><text:p>a</text:p></table:table-cell>" +
      "<table:table-cell><text:p>b</text:p></table:table-cell>" +
      "</table:table-row>" +
      "</table:table><text:p/>";
    const parsed = await parseOdt(await makeOdt(contentWith(body)));
    const table = parsed.sections[0].blocks[0];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    expect(table.columnWidthsCm).toHaveLength(2);
    expect(table.columnWidthsCm[0]).toBeCloseTo(table.columnWidthsCm[1], 6);
    expect(table.columnWidthsCm[0]).toBeGreaterThan(0);
  });

  it("parses lists inside cells as the cell's own list paragraphs (AC-8)", async () => {
    const body =
      "<table:table>" +
      "<table:table-column/>" +
      "<table:table-row><table:table-cell>" +
      "<text:list><text:list-item><text:p>one</text:p></text:list-item>" +
      "<text:list-item><text:p>two</text:p></text:list-item></text:list>" +
      "</table:table-cell></table:table-row>" +
      "</table:table><text:p/>";
    const parsed = await parseOdt(await makeOdt(contentWith(body)));
    const table = parsed.sections[0].blocks[0];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    const blocks = table.rows[0].cells[0].blocks;
    expect(blocks.map(blockText)).toEqual(["one", "two"]);
    expect(blocks.every((block) => block.list?.level === 1)).toBe(true);
  });
});
