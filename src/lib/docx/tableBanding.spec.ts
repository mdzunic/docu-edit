// SPEC-076 AC-12, AC-66, AC-67, AC-71 — the .docx half of the table look and
// style banding round trip.
import { describe, expect, it } from "vitest";
import { parseDocx } from "./reader";
import { buildDocumentXml, buildStylesXml, writeDocx } from "./writer";
import {
  createDocument,
  isTable,
  type DocumentModel,
  type NamedStyle,
  type Paragraph,
  type Table,
} from "../documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function cell(text: string) {
  return { blocks: [paragraph(text)] };
}

const BANDED_STYLE: NamedStyle = {
  id: "TableStyleAccent",
  name: "Accent",
  type: "table",
  cellBorders: {
    top: { style: "single", widthPt: 0.5, color: "#000000" },
    bottom: { style: "single", widthPt: 0.5, color: "#000000" },
    left: { style: "single", widthPt: 0.5, color: "#000000" },
    right: { style: "single", widthPt: 0.5, color: "#000000" },
  },
  banding: {
    firstRow: { shading: "#2f5496", bold: true },
    bandedRows: { size: 1, shading: "#d9e2f3" },
  },
};

function bandedModel(): DocumentModel {
  const model = createDocument("docx-banding-fixture");
  const table: Table = {
    type: "table",
    rows: [
      { isHeader: true, cells: [cell("Name"), cell("Qty")] },
      { cells: [cell("a"), cell("1")] },
      { cells: [cell("b"), cell("2")] },
    ],
    columnWidthsCm: [4, 4],
    styleId: "TableStyleAccent",
    look: { firstRow: true, bandedRows: true },
  };
  model.styles = [BANDED_STYLE];
  model.sections = [{ blocks: [paragraph("before"), table] }];
  return model;
}

describe("writeDocx — w:tblLook and w:tblStylePr (SPEC-076 AC-66)", () => {
  it("emits w:tblLook with noHBand/noVBand as the inverse of the banding flags", () => {
    const xml = buildDocumentXml(bandedModel());
    expect(xml).toContain('w:firstRow="1"');
    expect(xml).toContain('w:lastRow="0"');
    // bandedRows is on, so the "suppress horizontal bands" flag is off.
    expect(xml).toContain('w:noHBand="0"');
    expect(xml).toContain('w:noVBand="1"');
    expect(xml).toContain('<w:tblStyle w:val="TableStyleAccent"/>');
  });

  it("emits the style's regions as w:tblStylePr and its band sizes as w:tblStyle*BandSize", () => {
    const model = bandedModel();
    model.styles = [
      {
        ...BANDED_STYLE,
        banding: { ...BANDED_STYLE.banding, bandedRows: { size: 2, shading: "#d9e2f3" } },
      },
    ];
    const xml = buildStylesXml(model);
    expect(xml).toContain('<w:tblStylePr w:type="firstRow">');
    expect(xml).toContain('<w:tblStylePr w:type="band1Horz">');
    expect(xml).toContain('<w:tblStyleRowBandSize w:val="2"/>');
    // A column band the style does not declare emits nothing.
    expect(xml).not.toContain("band1Vert");
    expect(xml).not.toContain("tblStyleColBandSize");
  });
});

describe("parseDocx — the inverse (SPEC-076 AC-67)", () => {
  it("round-trips Table.look and NamedStyle.banding", async () => {
    const model = bandedModel();
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const table = reopened.sections[0].blocks[1];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    expect(table.styleId).toBe("TableStyleAccent");
    expect(table.look).toEqual({ firstRow: true, bandedRows: true });
    const style = reopened.styles.find((entry) => entry.id === "TableStyleAccent");
    expect(style?.banding?.firstRow).toEqual({ shading: "#2f5496", bold: true });
    expect(style?.banding?.bandedRows).toEqual({ size: 1, shading: "#d9e2f3" });
    expect(style?.cellBorders).toEqual(BANDED_STYLE.cellBorders);
  });

  it("keeps a band size other than 1 (AC-66)", async () => {
    const model = bandedModel();
    model.styles = [
      {
        ...BANDED_STYLE,
        banding: { ...BANDED_STYLE.banding, bandedRows: { size: 3, shading: "#d9e2f3" } },
      },
    ];
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const style = reopened.styles.find((entry) => entry.id === "TableStyleAccent");
    expect(style?.banding?.bandedRows).toEqual({ size: 3, shading: "#d9e2f3" });
  });

  it("leaves `look` unset for a table with no w:tblLook (AC-23)", async () => {
    const model = bandedModel();
    const table = model.sections[0].blocks[1] as Table;
    delete table.look;
    delete table.styleId;
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const reopenedTable = reopened.sections[0].blocks[1] as Table;
    expect(reopenedTable.look).toBeUndefined();
  });
});

describe("writeDocx — w:tblHeader for every header row (SPEC-076 AC-12)", () => {
  it("emits w:tblHeader for a NON-leading header row too", () => {
    const model = createDocument("docx-late-header");
    model.sections = [
      {
        blocks: [
          {
            type: "table",
            rows: [
              { cells: [cell("body")] },
              { isHeader: true, cells: [cell("late header")] },
            ],
            columnWidthsCm: [8],
          },
        ],
      },
    ];
    const xml = buildDocumentXml(model);
    // One header row in the model, one w:tblHeader in the output — on the
    // second row, not the first.
    expect(xml.split("<w:tblHeader/>")).toHaveLength(2);
    const rows = xml.slice(xml.indexOf("<w:tbl>")).split("<w:tr>");
    expect(rows[1]).not.toContain("<w:tblHeader/>");
    expect(rows[2]).toContain("<w:tblHeader/>");
  });
});
