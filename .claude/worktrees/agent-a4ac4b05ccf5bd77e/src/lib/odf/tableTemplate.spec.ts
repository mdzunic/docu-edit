// SPEC-076 AC-13, AC-14, AC-68, AC-69, AC-70 — the .odt half of the table
// look and style banding round trip, plus the two losses ODF's vocabulary
// forces and the drop records that report them.
import { describe, expect, it } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import { collectOdfTableDropRecords } from "../opaque";
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
  banding: {
    firstRow: { shading: "#2f5496", bold: true },
    bandedRows: { size: 1, shading: "#d9e2f3" },
  },
};

function bandedModel(): DocumentModel {
  const model = createDocument("odt-banding-fixture");
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

/** Save and reopen, restamping the id/title/format the app layer owns so the
 * result is a full DocumentModel again. */
async function roundTrip(model: DocumentModel): Promise<DocumentModel> {
  const { bytes } = await writeOdt(model);
  return { id: model.id, title: model.title, format: "odt", ...(await parseOdt(bytes)) } as DocumentModel;
}

describe("writeOdt / parseOdt — table:use-*-styles and table:table-template (AC-68, AC-69)", () => {
  it("round-trips Table.look and the style's banding", async () => {
    const reopened = await roundTrip(bandedModel());
    const table = reopened.sections[0].blocks[1];
    if (!isTable(table)) {
      throw new Error("expected a table");
    }
    expect(table.styleId).toBe("TableStyleAccent");
    expect(table.look).toEqual({ firstRow: true, bandedRows: true });
    const style = (reopened.styles ?? []).find((entry) => entry.id === "TableStyleAccent");
    expect(style?.banding?.firstRow).toEqual({ shading: "#2f5496", bold: true });
    expect(style?.banding?.bandedRows).toEqual({ size: 1, shading: "#d9e2f3" });
    // A style whose conditional formatting is all banding is modelled, so
    // Modify is available for it (AC-71's ODF half).
    expect(style?.unmodeled).toBeUndefined();
  });

  it("round-trips a band region's bold and borders through the template (AC-37)", async () => {
    const model = bandedModel();
    model.styles = [
      {
        ...BANDED_STYLE,
        banding: {
          ...BANDED_STYLE.banding,
          bandedRows: {
            size: 1,
            shading: "#d9e2f3",
            bold: true,
            borders: {
              top: { style: "dashed" as const, widthPt: 1.5, color: "#ff0000" },
              bottom: { style: "dashed" as const, widthPt: 1.5, color: "#ff0000" },
              left: { style: "dashed" as const, widthPt: 1.5, color: "#ff0000" },
              right: { style: "dashed" as const, widthPt: 1.5, color: "#ff0000" },
            },
          },
        },
      },
    ];
    const reopened = await roundTrip(model);
    const style = (reopened.styles ?? []).find((entry) => entry.id === "TableStyleAccent");
    expect(style?.banding?.bandedRows).toEqual(model.styles[0].banding?.bandedRows);
    expect(style?.unmodeled).toBeUndefined();
  });

  it("leaves `look` unset for a table declaring no table:use-* attributes (AC-23)", async () => {
    const model = bandedModel();
    const table = model.sections[0].blocks[1] as Table;
    delete table.look;
    delete table.styleId;
    model.styles = [];
    const reopened = await roundTrip(model);
    expect((reopened.sections[0].blocks[1] as Table).look).toBeUndefined();
  });

  it("keeps only a LEADING run of header rows in table:table-header-rows (AC-13, AC-15)", async () => {
    const model = bandedModel();
    const table = model.sections[0].blocks[1] as Table;
    table.rows[2].isHeader = true; // a late header row, outside the run
    const reopened = await roundTrip(model);
    const reopenedTable = reopened.sections[0].blocks[1] as Table;
    expect(reopenedTable.rows[0].isHeader).toBe(true);
    expect(reopenedTable.rows[1].isHeader).toBeUndefined();
    // The late header row came back as an ordinary row — the loss AC-14
    // reports rather than hides.
    expect(reopenedTable.rows[2].isHeader).toBeUndefined();
  });
});

describe("collectOdfTableDropRecords (AC-14, AC-70)", () => {
  it("reports a header row outside the leading run, naming the table and the 1-based row", () => {
    const model = bandedModel();
    (model.sections[0].blocks[1] as Table).rows[2].isHeader = true;
    const dropped = collectOdfTableDropRecords(model);
    expect(dropped).toHaveLength(1);
    expect(dropped[0].kind).toBe("tableStyling");
    expect(dropped[0].location).toBe("table at paragraph 2, row 3");
    expect(dropped[0].reason).toContain("only a leading run of header rows");
  });

  it("reports nothing for a table whose header rows are all leading", () => {
    expect(collectOdfTableDropRecords(bandedModel())).toEqual([]);
  });

  it("reports a band size other than 1, naming the style and the size (AC-70)", async () => {
    const model = bandedModel();
    model.styles = [
      { ...BANDED_STYLE, banding: { ...BANDED_STYLE.banding, bandedRows: { size: 3, shading: "#d9e2f3" } } },
    ];
    const dropped = collectOdfTableDropRecords(model);
    expect(dropped).toHaveLength(1);
    expect(dropped[0].kind).toBe("tableStyling");
    expect(dropped[0].location).toContain("Accent");
    expect(dropped[0].location).toContain("3");

    // …and the .odt itself carries size-1 banding, not the declared size.
    const reopened = await roundTrip(model);
    const style = (reopened.styles ?? []).find((entry) => entry.id === "TableStyleAccent");
    expect(style?.banding?.bandedRows).toEqual({ size: 1, shading: "#d9e2f3" });
  });
});
