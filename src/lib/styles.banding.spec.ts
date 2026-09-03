// SPEC-076 AC-37 — the style editor's banding half: the two band sizes and
// the six regions' shading, bold and borders, from the dialog's draft through
// StyleFormInput into the catalogue, and the re-render a save triggers.
import { describe, expect, it } from "vitest";
import { bandingDraftFrom, bandingFromDraft } from "./tableStyles";
import { styleFormInputFromDraft } from "./styles";
import { createStyle, modifyStyle } from "./commands.customStyles";
import { caretSelection } from "./commands";
import { createDocument, type DocumentModel, type Table, type TableStyleBanding } from "./documentModel";

const BANDING: TableStyleBanding = {
  firstRow: { shading: "#2f5496", bold: true },
  lastRow: { shading: "#dddddd" },
  firstColumn: { bold: true },
  bandedRows: { size: 2, shading: "#f2f2f2" },
  bandedColumns: { size: 3, shading: "#eaeaea" },
};

function draftFor(type = "table") {
  return {
    name: "Accent",
    type,
    basedOn: "",
    next: "",
    fontFamily: "",
    fontSizePt: undefined,
    bold: false,
    italic: false,
    underline: false,
    color: "",
    alignment: "",
    listPresetKey: "decimal",
    cellShadingFill: "",
    cellBorderStyle: "",
    cellBorderWidthPt: 1.5,
    cellBorderColor: "#123456",
    banding: bandingDraftFrom(BANDING),
  };
}

describe("bandingDraftFrom / bandingFromDraft (AC-37)", () => {
  it("lists the six regions in the dialog's order", () => {
    expect(bandingDraftFrom(undefined).regions.map((region) => region.key)).toEqual([
      "firstRow",
      "lastRow",
      "firstColumn",
      "lastColumn",
      "bandedRows",
      "bandedColumns",
    ]);
  });

  it("seeds the two band sizes from the style, defaulting to 1", () => {
    expect(bandingDraftFrom(BANDING).rowBandSize).toBe(2);
    expect(bandingDraftFrom(BANDING).columnBandSize).toBe(3);
    expect(bandingDraftFrom(undefined).rowBandSize).toBe(1);
    expect(bandingDraftFrom({ firstRow: { bold: true } }).columnBandSize).toBe(1);
  });

  it("round-trips shading, bold and band sizes", () => {
    const back = bandingFromDraft(bandingDraftFrom(BANDING), 1, "#000000");
    expect(back?.firstRow).toEqual({ shading: "#2f5496", bold: true });
    expect(back?.lastRow).toEqual({ shading: "#dddddd" });
    expect(back?.firstColumn).toEqual({ bold: true });
    expect(back?.bandedRows).toEqual({ size: 2, shading: "#f2f2f2" });
    expect(back?.bandedColumns).toEqual({ size: 3, shading: "#eaeaea" });
    // A region nothing set stays absent (AC-22's "independently absent").
    expect(back?.lastColumn).toBeUndefined();
  });

  it("builds a region's borders from its style plus the style's width and colour", () => {
    const draft = bandingDraftFrom(undefined);
    draft.regions[0].borderStyle = "double";
    const back = bandingFromDraft(draft, 2.25, "#ff0000");
    expect(back?.firstRow?.borders?.top).toEqual({ style: "double", widthPt: 2.25, color: "#ff0000" });
    expect(back?.firstRow?.borders?.right).toEqual({ style: "double", widthPt: 2.25, color: "#ff0000" });
  });

  it("round-trips a band region's bold and borders (AC-37)", () => {
    const draft = bandingDraftFrom(undefined);
    const bandedRows = draft.regions.find((region) => region.key === "bandedRows");
    const bandedColumns = draft.regions.find((region) => region.key === "bandedColumns");
    bandedRows!.bold = true;
    bandedRows!.borderStyle = "dashed";
    bandedColumns!.bold = true;
    const back = bandingFromDraft(draft, 1.5, "#00ff00");
    expect(back?.bandedRows?.size).toBe(1);
    expect(back?.bandedRows?.bold).toBe(true);
    expect(back?.bandedRows?.borders?.top).toEqual({ style: "dashed", widthPt: 1.5, color: "#00ff00" });
    // A band with bold alone and no shading is still a band (AC-22's
    // "independently settable").
    expect(back?.bandedColumns).toEqual({ size: 1, bold: true });
    // And the edit survives a second pass through the dialog.
    const again = bandingDraftFrom(back);
    expect(again.regions.find((region) => region.key === "bandedRows")).toMatchObject({
      bold: true,
      borderStyle: "dashed",
    });
    expect(again.regions.find((region) => region.key === "bandedColumns")).toMatchObject({ bold: true });
  });

  it("returns undefined when no region sets anything", () => {
    expect(bandingFromDraft(bandingDraftFrom(undefined), 1, "#000000")).toBeUndefined();
  });

  it("ignores a band whose shading is blank — a size alone is not banding", () => {
    const draft = bandingDraftFrom(undefined);
    draft.rowBandSize = 4;
    expect(bandingFromDraft(draft, 1, "#000000")).toBeUndefined();
  });
});

describe("styleFormInputFromDraft (AC-37)", () => {
  it("carries the banding for a table style", () => {
    const input = styleFormInputFromDraft(draftFor());
    expect(input.banding?.firstRow).toEqual({ shading: "#2f5496", bold: true });
    expect(input.banding?.bandedRows).toEqual({ size: 2, shading: "#f2f2f2" });
  });

  it("carries none for a paragraph style", () => {
    expect(styleFormInputFromDraft(draftFor("paragraph")).banding).toBeUndefined();
  });
});

function modelWithTable(styleId?: string): DocumentModel {
  const cell = (text: string) => ({
    blocks: [{ type: "paragraph" as const, runs: [{ text, formatting: {} }] }],
  });
  const table: Table = {
    type: "table",
    columnWidthsCm: [4, 4],
    rows: [
      { isHeader: true, cells: [cell("Name"), cell("Qty")] },
      { cells: [cell("a"), cell("1")] },
    ],
  };
  if (styleId) {
    table.styleId = styleId;
  }
  const model = createDocument("banding-style-fixture");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [] }, table] }];
  return model;
}

describe("createStyle / modifyStyle carry banding (AC-37)", () => {
  const inCell = caretSelection({ section: 0, block: 1, offset: 0, cell: { row: 0, col: 0, block: 0 } });

  it("a new table style keeps the dialog's banding", () => {
    const next = createStyle(modelWithTable(), inCell, styleFormInputFromDraft(draftFor())).model;
    const created = next.styles?.find((style) => style.name === "Accent");
    expect(created?.banding?.firstRow).toEqual({ shading: "#2f5496", bold: true });
  });

  it("modifying a table style replaces its banding, and every table using it re-renders", () => {
    const created = createStyle(modelWithTable(), inCell, styleFormInputFromDraft(draftFor())).model;
    const styleId = created.styles!.find((style) => style.name === "Accent")!.id;
    // createStyle applies the new style to the caret's table.
    expect((created.sections[0].blocks[1] as Table).styleId).toBe(styleId);

    const edited = draftFor();
    edited.banding.regions[0].shading = "#00ff00";
    const next = modifyStyle(created, inCell, styleId, styleFormInputFromDraft(edited)).model;
    expect(next.styles?.find((style) => style.id === styleId)?.banding?.firstRow?.shading).toBe("#00ff00");
    // The table block itself was replaced by the re-bake pass, so the canvas
    // re-renders it (banding is resolved at paint time from the catalogue).
    expect(next.sections[0].blocks[1]).not.toBe(created.sections[0].blocks[1]);
  });
});
