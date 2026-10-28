// SPEC-140 — a generated workbook opens in the in-browser editor with no
// download round trip, as openInEditor.spec.ts shows for documents: the
// Blob toBlob() produced, wrapped in a File (openAnyDocumentFile's
// signature, not a disk trip), goes through the same funnel HomeView's
// picker uses — openAnyDocumentFile, then storeOpenedFile — and lands at
// editorPathFor("spreadsheet", id), /sheets/:id. storeSpreadsheet(
// builder.toWorkbook()) is the model-level alternative. Runs under vitest's
// jsdom environment, which supplies Blob/File/FileReader as a browser does.
import { describe, expect, it } from "vitest";
import { editorPathFor } from "../documentKinds";
import { getSpreadsheetShell, storeOpenedFile, storeSpreadsheet } from "../documentStore";
import { openAnyDocumentFile } from "../openDocument";
import { spreadsheetMimeType } from "../spreadsheet/fileTypes";
import { WorkbookBuilder } from "./workbook";

function q3Sales(): WorkbookBuilder {
  return new WorkbookBuilder({ title: "Q3 sales" })
    .addSheet("Sales")
    .setRow("A1", ["Region", "July", "August", "September", "Total"])
    .setRows("A2", [["East", 1200, 1350, 1410]])
    .setFormula("E2", "SUM(B2:D2)")
    .numberFormat("B2:E2", "$#,##0.00")
    .freeze({ rows: 1 })
    .addSheet("Notes")
    .setCell("A1", "Generated");
}

describe("generated Blob → the in-browser editor (AC-38, AC-39)", () => {
  for (const format of ["xlsx", "ods"] as const) {
    it(`opens a .${format} with openAnyDocumentFile/storeOpenedFile at /sheets/:id, no download round trip`, async () => {
      const builder = q3Sales();
      const blob = await builder.toBlob(format);
      expect(blob.type).toBe(spreadsheetMimeType(format));
      const file = new File([blob], builder.fileName(format), { type: blob.type });

      const opened = await openAnyDocumentFile(file);
      expect(opened.kind).toBe("spreadsheet");
      storeOpenedFile(opened);

      const id = opened.model.id;
      expect(editorPathFor("spreadsheet", id)).toBe(`/sheets/${id}`);
      const stored = getSpreadsheetShell(id);
      expect(stored).toBeDefined();
      expect(stored!.title).toBe("Q3 sales");
      expect(stored!.format).toBe(format);
      expect(stored!.sheets.map((sheet) => sheet.name)).toEqual(["Sales", "Notes"]);
      expect(stored!.sheets[0]!.frozen).toEqual({ rows: 1, cols: 0 });
      const e2 = stored!.sheets[0]!.rows[1]!.cells[4]!;
      expect([e2.value, e2.formula]).toEqual([3960, "SUM(B2:D2)"]);
    });
  }

  it("storeSpreadsheet(builder.toWorkbook()) stores the model itself, the same route away", () => {
    const model = q3Sales().toWorkbook();
    storeSpreadsheet(model);
    const stored = getSpreadsheetShell(model.id);
    expect(stored?.sheets.map((sheet) => sheet.name)).toEqual(["Sales", "Notes"]);
    expect(stored?.preview?.sheetName).toBe("Sales");
    expect(editorPathFor("spreadsheet", model.id)).toBe(`/sheets/${model.id}`);
  });
});
