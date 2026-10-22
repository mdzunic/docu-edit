// SPEC-124 round trips on spec124-structure.xlsx and .ods: the e2e's edits —
// rows inserted, a column deleted, a resize, a hide, a merge, a freeze —
// written and re-read reproduce the adjusted formulas, merges, frozen panes,
// sizes and hidden flags (field-level toEqual on the re-read sheets, minted
// ids aside); every retained reference in the tables sits at its adjusted
// position; every part the edits did not touch keeps its bytes; and an
// untouched workbook writes byte-identical parts.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { getCell } from "./address";
import { freezePanes, resizeColumns, setRowsHidden } from "./bands";
import { buildSpec124StructureOdsBytes, buildSpec124StructureXlsxBytes } from "./fixtures/build";
import { mergeCells } from "./mergeCells";
import type { Sheet, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { applyStructuralEdit } from "./structure";
import { writeOds } from "./ods/writer";
import { writeXlsx } from "./xlsx/writer";

const text = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes);

/** What a re-read sheet must reproduce, its minted id and source aside. */
function shape(sheet: Sheet): unknown {
  return {
    name: sheet.name,
    formulas: sheet.rows.flatMap((row) => row.cells.filter((cell) => cell.formula !== undefined).map((cell) => [row.index, cell.col, cell.formula])),
    values: sheet.rows.flatMap((row) => row.cells.filter((cell) => cell.value !== null).map((cell) => [row.index, cell.col, cell.value])),
    merges: sheet.merges,
    frozen: sheet.frozen,
    widths: sheet.columns.filter((span) => span.widthPx !== undefined || span.hidden).map((span) => [span.min, span.max, span.widthPx, span.hidden]),
    hidden: sheet.rows.filter((row) => row.hidden).map((row) => row.index),
  };
}

/** The e2e flow on Sheet1, as lib calls. */
function edit(model: SpreadsheetModel): void {
  const id = model.sheets[0]!.id;
  expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: id, at: 2, count: 2 })).toEqual({ ok: true });
  expect(applyStructuralEdit(model, { kind: "deleteCols", sheetId: id, at: 1, count: 1 })).toEqual({ ok: true });
  expect(resizeColumns(model, id, [{ start: 2, end: 2 }], 120)).toEqual({ ok: true });
  expect(setRowsHidden(model, id, [{ start: 9, end: 9 }], true)).toEqual({ ok: true });
  expect(mergeCells(model, id, [{ top: 0, left: 0, bottom: 0, right: 2 }], "cells")).toEqual({ ok: true });
  expect(freezePanes(model, id, 1, 0)).toEqual({ ok: true });
}

describe("SPEC-124 round trip — spec124-structure.xlsx", () => {
  it("an untouched workbook writes every part byte-identical", async () => {
    const bytes = await buildSpec124StructureXlsxBytes();
    const source = await unzip(bytes);
    const out = await unzip(await writeXlsx(await readSpreadsheetBytes(bytes, "xlsx", "t")));
    expect([...out.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("reproduces the edits, moves every retained reference, and keeps the untouched parts' bytes", async () => {
    const bytes = await buildSpec124StructureXlsxBytes();
    const source = await unzip(bytes);
    const model = await readSpreadsheetBytes(bytes, "xlsx", "t");
    edit(model);
    const written = await writeXlsx(model);
    const out = await unzip(written);
    const reread = await readSpreadsheetBytes(written, "xlsx", "t2");
    expect(reread.sheets.map(shape)).toEqual(model.sheets.map(shape));
    const [sheet1, sheet2] = reread.sheets;
    expect(getCell(sheet2!, 0, 0)!.formula).toBe("Sheet1!A7");
    expect(getCell(sheet2!, 1, 0)!.formula).toBe("SUM('Sheet1'!#REF!)");
    expect(getCell(sheet2!, 3, 0)!.formula).toBe("Sheet1!#REF!*2");
    expect(getCell(sheet1!, 6, 6)!.formula).toBe("#REF!*2");
    expect(sheet1!.merges).toContainEqual({ top: 0, left: 0, bottom: 0, right: 2 });
    expect(sheet1!.frozen).toEqual({ rows: 1, cols: 0 });

    const sheet1Xml = text(out.get("xl/worksheets/sheet1.xml"));
    expect(sheet1Xml).toContain(`<conditionalFormatting sqref="B2:B33">`);
    expect(sheet1Xml).toContain(`<formula>B2&gt;AVERAGE($B$2:$B$33)</formula>`);
    expect(sheet1Xml).toContain(`sqref="C2:C33"`);
    // P1:P3 holds row 3, where the rows went in: it grows, then moves left.
    expect(sheet1Xml).toContain(`<formula1>$O$1:$O$5</formula1>`);
    expect(sheet1Xml).toContain(`<autoFilter ref="A1:E33"/>`);
    expect(sheet1Xml).toContain(`<hyperlink ref="A33" location="Sheet2!A1"`);
    expect(sheet1Xml).toContain(`<xm:sqref>A2:A33</xm:sqref>`);
    const workbook = text(out.get("xl/workbook.xml"));
    expect(workbook).toContain(`<definedName name="Prices">Sheet1!$B$2:$B$33</definedName>`);
    expect(workbook).toContain(`>Sheet1!$A$1:$E$33</definedName>`);
    expect(text(out.get("xl/tables/table1.xml"))).toContain(`ref="I1:J8"`);
    expect(text(out.get("xl/charts/chart1.xml"))).toContain(`<c:f>Sheet1!$B$2:$B$33</c:f>`);
    expect(text(out.get("xl/charts/chart1.xml"))).toContain(`<c:f>Sheet1!$B$1</c:f>`);
    // B3's comment moved to B5, whose column was then deleted: the comment
    // and its VML shape are gone.
    expect(text(out.get("xl/comments1.xml"))).not.toContain(`<comment ref=`);
    expect(text(out.get("xl/drawings/vmlDrawing1.vml"))).not.toContain(`<x:ClientData`);
    expect(out.has("xl/calcChain.xml")).toBe(false);
    for (const name of ["xl/theme/theme1.xml", "docProps/core.xml", "xl/media/image1.png"]) {
      expect(equalBytes(out.get(name)!, source.get(name)!), name).toBe(true);
    }
  });
});

describe("SPEC-124 round trip — spec124-structure.ods", () => {
  it("an untouched workbook writes every part byte-identical", async () => {
    const bytes = await buildSpec124StructureOdsBytes();
    const source = await unzip(bytes);
    const out = await unzip(await writeOds(await readSpreadsheetBytes(bytes, "ods", "t")));
    for (const [name, data] of source) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("inserting a column moves the named range, database range, conditional format, print ranges, validation base, frame end and chart ranges", async () => {
    const bytes = await buildSpec124StructureOdsBytes();
    const source = await unzip(bytes);
    const model = await readSpreadsheetBytes(bytes, "ods", "t");
    const id = model.sheets[0]!.id;
    expect(applyStructuralEdit(model, { kind: "insertCols", sheetId: id, at: 1, count: 1 })).toEqual({ ok: true });
    const written = await writeOds(model);
    const out = await unzip(written);
    const reread = await readSpreadsheetBytes(written, "ods", "t2");
    expect(reread.sheets.map(shape)).toEqual(model.sheets.map(shape));
    expect(getCell(reread.sheets[1]!, 0, 0)!.formula).toBe("Sheet1!C5");
    const content = text(out.get("content.xml"));
    expect(content).toContain(`table:cell-range-address="$Sheet1.$C$2:.$C$11"`);
    expect(content).toContain(`table:expression="of:=[$Sheet1.$C$2]*2"`);
    expect(content).toContain(`table:target-range-address="Sheet1.A1:Sheet1.D11"`);
    expect(content).toContain(`calcext:target-range-address="Sheet1.C2:Sheet1.C11"`);
    // SPEC-136: the print range is the model's _xlnm.Print_Area, moved by the names adjuster and written through canonicalToOdsRangeAddress.
    expect(content).toContain(`table:print-ranges="$Sheet1.$A$1:.$D$11"`);
    expect(content).toContain(`table:base-cell-address="Sheet1.D2"`);
    expect(content).toContain(`table:end-cell-address="Sheet1.H8"`);
    // SPEC-132: the chart is modelled — the charts adjuster moved its references and the
    // ods writer wrote the moved addresses through canonicalToOdsRangeAddress (the frame's
    // notify list and the object's patched addresses); the unmoved categories keep LibreOffice's spelling.
    expect(content).toContain(`draw:notify-on-update-of-ranges="$Sheet1.A1:.C11"`);
    const chart = text(out.get("Object 1/content.xml"));
    expect(chart).toContain(`chart:values-cell-range-address="$Sheet1.C2:.C11"`);
    expect(chart).toContain(`chart:label-cell-address="$Sheet1.C1"`);
    expect(chart).toContain(`table:cell-range-address="Sheet1.A2:Sheet1.A11"`);
    expect(equalBytes(out.get("styles.xml")!, source.get("styles.xml")!)).toBe(true);
    expect(equalBytes(out.get("settings.xml")!, source.get("settings.xml")!)).toBe(true);
  });

  it("a freeze, a resize, a hide and a merge on the ods reproduce on re-read", async () => {
    const model = await readSpreadsheetBytes(await buildSpec124StructureOdsBytes(), "ods", "t");
    const id = model.sheets[0]!.id;
    expect(model.sheets[0]!.frozen).toEqual({ rows: 0, cols: 1 });
    expect(freezePanes(model, id, 1, 0)).toEqual({ ok: true });
    expect(resizeColumns(model, id, [{ start: 2, end: 2 }], 150)).toEqual({ ok: true });
    expect(setRowsHidden(model, id, [{ start: 4, end: 4 }], true)).toEqual({ ok: true });
    expect(mergeCells(model, id, [{ top: 14, left: 0, bottom: 15, right: 1 }], "cells")).toEqual({ ok: true });
    const written = await writeOds(model);
    const reread = await readSpreadsheetBytes(written, "ods", "t2");
    expect(reread.sheets[0]!.frozen).toEqual({ rows: 1, cols: 0 });
    expect(reread.sheets.map(shape)).toEqual(model.sheets.map(shape));
  });
});
