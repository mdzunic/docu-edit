// SPEC-124's structure fixtures (spec124-structure.xlsx / spec124-structure.ods,
// built by fixtures/build.ts): the committed e2e files are the builders'
// bytes, both packages are whole, and what SPEC124_STRUCTURE_XLSX and
// SPEC124_STRUCTURE_ODS record is what the files hold and what the reader
// gives — values, formulas (every cached value is what its formula
// computes, so a full recalculation changes nothing), merges, frozen panes
// and column widths — and that the structural constructs a rebuilt sheet
// must carry verbatim (the table, the chart, the drawing, the legacy
// comment and its VML note, the validations and the named ranges) are
// present byte-for-byte in their parts. An untouched download re-emits
// every part byte-identical.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { formatA1, getCell, parseA1, parseRange } from "./address";
import { spreadsheetDownload } from "./download";
import { buildSpec124StructureOdsBytes, buildSpec124StructureXlsxBytes, SPEC124_STRUCTURE_ODS, SPEC124_STRUCTURE_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { recalculate } from "./formula/recalc";
import type { Cell, CellValue, Sheet, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const decoder = new TextDecoder();

function sheetNamed(model: SpreadsheetModel, name: string): Sheet {
  const sheet = model.sheets.find((candidate) => candidate.name === name);
  expect(sheet, name).toBeDefined();
  return sheet!;
}

function cellAt(sheet: Sheet, a1: string): Cell | undefined {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col);
}

/** An A1 address as the model's plain { row, col } (parseA1 also reports
 * which parts were absolute, which a view-state field never carries). */
function at(a1: string): { row: number; col: number } {
  const { row, col } = parseA1(a1)!;
  return { row, col };
}

/** Every stored cell of a sheet as A1 → value. */
function storedValues(sheet: Sheet): Record<string, CellValue> {
  return Object.fromEntries(sheet.rows.flatMap((row) => row.cells.map((cell) => [formatA1({ row: row.index, col: cell.col }), cell.value])));
}

/** Every formula cell of the workbook. */
function formulaCells(model: SpreadsheetModel): Cell[] {
  return model.sheets.flatMap((sheet) => sheet.rows.flatMap((row) => row.cells.filter((cell) => cell.formula !== undefined)));
}

describe("spec124-structure.xlsx (Excel-shaped)", () => {
  const F = SPEC124_STRUCTURE_XLSX;
  const bytes = buildSpec124StructureXlsxBytes();
  const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, "xlsx", "spec124-structure");
  const part = async (name: string): Promise<string> => decoder.decode((await unzip(await bytes)).get(name));

  it("is the committed e2e fixture, byte for byte, and a whole package in Excel's part order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec124-structure.xlsx")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(F.parts);
  });

  it("reads Sheet1 (active) and Sheet2 at their own parts", async () => {
    const model = await read();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(F.sheetNames);
    expect(model.sheets.map((sheet) => sheet.visibility)).toEqual(["visible", "visible"]);
    expect(model.sheets.map((sheet) => sheet.partName)).toEqual(["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]);
    expect(model.activeSheetIndex).toBe(F.activeSheetIndex);
    expect(model.calc).toBeUndefined();
  });

  it("reads every data row's values and the shared-formula total", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    for (const [ref, value] of Object.entries(F.headers)) {
      expect(cellAt(sheet1, ref)?.value, ref).toBe(value);
    }
    for (const row of F.rows) {
      expect(storedValues(sheet1)[`A${row.row}`], `A${row.row}`).toBe(row.item);
      expect(storedValues(sheet1)[`B${row.row}`], `B${row.row}`).toBe(row.qty);
      expect(storedValues(sheet1)[`C${row.row}`], `C${row.row}`).toBe(row.price);
      expect(storedValues(sheet1)[`D${row.row}`], `D${row.row}`).toBe(row.code);
      expect(storedValues(sheet1)[`E${row.row}`], `E${row.row}`).toBe(row.description);
      expect(storedValues(sheet1)[`F${row.row}`], `F${row.row}`).toBe(row.total);
      // The shared formula's canonical text shifts with the row.
      expect(cellAt(sheet1, `F${row.row}`)?.formula, `F${row.row}`).toBe(`B${row.row}*C${row.row}`);
      expect(cellAt(sheet1, `F${row.row}`)?.formulaUnparsed, `F${row.row}`).toBeUndefined();
    }
    // The longest description is at row 12, as the facts record.
    expect(F.rows.find((row) => row.description === F.rows[10]!.description && row.row !== 12)).toBeUndefined();
  });

  it("holds H5's plain formula and the H20:H21 legacy array, with no formula on the member cell", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    expect(cellAt(sheet1, F.simpleFormula.ref)).toMatchObject({ formula: F.simpleFormula.formula, value: F.simpleFormula.value });
    const anchor = cellAt(sheet1, F.arrayFormula.anchor)!;
    expect(anchor).toMatchObject({ formula: F.arrayFormula.formula, value: F.arrayFormula.values.H20 });
    expect(anchor.arrayRange).toEqual(parseRange(F.arrayFormula.ref));
    expect(anchor.dynamicArray).toBeUndefined();
    expect(cellAt(sheet1, "H21")).toMatchObject({ value: F.arrayFormula.values.H21 });
    expect(cellAt(sheet1, "H21")?.formula).toBeUndefined();
  });

  it("reads the layout: the frozen first column, the two merges and column C's width", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    expect(sheet1.frozen).toEqual(F.frozen);
    expect(sheet1.topLeftCell).toEqual(at(F.topLeftCell));
    expect(sheet1.merges).toEqual(F.merges.map((merge) => parseRange(merge.ref)));
    for (const merge of F.merges) {
      expect(cellAt(sheet1, merge.ref.split(":")[0]!)?.value, merge.ref).toBe(merge.value);
    }
    expect(sheet1.columns).toHaveLength(1);
    expect(sheet1.columns[0]).toMatchObject({ min: F.cols[0].min, max: F.cols[0].max, sourceWidthChars: Number(F.cols[0].width) });
  });

  it("holds the Rates table's own values and the local P1:P3 validation list", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    expect(cellAt(sheet1, "J1")?.value).toBe(F.table.columns[0]);
    expect(cellAt(sheet1, "K1")?.value).toBe(F.table.columns[1]);
    F.table.regions.forEach((region, index) => {
      const r = index + 2;
      expect(cellAt(sheet1, `J${r}`)?.value, `J${r}`).toBe(region);
      expect(cellAt(sheet1, `K${r}`)?.value, `K${r}`).toBe(F.table.rates[index]);
    });
    F.validationList.values.forEach((value, index) => {
      expect(cellAt(sheet1, `P${index + 1}`)?.value, `P${index + 1}`).toBe(value);
    });
  });

  it("reads Sheet2's cross-sheet, 3-D and literal formulas", async () => {
    const model = await read();
    const sheet2 = sheetNamed(model, "Sheet2");
    for (const [a1, fact] of Object.entries(F.sheet2.formulas)) {
      expect(cellAt(sheet2, a1), a1).toMatchObject({ formula: fact.formula, value: fact.value });
      expect(cellAt(sheet2, a1)?.formulaUnparsed, a1).toBeUndefined();
    }
    expect(storedValues(sheet2)).toMatchObject(F.sheet2.values);
    expect(formulaCells(model)).toHaveLength(F.rows.length + 2 /* H5, H20 */ + Object.keys(F.sheet2.formulas).length);
  });

  it("caches what the formulas compute: a full manual recalculation changes no value", async () => {
    const model = await read();
    const rowsBefore = model.sheets.map((sheet) => sheet.rows);
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.stale).toEqual([]);
    expect(result.circular).toEqual([]);
    model.sheets.forEach((sheet, index) => expect(sheet.rows, sheet.name).toBe(rowsBefore[index]));
    expect(model.forceFullCalcOnLoad).toBeUndefined();
  });

  it("downloads untouched: every part byte-identical, in the recorded order", async () => {
    const source = await bytes;
    const model = await read();
    const download = await spreadsheetDownload(model);
    const out = await unzip(download.bytes);
    const original = await unzip(source);
    expect([...original.keys()]).toEqual(F.parts);
    expect([...out.keys()]).toEqual(F.parts);
    for (const [name, original_] of original) {
      expect(equalBytes(out.get(name)!, original_), name).toBe(true);
    }
  });

  it("writes Excel's worksheet shape: CT_Worksheet order, the frozen pane, the autoFilter and the mergeCells", async () => {
    const sheet1 = await part("xl/worksheets/sheet1.xml");
    expect(sheet1).toContain(`<pane xSplit="${F.frozen.cols}" topLeftCell="${F.topLeftCell}" activePane="topRight" state="frozen"/>`);
    expect(sheet1).toContain(`<selection pane="topRight" activeCell="B2" sqref="B2"/>`);
    expect(sheet1).toContain(`<autoFilter ref="${F.autoFilter.ref}"/>`);
    expect(sheet1).toContain(`<mergeCells count="2"><mergeCell ref="${F.merges[0]!.ref}"/><mergeCell ref="${F.merges[1]!.ref}"/></mergeCells>`);
    expect(sheet1).toMatch(/<dimension[^>]*\/><sheetViews>.*<\/sheetViews><sheetFormatPr[^>]*\/><cols>.*<\/cols><sheetData>/s);
    expect(sheet1).toMatch(/<\/sheetData><autoFilter[^>]*\/><mergeCells.*<\/mergeCells><conditionalFormatting.*<dataValidations.*<hyperlinks.*<pageMargins[^>]*\/><drawing[^>]*\/><legacyDrawing[^>]*\/><tableParts.*<\/tableParts><extLst>/s);
  });

  it("holds the conditional format, the classic and x14 validations, and the hyperlink to Sheet2", async () => {
    const sheet1 = await part("xl/worksheets/sheet1.xml");
    expect(sheet1).toContain(
      `<conditionalFormatting sqref="${F.conditionalFormatting.sqref}"><cfRule type="expression" dxfId="${F.conditionalFormatting.dxfId}" priority="1">` +
        `<formula>C2&gt;AVERAGE($C$2:$C$31)</formula></cfRule></conditionalFormatting>`,
    );
    expect(sheet1).toContain(
      `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${F.dataValidation.sqref}" ` +
        `xr:uid="${F.dataValidation.uid}"><formula1>${F.dataValidation.formula1}</formula1></dataValidation></dataValidations>`,
    );
    expect(sheet1).toContain(
      `<hyperlinks><hyperlink ref="${F.hyperlink.ref}" location="${F.hyperlink.location}" display="${F.hyperlink.display}" xr:uid="${F.hyperlink.uid}"/></hyperlinks>`,
    );
    expect(sheet1).toContain(`<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" xr:uid="${F.x14Validation.uid}">`);
    expect(sheet1).toContain(`<xm:f>${F.x14Validation.formula1}</xm:f>`);
    expect(sheet1).toContain(`<xm:sqref>${F.x14Validation.sqref}</xm:sqref>`);
    const styles = await part("xl/styles.xml");
    expect(styles).toContain(`<dxfs count="1">`);
  });

  it("declares the Rates table with its own ref and autoFilter, related from Sheet1", async () => {
    const table = await part(F.table.partName);
    expect(table).toContain(`id="${F.table.id}" xr:uid="${F.table.uid}" name="${F.table.name}" displayName="${F.table.displayName}" ref="${F.table.ref}"`);
    expect(table).toContain(`<autoFilter ref="${F.table.ref}"/>`);
    const sheetRels = await part("xl/worksheets/_rels/sheet1.xml.rels");
    expect(sheetRels).toContain(`Target="../tables/table1.xml"`);
    const sheet1 = await part("xl/worksheets/sheet1.xml");
    expect(sheet1).toContain(`<tableParts count="1"><tablePart r:id="rId3"/></tableParts>`);
  });

  it("names Sheet1 in the chart's series, categories and values, with caches", async () => {
    const chart = await part(F.chart.partName);
    expect([...chart.matchAll(/<c:f>([^<]+)<\/c:f>/g)].map((match) => match[1])).toEqual(F.chart.formulas);
    expect(chart).toContain("<c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>");
    expect(chart).toContain(`<c:ptCount val="${F.rows.length}"/>`);
    // The drawing anchors the chart and a picture, each its own twoCellAnchor.
    const drawing = await part(F.drawing.partName);
    expect([...drawing.matchAll(/<xdr:twoCellAnchor/g)]).toHaveLength(2);
    expect(drawing).toContain(`<xdr:col>${F.drawing.anchors.picture.from.col}</xdr:col>`);
    expect(drawing).toContain(`<xdr:col>${F.drawing.anchors.chart.from.col}</xdr:col>`);
    expect(drawing).toContain(`r:embed="${F.drawing.anchors.picture.embedRId}"`);
    expect(drawing).toContain(`<c:chart xmlns:c="${CHART_NS}" r:id="${F.drawing.anchors.chart.rId}"/>`);
    const drawingRels = await part(F.drawing.rels);
    expect(drawingRels).toContain(`Target="../charts/chart1.xml"`);
    expect(drawingRels).toContain(`Target="../media/image1.png"`);
  });

  it("holds the legacy comment and its VML note, anchored at B3", async () => {
    const comments = await part(F.comment.partName);
    expect(comments).toContain(`<comment ref="${F.comment.ref}" authorId="0" shapeId="0">`);
    expect(comments).toContain(`<t>${F.comment.text}</t>`);
    expect(comments).toContain(`<author>${F.comment.author}</author>`);
    const vml = await part(F.vml.partName);
    expect(vml).toContain(`<x:Anchor>\r\n    ${F.vml.anchor}</x:Anchor>`);
    expect(vml).toContain(`<x:Row>${F.vml.row}</x:Row>\r\n   <x:Column>${F.vml.column}</x:Column>`);
    expect(vml).toContain(`<o:idmap v:ext="edit" data="${F.vml.idmap}"/>`);
  });

  it("holds the workbook's definedNames, calcChain and app.xml heading pairs", async () => {
    const workbook = await part("xl/workbook.xml");
    for (const name of F.definedNames) {
      expect(workbook, name.name).toContain(`name="${name.name}"`);
    }
    expect(workbook).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$F$31</definedName>`);
    expect(workbook).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$F$31</definedName>`);
    expect(workbook).toContain(`<definedName name="Prices">Sheet1!$C$2:$C$31</definedName>`);
    expect(workbook).toContain(`<definedName name="TaxRate">Sheet2!$B$1</definedName>`);
    const calcChain = await part("xl/calcChain.xml");
    for (const row of F.rows) {
      expect(calcChain, `F${row.row}`).toContain(`r="F${row.row}"`);
    }
    expect(calcChain).toContain(`<c r="H20" a="1"/>`);
    const app = await part("docProps/app.xml");
    expect(app).toContain(`<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4>`);
    expect(app).toContain(`<vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4>`);
    for (const title of F.titlesOfParts) {
      expect(app, title).toContain(`<vt:lpstr>${title}</vt:lpstr>`);
    }
  });
});

describe("spec124-structure.ods (LibreOffice-shaped)", () => {
  const O = SPEC124_STRUCTURE_ODS;
  const bytes = buildSpec124StructureOdsBytes();
  const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, "ods", "spec124-structure");
  const part = async (name: string): Promise<string> => decoder.decode((await unzip(await bytes)).get(name));

  it("is the committed e2e fixture, byte for byte, and a whole package in Calc's entry order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec124-structure.ods")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(O.parts);
  });

  it("reads Sheet1 (active) and Sheet2: headers, the 10 data rows and the doubling formula", async () => {
    const model = await read();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(O.tableNames);
    expect(model.activeSheetIndex).toBe(O.activeSheetIndex);
    const sheet1 = sheetNamed(model, "Sheet1");
    for (const [ref, value] of Object.entries(O.headers)) {
      expect(cellAt(sheet1, ref)?.value, ref).toBe(value);
    }
    for (const row of O.rows) {
      expect(storedValues(sheet1)[`A${row.row}`], `A${row.row}`).toBe(row.item);
      expect(storedValues(sheet1)[`B${row.row}`], `B${row.row}`).toBe(row.amount);
      expect(storedValues(sheet1)[`C${row.row}`], `C${row.row}`).toBe(row.code);
      expect(cellAt(sheet1, `D${row.row}`), `D${row.row}`).toMatchObject({ value: row.value });
      expect(sheet1.sourceXml, `D${row.row}`).toContain(`table:formula="${row.formula.replaceAll("&", "&amp;")}"`);
    }
  });

  it("holds the merged note at A13:C13, with B13 and C13 covered", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    expect(sheet1.merges).toEqual([parseRange(O.merge.ref)]);
    expect(cellAt(sheet1, "A13")?.value).toBe(O.merge.value);
    for (const a1 of O.merge.covered) {
      expect(cellAt(sheet1, a1), a1).toBeUndefined();
    }
  });

  it("freezes Sheet1's first column (settings.xml) and leaves Sheet2 unsplit", async () => {
    const model = await read();
    expect(sheetNamed(model, "Sheet1").frozen).toEqual(O.frozen);
    expect(sheetNamed(model, "Sheet2").frozen).toBeUndefined();
  });

  it("reads Sheet2's cross-sheet formulas", async () => {
    const model = await read();
    const sheet2 = sheetNamed(model, "Sheet2");
    for (const [a1, fact] of Object.entries(O.sheet2.formulas)) {
      expect(cellAt(sheet2, a1), a1).toMatchObject({ formula: fact.formula, value: fact.value });
      expect(sheet2.sourceXml, a1).toContain(`table:formula="${fact.wire}"`);
    }
  });

  it("caches what the formulas compute: a full manual recalculation changes no value", async () => {
    const model = await read();
    const rowsBefore = model.sheets.map((sheet) => sheet.rows);
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.stale).toEqual([]);
    expect(result.circular).toEqual([]);
    model.sheets.forEach((sheet, index) => expect(sheet.rows, sheet.name).toBe(rowsBefore[index]));
  });

  it("downloads untouched: every part byte-identical, mimetype first and STOREd; the envelope splices back", async () => {
    const source = await bytes;
    const model = await read();
    const download = await spreadsheetDownload(model);
    const out = await unzip(download.bytes);
    const original = await unzip(source);
    expect([...original.keys()]).toEqual(O.parts);
    expect([...out.keys()]).toEqual(O.parts);
    for (const [name, original_] of original) {
      expect(equalBytes(out.get(name)!, original_), name).toBe(true);
    }
    const envelope = model.odsContentEnvelope!;
    expect(envelope.prefix + model.sheets.map((sheet) => sheet.sourceXml).join("") + envelope.suffix).toBe(await part("content.xml"));
  });

  it("puts the content validation in the prefix, and the named range, named expression and database range in the suffix", async () => {
    const model = await read();
    const { prefix, suffix } = model.odsContentEnvelope!;
    expect(prefix).toContain(
      `<table:content-validations><table:content-validation table:name="${O.contentValidation.name}" table:condition="${O.contentValidation.condition.replaceAll("&", "&amp;")}" ` +
        `table:allow-empty-cell="true" table:base-cell-address="${O.contentValidation.baseCellAddress}" table:display-list="unsorted">`,
    );
    expect(suffix).toContain(`table:name="${O.namedRange.name}"`);
    expect(suffix).toContain(`table:cell-range-address="${O.namedRange.cellRangeAddress}"`);
    expect(suffix).toContain(`table:name="${O.namedExpression.name}"`);
    expect(suffix).toContain(`table:expression="${O.namedExpression.expression.replaceAll("&", "&amp;")}"`);
    expect(suffix).toContain(`<table:database-ranges><table:database-range table:name="${O.databaseRange.name}" table:target-range-address="${O.databaseRange.targetRangeAddress}"`);
    const sheet1 = sheetNamed(model, "Sheet1");
    for (const a1 of O.contentValidation.cells) {
      expect(sheet1.sourceXml, a1).toContain(`table:content-validation-name="${O.contentValidation.name}"`);
    }
  });

  it("keeps the table-local conditional format after the rows, and the print range on the table", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    expect(sheet1.sourceXml).toContain(`table:print-ranges="${O.printRanges}"`);
    expect(sheet1.sourceXml).toContain(
      `<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="${O.conditionalFormat.targetRangeAddress}">` +
        `<calcext:condition calcext:apply-style-name="${O.conditionalFormat.applyStyleName}" calcext:value="${O.conditionalFormat.value.replaceAll(">", "&gt;")}" ` +
        `calcext:base-cell-address="${O.conditionalFormat.baseCellAddress}"/>`,
    );
    expect(sheet1.sourceXml!.indexOf("<calcext:conditional-formats>")).toBeGreaterThan(sheet1.sourceXml!.lastIndexOf("</table:table-row>"));
    const styles = await part("styles.xml");
    expect(styles).toContain(`style:name="${O.conditionalFormat.applyStyleName}"`);
  });

  it("keeps E2's chart frame directly in the cell, before any text:p, pointing at the embedded chart object", async () => {
    const model = await read();
    const sheet1 = sheetNamed(model, "Sheet1");
    expect(sheet1.sourceXml).toMatch(/<table:table-cell><draw:frame [^>]*table:end-cell-address=/);
    // SPEC-131: the frame is the sheet's drawing anchored at E2, no longer the cell's preservedXml.
    expect(cellAt(sheet1, O.frame.ref)).toBeUndefined();
    const e2 = parseA1(O.frame.ref)!;
    const frame = sheet1.drawings!.find((drawing) => drawing.content.kind === "chart")!;
    expect(frame.anchor.from).toMatchObject({ row: e2.row, col: e2.col });
    expect(frame.content).toEqual({ kind: "chart", sourcePart: "Object 1/content.xml" });
    expect(frame.sourceXml).toContain(`table:end-cell-address="${O.frame.endCellAddress}"`);
    expect(frame.sourceXml).toContain(`draw:notify-on-update-of-ranges="${O.frame.notifyRanges}"`);
    expect(frame.sourceXml).toContain(`xlink:href="${O.frame.href}"`);
    expect(frame.sourceXml).toContain(`xlink:href="${O.frame.replacementHref}"`);
    const chart = await part("Object 1/content.xml");
    expect(chart).toContain(`table:cell-range-address="${O.chart.plotAreaRange}"`);
    expect(chart).toContain(`<chart:categories table:cell-range-address="${O.chart.categoriesRange}"/>`);
    expect(chart).toContain(`chart:values-cell-range-address="${O.chart.valuesRange}" chart:label-cell-address="${O.chart.labelCell}"`);
  });

  it("is a whole package: the manifest lists the embedded object and its replacement", async () => {
    const manifest = await part("META-INF/manifest.xml");
    expect(manifest).toContain(`manifest:full-path="Object 1/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.chart"`);
    expect(manifest).toContain(`manifest:full-path="${O.chart.replacement}"`);
    for (const objectPart of O.chart.objectParts) {
      expect(manifest, objectPart).toContain(`manifest:full-path="${objectPart}"`);
    }
  });
});
