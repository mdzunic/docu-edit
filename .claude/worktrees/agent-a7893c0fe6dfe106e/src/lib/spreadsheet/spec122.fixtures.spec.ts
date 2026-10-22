// SPEC-122's fixtures (spec122-shapes.xlsx / spec122-shapes.ods, built by
// fixtures/build.ts — AC-58): the committed e2e files are the builders'
// bytes, both packages are whole, and what SPEC122_SHAPES_XLSX and
// SPEC122_SHAPES_ODS record is what the readers give — values, formulas,
// arrays, the data table, merges, the freeze, number formats and the SPEC-122
// captures (sstIndex, preservedAttrs/Xml/Text, sheet protection, the
// unlocked cell's style). Every cached value is what the engine computes, and
// an edit on "Edit" recalculates no other sheet — the ground the e2e stands on.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { formatA1, getCell, parseA1, parseRange } from "./address";
import { buildSpec122ShapesOdsBytes, buildSpec122ShapesXlsxBytes, SPEC122_SHAPES_ODS, SPEC122_SHAPES_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { cellKey } from "./formula/dependencies";
import { recalculate } from "./formula/recalc";
import type { Cell, CellValue, Sheet, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { cellStyle, effectiveStyleId } from "./styles";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();

function sheetNamed(model: SpreadsheetModel, name: string): Sheet {
  const sheet = model.sheets.find((candidate) => candidate.name === name);
  expect(sheet, name).toBeDefined();
  return sheet!;
}

/** An A1 address as the model's plain { row, col }. */
function at(a1: string): { row: number; col: number } {
  const { row, col } = parseA1(a1)!;
  return { row, col };
}

function cellAt(sheet: Sheet, a1: string): Cell | undefined {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col);
}

/** Every stored cell of a sheet as A1 → value. */
function storedValues(sheet: Sheet): Record<string, CellValue> {
  return Object.fromEntries(sheet.rows.flatMap((row) => row.cells.map((cell) => [formatA1({ row: row.index, col: cell.col }), cell.value])));
}

/** Every stored cell carrying `field`, as A1 → its value. */
function cellsWith(sheet: Sheet, field: "preservedAttrs" | "preservedXml" | "preservedText"): Record<string, string> {
  return Object.fromEntries(
    sheet.rows.flatMap((row) => row.cells.flatMap((cell) => (cell[field] === undefined ? [] : [[formatA1({ row: row.index, col: cell.col }), cell[field]!]]))),
  );
}

function styleAt(model: SpreadsheetModel, sheet: Sheet, a1: string) {
  const ref = parseA1(a1)!;
  return cellStyle(model, effectiveStyleId(sheet, ref.row, ref.col));
}

/** Every formula text in the workbook, with its sheet. */
function formulaTexts(model: SpreadsheetModel): { sheet: string; formula: string }[] {
  return model.sheets.flatMap((sheet) => sheet.rows.flatMap((row) => row.cells.flatMap((cell) => (cell.formula === undefined ? [] : [{ sheet: sheet.name, formula: cell.formula }]))));
}

const VOLATILE = /\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\(/i;

describe("spec122-shapes.xlsx (Excel-shaped)", () => {
  const F = SPEC122_SHAPES_XLSX;
  const bytes = buildSpec122ShapesXlsxBytes();
  const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, "xlsx", "spec122-shapes");

  it("is the committed e2e fixture, byte for byte, and a whole package in Excel's part order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec122-shapes.xlsx")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(F.parts);
  });

  it("reads Edit (first, active), Locked and Inline at their parts; only Locked is protected; the structure is not", async () => {
    const model = await read();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(F.sheetNames);
    expect(model.sheets.map((sheet) => sheet.partName)).toEqual(F.sheets.map((sheet) => sheet.partName));
    expect(model.activeSheetIndex).toBe(F.activeSheetIndex);
    expect(model.sheets.map((sheet) => sheet.protection)).toEqual([undefined, { sheet: true }, undefined]);
    expect(model.structureProtected).toBeUndefined();
    const parts = await unzip(await bytes);
    expect(decoder.decode(parts.get("xl/workbook.xml"))).not.toContain("workbookProtection");
    expect(decoder.decode(parts.get(F.sheets[1].partName))).toContain(`</sheetData>${F.locked.sheetProtection}<pageMargins `);
  });

  it("reads every stored value of every sheet", async () => {
    const model = await read();
    expect(storedValues(sheetNamed(model, "Edit"))).toEqual(F.edit.values);
    expect(storedValues(sheetNamed(model, "Locked"))).toEqual(F.locked.values);
    expect(storedValues(sheetNamed(model, "Inline"))).toEqual(F.inline.values);
    const inline = sheetNamed(model, "Inline");
    expect(cellAt(inline, F.inline.rich.ref)?.runs).toEqual(F.inline.rich.runs);
  });

  it("gives the e2e its cell: Edit!B2 a plain General number, outside every array, block and merge, read by C2", async () => {
    const model = await read();
    const edit = sheetNamed(model, "Edit");
    const target = cellAt(edit, F.edit.target.ref)!;
    expect(target).toEqual({ col: 1, value: F.edit.target.value });
    expect(styleAt(model, edit, F.edit.target.ref).numberFormat).toBeUndefined();
    expect(edit.merges.some((range) => range.top <= 1 && 1 <= range.bottom && range.left <= 1 && 1 <= range.right)).toBe(false);
    expect(cellAt(edit, F.edit.target.dependent.ref)?.formula).toBe(F.edit.target.dependent.formula);
    expect(edit.selection?.active).toEqual(at(F.edit.activeCell));
    const locked = sheetNamed(model, "Locked");
    expect(locked.selection?.active).toEqual(at(F.locked.activeCell));
  });

  it("reads Edit's formulas: shared blocks expanded, the legacy and dynamic arrays, the data table verbatim", async () => {
    const model = await read();
    const edit = sheetNamed(model, "Edit");
    for (const [a1, formula] of Object.entries(F.edit.formulas)) {
      expect(cellAt(edit, a1)?.formula, a1).toBe(formula);
      expect(cellAt(edit, a1)?.formulaUnparsed, a1).toBeUndefined();
    }
    expect(formulaTexts(model).filter((entry) => entry.sheet === "Edit")).toHaveLength(Object.keys(F.edit.formulas).length);
    expect(cellAt(edit, F.edit.array.anchor)?.arrayRange).toEqual(parseRange(F.edit.array.ref));
    expect(cellAt(edit, F.edit.array.anchor)?.dynamicArray).toBeUndefined();
    const dynamic = cellAt(edit, F.edit.dynamicArray.anchor)!;
    expect([dynamic.arrayRange, dynamic.dynamicArray, dynamic.cm]).toEqual([parseRange(F.edit.dynamicArray.ref), true, F.edit.dynamicArray.cm]);
    expect(cellAt(edit, F.edit.dataTable.anchor)?.dataTableXml).toBe(F.edit.dataTable.xml);
    expect(cellAt(edit, F.edit.dataTable.anchor)?.formula).toBeUndefined();
    expect(sheetNamed(model, "Locked").rows.flatMap((row) => row.cells).filter((cell) => cell.formula !== undefined).map((cell) => cell.formula)).toEqual(
      Object.values(F.locked.formulas),
    );
  });

  it("reads Edit's layout: the merge, the frozen row, the outlined rows and columns, the tall row", async () => {
    const edit = sheetNamed(await read(), "Edit");
    expect(edit.merges).toEqual([parseRange(F.edit.merge)]);
    expect(edit.frozen).toEqual(F.edit.frozen);
    expect(edit.topLeftCell).toEqual(at(F.edit.topLeftCell));
    for (const row of edit.rows) {
      const outline = (F.edit.outlineRows as readonly number[]).includes(row.index) ? ` outlineLevel="1"` : "";
      expect(row.sourceAttrs, `row ${row.index + 1}`).toBe(`spans="${F.edit.spans}"${outline} x14ac:dyDescent="0.35"`);
    }
    expect(edit.rows.find((row) => row.index === F.edit.tallRow.index)).toMatchObject({ heightPt: F.edit.tallRow.heightPt });
    expect(edit.columns.map((span) => ({ min: span.min, max: span.max, sourceWidthChars: span.sourceWidthChars, sourceAttrs: span.sourceAttrs }))).toEqual(
      F.edit.columns.map((span) => ({ sourceAttrs: undefined, ...span })),
    );
  });

  it("reads the number formats: built-ins 14 and 22, custom 164 and 165", async () => {
    const model = await read();
    const edit = sheetNamed(model, "Edit");
    const N = F.edit.numberFormats;
    expect(styleAt(model, edit, "B10").builtinFormatId).toBe(N.B10.builtinFormatId);
    expect(styleAt(model, edit, "B11").builtinFormatId).toBe(N.B11.builtinFormatId);
    expect(styleAt(model, edit, "B12").numberFormat).toBe(N.B12.code);
    expect(styleAt(model, edit, "E2").numberFormat).toBe(N.E2.code);
    expect(styleAt(model, edit, F.edit.quotePrefix.ref).quotePrefix).toBe(true);
  });

  it("captures what SPEC-122's writer re-emits: every shared-string index, the rich and phonetic items, ph and vm", async () => {
    const model = await read();
    const parts = await unzip(await bytes);
    for (const sheet of model.sheets) {
      const xml = decoder.decode(parts.get(sheet.partName!));
      const shared = Object.fromEntries([...xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*? t="s"[^>]*><v>(\d+)<\/v>/g)].map((match) => [match[1]!, Number(match[2])]));
      const captured = Object.fromEntries(
        sheet.rows.flatMap((row) => row.cells.flatMap((cell) => (cell.sstIndex === undefined ? [] : [[formatA1({ row: row.index, col: cell.col }), cell.sstIndex]]))),
      );
      expect(captured, sheet.name).toEqual(shared);
    }
    const edit = sheetNamed(model, "Edit");
    expect(cellAt(edit, F.edit.rich.ref)?.runs?.map((run) => run.text)).toEqual(F.edit.rich.runs);
    expect(cellAt(edit, F.edit.rich.ref)?.sstIndex).toBe(F.edit.rich.sstIndex);
    expect(cellAt(edit, F.edit.phonetic.ref)).toMatchObject({ value: F.edit.phonetic.value, sstIndex: F.edit.phonetic.sstIndex });
    expect(cellAt(edit, F.edit.carriageReturn.ref)?.sstIndex).toBe(F.edit.carriageReturn.sstIndex);
    expect(cellsWith(edit, "preservedAttrs")).toEqual({ [F.edit.phonetic.ref]: F.edit.phonetic.preservedAttrs, [F.edit.valueMetadata.ref]: F.edit.valueMetadata.preservedAttrs });
    expect(model.sheets.map((sheet) => cellsWith(sheet, "preservedXml"))).toEqual([{}, {}, {}]);
    expect(decoder.decode(parts.get(F.sheets[2].partName))).not.toContain(` t="s"`);
  });

  it("locks Locked's A1 and unlocks exactly B2", async () => {
    const model = await read();
    const locked = sheetNamed(model, "Locked");
    expect(styleAt(model, locked, F.locked.protectedCell.ref).protection?.locked).not.toBe(false);
    const unlocked = locked.rows.flatMap((row) => row.cells.filter((cell) => styleAt(model, locked, formatA1({ row: row.index, col: cell.col })).protection?.locked === false).map((cell) => formatA1({ row: row.index, col: cell.col })));
    expect(unlocked).toEqual([F.locked.unlockedCell.ref]);
    expect(cellAt(locked, F.locked.unlockedCell.ref)?.styleId).toBe(F.locked.unlockedCell.xf);
  });

  it("caches what the formulas compute, no formula is volatile, none outside Edit reads it, and an edit to B2 recalculates only Edit", async () => {
    const model = await read();
    expect(formulaTexts(model).filter((entry) => VOLATILE.test(entry.formula))).toEqual([]);
    expect(formulaTexts(model).filter((entry) => entry.sheet !== "Edit" && /Edit!/i.test(entry.formula))).toEqual([]);
    const edit = sheetNamed(model, "Edit");
    const rowsBefore = model.sheets.map((sheet) => sheet.rows);
    const full = recalculate(model, [], { full: true, reason: "manual" });
    expect(full.valueChanged).toEqual([]);
    expect(full.circular).toEqual([]);
    // Only the data table — never evaluated here — is stale after a manual pass.
    const table = parseRange(F.edit.dataTable.ref)!;
    const tableKeys = [table.top, table.top + 1, table.top + 2].map((row) => cellKey(edit.id, row, table.left));
    expect(full.stale).toEqual(tableKeys);
    model.sheets.forEach((sheet, index) => expect(sheet.rows, sheet.name).toBe(rowsBefore[index]));

    const fresh = await read();
    const target = parseA1(F.edit.target.ref)!;
    const editSheet = sheetNamed(fresh, "Edit");
    fresh.sheets = fresh.sheets.map((sheet) =>
      sheet !== editSheet
        ? sheet
        : {
            ...sheet,
            rows: sheet.rows.map((row) =>
              row.index !== target.row ? row : { ...row, cells: row.cells.map((cell) => (cell.col === target.col ? { ...cell, value: 99 } : cell)) },
            ),
          },
    );
    const result = recalculate(fresh, [cellKey(editSheet.id, target.row, target.col)], { reason: "edit" });
    expect(result.changedSheetIds).toEqual([editSheet.id]);
    expect(cellAt(sheetNamed(fresh, "Edit"), F.edit.target.dependent.ref)?.value).toBe(198);
  });
});

describe("spec122-shapes.ods (LibreOffice-shaped)", () => {
  const O = SPEC122_SHAPES_ODS;
  const bytes = buildSpec122ShapesOdsBytes();
  const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, "ods", "spec122-shapes");

  it("is the committed e2e fixture, byte for byte, and a whole package in Calc's entry order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec122-shapes.ods")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(O.parts);
  });

  it("reads Data (active) and Locked: values, formulas, only Locked protected, the structure not", async () => {
    const model = await read();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(O.tableNames);
    expect(model.activeSheetIndex).toBe(O.activeSheetIndex);
    expect(model.sheets.map((sheet) => sheet.protection)).toEqual([undefined, { sheet: true }]);
    expect(model.structureProtected).toBeUndefined();
    const data = sheetNamed(model, "Data");
    const locked = sheetNamed(model, "Locked");
    expect(storedValues(data)).toEqual(O.data.values);
    expect(storedValues(locked)).toEqual(O.locked.values);
    for (const [sheet, formulas] of [
      [data, O.data.formulas],
      [locked, O.locked.formulas],
    ] as const) {
      for (const [a1, fact] of Object.entries(formulas)) {
        expect(cellAt(sheet, a1), a1).toMatchObject({ formula: fact.formula, value: fact.value });
        expect(sheet.sourceXml, a1).toContain(`table:formula="${fact.wire.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll(">", "&gt;")}"`);
      }
    }
  });

  it("keeps LibreOffice's shapes in Data's source: the tails, header rows, the row group, the span, local names and formats after the rows", async () => {
    const model = await read();
    const data = sheetNamed(model, "Data");
    const source = data.sourceXml!;
    expect(source).toContain(`table:number-rows-repeated="${O.data.trailingRowRun}"`);
    expect(source).toContain(`table:number-columns-repeated="${O.data.trailingColumnRun}" table:default-cell-style-name="Default"/>`);
    expect(O.data.usedRows + O.data.trailingRowRun).toBe(1_048_576);
    expect(O.data.usedColumns + O.data.trailingColumnRun).toBe(16_384);
    expect(source.indexOf("<table:table-header-rows>")).toBeLessThan(source.indexOf("<table:table-row-group>"));
    expect(source).toMatch(/<\/table:table-row><\/table:table>$|<\/calcext:conditional-formats><\/table:table>$/);
    expect(source.indexOf("<table:named-expressions>")).toBeGreaterThan(source.lastIndexOf("</table:table-row>"));
    expect(source.indexOf("<calcext:conditional-formats>")).toBeGreaterThan(source.indexOf("</table:named-expressions>"));
    expect(data.merges).toEqual([parseRange(O.data.merge.range)]);
    for (const a1 of O.data.merge.covered) {
      expect(cellAt(data, a1), a1).toBeUndefined();
    }
    expect(model.odsContentEnvelope!.prefix).toContain(`<table:content-validation table:name="${O.data.validation.name}"`);
  });

  it("captures nothing it models: the annotation is a note and the link paragraph a link (SPEC-135); the validation name is modelled (SPEC-130)", async () => {
    const model = await read();
    const data = sheetNamed(model, "Data");
    // SPEC-135: the annotation is lifted out of preservedXml into a note, and
    // the plain text:a is modelled inline content (a link), not preserved text.
    expect(cellsWith(data, "preservedXml")).toEqual({});
    expect(cellsWith(data, "preservedText")).toEqual({});
    expect((data.notes ?? []).map((note) => [formatA1(note.ref), note.author, note.text, note.source?.odsAnnotation])).toEqual([
      [O.data.annotation.a1, O.data.annotation.creator, O.data.annotation.text, O.data.annotation.xml],
    ]);
    expect((data.hyperlinks ?? []).map((link) => [formatA1({ row: link.range.top, col: link.range.left }), link.target])).toEqual([
      [O.data.link.a1, { kind: "url", href: O.data.link.href }],
    ]);
    expect(getCell(data, parseA1(O.data.link.a1)!.row, 0)?.value).toBe(O.data.link.text);
    // SPEC-130: B16's table:content-validation-name is no longer a preserved
    // attribute — it is the Data sheet's rule.
    expect(cellsWith(data, "preservedAttrs")).toEqual({});
    expect(data.validations.map((rule) => ({ ranges: rule.ranges, type: rule.type, operator: rule.operator, formula1: rule.formula1, formula2: rule.formula2, odsName: rule.odsName }))).toEqual([
      { ranges: [parseRange(O.data.validation.a1)], type: "whole", operator: "between", formula1: "1", formula2: "10", odsName: O.data.validation.name },
    ]);
    expect(model.sheets.map((sheet) => [cellsWith(sheet, "preservedXml"), cellsWith(sheet, "preservedText"), cellsWith(sheet, "preservedAttrs")]).slice(1)).toEqual([[{}, {}, {}]]);
  });

  it("reads the styles: the unprotected cell, de-ce1 by name, the red-negative map", async () => {
    const model = await read();
    const data = sheetNamed(model, "Data");
    const locked = sheetNamed(model, "Locked");
    expect(styleAt(model, locked, O.locked.unprotectedCell.a1)).toMatchObject({ protection: { locked: false }, sourceStyleName: O.locked.unprotectedCell.styleName });
    expect(styleAt(model, locked, O.locked.protectedCell.a1).protection?.locked).not.toBe(false);
    expect(locked.sourceXml).toContain(`table:protection-key="${O.locked.protectionKey}" table:protection-key-digest-algorithm="${O.locked.digestAlgorithm}">${O.locked.tableProtection}`);
    expect(styleAt(model, data, O.data.collision.a1).sourceStyleName).toBe(O.data.collision.styleName);
    expect(styleAt(model, data, O.data.redNegative.a1)).toMatchObject({ sourceStyleName: O.data.redNegative.styleName, numberFormat: O.data.redNegative.code });
  });

  it("caches what the formulas compute: a full manual recalculation changes nothing", async () => {
    const model = await read();
    expect(formulaTexts(model).filter((entry) => VOLATILE.test(entry.formula))).toEqual([]);
    const rowsBefore = model.sheets.map((sheet) => sheet.rows);
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.stale).toEqual([]);
    expect(result.circular).toEqual([]);
    model.sheets.forEach((sheet, index) => expect(sheet.rows, sheet.name).toBe(rowsBefore[index]));
  });
});
