// SPEC-133 round trip through .ods (ods/namedExpressions.ts, ods/tableXml.ts,
// ods/writer.ts): spec133-names.ods untouched writes content.xml
// byte-identical; workbook and sheet names defined, edited and deleted survive
// save → reopen (equalStructural) and a second write is byte-identical; a
// structural edit moves the names through the model alone (AC-22 retired the
// retained rows) and the writer writes the moved names; spec133-names-tables
// .xlsx downloaded as .ods reopens with identical values, its names but no
// tables, one database range per table, structured references as A1 ranges
// and each table's look baked into generated automatic styles — model.styles
// untouched. Opened and saved as the app does (readSpreadsheetBytes,
// writeOds, spreadsheetDownload).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { equalStructural } from "../equality";
import { getCell, parseA1 } from "./address";
import { spreadsheetDownload } from "./download";
import { buildSpec133NamesOdsBytes, buildSpec133NamesTablesXlsxBytes, SPEC133_NAMES_ODS, SPEC133_NAMES_TABLES_XLSX } from "./fixtures/build";
import type { CellValue, DefinedName, SpreadsheetModel } from "./model";
import { defineName, deleteName, updateName } from "./names";
import { writeOds } from "./ods/writer";
import { readSpreadsheetBytes } from "./open";
import { applyStructuralEdit } from "./structure";
import { cellStyle } from "./styles";

const decoder = new TextDecoder();

const readOds = (bytes: Uint8Array): Promise<SpreadsheetModel> => readSpreadsheetBytes(bytes, "ods", "names");

async function contentXml(bytes: Uint8Array): Promise<string> {
  return decoder.decode((await unzip(bytes)).get("content.xml")!);
}

/** A model's names as the .ods holds them — the workbook's, then each
 * sheet's in sheet order — each scope by its sheet's name, so two reads
 * compare. */
function namesByScope(model: SpreadsheetModel): unknown[] {
  const scoped = (scopeSheetId: string | undefined): unknown[] =>
    model.names
      .filter((name) => name.scopeSheetId === scopeSheetId)
      .map(({ scopeSheetId: _id, ...name }) => ({ ...name, scope: model.sheets.find((sheet) => sheet.id === scopeSheetId)?.name ?? null }));
  return [...scoped(undefined), ...model.sheets.flatMap((sheet) => scoped(sheet.id))];
}

function named(model: SpreadsheetModel, name: string, scope?: string): DefinedName {
  const scopeId = scope === undefined ? undefined : model.sheets.find((sheet) => sheet.name === scope)!.id;
  const found = model.names.find((each) => each.name === name && each.scopeSheetId === scopeId);
  expect(found, name).toBeDefined();
  return found!;
}

/** Every stored value, sheet → A1 → value. */
function values(model: SpreadsheetModel): Record<string, Record<string, CellValue>> {
  const out: Record<string, Record<string, CellValue>> = {};
  for (const sheet of model.sheets) {
    const cells: Record<string, CellValue> = {};
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.value !== null) {
          cells[`${row.index}:${cell.col}`] = cell.value;
        }
      }
    }
    out[sheet.name] = cells;
  }
  return out;
}

function formulaAt(model: SpreadsheetModel, sheet: string, a1: string): string | undefined {
  const ref = parseA1(a1)!;
  return getCell(model.sheets.find((each) => each.name === sheet)!, ref.row, ref.col)?.formula;
}

/** Each part of a written package that differs from `other`'s. */
async function differingParts(a: Uint8Array, b: Uint8Array): Promise<string[]> {
  const left = await unzip(a);
  const right = await unzip(b);
  return [...new Set([...left.keys(), ...right.keys()])].filter((name) => !left.has(name) || !right.has(name) || !equalBytes(left.get(name)!, right.get(name)!));
}

describe("SPEC-133 round trip — spec133-names.ods", () => {
  it("an untouched workbook writes content.xml byte-identical (and every other part)", async () => {
    const bytes = await buildSpec133NamesOdsBytes();
    const model = await readOds(bytes);
    expect(model.names.map((name) => name.name)).toEqual([...SPEC133_NAMES_ODS.globalNames, ...SPEC133_NAMES_ODS.localNames].map((name) => name.name));
    const written = await writeOds(model);
    expect(equalBytes((await unzip(written)).get("content.xml")!, (await unzip(bytes)).get("content.xml")!)).toBe(true);
    expect(await differingParts(written, bytes)).toEqual([]);
  });

  it("workbook and sheet names defined, edited and deleted → save → reopen: the names equalStructural; a second write is byte-identical", async () => {
    const model = await readOds(await buildSpec133NamesOdsBytes());
    const sheet1 = model.sheets[0]!;
    const sheet2 = model.sheets[1]!;
    // Workbook scope: a new named range, an edited expression, a deleted range.
    defineName(model, { name: "Margin", formula: "Sheet2!$B$2:$B$4" });
    const rate = named(model, "Rate");
    updateName(model, rate, { ...rate, formula: "0.25" });
    deleteName(model, named(model, "Report"));
    // Sheet scope: a renamed expression (its formulas follow), a deleted range, a new expression on the other (clean) sheet.
    const localExpr = named(model, "LocalExpr", "Sheet1");
    updateName(model, localExpr, { ...localExpr, name: "Doubled" });
    deleteName(model, named(model, "LocalRange", "Sheet1"));
    defineName(model, { name: "Half", scopeSheetId: sheet2.id, formula: "Sheet2!$B$2/2" });
    expect(formulaAt(model, "Sheet1", "C2")).toBe("Doubled");

    const first = await writeOds(model);
    const xml = await contentXml(first);
    // The workbook's element stays right after the last table; each sheet's is its table's last child.
    expect(xml).toContain(
      `</table:table><table:named-expressions>` +
        `<table:named-range table:name="Next" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet2.A2"/>` +
        `<table:named-range table:name="Prices" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet1.$A$2:.$A$4"/>` +
        `<table:named-expression table:name="Rate" table:base-cell-address="$Sheet1.$A$1" table:expression="of:=0.25"/>` +
        `<table:named-range table:name="Margin" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet2.$B$2:.$B$4"/>` +
        `</table:named-expressions>`,
    );
    expect(xml).toContain(`<table:named-expressions><table:named-expression table:name="Half" table:base-cell-address="$Sheet2.$A$1" table:expression="of:=[$Sheet2.$B$2]/2"/></table:named-expressions></table:table>`);
    expect(xml).toContain(`<table:named-expression table:name="Doubled" table:base-cell-address="$Sheet1.$A$1" table:expression="of:=[.XFD1]*2"/></table:named-expressions></table:table>`);
    expect(xml).not.toContain(`table:name="Report"`);
    expect(xml).not.toContain(`table:name="LocalRange"`);
    expect(xml).toContain(`table:formula="of:=Doubled"`);

    const reopened = await readOds(first);
    expect(equalStructural(namesByScope(reopened), namesByScope(model))).toBe(true);
    expect(reopened.sheets.map((sheet) => sheet.name)).toEqual([sheet1.name, sheet2.name]);
    expect(formulaAt(reopened, "Sheet1", "C3")).toBe("Doubled");

    const second = await writeOds(reopened);
    expect(await differingParts(second, first)).toEqual([]);
  });

  it("a structural edit moves the names through the model alone — the retained text is not rewritten (AC-22) — and the save writes the moved names", async () => {
    const source = await buildSpec133NamesOdsBytes();
    const model = await readOds(source);
    const suffix = model.odsContentEnvelope!.suffix;
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: model.sheets[0]!.id, at: 0, count: 2 })).toEqual({ ok: true });
    expect(named(model, "Prices").formula).toBe("Sheet1!$A$4:$A$6");
    expect(named(model, "LocalRange", "Sheet1").formula).toBe("Sheet1!$B$5:$B$7");
    // SPEC-124's rows for named ranges and expressions are retired: the retained element keeps its bytes.
    expect(model.odsContentEnvelope!.suffix).toBe(suffix);
    expect(model.odsContentEnvelope!.suffix).toContain(SPEC133_NAMES_ODS.namedExpressionsXml);

    const written = await writeOds(model);
    const xml = await contentXml(written);
    expect(xml).not.toContain(SPEC133_NAMES_ODS.namedExpressionsXml);
    expect(xml).toContain(`<table:named-range table:name="Prices" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet1.$A$4:.$A$6"/>`);
    expect(xml).toContain(`<table:named-range table:name="LocalRange" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet1.$B$5:.$B$7"/>`);
    const reopened = await readOds(written);
    expect(equalStructural(namesByScope(reopened), namesByScope(model))).toBe(true);
    expect(await differingParts(await writeOds(reopened), written)).toEqual([]);
  });
});

describe("SPEC-133 round trip — spec133-names-tables.xlsx downloaded as .ods", () => {
  const X = SPEC133_NAMES_TABLES_XLSX;

  async function converted(): Promise<{ model: SpreadsheetModel; bytes: Uint8Array; reopened: SpreadsheetModel; stylesBefore: string }> {
    const model = await readSpreadsheetBytes(await buildSpec133NamesTablesXlsxBytes(), "xlsx", "names");
    const stylesBefore = JSON.stringify(model.styles);
    const { bytes } = await spreadsheetDownload(model, "ods");
    return { model, bytes, reopened: await readOds(bytes), stylesBefore };
  }

  it("reopens with identical values, the names but no tables, and a second write byte-identical", async () => {
    const { model, bytes, reopened } = await converted();
    expect(values(reopened)).toEqual(values(model));
    expect(reopened.sheets.every((sheet) => sheet.tables === undefined)).toBe(true);
    // Every name but the reserved _xlnm.* ones and the _xlfn. placeholder, each in its scope, as the model holds it.
    const writable = (each: SpreadsheetModel): unknown[] =>
      (namesByScope(each) as { name: string; formula: string; scope: string | null }[]).map(({ name, formula, scope }) => ({ name, formula, scope }));
    expect(writable(reopened)).toEqual(writable(model).filter((name) => !/^_xl(?:nm|fn)\./i.test((name as { name: string }).name)));
    expect(reopened.names.map((name) => name.name)).toEqual(expect.arrayContaining(["TaxRate", "Above", "Rates", "Helper", "Local"]));
    expect(reopened.names.some((name) => name.name.startsWith("_xl"))).toBe(false);
    expect(await differingParts(await writeOds(reopened), bytes)).toEqual([]);
  });

  it("writes one database range per table, structured references as A1 ranges, and each table's look as generated automatic styles — model.styles untouched", async () => {
    const { model, bytes, reopened, stylesBefore } = await converted();
    expect(JSON.stringify(model.styles)).toBe(stylesBefore);
    const xml = await contentXml(bytes);
    const sales = X.tables.Sales;
    expect(xml).toContain(`<table:database-range table:name="Sales" table:target-range-address="Data.A1:Data.E6" table:contains-header="true" table:display-filter-buttons="true"/>`);
    expect(xml).toContain(`<table:database-range table:name="Table1" table:target-range-address="Lists.A2:Lists.B4" table:contains-header="false" table:display-filter-buttons="false"/>`);
    expect(xml.match(/<table:database-range table:name="(?:Sales|Table1)"/g)).toHaveLength(2);
    expect(xml).not.toContain("Sales[");

    // Structured references, hosted at their cells, became A1 ranges.
    expect(formulaAt(model, "Data", "G2")).toBe("SUM(Sales[Amount])");
    expect(formulaAt(reopened, "Data", "G2")).toBe("SUM($D$2:$D$5)");
    expect(formulaAt(reopened, "Data", "D2")).toBe("$B2*$C2");
    expect(formulaAt(reopened, "Data", "D6")).toBe("SUBTOTAL(109,$D$2:$D$5)");
    expect(formulaAt(reopened, "Data", "G3")).toBe("$D$6");
    expect(formulaAt(reopened, "Data", "G5")).toBe("TaxRate*G2");

    // The header row carries the look (Medium 2: fill accent1, bold white) beneath its own style; the body keeps its own number format.
    const data = reopened.sheets.find((sheet) => sheet.name === "Data")!;
    const at = (a1: string) => {
      const ref = parseA1(a1)!;
      return cellStyle(reopened, getCell(data, ref.row, ref.col)?.styleId ?? 0);
    };
    expect(at("A1").fill).toBeDefined();
    expect(at("A1").font?.bold).toBe(true);
    expect(at("D2").numberFormat).toBe("#,##0");
    expect(at("G2").fill).toBeUndefined();
    expect(sales.style.name).toBe("TableStyleMedium2");
  });
});
