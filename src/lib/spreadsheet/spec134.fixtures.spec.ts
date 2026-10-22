// SPEC-134's find-and-Go-To fixture (spec134-find.xlsx, built by
// fixtures/build.ts): the builder is deterministic, the committed e2e file is
// its bytes, the package is whole with exactly the listed parts, every XML
// part parses and Excel's shapes are really in them — and, read back, it
// holds what the e2e scenario's counts rest on: 3 / 2 / 4 / 8 matches for
// "north", the protected cell Replace all skips, the Totals name and the
// Formulas selection Go To Special makes.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { parseXml } from "../core/xml";
import { unzip } from "../core/zip";
import { formatRange } from "./address";
import { compileCellQuery, createSheetSearch, planReplace, replaceReport, resolveGoToTarget, selectSpecial, type SheetFindOptions } from "./find";
import { buildSpec134FindXlsxBytes, SPEC134_FIND_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { singleCell } from "./selection";

const F = SPEC134_FIND_XLSX;
const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const committed = (): Uint8Array => new Uint8Array(readFileSync(join(E2E_FIXTURES, F.file)));
const BASE: SheetFindOptions = { matchCase: false, entireCell: false, regex: false, lookIn: "values", within: "sheet" };

describe("spec134-find.xlsx — the package", () => {
  it("is deterministic, the committed bytes, whole, with exactly the listed parts, every XML part parsing", async () => {
    const built = await buildSpec134FindXlsxBytes();
    expect(equalBytes(built, await buildSpec134FindXlsxBytes())).toBe(true);
    expect(equalBytes(committed(), built)).toBe(true);
    expect(await packageIntegrity(built)).toEqual([]);
    const parts = await unzip(built);
    expect([...parts.keys()]).toEqual([...F.parts]);
    for (const [name, bytes] of parts) {
      expect(() => parseXml(decoder.decode(bytes)), name).not.toThrow();
    }
  });

  it("is Excel-shaped: shared strings, xr:uids, dyDescent, a numFmt 164, sheetProtection, a hidden sheet and the name", async () => {
    const parts = await unzip(committed());
    const text = (name: string): string => decoder.decode(parts.get(name));
    expect(text("xl/sharedStrings.xml")).toContain("<si><t>North</t></si>");
    expect(text("xl/sharedStrings.xml")).toContain("<si><t>north</t></si>");
    expect(text("xl/worksheets/sheet1.xml")).toContain(`xr:uid="${F.sheets.Sales.uid}"`);
    expect(text("xl/worksheets/sheet1.xml")).toContain('x14ac:dyDescent="0.35"');
    expect(text("xl/worksheets/sheet1.xml")).toContain('<c r="C2"><f>IF(A2="North",1,0)</f><v>1</v></c>');
    expect(text("xl/worksheets/sheet1.xml")).toContain('<c r="C5" t="e"><f>B5/0</f><v>#DIV/0!</v></c>');
    expect(text("xl/styles.xml")).toContain('<numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>');
    expect(text("xl/worksheets/sheet2.xml")).toContain('<sheetProtection sheet="1" objects="1" scenarios="1"/>');
    expect(text("xl/workbook.xml")).toContain('<sheet name="Archive" sheetId="3" state="hidden" r:id="rId3"/>');
    expect(text("xl/workbook.xml")).toContain('<definedNames><definedName name="Totals">Sales!$B$10</definedName></definedNames>');
  });
});

describe("spec134-find.xlsx — read back, what the e2e rests on", () => {
  async function open() {
    const model = await readSpreadsheetBytes(committed(), "xlsx", F.title);
    const ctx = workbookFormatContext(model, SPREADSHEET_LOCALES["en-US"]);
    const sales = model.sheets[0]!;
    const count = (options: Partial<SheetFindOptions>): string[] => {
      const all = { ...BASE, ...options };
      const query = compileCellQuery("north", all);
      if (query === null || "error" in query) {
        throw new Error("no query");
      }
      const search = createSheetSearch(model, query, all, sales.id, singleCell({ row: 0, col: 0 }), ctx);
      search.step(() => false);
      return search.hits.map((hit) => `${model.sheets.find((sheet) => sheet.id === hit.sheetId)!.name}!${formatRange({ top: hit.row, left: hit.col, bottom: hit.row, right: hit.col })}`);
    };
    return { model, ctx, sales, count };
  }

  it("reads three sheets — Archive hidden, Locked protected — and the currency amounts", async () => {
    const { model, sales, count } = await open();
    expect(model.sheets.map((sheet) => [sheet.name, sheet.visibility, sheet.protection?.sheet === true])).toEqual([
      ["Sales", "visible", false],
      ["Locked", "visible", true],
      ["Archive", "hidden", false],
    ]);
    expect(sales.rows.find((row) => row.index === 1)!.cells.find((cell) => cell.col === 1)!.value).toBe(1200);
    expect(count({ lookIn: "values" })).toEqual(["Sales!A2", "Sales!A4", "Sales!A6"]);
  });

  it("finds 3, 2 (entire cell), 4 (workbook) and 8 (formulas) cells for \"north\"", async () => {
    const { count } = await open();
    expect(count({})).toHaveLength(3);
    expect(count({ entireCell: true })).toEqual(["Sales!A2", "Sales!A6"]);
    expect(count({ within: "workbook" })).toEqual(["Sales!A2", "Sales!A4", "Sales!A6", "Locked!A1"]);
    expect(count({ within: "workbook", lookIn: "formulas" })).toEqual([
      "Sales!A2",
      "Sales!C2",
      "Sales!C3",
      "Sales!A4",
      "Sales!C4",
      "Sales!A6",
      "Sales!C6",
      "Locked!A1",
    ]);
  });

  it("plans the e2e's Replace all: 7 occurrences in 7 cells, the Locked cell skipped as protected", async () => {
    const { model, ctx, sales } = await open();
    const options: SheetFindOptions = { ...BASE, within: "workbook", lookIn: "formulas" };
    const query = compileCellQuery("north", options);
    if (query === null || "error" in query) {
      throw new Error("no query");
    }
    const search = createSheetSearch(model, query, options, sales.id, singleCell({ row: 0, col: 0 }), ctx);
    search.step(() => false);
    const plan = planReplace(model, search.hits, query, "N.", options, ctx);
    expect(replaceReport(plan, plan.edits.length)).toBe("Replaced 7 occurrences in 7 cells · skipped 1 protected");
    expect(plan.edits.find((edit) => edit.row === 1 && edit.col === 2)?.text).toBe('=IF(A2="N.",1,0)');
  });

  it("goes to Totals (B10) and selects C2:C6 and B10 for Go To Special ▸ Formulas", async () => {
    const { model, sales } = await open();
    expect(resolveGoToTarget(model, "Totals", sales.id)).toEqual({ sheetId: sales.id, ranges: [{ top: 9, left: 1, bottom: 9, right: 1 }] });
    expect(resolveGoToTarget(model, "Sales!B3:C4", sales.id)).toEqual({ sheetId: sales.id, ranges: [{ top: 2, left: 1, bottom: 3, right: 2 }] });
    const special = selectSpecial(model, sales, singleCell({ row: 9, col: 1 }), "formulas", { numbers: true, text: true, logicals: true, errors: true });
    expect(Array.isArray(special) ? special.map(formatRange) : special).toEqual(["C2:C6", "B10"]);
  });
});
