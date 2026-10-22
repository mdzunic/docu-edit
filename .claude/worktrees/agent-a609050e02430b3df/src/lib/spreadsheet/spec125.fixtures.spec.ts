// SPEC-125's fixtures (spec125-fill.xlsx / spec125-fill.ods, built by
// fixtures/build.ts): the committed e2e files are the builders' bytes, both
// packages are whole, the reader gives what SPEC125_FILL records — the copy
// block, the merge, the cut cells, the fill seeds, the cross-sheet formulas
// (every cached value what its formula computes) — and each of the "Series"
// sheet's seed columns continues as its series kind says.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { getCell, parseA1, parseRange } from "./address";
import { buildSpec125FillOdsBytes, buildSpec125FillXlsxBytes, SPEC125_FILL } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { recalculate } from "./formula/recalc";
import type { Cell, Sheet, SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { readSpreadsheetBytes } from "./open";
import { detectSeries, type SeriesKind } from "./series";
import { cellStyle, effectiveStyleId } from "./styles";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const EN_US = SPREADSHEET_LOCALES["en-US"];

const sheetNamed = (model: SpreadsheetModel, name: string): Sheet => model.sheets.find((sheet) => sheet.name === name)!;
const cellAt = (sheet: Sheet, a1: string): Cell | undefined => {
  const at = parseA1(a1)!;
  return getCell(sheet, at.row, at.col);
};

/** The series kind each "Series" column's label names. */
const KINDS: Record<string, SeriesKind> = {
  Number: "number",
  Step: "number",
  Trend: "number",
  Decimal: "number",
  Day: "date",
  "Month end": "date",
  Month: "date",
  Year: "date",
  Week: "date",
  Hour: "time",
  Weekday: "name",
  Months: "name",
  Quarter: "quarter",
  Qtr: "quarter",
  Item: "counter",
  Padded: "counter",
  "Odd items": "counter",
  Ordinal: "ordinal",
  Formula: "repeat",
  Logical: "repeat",
  Text: "repeat",
};

const FORMATS = [
  { name: "spec125-fill.xlsx", format: "xlsx" as const, build: buildSpec125FillXlsxBytes },
  { name: "spec125-fill.ods", format: "ods" as const, build: buildSpec125FillOdsBytes },
];

for (const { name, format, build } of FORMATS) {
  describe(`${name} (${format === "xlsx" ? "Excel" : "LibreOffice"}-shaped)`, () => {
    const F = SPEC125_FILL;
    const bytes = build();
    const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, format, "spec125-fill");

    it("is the committed e2e fixture, byte for byte, and a whole package", async () => {
      const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, name)));
      expect(equalBytes(committed, await bytes)).toBe(true);
      expect(await packageIntegrity(await bytes)).toEqual([]);
    });

    it("reads Sheet1's copy block, merge, cut cells and fill seeds", async () => {
      const model = await read();
      expect(model.sheets.map((sheet) => sheet.name)).toEqual(F.sheetNames);
      expect(model.activeSheetIndex).toBe(F.activeSheetIndex);
      const s = sheetNamed(model, "Sheet1");
      expect(cellAt(s, "B2")?.value).toBe(F.block.B2);
      expect(cellAt(s, "C2")?.value).toBe(F.block.C2);
      expect(cellAt(s, "B3")).toMatchObject({ formula: F.block.B3.formula, value: F.block.B3.value });
      expect(cellAt(s, "C3")).toMatchObject({ formula: F.block.C3.formula, value: F.block.C3.value });
      expect(s.merges).toEqual([parseRange(F.merge.ref)]);
      expect(cellAt(s, "E1")?.value).toBe(F.merge.value);
      expect(["A10", "A11", "A12"].map((ref) => cellAt(s, ref)?.value)).toEqual([F.cut.A10, F.cut.A11, F.cut.A12]);
      expect(cellAt(s, F.monday.ref)?.value).toBe(F.monday.value);
      expect(cellAt(s, "J3")).toBeUndefined();
      for (let r = F.guide.first; r <= F.guide.last; r += 1) {
        expect(cellAt(s, `${F.guide.column}${r}`)?.value).toBe(r - F.guide.first + 1);
      }
      expect(cellAt(s, F.doubleClick.ref)).toMatchObject({ formula: F.doubleClick.formula, value: F.doubleClick.value });
      expect(cellAt(s, F.fillDown.ref)?.value).toBe(F.fillDown.value);
    });

    it("reads Sheet2's cross-sheet formulas, every cached value what its formula computes", async () => {
      const model = await read();
      const s = sheetNamed(model, "Sheet2");
      expect(cellAt(s, "A1")).toMatchObject({ formula: F.sheet2.A1.formula, value: F.sheet2.A1.value });
      expect(cellAt(s, "A2")).toMatchObject({ formula: F.sheet2.A2.formula, value: F.sheet2.A2.value });
      const before = JSON.stringify(model.sheets.map((sheet) => sheet.rows));
      recalculate(model, [], { full: true, reason: "manual" });
      expect(JSON.stringify(model.sheets.map((sheet) => sheet.rows))).toBe(before);
    });

    it("holds a seed column per series kind on the Series sheet, each continuing as its kind", async () => {
      const model = await read();
      const s = sheetNamed(model, "Series");
      F.series.forEach((column, c) => {
        const letters = String.fromCharCode(65 + c);
        expect(cellAt(s, `${letters}1`)?.value, column.label).toBe(column.label);
        const seeds = column.seeds.map((_, r) => {
          const cell = cellAt(s, `${letters}${r + 2}`)!;
          const code = formatCodeOf(cellStyle(model, effectiveStyleId(s, r + 1, c)), EN_US);
          return { value: cell.value, ...(cell.formula !== undefined ? { formula: cell.formula } : {}), ...(code !== undefined ? { code } : {}) };
        });
        expect(detectSeries(seeds, EN_US, { dateSystem: model.dateSystem }).kind, column.label).toBe(KINDS[column.label]);
      });
    });
  });
}
