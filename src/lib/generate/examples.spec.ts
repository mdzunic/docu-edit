// @vitest-environment node
// SPEC-140 — the three workbook examples, each run exactly as a user runs it
// (`node examples/<name>.js`, SPEC-045's examples/<name>.js +
// examples/_impl/<name>.ts + lib/runTs.js pattern), then its output in
// examples/output/ re-read with readWorkbook and one computed total checked —
// the cached value the example's formula was given before the file was
// written.
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { getCell, parseA1 } from "../spreadsheet/address";
import type { CellValue, SpreadsheetModel } from "../spreadsheet/model";
import { readWorkbook } from "./readWorkbook";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..", "..", "..");
const output = (name: string): string => join(root, "examples", "output", name);

/** Runs `node examples/<name>.js` from the repository root. */
function run(name: string): string {
  return execFileSync(process.execPath, [join("examples", `${name}.js`)], { cwd: root, encoding: "utf8", timeout: 120_000 });
}

function cell(model: SpreadsheetModel, sheet: string, ref: string): { value: CellValue; formula?: string } | undefined {
  const at = parseA1(ref)!;
  const found = model.sheets.find((each) => each.name === sheet)!;
  return getCell(found, at.row, at.col);
}

describe("the workbook examples (AC-47, AC-48, AC-49, AC-50)", () => {
  it("sales-report.js: monthly sales by region with SUM/AVERAGE totals and a growth column", async () => {
    expect(run("sales-report")).toContain("sales-report.xlsx");
    const model = await readWorkbook(readFileSync(output("sales-report.xlsx")));
    expect(cell(model, "Sales", "E6")).toMatchObject({ value: 14115, formula: "SUM(E2:E5)" });
    expect(cell(model, "Sales", "B7")).toMatchObject({ value: 1110, formula: "AVERAGE(B2:B5)" });
    expect(cell(model, "Sales", "F2")?.value).toBeCloseTo(0.175, 10);
    const sheet = model.sheets[0]!;
    expect(sheet.frozen).toEqual({ rows: 1, cols: 0 });
    expect(model.styles[getCell(sheet, 0, 0)!.styleId!]).toMatchObject({ font: { bold: true }, fill: { kind: "solid", color: "#1f4e79" } });
    expect(model.styles[getCell(sheet, 1, 4)!.styleId!]!.numberFormat).toBe("$#,##0.00");
    expect(model.styles[getCell(sheet, 1, 5)!.styleId!]!.numberFormat).toBe("0.0%");
    expect(sheet.columns.map((span) => [span.min, span.max, span.sourceWidthChars])).toEqual([
      [0, 0, 12],
      [1, 4, 13],
      [5, 5, 9],
    ]);
  });

  it("csv-to-xlsx.js: orders.csv through parseCsvText and setRows' parseText, widths from content", async () => {
    expect(run("csv-to-xlsx")).toContain("orders.xlsx");
    const model = await readWorkbook(readFileSync(output("orders.xlsx")));
    const total = cell(model, "Orders", "H12");
    expect(total?.formula).toBe("SUM(H2:H11)");
    expect(total?.value).toBeCloseTo(11197.6, 6);
    // Typed where lossless, text otherwise — and never a formula.
    expect(cell(model, "Orders", "A2")?.value).toBe("00120");
    expect(cell(model, "Orders", "B2")?.value).toBe(46266);
    expect(cell(model, "Orders", "F3")?.value).toBe(1240);
    expect(cell(model, "Orders", "G3")?.value).toBe(0.05);
    expect(cell(model, "Orders", "C6")).toEqual({ col: 2, value: '=HYPERLINK("http://example.com","click")', sstIndex: expect.any(Number) });
    expect(model.sheets[0]!.columns.length).toBeGreaterThan(0);
    expect(model.sheets[0]!.columns.every((span) => span.widthPx !== undefined)).toBe(true);
  });

  it("budget.js: Income, Expenses and Summary from one description, as .xlsx and .ods", async () => {
    expect(run("budget")).toContain("budget.ods");
    for (const format of ["xlsx", "ods"] as const) {
      const model = await readWorkbook(readFileSync(output(`budget.${format}`)));
      expect(model.format).toBe(format);
      expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Summary", "Income", "Expenses"]);
      expect(cell(model, "Summary", "B3")).toMatchObject({ value: 4885, formula: "SUM(Income!B2:B4)" });
      expect(cell(model, "Summary", "B5")).toMatchObject({ value: 1645, formula: "B3-B4" });
      expect(model.sheets[0]!.merges).toEqual([{ top: 0, left: 0, bottom: 0, right: 1 }]);
    }
  });
});
