// SPEC-140: a household budget as ONE JSON workbook description — Income,
// Expenses and a Summary whose cross-sheet formulas total the other two,
// under a merged title — written as both .xlsx and .ods. The description is
// plain data (it could have come from a file or an HTTP request); the
// schema it follows is schema/workbook-description.json. Run via
// `node examples/budget.js`; writes examples/output/budget.xlsx and
// examples/output/budget.ods.
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { WorkbookBuilder } from "../../src/lib/generate/workbook";
import type { WorkbookDescription } from "../../src/lib/generate/workbookDescription";

const income = [
  ["Salary", 4200],
  ["Freelance", 650],
  ["Interest", 35],
];
const expenses = [
  ["Rent", 1450],
  ["Groceries", 520],
  ["Utilities", 210],
  ["Transport", 160],
  ["Savings", 900],
];

const listSheet = (name: string, rows: (string | number)[][], tabColor: string): NonNullable<WorkbookDescription["sheets"]>[number] => ({
  name,
  tabColor,
  cells: { A1: "Item", B1: "Monthly" },
  data: [{ at: "A2", rows }],
  styles: [
    { range: "A1:B1", style: "header" },
    { range: `B2:B${rows.length + 1}`, numberFormat: "$#,##0.00" },
  ],
  columns: [{ columns: "A", width: { chars: 16 } }, { columns: "B", width: { chars: 12 } }],
  freeze: { rows: 1 },
});

const budget: WorkbookDescription = {
  title: "Budget 2026",
  styles: {
    header: { font: { bold: true }, fill: "#d9e1f2", borders: { bottom: "thin" } },
    title: { font: { bold: true, sizePt: 16 }, alignment: { horizontal: "center" } },
  },
  sheets: [
    {
      name: "Summary",
      cells: {
        A1: { value: "Budget 2026", style: "title" },
        A3: "Income",
        B3: { formula: `SUM(Income!B2:B${income.length + 1})` },
        A4: "Expenses",
        B4: { formula: `SUM(Expenses!B2:B${expenses.length + 1})` },
        A5: { value: "Left over", style: { font: { bold: true } } },
        B5: { formula: "B3-B4", style: { font: { bold: true }, borders: { top: "thin", bottom: "double" } } },
        A6: "Saved",
        B6: { formula: "Expenses!B6/B3", numberFormat: "0.0%" },
      },
      styles: [{ range: "B3:B5", numberFormat: "$#,##0.00" }],
      columns: [{ columns: "A", width: { chars: 16 } }, { columns: "B", width: { chars: 14 } }],
      rows: [{ rows: 1, height: { pt: 24 } }],
      merges: ["A1:B1"],
    },
    listSheet("Income", income, "#70ad47"),
    listSheet("Expenses", expenses, "#c00000"),
  ],
};

const builder = WorkbookBuilder.fromDescription(budget);
const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
for (const format of ["xlsx", "ods"] as const) {
  const outPath = `${outDir}budget.${format}`;
  await writeFile(outPath, await builder.toBuffer(format));
  console.log(`Wrote ${outPath}`);
}
