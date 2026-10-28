// SPEC-140: a monthly sales-by-region report — SUM and AVERAGE totals, a
// growth-percentage formula column, currency and percent formats, a bold
// filled header, a frozen header row and column widths — generated from
// data with no UI involved. Every formula's cached value is computed before
// the file is written, so a preview pane or an indexer shows the totals
// too. Run via `node examples/sales-report.js`; writes
// examples/output/sales-report.xlsx.
import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { WorkbookBuilder } from "../../src/lib/generate/workbook";

const months = ["July", "August", "September"];
const sales = [
  { region: "East", values: [1200, 1350, 1410] },
  { region: "West", values: [980, 1010, 1130] },
  { region: "North", values: [1500, 1420, 1610] },
  { region: "South", values: [760, 840, 905] },
];

const first = 2;
const last = first + sales.length - 1;
const totalRow = last + 1;
const averageRow = last + 2;

const report = new WorkbookBuilder({ title: "Q3 sales by region" })
  .defineStyle("header", { font: { bold: true, color: "#ffffff" }, fill: "#1f4e79", borders: { bottom: "medium" }, alignment: { horizontal: "center" } })
  .defineStyle("totals", { font: { bold: true }, borders: { top: "thin" } })
  .addSheet("Sales", { tabColor: "#1f4e79" })
  .setRow("A1", ["Region", ...months, "Q3 total", "Growth"]);

sales.forEach((row, i) => {
  const r = first + i;
  report.setRow(`A${r}`, [row.region, ...row.values]).setFormula(`E${r}`, `SUM(B${r}:D${r})`).setFormula(`F${r}`, `(D${r}-B${r})/B${r}`);
});

// Totals and averages across regions, per month and for the quarter.
report.setCell(`A${totalRow}`, "Total").setCell(`A${averageRow}`, "Average");
for (const column of ["B", "C", "D", "E"]) {
  report.setFormula(`${column}${totalRow}`, `SUM(${column}${first}:${column}${last})`).setFormula(`${column}${averageRow}`, `AVERAGE(${column}${first}:${column}${last})`);
}
report.setFormula(`F${totalRow}`, `(D${totalRow}-B${totalRow})/B${totalRow}`);

report
  .style("A1:F1", "header")
  .style(`A${totalRow}:F${averageRow}`, "totals")
  .numberFormat(`B${first}:E${averageRow}`, "$#,##0.00")
  .numberFormat(`F${first}:F${totalRow}`, "0.0%")
  .columnWidth("A", { chars: 12 })
  .columnWidth("B:E", { chars: 13 })
  .columnWidth("F", { chars: 9 })
  .rowHeight(1, { pt: 20 })
  .freeze({ rows: 1 });

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outPath = `${outDir}sales-report.xlsx`;
await writeFile(outPath, await report.toBuffer("xlsx"));
console.log(`Wrote ${outPath}`);
