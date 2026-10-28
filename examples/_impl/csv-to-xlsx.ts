// SPEC-140: a CSV export turned into a formatted workbook. The committed
// examples/data/orders.csv is split into records by SPEC-126's parseCsvText,
// then written with setRows' `parseText` — every field through the same
// lossless inference the editor's CSV import uses, so "1,240.00", "5%" and
// "2026-09-02" become a number, a percentage and a date with their formats,
// while the order numbers keep their leading zeros and the customer field
// that starts with "=" stays text (parseText never makes a formula: no CSV
// formula injection). Column widths come from the content; a totals row
// sums the amounts. Run via `node examples/csv-to-xlsx.js`; writes
// examples/output/orders.xlsx.
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { colToLetters } from "../../src/lib/spreadsheet/address";
import { parseCsvText } from "../../src/lib/spreadsheet/csv/parse";
import { WorkbookBuilder } from "../../src/lib/generate/workbook";

const csv = await readFile(new URL("../data/orders.csv", import.meta.url), "utf8");
const records: string[][] = [];
parseCsvText(csv, ",", (fields) => {
  records.push(fields);
});
const [header, ...orders] = records;
if (!header) {
  throw new Error("orders.csv is empty");
}

const totalRow = orders.length + 2;
const amount = colToLetters(header.indexOf("Amount"));
const workbook = new WorkbookBuilder({ title: "September orders" })
  .addSheet("Orders")
  .setRow("A1", header)
  .setRows("A2", orders, { parseText: true })
  .setCell(`A${totalRow}`, "Total")
  .setCell(`${amount}${totalRow}`, { formula: `SUM(${amount}2:${amount}${totalRow - 1})`, numberFormat: "#,##0.00" })
  .style("A1:H1", { font: { bold: true }, fill: "#e2efda", borders: { bottom: "thin" } })
  .style(`A${totalRow}:H${totalRow}`, { font: { bold: true }, borders: { top: "double" } })
  .freeze({ rows: 1 });

// Widths from content: the longest field in each column, plus a little room.
header.forEach((title, col) => {
  const longest = Math.max(title.length, ...orders.map((fields) => (fields[col] ?? "").length));
  workbook.columnWidth(colToLetters(col), { chars: Math.min(255, longest + 2) });
});

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outPath = `${outDir}orders.xlsx`;
await writeFile(outPath, await workbook.toBuffer("xlsx"));
console.log(`Wrote ${outPath}`);
