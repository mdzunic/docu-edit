// SPEC-045 AC-21: an invoice with a table, generated from data with no UI
// involved. Run via `node examples/invoice.js` (see ../lib/runTs.js for why
// this file itself is TypeScript but the command line stays plain `node`).
import { writeFile, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { DocumentBuilder } from "../../src/lib/generate/builder";

const invoice = {
  number: "INV-1042",
  date: "2026-08-11",
  billTo: "Acme Robotics, 42 Industrial Way, Springfield",
  lineItems: [
    { description: "Widget assembly", qty: 12, unitPrice: 18.5 },
    { description: "On-site calibration", qty: 2, unitPrice: 240 },
    { description: "Support contract (annual)", qty: 1, unitPrice: 600 },
  ],
};

const total = invoice.lineItems.reduce((sum, item) => sum + item.qty * item.unitPrice, 0);

const builder = new DocumentBuilder({ title: `Invoice ${invoice.number}` });

builder
  .addHeading("Invoice", 1)
  .addParagraph([
    { text: `Invoice #: ${invoice.number}` },
  ])
  .addParagraph([{ text: `Date: ${invoice.date}` }])
  .addParagraph([{ text: "Bill to: ", bold: true }, invoice.billTo])
  .addTable({
    header: true,
    rows: [
      ["Description", "Qty", "Unit price", "Amount"],
      ...invoice.lineItems.map((item) => [
        item.description,
        String(item.qty),
        `$${item.unitPrice.toFixed(2)}`,
        `$${(item.qty * item.unitPrice).toFixed(2)}`,
      ]),
      [{ runs: [{ text: "Total", bold: true }] }, "", "", { runs: [{ text: `$${total.toFixed(2)}`, bold: true }] }],
    ],
  })
  .addParagraph("Payment due within 30 days. Thank you for your business.");

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
const outPath = fileURLToPath(new URL("../output/invoice.docx", import.meta.url));
await writeFile(outPath, await builder.toBuffer("docx"));
console.log(`Wrote ${outPath}`);
