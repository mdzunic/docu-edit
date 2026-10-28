// @vitest-environment node
// SPEC-140 — generation is DOM-free and fast. This file runs in vitest's
// plain Node environment (no jsdom): `document`, `window` and `DOMParser`
// are simply absent, so a module on the build and write path that touched
// one would throw — or, as the from-scratch .xlsx style table once did,
// silently write a different file. workbook.ts, workbookDescription.ts and
// both writers run here with no polyfill; `ensureNodeXmlGlobals` belongs to
// the READ path (readWorkbook) alone, which only this file's last test calls,
// to prove what Node wrote. And the performance bar, the shape
// performance.spec.ts has: a 10,000-row × 10-column workbook with a formula
// column, a styled header row and two number-formatted columns builds and
// writes in under 2 seconds, one spec per format.
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import { readWorkbook } from "./readWorkbook";
import { WorkbookBuilder, type WorkbookOutputFormat } from "./workbook";
import { workbookFromDescription } from "./workbookDescription";

const ROWS = 10_000;

function q3Sales(): WorkbookBuilder {
  return new WorkbookBuilder({ title: "Q3 sales" })
    .defineStyle("header", { font: { bold: true }, fill: "#dce6f1", borders: { bottom: "thin" } })
    .addSheet("Sales")
    .setRow("A1", ["Region", "July", "August", "September", "Total"])
    .setRows("A2", [["East", 1200, 1350, 1410], ["West", 980, 1010, 1130]])
    .setFormula("E2", "SUM(B2:D2)")
    .setFormula("E3", "SUM(B3:D3)")
    .setCell("F2", { date: "2026-10-08" })
    .style("A1:E1", "header")
    .numberFormat("B2:E3", "$#,##0.00")
    .columnWidth("A", { chars: 14 })
    .merge("A5:E5")
    .freeze({ rows: 1 });
}

describe("the build and write path is DOM-free (AC-43)", () => {
  it("runs with no document, window or DOMParser in scope", () => {
    expect(typeof document).toBe("undefined");
    expect(typeof window).toBe("undefined");
    expect(typeof DOMParser).toBe("undefined");
    expect(typeof XMLSerializer).toBe("undefined");
  });

  it("builds, materialises and writes both formats through workbook.ts with no polyfill", async () => {
    const builder = q3Sales().setRows("A6", [["1,234.50", "45%", "=SUM(A1:A2)"]], { parseText: true });
    expect(builder.toWorkbook().sheets[0]!.rows[1]!.cells[4]!.value).toBe(3960);
    for (const format of ["xlsx", "ods"] as const) {
      const parts = await unzip(await builder.toBytes(format));
      expect(parts.size).toBeGreaterThan(4);
    }
    expect(Buffer.isBuffer(await builder.toBuffer())).toBe(true);
    expect(workbookFromDescription({ sheets: [{ name: "S", cells: { A1: 1, A2: { formula: "A1*2" } } }] }).sheets[0]!.rows[1]!.cells[0]!.value).toBe(2);
    expect(typeof DOMParser).toBe("undefined");
  });

  it("writes in Node exactly what a browser reads back (the from-scratch style table needs no DOM)", async () => {
    const xlsx = await q3Sales().toBytes("xlsx");
    const ods = await q3Sales().toBytes("ods");
    expect(typeof DOMParser).toBe("undefined");
    // Reading needs the polyfill — readWorkbook installs it, so this runs last.
    for (const [bytes, format] of [[xlsx, "xlsx"], [ods, "ods"]] as const) {
      const model = await readWorkbook(bytes);
      const header = model.styles[model.sheets[0]!.rows[0]!.cells[0]!.styleId!]!;
      expect(header.font?.bold, format).toBe(true);
      expect(header.fill, format).toEqual({ kind: "solid", color: "#dce6f1" });
      expect(header.borders?.bottom?.style, format).toBe("thin");
      expect(model.styles[model.sheets[0]!.rows[1]!.cells[4]!.styleId!]!.numberFormat, format).toBe("$#,##0.00");
    }
  });
});

/** 10,000 data rows × 10 columns: a region, seven numbers, a date, and a
 * formula total; a styled header row; two number-formatted columns. */
function tenThousandRows(format: WorkbookOutputFormat): WorkbookBuilder {
  const builder = new WorkbookBuilder({ title: "Performance", format })
    .addSheet("Data")
    .setRow("A1", ["Region", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Updated", "Total"])
    .style("A1:J1", { font: { bold: true }, fill: "#dce6f1", borders: { bottom: "thin" } });
  const rows: (string | number | Date)[][] = [];
  for (let r = 0; r < ROWS; r += 1) {
    rows.push([`Region ${r + 1}`, ...Array.from({ length: 7 }, (_, c) => (r * 7 + c) % 1000), new Date(Date.UTC(2026, 0, 1 + (r % 365)))]);
  }
  builder.setRows("A2", rows);
  for (let r = 2; r <= ROWS + 1; r += 1) {
    builder.setFormula(`J${r}`, `SUM(B${r}:H${r})`);
  }
  return builder.numberFormat(`B2:B${ROWS + 1}`, "$#,##0.00").numberFormat(`J2:J${ROWS + 1}`, "#,##0.00");
}

describe("performance (AC-44)", () => {
  for (const format of ["xlsx", "ods"] as const) {
    it(`builds and writes a 10,000-row × 10-column .${format} in under 2 seconds`, async () => {
      const start = performance.now();
      const bytes = await tenThousandRows(format).toBytes();
      const elapsedMs = performance.now() - start;

      const parts = await unzip(bytes);
      const sheet = new TextDecoder().decode(parts.get(format === "xlsx" ? "xl/worksheets/sheet1.xml" : "content.xml"));
      expect(sheet).toContain(format === "xlsx" ? `<c r="J10001" s="` : "of:=SUM([.B10001:.H10001])");
      expect(elapsedMs).toBeLessThan(2000);
    });
  }
});
