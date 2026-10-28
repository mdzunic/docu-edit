// SPEC-140 "Materialisation" — what workbook.spec.ts cannot observe from the
// model alone: a builder's cells are materialised in batches, never one
// copy-on-write applyStylePatch per cell (quadratic on a large sheet), and a
// workbook past MAX_WORKBOOK_CELLS is refused with the product's own
// WORKBOOK_TOO_LARGE_MESSAGE rather than written. Both through module mocks
// — a passthrough spy on applyStylePatch, and a ceiling of 50 cells in place
// of 2,000,000 — which is why they live in a file of their own.
import { describe, expect, it, vi } from "vitest";
import { applyStylePatch } from "../spreadsheet/formatting";
import { MAX_WORKBOOK_CELLS } from "../spreadsheet/model";
import { WORKBOOK_TOO_LARGE_MESSAGE } from "../spreadsheet/errors";
import { DocumentGenerationError } from "./errors";
import { WorkbookBuilder } from "./workbook";

vi.mock("../spreadsheet/formatting", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../spreadsheet/formatting")>();
  return { ...actual, applyStylePatch: vi.fn(actual.applyStylePatch) };
});

vi.mock("../spreadsheet/model", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../spreadsheet/model")>()),
  MAX_WORKBOOK_CELLS: 50,
}));

const spy = vi.mocked(applyStylePatch);

describe("materialised once, in batches (AC-24)", () => {
  it("a column of dated cells takes one applyStylePatch, not one per cell", () => {
    const dates = Array.from({ length: 40 }, (_, r) => [{ date: `2026-01-${String((r % 28) + 1).padStart(2, "0")}` }]);
    const builder = new WorkbookBuilder().addSheet("S").setRows("A1", dates);
    spy.mockClear();
    const model = builder.toWorkbook();
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.mock.calls[0]![2]).toEqual([{ top: 0, left: 0, bottom: 39, right: 0 }]);
    expect(new Set(model.sheets[0]!.rows.map((row) => row.cells[0]!.styleId)).size).toBe(1);
  });

  it("keeps call order around range calls, and per cell", () => {
    const builder = new WorkbookBuilder()
      .addSheet("S")
      .setRows("A1", [[{ value: 1, numberFormat: "0.0" }], [{ value: 2, numberFormat: "0.0" }]])
      .numberFormat("A1:A2", "0.00")
      .setCell("A2", { value: 3, numberFormat: "0%", style: { borders: { all: "thin" }, font: { bold: true } } })
      .setCell("A1", { value: 4, style: { borders: { all: "thin" }, font: { bold: true } } });
    spy.mockClear();
    const model = builder.toWorkbook();
    // [cells 0.0], [range 0.00], [cells: borders, font, 0%] — the second
    // cell's borders and font join the first's groups.
    expect(spy).toHaveBeenCalledTimes(5);
    const style = (row: number): unknown => model.styles[model.sheets[0]!.rows[row]!.cells[0]!.styleId!];
    expect(style(0)).toMatchObject({ numberFormat: "0.00", font: { bold: true }, borders: { top: { style: "thin" } } });
    expect(style(1)).toMatchObject({ numberFormat: "0%", font: { bold: true }, borders: { top: { style: "thin" } } });
  });
});

/** `n` one-cell rows. */
const rows = (n: number): number[][] => Array.from({ length: n }, (_, r) => [r]);

describe("the cell ceiling (AC-26)", () => {
  it("refuses a workbook past MAX_WORKBOOK_CELLS with WORKBOOK_TOO_LARGE_MESSAGE", async () => {
    expect(MAX_WORKBOOK_CELLS).toBe(50);
    expect(new WorkbookBuilder().addSheet("A").setRows("A1", rows(30)).addSheet("B").setRows("A1", rows(20)).toWorkbook().sheets).toHaveLength(2);
    const over = new WorkbookBuilder().addSheet("A").setRows("A1", rows(30)).addSheet("B").setRows("A1", rows(21));
    expect(() => over.toWorkbook()).toThrow(DocumentGenerationError);
    expect(() => over.toWorkbook()).toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
    await expect(over.toBytes()).rejects.toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
    // Style-only cells count too: they are stored.
    const styled = new WorkbookBuilder().addSheet("A").setRows("A1", rows(40)).style("B1:B11", { fill: "#ffff00" });
    expect(() => styled.toWorkbook()).toThrow(`sheets: ${WORKBOOK_TOO_LARGE_MESSAGE}`);
  });
});
