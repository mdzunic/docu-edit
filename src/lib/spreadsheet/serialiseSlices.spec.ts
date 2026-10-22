// SPEC-137 AC-44/45: createSliceClock's due()/yield() contract, and that
// the xlsx and ODS sheet serialisers' async twins actually yield through it
// while still producing output identical to their synchronous originals.
import { describe, expect, it, vi } from "vitest";

vi.mock("../core/idleScheduler", () => ({ yieldToMainThread: vi.fn(() => Promise.resolve()) }));

import { yieldToMainThread } from "../core/idleScheduler";
import { createSliceClock, SERIALISE_SLICE_MS } from "./serialiseSlices";
import type { Row, Sheet } from "./model";
import { serializeWorksheet, serializeWorksheetAsync, type XlsxSheetContext } from "./xlsx/sheetXml";
import { serializeTable, serializeTableAsync, type OdsTableContext } from "./ods/tableXml";
import { OdsStyleRegistry } from "./ods/styles";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";

const yieldSpy = yieldToMainThread as unknown as ReturnType<typeof vi.fn>;

describe("createSliceClock", () => {
  it("is not due immediately, but is once sliceMs has elapsed", () => {
    const clock = createSliceClock(0);
    // sliceMs=0: any elapsed wall time (even sub-millisecond) trips it.
    let due = false;
    for (let i = 0; i < 1_000_000 && !due; i++) {
      due = clock.due();
    }
    expect(due).toBe(true);
  });

  it("yield() resets the clock so it is not immediately due again", async () => {
    const clock = createSliceClock(SERIALISE_SLICE_MS);
    await clock.yield();
    expect(clock.due()).toBe(false);
  });
});

function manyRows(count: number): Row[] {
  const rows: Row[] = [];
  for (let i = 0; i < count; i++) {
    rows.push({ index: i, cells: [{ col: 0, value: i }, { col: 1, value: `text ${i}` }] });
  }
  return rows;
}

function sheetOf(rows: Row[]): Sheet {
  return { conditionalFormats: [], validations: [], id: "s", name: "S", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

describe("serializeWorksheetAsync", () => {
  it("yields while serialising a large sheet and matches serializeWorksheet's output", async () => {
    yieldSpy.mockClear();
    const sheet = sheetOf(manyRows(300));
    const ctx: XlsxSheetContext = { active: false, theme: undefined, mdw: 7, strings: { indexFor: () => 0 }, styles: { xfFor: (id) => id } };
    const sync = serializeWorksheet(sheet, undefined, ctx);
    const clock = createSliceClock(0); // always due: every check yields
    const asyncOut = await serializeWorksheetAsync(sheet, undefined, ctx, clock);
    expect(asyncOut).toBe(sync);
    expect(yieldSpy.mock.calls.length).toBeGreaterThan(0);
  });

  it("produces identical output to serializeWorksheet when nothing is due to yield", async () => {
    yieldSpy.mockClear();
    const sheet = sheetOf(manyRows(10));
    const ctx: XlsxSheetContext = { active: false, theme: undefined, mdw: 7, strings: { indexFor: () => 0 }, styles: { xfFor: (id) => id } };
    const sync = serializeWorksheet(sheet, undefined, ctx);
    const clock = createSliceClock(SERIALISE_SLICE_MS);
    const asyncOut = await serializeWorksheetAsync(sheet, undefined, ctx, clock);
    expect(asyncOut).toBe(sync);
    expect(yieldSpy).not.toHaveBeenCalled();
  });
});

describe("serializeTableAsync", () => {
  const EN_US = SPREADSHEET_LOCALES["en-US"];

  it("yields while serialising a large sheet and matches serializeTable's output", async () => {
    yieldSpy.mockClear();
    const sheet = sheetOf(manyRows(300));
    const ctx: OdsTableContext = { model: { styles: [{}], dateSystem: "1900", format: "ods" }, registry: new OdsStyleRegistry(), locale: EN_US };
    const sync = serializeTable(sheet, ctx);
    const asyncCtx: OdsTableContext = { model: { styles: [{}], dateSystem: "1900", format: "ods" }, registry: new OdsStyleRegistry(), locale: EN_US };
    const clock = createSliceClock(0); // always due: every row yields
    const asyncOut = await serializeTableAsync(sheet, asyncCtx, clock);
    expect(asyncOut).toBe(sync);
    expect(yieldSpy.mock.calls.length).toBeGreaterThan(0);
  });
});
