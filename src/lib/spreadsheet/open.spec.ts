// SPEC-116 "Reader versions" and SPEC-117 "Reader version":
// SPREADSHEET_READER_VERSION and the upgrade list, assembly, and
// upgradeSpreadsheetModel — a current record comes back untouched; an older
// one with no dirty sheet is re-read from its retained package keeping its
// id, title and every sheet's id (by position); a dirty sheet instead runs
// every pending step, in order. SPEC-117 moves the version to 2 and appends a
// step that throws (no sheet can be dirty before SPEC-121), and a workbook a
// SPEC-116 reader stored reopens with its styles, sizes, merges and views.
// SPEC-119 moves it to 4 (its step throws "SPEC-119 upgrade: …"), and a
// record SPEC-118's reader stored re-reads its clean sheets and gains its
// formulas. SPEC-120 moves it to 5 (its step throws "SPEC-120 upgrade: …"): a
// record SPEC-119's reader stored gains its calculation settings, and a
// stored forceFullCalcOnLoad survives the re-read. SPEC-122 moves it to 6:
// its step is the first to run on a dirty sheet — preserved cell content and
// sheet protection come back from the retained source onto the cells whose
// value and formula are unchanged — and a dirty sheet's style ids move into
// the re-read style table.
import { afterEach, describe, expect, it } from "vitest";
import { getCell } from "./address";
import {
  buildMiniBookOdsBytes,
  buildMiniBookXlsxBytes,
  buildSpec117StylesOdsBytes,
  buildSpec117StylesXlsxBytes,
  buildSpec118FormatsOdsBytes,
  buildSpec118FormatsXlsxBytes,
  buildSpec118GeneratedXlsxBytes,
  buildSpec119FormulasOdsBytes,
  buildSpec119FormulasXlsxBytes,
  buildSpec120CalcOdsBytes,
  buildSpec120CalcXlsxBytes,
  buildSpec120FullCalcXlsxBytes,
  buildSpec120ManualXlsxBytes,
} from "./fixtures/build";
import type { Sheet, SpreadsheetModel } from "./model";
import {
  assembleSpreadsheet,
  READER_UPGRADES,
  readSpreadsheetBytes,
  SPREADSHEET_READER_VERSION,
  upgradeSpreadsheetModel,
  type ReaderUpgrade,
} from "./open";

/** The shipped upgrade list, restored after each test that swaps it. */
const SHIPPED_UPGRADES = [...READER_UPGRADES];

/** A record as an older reader would have stored it: version 0, the same
 * package retained, but nothing extracted from it yet (empty rows). */
async function versionZero(build: () => Promise<Uint8Array>, format: "xlsx" | "ods"): Promise<SpreadsheetModel> {
  const model = await readSpreadsheetBytes(await build(), format, "Renamed by the user");
  return {
    ...model,
    readerVersion: 0,
    sheets: model.sheets.map((sheet, index) => ({ ...sheet, id: `stored-sheet-${index}`, rows: [] })),
  };
}

/** A record exactly as SPEC-116's reader stored it (version 1): values and
 * runs, but none of SPEC-117's fields — no styles, no theme, no layout, no
 * view settings — on the model, its sheets, rows, cells or runs. */
async function spec116Record(build: () => Promise<Uint8Array>, format: "xlsx" | "ods"): Promise<SpreadsheetModel> {
  const model = await readSpreadsheetBytes(await build(), format, "Stored by SPEC-116");
  const plain = JSON.parse(JSON.stringify(model)) as Record<string, unknown> & { sheets: Record<string, unknown>[] };
  for (const key of ["styles", "defaultStyleId", "theme", "indexedColors"]) {
    delete plain[key];
  }
  plain.sheets = model.sheets.map((sheet, index) => ({
    id: `stored-sheet-${index}`,
    name: sheet.name,
    visibility: sheet.visibility,
    kind: sheet.kind,
    rows: sheet.rows
      .filter((row) => row.cells.some((cell) => cell.value !== null))
      .map((row) => ({
        index: row.index,
        cells: row.cells
          .filter((cell) => cell.value !== null)
          .map((cell) => (cell.runs ? { col: cell.col, value: cell.value, runs: cell.runs.map((run) => ({ text: run.text })) } : { col: cell.col, value: cell.value })),
      })),
    ...(sheet.partName !== undefined ? { partName: sheet.partName } : {}),
    ...(sheet.sourceXml !== undefined ? { sourceXml: sheet.sourceXml } : {}),
  }));
  plain.readerVersion = 1;
  return plain as unknown as SpreadsheetModel;
}

/** A record exactly as SPEC-117's reader stored it (version 2): styles,
 * layout and views, but no number format on any style and none of the
 * styles SPEC-118 derives for implied formats (nor the cells pointing at
 * them). */
async function spec117Record(build: () => Promise<Uint8Array>, format: "xlsx" | "ods"): Promise<SpreadsheetModel> {
  const model = JSON.parse(JSON.stringify(await readSpreadsheetBytes(await build(), format, "Stored by SPEC-117"))) as SpreadsheetModel;
  const derived = new Set(model.styles.flatMap((style, index) => (index > 0 && style.sourceXf === undefined && style.sourceStyleName === undefined ? [index] : [])));
  model.styles = model.styles
    .filter((_, index) => !derived.has(index))
    .map(({ numberFormat: _format, builtinFormatId: _builtin, ...style }) => style);
  model.sheets = model.sheets.map((sheet, index) => ({
    ...sheet,
    id: `stored-sheet-${index}`,
    rows: sheet.rows.map((row) => ({
      ...row,
      cells: row.cells.map(({ styleId, ...cell }) => (styleId === undefined || derived.has(styleId) ? cell : { ...cell, styleId })),
    })),
  }));
  model.readerVersion = 2;
  return model;
}

/** A record exactly as SPEC-118's reader stored it (version 3): no formula
 * fields on any cell (formula, formulaUnparsed, arrayRange, dynamicArray, cm,
 * dataTableXml), no showFormulas on any sheet — and no cell SPEC-119 stores
 * only for its formula (one with no cached value or style). */
async function spec118Record(build: () => Promise<Uint8Array>, format: "xlsx" | "ods"): Promise<SpreadsheetModel> {
  const model = JSON.parse(JSON.stringify(await readSpreadsheetBytes(await build(), format, "Stored by SPEC-118"))) as SpreadsheetModel;
  model.sheets = model.sheets.map(({ showFormulas: _showFormulas, ...sheet }, index) => ({
    ...sheet,
    id: `stored-sheet-${index}`,
    rows: sheet.rows.map((row) => ({
      ...row,
      cells: row.cells
        .filter((cell) => cell.value !== null || cell.styleId !== undefined)
        .map(({ formula: _f, formulaUnparsed: _u, arrayRange: _a, dynamicArray: _d, cm: _c, dataTableXml: _t, ...cell }) => cell),
    })),
  }));
  model.readerVersion = 3;
  return model;
}

afterEach(() => {
  READER_UPGRADES.splice(0, READER_UPGRADES.length, ...SHIPPED_UPGRADES);
});

describe("reader versions", () => {
  it("SPEC-117 moved the reader to version 2, SPEC-118 to 3, SPEC-119 to 4, SPEC-120 to 5 and SPEC-122 to 6, with one contiguous step each (from 1, 2, 3, 4, 5)", () => {
    expect(SPREADSHEET_READER_VERSION).toBe(6);
    expect(READER_UPGRADES.map((step) => step.from)).toEqual([1, 2, 3, 4, 5]);
  });

  it("SPEC-117's, SPEC-118's, SPEC-119's and SPEC-120's steps refuse a dirty sheet loudly — none can exist before SPEC-121", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    const [spec117, spec118, spec119, spec120] = READER_UPGRADES;
    for (const step of [spec117!, spec118!]) {
      expect(() => step.upgradeDirtySheet(model, model.sheets[0]!, new Map())).toThrow("unreachable: no sheet is dirty before SPEC-121");
    }
    expect(() => spec119!.upgradeDirtySheet(model, model.sheets[0]!, new Map())).toThrow(
      new Error("SPEC-119 upgrade: a sheet cannot be dirty before SPEC-121"),
    );
    expect(() => spec120!.upgradeDirtySheet(model, model.sheets[0]!, new Map())).toThrow(
      new Error("SPEC-120 upgrade: a sheet cannot be dirty before SPEC-121"),
    );
  });

  it("stamps every fresh read with the current version and a card preview", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    expect(model.kind).toBe("spreadsheet");
    expect(model.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(model.title).toBe("mini-book");
    expect(model.preview?.sheetName).toBe("Summary");
    expect(model.preview?.cells[1]?.[1]).toBe("Revenue");
  });

  it("omits empty optional fields when assembling (styles and the default style id are required)", () => {
    const model = assembleSpreadsheet("t", "xlsx", {
      sheets: [{ id: "s", name: "S", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] }],
      styles: [{}],
      activeSheetIndex: 0,
      dateSystem: "1900",
      opaque: [],
      properties: {},
    });
    expect(Object.keys(model).toSorted()).toEqual(
      [
        "activeSheetIndex",
        "dateSystem",
        "defaultStyleId",
        "format",
        "id",
        "kind",
        "opaque",
        "preview",
        "readerVersion",
        "sheets",
        "styles",
        "title",
      ].toSorted(),
    );
    expect(model.styles).toEqual([{}]);
    expect(model.defaultStyleId).toBe(0);
  });

  it("returns a current record as the same object", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    expect(await upgradeSpreadsheetModel(model)).toBe(model);
  });

  for (const [format, build] of [
    ["xlsx", buildMiniBookXlsxBytes],
    ["ods", buildMiniBookOdsBytes],
    ["xlsx", buildSpec117StylesXlsxBytes],
    ["ods", buildSpec117StylesOdsBytes],
  ] as const) {
    it(`re-reads a version-0 ${format} record (${build.name}) from opaque, keeping id, title and sheet ids`, async () => {
      const stored = await versionZero(build, format);
      const upgraded = await upgradeSpreadsheetModel(stored);
      expect(upgraded).not.toBe(stored);
      expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
      expect(upgraded.id).toBe(stored.id);
      expect(upgraded.title).toBe("Renamed by the user");
      expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(stored.sheets.map((sheet) => sheet.id));
      expect(upgraded.opaque).toEqual(stored.opaque);
      // What the retained source holds is extracted again.
      expect(upgraded.sheets.some((sheet) => sheet.rows.length > 0)).toBe(true);
      expect(upgraded.preview?.cells.length).toBeGreaterThan(0);
    });

    it(`a ${format} workbook (${build.name}) stored by SPEC-116's reader reopens with its styles, sizes, merges and views`, async () => {
      const stored = await spec116Record(build, format);
      const fresh = await readSpreadsheetBytes(await build(), format, "Stored by SPEC-116");
      const upgraded = await upgradeSpreadsheetModel(stored);
      expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
      expect(upgraded.id).toBe(stored.id);
      expect(upgraded.title).toBe("Stored by SPEC-116");
      expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(stored.sheets.map((sheet) => sheet.id));
      // Everything SPEC-117 reads comes back from the retained source, as a
      // fresh read of the same bytes has it.
      expect(upgraded.styles).toEqual(fresh.styles);
      expect(upgraded.theme).toEqual(fresh.theme);
      expect(upgraded.indexedColors).toEqual(fresh.indexedColors);
      upgraded.sheets.forEach((sheet, index) => {
        const { id: _upgradedId, ...layout } = sheet;
        const { id: _freshId, ...expected } = fresh.sheets[index]!;
        expect(layout, sheet.name).toEqual(expected);
      });
    });
  }

  for (const [format, build] of [
    ["xlsx", buildSpec118FormatsXlsxBytes],
    ["xlsx", buildSpec118GeneratedXlsxBytes],
    ["ods", buildSpec118FormatsOdsBytes],
  ] as const) {
    it(`a ${format} workbook (${build.name}) stored by SPEC-117's reader reopens with its number formats`, async () => {
      const stored = await spec117Record(build, format);
      expect(stored.styles.some((style) => style.numberFormat !== undefined)).toBe(false);
      const fresh = await readSpreadsheetBytes(await build(), format, "Stored by SPEC-117");
      const upgraded = await upgradeSpreadsheetModel(stored);
      // SPEC-119 has moved the current version on; the upgrade lands there.
      expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
      expect(upgraded.id).toBe(stored.id);
      expect(upgraded.title).toBe("Stored by SPEC-117");
      expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(stored.sheets.map((sheet) => sheet.id));
      // The formats (and the derived implied-format styles) come back from
      // the retained source, exactly as a fresh read has them.
      expect(upgraded.styles).toEqual(fresh.styles);
      expect(upgraded.styles.some((style) => style.numberFormat !== undefined)).toBe(true);
      upgraded.sheets.forEach((sheet, index) => {
        expect(sheet.rows).toEqual(fresh.sheets[index]!.rows);
      });
    });
  }

  for (const [format, build] of [
    ["xlsx", buildSpec119FormulasXlsxBytes],
    ["ods", buildSpec119FormulasOdsBytes],
  ] as const) {
    it(`a ${format} workbook (${build.name}) stored by SPEC-118's reader re-reads its clean sheets and gains its formulas`, async () => {
      const stored = await spec118Record(build, format);
      expect(stored.sheets.some((sheet) => sheet.rows.some((row) => row.cells.some((cell) => cell.formula !== undefined)))).toBe(false);
      const fresh = await readSpreadsheetBytes(await build(), format, "Stored by SPEC-118");
      const upgraded = await upgradeSpreadsheetModel(stored);
      expect(upgraded).not.toBe(stored);
      expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
      expect(upgraded.id).toBe(stored.id);
      expect(upgraded.title).toBe("Stored by SPEC-118");
      expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(stored.sheets.map((sheet) => sheet.id));
      upgraded.sheets.forEach((sheet, index) => {
        const { id: _upgradedId, ...read } = sheet;
        const { id: _freshId, ...expected } = fresh.sheets[index]!;
        expect(read, sheet.name).toEqual(expected);
      });
      const calc = upgraded.sheets[0]!;
      expect(getCell(calc, 4, 1)?.formula).toBe("SUM(B2:B4)");
      if (format === "xlsx") {
        expect(upgraded.sheets[1]!.showFormulas).toBe(true);
      }
    });
  }

  it("runs every pending step, in order, on a dirty sheet and re-reads the clean ones", async () => {
    const stored = await versionZero(buildMiniBookXlsxBytes, "xlsx");
    const dirty = stored.sheets[1] as Sheet & { dirty?: boolean };
    dirty.dirty = true;
    dirty.rows = [{ index: 0, cells: [{ col: 0, value: "edited" }] }];
    const calls: string[] = [];
    const step = (from: number): ReaderUpgrade => ({
      from,
      upgradeDirtySheet(model, sheet, sourceParts) {
        calls.push(`${from}:${sheet.id}:${sourceParts.has(sheet.partName!)}:${model.id}`);
      },
    });
    // The machinery under test, not SPEC-117's step (which refuses a dirty
    // sheet): two recording steps, registered out of order.
    READER_UPGRADES.splice(0, READER_UPGRADES.length, step(1), step(0));
    const upgraded = await upgradeSpreadsheetModel(stored);
    expect(calls).toEqual([`0:${dirty.id}:true:${stored.id}`, `1:${dirty.id}:true:${stored.id}`]);
    // The dirty sheet keeps its edits; the clean ones are re-read.
    expect(upgraded.sheets[1]).toBe(dirty);
    expect(getCell(upgraded.sheets[1]!, 0, 0)?.value).toBe("edited");
    expect(getCell(upgraded.sheets[0]!, 1, 1)?.value).toBe("Revenue");
    expect(upgraded.sheets[0]!.id).toBe(stored.sheets[0]!.id);
  });
});

/** A record exactly as SPEC-119's reader stored it (version 4): no `calc`. */
async function spec119Record(build: () => Promise<Uint8Array>, format: "xlsx" | "ods"): Promise<SpreadsheetModel> {
  const read = JSON.parse(JSON.stringify(await readSpreadsheetBytes(await build(), format, "Stored by SPEC-119"))) as SpreadsheetModel;
  const { calc: _calc, ...model } = read;
  return { ...model, readerVersion: 4, sheets: model.sheets.map((sheet, index) => ({ ...sheet, id: `stored-sheet-${index}` })) };
}

describe("SPEC-120's calculation settings in stored records", () => {
  for (const [format, build, calc] of [
    ["xlsx", buildSpec120FullCalcXlsxBytes, { fullCalcOnLoad: true }],
    ["xlsx", buildSpec120ManualXlsxBytes, { manual: true }],
    ["ods", buildSpec120CalcOdsBytes, { caseSensitive: true }],
  ] as const) {
    it(`a ${format} workbook (${build.name}) stored by SPEC-119's reader re-reads its clean sheets and gains its calculation settings`, async () => {
      const stored = await spec119Record(build, format);
      expect(stored.calc).toBeUndefined();
      const upgraded = await upgradeSpreadsheetModel(stored);
      expect(upgraded).not.toBe(stored);
      expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
      expect(upgraded.calc).toEqual(calc);
      expect(upgraded.id).toBe(stored.id);
      expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(stored.sheets.map((sheet) => sheet.id));
    });
  }

  it("a fresh read of spec120-calc.xlsx has no calc (its calcPr sets nothing modelled) and no forceFullCalcOnLoad", async () => {
    const model = await readSpreadsheetBytes(await buildSpec120CalcXlsxBytes(), "xlsx", "spec120-calc");
    expect(model.calc).toBeUndefined();
    expect(model.forceFullCalcOnLoad).toBeUndefined();
  });

  it("keeps a stored model's forceFullCalcOnLoad through the re-read — this product never clears it", async () => {
    const stored: SpreadsheetModel = { ...(await spec119Record(buildSpec120CalcXlsxBytes, "xlsx")), forceFullCalcOnLoad: true };
    expect((await upgradeSpreadsheetModel(stored)).forceFullCalcOnLoad).toBe(true);
  });
});
