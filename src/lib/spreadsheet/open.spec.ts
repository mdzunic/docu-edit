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
// the re-read style table. SPEC-128 moves it to 9: a version-8 record's
// dirty sheet gains its autofilter, sort records, sort/autofilter permits
// and filter-hidden rows from its retained source; a clean one is re-read.
import { afterEach, describe, expect, it } from "vitest";
import { decodedPartBytes } from "../core/packageParts";
import { unzip } from "../core/zip";
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
  buildSpec123StylesXlsxBytes,
  SPEC123_XF,
} from "./fixtures/build";
import type { Cell, Sheet, SpreadsheetModel } from "./model";
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
  it("SPEC-117 moved the reader to version 2, SPEC-118 to 3, SPEC-119 to 4, SPEC-120 to 5, SPEC-122 to 6, SPEC-123 to 7, SPEC-124 to 8 and SPEC-128 to 9, with one contiguous step each (from 1, 2, 3, 4, 5, 6, 7, 8)", () => {
    expect(SPREADSHEET_READER_VERSION).toBe(9);
    expect(READER_UPGRADES.map((step) => step.from)).toEqual([1, 2, 3, 4, 5, 6, 7, 8]);
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

// --- SPEC-122 (5 → 6): preserved cell content and protection on dirty sheets ----------------

const encodeText = (text: string): Uint8Array => new TextEncoder().encode(text);
const decodeText = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes);

/** `parts` with `name`'s text passed through `edit` (each replacement must
 * apply, so a fixture change cannot silently void the test). */
function patchPart(parts: Map<string, Uint8Array>, name: string, edits: [string, string][]): void {
  let text = decodeText(parts.get(name));
  for (const [from, to] of edits) {
    expect(text.includes(from), `${name}: ${from}`).toBe(true);
    text = text.replace(from, to);
  }
  parts.set(name, encodeText(text));
}

/** An ODS annotation (a note) as LibreOffice writes one inside a cell. */
const annotation = (text: string): string => `<office:annotation><dc:creator>Fixture Author</dc:creator><text:p>${text}</text:p></office:annotation>`;

const SPEC122_CELL_FIELDS = ["sstIndex", "preservedAttrs", "preservedXml", "preservedText"] as const;

/** A record as SPEC-120's reader (version 5) stored `read`: none of
 * SPEC-122's fields, no cell holding only them, then `edit`ed — the sheets it
 * names marked dirty, as SPEC-121's commands leave them. */
function spec120Record(read: SpreadsheetModel, dirty: number[], edit: (model: SpreadsheetModel) => void): SpreadsheetModel {
  const model = JSON.parse(JSON.stringify(read)) as SpreadsheetModel;
  delete model.structureProtected;
  model.readerVersion = 5;
  model.sheets = model.sheets.map(({ protection: _protection, ...sheet }, index) => ({
    ...sheet,
    id: `stored-sheet-${index}`,
    rows: sheet.rows.map((row) => ({
      ...row,
      cells: row.cells
        .map((cell) => {
          const plain: Cell = { ...cell };
          for (const field of SPEC122_CELL_FIELDS) {
            delete plain[field];
          }
          return plain;
        })
        .filter((cell) => cell.value !== null || cell.formula !== undefined || cell.styleId !== undefined),
    })),
  }));
  for (const index of dirty) {
    model.sheets[index] = { ...model.sheets[index]!, dirty: true };
  }
  edit(model);
  return model;
}

/** Sets a stored cell's value in place (the record is a private clone). */
function setStoredValue(model: SpreadsheetModel, sheet: number, a1: [number, number], value: string | number): void {
  const cell = getCell(model.sheets[sheet]!, a1[0], a1[1])!;
  cell.value = value;
}

describe("SPEC-122's reader step: a dirty sheet regains its preserved content and protection", () => {
  it("SPEC-122 moved the reader to version 6 with its step from 5 — the first that does real work on a dirty sheet", () => {
    expect(SPREADSHEET_READER_VERSION).toBeGreaterThanOrEqual(6);
    expect(READER_UPGRADES.findIndex((step) => step.from === 5)).toBe(4);
  });

  it("xlsx: sstIndex and preserved attributes and children come back onto unchanged cells only; protection and structure protection come back", async () => {
    const parts = await unzip(await buildMiniBookXlsxBytes());
    patchPart(parts, "xl/workbook.xml", [[`<workbookPr defaultThemeVersion="202300"/>`, `<workbookPr defaultThemeVersion="202300"/><workbookProtection lockStructure="1"/>`]]);
    patchPart(parts, "xl/worksheets/sheet1.xml", [
      [`<c r="A3" t="s"><v>6</v></c>`, `<c r="A3" t="s" ph="1"><v>6</v></c>`],
      [`<c r="B3"><v>1200</v></c>`, `<c r="B3" vm="1"><v>1200</v></c>`],
      [`<c r="C3"><v>800</v></c>`, `<c r="C3"><v>800</v><extLst><ext uri="{C3}"/></extLst></c>`],
      [`<c r="B4"><v>950.5</v></c>`, `<c r="B4"><v>950.5</v><extLst><ext uri="{B4}"/></extLst></c>`],
      [`</sheetData>`, `</sheetData><sheetProtection sheet="1" objects="1" scenarios="1"/>`],
    ]);
    const read = await readSpreadsheetBytes(new Uint8Array(0), "xlsx", "SPEC-122 shapes", parts);
    // The current reader captures the Model additions fields.
    const summary = read.sheets[0]!;
    expect(read.structureProtected).toBe(true);
    expect(summary.protection).toEqual({ sheet: true });
    expect(getCell(summary, 2, 0)).toMatchObject({ value: "North", sstIndex: 6, preservedAttrs: ' ph="1"' });
    expect(getCell(summary, 1, 1)).toMatchObject({ value: "Revenue", sstIndex: 2 });
    expect(getCell(summary, 2, 1)).toMatchObject({ value: 1200, preservedAttrs: ' vm="1"' });
    expect(getCell(summary, 2, 2)).toMatchObject({ value: 800, preservedXml: `<extLst><ext uri="{C3}"/></extLst>` });

    // Stored by SPEC-120's reader; then B2 and C3 were edited on the (dirty) Summary.
    const stored = spec120Record(read, [0], (model) => {
      setStoredValue(model, 0, [1, 1], "Turnover");
      setStoredValue(model, 0, [2, 2], 801);
    });
    expect(stored.structureProtected).toBeUndefined();
    expect(getCell(stored.sheets[0]!, 2, 0)!.sstIndex).toBeUndefined();

    const upgraded = await upgradeSpreadsheetModel(stored);
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.structureProtected).toBe(true);
    const sheet = upgraded.sheets[0]!;
    expect(sheet.id).toBe("stored-sheet-0");
    expect(sheet.protection).toEqual({ sheet: true });
    expect(getCell(sheet, 2, 0)).toMatchObject({ value: "North", sstIndex: 6, preservedAttrs: ' ph="1"' });
    expect(getCell(sheet, 2, 1)).toMatchObject({ value: 1200, preservedAttrs: ' vm="1"' });
    expect(getCell(sheet, 3, 1)).toMatchObject({ value: 950.5, preservedXml: `<extLst><ext uri="{B4}"/></extLst>` });
    // The edited cells keep their edits and regain nothing.
    expect(getCell(sheet, 1, 1)).toEqual({ col: 1, value: "Turnover", styleId: getCell(sheet, 1, 1)!.styleId });
    expect(getCell(sheet, 1, 1)!.sstIndex).toBeUndefined();
    expect(getCell(sheet, 2, 2)).toEqual({ col: 2, value: 801 });
    // A clean sheet is re-read and keeps its stored id.
    expect(upgraded.sheets[1]!.id).toBe("stored-sheet-1");
    expect(upgraded.sheets[1]!.protection).toBeUndefined();
  });

  it("xlsx: a default-width grouped column span and grouped rows without cells — stored by version 6, not 5 — come back onto a dirty sheet, so its outline is written", async () => {
    const parts = await unzip(await buildMiniBookXlsxBytes());
    patchPart(parts, "xl/worksheets/sheet1.xml", [
      [
        `<cols><col min="1" max="1" width="14.6328125" customWidth="1"/></cols>`,
        `<cols><col min="1" max="1" width="14.6328125" customWidth="1"/><col min="7" max="9" outlineLevel="1"/><col min="10" max="10" collapsed="1"/></cols>`,
      ],
      [`</sheetData>`, `<row r="13" outlineLevel="1"/><row r="14" spans="1:5" outlineLevel="1" x14ac:dyDescent="0.35"/><row r="15" collapsed="1"/></sheetData>`],
    ]);
    const read = await readSpreadsheetBytes(new Uint8Array(0), "xlsx", "SPEC-122 outline", parts);
    const summary = read.sheets[0]!;
    expect(summary.columns.slice(1)).toEqual([
      { min: 6, max: 8, sourceAttrs: `outlineLevel="1"` },
      { min: 9, max: 9, sourceAttrs: `collapsed="1"` },
    ]);
    expect(summary.rows.slice(-3)).toEqual([
      { index: 12, cells: [], sourceAttrs: `outlineLevel="1"` },
      { index: 13, cells: [], sourceAttrs: `spans="1:5" outlineLevel="1" x14ac:dyDescent="0.35"` },
      { index: 14, cells: [], sourceAttrs: `collapsed="1"` },
    ]);

    // Stored by SPEC-120's reader, which kept neither; then C3 was edited on the (dirty) Summary.
    const stored = spec120Record(read, [0], (model) => {
      const sheet = model.sheets[0]!;
      model.sheets[0] = { ...sheet, columns: sheet.columns.slice(0, 1), rows: sheet.rows.filter((row) => row.cells.length > 0) };
      setStoredValue(model, 0, [2, 2], 801);
    });
    const upgraded = await upgradeSpreadsheetModel(stored);
    const sheet = upgraded.sheets[0]!;
    expect(sheet.columns).toEqual(summary.columns);
    expect(sheet.rows.map((row) => row.index)).toEqual(summary.rows.map((row) => row.index));
    expect(sheet.rows.slice(-3)).toEqual(summary.rows.slice(-3));
    expect(getCell(sheet, 2, 2)).toEqual({ col: 2, value: 801 });
  });

  it("ods: preserved attributes, child content and text:p come back onto unchanged cells only; protection and structure protection come back", async () => {
    const parts = await unzip(await buildMiniBookOdsBytes());
    patchPart(parts, "content.xml", [
      [`<office:spreadsheet>`, `<office:spreadsheet table:structure-protected="true">`],
      [`<table:table table:name="Inventory" table:style-name="ta1">`, `<table:table table:name="Inventory" table:style-name="ta1" table:protected="true">`],
      [
        `<text:p>Widget</text:p>`,
        `<text:p><text:a xlink:href="https://example.com/" xlink:type="simple">Widget</text:a></text:p>`,
      ],
      [
        `<table:table-cell office:value-type="float" office:value="4" calcext:value-type="float">`,
        `<table:table-cell office:value-type="float" office:value="4" calcext:value-type="float" table:content-validation-name="val1">`,
      ],
      [`<text:p>Gadget</text:p>`, `${annotation("Edited away")}<text:p>Gadget</text:p>`],
      [`<text:p>Done</text:p>`, `${annotation("Kept")}<text:p>Done</text:p>`],
    ]);
    const read = await readSpreadsheetBytes(new Uint8Array(0), "ods", "SPEC-122 shapes", parts);
    const inventory = read.sheets[1]!;
    expect(read.structureProtected).toBe(true);
    expect(inventory.protection).toEqual({ sheet: true });
    const link = `<text:p><text:a xlink:href="https://example.com/" xlink:type="simple">Widget</text:a></text:p>`;
    expect(getCell(inventory, 1, 0)).toMatchObject({ value: "Widget", preservedText: link });
    expect(getCell(inventory, 1, 1)).toMatchObject({ value: 4, preservedAttrs: ' table:content-validation-name="val1"' });
    expect(getCell(inventory, 2, 0)).toMatchObject({ value: "Gadget", preservedXml: annotation("Edited away") });
    expect(getCell(inventory, 22, 0)).toMatchObject({ value: "Done", preservedXml: annotation("Kept") });

    // Stored by SPEC-120's reader; then A3 was edited on the (dirty) Inventory.
    const stored = spec120Record(read, [1], (model) => setStoredValue(model, 1, [2, 0], "Gizmo"));
    const upgraded = await upgradeSpreadsheetModel(stored);
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.structureProtected).toBe(true);
    const sheet = upgraded.sheets[1]!;
    expect(sheet.id).toBe("stored-sheet-1");
    expect(sheet.protection).toEqual({ sheet: true });
    expect(getCell(sheet, 1, 0)).toMatchObject({ value: "Widget", preservedText: link });
    expect(getCell(sheet, 1, 1)).toMatchObject({ value: 4, preservedAttrs: ' table:content-validation-name="val1"' });
    expect(getCell(sheet, 22, 0)).toMatchObject({ value: "Done", preservedXml: annotation("Kept") });
    // The edited cell keeps its edit and regains nothing.
    expect(getCell(sheet, 2, 0)!.value).toBe("Gizmo");
    expect(getCell(sheet, 2, 0)!.preservedXml).toBeUndefined();
    // The other (clean) tables are re-read, unprotected.
    expect(upgraded.sheets[0]!.protection).toBeUndefined();
    expect(upgraded.sheets[2]!.id).toBe("stored-sheet-2");
  });
});

describe("SPEC-123's reader step: named cell styles and the protection's format permits", () => {
  it("moved the reader to version 7 with its step from 6 (SPEC-124's from 7 follows it)", () => {
    const at = READER_UPGRADES.findIndex((step) => step.from === 6);
    expect(at).toBeGreaterThanOrEqual(0);
    expect(READER_UPGRADES[at + 1]!.from).toBe(7);
  });

  it("a dirty sheet's record regains namedStyles, keeps its source styles' ids, and the step is idempotent", async () => {
    const read = await readSpreadsheetBytes(await buildSpec123StylesXlsxBytes(), "xlsx", "spec123");
    expect(read.namedStyles?.map((style) => style.name)).toEqual(["Followed Hyperlink", "Good", "Heading 1", "Invoice total", "Normal"]);
    expect(read.styles[SPEC123_XF.good]!.namedStyle).toBe("Good");
    // Stored by SPEC-122's reader (version 6): no named styles anywhere; then
    // A2 was edited on the (dirty) Budget sheet.
    const stored = JSON.parse(JSON.stringify(read)) as SpreadsheetModel;
    stored.readerVersion = 6;
    delete stored.namedStyles;
    stored.styles = stored.styles.map(({ namedStyle: _named, ...style }) => style);
    stored.sheets = stored.sheets.map((sheet, index) => ({ ...sheet, id: `stored-sheet-${index}` }));
    stored.sheets[0] = { ...stored.sheets[0]!, dirty: true };
    getCell(stored.sheets[0]!, 1, 0)!.value = "Rent and rates";
    const upgraded = await upgradeSpreadsheetModel(stored);
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.namedStyles).toEqual(read.namedStyles);
    const sheet = upgraded.sheets[0]!;
    expect(getCell(sheet, 1, 0)!.value).toBe("Rent and rates");
    // F3 still names its source entry, now known to be "Good".
    expect(getCell(sheet, 2, 5)!.styleId).toBe(SPEC123_XF.good);
    expect(upgraded.styles[SPEC123_XF.good]!.namedStyle).toBe("Good");
    const named = upgraded.namedStyles;
    await READER_UPGRADES.find((step) => step.from === 6)!.upgradeDirtySheet(upgraded, sheet, decodedPartBytes(upgraded.opaque));
    expect(upgraded.namedStyles).toBe(named);
  });
});

// --- SPEC-128's step (8 → 9) ------------------------------------------------------------------

/** The rows' indices carrying a visibility flag. */
const flagged = (sheet: Sheet, flag: "hidden" | "filterHidden"): number[] => sheet.rows.filter((row) => row[flag]).map((row) => row.index);

/** The start of mini-book.ods's Inventory row for `item`, and the patch
 * giving it a `table:visibility`. */
const odsItemRow = (item: string): string => `<table:table-row table:style-name="ro1"><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>${item}</text:p>`;
const hide = (item: string, visibility: string): [string, string] => [odsItemRow(item), odsItemRow(item).replace(`"ro1"`, `"ro1" table:visibility="${visibility}"`)];

describe("SPEC-128's reader step: the autofilter, the sort records and filter-hidden rows", () => {
  it("moved the reader to version 9 with its step from 8, last", () => {
    expect(SPREADSHEET_READER_VERSION).toBe(9);
    expect(READER_UPGRADES.at(-1)!.from).toBe(8);
  });

  it("a version-8 record re-reads its clean filtered sheet, and its dirty one gains the autofilter, the sort records and its filter-hidden rows from the retained source", async () => {
    const parts = await unzip(await buildMiniBookXlsxBytes());
    // Summary (clean): A2:E5 filtered to North — South's row 4 hidden by it,
    // Total's row 5 rejected but shown (reading never changes what shows).
    patchPart(parts, "xl/worksheets/sheet1.xml", [
      [`<row r="4" spans="1:5" x14ac:dyDescent="0.35">`, `<row r="4" spans="1:5" hidden="1" x14ac:dyDescent="0.35">`],
      [`</sheetData>`, `</sheetData><autoFilter ref="A2:E5"><filterColumn colId="0"><filters><filter val="North"/></filters></filterColumn></autoFilter>`],
    ]);
    // Data (to be edited): A1:C500 filtered to odd parity and last sorted by
    // Amount descending — rows 3 and 5 (even) hidden by it, row 4 (odd) by
    // the user — plus a worksheet sort record and the sort/autofilter permits.
    patchPart(parts, "xl/worksheets/sheet2.xml", [
      [`<row><c><v>2</v></c><c><v>3</v></c>`, `<row hidden="1"><c><v>2</v></c><c><v>3</v></c>`],
      [`<row><c><v>3</v></c><c><v>4.5</v></c>`, `<row hidden="1"><c><v>3</v></c><c><v>4.5</v></c>`],
      [`<row><c><v>4</v></c><c><v>6</v></c>`, `<row hidden="1"><c><v>4</v></c><c><v>6</v></c>`],
      [
        `</sheetData>`,
        `</sheetData><sheetProtection sheet="1" sort="0" autoFilter="0"/>` +
          `<autoFilter ref="A1:C500"><filterColumn colId="2"><filters><filter val="odd"/></filters></filterColumn>` +
          `<sortState ref="A2:C500"><sortCondition descending="1" ref="B2:B500"/></sortState></autoFilter>` +
          `<sortState ref="E2:F9"><sortCondition ref="F2:F9"/></sortState>`,
      ],
    ]);
    const read = await readSpreadsheetBytes(new Uint8Array(0), "xlsx", "SPEC-128 filters", parts);
    // The current reader models both filters and classifies the hidden rows.
    const [summary, data] = read.sheets;
    expect(summary!.autoFilter?.range).toEqual({ top: 1, left: 0, bottom: 4, right: 4 });
    expect([flagged(summary!, "filterHidden"), flagged(summary!, "hidden")]).toEqual([[3], []]);
    expect(data!.autoFilter?.sort?.keys).toEqual([{ col: 1, descending: true }]);
    expect(data!.sortState?.keys).toEqual([{ col: 5, descending: false }]);
    expect(data!.protection).toMatchObject({ sheet: true, allowSort: true, allowAutoFilter: true });
    expect([flagged(data!, "filterHidden"), flagged(data!, "hidden")]).toEqual([[2, 4], [3]]);

    // Stored by SPEC-124's reader (version 8): no autofilter or sort record
    // anywhere, every hidden row the user's, no sort or autofilter permit;
    // then A2 was edited on the (dirty) Data sheet.
    const stored = JSON.parse(JSON.stringify(read)) as SpreadsheetModel;
    stored.readerVersion = 8;
    stored.sheets = stored.sheets.map(({ autoFilter: _filter, sortState: _sort, protection, ...sheet }, index) => {
      const old: Sheet = {
        ...sheet,
        id: `stored-sheet-${index}`,
        rows: sheet.rows.map(({ filterHidden, ...row }) => (filterHidden ? { ...row, hidden: true as const } : row)),
      };
      if (protection) {
        const { allowSort: _s, allowAutoFilter: _a, ...rest } = protection;
        old.protection = rest;
      }
      return old;
    });
    stored.sheets[1] = { ...stored.sheets[1]!, dirty: true };
    getCell(stored.sheets[1]!, 1, 0)!.value = 100;
    expect(flagged(stored.sheets[1]!, "hidden")).toEqual([2, 3, 4]);

    const upgraded = await upgradeSpreadsheetModel(stored);
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    // The dirty sheet keeps its edit and gains what the source says.
    const dirty = upgraded.sheets[1]!;
    expect(dirty.id).toBe("stored-sheet-1");
    expect(getCell(dirty, 1, 0)!.value).toBe(100);
    expect(dirty.autoFilter).toEqual(data!.autoFilter);
    expect(dirty.sortState).toEqual(data!.sortState);
    expect(dirty.protection).toEqual(data!.protection);
    expect([flagged(dirty, "filterHidden"), flagged(dirty, "hidden")]).toEqual([[2, 4], [3]]);
    // The clean sheet is re-read, keeping its stored id.
    const clean = upgraded.sheets[0]!;
    expect(clean.id).toBe("stored-sheet-0");
    expect(clean.autoFilter).toEqual(summary!.autoFilter);
    expect([flagged(clean, "filterHidden"), flagged(clean, "hidden")]).toEqual([[3], []]);
  });

  it("ods: a dirty sheet gains its database range's autofilter, and its hidden rows inside it are reclassified by the read rule", async () => {
    const parts = await unzip(await buildMiniBookOdsBytes());
    // Inventory A1:D6 filtered to Widget: Gadget hidden by it, Widget by the
    // user, the three Spare rows rejected but shown.
    patchPart(parts, "content.xml", [
      hide("Widget", "collapse"),
      hide("Gadget", "filter"),
      [
        `</office:spreadsheet>`,
        `<table:database-ranges><table:database-range table:name="__Anonymous_Sheet_DB__1" table:target-range-address="Inventory.A1:Inventory.D6" table:display-filter-buttons="true">` +
          `<table:filter><table:filter-condition table:field-number="0" table:value="Widget" table:operator="="/></table:filter></table:database-range></table:database-ranges></office:spreadsheet>`,
      ],
    ]);
    const read = await readSpreadsheetBytes(new Uint8Array(0), "ods", "SPEC-128 ods filter", parts);
    const inventory = read.sheets[1]!;
    expect(inventory.autoFilter?.range).toEqual({ top: 0, left: 0, bottom: 5, right: 3 });
    expect([flagged(inventory, "filterHidden"), flagged(inventory, "hidden")]).toEqual([[2], [1]]);

    // Stored by SPEC-124's reader: both `table:visibility` values read as
    // the user's `hidden`, no autofilter; then a cell was edited.
    const stored = JSON.parse(JSON.stringify(read)) as SpreadsheetModel;
    stored.readerVersion = 8;
    stored.sheets = stored.sheets.map(({ autoFilter: _filter, ...sheet }, index) => ({
      ...sheet,
      id: `stored-sheet-${index}`,
      rows: sheet.rows.map(({ filterHidden, ...each }) => (filterHidden ? { ...each, hidden: true as const } : each)),
    }));
    stored.sheets[1] = { ...stored.sheets[1]!, dirty: true };
    getCell(stored.sheets[1]!, 3, 1)!.value = 2;
    const upgraded = await upgradeSpreadsheetModel(stored);
    const dirty = upgraded.sheets[1]!;
    expect(dirty.autoFilter).toEqual(inventory.autoFilter);
    expect([flagged(dirty, "filterHidden"), flagged(dirty, "hidden")]).toEqual([[2], [1]]);
    expect(getCell(dirty, 3, 1)!.value).toBe(2);
  });

  it("the step leaves a dirty sheet whose source has no autofilter as it is", async () => {
    const read = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    const sheet = { ...read.sheets[0]!, rows: [{ index: 2, cells: [], hidden: true as const }], dirty: true as const };
    await READER_UPGRADES.at(-1)!.upgradeDirtySheet(read, sheet, decodedPartBytes(read.opaque));
    expect(sheet.autoFilter).toBeUndefined();
    expect(sheet.rows).toEqual([{ index: 2, cells: [], hidden: true }]);
  });
});
