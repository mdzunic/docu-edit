// SPEC-116 "The workbook model" (with SPEC-117's additions): the grid and
// workbook limits; the model is plain JSON — a hand-built model with every
// optional field (SPEC-117's styles, theme, layout and view fields among
// them), and every model
// the readers build from the fixtures (the Numbers export included), survive
// JSON.parse(JSON.stringify(…)) and hold no Map, Set or class instance; the
// invariants hold on read models (rows strictly ascending by index, cells
// strictly ascending by column, no stored null cell, no stored empty row,
// unique sheet ids, every sheet's source retained); the error helpers; and
// model.ts's header states the four invariants as normative.
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { formatA1 } from "./address";
import {
  buildMiniBookOdsBytes,
  buildMiniBookXlsxBytes,
  buildSpec1904XlsxBytes,
  buildSpecMacrosXlsmBytes,
  buildSpecTemplateXltxBytes,
  buildSpecWideOdsBytes,
  buildSpec117StylesOdsBytes,
  buildSpec117StylesXlsxBytes,
} from "./fixtures/build";
import {
  isCellError,
  MAX_COLS,
  MAX_ROWS,
  MAX_WORKBOOK_CELLS,
  toCellError,
  type CellValue,
  type ErrorCode,
  type SheetColorRef,
  type SpreadsheetFormat,
  type SpreadsheetModel,
} from "./model";
import { readSpreadsheetBytes } from "./open";

const spreadsheetDir = dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = resolve(spreadsheetDir, "..", "..", "..", "e2e", "fixtures");

const ERROR_CODES: ErrorCode[] = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
  "#FIELD!",
  "#BLOCKED!",
  "#CONNECT!",
  "#BUSY!",
  "#UNKNOWN!",
];

/** A theme colour as the xlsx reader keeps it: the transform plus the
 * source `tint` text (SPEC-117's SheetColorRef). */
const ACCENT1_DARKER_25: SheetColorRef = { theme: "accent1", lumMod: 0.75, xlsxTint: "-0.249977111117893" };

/** A model with every optional field populated and every kind of value. */
function handBuiltModel(): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "doc-1",
    title: "Quarterly budget",
    format: "ods",
    packageVariant: "template",
    readerVersion: 1,
    sheets: [
      {
        id: "sheet-a",
        name: "Summary",
        visibility: "visible",
        kind: "grid",
        rows: [
          {
            index: 0,
            cells: [
              { col: 0, value: "Region" },
              { col: 1, value: "Revenue" },
            ],
          },
          {
            index: 1,
            cells: [
              { col: 0, value: "North" },
              { col: 1, value: 1234.5 },
              { col: 2, value: true },
              { col: 3, value: { error: "#DIV/0!" } },
              { col: 4, value: 46_303.75 },
              {
                col: 5,
                value: "Bold and plain",
                runs: [{ text: "Bold", font: { bold: true, color: ACCENT1_DARKER_25 } }, { text: " and plain" }],
                styleId: 1,
              },
            ],
          },
          // SPEC-117: a styled empty row, a style-only cell, row properties.
          { index: 2, cells: [], heightPt: 30, customHeight: true, sourceAttrs: 'spans="1:6" x14ac:dyDescent="0.25"' },
          { index: 3, cells: [], hidden: true, styleId: 2, sourceStyleName: "ro2" },
          { index: 4, cells: [{ col: 1, value: false }, { col: 6, value: null, styleId: 1 }] },
          { index: MAX_ROWS - 1, cells: [{ col: MAX_COLS - 1, value: -0.5 }] },
        ],
        sourceXml: '\n   <table:table table:name="Summary" table:style-name="ta1">…</table:table>',
        columns: [
          { min: 0, max: 0, widthPx: 120, sourceWidthChars: 16.42578125, sourceAttrs: 'bestFit="1" customWidth="1"' },
          { min: 3, max: 5, styleId: 1, sourceStyleName: "co2" },
          { min: 6, max: 6, hidden: true },
        ],
        defaultColWidthPx: 85.33,
        defaultRowHeightPt: 12.8,
        merges: [{ top: 0, left: 2, bottom: 0, right: 4 }],
        frozen: { rows: 1, cols: 1 },
        showGridLines: false,
        showZeros: false,
        zoomPercent: 85,
        tabColor: { theme: "accent2", lumMod: 0.6, lumOff: 0.4 },
        selection: {
          active: { row: 1, col: 1 },
          ranges: [
            { top: 1, left: 1, bottom: 4, right: 1 },
            { top: 6, left: 3, bottom: 7, right: 4 },
          ],
          activeRangeIndex: 0,
        },
        topLeftCell: { row: 1, col: 1 },
      },
      {
        id: "sheet-b",
        name: "Chart1",
        visibility: "hidden",
        kind: "foreign",
        rows: [],
        partName: "xl/chartsheets/sheet1.xml",
        columns: [],
        defaultColWidthPx: 64,
        defaultRowHeightPt: 15,
        merges: [],
      },
      {
        id: "sheet-c",
        name: "Secret",
        visibility: "veryHidden",
        kind: "grid",
        rows: [{ index: 3, cells: [{ col: 2, value: { error: "#N/A" } }] }],
        sourceXml: '<table:table table:name="Secret">…</table:table>',
        columns: [],
        defaultColWidthPx: 85.33,
        defaultRowHeightPt: 12.8,
        merges: [],
      },
    ],
    activeSheetIndex: 0,
    dateSystem: "1900",
    odsNullDate: "1900-01-01",
    odsContentEnvelope: { prefix: '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content>…<office:spreadsheet>', suffix: "</office:spreadsheet>…" },
    opaque: [
      { name: "mimetype", data: "YXBwbGljYXRpb24vdm5kLm9hc2lzLm9wZW5kb2N1bWVudC5zcHJlYWRzaGVldA==" },
      { name: "content.xml", data: "PD94bWwgdmVyc2lvbj0iMS4wIj8+" },
    ],
    preview: {
      sheetName: "Summary",
      cells: [
        ["Region", "Revenue", "", "", "", ""],
        ["North", "1234.5", "TRUE", "#DIV/0!", "46303.75", "Bold and plain"],
      ],
    },
    properties: {
      title: "Budget",
      author: "Fixture Author",
      created: "2026-10-08T13:45:00Z",
      customProperties: [{ name: "Department", type: "text", value: "Finance" }],
      extraFields: [{ name: "generator", xml: "<meta:generator>LibreOffice/24.8</meta:generator>", value: "LibreOffice/24.8" }],
    },
    media: [],
    styles: [
      {},
      {
        font: { family: { theme: "minor" }, sizePt: 11, bold: true, italic: true, underline: "double", strike: true, color: "#ff0000", vertAlign: "superscript" },
        fill: { kind: "solid", color: "#ffff00" },
        borders: {
          top: { style: "thin" },
          right: { style: "mediumDashDot", color: { theme: "dk1" } },
          bottom: { style: "double", color: "#0000ff" },
          left: { style: "hair" },
          diagonal: { style: "dashed", up: true, down: true },
        },
        alignment: { horizontal: "center", vertical: "top", wrap: true, indent: 2, textRotation: 255, shrinkToFit: true },
        protection: { locked: false, hidden: true },
        quotePrefix: true,
        sourceXf: 1,
      },
      { fill: { kind: "pattern", pattern: "darkTrellis", fg: { theme: "accent2", tint: 0.5 }, bg: "#ffffff" }, sourceStyleName: "ce2" },
      {
        fill: {
          kind: "gradient",
          type: "path",
          left: 0.5,
          right: 0.5,
          top: 0.5,
          bottom: 0.5,
          stops: [
            { position: 0, color: "#ffffff" },
            { position: 1, color: { theme: "accent1" } },
          ],
        },
      },
    ],
    defaultStyleId: 0,
    theme: {
      name: "Office Theme",
      colors: {
        dk1: "#000000",
        lt1: "#ffffff",
        dk2: "#44546a",
        lt2: "#e7e6e6",
        accent1: "#4472c4",
        accent2: "#ed7d31",
        accent3: "#a5a5a5",
        accent4: "#ffc000",
        accent5: "#5b9bd5",
        accent6: "#70ad47",
        hlink: "#0563c1",
        folHlink: "#954f72",
      },
      fonts: { major: "Calibri Light", minor: "Calibri" },
    },
    indexedColors: Array.from({ length: 64 }, (_, i) => `#${i.toString(16).padStart(2, "0").repeat(3)}`),
  };
}

/** Every way `model` breaks the sparse invariants, as readable messages. */
function invariantViolations(model: SpreadsheetModel): string[] {
  const problems: string[] = [];
  const ids = new Set<string>();
  for (const sheet of model.sheets) {
    if (sheet.id.length === 0 || ids.has(sheet.id)) {
      problems.push(`${sheet.name}: duplicate or empty sheet id "${sheet.id}"`);
    }
    ids.add(sheet.id);
    let previousRow = -1;
    for (const row of sheet.rows) {
      const where = `${sheet.name}!row ${row.index}`;
      if (!Number.isInteger(row.index) || row.index <= previousRow || row.index >= MAX_ROWS) {
        problems.push(`${where}: not strictly ascending inside the grid (after row ${previousRow})`);
      }
      previousRow = row.index;
      // SPEC-117: a row with no cells is stored only for its properties.
      if (row.cells.length === 0 && row.heightPt === undefined && !row.hidden && row.styleId === undefined) {
        problems.push(`${where}: stored with no cells`);
      }
      let previousCol = -1;
      for (const cell of row.cells) {
        const at = `${sheet.name}!${Number.isInteger(cell.col) && cell.col >= 0 ? formatA1({ row: row.index, col: cell.col }) : `col ${cell.col}`}`;
        if (!Number.isInteger(cell.col) || cell.col <= previousCol || cell.col >= MAX_COLS) {
          problems.push(`${at}: column not strictly ascending inside the grid (after col ${previousCol})`);
        }
        previousCol = cell.col;
        // SPEC-117: a style-only cell (value null, styleId set) is stored.
        if (cell.value === null && cell.styleId === undefined) {
          problems.push(`${at}: stored with value null`);
        }
        if (cell.runs && cell.value !== cell.runs.map((run) => run.text).join("")) {
          problems.push(`${at}: value is not its runs' concatenated text`);
        }
      }
    }
  }
  return problems;
}

/** Paths in `value` holding anything JSON cannot carry as-is: a Map, Set,
 * typed array, Date or other class instance, a function, a symbol, a
 * bigint, or a hole/undefined in an array. */
function nonPlainPaths(value: unknown, path: string): string[] {
  if (value === null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return [];
  }
  if (typeof value !== "object") {
    return [`${path} (${typeof value})`];
  }
  if (Array.isArray(value)) {
    return Array.from(value, (item: unknown, i) => (item === undefined ? [`${path}[${i}] (undefined)`] : nonPlainPaths(item, `${path}[${i}]`))).flat();
  }
  if (Object.getPrototypeOf(value) !== Object.prototype) {
    return [`${path} (${Object.prototype.toString.call(value)})`];
  }
  return Object.entries(value).flatMap(([key, item]: [string, unknown]) => (item === undefined ? [] : nonPlainPaths(item, `${path}.${key}`)));
}

interface ReadFixture {
  name: string;
  format: SpreadsheetFormat;
  title: string;
  bytes: () => Promise<Uint8Array>;
}

/** The numbers fixture: exported by Apple Numbers 14.4 and committed, not
 * built — a real producer's shapes. */
async function numbersExportBytes(): Promise<Uint8Array> {
  return new Uint8Array(readFileSync(resolve(e2eFixturesDir, "spec116-numbers.xlsx")));
}

const READ_FIXTURES: ReadFixture[] = [
  { name: "mini-book.xlsx", format: "xlsx", title: "Mini book", bytes: buildMiniBookXlsxBytes },
  { name: "mini-book.ods", format: "ods", title: "Mini book", bytes: buildMiniBookOdsBytes },
  { name: "spec116-1904.xlsx", format: "xlsx", title: "spec116-1904", bytes: buildSpec1904XlsxBytes },
  { name: "spec116-macros.xlsm", format: "xlsx", title: "spec116-macros", bytes: buildSpecMacrosXlsmBytes },
  { name: "spec116-template.xltx", format: "xlsx", title: "spec116-template", bytes: buildSpecTemplateXltxBytes },
  { name: "spec116-wide.ods", format: "ods", title: "spec116-wide", bytes: buildSpecWideOdsBytes },
  { name: "spec116-numbers.xlsx", format: "xlsx", title: "spec116-numbers", bytes: numbersExportBytes },
  { name: "spec117-styles.xlsx", format: "xlsx", title: "spec117-styles", bytes: buildSpec117StylesXlsxBytes },
  { name: "spec117-styles.ods", format: "ods", title: "spec117-styles", bytes: buildSpec117StylesOdsBytes },
];

const readModels = new Map<string, Promise<SpreadsheetModel>>();

/** The model readSpreadsheetBytes builds from a fixture — read once per
 * file and never mutated by these tests. */
function readFixture(fixture: ReadFixture): Promise<SpreadsheetModel> {
  let model = readModels.get(fixture.name);
  if (!model) {
    model = fixture.bytes().then((bytes) => readSpreadsheetBytes(bytes, fixture.format, fixture.title));
    readModels.set(fixture.name, model);
  }
  return model;
}

describe("the limits", () => {
  it("pins the grid at 1,048,576 rows × 16,384 columns and the workbook at 2,000,000 stored cells", () => {
    expect(MAX_ROWS).toBe(1_048_576);
    expect(MAX_COLS).toBe(16_384);
    expect(MAX_WORKBOOK_CELLS).toBe(2_000_000);
  });
});

describe("the model is plain JSON", () => {
  it("a hand-built model with every optional field survives JSON.parse(JSON.stringify(…))", () => {
    const model = handBuiltModel();
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
    expect(nonPlainPaths(model, "model")).toEqual([]);
  });

  it.each(READ_FIXTURES)("the model read from $name survives JSON.parse(JSON.stringify(…))", async (fixture) => {
    const model = await readFixture(fixture);
    expect(model.kind).toBe("spreadsheet");
    expect(model.title).toBe(fixture.title);
    expect(model.format).toBe(fixture.format);
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
  });

  it.each(READ_FIXTURES)("the model read from $name holds no Map, Set or class instance", async (fixture) => {
    expect(nonPlainPaths(await readFixture(fixture), "model")).toEqual([]);
  });

  it("the plain-data check itself flags what JSON cannot carry", () => {
    const tainted = { ...handBuiltModel(), extras: [new Map(), new Set(), new Uint8Array(1), new Date(0), undefined] };
    expect(nonPlainPaths(tainted, "model")).toHaveLength(5);
  });
});

describe("the invariants hold on read models", () => {
  it("the hand-built model keeps them", () => {
    expect(invariantViolations(handBuiltModel())).toEqual([]);
  });

  it.each(READ_FIXTURES)(
    "$name: rows strictly ascending by index, cells strictly ascending by column, no bare null cell, no bare empty row, unique sheet ids",
    async (fixture) => {
      const model = await readFixture(fixture);
      expect(model.sheets.length).toBeGreaterThan(0);
      expect(model.sheets.some((sheet) => sheet.rows.length > 0)).toBe(true);
      expect(invariantViolations(model)).toEqual([]);
      expect(model.activeSheetIndex).toBeGreaterThanOrEqual(0);
      expect(model.activeSheetIndex).toBeLessThan(model.sheets.length);
    },
  );

  it.each(READ_FIXTURES)("$name keeps every sheet's source — the splice base", async (fixture) => {
    const model = await readFixture(fixture);
    const partNames = new Set(model.opaque.map((part) => part.name));
    for (const sheet of model.sheets) {
      if (model.format === "xlsx") {
        expect(sheet.partName !== undefined && partNames.has(sheet.partName), `${sheet.name}: ${sheet.partName}`).toBe(true);
      } else {
        expect(sheet.sourceXml, sheet.name).toMatch(/^\s*<table:table[\s>]/);
      }
    }
  });

  it("the invariant check itself flags each violation", () => {
    const broken = handBuiltModel();
    broken.sheets[2]!.id = "sheet-a";
    broken.sheets[0]!.rows = [
      { index: 1, cells: [{ col: 0, value: 1 }] },
      { index: 1, cells: [{ col: 0, value: 2 }] },
      { index: 2, cells: [] },
      {
        index: 3,
        cells: [
          { col: 2, value: 1 },
          { col: 2, value: 2 },
          { col: 1, value: null },
          { col: 3, value: "ab", runs: [{ text: "a" }] },
        ],
      },
    ];
    const problems = invariantViolations(broken);
    expect(problems).toHaveLength(7);
    expect(problems).toEqual(
      expect.arrayContaining([
        expect.stringContaining("duplicate or empty sheet id"),
        expect.stringContaining("Summary!row 1: not strictly ascending"),
        expect.stringContaining("Summary!row 2: stored with no cells"),
        expect.stringContaining("Summary!C4: column not strictly ascending"),
        expect.stringContaining("Summary!B4: column not strictly ascending"),
        expect.stringContaining("Summary!B4: stored with value null"),
        expect.stringContaining("Summary!D4: value is not its runs' concatenated text"),
      ]),
    );
  });
});

describe("toCellError and isCellError", () => {
  it("keeps every error code the model defines", () => {
    for (const code of ERROR_CODES) {
      expect(toCellError(code)).toEqual({ error: code });
    }
    expect(toCellError("#N/A")).toEqual({ error: "#N/A" });
  });

  it.each(["#BOGUS!", "", "N/A", "#DIV/0", "#NULL", "DIV/0!", "#VALUE"])("reads the unknown code %j as #VALUE!", (code) => {
    expect(toCellError(code)).toEqual({ error: "#VALUE!" });
  });

  it("tells an error from every other cell value", () => {
    expect(isCellError({ error: "#N/A" })).toBe(true);
    expect(isCellError(toCellError("#BOGUS!"))).toBe(true);
    const others: CellValue[] = [null, 0, -1.5, 46_303.75, "", "#N/A", "text", true, false];
    for (const value of others) {
      expect(isCellError(value), JSON.stringify(value)).toBe(false);
    }
  });
});

describe("model.ts's header", () => {
  it("states the four invariants as normative for every later spec", () => {
    const source = readFileSync(resolve(spreadsheetDir, "model.ts"), "utf8");
    const header = source.slice(0, source.search(/^(?:import|export)\b/m));
    expect(header.length).toBeGreaterThan(0);
    expect(header).toMatch(/normative/i);
    for (const invariant of ["Copy-on-write rows", "Sparse", "The source is the splice base", "Dates are serials"]) {
      expect(header).toContain(invariant);
    }
  });
});
