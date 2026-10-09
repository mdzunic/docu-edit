// SPEC-116 "The workbook model": the grid and workbook limits; the model is
// plain JSON — a hand-built model with every optional field, and every model
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
} from "./fixtures/build";
import {
  isCellError,
  MAX_COLS,
  MAX_ROWS,
  MAX_WORKBOOK_CELLS,
  toCellError,
  type CellValue,
  type ErrorCode,
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
              { col: 5, value: "Bold and plain", runs: [{ text: "Bold" }, { text: " and plain" }] },
            ],
          },
          { index: 4, cells: [{ col: 1, value: false }] },
          { index: MAX_ROWS - 1, cells: [{ col: MAX_COLS - 1, value: -0.5 }] },
        ],
        sourceXml: '\n   <table:table table:name="Summary" table:style-name="ta1">…</table:table>',
      },
      {
        id: "sheet-b",
        name: "Chart1",
        visibility: "hidden",
        kind: "foreign",
        rows: [],
        partName: "xl/chartsheets/sheet1.xml",
      },
      {
        id: "sheet-c",
        name: "Secret",
        visibility: "veryHidden",
        kind: "grid",
        rows: [{ index: 3, cells: [{ col: 2, value: { error: "#N/A" } }] }],
        sourceXml: '<table:table table:name="Secret">…</table:table>',
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
      if (row.cells.length === 0) {
        problems.push(`${where}: stored with no cells`);
      }
      let previousCol = -1;
      for (const cell of row.cells) {
        const at = `${sheet.name}!${Number.isInteger(cell.col) && cell.col >= 0 ? formatA1({ row: row.index, col: cell.col }) : `col ${cell.col}`}`;
        if (!Number.isInteger(cell.col) || cell.col <= previousCol || cell.col >= MAX_COLS) {
          problems.push(`${at}: column not strictly ascending inside the grid (after col ${previousCol})`);
        }
        previousCol = cell.col;
        if (cell.value === null) {
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
    "$name: rows strictly ascending by index, cells strictly ascending by column, no null cell, no empty row, unique sheet ids",
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
