// SPEC-140 — the JSON workbook description: every row of the Reference's key
// table produces exactly what the matching WorkbookBuilder calls produce
// (1:1), unknown keys at each level are validation errors in
// rejectUnknownKeys' own wording, `{}` describes a one-sheet empty workbook,
// a description and the program that would build it fail with the same
// path, and the published schema/workbook-description.json has the runtime
// validator's closed shape (beside SPEC-045's document-description.json).
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { ERROR_CODES, type SpreadsheetModel } from "../spreadsheet/model";
import { SPREADSHEET_LOCALE_TAGS } from "../spreadsheet/numberFormat/locale";
import { DocumentGenerationError } from "./errors";
import { WorkbookBuilder } from "./workbook";
import {
  COLUMN_ENTRY_KEYS,
  DATA_KEYS,
  DESCRIPTION_KEYS,
  ROW_ENTRY_KEYS,
  SHEET_KEYS,
  STYLE_ENTRY_KEYS,
  workbookFromDescription,
  type WorkbookDescription,
} from "./workbookDescription";
import { CELL_INPUT_KEYS } from "./workbookPlan";
import { ALIGNMENT_SPEC_KEYS, BORDER_EDGE_KEYS, BORDER_SPEC_KEYS, FONT_SPEC_KEYS, SHEET_BORDER_STYLES, STYLE_SPEC_KEYS } from "./workbookStyles";

/** A model with its minted ids masked. */
function masked(model: SpreadsheetModel): unknown {
  return { ...model, id: "", sheets: model.sheets.map((sheet) => ({ ...sheet, id: "" })) };
}

function expectSame(description: unknown, builder: WorkbookBuilder): SpreadsheetModel {
  const fromJson = workbookFromDescription(description);
  expect(masked(fromJson)).toEqual(masked(builder.toWorkbook()));
  return fromJson;
}

const fails = (build: () => unknown): string => {
  try {
    build();
  } catch (error) {
    expect(error).toBeInstanceOf(DocumentGenerationError);
    return (error as Error).message;
  }
  throw new Error("expected a DocumentGenerationError");
};

/** The Reference's description, verbatim. */
const REFERENCE = {
  title: "Q3 sales",
  format: "xlsx",
  dateSystem: "1900",
  locale: "en-US",
  styles: { header: { font: { bold: true }, fill: "#dce6f1" } },
  sheets: [
    {
      name: "Sales",
      tabColor: "#4472c4",
      hidden: false,
      showGridLines: true,
      cells: { A1: "Region", B5: { formula: "SUM(B2:B4)", numberFormat: "$#,##0.00" }, F2: { date: "2026-10-08" } },
      data: [
        {
          at: "A2",
          rows: [
            ["East", 1200],
            ["West", 980],
          ],
          parseText: false,
        },
      ],
      styles: [
        { range: "A1:E1", style: "header" },
        { range: "B2:E5", numberFormat: "$#,##0.00" },
      ],
      columns: [{ columns: "A", width: { chars: 14 } }],
      rows: [{ rows: "1", height: { pt: 24 } }],
      merges: ["A10:E10"],
      freeze: { rows: 1, cols: 0 },
    },
  ],
};

describe("workbookFromDescription — one row of the key table each (AC-30)", () => {
  it("title, format, dateSystem and locale are the constructor options", () => {
    const description = { title: "Budget", format: "ods", dateSystem: "1904", locale: "en-GB", sheets: [{ name: "S", data: [{ at: "A1", rows: [["10/08/2026"]], parseText: true }] }] };
    const model = expectSame(
      description,
      new WorkbookBuilder({ title: "Budget", format: "ods", dateSystem: "1904", locale: "en-GB" }).addSheet("S").setRows("A1", [["10/08/2026"]], { parseText: true }),
    );
    expect([model.title, model.format, model.dateSystem]).toEqual(["Budget", "ods", "1904"]);
    expect(typeof model.sheets[0]!.rows[0]!.cells[0]!.value).toBe("number");
  });

  it("top-level styles is defineStyle per entry", () => {
    const header = { font: { bold: true }, borders: { bottom: "thin" } } as const;
    expectSame(
      { styles: { header }, sheets: [{ name: "S", styles: [{ range: "A1:C1", style: "header" }] }] },
      new WorkbookBuilder().defineStyle("header", header).addSheet("S").style("A1:C1", "header"),
    );
  });

  it("sheets[i] with name, tabColor, hidden and showGridLines is addSheet", () => {
    expectSame(
      { sheets: [{ name: "One", tabColor: "#ff0000", showGridLines: false }, { name: "Two", hidden: true }] },
      new WorkbookBuilder().addSheet("One", { tabColor: "#ff0000", showGridLines: false }).addSheet("Two", { hidden: true }),
    );
  });

  it("cells is setCell, in key order", () => {
    const model = expectSame(
      { sheets: [{ name: "S", cells: { C3: 3, A1: { value: "a", style: { font: { italic: true } } }, B2: { formula: "C3*2" }, $A$1: "again" } }] },
      new WorkbookBuilder().addSheet("S").setCell("C3", 3).setCell("A1", { value: "a", style: { font: { italic: true } } }).setCell("B2", { formula: "C3*2" }).setCell("$A$1", "again"),
    );
    expect(model.sheets[0]!.rows.map((row) => row.cells.map((cell) => cell.value))).toEqual([["again"], [6], [3]]);
  });

  it("data[] is setRows, with parseText", () => {
    expectSame(
      {
        sheets: [
          {
            name: "S",
            data: [
              { at: "A1", rows: [["Item", "Price"], ["Tea", 3.5]] },
              { at: "D1", rows: [["1,234.50", "45%"]], parseText: true },
            ],
          },
        ],
      },
      new WorkbookBuilder()
        .addSheet("S")
        .setRows("A1", [
          ["Item", "Price"],
          ["Tea", 3.5],
        ])
        .setRows("D1", [["1,234.50", "45%"]], { parseText: true }),
    );
  });

  it("styles[] is style and/or numberFormat over a range, style first", () => {
    expectSame(
      { sheets: [{ name: "S", cells: { A1: 1 }, styles: [{ range: "A1:B2", style: { fill: "#eeeeee", numberFormat: "0.0" }, numberFormat: "0.00" }, { range: "B:B", numberFormat: "0%" }] }] },
      new WorkbookBuilder().addSheet("S").setCell("A1", 1).style("A1:B2", { fill: "#eeeeee", numberFormat: "0.0" }).numberFormat("A1:B2", "0.00").numberFormat("B:B", "0%"),
    );
  });

  it("columns[] and rows[] are columnWidth and rowHeight", () => {
    expectSame(
      {
        sheets: [
          {
            name: "S",
            columns: [{ columns: "A", width: { chars: 14 } }, { columns: "B:D", width: { px: 90 } }],
            rows: [{ rows: "1", height: { pt: 24 } }, { rows: 3, height: { px: 40 } }, { rows: "5:6", height: { pt: 30 } }],
          },
        ],
      },
      new WorkbookBuilder()
        .addSheet("S")
        .columnWidth("A", { chars: 14 })
        .columnWidth("B:D", { px: 90 })
        .rowHeight("1", { pt: 24 })
        .rowHeight(3, { px: 40 })
        .rowHeight("5:6", { pt: 30 }),
    );
  });

  it("merges and freeze are merge and freeze", () => {
    expectSame({ sheets: [{ name: "S", cells: { A1: "Title" }, merges: ["A1:E1", "A3:A4"], freeze: { rows: 1, cols: 2 } }] }, new WorkbookBuilder().addSheet("S").setCell("A1", "Title").merge("A1:E1").merge("A3:A4").freeze({ rows: 1, cols: 2 }));
  });

  it("the Reference's description builds what its builder calls build", () => {
    const model = expectSame(
      REFERENCE,
      new WorkbookBuilder({ title: "Q3 sales", format: "xlsx", dateSystem: "1900", locale: "en-US" })
        .defineStyle("header", { font: { bold: true }, fill: "#dce6f1" })
        .addSheet("Sales", { tabColor: "#4472c4", hidden: false, showGridLines: true })
        .setCell("A1", "Region")
        .setCell("B5", { formula: "SUM(B2:B4)", numberFormat: "$#,##0.00" })
        .setCell("F2", { date: "2026-10-08" })
        .setRows("A2", [
          ["East", 1200],
          ["West", 980],
        ], { parseText: false })
        .style("A1:E1", "header")
        .numberFormat("B2:E5", "$#,##0.00")
        .columnWidth("A", { chars: 14 })
        .rowHeight("1", { pt: 24 })
        .merge("A10:E10")
        .freeze({ rows: 1, cols: 0 }),
    );
    const b5 = model.sheets[0]!.rows.find((row) => row.index === 4)!.cells[0]!;
    expect(b5).toMatchObject({ col: 1, value: 2180, formula: "SUM(B2:B4)" });
    const typed: WorkbookDescription = REFERENCE as WorkbookDescription;
    expect(typed.sheets).toHaveLength(1);
  });
});

describe("validation (AC-32, AC-33)", () => {
  it("{} describes a one-sheet empty workbook", () => {
    for (const description of [{}, { sheets: [] }]) {
      const model = workbookFromDescription(description);
      expect(model.sheets.map((sheet) => [sheet.name, sheet.rows.length])).toEqual([["Sheet1", 0]]);
      expect(masked(model)).toEqual(masked(new WorkbookBuilder().toWorkbook()));
    }
  });

  it("an unknown key at any level is an unknown property, named by its JSON path", () => {
    const sheet = (extra: Record<string, unknown>): unknown => ({ sheets: [{ name: "One" }, { name: "Two", ...extra }] });
    const cases: [unknown, string][] = [
      [{ names: [] }, "names: unknown property"],
      [{ title: "x", charts: [] }, "charts: unknown property"],
      [{ styles: { header: { font: { bold: true }, colour: "#000000" } } }, "styles.header.colour: unknown property"],
      [{ styles: { header: { font: { weight: 700 } } } }, "styles.header.font.weight: unknown property"],
      [{ styles: { header: { borders: { inside: "thin" } } } }, "styles.header.borders.inside: unknown property"],
      [{ styles: { header: { borders: { top: { style: "thin", width: 1 } } } } }, "styles.header.borders.top.width: unknown property"],
      [{ styles: { header: { alignment: { shrink: true } } } }, "styles.header.alignment.shrink: unknown property"],
      [sheet({ validations: [] }), "sheets[1].validations: unknown property"],
      [sheet({ conditionalFormats: [] }), "sheets[1].conditionalFormats: unknown property"],
      [sheet({ hyperlinks: [] }), "sheets[1].hyperlinks: unknown property"],
      [sheet({ cells: { C4: { fomula: "1" } } }), "sheets[1].cells.C4.fomula: unknown property"],
      [sheet({ cells: { C4: { value: 1, comment: "n" } } }), "sheets[1].cells.C4.comment: unknown property"],
      [sheet({ cells: { C4: { value: 1, style: { font: { size: 3 } } } } }), "sheets[1].cells.C4.style.font.size: unknown property"],
      [sheet({ data: [{ at: "A1", rows: [], header: true }] }), "sheets[1].data[0].header: unknown property"],
      [sheet({ data: [{ at: "A1", rows: [[{ formula: "1", note: "x" }]] }] }), "sheets[1].data[0].rows[0][0].note: unknown property"],
      [sheet({ styles: [{ range: "A1", fill: "#000000" }] }), "sheets[1].styles[0].fill: unknown property"],
      [sheet({ styles: [{ range: "A1", style: { fill: "#000000", pattern: "solid" } }] }), "sheets[1].styles[0].style.pattern: unknown property"],
      [sheet({ columns: [{ columns: "A", width: { chars: 1 }, hidden: true }] }), "sheets[1].columns[0].hidden: unknown property"],
      [sheet({ columns: [{ columns: "A", width: { em: 1 } }] }), "sheets[1].columns[0].width.em: unknown property"],
      [sheet({ rows: [{ rows: 1, height: { pt: 1 }, outline: 1 }] }), "sheets[1].rows[0].outline: unknown property"],
      [sheet({ rows: [{ rows: 1, height: { cm: 1 } }] }), "sheets[1].rows[0].height.cm: unknown property"],
      [sheet({ freeze: { rows: 1, split: true } }), "sheets[1].freeze.split: unknown property"],
    ];
    for (const [description, message] of cases) {
      expect(fails(() => workbookFromDescription(description))).toBe(message);
    }
  });

  it("names the offending path from the description's root", () => {
    const cases: [unknown, string][] = [
      ["workbook", "description: expected an object"],
      [{ sheets: {} }, "sheets: expected an array"],
      [{ sheets: ["Sales"] }, "sheets[0]: expected an object"],
      [{ sheets: [{}] }, "sheets[0].name: expected a string"],
      [{ sheets: [{ name: "A" }, { name: "a" }] }, `sheets[1].name: There's already a sheet called "a"`],
      [{ sheets: [{ name: "S", cells: [] }] }, "sheets[0].cells: expected an object"],
      [{ sheets: [{ name: "S", cells: { B5: { formula: "SUMM(B2:B4)" } } }] }, "sheets[0].cells.B5.formula: unknown function SUMM"],
      [{ sheets: [{ name: "S", data: [{ at: "A1", rows: [[1, 2], [3, { date: "2026-10-08Z" }]] }] }] }, 'sheets[0].data[0].rows[1][1].date: expected "YYYY-MM-DD[THH:MM[:SS[.sss]]]", got "2026-10-08Z"'],
      [{ sheets: [{ name: "S", data: [{ rows: [] }] }] }, "sheets[0].data[0].at: expected a string"],
      [{ sheets: [{ name: "S", styles: [{ range: "A1" }] }] }, 'sheets[0].styles[0]: expected "style" or "numberFormat"'],
      [{ styles: { header: { font: { sizePt: 0 } } } }, "styles.header.font.sizePt: expected a size between 1 and 409 pt"],
      [{ styles: [] }, "styles: expected an object"],
      [{ sheets: [{ name: "S", merges: "A1:B1" }] }, "sheets[0].merges: expected an array"],
      [{ sheets: [{ name: "S", merges: ["A1:B1", "B1:C1"] }] }, "sheets[0].merges[1]: B1:C1 overlaps the merge A1:B1"],
      [{ format: "numbers" }, 'format: expected one of "xlsx", "ods", got "numbers"'],
    ];
    for (const [description, message] of cases) {
      expect(fails(() => workbookFromDescription(description))).toBe(message);
    }
    expect(fails(() => workbookFromDescription({ sheets: [{ name: "Sales", cells: { B2: { formula: "B2*2" } } }] }))).toBe("sheets[0].cells.B2.formula: circular reference (Sales!B2)");
  });

  it("a description and the program that would build it fail identically", () => {
    const pairs: [unknown, () => unknown][] = [
      [{ sheets: [{ name: "S", cells: { B5: { formula: "SUMM(1)" } } }] }, () => new WorkbookBuilder().addSheet("S").setFormula("B5", "SUMM(1)")],
      [{ sheets: [{ name: "S", cells: { C4: { fomula: "1" } } }] }, () => new WorkbookBuilder().addSheet("S").setCell("C4", { fomula: "1" } as never)],
      [
        { sheets: [{ name: "A" }, { name: "B", data: [{ at: "A1", rows: [[1]] }, { at: "A2", rows: [[1, 2, 3, Number.NaN]] }] }] },
        () => new WorkbookBuilder().addSheet("A").addSheet("B").setRow("A1", [1]).setRows("A2", [[1, 2, 3, Number.NaN]]),
      ],
      [
        { sheets: [{ name: "S", styles: [{ range: "A1", numberFormat: "0" }, { range: "A1", style: { font: { sizePt: 0 } } }] }] },
        () => new WorkbookBuilder().addSheet("S").numberFormat("A1", "0").style("A1", { font: { sizePt: 0 } }),
      ],
      [{ sheets: [{ name: "S", columns: [{ columns: "A", width: { chars: 300 } }] }] }, () => new WorkbookBuilder().addSheet("S").columnWidth("A", { chars: 300 })],
      [{ sheets: [{ name: "S", rows: [{ rows: "2", height: { pt: 500 } }] }] }, () => new WorkbookBuilder().addSheet("S").rowHeight("2", { pt: 500 })],
      [{ sheets: [{ name: "S", freeze: { rows: 1.5 } }] }, () => new WorkbookBuilder().addSheet("S").freeze({ rows: 1.5 })],
      [{ sheets: [{ name: "S", cells: { A1: 1 }, merges: ["B1:C1", "C3:D3", "A1:D1"] }] }, () => new WorkbookBuilder().addSheet("S").setCell("A1", 1).merge("B1:C1").merge("C3:D3").merge("A1:D1").toWorkbook()],
      [{ sheets: [{ name: "S", cells: { A1: 1, C1: "x" }, merges: ["A1:D1"] }] }, () => new WorkbookBuilder().addSheet("S").setCell("A1", 1).setCell("C1", "x").merge("A1:D1").toWorkbook()],
      [{ sheets: [{ name: "Sales", tabColor: "red" }] }, () => new WorkbookBuilder().addSheet("Sales", { tabColor: "red" })],
      [{ styles: { header: { fill: "red" } } }, () => new WorkbookBuilder().defineStyle("header", { fill: "red" })],
      [{ title: 1 }, () => new WorkbookBuilder({ title: 1 } as never)],
    ];
    for (const [description, program] of pairs) {
      const fromJson = fails(() => workbookFromDescription(description));
      expect(fails(program)).toBe(fromJson);
    }
  });
});

describe("WorkbookBuilder.fromDescription (AC-31)", () => {
  it("is the static factory: the description's workbook, ready for more calls on its last sheet", async () => {
    const builder = WorkbookBuilder.fromDescription(REFERENCE);
    expect(builder).toBeInstanceOf(WorkbookBuilder);
    expect(masked(builder.toWorkbook())).toEqual(masked(workbookFromDescription(REFERENCE)));
    builder.setCell("B4", 400).addSheet("Notes").setCell("A1", "more");
    const model = builder.toWorkbook();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Sales", "Notes"]);
    expect(model.sheets[0]!.rows.find((row) => row.index === 4)!.cells[0]!.value).toBe(2580);
    expect(builder.fileName()).toBe("Q3 sales.xlsx");
    expect((await builder.toBytes()).length).toBeGreaterThan(0);
  });
});

// The published schema describes the same closed shape the runtime
// validator enforces — checked structurally, as SPEC-045's and SPEC-106's
// schemas are (the project carries no JSON Schema engine).
describe("schema/workbook-description.json", () => {
  const schemaDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "schema");
  const schema = JSON.parse(readFileSync(resolve(schemaDir, "workbook-description.json"), "utf8"));
  const documentSchema = JSON.parse(readFileSync(resolve(schemaDir, "document-description.json"), "utf8"));
  const defs = schema.definitions;
  const keys = (shape: { properties: Record<string, unknown> }): string[] => Object.keys(shape.properties).toSorted();
  const sorted = (list: readonly string[]): string[] => [...list].toSorted();

  it("is a closed draft-07 object with the description's top-level keys", () => {
    expect(schema.$schema).toBe(documentSchema.$schema);
    expect(schema.type).toBe("object");
    expect(schema.additionalProperties).toBe(false);
    expect(schema.required).toBeUndefined();
    expect(keys(schema)).toEqual(sorted(DESCRIPTION_KEYS));
    expect(schema.properties.format.enum).toEqual(["xlsx", "ods"]);
    expect(schema.properties.dateSystem.enum).toEqual(["1900", "1904"]);
    expect(schema.properties.locale.enum).toEqual([...SPREADSHEET_LOCALE_TAGS]);
    expect(schema.properties.styles.additionalProperties).toEqual({ $ref: "#/definitions/cellStyle" });
  });

  it("keys every nested object exactly as the validator does", () => {
    const sheet = defs.sheet;
    expect(keys(sheet)).toEqual(sorted(SHEET_KEYS));
    expect(sheet.required).toEqual(["name"]);
    expect(keys(defs.cellObject)).toEqual(sorted(CELL_INPUT_KEYS));
    expect(keys(sheet.properties.data.items)).toEqual(sorted(DATA_KEYS));
    expect(keys(sheet.properties.styles.items)).toEqual(sorted(STYLE_ENTRY_KEYS));
    expect(keys(sheet.properties.columns.items)).toEqual(sorted(COLUMN_ENTRY_KEYS));
    expect(keys(sheet.properties.rows.items)).toEqual(sorted(ROW_ENTRY_KEYS));
    expect(keys(sheet.properties.freeze)).toEqual(["cols", "rows"]);
    expect(keys(defs.cellStyle)).toEqual(sorted(STYLE_SPEC_KEYS));
    expect(keys(defs.cellStyle.properties.font)).toEqual(sorted(FONT_SPEC_KEYS));
    expect(keys(defs.cellStyle.properties.borders)).toEqual(sorted(BORDER_SPEC_KEYS));
    expect(keys(defs.border.oneOf[1])).toEqual(sorted(BORDER_EDGE_KEYS));
    expect(keys(defs.cellStyle.properties.alignment)).toEqual(sorted(ALIGNMENT_SPEC_KEYS));
    expect(defs.borderStyle.enum).toEqual([...SHEET_BORDER_STYLES]);
    expect(defs.cellObject.properties.error.enum).toEqual([...ERROR_CODES]);
  });

  it("has additionalProperties: false on every object with properties", () => {
    const open: string[] = [];
    const visit = (node: unknown, path: string): void => {
      if (Array.isArray(node)) {
        node.forEach((item, i) => visit(item, `${path}[${i}]`));
        return;
      }
      if (typeof node !== "object" || node === null) {
        return;
      }
      const record = node as Record<string, unknown>;
      if (record.properties !== undefined && record.additionalProperties !== false) {
        open.push(path);
      }
      for (const [key, child] of Object.entries(record)) {
        visit(child, `${path}.${key}`);
      }
    };
    visit(schema, "$");
    expect(open).toEqual([]);
  });

  it("the Reference's description is the shape the schema describes", () => {
    expect(keys(schema)).toEqual(expect.arrayContaining(Object.keys(REFERENCE)));
    expect(keys(defs.sheet)).toEqual(expect.arrayContaining(Object.keys(REFERENCE.sheets[0]!)));
  });
});
