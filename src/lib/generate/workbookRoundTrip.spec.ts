// SPEC-140 "Round-trip relation": a workbook built with every construct this
// API emits, written as .xlsx and as .ods and read back with readWorkbook, is
// `toEqual` the generated model under comparableWorkbook — per sheet in
// order its name, visibility, tab colour, frozen panes, merges, column
// widths (±1 px), row heights, and per stored cell its value, formula and
// effective CellStyle. Also: byte determinism with a fixed `modified`, and
// every formula's cached value equal to what SPEC-120's engine computes on
// the re-read workbook.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { DEFAULT_THEME, resolveColor, resolveFont, type ModelColor } from "../core/theme";
import { unzip } from "../core/zip";
import { formatA1 } from "../spreadsheet/address";
import { cellKey } from "../spreadsheet/formula/dependencies";
import { recalculate } from "../spreadsheet/formula/recalc";
import type { CellBorderEdge, CellStyle, Sheet, SpreadsheetModel } from "../spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../spreadsheet/numberFormat/locale";
import { formatValue } from "../spreadsheet/numberFormat/render";
import { cellStyle, effectiveStyleId } from "../spreadsheet/styles";
import { readWorkbook } from "./readWorkbook";
import { WorkbookBuilder, type WorkbookOutputFormat } from "./workbook";
import { SHEET_BORDER_STYLES } from "./workbookStyles";

const MODIFIED = new Date(Date.UTC(2026, 9, 8, 9, 30));
const BLACK = "#000000";

/** A cell's effective style as compared: the CellStyle its cell, row or
 * column gives it, without sourceXf/sourceStyleName/namedStyle (style
 * indices and source names differ between a generated and a re-read model
 * by design), with every colour and font as it renders — theme references
 * resolved against the workbook's theme (an .ods has none: its writer
 * resolves them), an absent (automatic) font or border colour as the black
 * it is. In .ods the number format is compared by its rendered text
 * (SPEC-118 formatValue, en-US), and slantDashDot as the mediumDashDot
 * SPEC-122's ODS border table writes for it (ODF has no slanted dash-dot). */
function comparableStyle(model: SpreadsheetModel, style: CellStyle, value: SpreadsheetModel["sheets"][number]["rows"][number]["cells"][number]["value"], format: WorkbookOutputFormat): Record<string, unknown> {
  const theme = model.theme ?? DEFAULT_THEME;
  const color = (c: ModelColor | undefined): string => (c === undefined ? BLACK : resolveColor(c, theme).toLowerCase());
  const { sourceXf: _xf, sourceStyleName: _name, namedStyle: _named, builtinFormatId, numberFormat, font, fill, borders, ...rest } = style;
  const out: Record<string, unknown> = { ...rest };
  out.font = { ...font, family: font?.family === undefined ? undefined : resolveFont(font.family, theme), color: color(font?.color) };
  if (fill !== undefined) {
    out.fill = fill.kind === "solid" ? { kind: "solid", color: color(fill.color) } : fill;
  }
  if (borders !== undefined) {
    const edge = (e: CellBorderEdge | undefined): unknown =>
      e === undefined ? undefined : { style: format === "ods" && e.style === "slantDashDot" ? "mediumDashDot" : e.style, color: color(e.color) };
    out.borders = { top: edge(borders.top), right: edge(borders.right), bottom: edge(borders.bottom), left: edge(borders.left) };
  }
  if (format === "xlsx") {
    out.numberFormat = numberFormat;
    out.builtinFormatId = builtinFormatId;
  } else {
    out.rendered = formatValue(value, numberFormat, { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: model.dateSystem }).text;
  }
  return JSON.parse(JSON.stringify(out));
}

/** Every column's width, one entry per column a span sizes or hides — a
 * span at the sheet's default width (the one an .xlsx `<col>` holding only a
 * style must state) sizes nothing. */
function columnWidths(sheet: Pick<Sheet, "columns" | "defaultColWidthPx">): { col: number; widthPx?: number; hidden?: true }[] {
  return sheet.columns
    .filter((span) => span.hidden || (span.widthPx !== undefined && Math.abs(span.widthPx - sheet.defaultColWidthPx) >= 1))
    .flatMap((span) => Array.from({ length: span.max - span.min + 1 }, (_, k) => ({ col: span.min + k, ...(span.hidden ? { hidden: true as const } : { widthPx: span.widthPx }) })));
}

/** SPEC-140's comparableWorkbook (column widths split out: they compare
 * within 1 px). */
function comparableWorkbook(model: SpreadsheetModel, format: WorkbookOutputFormat): unknown {
  const theme = model.theme ?? DEFAULT_THEME;
  return model.sheets.map((sheet) => ({
    name: sheet.name,
    visibility: sheet.visibility,
    tabColor: sheet.tabColor === undefined ? null : resolveColor(sheet.tabColor, theme).toLowerCase(),
    frozen: sheet.frozen ?? null,
    merges: sheet.merges,
    rowHeights: sheet.rows.filter((row) => row.heightPt !== undefined || row.hidden).map((row) => ({ index: row.index, heightPt: row.heightPt, customHeight: row.customHeight, hidden: row.hidden })),
    cells: sheet.rows.flatMap((row) =>
      row.cells.map((cell) => ({
        ref: formatA1({ row: row.index, col: cell.col }),
        value: cell.value,
        formula: cell.formula ?? null,
        style: comparableStyle(model, cellStyle(model, effectiveStyleId(sheet, row.index, cell.col)), cell.value, format),
      })),
    ),
  }));
}

/** A workbook with every construct the API emits. */
function everyConstruct(format: WorkbookOutputFormat): WorkbookBuilder {
  const builder = new WorkbookBuilder({ title: "Every construct", format, locale: "en-GB" })
    .defineStyle("header", { font: { bold: true, color: "#ffffff" }, fill: "#1f4e79", borders: { bottom: { style: "medium", color: "#000000" } }, alignment: { horizontal: "center" } })
    .addSheet("Data", { tabColor: "#4472c4", showGridLines: false })
    .setRow("A1", ["Text", "Number", "Date", "Flag", "Formula"])
    .style("A1:E1", "header")
    .setRow("A2", ["plain", 42, new Date(Date.UTC(2026, 9, 8, 13, 45)), true, { formula: "B2*2" }])
    .setRow("A3", ["=not a formula", -1.5, { date: "2026-10-08" }, false, { formula: "'Q3 Notes'!A1&\"!\"" }])
    .setRow("A4", ["Ünïcödé — ✓", 1e15, { date: "1904-03-01T06:00:30" }, { value: null, style: { fill: "#eeeeee" } }, { formula: "C3>B2" }])
    .setRow("A5", [{ error: "#N/A" }, { error: "#DIV/0!" }, { error: "#VALUE!" }, { error: "#REF!" }, { formula: "1/0" }])
    .setRow("A6", [{ error: "#NAME?" }, { error: "#NUM!" }, { error: "#NULL!" }, 0.000001, { formula: 'IFERROR(1/0,"none")' }])
    .setRows("A7", [["1,234.50", "45%", "10/08/2026", "TRUE", "007"]], { parseText: true })
    .setFormula("F2", "SUM(B2:B4)")
    .setCell("F3", { formula: "AVERAGE(B2:B3)", numberFormat: "0.00" })
    .setCell("F4", { formula: "SUM(Data!B2,B3)", style: { font: { italic: true } } })
    .numberFormat("B2:B4", "#,##0.00")
    .numberFormat("D6", "0.00E+00");
  // Every font, fill, alignment and number-format field, one cell each.
  builder
    .setRow("A9", ["font", "underline", "fill", "align", "rotate", "stacked", "currency", "percent", "units", "time"])
    .style("A9", { font: { family: "Arial", sizePt: 14, bold: true, italic: true, underline: "single", strike: true, color: "#1f4e79" } })
    .style("B9", { font: { underline: "double", family: "Times New Roman", sizePt: 9 } })
    .style("C9", { fill: "#ffff00" })
    .style("D9", { alignment: { horizontal: "right", vertical: "top", wrap: true, indent: 2 } })
    .style("E9", { alignment: { horizontal: "justify", vertical: "center", rotation: 45 } })
    .style("F9", { alignment: { horizontal: "left", vertical: "bottom", rotation: 255 } })
    .setRow("G9", [1234.5, 0.125, 72.25, 0.75])
    .numberFormat("G9", "$#,##0.00")
    .numberFormat("H9", "0.00%")
    .numberFormat("I9", '0.0 "kg"')
    .numberFormat("J9", "hh:mm:ss")
    .style("E10", { alignment: { rotation: 135 } })
    .setCell("E10", "down");
  // Every SPEC-117 border style, all four edges of one cell each; the
  // shorthands on a block.
  SHEET_BORDER_STYLES.forEach((style, k) => {
    builder.setCell(`${String.fromCharCode(65 + k)}12`, style).style(`${String.fromCharCode(65 + k)}12`, { borders: { all: { style, color: "#c00000" } } });
  });
  builder
    .style("B14:D16", { borders: { outline: "thick", top: "dashed" } })
    .style("F14:G15", { borders: { all: "hair", left: { style: "double", color: "#00b050" } } })
    // Sizes, merges and panes.
    .columnWidth("A", { chars: 14 })
    .columnWidth("B:C", { px: 100 })
    .columnWidth("E", { chars: 20.5 })
    .columnWidth("K", { chars: 0.4 })
    .rowHeight(1, { pt: 24 })
    .rowHeight("3:4", { px: 30 })
    .rowHeight(9, { pt: 31.5 })
    .setCell("A18", { value: "Merged title", style: { alignment: { horizontal: "center" } } })
    .merge("A18:E18")
    .merge("G18:H20")
    .freeze({ rows: 1, cols: 1 })
    // A whole-column and a whole-row style.
    .style("L:L", { fill: "#f2f2f2" })
    .setCell("L2", "in L")
    .numberFormat("22:22", "0.0")
    .setRow("A22", [1.25, 2.5])
    .addSheet("Q3 Notes", { hidden: true })
    .setCell("A1", "notes")
    .freeze({ rows: 2 })
    .addSheet("Plain")
    .setCell("B2", "x")
    .setFormula("B3", "Data!F2/2");
  return builder;
}

describe("the round-trip relation (AC-36, AC-37)", () => {
  for (const format of ["xlsx", "ods"] as const) {
    it(`build → write .${format} → readWorkbook reproduces every construct`, async () => {
      const builder = everyConstruct(format);
      const generated = builder.toWorkbook();
      const read = await readWorkbook(await builder.toBytes(format, { modified: MODIFIED }));
      expect(read.format).toBe(format);
      expect(comparableWorkbook(read, format)).toEqual(comparableWorkbook(generated, format));
      // Column widths within 1 px.
      generated.sheets.forEach((sheet, i) => {
        const want = columnWidths(sheet);
        const got = columnWidths(read.sheets[i]!);
        expect(got.map(({ col, hidden }) => [col, hidden]), sheet.name).toEqual(want.map(({ col, hidden }) => [col, hidden]));
        want.forEach((column, k) => {
          if (column.widthPx !== undefined) {
            expect(Math.abs(got[k]!.widthPx! - column.widthPx), `${sheet.name} column ${column.col}`).toBeLessThanOrEqual(1);
          }
        });
      });
    });
  }

  it("the fixture holds what the relation checks (a guard on the guard)", () => {
    const model = everyConstruct("xlsx").toWorkbook();
    const data = model.sheets[0]!;
    expect(model.sheets.map((sheet) => [sheet.name, sheet.visibility])).toEqual([
      ["Data", "visible"],
      ["Q3 Notes", "hidden"],
      ["Plain", "visible"],
    ]);
    expect(data.frozen).toEqual({ rows: 1, cols: 1 });
    expect(data.merges).toHaveLength(2);
    expect(data.rows.filter((row) => row.customHeight)).toHaveLength(4);
    expect(columnWidths(data).length).toBeGreaterThanOrEqual(5);
    const usedStyles = new Set(data.rows.flatMap((row) => row.cells.map((cell) => cell.styleId)));
    expect(usedStyles.size).toBeGreaterThan(25);
    expect(data.rows.flatMap((row) => row.cells).filter((cell) => cell.formula !== undefined)).toHaveLength(8);
  });

  it("number formats round-trip by code in .xlsx and by rendered text in .ods", async () => {
    const builder = new WorkbookBuilder().addSheet("S").setRow("A1", [1234.5, 0.125, 72.25]).numberFormat("A1", "$#,##0.00").numberFormat("B1", "0.00%").numberFormat("C1", '0.0 "kg"');
    const xlsx = await readWorkbook(await builder.toBytes("xlsx"));
    expect(xlsx.sheets[0]!.rows[0]!.cells.map((cell) => xlsx.styles[cell.styleId!]!.numberFormat)).toEqual(["$#,##0.00", "0.00%", '0.0 "kg"']);
    const ods = await readWorkbook(await builder.toBytes("ods"));
    const ctx = { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" as const };
    expect(ods.sheets[0]!.rows[0]!.cells.map((cell) => formatValue(cell.value, ods.styles[cell.styleId!]!.numberFormat, ctx).text)).toEqual(["$1,234.50", "12.50%", "72.3 kg"]);
  });
});

describe("determinism and cached values", () => {
  for (const format of ["xlsx", "ods"] as const) {
    it(`the same input with the same modified yields equalBytes .${format} parts`, async () => {
      const first = await unzip(await everyConstruct(format).toBytes(format, { modified: MODIFIED }));
      const second = await unzip(await everyConstruct(format).toBytes(format, { modified: MODIFIED }));
      expect([...second.keys()]).toEqual([...first.keys()]);
      for (const [name, data] of first) {
        expect(equalBytes(second.get(name)!, data), name).toBe(true);
      }
    });

    it(`every formula's cached value in the .${format} is what SPEC-120 computes on the re-read workbook`, async () => {
      const read = await readWorkbook(await everyConstruct(format).toBytes(format));
      const cached = new Map<string, unknown>();
      for (const sheet of read.sheets) {
        for (const row of sheet.rows) {
          for (const cell of row.cells) {
            if (cell.formula !== undefined) {
              cached.set(cellKey(sheet.id, row.index, cell.col), cell.value);
            }
          }
        }
      }
      expect(cached.size).toBe(9);
      const result = recalculate(read, [...cached.keys()], { reason: "manual", full: true });
      expect(result.circular).toEqual([]);
      expect(result.stale).toEqual([]);
      expect(result.valueChanged).toEqual([]);
      const values = [...cached.values()];
      expect(values).toContain(84);
      expect(values).toContain(1e15 + 40.5);
      expect(values).toContain("notes!");
      expect(values).toContain("none");
      expect(values).toContainEqual({ error: "#DIV/0!" });
    });
  }
});
