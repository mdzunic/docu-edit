// SPEC-118's three fixtures through the reader and the grid's own formatting
// path (cellDisplay with the column's width and the font's MDW) — the facts
// SPEC118_FORMATS_XLSX and SPEC118_FORMATS_ODS record and e2e/SPEC-118.spec.ts
// asserts in the browser, pinned here so the two can never drift.
import { describe, expect, it } from "vitest";
import { DEFAULT_THEME } from "../core/theme";
import { getCell, parseA1 } from "./address";
import { cellDisplay } from "./cellRender";
import { buildSpec118FormatsOdsBytes, buildSpec118FormatsXlsxBytes, SPEC118_FORMATS_ODS, SPEC118_FORMATS_XLSX, SPEC118_NUMFMTS } from "./fixtures/build";
import { buildGridMetrics } from "./geometry";
import type { SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "./numberFormat/locale";
import { editText, workbookFormatContext, type FormattedValue } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { cellStyle, effectiveStyleId, fontMaxDigitWidth } from "./styles";

/** What the grid shows at an A1 address of the first sheet. */
function shown(model: SpreadsheetModel, a1: string, tag: SpreadsheetLocaleTag = "en-US"): FormattedValue {
  const sheet = model.sheets[0]!;
  const ref = parseA1(a1)!;
  const style = cellStyle(model, effectiveStyleId(sheet, ref.row, ref.col));
  return cellDisplay(getCell(sheet, ref.row, ref.col)?.value ?? null, style, sheet, {
    ...workbookFormatContext(model, SPREADSHEET_LOCALES[tag]),
    widthPx: buildGridMetrics(sheet).colWidth(ref.col),
    digitWidthPx: fontMaxDigitWidth(style.font, model.theme ?? DEFAULT_THEME),
  });
}

describe("spec118-formats.xlsx (Excel-shaped)", () => {
  const read = buildSpec118FormatsXlsxBytes().then((bytes) => readSpreadsheetBytes(bytes, "xlsx", "spec118-formats"));

  it("reads every custom numFmt and the explicit 43/44 as written, the locale-dependent 14 and 22 as builtinFormatId", async () => {
    const model = await read;
    const formats = new Set(model.styles.map((style) => style.numberFormat));
    for (const code of Object.values(SPEC118_NUMFMTS)) {
      expect(formats.has(code), code).toBe(true);
    }
    const byXf = (xf: number) => model.styles[xf]!;
    expect(byXf(SPEC118_FORMATS_XLSX.cells.A2.xf)).toMatchObject({ builtinFormatId: 14 });
    expect(byXf(SPEC118_FORMATS_XLSX.cells.A3.xf)).toMatchObject({ builtinFormatId: 22 });
    expect(byXf(SPEC118_FORMATS_XLSX.cells.E7.xf).builtinFormatId).toBeUndefined(); // 43 written explicitly
    expect(byXf(SPEC118_FORMATS_XLSX.cells.E8.xf).builtinFormatId).toBeUndefined(); // 44 written explicitly
  });

  it("lays out the wide A, B and E, the narrow F and the too-narrow G", async () => {
    const metrics = buildGridMetrics((await read).sheets[0]!);
    const W = SPEC118_FORMATS_XLSX.columnWidthsPx;
    expect([0, 1, 4, 5, 6].map((col) => metrics.colWidth(col))).toEqual([W.A, W.B, W.E, W.F, W.G]);
  });

  it("shows every cell as en-US Excel does (colour included), the narrow F as 1.23E+11 and G only #", async () => {
    const model = await read;
    for (const [a1, cell] of Object.entries(SPEC118_FORMATS_XLSX.cells)) {
      if (!("enUS" in cell)) {
        continue;
      }
      const fv = shown(model, a1);
      expect(fv.text, a1).toBe(cell.enUS);
      expect(fv.color, a1).toBe("color" in cell ? cell.color : undefined);
    }
    expect(shown(model, "G2").overflow).toBe("hash");
    expect(shown(model, "F2").overflow).toBeUndefined();
  });

  it("shows the de-DE texts in de-DE, and the formula bar's edit text in both", async () => {
    const model = await read;
    for (const [a1, cell] of Object.entries(SPEC118_FORMATS_XLSX.cells)) {
      if ("deDE" in cell) {
        expect(shown(model, a1, "de-DE").text, a1).toBe(cell.deDE);
      }
    }
    const sheet = model.sheets[0]!;
    const a4 = getCell(sheet, 3, 0)!;
    const style = cellStyle(model, effectiveStyleId(sheet, 3, 0));
    expect(editText(a4, { ...workbookFormatContext(model, SPREADSHEET_LOCALES["en-US"]), style })).toBe(SPEC118_FORMATS_XLSX.editText.enUS);
    expect(editText(a4, { ...workbookFormatContext(model, SPREADSHEET_LOCALES["de-DE"]), style })).toBe(SPEC118_FORMATS_XLSX.editText.deDE);
  });

  it("gives the accounting formats a space fill", async () => {
    const model = await read;
    expect(shown(model, "E7")).toMatchObject({ text: " 1,234.50 ", fill: { char: " ", index: 1 } });
    expect(shown(model, "E8")).toMatchObject({ text: " $(1,234.50)", fill: { char: " ", index: 2 } });
  });
});

describe("spec118-formats.ods (LibreOffice-shaped)", () => {
  const read = buildSpec118FormatsOdsBytes().then((bytes) => readSpreadsheetBytes(bytes, "ods", "spec118-formats"));

  it("maps every data style to its code and shows it", async () => {
    const model = await read;
    for (const [row, fact] of Object.entries(SPEC118_FORMATS_ODS.rows)) {
      const style = cellStyle(model, effectiveStyleId(model.sheets[0]!, Number(row) - 1, 1));
      expect(style.numberFormat, fact.label).toBe(fact.code);
      const fv = shown(model, `B${row}`);
      expect(fv.text, fact.label).toBe(fact.enUS);
      expect(fv.color, fact.label).toBe("color" in fact ? fact.color : undefined);
    }
  });

  it("follows the locale where the code does — and not where an LCID pins the names", async () => {
    const model = await read;
    expect(shown(model, "B2", "de-DE")).toMatchObject({ text: SPEC118_FORMATS_ODS.rows[2].deDE, color: "#ff0000" });
    expect(shown(model, "B3", "de-DE").text).toBe(SPEC118_FORMATS_ODS.rows[3].deDE);
    expect(shown(model, "B4", "fr-FR").text).toBe("Donnerstag, 8. Oktober 2026");
    expect(shown(model, "B14", "de-DE").text).toBe("1899-12-25");
  });
});
