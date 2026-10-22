// SPEC-118 "Implied formats": a value its source marks as a date, a time or a
// percentage but leaves unformatted points at a derived style carrying the
// implied code — one per (base style, code) pair, appended to the style
// table with no source identity — through both readers: an xlsx `t="d"`
// value under a General style (spec118-generated.xlsx, openpyxl-shaped) and
// ODS date, time and percentage cells with no data style (spec118-formats.ods).
import { describe, expect, it } from "vitest";
import { getCell, parseA1 } from "./address";
import { buildSpec118FormatsOdsBytes, buildSpec118GeneratedXlsxBytes, SPEC118_FORMATS_ODS, SPEC118_GENERATED_XLSX } from "./fixtures/build";
import {
  IMPLIED_DATE,
  IMPLIED_DATE_TIME,
  IMPLIED_PERCENTAGE,
  IMPLIED_TIME,
  impliedDateCode,
  ImpliedFormatStyles,
  isGeneralStyle,
} from "./impliedFormats";
import type { CellStyle, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { cellStyle, effectiveStyleId } from "./styles";

/** The effective style of an A1 address of the first sheet. */
function styleAt(model: SpreadsheetModel, a1: string): CellStyle {
  const ref = parseA1(a1)!;
  return cellStyle(model, effectiveStyleId(model.sheets[0]!, ref.row, ref.col));
}

describe("the implied codes", () => {
  it("pins the four codes", () => {
    expect([IMPLIED_DATE, IMPLIED_DATE_TIME, IMPLIED_TIME, IMPLIED_PERCENTAGE]).toEqual(["yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss", "[hh]:mm:ss", "0.00%"]);
  });

  it("a date gets the time only when it has one (to the millisecond), negative serials included", () => {
    expect(impliedDateCode(46303)).toBe("yyyy-mm-dd");
    expect(impliedDateCode(46303.5729166667)).toBe("yyyy-mm-dd hh:mm:ss");
    expect(impliedDateCode(-6)).toBe("yyyy-mm-dd");
    expect(impliedDateCode(-5.75)).toBe("yyyy-mm-dd hh:mm:ss");
    expect(impliedDateCode(46303 + 0.0000000001)).toBe("yyyy-mm-dd"); // rounds to midnight
  });

  it("General is a style with neither numberFormat nor builtinFormatId", () => {
    expect(isGeneralStyle({})).toBe(true);
    expect(isGeneralStyle(undefined)).toBe(true);
    expect(isGeneralStyle({ numberFormat: "0.00" })).toBe(false);
    expect(isGeneralStyle({ builtinFormatId: 14, numberFormat: "m/d/yyyy" })).toBe(false);
  });
});

describe("ImpliedFormatStyles", () => {
  it("appends one derived copy per (base, code) — no source identity — and reuses it", () => {
    const styles: CellStyle[] = [{ sourceXf: 0 }, { font: { bold: true }, alignment: { horizontal: "center" }, sourceXf: 1 }, { numberFormat: "0.0", sourceXf: 2 }];
    const implied = new ImpliedFormatStyles(styles);
    expect(implied.styleFor(1, "yyyy-mm-dd")).toBe(3);
    expect(styles[3]).toEqual({ font: { bold: true }, alignment: { horizontal: "center" }, numberFormat: "yyyy-mm-dd" });
    expect(styles[3]!.font).not.toBe(styles[1]!.font); // a copy, not a shared object
    expect(implied.styleFor(1, "yyyy-mm-dd")).toBe(3);
    expect(implied.styleFor(1, "0.00%")).toBe(4);
    expect(implied.styleFor(0, "yyyy-mm-dd")).toBe(5);
    expect(styles[5]).toEqual({ numberFormat: "yyyy-mm-dd" });
    expect(styles).toHaveLength(6);
  });

  it("leaves a base that already formats its value alone", () => {
    const styles: CellStyle[] = [{}, { numberFormat: "0.0" }];
    expect(new ImpliedFormatStyles(styles).styleFor(1, "yyyy-mm-dd")).toBe(1);
    expect(styles).toHaveLength(2);
  });
});

describe("spec118-generated.xlsx (openpyxl-shaped): t=\"d\" values under a General style", () => {
  it("shows each date through a derived yyyy-mm-dd (+ hh:mm:ss) style, the base's alignment kept", async () => {
    const model = await readSpreadsheetBytes(await buildSpec118GeneratedXlsxBytes(), "xlsx", "generated");
    expect(model.theme).toBeUndefined();
    expect(model.styles.slice(0, 2).every((style) => style.numberFormat === undefined)).toBe(true);
    expect(styleAt(model, "B1").numberFormat).toBe("yyyy-mm-dd hh:mm:ss");
    expect(styleAt(model, "B2").numberFormat).toBe("yyyy-mm-dd");
    expect(styleAt(model, "B3")).toMatchObject({ numberFormat: "yyyy-mm-dd", alignment: { horizontal: "center" } });
    // One derived style per (base, code): B1 (no s → 0) and B2 (s=0) differ by code.
    expect(model.styles).toHaveLength(5);
    expect(model.styles.slice(2).every((style) => style.sourceXf === undefined)).toBe(true);
    // The strings and the values are untouched.
    const sheet = model.sheets[0]!;
    expect(getCell(sheet, 0, 0)?.value).toBe(SPEC118_GENERATED_XLSX.cells.A1);
    expect(getCell(sheet, 1, 1)?.value).toBe(46303);
  });
});

describe("spec118-formats.ods (LibreOffice-shaped): date, time and percentage cells without a data style", () => {
  it("gives each its implied format, and leaves the styled and the era-styled cells to their own", async () => {
    const model = await readSpreadsheetBytes(await buildSpec118FormatsOdsBytes(), "ods", "formats");
    const rows = SPEC118_FORMATS_ODS.rows;
    for (const index of [14, 15, 16, 17] as const) {
      expect(styleAt(model, `B${index}`).numberFormat, rows[index].label).toBe(rows[index].code);
    }
    expect(styleAt(model, "B4").numberFormat).toBe(rows[4].code);
    expect(styleAt(model, "B13").numberFormat).toBeUndefined(); // an era style is General, not "unformatted"
    expect(styleAt(model, "A14").numberFormat).toBeUndefined(); // the string label is not a date
    // 1899-12-25 is an ODS-only date: a serial before 1900's.
    expect(getCell(model.sheets[0]!, 13, 1)?.value).toBe(-6);
  });
});
