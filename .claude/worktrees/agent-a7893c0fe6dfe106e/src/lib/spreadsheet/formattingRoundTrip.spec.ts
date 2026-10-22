// SPEC-123 round trips: spec123-styles.xlsx and .ods formatted per the e2e
// flow (B2:C3 bold + Accent 1 lighter 40% + All borders + Currency, D2 0.0%
// then "Good", B2's pattern painted onto E2:E4, column F bold) are written and
// re-read. Relations: field-level toEqual on the re-read resolved styles
// (sourceXf/sourceStyleName and the xlsx tint's source text ignored — where a
// style came from, not how it looks; an ods code compared by what it renders,
// since a data style reads `"$"` back bare) and each cell's namedStyle;
// part-level equalBytes for every untouched part; and every source
// numFmts/fonts/fills/borders/cellStyleXfs/cellXfs/cellStyles entry keeping
// its index and bytes (append-only).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import type { GridRange } from "./address";
import { buildSpec123StylesOdsBytes, buildSpec123StylesXlsxBytes } from "./fixtures/build";
import { applyNamedStyle, applyPattern, applyStylePatch, applyStylePatches, borderPresetPatches, capturePattern, presetNumberFormat } from "./formatting";
import { MAX_ROWS, type CellStyle, type SpreadsheetFormat, type SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { formatValue } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { cellStyle, effectiveStyleId } from "./styles";
import { writeOds } from "./ods/writer";
import { writeXlsx } from "./xlsx/writer";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const at = (top: number, left: number, bottom = top, right = left): GridRange => ({ top, left, bottom, right });
const REFS = ["B2", "C2", "B3", "C3", "D2", "E2", "E3", "E4", "F2", "F3", "F4", "F6", "F9", "A6", "H1"];

function formatLikeTheFlow(model: SpreadsheetModel): void {
  const id = model.sheets[0]!.id;
  const b2c3 = [at(1, 1, 2, 2)];
  applyStylePatch(model, id, b2c3, { font: { bold: true } });
  applyStylePatch(model, id, b2c3, { fill: { kind: "solid", color: { theme: "accent1", lumMod: 0.6, lumOff: 0.4 } } });
  applyStylePatches(model, id, borderPresetPatches(b2c3[0]!, "all", { style: "thin" }));
  const currency = presetNumberFormat("currency", EN_US);
  applyStylePatch(model, id, b2c3, { numberFormat: currency.numberFormat, builtinFormatId: null });
  applyStylePatch(model, id, [at(1, 3)], { numberFormat: "0.0%" });
  applyNamedStyle(model, id, [at(1, 3)], "Good");
  applyPattern(model, id, at(1, 4, 3, 4), capturePattern(model.sheets[0]!, at(1, 1))!);
  applyStylePatch(model, id, [{ top: 0, left: 5, bottom: MAX_ROWS - 1, right: 5 }], { font: { bold: true } });
}

/** A position's resolved style without provenance (and, for ods, its code
 * as what it renders). */
function resolved(model: SpreadsheetModel, ref: string, format: SpreadsheetFormat): CellStyle {
  const [, letters, digits] = /^([A-Z])(\d+)$/.exec(ref)!;
  const style = structuredClone(cellStyle(model, effectiveStyleId(model.sheets[0]!, Number(digits) - 1, letters!.charCodeAt(0) - 65)));
  delete style.sourceXf;
  delete style.sourceStyleName;
  const json = JSON.parse(JSON.stringify(style, (key, value) => (key === "xlsxTint" ? undefined : value))) as CellStyle;
  if (format === "ods" && json.numberFormat !== undefined) {
    json.numberFormat = formatValue(-1234.5, formatCodeOf(json, EN_US), { locale: EN_US, dateSystem: "1900" }).text;
    delete json.builtinFormatId;
  }
  return json;
}

/** Each collection's inner XML in a styles part. */
function collection(xml: string, name: string): string {
  return new RegExp(`<${name}[^>]*>([\\s\\S]*?)</${name}>`).exec(xml)?.[1] ?? "";
}

describe("SPEC-123 formatting round trips", () => {
  it("xlsx: every changed style field and namedStyle survives; untouched parts and source style entries keep their bytes", async () => {
    const bytes = await buildSpec123StylesXlsxBytes();
    const model = await readSpreadsheetBytes(bytes, "xlsx", "spec123");
    formatLikeTheFlow(model);
    const written = await writeXlsx(model);
    const reread = await readSpreadsheetBytes(written, "xlsx", "spec123");
    for (const ref of REFS) {
      expect(resolved(reread, ref, "xlsx"), ref).toEqual(resolved(model, ref, "xlsx"));
    }
    expect(resolved(reread, "D2", "xlsx").namedStyle).toBe("Good");
    expect(reread.namedStyles?.map((style) => style.name)).toEqual(model.namedStyles?.map((style) => style.name));
    const source = await unzip(bytes);
    const out = await unzip(written);
    for (const [name, data] of source) {
      if (name !== "xl/styles.xml" && name !== "xl/worksheets/sheet1.xml" && name !== "xl/sharedStrings.xml" && name !== "xl/workbook.xml" && name !== "docProps/app.xml" && name !== "[Content_Types].xml" && name !== "xl/_rels/workbook.xml.rels") {
        expect(equalBytes(out.get(name)!, data), name).toBe(true);
      }
    }
    expect(equalBytes(out.get("xl/worksheets/sheet2.xml")!, source.get("xl/worksheets/sheet2.xml")!)).toBe(true);
    const before = new TextDecoder().decode(source.get("xl/styles.xml")!);
    const after = new TextDecoder().decode(out.get("xl/styles.xml")!);
    for (const name of ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles"]) {
      expect(collection(after, name).startsWith(collection(before, name)), name).toBe(true);
    }
    for (const kept of ["<dxfs", "<x14:slicerStyles", "<x15:timelineStyles"]) {
      expect(after.slice(after.indexOf(kept))).toBe(before.slice(before.indexOf(kept)));
    }
  });

  it("ods: every changed style field and namedStyle survives; styles.xml is untouched", async () => {
    const bytes = await buildSpec123StylesOdsBytes();
    const model = await readSpreadsheetBytes(bytes, "ods", "spec123");
    formatLikeTheFlow(model);
    const written = await writeOds(model);
    const reread = await readSpreadsheetBytes(written, "ods", "spec123");
    for (const ref of REFS) {
      const expected = resolved(model, ref, "ods");
      // ods has no theme: a theme colour is written resolved.
      const actual = resolved(reread, ref, "ods");
      expect(actual.fill !== undefined, ref).toBe(expected.fill !== undefined);
      expect(actual.font?.bold, ref).toBe(expected.font?.bold);
      expect(actual.numberFormat, ref).toBe(expected.numberFormat);
      expect(actual.namedStyle, ref).toBe(expected.namedStyle);
      expect(Object.keys(actual.borders ?? {}).toSorted(), ref).toEqual(Object.keys(expected.borders ?? {}).toSorted());
    }
    expect(resolved(reread, "D2", "ods").namedStyle).toBe("Good");
    const source = await unzip(bytes);
    const out = await unzip(written);
    expect(equalBytes(out.get("styles.xml")!, source.get("styles.xml")!)).toBe(true);
    expect(equalBytes(out.get("meta.xml")!, source.get("meta.xml")!)).toBe(true);
  });
});
