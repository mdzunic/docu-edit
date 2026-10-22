// SPEC-123's fixtures (spec123-styles.xlsx / spec123-styles.ods, built by
// fixtures/build.ts — the "Fixtures" AC): the committed e2e files are the
// builders' bytes, both packages are whole, each holds every construct the
// AC names in its producer's own shape, and what SPEC123_STYLES_XLSX and
// SPEC123_STYLES_ODS record is what the readers give — values, formulas, the
// styled row's Row.styleId, column H's span (xlsx) / default cell style
// (ods), and the formatting each documented position shows. The named-style
// read itself (SpreadsheetModel.namedStyles, CellStyle.namedStyle) is
// asserted in xlsx/styles.spec.ts and ods/styles.spec.ts, beside the readers
// that learn it.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { FO_NAMESPACE, NUMBER_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TABLE_NAMESPACE } from "../core/odfXml";
import { parseXml } from "../core/xml";
import { unzip } from "../core/zip";
import { formatA1, getCell, parseA1 } from "./address";
import { buildSpec123StylesOdsBytes, buildSpec123StylesXlsxBytes, SPEC123_STYLES_ODS, SPEC123_STYLES_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { recalculate } from "./formula/recalc";
import type { CellStyle, CellValue, Sheet, SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { formatValue, workbookFormatContext } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { cellStyle, effectiveStyleId } from "./styles";
import { tintTransforms } from "./xlsx/themeColor";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REVISION_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const decoder = new TextDecoder();

/** SPEC-123's StyleAspect order (styles.ts STYLE_ASPECTS) and the xlsx
 * `applyX` attribute that excludes each from a named style. */
const ASPECT_APPLY = [
  ["numberFormat", "applyNumberFormat"],
  ["font", "applyFont"],
  ["fill", "applyFill"],
  ["borders", "applyBorder"],
  ["alignment", "applyAlignment"],
  ["protection", "applyProtection"],
] as const;

const styleName = (el: Element): string | null => el.getAttributeNS(STYLE_NAMESPACE, "name");
const tableAttr = (el: Element, local: string): string | null => el.getAttributeNS(TABLE_NAMESPACE, local);
const all = (root: Element | Document, ns: string, local: string): Element[] => Array.from(root.getElementsByTagNameNS(ns, local));
const one = (root: Element | Document, ns: string, local: string): Element => {
  const found = all(root, ns, local);
  expect(found, `${local} elements`).toHaveLength(1);
  return found[0]!;
};

function sheetNamed(model: SpreadsheetModel, name: string): Sheet {
  const sheet = model.sheets.find((candidate) => candidate.name === name);
  expect(sheet, name).toBeDefined();
  return sheet!;
}

/** Every stored cell of a sheet as A1 → value. */
function storedValues(sheet: Sheet): Record<string, CellValue> {
  return Object.fromEntries(sheet.rows.flatMap((row) => row.cells.map((cell) => [formatA1({ row: row.index, col: cell.col }), cell.value])));
}

/** Every stored cell's own styleId, as A1 → id (cells without one omitted). */
function storedStyleIds(sheet: Sheet): Record<string, number> {
  return Object.fromEntries(
    sheet.rows.flatMap((row) => row.cells.flatMap((cell) => (cell.styleId === undefined ? [] : [[formatA1({ row: row.index, col: cell.col }), cell.styleId]]))),
  );
}

function effectiveAt(sheet: Sheet, a1: string): number {
  const { row, col } = parseA1(a1)!;
  return effectiveStyleId(sheet, row, col);
}

function styleAt(model: SpreadsheetModel, sheet: Sheet, a1: string): CellStyle {
  return cellStyle(model, effectiveAt(sheet, a1));
}

/** The documented `looks` keys of what a position shows (en-US). */
function observedLook(model: SpreadsheetModel, sheet: Sheet, a1: string, keys: readonly string[]): Record<string, unknown> {
  const style = styleAt(model, sheet, a1);
  const { row, col } = parseA1(a1)!;
  const value = getCell(sheet, row, col)?.value ?? null;
  const observed: Record<string, unknown> = {
    bold: style.font?.bold,
    italic: style.font?.italic,
    sizePt: style.font?.sizePt,
    color: style.font?.color,
    fill: style.fill?.kind === "solid" ? style.fill.color : undefined,
    horizontal: style.alignment?.horizontal,
    numberFormat: style.numberFormat,
    text: formatValue(value, style.numberFormat, workbookFormatContext(model, SPREADSHEET_LOCALES["en-US"])).text,
    topEdge: style.borders?.top?.style,
    bottomEdge: style.borders?.bottom?.style,
    borderColor: style.borders?.bottom?.color,
  };
  return Object.fromEntries(keys.map((key) => [key, observed[key]]));
}

describe("spec123-styles.xlsx (Excel-shaped)", () => {
  const F = SPEC123_STYLES_XLSX;
  const B = F.budget;
  const bytes = buildSpec123StylesXlsxBytes();
  const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, "xlsx", "spec123-styles");
  const part = async (name: string): Promise<string> => decoder.decode((await unzip(await bytes)).get(name));

  it("is the committed e2e fixture, byte for byte, and a whole package in Excel's part order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec123-styles.xlsx")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(F.parts);
    expect(Object.values(F.partNames).every((name) => (F.parts as readonly string[]).includes(name))).toBe(true);
    expect(F.parts).not.toContain("xl/calcChain.xml");
  });

  it("holds the named styles: cellStyleXfs with Excel's applyX exclusions, cellStyles with builtinIds, a hidden and a custom one", async () => {
    const styles = parseXml(await part(F.partNames.styles));
    const styleXfs = all(one(styles, MAIN_NS, "cellStyleXfs"), MAIN_NS, "xf");
    const cellStyles = all(one(styles, MAIN_NS, "cellStyles"), MAIN_NS, "cellStyle");
    expect(styleXfs).toHaveLength(F.namedStyles.length);
    expect(one(styles, MAIN_NS, "cellStyleXfs").getAttribute("count")).toBe(String(styleXfs.length));
    expect(one(styles, MAIN_NS, "cellStyles").getAttribute("count")).toBe(String(cellStyles.length));
    expect(
      cellStyles.map((el) => ({
        name: el.getAttribute("name"),
        xfId: Number(el.getAttribute("xfId")),
        ...(el.hasAttribute("builtinId") ? { builtinId: Number(el.getAttribute("builtinId")) } : {}),
        ...(el.getAttribute("hidden") === "1" ? { hidden: true } : {}),
        ...(el.hasAttributeNS(REVISION_NS, "uid") ? { uid: el.getAttributeNS(REVISION_NS, "uid") } : {}),
        // An absent applyX includes the aspect; applyX="0" excludes it.
        includes: ASPECT_APPLY.filter(([, apply]) => styleXfs[Number(el.getAttribute("xfId"))]!.getAttribute(apply) !== "0").map(([aspect]) => aspect),
      })),
    ).toEqual(F.namedStyles);
    // Excel's own order: alphabetical, Normal (xfId 0) not first.
    const names = cellStyles.map((el) => el.getAttribute("name")!);
    expect(names).toEqual(names.toSorted());
    expect(cellStyles.find((el) => el.getAttribute("name") === "Invoice total")!.hasAttribute("builtinId")).toBe(false);
  });

  it("holds cellXfs naming their styles by xfId, applyX only where a cell differs, tinted theme colours, numFmt 164, the dxf and the extLst", async () => {
    const text = await part(F.partNames.styles);
    const styles = parseXml(text);
    const cellXfs = all(one(styles, MAIN_NS, "cellXfs"), MAIN_NS, "xf");
    expect(cellXfs.map((xf) => Number(xf.getAttribute("xfId")))).toEqual(F.xfIds);
    expect(cellXfs.map((xf) => xf.getAttributeNames().filter((name) => name.startsWith("apply")))).toEqual([
      [],
      ["applyNumberFormat"],
      [],
      [],
      [],
      ["applyFont"],
      ["applyNumberFormat"],
      ["applyAlignment"],
    ]);
    expect(cellXfs[F.xf.goodCentred]!.getElementsByTagNameNS(MAIN_NS, "alignment")[0]?.getAttribute("horizontal")).toBe("center");
    expect(cellXfs.map((xf) => xf.getAttribute("numFmtId"))).toEqual(["0", "2", "0", "0", "164", "0", "10", "0"]);
    const numFmt = one(one(styles, MAIN_NS, "numFmts"), MAIN_NS, "numFmt");
    expect([numFmt.getAttribute("numFmtId"), numFmt.getAttribute("formatCode")]).toEqual(["164", F.numFmts[164]]);
    expect(text).toContain(`<color theme="4" tint="${F.tints.columnHFont}"/>`);
    expect(text).toContain(`<fgColor theme="4" tint="${F.tints.totalFill}"/><bgColor indexed="64"/>`);
    // Good's fill as Excel writes the style's: no bgColor.
    expect(text).toContain(`<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>`);
    expect(all(one(styles, MAIN_NS, "dxfs"), MAIN_NS, "dxf")).toHaveLength(1);
    expect(text).toContain(`<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst>`);
    expect(text).toContain(`<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>`);
    expect(text).toContain(`<x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>`);
    expect(text).toMatch(/<\/extLst><\/styleSheet>$/);
  });

  it("writes Budget as Excel does: the <col style>, the <row s customFormat>, H6 where they meet, no F1 and no row 5", async () => {
    const sheet = await part(F.partNames.budget);
    const H = B.styledColumn;
    expect(sheet).toContain(`<dimension ref="${B.dimension}"/>`);
    expect(sheet).toContain(`<cols><col min="${H.index + 1}" max="${H.index + 1}" width="${H.sourceWidthChars}" style="${H.styleId}" customWidth="1"/></cols>`);
    expect(sheet).toContain(`<row r="6" spans="${B.spans}" s="${B.styledRow.styleId}" customFormat="1" x14ac:dyDescent="0.35">`);
    expect(sheet).toContain(`<c r="H6" s="${B.styledRow.styleId}"/></row>`);
    expect([...sheet.matchAll(/<row r="(\d+)" spans="([^"]+)"/g)].map((match) => [match[1], match[2]])).toEqual(
      ["1", "2", "3", "4", "6"].map((r) => [r, B.spans]),
    );
    expect(sheet).not.toContain(`r="${B.emptyCell}"`);
    expect(sheet).toContain(`<row r="6" `);
    expect(sheet).not.toContain(`<row r="5" `);
  });

  it("reads Budget (first, active) and Notes: every stored value, formula and style id", async () => {
    const model = await read();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(F.sheetNames);
    expect(model.sheets.map((sheet) => sheet.partName)).toEqual(F.sheets.map((sheet) => sheet.partName));
    expect(model.activeSheetIndex).toBe(F.activeSheetIndex);
    const budget = sheetNamed(model, "Budget");
    expect(storedValues(budget)).toEqual({ ...B.values, ...Object.fromEntries(B.styleOnly.map((ref) => [ref, null])) });
    expect(storedStyleIds(budget)).toEqual(B.styles);
    expect(Object.fromEntries(budget.rows.flatMap((row) => row.cells.flatMap((cell) => (cell.formula ? [[formatA1({ row: row.index, col: cell.col }), cell.formula]] : []))))).toEqual(
      B.formulas,
    );
    expect(Math.max(...budget.rows.map((row) => row.index)) + 1).toBe(B.usedRows);
    expect(budget.rows.some((row) => row.index === B.emptyRow.index)).toBe(false);
    expect(budget.selection?.active).toEqual({ row: 0, col: 0 });
    expect(storedValues(sheetNamed(model, "Notes"))).toEqual(F.notes.values);
  });

  it("reads the styled row's Row.styleId, column H's span, and H6 where they meet", async () => {
    const budget = sheetNamed(await read(), "Budget");
    const H = B.styledColumn;
    const R = B.styledRow;
    expect(budget.rows.filter((row) => row.styleId !== undefined).map((row) => [row.index, row.styleId])).toEqual([[R.index, R.styleId]]);
    expect(budget.rows.find((row) => row.index === R.index)?.sourceAttrs).toBe(`spans="${B.spans}" x14ac:dyDescent="0.35"`);
    expect(budget.columns).toHaveLength(1);
    expect(budget.columns[0]).toMatchObject({ min: H.index, max: H.index, styleId: H.styleId, sourceWidthChars: H.sourceWidthChars });
    // Resolution cell → row → column → 0 (SPEC-117).
    expect(["H1", "H2", "H5", "H7", "H1048576"].map((a1) => effectiveAt(budget, a1))).toEqual(Array(5).fill(H.styleId));
    expect(["A6", "D6", "F6", "G6", "H6", "XFD6"].map((a1) => effectiveAt(budget, a1))).toEqual(Array(6).fill(R.styleId));
    expect(["F1", "E5", "A7", "G1"].map((a1) => effectiveAt(budget, a1))).toEqual(Array(4).fill(F.xf.normal));
  });

  it("shows the documented formatting: Heading 1, 0.00, 0.00%, Good, Good centred, the italic column and the Invoice total row", async () => {
    const model = await read();
    const budget = sheetNamed(model, "Budget");
    for (const [a1, look] of Object.entries(B.looks)) {
      expect(observedLook(model, budget, a1, Object.keys(look)), a1).toEqual(look);
    }
    // The theme colours: Excel's index swap (3 = dk2, 4 = accent1) and tints as lumMod/lumOff.
    expect(styleAt(model, budget, "A1")).toMatchObject({ font: { color: { theme: "dk2" } }, borders: { bottom: { style: "thick", color: { theme: "accent1" } } } });
    expect(styleAt(model, budget, "H1").font?.color).toEqual({ theme: "accent1", ...tintTransforms(Number(F.tints.columnHFont)), xlsxTint: F.tints.columnHFont });
    const total = styleAt(model, budget, "B6");
    expect(total.fill).toEqual({ kind: "solid", color: { theme: "accent1", ...tintTransforms(Number(F.tints.totalFill)), xlsxTint: F.tints.totalFill } });
    expect(total.borders).toEqual({ top: { style: "thin", color: { theme: "accent1" } }, bottom: { style: "double", color: { theme: "accent1" } } });
    expect(styleAt(model, budget, "G6")).toEqual(total);
    expect(styleAt(model, budget, "F3").font?.color).toBe("#006100");
  });

  it("caches what its formulas compute: a full recalculation changes nothing", async () => {
    const model = await read();
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.circular).toEqual([]);
  });
});

describe("spec123-styles.ods (LibreOffice-shaped)", () => {
  const O = SPEC123_STYLES_ODS;
  const B = O.budget;
  const bytes = buildSpec123StylesOdsBytes();
  const read = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await bytes, "ods", "spec123-styles");
  const part = async (name: string): Promise<string> => decoder.decode((await unzip(await bytes)).get(name));

  it("is the committed e2e fixture, byte for byte, and a whole package in Calc's entry order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec123-styles.ods")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(O.parts);
  });

  it("holds the common styles: Default, Heading, Heading_20_1, Good and Invoice_20_total with display names and parents", async () => {
    const styles = parseXml(await part("styles.xml"));
    const common = all(one(styles, OFFICE_NAMESPACE, "styles"), STYLE_NAMESPACE, "style").filter((el) => el.getAttributeNS(STYLE_NAMESPACE, "family") === "table-cell");
    const byName = new Map(common.map((el) => [styleName(el)!, el]));
    /** The aspects a common style sets, itself or through its parents
     * (Default sets none), in STYLE_ASPECTS order. */
    const includes = (name: string): string[] => {
      const set = new Set<string>();
      for (let el = byName.get(name); el && styleName(el) !== "Default"; el = byName.get(el.getAttributeNS(STYLE_NAMESPACE, "parent-style-name") ?? "")) {
        if (el.hasAttributeNS(STYLE_NAMESPACE, "data-style-name")) {
          set.add("numberFormat");
        }
        if (el.getElementsByTagNameNS(STYLE_NAMESPACE, "text-properties").length > 0) {
          set.add("font");
        }
        for (const props of Array.from(el.getElementsByTagNameNS(STYLE_NAMESPACE, "table-cell-properties"))) {
          if (props.hasAttributeNS(FO_NAMESPACE, "background-color")) {
            set.add("fill");
          }
          if (props.getAttributeNames().some((attr) => attr.startsWith("fo:border"))) {
            set.add("borders");
          }
        }
        if (el.getElementsByTagNameNS(STYLE_NAMESPACE, "paragraph-properties").length > 0) {
          set.add("alignment");
        }
      }
      return ASPECT_APPLY.map(([aspect]) => aspect).filter((aspect) => set.has(aspect));
    };
    expect(
      common.map((el) => {
        const name = styleName(el)!;
        const displayName = el.getAttributeNS(STYLE_NAMESPACE, "display-name");
        const parent = el.getAttributeNS(STYLE_NAMESPACE, "parent-style-name");
        return { name, ...(displayName ? { displayName } : {}), ...(parent ? { parent, includes: includes(name) } : {}) };
      }),
    ).toEqual(O.commonStyles);
  });

  it("holds Calc's automatic styles: ce styles parented to the common ones, N1xx data styles with the style:map pair", async () => {
    const content = parseXml(await part("content.xml"));
    const automatic = one(content, OFFICE_NAMESPACE, "automatic-styles");
    const cellStyles = all(automatic, STYLE_NAMESPACE, "style").filter((el) => el.getAttributeNS(STYLE_NAMESPACE, "family") === "table-cell");
    expect(
      Object.fromEntries(
        cellStyles.map((el) => {
          const dataStyle = el.getAttributeNS(STYLE_NAMESPACE, "data-style-name");
          const align = el.getElementsByTagNameNS(STYLE_NAMESPACE, "paragraph-properties")[0]?.getAttributeNS(FO_NAMESPACE, "text-align");
          const italic = el.getElementsByTagNameNS(STYLE_NAMESPACE, "text-properties")[0]?.getAttributeNS(FO_NAMESPACE, "font-style") === "italic";
          return [
            styleName(el),
            {
              parent: el.getAttributeNS(STYLE_NAMESPACE, "parent-style-name"),
              ...(dataStyle ? { dataStyle } : {}),
              ...(align ? { horizontal: align } : {}),
              ...(italic ? { italic } : {}),
            },
          ];
        }),
      ),
    ).toEqual(O.cellStyles);
    const dataStyles = Array.from(automatic.children).filter((el) => el.namespaceURI === NUMBER_NAMESPACE);
    expect(dataStyles.map(styleName)).toEqual(["N104", "N105P0", "N105", "N111"]);
    const map = one(dataStyles.find((el) => styleName(el) === "N105")!, STYLE_NAMESPACE, "map");
    expect([map.getAttributeNS(STYLE_NAMESPACE, "condition"), map.getAttributeNS(STYLE_NAMESPACE, "apply-style-name")]).toEqual(["value()>=0", "N105P0"]);
  });

  it("writes Budget as Calc does: column H's default cell style, row 6 formatted to XFD by a styled run, the trailing run", async () => {
    const content = parseXml(await part("content.xml"));
    const budget = all(content, TABLE_NAMESPACE, "table").find((el) => tableAttr(el, "name") === "Budget")!;
    const H = B.columnDefault;
    expect(
      all(budget, TABLE_NAMESPACE, "table-column").map((el) => [tableAttr(el, "style-name"), tableAttr(el, "number-columns-repeated"), tableAttr(el, "default-cell-style-name")]),
    ).toEqual([
      ["co1", String(H.index), "Default"],
      [H.columnStyle, null, H.styleName],
      ["co1", String(16_384 - H.index - 1), "Default"],
    ]);
    const rows = all(budget, TABLE_NAMESPACE, "table-row");
    expect(rows).toHaveLength(B.usedRows + 1);
    const styledRow = rows[B.styledRow.index]!;
    const run = Array.from(styledRow.children).at(-1)!;
    expect([tableAttr(run, "style-name"), tableAttr(run, "number-columns-repeated")]).toEqual([B.styledRow.styleName, String(B.styledRow.repeatedRun)]);
    expect(Array.from(rows[B.emptyRow.index]!.children).map((cell) => tableAttr(cell, "number-columns-repeated"))).toEqual(["16384"]);
    expect(tableAttr(rows.at(-1)!, "number-rows-repeated")).toBe(String(B.trailingRowRun));
    expect(B.usedRows + B.trailingRowRun).toBe(1_048_576);
  });

  it("reads Budget (first, active) and Notes: every stored value and formula; F1, row 5 and H6 store nothing", async () => {
    const model = await read();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(O.tableNames);
    expect(model.activeSheetIndex).toBe(O.activeSheetIndex);
    const budget = sheetNamed(model, "Budget");
    expect(storedValues(budget)).toEqual(B.values);
    for (const [a1, fact] of Object.entries(B.formulas)) {
      const { row, col } = parseA1(a1)!;
      expect(getCell(budget, row, col), a1).toMatchObject({ formula: fact.formula, value: fact.value });
    }
    expect(Object.fromEntries(Object.entries(storedStyleIds(budget)).map(([a1, id]) => [a1, cellStyle(model, id).sourceStyleName]))).toEqual(B.styles);
    expect(budget.rows.some((row) => row.index === B.emptyRow.index)).toBe(false);
    expect(Math.max(...budget.rows.map((row) => row.index)) + 1).toBe(B.usedRows);
    expect(storedValues(sheetNamed(model, "Notes"))).toEqual(O.notes.values);
  });

  it("reads the styled run as row 6's Row.styleId and column H's default as its span's styleId", async () => {
    const model = await read();
    const budget = sheetNamed(model, "Budget");
    const H = B.columnDefault;
    const R = B.styledRow;
    const styled = budget.rows.filter((row) => row.styleId !== undefined);
    expect(styled.map((row) => row.index)).toEqual([R.index]);
    expect(cellStyle(model, styled[0]!.styleId!).sourceStyleName).toBe(R.styleName);
    expect(budget.columns).toHaveLength(1);
    expect(budget.columns[0]).toMatchObject({ min: H.index, max: H.index, sourceStyleName: H.columnStyle });
    expect(cellStyle(model, budget.columns[0]!.styleId!).sourceStyleName).toBe(H.styleName);
    const named = (a1: string): string | undefined => styleAt(model, budget, a1).sourceStyleName;
    expect(["H1", "H5", "H7", "H1048576"].map(named)).toEqual(Array(4).fill(H.styleName));
    expect(["A6", "D6", "G6", "H6", "XFD6"].map(named)).toEqual(Array(5).fill(R.styleName));
    expect(["F1", "E5", "A7"].map(named)).toEqual(Array(3).fill("Default"));
  });

  it("shows the documented formatting: Heading 1, 0.00, 0.00%, Good, Good centred, the italic column and the Invoice total row", async () => {
    const model = await read();
    const budget = sheetNamed(model, "Budget");
    for (const [a1, look] of Object.entries(B.looks)) {
      expect(observedLook(model, budget, a1, Object.keys(look)), a1).toEqual(look);
    }
    for (const [name, code] of Object.entries(O.dataStyles)) {
      const ce = Object.entries(O.cellStyles).find(([, style]) => "dataStyle" in style && style.dataStyle === name)![0];
      expect(model.styles.find((style) => style.sourceStyleName === ce)?.numberFormat, name).toBe(code);
    }
    expect(styleAt(model, budget, "B6").borders).toEqual({ top: { style: "thin", color: "#156082" }, bottom: { style: "double", color: "#156082" } });
  });

  it("caches what its formulas compute: a full recalculation changes nothing", async () => {
    const model = await read();
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.circular).toEqual([]);
  });
});
