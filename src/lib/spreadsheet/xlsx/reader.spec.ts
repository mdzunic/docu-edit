// SPEC-116 "Reading .xlsx": every rule, over the Excel-shaped fixtures and
// synthetic packages — sheets and their kinds, shared strings (rich runs,
// `_xHHHH_` escapes, phonetic runs), every cell type, sequential placement
// without `r`, cached formula values, the date system, the active tab, core
// properties, package variants, out-of-grid addresses and the cell ceiling.
// SPEC-117 changes three of those rules, asserted here in their new form:
// rich runs carry their `<rPr>` as `font`, a `<c>` carrying only a style is
// stored as a style-only cell, and a row with a height is stored even with no
// cells. (The layout, view and style reading itself is in sheetXml.spec.ts
// and styles.spec.ts.) SPEC-119 adds formulas: every "Reading formulas from
// .xlsx" row over spec119-formulas.xlsx and synthetic sheets, and the
// Numbers export's formula reading without formulaUnparsed. SPEC-120 adds the
// calculation settings calcPr gives `calc`.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { getCell, parseA1 } from "../address";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import {
  buildMiniBookXlsxBytes,
  buildSpec1904XlsxBytes,
  buildSpec119FormulasXlsxBytes,
  buildSpec120CalcXlsxBytes,
  buildSpec120FullCalcXlsxBytes,
  buildSpec120ManualXlsxBytes,
  buildSpecMacrosXlsmBytes,
  buildSpecTemplateXltxBytes,
  FIXTURE_WORKBOOK_AUTHOR,
  MINI_BOOK_XLSX,
  SPEC116_1904,
  SPEC119_FORMULAS_XLSX,
} from "../fixtures/build";
import { MAX_WORKBOOK_CELLS, type Cell, type Sheet } from "../model";
import { dateToSerial } from "../serial";
import { decodeXmlEscapes, parseXlsx, readSharedStrings } from "./reader";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const encoder = new TextEncoder();

/** The value at an A1 address of a parsed sheet (undefined = no cell). */
function at(sheet: Sheet, a1: string): unknown {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col)?.value;
}

interface SyntheticSheet {
  name: string;
  /** The worksheet's `<sheetData>` content (omitted for foreign sheets). */
  data?: string;
  state?: string;
  /** Relationship type suffix: "worksheet" (default), "chartsheet", … */
  type?: string;
}

/** A minimal SpreadsheetML package as the parts map the readers accept. */
function synthetic(options: {
  sheets: SyntheticSheet[];
  sharedStrings?: string;
  workbookPr?: string;
  activeTab?: number;
  mainContentType?: string;
}): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>();
  const put = (name: string, text: string) => parts.set(name, encoder.encode(text));
  put(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="${options.mainContentType ?? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}"/></Types>`,
  );
  put(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  const rels = options.sheets.map(
    (sheet, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/${sheet.type ?? "worksheet"}" Target="worksheets/sheet${i + 1}.xml"/>`,
  );
  if (options.sharedStrings !== undefined) {
    rels.push(`<Relationship Id="rIdS" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`);
    put("xl/sharedStrings.xml", `<sst xmlns="${MAIN}">${options.sharedStrings}</sst>`);
  }
  put("xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);
  const sheetsXml = options.sheets
    .map((sheet, i) => `<sheet name="${sheet.name}" sheetId="${i + 1}"${sheet.state ? ` state="${sheet.state}"` : ""} r:id="rId${i + 1}"/>`)
    .join("");
  put(
    "xl/workbook.xml",
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}">${options.workbookPr ?? ""}` +
      `<bookViews><workbookView${options.activeTab !== undefined ? ` activeTab="${options.activeTab}"` : ""}/></bookViews>` +
      `<sheets>${sheetsXml}</sheets></workbook>`,
  );
  options.sheets.forEach((sheet, i) => {
    put(`xl/worksheets/sheet${i + 1}.xml`, `<worksheet xmlns="${MAIN}"><sheetData>${sheet.data ?? ""}</sheetData></worksheet>`);
  });
  return parts;
}

/** Parses one synthetic sheet's data and returns that sheet. */
async function oneSheet(data: string, sharedStrings?: string, workbookPr?: string): Promise<Sheet> {
  const parsed = await parseXlsx(new Uint8Array(0), synthetic({ sheets: [{ name: "S", data }], sharedStrings, workbookPr }));
  return parsed.sheets[0]!;
}

describe("mini-book.xlsx (Excel-shaped)", () => {
  it("reads the sheets in <sheets> order with names, visibility and kind", async () => {
    const parsed = await parseXlsx(await buildMiniBookXlsxBytes());
    expect(parsed.sheets.map((sheet) => sheet.name)).toEqual(MINI_BOOK_XLSX.sheetNames);
    expect(parsed.sheets.map((sheet) => sheet.visibility)).toEqual(["visible", "visible", "hidden"]);
    expect(parsed.sheets.map((sheet) => sheet.kind)).toEqual(["grid", "grid", "grid"]);
    expect(parsed.sheets.map((sheet) => sheet.partName)).toEqual([
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/sheet3.xml",
    ]);
    expect(new Set(parsed.sheets.map((sheet) => sheet.id)).size).toBe(3);
    expect(parsed.activeSheetIndex).toBe(0);
    expect(parsed.dateSystem).toBe("1900");
    expect(parsed.packageVariant).toBeUndefined();
  });

  it("reads every Summary value, formulas contributing their cached <v>", async () => {
    const summary = (await parseXlsx(await buildMiniBookXlsxBytes())).sheets[0]!;
    for (const [a1, expected] of Object.entries(MINI_BOOK_XLSX.summary)) {
      expect(at(summary, a1), a1).toEqual(expected);
    }
    // An absent cell stores nothing.
    expect(getCell(summary, 3, 4)).toBeUndefined(); // E4
  });

  it("keeps a rich shared string's runs, its value their concatenation", async () => {
    const summary = (await parseXlsx(await buildMiniBookXlsxBytes())).sheets[0]!;
    const a1 = getCell(summary, 0, 0)!;
    expect(a1.value).toBe("Q3 report");
    // SPEC-117: each run carries its <rPr> as a font (scheme minor wins over
    // rFont; theme="1" is dk1 by Excel's swap).
    const runFont = { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } };
    expect(a1.runs).toEqual([
      { text: "Q3 ", font: { ...runFont, bold: true } },
      { text: "report", font: runFont },
    ]);
    expect(getCell(summary, 1, 1)!.runs).toBeUndefined();
  });

  it("places Data's rows and cells, written without r, sequentially", async () => {
    const data = (await parseXlsx(await buildMiniBookXlsxBytes())).sheets[1]!;
    expect(data.rows).toHaveLength(MINI_BOOK_XLSX.dataRows);
    expect(data.rows[0]!.cells.map((cell) => cell.value)).toEqual(["ID", "Amount", "Parity"]);
    expect(data.rows.at(-1)!.index).toBe(499);
    expect(data.rows.at(-1)!.cells).toEqual([
      { col: 0, value: 499 },
      { col: 1, value: 748.5 },
      { col: 2, value: "odd" },
    ]);
  });

  it("ignores phonetic runs", async () => {
    const hidden = (await parseXlsx(await buildMiniBookXlsxBytes())).sheets[2]!;
    expect(at(hidden, "A2")).toBe("東京");
  });

  it("reads docProps/core.xml into properties", async () => {
    const parsed = await parseXlsx(await buildMiniBookXlsxBytes());
    expect(parsed.properties.title).toBe("Mini book");
    expect(parsed.properties.author).toBe(FIXTURE_WORKBOOK_AUTHOR);
    expect(parsed.properties.created).toBe("2026-10-01T09:00:00Z");
  });

  it("retains every source part, verbatim and in order", async () => {
    const parsed = await parseXlsx(await buildMiniBookXlsxBytes());
    expect(parsed.opaque.map((part) => part.name)).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/sheet3.xml",
      "xl/theme/theme1.xml",
      "xl/styles.xml",
      "xl/sharedStrings.xml",
      "xl/calcChain.xml",
      "docProps/core.xml",
      "docProps/app.xml",
    ]);
  });
});

describe("the variant fixtures", () => {
  it("reads workbookPr@date1904 as the 1904 system", async () => {
    const parsed = await parseXlsx(await buildSpec1904XlsxBytes());
    expect(parsed.dateSystem).toBe("1904");
    expect(at(parsed.sheets[0]!, "B1")).toBe(SPEC116_1904.dateSerial);
    expect(at(parsed.sheets[0]!, "A1")).toBe("Shipped");
  });

  it("records the package variant a macro-enabled workbook and a template declare", async () => {
    expect((await parseXlsx(await buildSpecMacrosXlsmBytes())).packageVariant).toBe("macroEnabled");
    expect((await parseXlsx(await buildSpecTemplateXltxBytes())).packageVariant).toBe("template");
    const xltm = synthetic({
      sheets: [{ name: "S", data: "" }],
      mainContentType: "application/vnd.ms-excel.template.macroEnabled.main+xml",
    });
    expect((await parseXlsx(new Uint8Array(0), xltm)).packageVariant).toBe("macroEnabledTemplate");
  });
});

describe("sheets, visibility and the active tab", () => {
  it("makes chartsheets, dialogsheets and macrosheets foreign sheets", async () => {
    const parsed = await parseXlsx(
      new Uint8Array(0),
      synthetic({
        sheets: [
          { name: "Grid", data: `<row r="1"><c r="A1"><v>1</v></c></row>` },
          { name: "Chart", type: "chartsheet" },
          { name: "Dialog", type: "dialogsheet" },
          { name: "Macro", type: "xlMacrosheet" },
        ],
      }),
    );
    expect(parsed.sheets.map((sheet) => sheet.kind)).toEqual(["grid", "foreign", "foreign", "foreign"]);
    expect(parsed.sheets[1]!.rows).toEqual([]);
    expect(parsed.sheets[1]!.partName).toBe("xl/worksheets/sheet2.xml");
  });

  it("maps state to visibility", async () => {
    const parsed = await parseXlsx(
      new Uint8Array(0),
      synthetic({ sheets: [{ name: "A" }, { name: "B", state: "hidden" }, { name: "C", state: "veryHidden" }] }),
    );
    expect(parsed.sheets.map((sheet) => sheet.visibility)).toEqual(["visible", "hidden", "veryHidden"]);
  });

  it("reads activeTab, clamped, and moved off a hidden sheet to the first visible one", async () => {
    const sheets = [{ name: "A", state: "hidden" }, { name: "B" }, { name: "C" }];
    expect((await parseXlsx(new Uint8Array(0), synthetic({ sheets, activeTab: 2 }))).activeSheetIndex).toBe(2);
    expect((await parseXlsx(new Uint8Array(0), synthetic({ sheets, activeTab: 9 }))).activeSheetIndex).toBe(2);
    expect((await parseXlsx(new Uint8Array(0), synthetic({ sheets, activeTab: 0 }))).activeSheetIndex).toBe(1);
    expect((await parseXlsx(new Uint8Array(0), synthetic({ sheets }))).activeSheetIndex).toBe(1);
  });
});

describe("shared strings", () => {
  it("decodes _xHHHH_ escapes, _x005F_ being a literal underscore", () => {
    expect(decodeXmlEscapes("a_x000D_b")).toBe("a\rb");
    expect(decodeXmlEscapes("_x0041__x0042_")).toBe("AB");
    expect(decodeXmlEscapes("_x005F_x0041_")).toBe("_x0041_");
    expect(decodeXmlEscapes("_x12_")).toBe("_x12_");
  });

  it("reads plain items, rich runs (their rPr as the run font, SPEC-117) and skips phonetic runs", () => {
    const strings = readSharedStrings(
      `<sst xmlns="${MAIN}"><si><t>plain</t></si><si/><si><r><t>Bo</t></r><r><rPr><b/></rPr><t xml:space="preserve">ld </t></r></si>` +
        `<si><t>漢字</t><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="0"/></si><si><t>a_x005F_x0042_</t></si></sst>`,
    );
    expect(strings).toEqual([
      { text: "plain" },
      { text: "" },
      { text: "Bold ", runs: [{ text: "Bo" }, { text: "ld ", font: { bold: true } }] },
      { text: "漢字" },
      { text: "a_x0042_" },
    ]);
  });
});

describe("cells", () => {
  it("reads every cell type: n, s, str, inlineStr, b, e and d", async () => {
    const sheet = await oneSheet(
      `<row r="1">` +
        `<c r="A1"><v>1.5</v></c><c r="B1" t="n"><v>-2</v></c><c r="C1" t="s"><v>1</v></c>` +
        `<c r="D1" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="E1" t="inlineStr"><is><t>inline</t></is></c>` +
        `<c r="F1" t="inlineStr"><is><r><t>ri</t></r><r><rPr><i/></rPr><t>ch</t></r></is></c>` +
        `<c r="G1" t="b"><v>1</v></c><c r="H1" t="b"><v>0</v></c><c r="I1" t="e"><v>#N/A</v></c>` +
        `<c r="J1" t="e"><v>#WHAT?</v></c><c r="K1" t="d"><v>2026-10-08T12:00:00</v></c><c r="L1" t="e"><v>#SPILL!</v></c>` +
        `</row>`,
      `<si><t>zero</t></si><si><t>one</t></si>`,
    );
    expect(sheet.rows[0]!.cells.map((cell) => cell.value)).toEqual([
      1.5,
      -2,
      "one",
      "xy",
      "inline",
      "rich",
      true,
      false,
      { error: "#N/A" },
      { error: "#VALUE!" },
      dateToSerial(2026, 10, 8, "1900") + 0.5,
      { error: "#SPILL!" },
    ]);
    // SPEC-117: an inline run's <rPr> is its font.
    expect(getCell(sheet, 0, 5)!.runs).toEqual([{ text: "ri" }, { text: "ch", font: { italic: true } }]);
  });

  it("converts t=\"d\" in the workbook's date system", async () => {
    const sheet = await oneSheet(`<row r="1"><c r="A1" t="d"><v>2026-10-08</v></c></row>`, undefined, `<workbookPr date1904="true"/>`);
    expect(at(sheet, "A1")).toBe(dateToSerial(2026, 10, 8, "1904"));
  });

  it("stores no cell for an empty <v>; a style alone is a style-only cell (SPEC-117); a formula without a cached value is stored with value null (SPEC-119)", async () => {
    const sheet = await oneSheet(
      `<row r="1"><c r="A1"><v></v></c><c r="B1" s="3"/><c r="C1"><f>1+1</f></c><c r="D1"><v>4</v></c></row><row r="2" ht="30" customHeight="1"/>`,
    );
    // SPEC-116 stored neither B1 nor row 2; SPEC-117 keeps the style-only
    // cell and the row with a height (and nothing for A1). SPEC-116 and
    // SPEC-117 stored nothing for C1 either; SPEC-119 reads its formula and
    // keeps it with no cached value.
    expect(sheet.rows).toEqual([
      { index: 0, cells: [{ col: 1, value: null, styleId: 3 }, { col: 2, value: null, formula: "1+1" }, { col: 3, value: 4 }] },
      { index: 1, cells: [], heightPt: 30, customHeight: true },
    ]);
  });

  it("places rows and cells without r after the previous one", async () => {
    const sheet = await oneSheet(
      `<row r="3"><c r="B3"><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c><c r="E4"><v>4</v></c><c><v>5</v></c></row>`,
    );
    expect(sheet.rows).toEqual([
      { index: 2, cells: [{ col: 1, value: 1 }, { col: 2, value: 2 }] },
      { index: 3, cells: [{ col: 0, value: 3 }, { col: 4, value: 4 }, { col: 5, value: 5 }] },
    ]);
  });

  it("restores ascending order when a producer writes rows out of order", async () => {
    const sheet = await oneSheet(`<row r="5"><c r="A5"><v>5</v></c></row><row r="2"><c r="B2"><v>2</v></c><c r="A2"><v>1</v></c></row>`);
    expect(sheet.rows.map((row) => row.index)).toEqual([1, 4]);
    expect(sheet.rows[0]!.cells.map((cell) => cell.col)).toEqual([0, 1]);
  });

  it("accepts $ markers in c@r", async () => {
    expect(at(await oneSheet(`<row r="2"><c r="$C$2"><v>9</v></c></row>`), "C2")).toBe(9);
  });
});

/** The error parsing `parts` is refused with (failing if it is not). */
async function failure(parts: Map<string, Uint8Array>): Promise<Error> {
  try {
    await parseXlsx(new Uint8Array(0), parts);
  } catch (error) {
    return error as Error;
  }
  throw new Error("expected the workbook to be refused");
}

describe("refusals", () => {
  it("names an out-of-grid cell's address", async () => {
    const error = await failure(synthetic({ sheets: [{ name: "Wide", data: `<row r="1"><c r="XFE1"><v>1</v></c></row>` }] }));
    expect(error).toBeInstanceOf(SpreadsheetOpenError);
    expect(error.message).toContain("XFE1");
    expect(error.message).toContain("Wide");
  });

  it("names an out-of-grid row", async () => {
    const error = await failure(synthetic({ sheets: [{ name: "Tall", data: `<row r="1048577"><c><v>1</v></c></row>` }] }));
    expect(error.message).toContain("row 1048577");
  });

  it("names the address a run of cells without r walks off the grid at", async () => {
    const error = await failure(synthetic({ sheets: [{ name: "S", data: `<row r="7"><c r="XFD7"><v>1</v></c><c><v>2</v></c></row>` }] }));
    expect(error.message).toContain("XFE7");
  });

  it("refuses a shared-string index that does not exist", async () => {
    const error = await failure(synthetic({ sheets: [{ name: "S", data: `<row r="1"><c r="A1" t="s"><v>4</v></c></row>` }], sharedStrings: "<si><t>a</t></si>" }));
    expect(error.message).toMatch(/shared string 4/);
  });

  it("refuses a worksheet part containing a DTD", async () => {
    const parts = synthetic({ sheets: [{ name: "S", data: "" }] });
    parts.set("xl/worksheets/sheet1.xml", encoder.encode(`<!DOCTYPE x [<!ENTITY e "boom">]><worksheet xmlns="${MAIN}"/>`));
    const error = await failure(parts);
    expect(error).toBeInstanceOf(SpreadsheetOpenError);
    expect(error.message).toMatch(/DTD/);
  });

  it("refuses a package that is not a SpreadsheetML workbook", async () => {
    const parts = synthetic({ sheets: [{ name: "S" }], mainContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" });
    expect(await failure(parts)).toBeInstanceOf(SpreadsheetOpenError);
  });

  it("refuses a 2,000,001-cell workbook at the ceiling, without scanning on", async () => {
    // 2,001 rows of 1,000 cells: one past the ceiling is cell 2,000,001. The
    // part is malformed after that point — a reader that kept scanning (and
    // materialising) would fail on it with a parse error instead.
    const row = `<row>${"<c><v>1</v></c>".repeat(1000)}</row>`;
    const data = row.repeat(Math.ceil((MAX_WORKBOOK_CELLS + 1) / 1000)) + "<row><c><v>1</v></broken>";
    const error = await failure(synthetic({ sheets: [{ name: "Huge", data }] }));
    expect(error).toBeInstanceOf(SpreadsheetOpenError);
    expect(error.message).toBe(WORKBOOK_TOO_LARGE_MESSAGE);
  });

  it("counts the ceiling across every sheet", async () => {
    const row = `<row>${"<c><v>1</v></c>".repeat(1000)}</row>`;
    const half = row.repeat(MAX_WORKBOOK_CELLS / 2000);
    const exactly = synthetic({ sheets: [{ name: "A", data: half }, { name: "B", data: half }] });
    const parsed = await parseXlsx(new Uint8Array(0), exactly);
    expect(parsed.sheets.reduce((sum, sheet) => sum + sheet.rows.reduce((n, r) => n + r.cells.length, 0), 0)).toBe(MAX_WORKBOOK_CELLS);
    const over = synthetic({ sheets: [{ name: "A", data: half }, { name: "B", data: half + "<row><c><v>1</v></c></row>" }] });
    expect((await failure(over)).message).toBe(WORKBOOK_TOO_LARGE_MESSAGE);
  });
});

/** The stored cell at an A1 address. */
function cellAt(sheet: Sheet, a1: string): Cell | undefined {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col);
}

describe("formulas (SPEC-119 \"Reading formulas from .xlsx\")", () => {
  it("reads every formula of spec119-formulas.xlsx: canonical text, cached values as before", async () => {
    const calc = (await parseXlsx(await buildSpec119FormulasXlsxBytes())).sheets[0]!;
    for (const [a1, formula] of Object.entries(SPEC119_FORMULAS_XLSX.formulas)) {
      expect(cellAt(calc, a1)?.formula, a1).toBe(formula);
      expect(cellAt(calc, a1)?.formulaUnparsed, a1).toBeUndefined();
    }
    for (const [a1, value] of Object.entries(SPEC119_FORMULAS_XLSX.values)) {
      expect(cellAt(calc, a1)?.value, a1).toEqual(value);
    }
  });

  it("expands shared formulas per follower — Excel's order, a follower before its master, and an off-grid shift to #REF!", async () => {
    const calc = (await parseXlsx(await buildSpec119FormulasXlsxBytes())).sheets[0]!;
    expect(["C2", "C3", "C9"].map((a1) => cellAt(calc, a1)?.formula)).toEqual(["B2*$F$1", "B3*$F$1", "B9*$F$1"]);
    expect(["D2", "D3", "D4", "D5"].map((a1) => cellAt(calc, a1)?.formula)).toEqual(["A1048574*2", "A1048575*2", "A1048576*2", "#REF!*2"]);
    expect(cellAt(calc, "D5")?.value).toEqual({ error: "#REF!" });
  });

  it("reads t=\"array\" as the anchor's formula + arrayRange — members keep their cached values only — and cm as a dynamic array", async () => {
    const calc = (await parseXlsx(await buildSpec119FormulasXlsxBytes())).sheets[0]!;
    expect(cellAt(calc, "E2")).toEqual({ col: 4, value: 6, formula: "A2:A4*2", arrayRange: SPEC119_FORMULAS_XLSX.arrays.E2 });
    expect(cellAt(calc, "E3")).toEqual({ col: 4, value: 2 });
    expect(cellAt(calc, "H2")).toEqual({ col: 7, value: 1, formula: "SORT(A2:A6)", arrayRange: SPEC119_FORMULAS_XLSX.arrays.H2, dynamicArray: true, cm: 1 });
    expect(cellAt(calc, "H6")).toEqual({ col: 7, value: 5 });
  });

  it("keeps a data table's <f t=\"dataTable\"> element verbatim on its anchor, with no formula", async () => {
    const calc = (await parseXlsx(await buildSpec119FormulasXlsxBytes())).sheets[0]!;
    expect(cellAt(calc, "L3")).toEqual({ col: 11, value: 10, dataTableXml: SPEC119_FORMULAS_XLSX.dataTable.L3 });
    expect(cellAt(calc, "L4")).toEqual({ col: 11, value: 20 });
  });

  it("keeps the immediately-invoked LAMBDA Excel writes as its <f> text, formulaUnparsed", async () => {
    const calc = (await parseXlsx(await buildSpec119FormulasXlsxBytes())).sheets[0]!;
    expect(cellAt(calc, "S2")).toEqual({ col: 18, value: 3, formula: SPEC119_FORMULAS_XLSX.unparsed.S2, formulaUnparsed: true });
  });

  it("reads sheetView@showFormulas=\"1\" as Sheet.showFormulas (the second sheet only)", async () => {
    const sheets = (await parseXlsx(await buildSpec119FormulasXlsxBytes())).sheets;
    expect(sheets.map((sheet) => sheet.name)).toEqual(SPEC119_FORMULAS_XLSX.sheetNames);
    expect(sheets.map((sheet) => sheet.showFormulas)).toEqual([undefined, true, undefined, undefined, undefined]);
    expect(cellAt(sheets[1]!, "C2")).toEqual({ col: 2, value: 2, formula: "B2*2" });
  });

  it("keeps a formula with no <v> as value: null; drops one stray leading =; reads c@cm on any cell", async () => {
    const sheet = await oneSheet(
      `<row r="1"><c r="A1"><f>=SUM(B1:B2)</f></c><c r="B1" cm="2"><v>4</v></c><c r="C1" t="str"><f t="array" ref="C1">_xlfn.CONCAT(A1,"x")</f><v>4x</v></c></row>`,
    );
    expect(sheet.rows[0]!.cells).toEqual([
      { col: 0, value: null, formula: "SUM(B1:B2)" },
      { col: 1, value: 4, cm: 2 },
      { col: 2, value: "4x", formula: 'CONCAT(A1,"x")', arrayRange: { top: 0, left: 2, bottom: 0, right: 2 } },
    ]);
  });

  it("keeps a follower whose si no master defines as its value alone — and stores nothing for an empty one", async () => {
    const sheet = await oneSheet(
      `<row r="1"><c r="A1"><f t="shared" si="7"/><v>3</v></c><c r="B1"><f t="shared" si="8"/></c></row><row r="2"><c r="A2"><f t="shared" si="9"/></c></row>`,
    );
    expect(sheet.rows).toEqual([{ index: 0, cells: [{ col: 0, value: 3 }] }]);
  });

  it("resolves a follower met before its master at the end of the sheet, and parses a shared block once", async () => {
    const sheet = await oneSheet(
      `<row r="1"><c r="B1"><f t="shared" si="0"/><v>2</v></c></row>` +
        `<row r="2"><c r="B2"><f t="shared" ref="B1:B3" si="0">A2*2</f><v>4</v></c></row>` +
        `<row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>`,
    );
    expect(sheet.rows.map((row) => row.cells[0]!.formula)).toEqual(["A1*2", "A2*2", "A3*2"]);
  });

  it("shifts an unparsed master's wire text for its followers, still unparsed", async () => {
    const sheet = await oneSheet(
      `<row r="1"><c r="B1"><f t="shared" ref="B1:B2" si="0">_xlfn.LAMBDA(_xlpm.x,_xlpm.x+A1)(2)</f><v>2</v></c></row>` +
        `<row r="2"><c r="B2"><f t="shared" si="0"/><v>3</v></c></row>`,
    );
    expect(sheet.rows.map((row) => row.cells[0])).toEqual([
      { col: 1, value: 2, formula: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+A1)(2)", formulaUnparsed: true },
      { col: 1, value: 3, formula: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+A2)(2)", formulaUnparsed: true },
    ]);
  });

  it("reads every formula cell of the Numbers export without formulaUnparsed", async () => {
    const bytes = new Uint8Array(readFileSync(join(__dirname, "../../../../e2e/fixtures/spec116-numbers.xlsx")));
    const parsed = await parseXlsx(bytes);
    const formulaCells = parsed.sheets.flatMap((sheet) => sheet.rows.flatMap((row) => row.cells.filter((cell) => cell.formula !== undefined)));
    expect(formulaCells.length).toBeGreaterThan(0);
    expect(formulaCells.filter((cell) => cell.formulaUnparsed)).toEqual([]);
    expect(cellAt(parsed.sheets[1]!, "B5")).toMatchObject({
      value: 850.5,
      formula: "B3-B4",
      arrayRange: { top: 4, left: 1, bottom: 4, right: 1 },
      dynamicArray: true,
      cm: 1,
    });
  });
});

describe('calculation settings (SPEC-120 "Calculation settings")', () => {
  async function calcOf(calcPr: string) {
    return (await parseXlsx(new Uint8Array(0), synthetic({ sheets: [{ name: "S" }], workbookPr: calcPr }))).calc;
  }

  it("reads calcPr@fullCalcOnLoad, @calcMode=\"manual\" and @iterate as xsd:booleans", async () => {
    expect(await calcOf(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`)).toEqual({ fullCalcOnLoad: true });
    expect(await calcOf(`<calcPr calcId="191029" fullCalcOnLoad="true"/>`)).toEqual({ fullCalcOnLoad: true });
    expect(await calcOf(`<calcPr calcId="191029" calcMode="manual"/>`)).toEqual({ manual: true });
    expect(await calcOf(`<calcPr calcId="191029" iterate="1" iterateCount="100"/>`)).toEqual({ iterate: true });
    expect(await calcOf(`<calcPr fullCalcOnLoad="1" calcMode="manual" iterate="true"/>`)).toEqual({ fullCalcOnLoad: true, manual: true, iterate: true });
  });

  it("leaves calc out when calcPr is absent or says nothing it models — autoNoTable and auto are automatic", async () => {
    expect(await calcOf("")).toBeUndefined();
    expect(await calcOf(`<calcPr calcId="191029"/>`)).toBeUndefined();
    expect(await calcOf(`<calcPr calcId="191029" calcMode="autoNoTable" fullCalcOnLoad="0"/>`)).toBeUndefined();
    expect(await calcOf(`<calcPr calcMode="auto" iterate="0"/>`)).toBeUndefined();
  });

  it("reads the SPEC-120 fixtures: spec120-calc.xlsx none, spec120-fullcalc.xlsx fullCalcOnLoad, spec120-manual.xlsx manual", async () => {
    expect((await parseXlsx(await buildSpec120CalcXlsxBytes())).calc).toBeUndefined();
    expect((await parseXlsx(await buildSpec120FullCalcXlsxBytes())).calc).toEqual({ fullCalcOnLoad: true });
    expect((await parseXlsx(await buildSpec120ManualXlsxBytes())).calc).toEqual({ manual: true });
  });
});
