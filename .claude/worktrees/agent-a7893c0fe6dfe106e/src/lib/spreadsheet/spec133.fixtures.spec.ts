// SPEC-133's names-and-tables fixtures (spec133-names-tables.xlsx and
// spec133-names.ods, built by fixtures/build.ts): package-level facts only —
// each builder is deterministic, the committed e2e file is its bytes, the
// package is whole with exactly the listed parts, every XML part parses, the
// producer shapes the name and table readers are proven on are really in the
// parts, and every cached value is what its formula computes over the
// recorded cells — worked out by hand here, since SPEC-120's evaluator
// resolves neither names nor structured references until SPEC-133. Nothing
// here asserts how the app reads them.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { parseXml } from "../core/xml";
import { unzip } from "../core/zip";
import { formatA1, parseA1, parseRange } from "./address";
import { buildSpec133NamesOdsBytes, buildSpec133NamesTablesXlsxBytes, SPEC133_NAMES_ODS, SPEC133_NAMES_TABLES_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { MAX_COLS, MAX_ROWS } from "./model";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REVISION_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const REVISION3_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";
const TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const decoder = new TextDecoder();
const committed = (name: string): Uint8Array => new Uint8Array(readFileSync(join(E2E_FIXTURES, name)));
const textOf = (parts: Map<string, Uint8Array>, name: string): string => decoder.decode(parts.get(name));
const xmlOf = (parts: Map<string, Uint8Array>, name: string): Element => parseXml(textOf(parts, name)).documentElement;
const elements = (root: Element, ns: string, local: string): Element[] => Array.from(root.getElementsByTagNameNS(ns, local));
const children = (parent: Element): Element[] => Array.from(parent.children);
/** An element's attributes as written, name → value. */
const attrs = (element: Element): Record<string, string> => Object.fromEntries(Array.from(element.attributes).map((attr) => [attr.name, attr.value]));
/** Excel's case-insensitive name order. */
const byLowerName = (a: string, b: string): number => (a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0);

/** A relative cell reference moved by (dRow, dCol), wrapping round the grid
 * as Excel and Calc do for names. */
function shifted(a1: string, dRow: number, dCol: number): string {
  const ref = parseA1(a1)!;
  return formatA1({ row: (((ref.row + dRow) % MAX_ROWS) + MAX_ROWS) % MAX_ROWS, col: (((ref.col + dCol) % MAX_COLS) + MAX_COLS) % MAX_COLS });
}

describe("spec133-names-tables.xlsx", () => {
  const F = SPEC133_NAMES_TABLES_XLSX;
  const { Sales, Table1 } = F.tables;
  const parts = async (): Promise<Map<string, Uint8Array>> => unzip(committed(F.file));

  it("is deterministic, the committed bytes, whole, with exactly the listed parts, every XML part parsing", async () => {
    const built = await buildSpec133NamesTablesXlsxBytes();
    expect(equalBytes(built, await buildSpec133NamesTablesXlsxBytes())).toBe(true);
    expect(equalBytes(committed(F.file), built)).toBe(true);
    expect(await packageIntegrity(built)).toEqual([]);
    const all = await parts();
    expect([...all.keys()]).toEqual([...F.parts]);
    for (const [name] of all) {
      expect(() => parseXml(textOf(all, name)), name).not.toThrow();
    }
  });

  it("lists Data, Lists and the hidden Hidden, then the definedNames in Excel's order, scoped by sheet position", async () => {
    const workbook = textOf(await parts(), "xl/workbook.xml");
    expect(workbook).toContain(
      `<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Lists" sheetId="2" r:id="rId2"/>` +
        `<sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/></sheets><definedNames>`,
    );
    expect(workbook).not.toContain("activeTab");
    expect(workbook).toMatch(/<\/definedNames><calcPr calcId="191029"\/><\/workbook>$/);
    // Excel's attribute order (CT_DefinedName: name, comment, localSheetId, hidden).
    expect(workbook).toContain(`<definedName name="_xlfn.IFS" hidden="1">#NAME?</definedName>`);
    expect(workbook).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Lists!$D$1:$E$4</definedName>`);
    expect(workbook).toContain(`<definedName name="Local" comment="Codes on Lists" localSheetId="1">Lists!$A$2:$A$4</definedName>`);
    const written = elements(parseXml(workbook).documentElement, MAIN_NS, "definedName").map((element) => {
      const { name, localSheetId, hidden, comment } = attrs(element);
      return {
        name: name!,
        ...(comment === undefined ? {} : { comment }),
        ...(localSheetId === undefined ? {} : { localSheetId: Number(localSheetId), scope: F.sheetNames[Number(localSheetId)] }),
        ...(hidden === "1" ? { hidden: true } : {}),
        text: element.textContent,
      };
    });
    expect(written).toEqual(F.definedNames);
    expect(F.definedNames.map((name) => name.name)).toEqual(F.definedNames.map((name) => name.name).toSorted(byLowerName));
    const visible = F.definedNames.filter((name) => !name.hidden && !/^_xl(nm|fn)\./.test(name.name)).map((name) => name.name);
    expect(visible).toEqual([...F.visibleNames]);
  });

  it("app.xml: Excel's Worksheets and Named Ranges heading pairs, hidden names left out of TitlesOfParts", async () => {
    const app = textOf(await parts(), "docProps/app.xml");
    expect(app).toContain(
      `<vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant>` +
        `<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>5</vt:i4></vt:variant></vt:vector>`,
    );
    const titles = [...app.matchAll(/<vt:lpstr>([^<]*)<\/vt:lpstr>/g)].map((match) => match[1]).slice(2);
    expect(titles).toEqual([...F.titlesOfParts]);
    const named = F.definedNames.filter((name) => !name.hidden).map((name) => `${name.scope ? `${name.scope}!` : ""}${name.name.replace(/^_xlnm\./, "")}`);
    expect(titles).toEqual([...F.sheetNames, ...named]);
  });

  it("Sales: totals row, its autoFilter over header + data, xr:uid/xr3:uids, the calculated column, dataDxfId, the style", async () => {
    const all = await parts();
    const text = textOf(all, Sales.partName);
    expect(text).toContain(`mc:Ignorable="xr xr3"`);
    const root = parseXml(text).documentElement;
    expect(root.getAttribute("id")).toBe(String(Sales.id));
    expect(root.getAttributeNS(REVISION_NS, "uid")).toBe(Sales.uid);
    expect([root.getAttribute("name"), root.getAttribute("displayName"), root.getAttribute("ref")]).toEqual([Sales.name, Sales.displayName, Sales.ref]);
    expect(root.getAttribute("totalsRowCount")).toBe("1");
    expect(root.hasAttribute("headerRowCount")).toBe(false);
    expect(text).toContain(`>${Sales.autoFilter.xml}<tableColumns count="5">`);
    expect(parseRange(Sales.ref)).toEqual(Sales.range);
    expect(parseRange(Sales.autoFilter.ref)).toEqual({ ...Sales.range, bottom: Sales.range.bottom - 1 });
    const columns = elements(root, MAIN_NS, "tableColumn").map((column) => {
      const { id, name, totalsRowFunction, totalsRowLabel, dataDxfId } = attrs(column);
      const formula = elements(column, MAIN_NS, "calculatedColumnFormula")[0]?.textContent;
      return {
        id: Number(id),
        uid: column.getAttributeNS(REVISION3_NS, "uid"),
        name,
        ...(totalsRowLabel === undefined ? {} : { totalsRowLabel }),
        ...(totalsRowFunction === undefined ? {} : { totalsRowFunction }),
        ...(dataDxfId === undefined ? {} : { dataDxfId: Number(dataDxfId) }),
        ...(formula === undefined ? {} : { calculatedColumnFormula: formula }),
      };
    });
    expect(columns).toEqual(Sales.columns.map(({ calculatedFormula: _canonical, ...column }) => column));
    expect(text).toContain(
      `<tableColumn id="4" xr3:uid="${Sales.columns[3]!.uid}" name="Amount" totalsRowFunction="sum" dataDxfId="0">` +
        `<calculatedColumnFormula>Sales[[#This Row],[Price]]*Sales[[#This Row],[Qty]]</calculatedColumnFormula></tableColumn>`,
    );
    expect(text).toMatch(/<\/tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"\/><\/table>$/);
    // dataDxfId 0 is styles.xml's one dxf — the #,##0 Amount's data cells carry.
    const styles = textOf(all, "xl/styles.xml");
    expect(styles).toContain(`<dxfs count="1"><dxf><numFmt numFmtId="${F.dxf.numFmtId}" formatCode="${F.dxf.formatCode}"/></dxf></dxfs>`);
    expect(elements(xmlOf(all, "xl/styles.xml"), MAIN_NS, "cellXfs")[0]!.children[F.xf.amount]!.getAttribute("numFmtId")).toBe(String(F.dxf.numFmtId));
  });

  it("Table1 on Lists has no header row (so no autoFilter) and TableStyleLight9; Lists' own autoFilter is what _FilterDatabase names", async () => {
    const all = await parts();
    const text = textOf(all, Table1.partName);
    const root = parseXml(text).documentElement;
    expect(attrs(root)).toMatchObject({ id: String(Table1.id), name: "Table1", displayName: "Table1", ref: Table1.ref, headerRowCount: "0", totalsRowShown: "0" });
    expect(root.getAttributeNS(REVISION_NS, "uid")).toBe(Table1.uid);
    expect(text).not.toContain("<autoFilter");
    expect(elements(root, MAIN_NS, "tableColumn").map((column) => [Number(column.getAttribute("id")), column.getAttributeNS(REVISION3_NS, "uid"), column.getAttribute("name")])).toEqual(
      Table1.columns.map((column) => [column.id, column.uid, column.name]),
    );
    expect(text).toContain(`<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>`);
    expect(parseRange(Table1.ref)).toEqual(Table1.range);
    const lists = textOf(all, F.sheets.Lists.partName);
    expect(lists).toContain(`</sheetData>${F.listsAutoFilter.xml}<pageMargins `);
    const filterDatabase = F.definedNames.find((name) => name.name === "_xlnm._FilterDatabase")!;
    expect(filterDatabase.text.replaceAll("$", "")).toBe(`Lists!${F.listsAutoFilter.ref}`);
  });

  it("each worksheet relates its table, tableParts last (CT_Worksheet order), with the table content type", async () => {
    const all = await parts();
    const types = textOf(all, "[Content_Types].xml");
    for (const table of [Sales, Table1]) {
      const sheet = F.sheets[table.sheet];
      expect(textOf(all, sheet.partName), table.name).toMatch(
        new RegExp(`<pageMargins [^>]*/><tableParts count="1"><tablePart r:id="${table.rId}"/></tableParts></worksheet>$`),
      );
      expect(textOf(all, sheet.rels)).toContain(
        `<Relationship Id="${table.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../${table.partName.slice(3)}"/>`,
      );
      expect(types).toContain(`<Override PartName="/${table.partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
    }
    expect(textOf(all, F.sheets.Hidden.partName)).not.toContain("tableParts");
    expect(textOf(all, F.sheets.Data.partName)).toContain(`mc:Ignorable="x14ac xr xr2 xr3"`);
  });

  it("holds the recorded cells: shared strings, numbers, Excel's wire formulas with their caches, the #,##0 xf on Amount's data cells", async () => {
    const all = await parts();
    const strings = elements(xmlOf(all, "xl/sharedStrings.xml"), MAIN_NS, "si").map((si) => si.textContent);
    for (const name of F.sheetNames) {
      const sheet = F.sheets[name];
      const cells: Record<string, string | number> = {};
      const wires: Record<string, string> = {};
      const styled: string[] = [];
      for (const c of elements(xmlOf(all, sheet.partName), MAIN_NS, "c")) {
        const ref = c.getAttribute("r")!;
        const v = elements(c, MAIN_NS, "v")[0]!.textContent!;
        cells[ref] = c.getAttribute("t") === "s" ? strings[Number(v)]! : Number(v);
        const f = elements(c, MAIN_NS, "f")[0];
        if (f) {
          wires[ref] = f.textContent!;
        }
        if (c.hasAttribute("s")) {
          styled.push(`${ref}:${c.getAttribute("s")}`);
        }
      }
      expect(cells, name).toEqual(F.values[name]);
      expect(wires, name).toEqual(name === "Data" ? Object.fromEntries(Object.entries(F.formulas).map(([ref, formula]) => [ref, formula.wire])) : {});
      expect(styled, name).toEqual(name === "Data" ? F.amountCells.map((ref) => `${ref}:${F.xf.amount}`) : []);
    }
    expect(textOf(all, "xl/calcChain.xml")).toContain(`<c r="${F.calcChain[0]}" i="1" l="1"/><c r="${F.calcChain[1]}"/>`);
    expect([...textOf(all, "xl/calcChain.xml").matchAll(/<c r="([A-Z]+\d+)"/g)].map((match) => match[1])).toEqual(Object.keys(F.formulas));
    expect([...F.calcChain]).toEqual(Object.keys(F.formulas));
  });

  it("every cached value is what Excel computes over the recorded cells", () => {
    const V: Readonly<Record<string, string | number>> = F.values.Data;
    const n = (ref: string): number => V[ref] as number;
    const rows = [2, 3, 4, 5];
    for (const [ref, formula] of Object.entries(F.formulas)) {
      expect(formula.value, ref).toBe(V[ref]);
      // Canonical = wire with Excel's [#This Row] spelled `@`; every other structured reference verbatim.
      expect(formula.formula, ref).toBe(formula.wire.replaceAll(/\[\[#This Row\],\[([^\]]+)\]\]/g, "[@$1]"));
    }
    expect(Sales.columns[3]!.calculatedFormula).toBe(F.formulas.D2.formula);
    for (const r of rows) {
      expect(n(`D${r}`), `D${r}`).toBe(n(`B${r}`) * n(`C${r}`));
    }
    const amounts = rows.reduce((sum, r) => sum + n(`D${r}`), 0);
    expect(n("D6")).toBe(amounts); // SUBTOTAL(109, …): sum
    expect(n("E6")).toBe(rows.filter((r) => V[`E${r}`] !== undefined).length); // SUBTOTAL(103, …): count
    expect(n("G2")).toBe(amounts);
    expect(n("G3")).toBe(n("D6")); // the totals cell
    expect(n("G4")).toBe(n("E4")); // implicit intersection with G4's row
    const taxRate = Number(F.definedNames.find((name) => name.name === "TaxRate")!.text);
    expect(n("G5")).toBe(taxRate * n("G2"));
    // Above (Data!A1048576) is relative to A1: from G6 it names G5.
    const above = parseA1(F.definedNames.find((name) => name.name === "Above")!.text.split("!")[1]!)!;
    const { host, refersTo } = F.relativeName;
    expect(shifted(host, above.row, above.col)).toBe(refersTo);
    expect(n(host)).toBe(n(refersTo));
    // Column F is empty, so the table can grow right.
    expect(Object.keys(V).filter((ref) => ref.startsWith("F"))).toEqual([]);
  });
});

describe("spec133-names.ods", () => {
  const O = SPEC133_NAMES_ODS;
  const parts = async (): Promise<Map<string, Uint8Array>> => unzip(committed(O.file));

  it("is deterministic, the committed bytes, whole, with exactly the listed parts, every XML part parsing", async () => {
    const built = await buildSpec133NamesOdsBytes();
    expect(equalBytes(built, await buildSpec133NamesOdsBytes())).toBe(true);
    expect(equalBytes(committed(O.file), built)).toBe(true);
    expect(await packageIntegrity(built)).toEqual([]);
    const all = await parts();
    expect([...all.keys()]).toEqual([...O.parts]);
    for (const [name] of all) {
      if (name.endsWith(".xml") || name.endsWith(".rdf")) {
        expect(() => parseXml(textOf(all, name)), name).not.toThrow();
      }
    }
    expect(textOf(all, "meta.xml")).toContain(`meta:table-count="2" meta:cell-count="24" meta:object-count="0"`);
    expect(textOf(all, "meta.xml")).toContain(`<meta:generator>LibreOffice/24.8`);
  });

  it("Sheet1's own names are its table's last child; after the last table the workbook's names, then the database ranges", async () => {
    const content = textOf(await parts(), "content.xml");
    expect(content).toContain(`${O.localNamedExpressionsXml}</table:table><table:table table:name="Sheet2"`);
    expect(content).toContain(`</table:table>${O.namedExpressionsXml}${O.databaseRangesXml}</office:spreadsheet>`);
    const spreadsheet = elements(parseXml(content).documentElement, OFFICE_NS, "spreadsheet")[0]!;
    expect(children(spreadsheet).map((child) => child.localName)).toEqual(["calculation-settings", "table", "table", "named-expressions", "database-ranges"]);
    const [sheet1, sheet2] = children(spreadsheet).filter((child) => child.localName === "table");
    expect(sheet1!.lastElementChild!.localName).toBe("named-expressions");
    expect(elements(sheet2!, TABLE_NS, "named-expressions")).toEqual([]);
    // Calc writes each block in upper-cased name order, ranges and expressions interleaved.
    for (const names of [O.localNames, O.globalNames]) {
      expect(names.map((name) => name.name)).toEqual(names.map((name) => name.name).toSorted((a, b) => (a.toUpperCase() < b.toUpperCase() ? -1 : 1)));
    }
    expect(O.globalNames.map((name) => name.kind)).toEqual(["named-range", "named-range", "named-expression", "named-range"]);
    expect(content).toContain(`<table:named-range table:name="Report" table:base-cell-address="$Sheet1.$A$1" table:cell-range-address="$Sheet1.$A$1:.$C$5" table:range-usable-as="print-range filter"/>`);
    expect(O.databaseRanges.map((range) => range.name)).toEqual(["__Anonymous_Sheet_DB__0", "Inventory"]);
    expect(O.databaseRanges[0].xml).toContain(`table:display-filter-buttons="true"`);
    expect(O.databaseRanges[1].xml).toContain(`table:contains-header="true"`);
    for (const range of O.databaseRanges) {
      const [from, to] = range.targetRangeAddress.split(":").map((end) => end.split(".")[1]!);
      expect(parseRange(`${from}:${to}`), range.name).toEqual(range.range);
    }
  });

  it("holds the recorded cells and formulas with their caches, then Calc's trailing runs", async () => {
    const content = textOf(await parts(), "content.xml");
    for (const name of O.tableNames) {
      const table = elements(parseXml(content).documentElement, TABLE_NS, "table").find((element) => element.getAttributeNS(TABLE_NS, "name") === name)!;
      const cells: Record<string, string | number> = {};
      const wires: Record<string, string> = {};
      const rows = elements(table, TABLE_NS, "table-row");
      rows.slice(0, O.rowsWritten[name]).forEach((row, r) => {
        let col = 0;
        for (const cell of elements(row, TABLE_NS, "table-cell")) {
          const ref = formatA1({ row: r, col });
          const type = cell.getAttributeNS(OFFICE_NS, "value-type");
          if (type !== null) {
            cells[ref] = type === "float" ? Number(cell.getAttributeNS(OFFICE_NS, "value")) : cell.textContent!;
          }
          if (cell.hasAttributeNS(TABLE_NS, "formula")) {
            wires[ref] = cell.getAttributeNS(TABLE_NS, "formula")!;
          }
          col += Number(cell.getAttributeNS(TABLE_NS, "number-columns-repeated") ?? 1);
        }
        expect(col, `${name} row ${r + 1}`).toBe(MAX_COLS);
      });
      expect(cells, name).toEqual(O.values[name]);
      expect(wires, name).toEqual(Object.fromEntries(Object.entries(O.formulas[name]).map(([ref, formula]) => [ref, formula.wire])));
      expect(rows).toHaveLength(O.rowsWritten[name] + 1);
      expect(rows.at(-1)!.getAttributeNS(TABLE_NS, "number-rows-repeated")).toBe(String(MAX_ROWS - O.rowsWritten[name]));
    }
  });

  it("every cached value is what Calc computes, and each name's A1-rebased formula is its base-relative text moved to A1", () => {
    const S1: Readonly<Record<string, string | number>> = O.values.Sheet1;
    const S2: Readonly<Record<string, string | number>> = O.values.Sheet2;
    const n = (ref: string): number => S1[ref] as number;
    const names = new Map([...O.localNames, ...O.globalNames].map((name) => [name.name, name]));
    // LocalExpr's [.B2] is relative to its base C2: from a host, the cell one column left.
    const base = parseA1(names.get("LocalExpr")!.baseCellAddress.split(".")[1]!)!;
    const left = parseA1(/\[\.([A-Z]+\d+)\]/.exec(names.get("LocalExpr")!.expression!)![1]!)!;
    for (const host of ["C2", "C3"] as const) {
      const at = parseA1(host)!;
      expect(O.formulas.Sheet1[host].value, host).toBe(n(formatA1({ row: at.row + left.row - base.row, col: at.col + left.col - base.col })) * 2);
    }
    expect(O.formulas.Sheet1.C4.value).toBe(n("B3") + n("B4") + n("B5")); // LocalRange = $B$3:$B$5
    const prices = n("A2") + n("A3") + n("A4"); // Prices = $A$2:$A$4
    expect(O.formulas.Sheet1.C5.value).toBe(prices * 0.15); // Rate = 0.15
    expect(O.formulas.Sheet2.B6.value).toBe(prices);
    for (const sheet of ["Sheet1", "Sheet2"] as const) {
      for (const [ref, formula] of Object.entries(O.formulas[sheet])) {
        expect(formula.value, `${sheet}!${ref}`).toBe(O.values[sheet][ref as keyof (typeof O.values)[typeof sheet]]);
        expect(formula.wire, `${sheet}!${ref}`).toBe(`of:=${formula.formula}`);
      }
    }
    // Rebased to A1 with the grid's wrap: LocalExpr's B2 from C2 → XFD1; Next's B3 from B2 → A2.
    expect(names.get("LocalExpr")!.formula).toBe(`${shifted(formatA1(left), -base.row, -base.col)}*2`);
    const next = names.get("Next")!;
    const nextBase = parseA1(next.baseCellAddress.split(".")[1]!.replaceAll("$", ""))!;
    expect(next.formula).toBe(`Sheet2!${shifted(next.cellRangeAddress!.split(".")[1]!, -nextBase.row, -nextBase.col)}`);
    expect(S2.B2).toBe(12);
    // Absolute ranges and expressions do not move: their formula is the address in canonical form.
    for (const name of ["LocalRange", "Prices", "Report"]) {
      const entry = names.get(name)!;
      const [start, end] = entry.cellRangeAddress!.split(":");
      expect(entry.formula, name).toBe(`${start!.slice(1).replace(".", "!")}:${end!.slice(1)}`);
    }
    expect(names.get("Rate")!.formula).toBe(names.get("Rate")!.expression!.slice("of:=".length));
  });
});
