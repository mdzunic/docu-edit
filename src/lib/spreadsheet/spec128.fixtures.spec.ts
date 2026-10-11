// SPEC-128's sort-and-filter fixtures (spec128-filter.xlsx / spec128-filter.ods,
// built by fixtures/build.ts): the committed e2e files are the builders'
// bytes, both packages are whole, and the producer shapes the readers and
// writers are proven on are really in the parts — Excel's `xr:uid`s,
// `sheetPr filterMode`, the `autoFilter` with a `filters blank="1"` list, a
// `customFilters` pair and its `sortState`, a `colorFilter` with its dxf, a
// worksheet `sortState` by cell colour, the hidden `_xlnm._FilterDatabase`
// names, `x14ac:dyDescent` rows and rows without `r`; LibreOffice's
// `__Anonymous_Sheet_DB__0` with `table:filter-set-item`s, `filter` and
// `collapse` row visibility and the trailing repeated rows. Deliberately
// byte- and XML-level only: what the readers make of these files is
// asserted by the reader, filter and round-trip specs. The facts constants
// are checked against the parts and against themselves (the criteria really
// reject exactly the rows recorded as filter-hidden, so "8 of 12" holds).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { scanXml } from "../core/xmlScanner";
import { unzip } from "../core/zip";
import { buildSpec128FilterOdsBytes, buildSpec128FilterXlsxBytes, SPEC128_FILTER_ODS, SPEC128_FILTER_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { dateToSerial } from "./serial";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();

/** One element of a part: its qualified name and attributes, and its
 * element children (text joined into `text`). */
interface Node {
  name: string;
  attrs: Record<string, string>;
  children: Node[];
  text: string;
}

/** A small tree of the part, built from the scanner's events. */
function tree(xml: string): Node {
  const root: Node = { name: "#document", attrs: {}, children: [], text: "" };
  const stack = [root];
  scanXml(xml, (event) => {
    const top = stack[stack.length - 1]!;
    if (event.type === "open") {
      const node: Node = { name: event.name, attrs: Object.fromEntries(event.attrs.map((attr) => [attr.name, attr.value])), children: [], text: "" };
      top.children.push(node);
      if (!event.selfClosing) {
        stack.push(node);
      }
    } else if (event.type === "close") {
      stack.pop();
    } else {
      top.text += event.text;
    }
  });
  return root.children[0]!;
}

function child(node: Node, name: string): Node {
  const found = node.children.find((candidate) => candidate.name === name);
  expect(found, `${node.name} > ${name}`).toBeDefined();
  return found!;
}

const childNames = (node: Node): string[] => node.children.map((candidate) => candidate.name);

/** Where `needle` sits in `haystack`, asserting it is there exactly once. */
function onlyAt(haystack: string, needle: string): number {
  const at = haystack.indexOf(needle);
  expect(at, needle).toBeGreaterThanOrEqual(0);
  expect(haystack.indexOf(needle, at + 1), `${needle} occurs once`).toBe(-1);
  return at;
}

/** The facts' own rule: the Region list (or a blank) AND 100 <= Amount <= 900. */
function salesRowPasses(row: (typeof SPEC128_FILTER_XLSX.sales.rows)[number]): boolean {
  const A = SPEC128_FILTER_XLSX.sales.autoFilter;
  const regionPasses = row.region === null ? A.region.blanks : (A.region.values as readonly string[]).includes(row.region);
  return regionPasses && row.amount >= Number(A.amount.conditions[0].value) && row.amount <= Number(A.amount.conditions[1].value);
}

describe("SPEC128 facts", () => {
  const X = SPEC128_FILTER_XLSX.sales;

  it("records exactly the rows the criteria reject as filter-hidden, and one hidden row they all pass", () => {
    expect(X.rows.map((row) => row.row)).toEqual([2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]);
    expect(X.rows.filter((row) => !salesRowPasses(row)).map((row) => row.row)).toEqual(X.filterHiddenRows);
    expect(X.rows.filter((row) => row.state === "filter").map((row) => row.row)).toEqual(X.filterHiddenRows);
    expect(X.rows.filter((row) => row.state === "user").map((row) => row.row)).toEqual([X.userHiddenRow]);
    expect(salesRowPasses(X.rows.find((row) => row.row === X.userHiddenRow)!)).toBe(true);
    expect(X.filterHiddenIndices).toEqual(X.filterHiddenRows.map((row) => row - 1));
    expect(X.userHiddenIndex).toBe(X.userHiddenRow - 1);
    // Both South rows fail Region, the 50 and the 950 fail Amount; the blank Region passes.
    expect(X.rows.filter((row) => row.region === "South").map((row) => row.row)).toEqual([4, 9]);
    expect(X.rows.filter((row) => row.region === null).map((row) => [row.row, row.state])).toEqual([[7, "shown"]]);
  });

  it("says 8 of 12 records, and 6 once West is unchecked", () => {
    const shown = X.records - X.filterHiddenRows.length;
    expect(X.status).toBe(`${shown} of ${X.records} records found`);
    const west = X.rows.filter((row) => row.region === "West" && row.state !== "filter").length;
    expect(X.statusWithoutWest).toBe(`${shown - west} of ${X.records} records found`);
    expect(X.records).toBe(X.autoFilter.range.bottom - X.autoFilter.range.top);
  });

  it("lists the visible Amounts as opened, after Clear filters and after Sort Z to A", () => {
    expect(X.visibleAmounts).toEqual(X.rows.filter((row) => row.state === "shown").map((row) => row.amount));
    expect(X.visibleAmountsCleared).toEqual(X.rows.filter((row) => row.state !== "user").map((row) => row.amount));
    expect(X.visibleAmountsDescending).toEqual(X.visibleAmountsCleared.toSorted((a, b) => b - a));
  });

  it("holds the records in Date order, as the sort record says, each serial its date's", () => {
    for (const row of X.rows) {
      expect(row.serial, `row ${row.row}`).toBe(dateToSerial(row.date[0], row.date[1], row.date[2], "1900"));
    }
    const serials = X.rows.map((row) => row.serial);
    expect(serials).toEqual(serials.toSorted((a, b) => a - b));
    expect(X.autoFilter.sort.keys).toEqual([{ col: 3, descending: false }]);
  });

  it("gives the ODS twin the same records and the same hidden rows", () => {
    const O = SPEC128_FILTER_ODS;
    expect(O.rows).toBe(X.rows);
    expect(O.headers).toEqual(X.headers);
    expect(O.filterRows).toEqual(X.filterHiddenRows);
    expect(O.filterIndices).toEqual(X.filterHiddenIndices);
    expect(O.collapseRow).toBe(X.userHiddenRow);
    expect(O.collapseIndex).toBe(X.userHiddenIndex);
    expect(O.status).toBe(X.status);
    expect(O.databaseRange.range).toEqual(X.autoFilter.range);
    expect(O.databaseRange.region.values).toEqual(X.autoFilter.region.values);
    expect(O.trailingRowRun + 13).toBe(1_048_576);
  });
});

describe("spec128-filter.xlsx (Excel-shaped)", () => {
  const F = SPEC128_FILTER_XLSX;
  const bytes = buildSpec128FilterXlsxBytes();
  const part = async (name: string): Promise<string> => decoder.decode((await unzip(await bytes)).get(name));
  const sst = async (): Promise<string[]> =>
    tree(await part("xl/sharedStrings.xml"))
      .children.filter((item) => item.name === "si")
      .map((item) => child(item, "t").text);

  it("is the committed e2e fixture, byte for byte, and a whole package in Excel's part order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec128-filter.xlsx")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()]).toEqual(F.parts);
  });

  it("names the three sheets in order, Sales the selected tab, and a hidden _xlnm._FilterDatabase per autofilter", async () => {
    const workbook = await part("xl/workbook.xml");
    const root = tree(workbook);
    expect(child(root, "sheets").children.map((sheet) => sheet.attrs.name)).toEqual(F.sheetNames);
    expect(child(root, "definedNames").children.map((name) => ({ ...name.attrs, text: name.text }))).toEqual(
      F.definedNames.map((name) => ({ name: name.name, localSheetId: String(name.localSheetId), hidden: "1", text: name.text })),
    );
    // Excel's place for them: after </sheets>, before <calcPr>.
    const names = onlyAt(workbook, "<definedNames>");
    expect(workbook.indexOf("</sheets>")).toBeLessThan(names);
    expect(workbook.indexOf("<calcPr")).toBeGreaterThan(names);
    const selected = await Promise.all([F.sales.partName, F.colours.partName, F.log.partName].map(async (name) => (await part(name)).includes(`tabSelected="1"`)));
    expect(selected).toEqual([true, false, false]);
  });

  it("writes Excel 365's worksheet roots: mc:Ignorable, xr:uid, the Calibri-era sheetFormatPr", async () => {
    for (const name of [F.sales.partName, F.colours.partName, F.log.partName]) {
      const xml = await part(name);
      const root = tree(xml);
      expect(root.attrs["mc:Ignorable"], name).toBe("x14ac xr xr2 xr3");
      expect(root.attrs["xr:uid"], name).toMatch(/^\{[0-9A-F-]{36}\}$/);
      expect(child(root, "sheetFormatPr").attrs, name).toEqual({ defaultRowHeight: "15", "x14ac:dyDescent": "0.25" });
    }
  });

  describe("Sales", () => {
    const S = F.sales;

    it("carries sheetPr filterMode before the dimension, and its autoFilter verbatim after sheetData", async () => {
      const xml = await part(S.partName);
      const root = tree(xml);
      expect(childNames(root)).toEqual(["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "sheetData", "autoFilter", "pageMargins"]);
      expect(child(root, "sheetPr").attrs).toEqual({ filterMode: "1" });
      expect(child(root, "dimension").attrs.ref).toBe(S.dimension);
      const at = onlyAt(xml, S.autoFilter.xml);
      expect(xml.indexOf("</sheetData>")).toBeLessThan(at);
      expect(xml.indexOf("<pageMargins")).toBeGreaterThan(at);
    });

    it("filters Region by a list with blank=\"1\" and Amount by an and-ed customFilters pair, sorted by Date", async () => {
      const autoFilter = child(tree(await part(S.partName)), "autoFilter");
      expect(autoFilter.attrs).toEqual({ ref: S.autoFilter.ref, "xr:uid": S.autoFilter.uid });
      expect(childNames(autoFilter)).toEqual(["filterColumn", "filterColumn", "sortState"]);
      const [region, amount, sortState] = autoFilter.children as [Node, Node, Node];
      expect(region.attrs).toEqual({ colId: String(S.autoFilter.region.colId) });
      const filters = child(region, "filters");
      expect(filters.attrs).toEqual({ blank: "1" });
      expect(filters.children.map((filter) => filter.attrs.val)).toEqual(S.autoFilter.region.values);
      expect(amount.attrs).toEqual({ colId: String(S.autoFilter.amount.colId) });
      const custom = child(amount, "customFilters");
      expect(custom.attrs).toEqual({ and: "1" });
      expect(custom.children.map((condition) => ({ operator: condition.attrs.operator, value: condition.attrs.val }))).toEqual(S.autoFilter.amount.conditions);
      expect(sortState.attrs).toEqual({ "xmlns:xlrd2": "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2", ref: S.autoFilter.sort.ref });
      expect(sortState.children.map((condition) => condition.attrs)).toEqual([{ ref: S.autoFilter.sort.conditionRef }]);
    });

    it("writes every row with spans and x14ac:dyDescent, half of them (and their cells) without r", async () => {
      const rows = child(tree(await part(S.partName)), "sheetData").children;
      expect(rows).toHaveLength(13);
      rows.forEach((row, index) => {
        const r = index + 1; // consecutive from row 1, so a row without r lands here
        const bare = (S.rowsWithoutR as readonly number[]).includes(r);
        expect(row.attrs.r, `row ${r}`).toBe(bare ? undefined : String(r));
        expect(row.attrs.spans, `row ${r}`).toBe("1:4");
        expect(row.attrs["x14ac:dyDescent"], `row ${r}`).toBe("0.25");
        for (const cell of row.children) {
          expect(cell.attrs.r === undefined, `row ${r} cell refs`).toBe(bare);
        }
      });
      expect(rows.filter((row) => row.attrs.r === undefined)).toHaveLength(6);
      // Hidden rows, with and without r, are hidden="1" whichever kind they are.
      expect(rows.flatMap((row, index) => (row.attrs.hidden === "1" ? [index + 1] : []))).toEqual([...S.filterHiddenRows, S.userHiddenRow].toSorted((a, b) => a - b));
      // Row 7 has no A cell — its first cell is B7.
      expect(rows[6]!.children.map((cell) => cell.attrs.r)).toEqual(["B7", "C7", "D7"]);
    });

    it("holds the recorded values: bold shared-string headers, the records, dates on numFmt 14", async () => {
      const strings = await sst();
      const rows = child(tree(await part(S.partName)), "sheetData").children;
      const header = rows[0]!.children;
      expect(header.map((cell) => [cell.attrs.r, cell.attrs.s, cell.attrs.t, strings[Number(child(cell, "v").text)]])).toEqual(
        Object.entries(S.headers).map(([ref, text]) => [ref, String(F.xf.header), "s", text]),
      );
      S.rows.forEach((record) => {
        const cells = rows[record.row - 1]!.children;
        const values = cells.map((cell) => (cell.attrs.t === "s" ? strings[Number(child(cell, "v").text)] : Number(child(cell, "v").text)));
        expect(values, `row ${record.row}`).toEqual([...(record.region === null ? [] : [record.region]), record.rep, record.amount, record.serial]);
        expect(cells.map((cell) => cell.attrs.s ?? "0").at(-1), `row ${record.row} date style`).toBe(String(F.xf.date));
      });
      const styles = tree(await part("xl/styles.xml"));
      const xfs = child(styles, "cellXfs").children;
      expect(xfs[F.xf.date]!.attrs.numFmtId).toBe("14");
      const bold = child(styles, "fonts").children[Number(xfs[F.xf.header]!.attrs.fontId)]!;
      expect(childNames(bold)).toContain("b");
    });
  });

  describe("Colours", () => {
    const C = F.colours;

    it("filters by cell colour: a colorFilter naming dxf 0, its two unfilled rows hidden, filterMode on", async () => {
      const xml = await part(C.partName);
      const root = tree(xml);
      expect(childNames(root)).toEqual(["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "sheetData", "autoFilter", "pageMargins"]);
      expect(child(root, "sheetPr").attrs).toEqual({ filterMode: "1" });
      onlyAt(xml, C.autoFilter.xml);
      onlyAt(xml, C.autoFilter.colorFilter.xml);
      const column = child(child(root, "autoFilter"), "filterColumn");
      expect(column.attrs.colId).toBe(String(C.autoFilter.colorFilter.colId));
      expect(child(column, "colorFilter").attrs).toEqual({ dxfId: String(C.autoFilter.colorFilter.dxfId) });
      const rows = child(root, "sheetData").children;
      expect(rows.filter((row) => row.attrs.hidden === "1").map((row) => Number(row.attrs.r))).toEqual(C.filterHiddenRows);
      for (const record of C.rows) {
        const a = rows[record.row - 1]!.children[0]!;
        expect(a.attrs.s === String(F.xf.yellow), `A${record.row}`).toBe(record.yellow);
        expect(rows[record.row - 1]!.attrs.hidden === "1", `row ${record.row}`).toBe(!record.yellow);
      }
    });

    it("has the dxf in styles.xml: one solid yellow fill, and the yellow cell xf to match", async () => {
      const styles = tree(await part("xl/styles.xml"));
      const dxfs = child(styles, "dxfs");
      expect(dxfs.attrs.count).toBe("1");
      const fill = child(child(child(dxfs.children[F.dxf.id]!, "fill"), "patternFill"), "bgColor");
      expect(fill.attrs.rgb).toBe(`FF${F.dxf.fill.slice(1).toUpperCase()}`);
      const yellowXf = child(styles, "cellXfs").children[F.xf.yellow]!;
      const cellFill = child(child(child(styles, "fills").children[Number(yellowXf.attrs.fillId)]!, "patternFill"), "fgColor");
      expect(cellFill.attrs.rgb).toBe(`FF${F.dxf.fill.slice(1).toUpperCase()}`);
    });
  });

  describe("Log", () => {
    const L = F.log;

    it("carries a worksheet-level sortState by cell colour after sheetData, and no autoFilter or hidden row", async () => {
      const xml = await part(L.partName);
      const root = tree(xml);
      expect(childNames(root)).toEqual(["dimension", "sheetViews", "sheetFormatPr", "sheetData", "sortState", "pageMargins"]);
      const at = onlyAt(xml, L.sortState.xml);
      expect(xml.indexOf("</sheetData>")).toBeLessThan(at);
      const sortState = child(root, "sortState");
      expect(sortState.attrs.ref).toBe(L.sortState.ref);
      expect(sortState.children.map((condition) => condition.attrs)).toEqual([
        { sortBy: L.sortState.condition.sortBy, ref: L.sortState.condition.ref, dxfId: String(L.sortState.condition.dxfId) },
      ]);
      expect(child(root, "sheetData").children.some((row) => row.attrs.hidden !== undefined)).toBe(false);
      // The yellow entries sit on top, as the colour sort left them.
      expect(L.rows.map((row) => row.yellow)).toEqual([true, true, false, false, false]);
    });
  });

  it("counts every shared-string reference in the sst", async () => {
    const sstRoot = tree(await part("xl/sharedStrings.xml"));
    let references = 0;
    for (const name of [F.sales.partName, F.colours.partName, F.log.partName]) {
      references += (await part(name)).split(` t="s"`).length - 1;
    }
    expect(sstRoot.attrs.count).toBe(String(references));
    expect(sstRoot.attrs.uniqueCount).toBe(String((await sst()).length));
    expect((await sst()).slice(0, 4)).toEqual(Object.values(F.sales.headers));
  });
});

describe("spec128-filter.ods (LibreOffice-shaped)", () => {
  const O = SPEC128_FILTER_ODS;
  const bytes = buildSpec128FilterOdsBytes();
  const content = async (): Promise<string> => decoder.decode((await unzip(await bytes)).get("content.xml"));

  it("is the committed e2e fixture, byte for byte, and a whole package", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec128-filter.ods")));
    expect(equalBytes(committed, await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    expect([...(await unzip(await bytes)).keys()][0]).toBe("mimetype");
  });

  it("holds one table, Sales, then the named expressions, then the anonymous database range verbatim", async () => {
    const xml = await content();
    const spreadsheet = child(child(tree(xml), "office:body"), "office:spreadsheet");
    expect(childNames(spreadsheet)).toEqual(["table:calculation-settings", "table:table", "table:named-expressions", "table:database-ranges"]);
    expect(child(spreadsheet, "table:table").attrs["table:name"]).toBe(O.tableNames[0]);
    const at = onlyAt(xml, O.databaseRange.xml);
    expect(xml.indexOf("<table:named-expressions/>")).toBeLessThan(at);
    expect(xml.indexOf("</office:spreadsheet>")).toBeGreaterThan(at);
  });

  it("names the range __Anonymous_Sheet_DB__0 with filter buttons, its filter-and split per field", async () => {
    const spreadsheet = child(child(tree(await content()), "office:body"), "office:spreadsheet");
    const ranges = child(spreadsheet, "table:database-ranges").children;
    expect(ranges).toHaveLength(1);
    const range = ranges[0]!;
    expect(range.attrs).toEqual({
      "table:name": O.databaseRange.name,
      "table:target-range-address": O.databaseRange.targetRangeAddress,
      "table:display-filter-buttons": "true",
    });
    const and = child(child(range, "table:filter"), "table:filter-and");
    const [region, low, high] = and.children as [Node, Node, Node];
    expect(and.children).toHaveLength(3);
    expect(region.attrs).toEqual({ "table:field-number": String(O.databaseRange.region.fieldNumber), "table:value": "East", "table:operator": "=" });
    expect(region.children.map((item) => [item.name, item.attrs["table:value"]])).toEqual(
      O.databaseRange.region.setItems.map((value) => ["table:filter-set-item", value]),
    );
    expect([low, high].map((condition) => condition.attrs)).toEqual(
      O.databaseRange.amount.conditions.map((condition) => ({
        "table:field-number": String(O.databaseRange.amount.fieldNumber),
        "table:value": condition.value,
        "table:operator": condition.operator,
        "table:data-type": "number",
      })),
    );
  });

  it("marks rows 4, 5, 9 and 10 visibility filter and row 12 collapse, then Calc's trailing repeated run", async () => {
    const xml = await content();
    expect(xml.split(`table:visibility="filter"`)).toHaveLength(5);
    expect(xml.split(`table:visibility="collapse"`)).toHaveLength(2);
    const table = child(child(child(tree(xml), "office:body"), "office:spreadsheet"), "table:table");
    const rows = table.children.filter((node) => node.name === "table:table-row");
    expect(rows).toHaveLength(14);
    const visibility = rows.slice(0, 13).map((row, index) => [index + 1, row.attrs["table:visibility"]]);
    expect(visibility.filter(([, value]) => value === "filter").map(([r]) => r)).toEqual(O.filterRows);
    expect(visibility.filter(([, value]) => value === "collapse").map(([r]) => r)).toEqual([O.collapseRow]);
    expect(rows.slice(0, 13).every((row) => row.attrs["table:number-rows-repeated"] === undefined)).toBe(true);
    const tail = rows[13]!;
    expect(tail.attrs).toEqual({ "table:style-name": "ro1", "table:number-rows-repeated": String(O.trailingRowRun) });
    expect(tail.children.map((cell) => cell.attrs)).toEqual([{ "table:number-columns-repeated": "16384" }]);
    // Calc's attribute order: style, then visibility.
    expect(xml).toContain(`<table:table-row table:style-name="ro1" table:visibility="filter">`);
  });

  it("holds the records: bold headers, a blank A7, floats and office:value-type date cells under a date data style", async () => {
    const xml = await content();
    const root = tree(xml);
    const rows = child(child(child(root, "office:body"), "office:spreadsheet"), "table:table").children.filter((node) => node.name === "table:table-row");
    const header = rows[0]!.children;
    expect(header.slice(0, 4).map((cell) => [cell.attrs["table:style-name"], child(cell, "text:p").text])).toEqual(
      Object.values(O.headers).map((text) => [O.styles.header, text]),
    );
    for (const record of O.rows) {
      const cells = rows[record.row - 1]!.children;
      expect(cells).toHaveLength(5);
      const [region, rep, amount, date, tail] = cells as [Node, Node, Node, Node, Node];
      if (record.region === null) {
        expect(region.attrs, `A${record.row}`).toEqual({});
      } else {
        expect(child(region, "text:p").text, `A${record.row}`).toBe(record.region);
      }
      expect(child(rep, "text:p").text).toBe(record.rep);
      expect(amount.attrs["office:value-type"]).toBe("float");
      expect(Number(amount.attrs["office:value"])).toBe(record.amount);
      const [y, m, d] = record.date;
      expect(date.attrs).toMatchObject({
        "table:style-name": O.styles.date,
        "office:value-type": "date",
        "office:date-value": `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`,
      });
      expect(tail.attrs).toEqual({ "table:number-columns-repeated": "16380" });
    }
    const automatic = child(root, "office:automatic-styles");
    const dateStyle = automatic.children.find((node) => node.attrs["style:name"] === O.styles.date)!;
    expect(dateStyle.attrs["style:data-style-name"]).toBe(O.styles.dateDataStyle);
    expect(automatic.children.find((node) => node.attrs["style:name"] === O.styles.dateDataStyle)?.name).toBe("number:date-style");
    const bold = automatic.children.find((node) => node.attrs["style:name"] === O.styles.header)!;
    expect(child(bold, "style:text-properties").attrs["fo:font-weight"]).toBe("bold");
  });
});
