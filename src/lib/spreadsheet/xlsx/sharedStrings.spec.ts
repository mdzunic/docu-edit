// SPEC-122 "Shared strings": the SharedStringsWriter one save shares across
// its rebuilt sheets — every indexFor rule (a cell's sstIndex; the first
// source item that is a lone <t> with the same text; an earlier append; a new
// item), rich runs appended with their font XML and deduplicated among
// appends, inline strings converted to shared ones — and the part it writes:
// unchanged unless something was appended, else the source bytes with the
// appends spliced before </sst>, uniqueCount = items and count removed (the
// source <si> bytes untouched), and a workbook without the part gaining it
// with its relationship and Override. (encodeXlsxText's round trip is
// property-tested in sheetXml.spec.ts beside the cells it encodes.)
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { parseA1 } from "../address";
import { buildMiniBookXlsxBytes, buildSpec118GeneratedXlsxBytes, SPEC118_GENERATED_XLSX } from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import type { Cell, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { CellBudget } from "../parsed";
import { SharedStringsWriter } from "./sharedStrings";
import { readSharedStrings, readWorksheet } from "./sheetXml";
import { writeXlsx } from "./writer";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const decoder = new TextDecoder();

/** A source table: plain, rich, phonetic, encoded and escaped items, a
 * repeated text, an empty `<si/>` and an empty `<t/>`. */
const ITEMS = [
  `<si><t>alpha</t></si>`,
  `<si><r><t>beta</t></r></si>`,
  `<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>`,
  `<si><t xml:space="preserve">line_x000D_\nnext</t></si>`,
  `<si><t>a &amp; b</t></si>`,
  `<si><t>alpha</t></si>`,
  `<si/>`,
  `<si><t/></si>`,
];
const SOURCE = `${DECLARATION}\r\n<sst xmlns="${MAIN}" count="12" uniqueCount="${ITEMS.length}">${ITEMS.join("")}</sst>`;

const plain = (value: string): Pick<Cell, "value"> => ({ value });

describe("indexFor", () => {
  it("returns a cell's sstIndex — the item it was read from, rich runs and phonetics intact — and appends nothing", () => {
    const writer = new SharedStringsWriter(SOURCE, undefined);
    expect(writer.indexFor({ value: "beta", runs: [{ text: "beta" }], sstIndex: 1 })).toBe(1);
    expect(writer.indexFor({ value: "東京", sstIndex: 2 })).toBe(2);
    expect(writer.indexFor({ value: "alpha", sstIndex: 5 })).toBe(5);
    expect(writer.changed).toBe(false);
    expect(writer.xml()).toBeUndefined();
  });

  it("gives a plain string the first source item that is a lone <t> with the same (decoded) text — never a rich, phonetic or empty <si/> one", () => {
    const writer = new SharedStringsWriter(SOURCE, undefined);
    expect(writer.indexFor(plain("alpha"))).toBe(0);
    expect(writer.indexFor(plain("line\r\nnext"))).toBe(3);
    expect(writer.indexFor(plain("a & b"))).toBe(4);
    expect(writer.indexFor(plain(""))).toBe(7);
    expect(writer.changed).toBe(false);
    // "beta" exists only as a rich item, 東京 only with its phonetic run: both are appended.
    expect(writer.indexFor(plain("beta"))).toBe(ITEMS.length);
    expect(writer.indexFor(plain("東京"))).toBe(ITEMS.length + 1);
  });

  it("reuses an earlier append, else appends a new <si><t>…</t></si> — encoded, xml:space where whitespace needs it", () => {
    const writer = new SharedStringsWriter(SOURCE, undefined);
    expect(writer.indexFor(plain("new"))).toBe(8);
    expect(writer.indexFor(plain(" padded\r"))).toBe(9);
    expect(writer.indexFor(plain("new"))).toBe(8);
    expect(writer.indexFor(plain("_x0041_ & <x>"))).toBe(10);
    expect(writer.xml()).toBe(
      `${DECLARATION}\r\n<sst xmlns="${MAIN}" uniqueCount="11">${ITEMS.join("")}` +
        `<si><t>new</t></si><si><t xml:space="preserve"> padded_x000D_</t></si><si><t>_x005F_x0041_ &amp; &lt;x&gt;</t></si></sst>`,
    );
  });

  it("appends rich runs as <r><rPr>(the font XML)</rPr><t>…</t></r>, deduplicated among appends only", () => {
    const writer = new SharedStringsWriter(SOURCE, undefined);
    const runs = [{ text: "Bold ", font: { family: { theme: "minor" as const }, sizePt: 11, bold: true as const, color: { theme: "dk1" as const } } }, { text: "plain" }];
    expect(writer.indexFor({ value: "Bold plain", runs })).toBe(8);
    expect(writer.indexFor({ value: "Bold plain", runs: structuredClone(runs) })).toBe(8);
    // The same text without runs is another item; a rich "alpha" is not the plain source item.
    expect(writer.indexFor(plain("Bold plain"))).toBe(9);
    expect(writer.indexFor({ value: "alpha", runs: [{ text: "al", font: { italic: true } }, { text: "pha", font: { strike: true, family: "Cambria" } }] })).toBe(10);
    const xml = writer.xml()!;
    expect(xml.slice(xml.indexOf(ITEMS.at(-1)!) + ITEMS.at(-1)!.length)).toBe(
      `<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">Bold </t></r><r><t>plain</t></r></si>` +
        `<si><t>Bold plain</t></si><si><r><rPr><i/></rPr><t>al</t></r><r><rPr><strike/><rFont val="Cambria"/></rPr><t>pha</t></r></si></sst>`,
    );
    // Read back: the runs and their fonts as written.
    expect(readSharedStrings(xml).slice(8)).toEqual([{ text: "Bold plain", runs }, { text: "Bold plain" }, { text: "alpha", runs: [{ text: "al", font: { italic: true } }, { text: "pha", font: { strike: true, family: "Cambria" } }] }]);
  });

  it("converts inline strings to shared ones: a plain one finds its source item, a rich one is appended", () => {
    const content = readWorksheet(
      `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c>` +
        `<c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>in</t></r><r><t>line</t></r></is></c></row></sheetData></worksheet>`,
      "S",
      { sharedStrings: [], dateSystem: "1900", budget: new CellBudget(), mdw: 7 },
    )!;
    const [a1, b1] = content.rows[0]!.cells;
    expect(a1).not.toHaveProperty("sstIndex");
    const writer = new SharedStringsWriter(SOURCE, undefined);
    expect(writer.indexFor(a1!)).toBe(0);
    expect(writer.indexFor(b1!)).toBe(8);
    expect(writer.xml()).toContain(`<si><r><rPr><b/></rPr><t>in</t></r><r><t>line</t></r></si></sst>`);
  });

  it("without a source part, ignores sstIndex (it names no item) and numbers the appends from 0", () => {
    const writer = new SharedStringsWriter(undefined, undefined);
    expect(writer.indexFor({ value: "b", sstIndex: 4 })).toBe(0);
    expect(writer.indexFor(plain("a"))).toBe(1);
    expect(writer.indexFor(plain("b"))).toBe(0);
    expect(writer.xml()).toBe(`${DECLARATION}<sst xmlns="${MAIN}" uniqueCount="2"><si><t>b</t></si><si><t>a</t></si></sst>`);
  });
});

describe("the part", () => {
  it("splices the appends before </sst> with uniqueCount = items and count removed — every source byte else unchanged", () => {
    const writer = new SharedStringsWriter(SOURCE, undefined);
    writer.indexFor(plain("gamma"));
    const xml = writer.xml()!;
    expect(xml).toBe(`${DECLARATION}\r\n<sst xmlns="${MAIN}" uniqueCount="9">${ITEMS.join("")}<si><t>gamma</t></si></sst>`);
    // Removing the inserted run and restoring the counts gives the source exactly.
    expect(xml.replace(`<si><t>gamma</t></si>`, "").replace(` uniqueCount="9"`, ` count="12" uniqueCount="8"`)).toBe(SOURCE);
  });

  it("expands a self-closing root and writes a prefixed root's prefix on the appends", () => {
    const empty = new SharedStringsWriter(`<sst xmlns="${MAIN}" count="0" uniqueCount="0"/>`, undefined);
    expect(empty.indexFor(plain("x"))).toBe(0);
    expect(empty.xml()).toBe(`<sst xmlns="${MAIN}" uniqueCount="1"><si><t>x</t></si></sst>`);
    const prefixed = new SharedStringsWriter(`<x:sst xmlns:x="${MAIN}" uniqueCount="1"><x:si><x:t>a</x:t></x:si></x:sst>`, undefined);
    expect(prefixed.indexFor(plain("a"))).toBe(0);
    expect(prefixed.indexFor({ value: "bc", runs: [{ text: "b", font: { bold: true } }, { text: "c" }] })).toBe(1);
    expect(prefixed.xml()).toBe(
      `<x:sst xmlns:x="${MAIN}" uniqueCount="2"><x:si><x:t>a</x:t></x:si><x:si><x:r><x:rPr><x:b/></x:rPr><x:t>b</x:t></x:r><x:r><x:t>c</x:t></x:r></x:si></x:sst>`,
    );
  });
});

/** `model` with one cell set (copy-on-write: a new row, cell and rows array)
 * and its sheet marked dirty — what an edit command leaves behind. */
function setCell(model: SpreadsheetModel, sheetIndex: number, a1: string, fields: Omit<Cell, "col">): void {
  const { row: rowIndex, col } = parseA1(a1)!;
  const sheet = model.sheets[sheetIndex]!;
  const cell: Cell = { col, ...fields };
  const found = sheet.rows.find((row) => row.index === rowIndex);
  const rows = found
    ? sheet.rows.map((row) => (row !== found ? row : { ...row, cells: [...row.cells.filter((each) => each.col !== col), cell].toSorted((a, b) => a.col - b.col) }))
    : [...sheet.rows, { index: rowIndex, cells: [cell] }].toSorted((a, b) => a.index - b.index);
  model.sheets = model.sheets.map((each, at) => (at === sheetIndex ? { ...each, rows, dirty: true } : each));
}

describe("through writeXlsx", () => {
  it("mini-book.xlsx: an edit that only reuses items leaves the part equalBytes; a new string is appended, the source <si> bytes untouched", async () => {
    const bytes = await buildMiniBookXlsxBytes();
    const source = decoder.decode((await unzip(bytes)).get("xl/sharedStrings.xml"));
    const model = await readSpreadsheetBytes(bytes, "xlsx", "book");
    setCell(model, 0, "B2", { value: "Cost" });
    expect(equalBytes((await unzip(await writeXlsx(model))).get("xl/sharedStrings.xml")!, (await unzip(bytes)).get("xl/sharedStrings.xml")!)).toBe(true);
    setCell(model, 0, "C2", { value: "Brand new" });
    const written = decoder.decode((await unzip(await writeXlsx(model))).get("xl/sharedStrings.xml"));
    const uniqueCount = Number(/uniqueCount="(\d+)"/.exec(source)![1]);
    expect(written).toBe(source.replace(/ count="\d+"/, "").replace(`uniqueCount="${uniqueCount}"`, `uniqueCount="${uniqueCount + 1}"`).replace("</sst>", "<si><t>Brand new</t></si></sst>"));
    const summary = decoder.decode((await unzip(await writeXlsx(model))).get("xl/worksheets/sheet1.xml"));
    // B2 reuses "Cost" (item 3); C2 is the new item (the cells were replaced whole, so unstyled).
    expect(summary).toContain(`<c r="B2" t="s"><v>3</v></c><c r="C2" t="s"><v>${uniqueCount}</v></c>`);
  });

  it("spec118-generated.xlsx (inline strings, no shared-strings part) gains the part with its relationship and Override", async () => {
    const bytes = await buildSpec118GeneratedXlsxBytes();
    const source = await unzip(bytes);
    expect(source.has("xl/sharedStrings.xml")).toBe(false);
    const model = await readSpreadsheetBytes(bytes, "xlsx", "generated");
    model.sheets[0] = { ...model.sheets[0]!, dirty: true };
    const out = await writeXlsx(model);
    expect(await packageIntegrity(out)).toEqual([]);
    const parts = await unzip(out);
    const G = SPEC118_GENERATED_XLSX.cells;
    expect(decoder.decode(parts.get("xl/sharedStrings.xml"))).toBe(
      `${DECLARATION}<sst xmlns="${MAIN}" uniqueCount="3"><si><t>${G.A1}</t></si><si><t>${G.A2}</t></si><si><t>${G.A3}</t></si></sst>`,
    );
    expect(decoder.decode(parts.get("xl/_rels/workbook.xml.rels"))).toMatch(
      /<Relationship Id="rId\d+" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/sharedStrings" Target="sharedStrings\.xml"\/><\/Relationships>$/,
    );
    expect(decoder.decode(parts.get("[Content_Types].xml"))).toContain(
      `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
    );
    const sheet = decoder.decode(parts.get("xl/worksheets/sheet1.xml"));
    expect(sheet).not.toContain("inlineStr");
    expect(sheet).toContain(`<c r="A1" t="s"><v>0</v></c>`);
    const reread = await readSpreadsheetBytes(out, "xlsx", "generated");
    expect(reread.sheets[0]!.rows.map((row) => row.cells[0]!.value)).toEqual([G.A1, G.A2, G.A3]);
    expect(reread.sheets[0]!.rows.map((row) => row.cells[0]!.sstIndex)).toEqual([0, 1, 2]);
  });
});
