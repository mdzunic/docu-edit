// SPEC-133 "Names on the wire", "Reserved names" and "Reader upgrade" (the
// xlsx read/write AC, the reserved-names AC's wire side, the reader-upgrade
// AC's xlsx half; AC-60's `xlsx/definedNames.spec.ts`):
// `workbook.xml` `definedNames/definedName` read into `names` — scope from
// `localSheetId`'s position in `<sheets>`, orphans, `hidden`, `comment`,
// unmodelled attributes verbatim, the text through the dialect and the
// structured-reference mapping, unparseable text kept — and written back by
// splicing the source part: an unchanged list re-emits the source bytes, a
// matched element keeps its tag with only what the model changed patched, an
// unmatched one is written from the model; the from-scratch serialisation;
// the reader upgrade's source-order mapping; and the model order
// (names.ts setReservedName) the writer emits. Synthetic parts, plus
// spec133-names-tables.xlsx's workbook part read and written.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { buildSpec133NamesTablesXlsxBytes, SPEC133_NAMES_TABLES_XLSX } from "../fixtures/build";
import type { SheetRemap } from "../formula/sheetRefs";
import type { DefinedName, Sheet, SpreadsheetModel } from "../model";
import { setReservedName } from "../names";
import { sourceSheetRemap } from "../sourceNames";
import { applySplices } from "../xmlSplice";
import {
  definedNamesSplices,
  definedNamesXml,
  definedNameXml,
  readNameFormula,
  readSourceDefinedNames,
  readXlsxDefinedNames,
  scanDefinedNames,
  type DefinedNamesWriteContext,
} from "./definedNames";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;
const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** A workbook part whose root holds `body`. */
const part = (body: string): string => `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}">${body}</workbook>`;

/** Data, Lists and a hidden sheet, in that order; IDS are their model ids. */
const SHEETS =
  `<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Lists" sheetId="2" r:id="rId2"/>` +
  `<sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/></sheets>`;
const IDS: readonly string[] = ["data", "lists", "hidden"];

/** An indented block with a placeholder, a reserved name, a scoped name with
 * its attributes out of Excel's order and an unmodelled one, an orphan, an
 * unparseable name and a plain one. */
const NAMES_BLOCK =
  `<definedNames>\n` +
  `    <definedName name="_xlfn.IFS" hidden="1">#NAME?</definedName>\n` +
  `    <definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Lists!$D$1:$E$4</definedName>\n` +
  `    <definedName localSheetId="1" name="Local" publishToServer="1" comment="Codes">Lists!$A$2:$A$4</definedName>\n` +
  `    <definedName name="Far" localSheetId="9">Data!$A$1</definedName>\n` +
  `    <definedName name="Broken">Data!$A$1 +</definedName>\n` +
  `    <definedName name="Rates">Lists!$B$2:$B$4</definedName>\n` +
  `  </definedNames>`;
const SOURCE = part(`<workbookPr/>${SHEETS}${NAMES_BLOCK}<calcPr calcId="191029"/>`);

/** What SOURCE reads as. */
const SOURCE_NAMES: readonly DefinedName[] = [
  { name: "_xlfn.IFS", formula: "#NAME?", hidden: true },
  { name: "_xlnm._FilterDatabase", scopeSheetId: "lists", formula: "Lists!$D$1:$E$4", hidden: true },
  { name: "Local", scopeSheetId: "lists", formula: "Lists!$A$2:$A$4", comment: "Codes", sourceAttrs: ` publishToServer="1"` },
  { name: "Far", formula: "Data!$A$1", hidden: true, sourceAttrs: ` localSheetId="9"` },
  { name: "Broken", formula: "Data!$A$1 +", formulaUnparsed: true },
  { name: "Rates", formula: "Lists!$B$2:$B$4" },
];

const read = (xml: string, ids: readonly (string | undefined)[] = IDS): DefinedName[] => readXlsxDefinedNames(xml, ids);

/** A copy of SOURCE's names with `edit` applied to the one called `name`. */
function sourceNamesWith(name: string, edit: (name: DefinedName) => DefinedName): DefinedName[] {
  return read(SOURCE).map((entry) => (entry.name === name ? edit(entry) : entry));
}

/** A sheet's position in `order` (model ids); undefined when absent. */
const positionIn =
  (order: readonly string[]) =>
  (id: string): number | undefined =>
    order.includes(id) ? order.indexOf(id) : undefined;

/** writer.ts's context: `order` the written `<sheets>` (model ids), the
 * source's IDS. */
function context(order: readonly string[] = IDS, remap?: SheetRemap): DefinedNamesWriteContext {
  const ctx: DefinedNamesWriteContext = {
    indexOf: positionIn(order),
    sourceIndexOf: positionIn(IDS),
  };
  if (remap !== undefined) {
    ctx.remap = remap;
  }
  return ctx;
}

/** The part with `names` spliced into it, as writer.ts does. */
function write(xml: string, names: readonly DefinedName[], order: readonly string[] = IDS, remap?: SheetRemap): string {
  return applySplices(xml, definedNamesSplices(xml, scanDefinedNames(xml), names, context(order, remap)));
}

/** A one-name element of the indented block, with the whitespace before it. */
const line = (element: string): string => `\n    ${element}`;

function sheetOf(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], validations: [], id, name, visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function modelOf(sheets: Sheet[], names: DefinedName[] = []): SpreadsheetModel {
  return { names, dxfs: [], kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 14, sheets, activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0 };
}

describe("readXlsxDefinedNames — workbook.xml definedNames → names", () => {
  it("reads every name in source order", () => {
    expect(read(SOURCE)).toEqual(SOURCE_NAMES);
  });

  it("localSheetId is the 0-based position in <sheets>, hidden and chart sheets counted → scopeSheetId", () => {
    const xml = part(
      `<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Secret" sheetId="2" state="hidden" r:id="rId2"/>` +
        `<sheet name="Chart1" sheetId="3" r:id="rId3"/><sheet name="Lists" sheetId="4" r:id="rId4"/></sheets>` +
        `<definedNames><definedName name="A" localSheetId="0">Data!$A$1</definedName><definedName name="B" localSheetId="1">Secret!$A$1</definedName>` +
        `<definedName name="C" localSheetId="2">Data!$B$1</definedName><definedName name="D" localSheetId="3">Lists!$A$1</definedName>` +
        `<definedName name="E">Lists!$B$1</definedName></definedNames>`,
    );
    expect(read(xml, ["data", "secret", "chart1", "lists"])).toEqual([
      { name: "A", scopeSheetId: "data", formula: "Data!$A$1" },
      { name: "B", scopeSheetId: "secret", formula: "Secret!$A$1" },
      { name: "C", scopeSheetId: "chart1", formula: "Data!$B$1" },
      { name: "D", scopeSheetId: "lists", formula: "Lists!$A$1" },
      { name: "E", formula: "Lists!$B$1" },
    ]);
  });

  it("an out-of-range or non-numeric localSheetId stays verbatim in sourceAttrs, the name hidden (an orphan)", () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName name="Far"  localSheetId = '3'>Data!$A$1</definedName>` +
        `<definedName name="Odd" localSheetId="x" comment="c">Data!$B$1</definedName><definedName name="Neg" localSheetId="-1">1</definedName></definedNames>`,
    );
    expect(read(xml)).toEqual([
      { name: "Far", formula: "Data!$A$1", hidden: true, sourceAttrs: `  localSheetId = '3'` },
      { name: "Odd", formula: "Data!$B$1", comment: "c", hidden: true, sourceAttrs: ` localSheetId="x"` },
      { name: "Neg", formula: "1", hidden: true, sourceAttrs: ` localSheetId="-1"` },
    ]);
  });

  it(`hidden="1" and hidden="true" set hidden; "0" and "false" do not`, () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName name="A" hidden="1">1</definedName><definedName name="B" hidden="true">2</definedName>` +
        `<definedName name="C" hidden="0">3</definedName><definedName name="D" hidden="false">4</definedName></definedNames>`,
    );
    expect(read(xml)).toEqual([
      { name: "A", formula: "1", hidden: true },
      { name: "B", formula: "2", hidden: true },
      { name: "C", formula: "3" },
      { name: "D", formula: "4" },
    ]);
  });

  it("comment → comment, decoded", () => {
    const xml = part(`${SHEETS}<definedNames><definedName name="Rate" comment="Tax &amp; &quot;duty&quot;">0.2</definedName></definedNames>`);
    expect(read(xml)).toEqual([{ name: "Rate", formula: "0.2", comment: `Tax & "duty"` }]);
  });

  it("every other attribute is kept verbatim, with its leading whitespace, in source order, in sourceAttrs", () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName function="1" name="Fn" vbProcedure="1"\n      xlm="1" localSheetId="0" functionGroupId="14" hidden="0" shortcutKey='K' comment="x">Data!$A$1</definedName></definedNames>`,
    );
    expect(read(xml)).toEqual([
      { name: "Fn", scopeSheetId: "data", formula: "Data!$A$1", comment: "x", sourceAttrs: ` function="1" vbProcedure="1"\n      xlm="1" functionGroupId="14" shortcutKey='K'` },
    ]);
  });

  it("the text goes through fromXlsxFormula then structuredFromXlsx: the canonical form", () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName name="Each">Sales[[#This Row],[Price]]*Sales[[#This Row],[Qty]]</definedName>` +
        `<definedName name="Row">Sales[#This Row]</definedName><definedName name="Total">Sales[[#Totals],[Amount]]</definedName>` +
        `<definedName name="Look">_xlfn.XLOOKUP(1,Data!$A:$A,Data!$B:$B)</definedName></definedNames>`,
    );
    expect(read(xml)).toEqual([
      { name: "Each", formula: "Sales[@Price]*Sales[@Qty]" },
      { name: "Row", formula: "Sales[@]" },
      { name: "Total", formula: "Sales[[#Totals],[Amount]]" },
      { name: "Look", formula: "XLOOKUP(1,Data!$A:$A,Data!$B:$B)" },
    ]);
  });

  it("XML-escaped text is decoded", () => {
    const xml = part(`${SHEETS}<definedNames><definedName name="Label">IF(Data!$A$1&lt;1,"a&amp;b","&gt;")</definedName></definedNames>`);
    expect(read(xml)).toEqual([{ name: "Label", formula: `IF(Data!$A$1<1,"a&b",">")` }]);
  });

  it("text that does not parse is kept verbatim with formulaUnparsed", () => {
    const xml = part(`${SHEETS}<definedNames><definedName name="Open">SUM(Data!$A$1</definedName><definedName name="Tail">Data!$A$1 &amp;</definedName></definedNames>`);
    expect(read(xml)).toEqual([
      { name: "Open", formula: "SUM(Data!$A$1", formulaUnparsed: true },
      { name: "Tail", formula: "Data!$A$1 &", formulaUnparsed: true },
    ]);
  });

  it("_xlnm.* names are kept spelled as in the file, sheet-scoped like any name", () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Lists!$D$1:$E$4</definedName>` +
        `<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$E$6</definedName><definedName name="_xlnm.print_titles" localSheetId="0">Data!$1:$1</definedName>` +
        `<definedName name="_xlnm.Extract" localSheetId="2">Hidden!$A$1</definedName></definedNames>`,
    );
    expect(read(xml)).toEqual([
      { name: "_xlnm._FilterDatabase", scopeSheetId: "lists", formula: "Lists!$D$1:$E$4", hidden: true },
      { name: "_xlnm.Print_Area", scopeSheetId: "data", formula: "Data!$A$1:$E$6" },
      { name: "_xlnm.print_titles", scopeSheetId: "data", formula: "Data!$1:$1" },
      { name: "_xlnm.Extract", scopeSheetId: "hidden", formula: "Hidden!$A$1" },
    ]);
  });

  it("a _xlfn. placeholder is read like any name: its #NAME? text parses", () => {
    const xml = part(`${SHEETS}<definedNames><definedName name="_xlfn.IFS" hidden="1">#NAME?</definedName><definedName name="_xlfn.SINGLE">#NAME?</definedName></definedNames>`);
    expect(read(xml)).toEqual([
      { name: "_xlfn.IFS", formula: "#NAME?", hidden: true },
      { name: "_xlfn.SINGLE", formula: "#NAME?" },
    ]);
  });

  it("a workbook without definedNames reads as []", () => {
    expect(read(part(`${SHEETS}<calcPr calcId="1"/>`))).toEqual([]);
    expect(read(part(`${SHEETS}<definedNames/><calcPr calcId="1"/>`))).toEqual([]);
    expect(read(part(`${SHEETS}<definedNames></definedNames>`))).toEqual([]);
  });
});

describe("readNameFormula", () => {
  it("parseable text → its canonical formula, relative to A1", () => {
    expect(readNameFormula("Lists!$B$2:$B$4")).toEqual({ formula: "Lists!$B$2:$B$4" });
    expect(readNameFormula("0.2")).toEqual({ formula: "0.2" });
    expect(readNameFormula("Data!A1048576")).toEqual({ formula: "Data!A1048576" });
    expect(readNameFormula("Sales[[#This Row],[Price]]")).toEqual({ formula: "Sales[@Price]" });
    expect(readNameFormula("#NAME?")).toEqual({ formula: "#NAME?" });
  });

  it("unparseable text → itself, formulaUnparsed", () => {
    expect(readNameFormula("SUM(")).toEqual({ formula: "SUM(", formulaUnparsed: true });
    expect(readNameFormula("Data!$A$1,")).toEqual({ formula: "Data!$A$1,", formulaUnparsed: true });
  });
});

describe("definedNamesSplices — the source workbook part's <definedNames> from the model", () => {
  it("an unchanged name list re-emits the source bytes: no splices", () => {
    const names = read(SOURCE);
    expect(definedNamesSplices(SOURCE, scanDefinedNames(SOURCE), names, context())).toEqual([]);
    expect(write(SOURCE, names)).toBe(SOURCE);
  });

  it("spec133-names-tables.xlsx: its names read as the fixture records them and write back byte-identical", async () => {
    const F = SPEC133_NAMES_TABLES_XLSX;
    const xml = decoder.decode((await unzip(await buildSpec133NamesTablesXlsxBytes())).get("xl/workbook.xml")!);
    const ids: readonly string[] = F.sheetNames;
    const names = readXlsxDefinedNames(xml, ids);
    expect(names).toEqual(
      F.definedNames.map((name) => ({
        name: name.name,
        formula: name.text,
        ...(name.scope === undefined ? {} : { scopeSheetId: name.scope }),
        ...(name.hidden ? { hidden: true } : {}),
        ...(name.comment === undefined ? {} : { comment: name.comment }),
      })),
    );
    expect(definedNamesSplices(xml, scanDefinedNames(xml), names, { indexOf: positionIn(ids), sourceIndexOf: positionIn(ids) })).toEqual([]);
  });

  it("a scope sheet's index change recomputes localSheetId, keeping the tag's attribute order and unmodelled attributes", () => {
    expect(write(SOURCE, read(SOURCE), ["lists", "data", "hidden"])).toBe(
      SOURCE.replace(`name="_xlnm._FilterDatabase" localSheetId="1"`, `name="_xlnm._FilterDatabase" localSheetId="0"`).replace(
        `<definedName localSheetId="1" name="Local" publishToServer="1" comment="Codes">`,
        `<definedName localSheetId="0" name="Local" publishToServer="1" comment="Codes">`,
      ),
    );
  });

  it("a changed formula rewrites only the element's text, in the wire form, escaped", () => {
    const names = sourceNamesWith("Rates", (name) => ({ ...name, formula: "Sales[@Amount]" }));
    expect(write(SOURCE, names)).toBe(SOURCE.replace(`<definedName name="Rates">Lists!$B$2:$B$4<`, `<definedName name="Rates">Sales[[#This Row],[Amount]]<`));
    const local = sourceNamesWith("Local", (name) => ({ ...name, formula: `Lists!$A$2:$A$9&"<"` }));
    expect(write(SOURCE, local)).toBe(SOURCE.replace(`comment="Codes">Lists!$A$2:$A$4<`, `comment="Codes">Lists!$A$2:$A$9&amp;"&lt;"<`));
  });

  it("a changed hidden, comment or name case patches only that attribute of the source tag", () => {
    const names = read(SOURCE).map((name): DefinedName => {
      if (name.name === "_xlnm._FilterDatabase") {
        const { hidden: _hidden, ...rest } = name;
        return rest;
      }
      if (name.name === "Local") {
        return { ...name, comment: `Region & "codes"` };
      }
      return name.name === "Rates" ? { ...name, name: "RATES", hidden: true } : name;
    });
    expect(write(SOURCE, names)).toBe(
      SOURCE.replace(`name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">`, `name="_xlnm._FilterDatabase" localSheetId="1">`)
        .replace(`comment="Codes">`, `comment="Region &amp; &quot;codes&quot;">`)
        .replace(`<definedName name="Rates">`, `<definedName name="RATES" hidden="1">`),
    );
  });

  it("a comment removed drops the attribute; one added is appended to the tag", () => {
    const names = read(SOURCE).map((name): DefinedName => {
      if (name.name === "Local") {
        const { comment: _comment, ...rest } = name;
        return rest;
      }
      return name.name === "Rates" ? { ...name, comment: "Per region" } : name;
    });
    expect(write(SOURCE, names)).toBe(
      SOURCE.replace(` publishToServer="1" comment="Codes">`, ` publishToServer="1">`).replace(`<definedName name="Rates">`, `<definedName name="Rates" comment="Per region">`),
    );
  });

  it("a renamed name no longer matches its source element: it is written from the model in its place", () => {
    const names = sourceNamesWith("Local", (name) => ({ ...name, name: "Codes" }));
    expect(write(SOURCE, names)).toBe(
      SOURCE.replace(
        `<definedName localSheetId="1" name="Local" publishToServer="1" comment="Codes">`,
        `<definedName name="Codes" comment="Codes" localSheetId="1" publishToServer="1">`,
      ),
    );
  });

  it("names are matched by name, case-insensitively, AND scope: same-named names keep their own elements in model order", () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName name="Local">Data!$A$1</definedName>` +
        `<definedName name="Local" localSheetId="1" comment="L">Lists!$A$1</definedName></definedNames><calcPr calcId="1"/>`,
    );
    const [workbookLocal, listsLocal] = read(xml);
    expect(write(xml, [{ ...listsLocal!, name: "LOCAL" }, workbookLocal!], ["lists", "data", "hidden"])).toBe(
      part(
        `${SHEETS}<definedNames><definedName name="LOCAL" localSheetId="0" comment="L">Lists!$A$1</definedName>` +
          `<definedName name="Local">Data!$A$1</definedName></definedNames><calcPr calcId="1"/>`,
      ),
    );
  });

  it("a new name is appended after the source's, with the first element's leading whitespace", () => {
    const names: DefinedName[] = [
      ...read(SOURCE),
      { name: "Margin", formula: "Data!$B$2:$B$5" },
      { name: "Scoped", scopeSheetId: "hidden", formula: "Hidden!$A$1", hidden: true, comment: "c" },
    ];
    expect(write(SOURCE, names)).toBe(
      SOURCE.replace(
        `</definedName>\n  </definedNames>`,
        `</definedName>` +
          line(`<definedName name="Margin">Data!$B$2:$B$5</definedName>`) +
          line(`<definedName name="Scoped" comment="c" localSheetId="2" hidden="1">Hidden!$A$1</definedName>`) +
          `\n  </definedNames>`,
      ),
    );
  });

  it("names are written in model order: a new _FilterDatabase, which setReservedName puts first, is written first", () => {
    const model = modelOf([sheetOf("data", "Data"), sheetOf("lists", "Lists"), sheetOf("hidden", "Hidden")], read(SOURCE));
    setReservedName(model, "data", "_FilterDatabase", "Data!$A$1:$B$5");
    setReservedName(model, "data", "Print_Area", "Data!$A$1:$E$6");
    expect(write(SOURCE, model.names)).toBe(
      SOURCE.replace(
        `<definedNames>\n`,
        `<definedNames>` + line(`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$5</definedName>`) + `\n`,
      ).replace(`</definedName>\n  </definedNames>`, `</definedName>` + line(`<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$E$6</definedName>`) + `\n  </definedNames>`),
    );
  });

  it("a name scoped to a sheet no longer written is not written", () => {
    expect(write(SOURCE, read(SOURCE), ["data", "hidden"])).toBe(
      SOURCE.replace(line(`<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Lists!$D$1:$E$4</definedName>`), "").replace(
        line(`<definedName localSheetId="1" name="Local" publishToServer="1" comment="Codes">Lists!$A$2:$A$4</definedName>`),
        "",
      ),
    );
  });

  it("a <definedNames> the part lacks is created before calcPr, else a later CT_Workbook child, else at the root's end", () => {
    const names: DefinedName[] = [
      { name: "TaxRate", formula: "0.2" },
      { name: "Local", scopeSheetId: "lists", formula: "Lists!$A$1" },
    ];
    const block = `<definedNames><definedName name="TaxRate">0.2</definedName><definedName name="Local" localSheetId="1">Lists!$A$1</definedName></definedNames>`;
    expect(write(part(`<workbookPr/>${SHEETS}<calcPr calcId="1"/><extLst/>`), names)).toBe(part(`<workbookPr/>${SHEETS}${block}<calcPr calcId="1"/><extLst/>`));
    expect(write(part(`${SHEETS}<pivotCaches><pivotCache cacheId="1" r:id="rId9"/></pivotCaches><extLst/>`), names)).toBe(
      part(`${SHEETS}${block}<pivotCaches><pivotCache cacheId="1" r:id="rId9"/></pivotCaches><extLst/>`),
    );
    expect(write(part(`${SHEETS}<extLst/>`), names)).toBe(part(`${SHEETS}${block}<extLst/>`));
    expect(write(part(SHEETS), names)).toBe(part(`${SHEETS}${block}`));
  });

  it("no names and no <definedNames>: nothing is created", () => {
    const xml = part(`${SHEETS}<calcPr calcId="1"/>`);
    expect(definedNamesSplices(xml, scanDefinedNames(xml), [], context())).toEqual([]);
  });

  it("a self-closing <definedNames/> is opened up; with no names it is left alone", () => {
    const xml = part(`${SHEETS}<definedNames/><calcPr calcId="1"/>`);
    expect(write(xml, [{ name: "TaxRate", formula: "0.2" }])).toBe(part(`${SHEETS}<definedNames><definedName name="TaxRate">0.2</definedName></definedNames><calcPr calcId="1"/>`));
    expect(definedNamesSplices(xml, scanDefinedNames(xml), [], context())).toEqual([]);
  });

  it("zero names leave an existing container empty, keeping the block's trailing whitespace", () => {
    const xml = part(`${SHEETS}<definedNames><definedName name="TaxRate">0.2</definedName></definedNames><calcPr calcId="1"/>`);
    expect(write(xml, [])).toBe(part(`${SHEETS}<definedNames></definedNames><calcPr calcId="1"/>`));
    expect(write(SOURCE, [])).toBe(SOURCE.replace(NAMES_BLOCK, `<definedNames>\n  </definedNames>`));
  });

  it("an out-of-range orphan is written back as read, whatever the sheet order (its hidden left as read)", () => {
    const xml = part(
      `${SHEETS}<definedNames><definedName name="Far" localSheetId="9" hidden="0">Data!$A$1</definedName>` +
        `<definedName name="Near" localSheetId="3">2</definedName></definedNames><calcPr calcId="1"/>`,
    );
    const names = read(xml);
    expect(names.every((name) => name.hidden === true)).toBe(true);
    expect(write(xml, names, ["hidden", "lists", "data"])).toBe(xml);
  });

  it("a non-numeric orphan is written back as read", () => {
    const xml = part(`${SHEETS}<definedNames><definedName name="Odd" localSheetId="x">1</definedName></definedNames><calcPr calcId="1"/>`);
    const names = read(xml);
    expect(names).toEqual([{ name: "Odd", formula: "1", hidden: true, sourceAttrs: ` localSheetId="x"` }]);
    expect(write(xml, names, ["hidden", "lists", "data"])).toBe(xml);
  });

  it("an unparsed name is written verbatim, with the save's sheet remap applied when given", () => {
    const remap: SheetRemap = { renamed: { Data: "Facts" }, deleted: [], orderBefore: ["Data", "Lists", "Hidden"] };
    expect(write(SOURCE, read(SOURCE), IDS, remap)).toBe(SOURCE.replace(`<definedName name="Broken">Data!$A$1 +<`, `<definedName name="Broken">Facts!$A$1 +<`));
    const added: DefinedName[] = [...read(SOURCE), { name: "Tail", formula: "Data!$A$1 &", formulaUnparsed: true }];
    expect(write(SOURCE, added)).toBe(SOURCE.replace(`</definedName>\n  </definedNames>`, `</definedName>` + line(`<definedName name="Tail">Data!$A$1 &amp;</definedName>`) + `\n  </definedNames>`));
    expect(write(SOURCE, added, IDS, remap)).toBe(
      SOURCE.replace(`<definedName name="Broken">Data!$A$1 +<`, `<definedName name="Broken">Facts!$A$1 +<`).replace(
        `</definedName>\n  </definedNames>`,
        `</definedName>` + line(`<definedName name="Tail">Facts!$A$1 &amp;</definedName>`) + `\n  </definedNames>`,
      ),
    );
  });

  it("a prefixed (x:) workbook root keeps its prefix", () => {
    const root = `${DECL}<x:workbook xmlns:x="${MAIN}" xmlns:r="${R}"><x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets>`;
    const created = `${root}<x:calcPr calcId="1"/></x:workbook>`;
    expect(write(created, [{ name: "TaxRate", formula: "0.2" }], ["data"])).toBe(
      `${root}<x:definedNames><x:definedName name="TaxRate">0.2</x:definedName></x:definedNames><x:calcPr calcId="1"/></x:workbook>`,
    );
    const existing = `${root}<x:definedNames><x:definedName name="A">1</x:definedName></x:definedNames><x:calcPr calcId="1"/></x:workbook>`;
    expect(write(existing, [{ name: "A", formula: "2" }, { name: "B", scopeSheetId: "data", formula: "3" }], ["data"])).toBe(
      `${root}<x:definedNames><x:definedName name="A">2</x:definedName><x:definedName name="B" localSheetId="0">3</x:definedName></x:definedNames><x:calcPr calcId="1"/></x:workbook>`,
    );
  });
});

describe("definedNameXml / definedNamesXml — from scratch", () => {
  it("attributes in the order name, comment, localSheetId, hidden, then sourceAttrs", () => {
    const name: DefinedName = { name: "Local", scopeSheetId: "lists", formula: "Lists!$A$2", hidden: true, comment: `A & "B"`, sourceAttrs: ` publishToServer="1" function="0"` };
    expect(definedNameXml(name, 1, "")).toBe(`<definedName name="Local" comment="A &amp; &quot;B&quot;" localSheetId="1" hidden="1" publishToServer="1" function="0">Lists!$A$2</definedName>`);
    expect(definedNameXml({ name: "TaxRate", formula: "0.2" }, undefined, "x:")).toBe(`<x:definedName name="TaxRate">0.2</x:definedName>`);
  });

  it("the text is toXlsxFormula(structuredToXlsx(formula)), escaped", () => {
    expect(definedNameXml({ name: "Each", formula: "Sales[@Amount]" }, undefined, "")).toBe(`<definedName name="Each">Sales[[#This Row],[Amount]]</definedName>`);
    expect(definedNameXml({ name: "Row", formula: "Sales[@]" }, undefined, "")).toBe(`<definedName name="Row">Sales[#This Row]</definedName>`);
    expect(definedNameXml({ name: "Label", formula: `IF(Data!$A$1<1,"a&b",">")` }, undefined, "")).toBe(`<definedName name="Label">IF(Data!$A$1&lt;1,"a&amp;b","&gt;")</definedName>`);
    expect(definedNameXml({ name: "Look", formula: "XLOOKUP(1,Data!$A:$A,Data!$B:$B)" }, undefined, "")).toBe(
      `<definedName name="Look">_xlfn.XLOOKUP(1,Data!$A:$A,Data!$B:$B)</definedName>`,
    );
  });

  it("an unparsed name's text is written verbatim, escaped", () => {
    expect(definedNameXml({ name: "Tail", formula: "Data!$A$1 &", formulaUnparsed: true }, undefined, "")).toBe(`<definedName name="Tail">Data!$A$1 &amp;</definedName>`);
  });

  it("definedNamesXml: localSheetId from the written order, names of sheets not written skipped, \"\" for none", () => {
    const indexOf = positionIn(["data", "lists"]);
    const names: DefinedName[] = [
      { name: "_xlnm._FilterDatabase", scopeSheetId: "lists", formula: "Lists!$D$1:$E$4", hidden: true },
      { name: "Gone", scopeSheetId: "deleted", formula: "1" },
      { name: "TaxRate", formula: "0.2", sourceAttrs: ` publishToServer="1" xr:uid="{0}"` },
    ];
    expect(definedNamesXml(names, indexOf)).toBe(
      `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Lists!$D$1:$E$4</definedName>` +
        `<definedName name="TaxRate" publishToServer="1">0.2</definedName></definedNames>`,
    );
    expect(definedNamesXml([], indexOf)).toBe("");
    expect(definedNamesXml([{ name: "Gone", scopeSheetId: "deleted", formula: "1" }], indexOf)).toBe("");
  });

  it("a from-scratch write reads back as the model", () => {
    const names: DefinedName[] = [
      { name: "_xlnm._FilterDatabase", scopeSheetId: "lists", formula: "Lists!$D$1:$E$4", hidden: true },
      { name: "Each", scopeSheetId: "data", formula: "Sales[@Price]*Sales[@Qty]", comment: "per row" },
      { name: "Label", formula: `IF(Data!$A$1<1,"a&b","")` },
    ];
    const xml = part(`${SHEETS}${definedNamesXml(names, positionIn(IDS))}`);
    expect(read(xml)).toEqual(names);
  });
});

describe("readSourceDefinedNames — the reader upgrade (SPEC-133 \"Reader upgrade\")", () => {
  const WORKSHEET = `${R}/worksheet`;
  /** Source order: Data, Secret (hidden), Chart1 (a chartsheet), Lists. */
  function sourceParts(definedNames: string): Map<string, Uint8Array> {
    const texts: [string, string][] = [
      ["_rels/.rels", `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      [
        "xl/workbook.xml",
        part(
          `<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Secret" sheetId="2" state="hidden" r:id="rId2"/>` +
            `<sheet name="Chart1" sheetId="3" r:id="rId3"/><sheet name="Lists" sheetId="4" r:id="rId4"/></sheets>${definedNames}<calcPr calcId="1"/>`,
        ),
      ],
      [
        "xl/_rels/workbook.xml.rels",
        `<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${WORKSHEET}" Target="worksheets/sheet1.xml"/>` +
          `<Relationship Id="rId2" Type="${WORKSHEET}" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${R}/chartsheet" Target="chartsheets/sheet1.xml"/>` +
          `<Relationship Id="rId4" Type="${WORKSHEET}" Target="worksheets/sheet3.xml"/></Relationships>`,
      ],
    ];
    return new Map(texts.map(([name, text]) => [name, encoder.encode(text)]));
  }

  const NAMES =
    `<definedNames><definedName name="OnData" localSheetId="0">Data!$A$1</definedName><definedName name="OnSecret" localSheetId="1">Secret!$A$1</definedName>` +
    `<definedName name="OnChart" localSheetId="2">Lists!$A$1</definedName><definedName name="OnLists" localSheetId="3">Lists!$B$1</definedName>` +
    `<definedName name="Far" localSheetId="9">Data!$A$1</definedName><definedName name="Broken">Data!$A$1 +</definedName>` +
    `<definedName name="Peek">Secret!$B$1+Data!$B$1</definedName></definedNames>`;

  /** The model's sheets, reordered from the source's (and Data renamed Facts). */
  const reordered = (): Sheet[] => [
    sheetOf("lists", "Lists", { partName: "xl/worksheets/sheet3.xml" }),
    sheetOf("facts", "Facts", { partName: "xl/worksheets/sheet1.xml" }),
    sheetOf("chart", "Chart1", { kind: "chart", partName: "xl/chartsheets/sheet1.xml" }),
    sheetOf("secret", "Secret", { visibility: "hidden", partName: "xl/worksheets/sheet2.xml" }),
  ];

  it("maps localSheetId through the SOURCE sheet order to the model sheet holding that part", () => {
    expect(readSourceDefinedNames(sourceParts(NAMES), reordered())).toEqual([
      { name: "OnData", scopeSheetId: "facts", formula: "Data!$A$1" },
      { name: "OnSecret", scopeSheetId: "secret", formula: "Secret!$A$1" },
      { name: "OnChart", scopeSheetId: "chart", formula: "Lists!$A$1" },
      { name: "OnLists", scopeSheetId: "lists", formula: "Lists!$B$1" },
      { name: "Far", formula: "Data!$A$1", hidden: true, sourceAttrs: ` localSheetId="9"` },
      { name: "Broken", formula: "Data!$A$1 +", formulaUnparsed: true },
      { name: "Peek", formula: "Secret!$B$1+Data!$B$1" },
    ]);
  });

  it("a copy holding its origin's part is not the sheet a source name was scoped to", () => {
    const sheets = [...reordered(), sheetOf("copy", "Lists (2)", { partName: "xl/worksheets/sheet3.xml", copiedFrom: "Lists" })];
    expect(readSourceDefinedNames(sourceParts(NAMES), sheets).find((name) => name.name === "OnLists")?.scopeSheetId).toBe("lists");
  });

  it("drops a name scoped to a sheet no longer present and applies the remap to every parsed formula", () => {
    const parts = sourceParts(NAMES);
    const sheets = reordered().filter((sheet) => sheet.id !== "secret");
    const remap = sourceSheetRemap(modelOf(sheets), parts);
    expect(remap).toEqual({ renamed: { Data: "Facts" }, deleted: ["Secret"], orderBefore: ["Data", "Secret", "Chart1", "Lists"] });
    expect(readSourceDefinedNames(parts, sheets, remap)).toEqual([
      { name: "OnData", scopeSheetId: "facts", formula: "Facts!$A$1" },
      { name: "OnChart", scopeSheetId: "chart", formula: "Lists!$A$1" },
      { name: "OnLists", scopeSheetId: "lists", formula: "Lists!$B$1" },
      { name: "Far", formula: "Facts!$A$1", hidden: true, sourceAttrs: ` localSheetId="9"` },
      { name: "Broken", formula: "Data!$A$1 +", formulaUnparsed: true },
      { name: "Peek", formula: "#REF!+Facts!$B$1" },
    ]);
  });

  it("a package without a workbook part, or one without names, gives []", () => {
    expect(readSourceDefinedNames(new Map(), reordered())).toEqual([]);
    expect(readSourceDefinedNames(sourceParts(""), reordered())).toEqual([]);
  });
});

describe("the model order the writer emits (names.ts setReservedName)", () => {
  const sheets = (): Sheet[] => [sheetOf("data", "Data"), sheetOf("lists", "Lists")];
  const RATES: DefinedName = { name: "Rates", formula: "Lists!$B$2:$B$4" };
  const LOCAL: DefinedName = { name: "Local", scopeSheetId: "lists", formula: "Lists!$A$2:$A$4" };

  it("a new _FilterDatabase goes FIRST, hidden", () => {
    const model = modelOf(sheets(), [RATES, LOCAL]);
    setReservedName(model, "data", "_FilterDatabase", "Data!$A$1:$B$5");
    expect(model.names).toEqual([{ name: "_xlnm._FilterDatabase", scopeSheetId: "data", formula: "Data!$A$1:$B$5", hidden: true }, RATES, LOCAL]);
    expect(model.namesDirty).toBe(true);
  });

  it("any other new reserved name is appended", () => {
    const model = modelOf(sheets(), [RATES, LOCAL]);
    setReservedName(model, "lists", "Print_Area", "Lists!$A$1:$E$4");
    expect(model.names).toEqual([RATES, LOCAL, { name: "_xlnm.Print_Area", scopeSheetId: "lists", formula: "Lists!$A$1:$E$4" }]);
  });

  it("an existing one keeps its place, spelling and other fields", () => {
    const filter: DefinedName = { name: "_xlnm._filterdatabase", scopeSheetId: "lists", formula: "Lists!$D$1:$E$4", hidden: true, sourceAttrs: ` publishToServer="1"` };
    const model = modelOf(sheets(), [RATES, filter, LOCAL]);
    setReservedName(model, "lists", "_FilterDatabase", "Lists!$D$1:$E$9");
    expect(model.names).toEqual([RATES, { ...filter, formula: "Lists!$D$1:$E$9" }, LOCAL]);
  });
});
