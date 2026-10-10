// SPEC-121's .xlsx writer: the from-scratch goldens ("From scratch"), the
// verbatim path, the substitution rules ("The .xlsx substitution path",
// "Deleting a sheet", "Duplicating a sheet", "Interim: dirty sheets before
// SPEC-122"), the retained sheet-reference pass ("Sheet references in
// formulas") and package integrity over every fixture × every command.
import { describe, expect, it } from "vitest";
import { setClock } from "../../core/clock";
import { equalBytes } from "../../core/equality";
import { OFFICE_THEME_XML } from "../../core/officeThemeXml";
import { base64ToBytes, bytesToBase64 } from "../../core/packageParts";
import { unzip } from "../../core/zip";
import { formatA1 } from "../address";
import { spreadsheetDownload } from "../download";
import * as F from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import { workbookForCompare } from "../fixtures/workbookForCompare";
import type { CellValue, Sheet, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import {
  addSheet,
  cannotDuplicateReason,
  createBlankWorkbook,
  deleteSheet,
  duplicateSheet,
  isLastVisibleSheet,
  moveSheet,
  renameSheet,
  setSheetHidden,
  setTabColor,
} from "../sheets";
import { emptyWorksheetXml, FROM_SCRATCH_APP_XML, FROM_SCRATCH_STYLES_XML, writeXlsx } from "./writer";

/** A fixed instant for docProps/core.xml — injected, never "today". */
const MODIFIED = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const WS_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const decoder = new TextDecoder();
const encoder = new TextEncoder();

type Parts = Map<string, Uint8Array>;
const text = (parts: Parts, name: string): string => {
  expect(parts.has(name), name).toBe(true);
  return decoder.decode(parts.get(name));
};

async function open(build: () => Promise<Uint8Array>): Promise<{ source: Parts; model: SpreadsheetModel }> {
  const bytes = await build();
  return { source: await unzip(bytes), model: await readSpreadsheetBytes(bytes, "xlsx", "book") };
}

async function save(model: SpreadsheetModel, options: { fullCalcOnLoad?: boolean } = {}): Promise<{ bytes: Uint8Array; parts: Parts }> {
  const bytes = await writeXlsx(model, { modified: MODIFIED, ...options });
  return { bytes, parts: await unzip(bytes) };
}

const links = (): Promise<{ source: Parts; model: SpreadsheetModel }> => open(F.buildSpec121LinksXlsxBytes);

function indexOf(model: SpreadsheetModel, name: string): number {
  const index = model.sheets.findIndex((sheet) => sheet.name === name);
  expect(index, name).toBeGreaterThanOrEqual(0);
  return index;
}

/** Replaces a retained part's text in `opaque` (a source with that shape). */
function editPart(model: SpreadsheetModel, name: string, edit: (xml: string) => string): void {
  model.opaque = model.opaque.map((part) =>
    part.name === name ? { ...part, data: bytesToBase64(encoder.encode(edit(decoder.decode(base64ToBytes(part.data))))) } : part,
  );
}

/** Every self-closing element `<local …/>` of a text, verbatim. */
const elements = (xml: string, local: string): string[] => [...xml.matchAll(new RegExp(`<${local}\\b[^>]*/>`, "g"))].map((match) => match[0]);

function storedValues(sheet: Sheet): Record<string, CellValue> {
  return Object.fromEntries(sheet.rows.flatMap((row) => row.cells.map((cell) => [formatA1({ row: row.index, col: cell.col }), cell.value])));
}

function formulas(sheet: Sheet): Record<string, string | undefined> {
  return Object.fromEntries(sheet.rows.flatMap((row) => row.cells.filter((cell) => cell.formula !== undefined).map((cell) => [formatA1({ row: row.index, col: cell.col }), cell.formula])));
}

/** Whether a worksheet's first sheetView is tabSelected. */
const tabSelected = (xml: string): boolean => /^(1|true)$/.test(/<(?:\w+:)?sheetView\b[^>]*\btabSelected="([^"]*)"/.exec(xml)?.[1] ?? "");

async function expectReopensAs(bytes: Uint8Array, model: SpreadsheetModel): Promise<SpreadsheetModel> {
  expect(await packageIntegrity(bytes)).toEqual([]);
  const reread = await readSpreadsheetBytes(bytes, "xlsx", "reread");
  expect(workbookForCompare(reread)).toEqual(workbookForCompare(model));
  return reread;
}

/** workbook.xml less what the substitution path rewrites. */
const strip = (xml: string): string =>
  xml.replace(/<sheets>.*<\/sheets>/, "").replace(/<definedNames>.*<\/definedNames>/, "").replace(/<workbookView\b[^>]*\/>/, "").replace(/<calcPr[^>]*\/>/, "");

/** A part's Override content type in a [Content_Types].xml text. */
const typeOf = (xml: string, name: string): string | undefined => new RegExp(`PartName="/${name.replace(/[.]/g, "\\.")}" ContentType="([^"]*)"`).exec(xml)?.[1];

// --- from scratch -------------------------------------------------------------------------

/** createBlankWorkbook() + addSheet + a tab colour on the new (active) sheet. */
function blankPlus(): SpreadsheetModel {
  const model = createBlankWorkbook();
  addSheet(model, 0);
  setTabColor(model, 1, { theme: "accent2", lumMod: 0.6, lumOff: 0.4 });
  return model;
}

const GOLDEN = {
  "[Content_Types].xml":
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WS_TYPE}"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="${WS_TYPE}"/>` +
    `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  "_rels/.rels":
    `${DECL}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${R}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  "docProps/core.xml":
    `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
    `xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Untitled spreadsheet</dc:title>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:modified></cp:coreProperties>`,
  "docProps/app.xml":
    `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>docu-edit</Application></Properties>`,
  "xl/workbook.xml":
    `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}"><workbookPr/><bookViews><workbookView activeTab="1"/></bookViews>` +
    `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="191029"/></workbook>`,
  "xl/_rels/workbook.xml.rels":
    `${DECL}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>` +
    `<Relationship Id="rId2" Type="${R}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${R}/theme" Target="theme/theme1.xml"/>` +
    `<Relationship Id="rId4" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
  "xl/theme/theme1.xml": OFFICE_THEME_XML,
  "xl/styles.xml":
    `${DECL}<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
    `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>`,
  "xl/worksheets/sheet1.xml":
    `${DECL}<worksheet xmlns="${MAIN}" xmlns:r="${R}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
  "xl/worksheets/sheet2.xml":
    `${DECL}<worksheet xmlns="${MAIN}" xmlns:r="${R}"><sheetPr><tabColor theme="5" tint="0.4"/></sheetPr><dimension ref="A1"/>` +
    `<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
} as const;

describe("from scratch: createBlankWorkbook() + addSheet + a tab colour", () => {
  it("writes exactly the listed parts, in order", async () => {
    const { parts } = await save(blankPlus());
    expect([...parts.keys()]).toEqual(Object.keys(GOLDEN));
  });

  for (const [name, golden] of Object.entries(GOLDEN)) {
    it(`${name} matches its golden byte for byte`, async () => {
      const { parts } = await save(blankPlus());
      expect(text(parts, name)).toBe(golden);
    });
  }

  it("exports the styles and app parts as the goldens", () => {
    expect(FROM_SCRATCH_STYLES_XML).toBe(GOLDEN["xl/styles.xml"]);
    expect(FROM_SCRATCH_APP_XML).toBe(GOLDEN["docProps/app.xml"]);
    const model = blankPlus();
    expect(DECL + emptyWorksheetXml(model.sheets[0]!, false, model).slice(DECL.length)).toBe(GOLDEN["xl/worksheets/sheet1.xml"]);
  });

  it("the 1904 system writes workbookPr date1904; hidden and very hidden sheets their state", async () => {
    const model = createBlankWorkbook();
    model.dateSystem = "1904";
    addSheet(model, 0);
    addSheet(model, 1);
    setSheetHidden(model, [1], true);
    model.sheets[2] = { ...model.sheets[2]!, visibility: "veryHidden" };
    model.activeSheetIndex = 0;
    const workbook = text((await save(model)).parts, "xl/workbook.xml");
    expect(workbook).toContain(`<workbookPr date1904="1"/>`);
    expect(workbook).toContain(
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" state="hidden" r:id="rId2"/>` +
        `<sheet name="Sheet3" sheetId="3" state="veryHidden" r:id="rId3"/></sheets>`,
    );
    expect(workbook).toContain(`<workbookView activeTab="0"/>`);
  });

  it("fullCalcOnLoad comes from the option, model.forceFullCalcOnLoad and model.calc.fullCalcOnLoad", async () => {
    const calcPr = async (model: SpreadsheetModel, options: { fullCalcOnLoad?: boolean } = {}): Promise<string> =>
      /<calcPr[^>]*\/>/.exec(text((await save(model, options)).parts, "xl/workbook.xml"))![0];
    expect(await calcPr(createBlankWorkbook())).toBe(`<calcPr calcId="191029"/>`);
    expect(await calcPr(createBlankWorkbook(), { fullCalcOnLoad: true })).toBe(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`);
    expect(await calcPr({ ...createBlankWorkbook(), forceFullCalcOnLoad: true })).toBe(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`);
    expect(await calcPr({ ...createBlankWorkbook(), calc: { fullCalcOnLoad: true } })).toBe(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`);
  });

  it("downloads as 'Untitled spreadsheet.xlsx' with `modified` from core now()", async () => {
    const instant = new Date(Date.UTC(2031, 5, 7, 8, 9, 10));
    setClock(instant);
    try {
      const download = await spreadsheetDownload(createBlankWorkbook());
      expect(download.fileName).toBe("Untitled spreadsheet.xlsx");
      expect(download.notices).toEqual([]);
      expect(text(await unzip(download.bytes), "docProps/core.xml")).toContain(`<dcterms:modified xsi:type="dcterms:W3CDTF">2031-06-07T08:09:10Z</dcterms:modified>`);
    } finally {
      setClock(null);
    }
  });
});

describe("from scratch: every package re-opens to the same workbook (AC-3)", () => {
  const cases: [string, () => SpreadsheetModel][] = [
    ["the blank workbook", createBlankWorkbook],
    ["blank + a sheet + a theme tab colour", blankPlus],
    [
      "hidden and very hidden sheets, a later active sheet, literal and darkened colours, 1904",
      () => {
        const model = createBlankWorkbook();
        model.dateSystem = "1904";
        for (let i = 0; i < 4; i += 1) {
          addSheet(model, model.sheets.length - 1);
        }
        renameSheet(model, 1, `R&D "Q1" <x>`);
        setSheetHidden(model, [2], true);
        model.sheets[3] = { ...model.sheets[3]!, visibility: "veryHidden" };
        setTabColor(model, 0, "#ff0000");
        setTabColor(model, 4, { theme: "accent1", lumMod: 0.75 });
        setTabColor(model, 1, { theme: "accent6", lumMod: 0.5, lumOff: 0.2 });
        model.activeSheetIndex = 4;
        return model;
      },
    ],
    [
      "custom default sizes",
      () => {
        const model = createBlankWorkbook();
        model.sheets[0] = { ...model.sheets[0]!, defaultColWidthPx: 100, defaultRowHeightPt: 20.25 };
        return model;
      },
    ],
  ];
  for (const [label, build] of cases) {
    it(label, async () => {
      const model = build();
      await expectReopensAs((await save(model)).bytes, model);
    });
  }
});

// --- verbatim -----------------------------------------------------------------------------

describe("verbatim: nothing changed", () => {
  for (const build of [F.buildSpec121LinksXlsxBytes, F.buildMiniBookXlsxBytes]) {
    it(`${build.name}: no command, and commands all undone, re-emit every part equalBytes in order`, async () => {
      const { source, model } = await open(build);
      const check = async (): Promise<void> => {
        const { parts } = await save(model);
        expect([...parts.keys()]).toEqual([...source.keys()]);
        for (const [name, bytes] of source) {
          expect(equalBytes(parts.get(name)!, bytes), name).toBe(true);
        }
      };
      await check();
      const snapshot = structuredClone(model);
      addSheet(model, 0);
      renameSheet(model, 0, "Renamed");
      duplicateSheet(model, 0);
      moveSheet(model, 0, model.sheets.length - 1);
      setTabColor(model, 1, "#123456");
      deleteSheet(model, 1);
      for (const key of Object.keys(model) as (keyof SpreadsheetModel)[]) {
        if (!(key in snapshot)) {
          delete model[key];
        }
      }
      Object.assign(model, structuredClone(snapshot));
      await check();
    });
  }
});

// --- substitution on spec121-links.xlsx ------------------------------------------------------

describe("substitution: xl/workbook.xml", () => {
  it("splices <sheets>: originals keep sheetId and rId, a copy and a new sheet get the next free ones", async () => {
    const { source, model } = await links();
    duplicateSheet(model, indexOf(model, "Data"));
    addSheet(model, indexOf(model, "Data (2)"));
    const { bytes, parts } = await save(model);
    const workbook = text(parts, "xl/workbook.xml");
    expect(/<sheets>.*<\/sheets>/.exec(workbook)![0]).toBe(
      `<sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/>` +
        `<sheet name="Data (2)" sheetId="6" r:id="rId11"/><sheet name="Sheet1" sheetId="7" r:id="rId12"/><sheet name="Lookup" sheetId="3" r:id="rId3"/>` +
        `<sheet name="Archive" sheetId="4" state="hidden" r:id="rId4"/><sheet name="Vault" sheetId="5" state="veryHidden" r:id="rId5"/></sheets>`,
    );
    // Everything outside <sheets>, <definedNames>, <workbookView> and <calcPr> is the source's.
    expect(strip(workbook)).toBe(strip(text(source, "xl/workbook.xml")));
    expect(await packageIntegrity(bytes)).toEqual([]);
  });

  it("renumbers definedName@localSheetId by the moved sheet's new index; activeTab follows, firstSheet goes", async () => {
    const { model } = await links();
    editPart(model, "xl/workbook.xml", (xml) => xml.replace(` xr2:uid=`, ` firstSheet="1" activeTab="1" xr2:uid=`));
    moveSheet(model, indexOf(model, "Data"), 4);
    const workbook = text((await save(model)).parts, "xl/workbook.xml");
    expect(workbook).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="4">Data!$A$1:$B$5</definedName>`);
    expect(workbook).toContain(`<definedName name="Amounts" localSheetId="4">Data!$B$2:$B$5</definedName>`);
    expect(workbook).not.toContain("firstSheet");
    expect(/<workbookView\b[^>]*>/.exec(workbook)![0]).toMatch(/\sactiveTab="4"/);
    expect(workbook.match(/activeTab=/g)).toHaveLength(1);
  });

  it("renumbers names after an earlier sheet is deleted and removes the deleted sheet's own names", async () => {
    const { model } = await links();
    deleteSheet(model, indexOf(model, "Summary"));
    let workbook = text((await save(model)).parts, "xl/workbook.xml");
    expect(workbook).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$5</definedName>`);
    expect(workbook).toContain(`<definedName name="Amounts" localSheetId="0">Data!$B$2:$B$5</definedName>`);

    const second = (await links()).model;
    deleteSheet(second, indexOf(second, "Data"));
    workbook = text((await save(second)).parts, "xl/workbook.xml");
    expect(workbook).not.toContain("_xlnm.Print_Area");
    expect(workbook).not.toContain(`name="Amounts"`);
    expect(workbook).toContain(`<definedName name="LookupRange">Lookup!$A$1:$B$5</definedName><definedName name="Rate">#REF!</definedName>`);
  });

  it("duplicates a sheet-scoped name for a copy, its text remapped to the copy's name", async () => {
    const { model } = await links();
    duplicateSheet(model, indexOf(model, "Data"));
    const workbook = text((await save(model)).parts, "xl/workbook.xml");
    expect(workbook).toContain(
      `<definedName name="_xlnm.Print_Area" localSheetId="1">Data!$A$1:$B$5</definedName>` +
        `<definedName name="_xlnm.Print_Area" localSheetId="2">'Data (2)'!$A$1:$B$5</definedName>`,
    );
    expect(workbook).toContain(
      `<definedName name="Amounts" localSheetId="1">Data!$B$2:$B$5</definedName><definedName name="Amounts" localSheetId="2">'Data (2)'!$B$2:$B$5</definedName>`,
    );
    expect(workbook).toContain(`<definedName name="Rate">Data!$B$2</definedName>`);
  });
});

describe("substitution: relationships, content types, calcChain, app.xml", () => {
  it("splices the workbook rels and [Content_Types].xml: untouched elements byte-identical, TargetMode kept, calcChain gone", async () => {
    const external = `<Relationship Id="rId42" Type="${R}/hyperlink" Target="https://example.com/a?b=1&amp;c=2" TargetMode="External"/>`;
    const { model } = await links();
    editPart(model, "xl/_rels/workbook.xml.rels", (xml) => xml.replace("</Relationships>", `${external}</Relationships>`));
    const sourceRels = decoder.decode(base64ToBytes(model.opaque.find((part) => part.name === "xl/_rels/workbook.xml.rels")!.data));
    const { source } = await links();
    addSheet(model, 0);
    const { bytes, parts } = await save(model);
    const rels = text(parts, "xl/_rels/workbook.xml.rels");
    for (const element of elements(sourceRels, "Relationship").filter((each) => !each.includes("calcChain"))) {
      expect(rels, element).toContain(element);
    }
    expect(rels).toContain(external);
    expect(rels).not.toContain("calcChain");
    expect(rels).toContain(`<Relationship Id="rId43" Type="${R}/worksheet" Target="worksheets/sheet6.xml"/>`);
    const types = text(parts, "[Content_Types].xml");
    for (const element of [...elements(text(source, "[Content_Types].xml"), "Override"), ...elements(text(source, "[Content_Types].xml"), "Default")].filter(
      (each) => !each.includes("calcChain"),
    )) {
      expect(types, element).toContain(element);
    }
    expect(types).not.toContain("calcChain");
    expect(types).toContain(`<Override PartName="/xl/worksheets/sheet6.xml" ContentType="${WS_TYPE}"/>`);
    expect(parts.has("xl/calcChain.xml")).toBe(false);
    expect(await packageIntegrity(bytes)).toEqual([]);
  });

  it("rebuilds docProps/app.xml's sheet titles (count updated) and leaves no deleted name behind", async () => {
    const { model } = await links();
    addSheet(model, 0);
    deleteSheet(model, indexOf(model, "Lookup"));
    renameSheet(model, indexOf(model, "Data"), "Facts");
    const app = text((await save(model)).parts, "docProps/app.xml");
    expect(app).toContain(
      `<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>5</vt:i4></vt:variant>`,
    );
    const titles = [.../<TitlesOfParts>(.*)<\/TitlesOfParts>/.exec(app)![1]!.matchAll(/<vt:lpstr>([^<]*)<\/vt:lpstr>/g)].map((match) => match[1]);
    expect(titles.slice(0, 5)).toEqual(["Summary", "Sheet1", "Facts", "Archive", "Vault"]);
    expect(titles.some((title) => /^Lookup(!|$)/.test(title!))).toBe(false);
    const counts = [...app.matchAll(/<vt:i4>(\d+)<\/vt:i4>/g)].reduce((sum, match) => sum + Number(match[1]), 0);
    expect(counts).toBe(titles.length);
    expect(app).toContain(`<vt:vector size="${titles.length}" baseType="lpstr">`);
    expect(app.startsWith(`${DECL}\r\n<Properties`)).toBe(true);
    expect(app).toContain(`<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs>`);
  });

  it("removes HeadingPairs and TitlesOfParts when the sheet titles cannot be identified, every other byte kept", async () => {
    const { model } = await links();
    editPart(model, "docProps/app.xml", (xml) => xml.replace("<vt:lpstr>Vault</vt:lpstr>", "<vt:lpstr>Stale</vt:lpstr>"));
    const before = decoder.decode(base64ToBytes(model.opaque.find((part) => part.name === "docProps/app.xml")!.data));
    addSheet(model, 0);
    const app = text((await save(model)).parts, "docProps/app.xml");
    expect(app).toBe(before.replace(/<HeadingPairs>.*<\/HeadingPairs>/, "").replace(/<TitlesOfParts>.*<\/TitlesOfParts>/, ""));
  });
});

describe("substitution: sheet parts", () => {
  it("every worksheet part the commands did not touch is equalBytes its source; tabSelected agrees with the active sheet", async () => {
    const { source, model } = await links();
    addSheet(model, model.sheets.length - 1);
    model.activeSheetIndex = 0; // view state: Summary stays the active tab
    setSheetHidden(model, [indexOf(model, "Lookup")], true);
    const { bytes, parts } = await save(model);
    for (const name of ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml", "xl/worksheets/sheet4.xml", "xl/worksheets/sheet5.xml"]) {
      expect(equalBytes(parts.get(name)!, source.get(name)!), name).toBe(true);
    }
    expect(text(parts, "xl/worksheets/sheet6.xml")).not.toContain("tabSelected");
    expect(await packageIntegrity(bytes)).toEqual([]);

    // Hiding the active sheet moves the selection: exactly one tab selected, the active one.
    setSheetHidden(model, [0], true);
    const after = (await save(model)).parts;
    const active = model.sheets[model.activeSheetIndex]!;
    for (const sheet of model.sheets) {
      const partName = sheet.partName ?? "xl/worksheets/sheet6.xml";
      expect(tabSelected(text(after, partName)), sheet.name).toBe(sheet === active);
    }
    expect(text(after, "xl/worksheets/sheet1.xml")).toBe(text(source, "xl/worksheets/sheet1.xml").replace(`<sheetView tabSelected="1" workbookViewId="0">`, `<sheetView workbookViewId="0">`));
  });

  it("splices <tabColor> (theme refs as theme + tint, else rgb) and removes it when cleared", async () => {
    const { source, model } = await links();
    setTabColor(model, indexOf(model, "Data"), { theme: "accent1", lumMod: 0.75 });
    setTabColor(model, indexOf(model, "Summary"), { theme: "accent6", lumMod: 0.6, lumOff: 0.4 });
    setTabColor(model, indexOf(model, "Lookup"), { theme: "accent1", lumMod: 0.5, lumOff: 0.2 });
    let parts = (await save(model)).parts;
    const data = text(source, "xl/worksheets/sheet2.xml");
    expect(text(parts, "xl/worksheets/sheet2.xml")).toBe(data.replace(`<tabColor theme="5" tint="0.39997558519241921"/>`, `<tabColor theme="4" tint="-0.25"/>`));
    const summary = text(source, "xl/worksheets/sheet1.xml");
    expect(text(parts, "xl/worksheets/sheet1.xml")).toBe(summary.replace(`<dimension `, `<sheetPr><tabColor theme="9" tint="0.4"/></sheetPr><dimension `));
    // accent1 (#156082 in this workbook's theme) lumMod .5 lumOff .2 is no tint: written as the rgb it resolves to.
    expect(text(parts, "xl/worksheets/sheet3.xml")).toMatch(/<sheetPr><tabColor rgb="FF[0-9A-F]{6}"\/><\/sheetPr><dimension /);
    const reread = await readSpreadsheetBytes(await writeXlsx(model, { modified: MODIFIED }), "xlsx", "t");
    expect(workbookForCompare(reread)).toEqual(workbookForCompare(model));

    setTabColor(model, indexOf(model, "Data"), undefined);
    parts = (await save(model)).parts;
    expect(text(parts, "xl/worksheets/sheet2.xml")).not.toContain("tabColor");
    expect(text(parts, "xl/worksheets/sheet2.xml")).toBe(data.replace(`<tabColor theme="5" tint="0.39997558519241921"/>`, ""));
  });
});

describe("substitution: the retained sheet-reference pass", () => {
  it('renaming "Data" to "Facts" rewrites <f>, chart c:f and definedName text — every other byte verbatim', async () => {
    const { source, model } = await links();
    renameSheet(model, indexOf(model, "Data"), "Facts");
    const { bytes, parts } = await save(model);
    expect(text(parts, "xl/worksheets/sheet1.xml")).toBe(
      text(source, "xl/worksheets/sheet1.xml").replace("<f>Data!B2*2</f>", "<f>Facts!B2*2</f>").replace("<f>SUM(Data:Lookup!A1)</f>", "<f>SUM(Facts:Lookup!A1)</f>"),
    );
    expect(text(parts, "xl/charts/chart1.xml")).toBe(text(source, "xl/charts/chart1.xml").replaceAll("<c:f>Data!", "<c:f>Facts!"));
    for (const name of ["xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml", "xl/tables/table1.xml", "xl/drawings/drawing1.xml", "xl/sharedStrings.xml", "docProps/core.xml"]) {
      expect(equalBytes(parts.get(name)!, source.get(name)!), name).toBe(true);
    }
    const workbook = text(parts, "xl/workbook.xml");
    expect(workbook).toContain(`<sheet name="Facts" sheetId="2" r:id="rId2"/>`);
    expect(workbook).toContain(
      `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">Facts!$A$1:$B$5</definedName><definedName name="Amounts" localSheetId="1">Facts!$B$2:$B$5</definedName>` +
        `<definedName name="LookupRange">Lookup!$A$1:$B$5</definedName><definedName name="Rate">Facts!$B$2</definedName></definedNames>`,
    );
    const reread = await expectReopensAs(bytes, model);
    expect(formulas(reread.sheets[0]!)).toEqual(formulas(model.sheets[0]!));
  });

  it('renaming "Lookup" to a name that needs quoting rewrites the cfRule, the validation, the hyperlink location and the names', async () => {
    const { source, model } = await links();
    renameSheet(model, indexOf(model, "Lookup"), "Look Up");
    const { bytes, parts } = await save(model);
    expect(text(parts, "xl/worksheets/sheet2.xml")).toBe(
      text(source, "xl/worksheets/sheet2.xml")
        .replace("<formula>B2&gt;Lookup!$B$1</formula>", "<formula>B2&gt;'Look Up'!$B$1</formula>")
        .replace("<formula1>Lookup!$A$1:$A$5</formula1>", "<formula1>'Look Up'!$A$1:$A$5</formula1>")
        .replace(`location="Lookup!A1"`, `location="'Look Up'!A1"`),
    );
    expect(text(parts, "xl/worksheets/sheet1.xml")).toBe(
      text(source, "xl/worksheets/sheet1.xml")
        .replace("<f>VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)</f>", "<f>VLOOKUP(A2,'Look Up'!$A$1:$B$5,2,FALSE)</f>")
        .replace("<f>SUM(Data:Lookup!A1)</f>", "<f>SUM('Data:Look Up'!A1)</f>")
        .replace("<f>Lookup!B2</f>", "<f>'Look Up'!B2</f>"),
    );
    expect(text(parts, "xl/workbook.xml")).toContain(`<definedName name="LookupRange">'Look Up'!$A$1:$B$5</definedName>`);
    await expectReopensAs(bytes, model);
  });

  it("a swap of two names (A→tmp, B→A, tmp→B) is one simultaneous remap and round-trips", async () => {
    const { source, model } = await links();
    renameSheet(model, indexOf(model, "Data"), "tmp");
    renameSheet(model, indexOf(model, "Lookup"), "Data");
    renameSheet(model, indexOf(model, "tmp"), "Lookup");
    const { bytes, parts } = await save(model);
    expect(text(parts, "xl/worksheets/sheet1.xml")).toBe(
      text(source, "xl/worksheets/sheet1.xml")
        .replace("<f>VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)</f>", "<f>VLOOKUP(A2,Data!$A$1:$B$5,2,FALSE)</f>")
        .replace("<f>Data!B2*2</f>", "<f>Lookup!B2*2</f>")
        .replace("<f>SUM(Data:Lookup!A1)</f>", "<f>SUM(Lookup:Data!A1)</f>")
        .replace("<f>Lookup!B2</f>", "<f>Data!B2</f>"),
    );
    expect(text(parts, "xl/charts/chart1.xml")).toBe(text(source, "xl/charts/chart1.xml").replaceAll("<c:f>Data!", "<c:f>Lookup!"));
    expect(text(parts, "xl/worksheets/sheet2.xml")).toContain(`<formula1>Data!$A$1:$A$5</formula1>`);
    const reread = await expectReopensAs(bytes, model);
    for (const [index, sheet] of model.sheets.entries()) {
      expect(formulas(reread.sheets[index]!), sheet.name).toEqual(formulas(sheet));
    }
  });

  it('deleting "Lookup" turns its references into #REF! and skips its part, relationship and Override', async () => {
    const { source, model } = await links();
    deleteSheet(model, indexOf(model, "Lookup"));
    const { bytes, parts } = await save(model);
    expect(parts.has("xl/worksheets/sheet3.xml")).toBe(false);
    expect(text(parts, "xl/_rels/workbook.xml.rels")).not.toContain("sheet3.xml");
    expect(text(parts, "[Content_Types].xml")).not.toContain("sheet3.xml");
    expect(text(parts, "xl/worksheets/sheet1.xml")).toBe(
      text(source, "xl/worksheets/sheet1.xml")
        .replace("<f>VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)</f>", "<f>VLOOKUP(A2,#REF!,2,FALSE)</f>")
        .replace("<f>SUM(Data:Lookup!A1)</f>", "<f>SUM(Data!A1)</f>")
        .replace("<f>Lookup!B2</f>", "<f>#REF!</f>"),
    );
    expect(text(parts, "xl/worksheets/sheet2.xml")).toBe(
      text(source, "xl/worksheets/sheet2.xml")
        .replace("<formula>B2&gt;Lookup!$B$1</formula>", "<formula>B2&gt;#REF!</formula>")
        .replace("<formula1>Lookup!$A$1:$A$5</formula1>", "<formula1>#REF!</formula1>")
        .replace(`location="Lookup!A1"`, `location="#REF!"`),
    );
    expect(text(parts, "xl/workbook.xml")).toContain(`<definedName name="LookupRange">#REF!</definedName>`);
    await expectReopensAs(bytes, model);
  });

  it('deleting "Data" skips its part closure with the Overrides; shared and never-reachable parts stay', async () => {
    const { model } = await links();
    model.opaque = [...model.opaque, { name: "xl/unreferenced.xml", data: bytesToBase64(encoder.encode(`${DECL}<x/>`)) }];
    deleteSheet(model, indexOf(model, "Data"));
    const { bytes, parts } = await save(model);
    const closure = [
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/_rels/sheet2.xml.rels",
      "xl/drawings/drawing1.xml",
      "xl/drawings/_rels/drawing1.xml.rels",
      "xl/charts/chart1.xml",
      "xl/charts/_rels/chart1.xml.rels",
      "xl/charts/style1.xml",
      "xl/charts/colors1.xml",
      "xl/tables/table1.xml",
      "xl/drawings/vmlDrawing1.vml",
      "xl/comments1.xml",
      "xl/threadedComments/threadedComment1.xml",
    ];
    const types = text(parts, "[Content_Types].xml");
    for (const name of closure) {
      expect(parts.has(name), name).toBe(false);
      expect(types, name).not.toContain(`PartName="/${name}"`);
    }
    expect(parts.has("xl/persons/person.xml")).toBe(true);
    expect(parts.has("xl/unreferenced.xml")).toBe(true);
    expect(await packageIntegrity(bytes)).toEqual([]);
  });
});

describe("interim: dirty sheets before SPEC-122", () => {
  it("a dirty sheet is written from its retained source and the workbook gets calcPr@fullCalcOnLoad", async () => {
    const { source, model } = await links();
    const lookup = indexOf(model, "Lookup");
    model.sheets[lookup] = { ...model.sheets[lookup]!, dirty: true };
    const { parts } = await save(model);
    expect(equalBytes(parts.get("xl/worksheets/sheet3.xml")!, source.get("xl/worksheets/sheet3.xml")!)).toBe(true);
    expect(text(parts, "xl/workbook.xml")).toContain(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`);
    expect(parts.has("xl/calcChain.xml")).toBe(false);
  });

  it("a rename that rewrites another sheet's formulas dirties it, so fullCalcOnLoad; an add alone does not", async () => {
    const renamed = (await links()).model;
    renameSheet(renamed, indexOf(renamed, "Data"), "Facts");
    expect(text((await save(renamed)).parts, "xl/workbook.xml")).toContain(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`);
    const added = (await links()).model;
    addSheet(added, 0);
    expect(text((await save(added)).parts, "xl/workbook.xml")).toContain(`<calcPr calcId="191029"/>`);
  });

  it("an absent calcPr is created (here before </workbook>)", async () => {
    const { model } = await links();
    editPart(model, "xl/workbook.xml", (xml) => xml.replace(`<calcPr calcId="191029"/>`, ""));
    model.sheets[0] = { ...model.sheets[0]!, dirty: true };
    expect(text((await save(model)).parts, "xl/workbook.xml")).toMatch(/<\/definedNames><calcPr fullCalcOnLoad="1"\/><\/workbook>$/);
  });
});

describe("duplicating Data", () => {
  it("writes every copied part with its origin's content type, unique table id/name, a new VML block, re-minted thread ids — and reopens with the copy's cells", async () => {
    const { source, model } = await links();
    duplicateSheet(model, indexOf(model, "Data"));
    const { bytes, parts } = await save(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const types = text(parts, "[Content_Types].xml");
    const sourceTypes = text(source, "[Content_Types].xml");
    const copies: [string, string][] = [
      ["xl/worksheets/sheet6.xml", "xl/worksheets/sheet2.xml"],
      ["xl/tables/table2.xml", "xl/tables/table1.xml"],
      ["xl/drawings/drawing2.xml", "xl/drawings/drawing1.xml"],
      ["xl/charts/chart2.xml", "xl/charts/chart1.xml"],
      ["xl/charts/style2.xml", "xl/charts/style1.xml"],
      ["xl/charts/colors2.xml", "xl/charts/colors1.xml"],
      ["xl/comments2.xml", "xl/comments1.xml"],
      ["xl/threadedComments/threadedComment2.xml", "xl/threadedComments/threadedComment1.xml"],
    ];
    for (const [copy, origin] of copies) {
      expect(parts.has(copy), copy).toBe(true);
      expect(typeOf(types, copy), copy).toBe(typeOf(sourceTypes, origin));
    }
    expect(parts.has("xl/drawings/vmlDrawing2.vml")).toBe(true);
    expect(parts.has("xl/persons/person2.xml")).toBe(false);
    expect(text(parts, "xl/worksheets/_rels/sheet6.xml.rels")).not.toMatch(/Target="[^"]*(table1|drawing1|vmlDrawing1|comments1|threadedComment1)\./);

    const table = /<table\b[^>]*>/.exec(text(parts, "xl/tables/table2.xml"))![0];
    expect(table).toMatch(/\sid="2"/);
    expect(table).toMatch(/\sname="Table1_2"/);
    expect(table).toMatch(/\sdisplayName="Table1_2"/);
    expect(text(parts, "xl/drawings/vmlDrawing2.vml")).toContain(`<o:idmap v:ext="edit" data="2"/>`);
    expect(text(parts, "xl/drawings/vmlDrawing2.vml")).toContain(`id="_x0000_s2049"`);
    const threadId = /<threadedComment\b[^>]*\sid="([^"]*)"/.exec(text(parts, "xl/threadedComments/threadedComment2.xml"))![1]!;
    expect(threadId).not.toBe(F.SPEC121_LINKS_XLSX.threadedComment.id);
    expect(threadId).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
    expect(/<comment\b[^>]*\sxr:uid="([^"]*)"/.exec(text(parts, "xl/comments2.xml"))![1]).toBe(threadId);
    expect(text(parts, "xl/charts/chart2.xml")).toContain(`<c:f>'Data (2)'!$B$2:$B$5</c:f>`);
    expect(text(parts, "xl/charts/chart1.xml")).toBe(text(source, "xl/charts/chart1.xml"));

    const reread = await expectReopensAs(bytes, model);
    expect(storedValues(reread.sheets[indexOf(reread, "Data (2)")]!)).toEqual(storedValues(reread.sheets[indexOf(reread, "Data")]!));
    expect(storedValues(reread.sheets[indexOf(reread, "Data (2)")]!)).toEqual(storedValues(model.sheets[indexOf(model, "Data")]!));
  });
});

describe("substitution on mini-book.xlsx", () => {
  it("add, rename, delete: integrity, re-open, calcChain dropped, titles rebuilt, untouched parts verbatim", async () => {
    const { source, model } = await open(F.buildMiniBookXlsxBytes);
    addSheet(model, model.sheets.length - 1);
    model.activeSheetIndex = 0;
    renameSheet(model, indexOf(model, "Data"), "Numbers");
    deleteSheet(model, indexOf(model, "Hidden"));
    const { bytes, parts } = await save(model);
    await expectReopensAs(bytes, model);
    expect(parts.has("xl/calcChain.xml")).toBe(false);
    expect(parts.has("xl/worksheets/sheet3.xml")).toBe(false);
    for (const name of ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/styles.xml", "xl/theme/theme1.xml", "xl/sharedStrings.xml", "docProps/core.xml"]) {
      expect(equalBytes(parts.get(name)!, source.get(name)!), name).toBe(true);
    }
    expect(text(parts, "docProps/app.xml")).toContain(
      `<TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Summary</vt:lpstr><vt:lpstr>Numbers</vt:lpstr><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>`,
    );
  });
});

// --- every fixture × every command ------------------------------------------------------------

const XLSX_BUILDERS = [
  F.buildMiniBookXlsxBytes,
  F.buildSpec1904XlsxBytes,
  F.buildSpecMacrosXlsmBytes,
  F.buildSpecTemplateXltxBytes,
  F.buildSpec117StylesXlsxBytes,
  F.buildSpec118FormatsXlsxBytes,
  F.buildSpec118GeneratedXlsxBytes,
  F.buildSpec119FormulasXlsxBytes,
  F.buildSpec120CalcXlsxBytes,
  F.buildSpec120FullCalcXlsxBytes,
  F.buildSpec120ManualXlsxBytes,
  F.buildSpec121LinksXlsxBytes,
];

const visible = (model: SpreadsheetModel): number[] => model.sheets.flatMap((sheet, index) => (sheet.visibility === "visible" ? [index] : []));
const ensureTwoVisible = (model: SpreadsheetModel): void => {
  if (visible(model).length < 2) {
    addSheet(model, model.sheets.length - 1);
  }
};

/** Each command, run where it applies; false when it cannot (a single-sheet
 * workbook has nothing deletable). */
const COMMANDS: [string, (model: SpreadsheetModel) => boolean][] = [
  ["add", (model) => (addSheet(model, model.activeSheetIndex), true)],
  ["rename", (model) => (renameSheet(model, 0, `R&D "Q1" <x>`), true)],
  [
    "duplicate",
    (model) => {
      const index = model.sheets.findIndex((sheet, at) => sheet.visibility === "visible" && cannotDuplicateReason(model, at) === null);
      return index >= 0 && (duplicateSheet(model, index), true);
    },
  ],
  [
    "delete",
    (model) => {
      const index = model.sheets.findIndex((_, at) => !isLastVisibleSheet(model, at));
      return index >= 0 && model.sheets.length > 1 && (deleteSheet(model, index), true);
    },
  ],
  [
    "move",
    (model) => {
      if (model.sheets.length < 2) {
        addSheet(model, 0);
      }
      const from = model.activeSheetIndex;
      moveSheet(model, from, from === model.sheets.length - 1 ? 0 : model.sheets.length - 1);
      return true;
    },
  ],
  ["hide", (model) => (ensureTwoVisible(model), setSheetHidden(model, [model.activeSheetIndex], true), true)],
  [
    "unhide",
    (model) => {
      if (!model.sheets.some((sheet) => sheet.visibility === "hidden")) {
        ensureTwoVisible(model);
        setSheetHidden(model, [visible(model).at(-1)!], true);
      }
      setSheetHidden(model, model.sheets.flatMap((sheet, index) => (sheet.visibility === "hidden" ? [index] : [])), false);
      return true;
    },
  ],
  [
    "tab colour",
    (model) => {
      setTabColor(model, 0, { theme: "accent1", lumMod: 0.75 });
      const coloured = model.sheets.findIndex((sheet, index) => index > 0 && sheet.tabColor !== undefined);
      if (coloured > 0) {
        setTabColor(model, coloured, undefined);
      }
      if (model.sheets.length > 1) {
        setTabColor(model, model.sheets.length - 1, "#00ff7f");
      }
      return true;
    },
  ],
];

describe("package integrity: every xlsx fixture × every command, written and re-opened", () => {
  for (const build of XLSX_BUILDERS) {
    it(`${build.name}: the source is sound`, async () => {
      expect(await packageIntegrity(await build())).toEqual([]);
    });
    for (const [label, run] of COMMANDS) {
      it(`${build.name} × ${label}`, async () => {
        const { model } = await open(build);
        if (!run(model)) {
          expect(model.sheets.length).toBe(1);
          return;
        }
        expect(model.structureDirty).toBe(true);
        await expectReopensAs((await save(model)).bytes, model);
      });
    }
  }
});
