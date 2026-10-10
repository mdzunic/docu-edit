// SPEC-121's .ods writer: from-scratch goldens, the verbatim path, "The .ods
// substitution path" (content.xml = prefix + tables + suffix, settings.xml,
// the manifest), the retained sheet-reference pass over content.xml and the
// chart objects, deletes and duplicates, and integrity over every fixture ×
// every command.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { base64ToBytes, bytesToBase64 } from "../../core/packageParts";
import { unzip } from "../../core/zip";
import * as F from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import { workbookForCompare } from "../fixtures/workbookForCompare";
import type { SpreadsheetModel } from "../model";
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
import { FROM_SCRATCH_ODS_MANIFEST_XML, FROM_SCRATCH_ODS_STYLES_XML, writeOds } from "./writer";

const MODIFIED = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
const DECL = `<?xml version="1.0" encoding="UTF-8"?>`;
const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
};
const decoder = new TextDecoder();
const encoder = new TextEncoder();
type Parts = Map<string, Uint8Array>;
const text = (parts: Parts, name: string): string => {
  expect(parts.has(name), name).toBe(true);
  return decoder.decode(parts.get(name));
};

async function open(build: () => Promise<Uint8Array>): Promise<{ source: Parts; model: SpreadsheetModel }> {
  const bytes = await build();
  return { source: await unzip(bytes), model: await readSpreadsheetBytes(bytes, "ods", "book") };
}
async function save(model: SpreadsheetModel): Promise<{ bytes: Uint8Array; parts: Parts }> {
  const bytes = await writeOds(model, { modified: MODIFIED });
  return { bytes, parts: await unzip(bytes) };
}
const links = (): Promise<{ source: Parts; model: SpreadsheetModel }> => open(F.buildSpec121LinksOdsBytes);
const indexOf = (model: SpreadsheetModel, name: string): number => model.sheets.findIndex((sheet) => sheet.name === name);
const sourceXmlOf = (model: SpreadsheetModel, name: string): string => model.sheets[indexOf(model, name)]!.sourceXml!;
function firstEntry(bytes: Uint8Array): { name: string; stored: boolean } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { name: decoder.decode(bytes.subarray(30, 30 + view.getUint16(26, true))), stored: view.getUint16(8, true) === 0 };
}
function editPart(model: SpreadsheetModel, name: string, edit: (xml: string) => string): void {
  model.opaque = model.opaque.map((part) => (part.name === name ? { ...part, data: bytesToBase64(encoder.encode(edit(decoder.decode(base64ToBytes(part.data))))) } : part));
}
const replaceAll = (xml: string, pairs: [string, string][]): string => pairs.reduce((out, [from, to]) => {
  expect(out, from).toContain(from);
  return out.replaceAll(from, to);
}, xml);
async function expectReopensAs(bytes: Uint8Array, model: SpreadsheetModel): Promise<SpreadsheetModel> {
  expect(await packageIntegrity(bytes)).toEqual([]);
  const reread = await readSpreadsheetBytes(bytes, "ods", "reread");
  expect(workbookForCompare(reread)).toEqual(workbookForCompare(model));
  return reread;
}

// --- from scratch ---------------------------------------------------------------------------

const blank = (): SpreadsheetModel => ({ ...createBlankWorkbook(), format: "ods" });

function blankPlus(): SpreadsheetModel {
  const model = { ...createBlankWorkbook(), format: "ods" as const };
  addSheet(model, 0);
  setTabColor(model, 1, { theme: "accent2", lumMod: 0.6, lumOff: 0.4 });
  return model;
}

const TABLE = (name: string, style: string): string =>
  `<table:table table:name="${name}" table:style-name="${style}"><table:table-column table:style-name="de-co1" table:default-cell-style-name="Default"/>` +
  `<table:table-row table:style-name="de-ro1"><table:table-cell/></table:table-row></table:table>`;
const GOLDEN: Record<string, string> = {
  mimetype: "application/vnd.oasis.opendocument.spreadsheet",
  "META-INF/manifest.xml":
    `${DECL}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/></manifest:manifest>`,
  "meta.xml":
    `${DECL}<office:document-meta xmlns:office="${NS.office}" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
    `office:version="1.2"><office:meta><dc:title>Untitled spreadsheet</dc:title><dc:date>2026-01-02T03:04:05Z</dc:date></office:meta></office:document-meta>`,
  "styles.xml":
    `${DECL}<office:document-styles xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:text="${NS.text}" xmlns:table="${NS.table}" xmlns:fo="${NS.fo}" ` +
    `xmlns:svg="${NS.svg}" office:version="1.2"><office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri"/></office:font-face-decls>` +
    `<office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:default-style>` +
    `<style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-size="11pt"/></style:style></office:styles>` +
    `<office:automatic-styles><style:page-layout style:name="pm1"/></office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>`,
  "content.xml":
    `${DECL}<office:document-content xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:text="${NS.text}" xmlns:table="${NS.table}" xmlns:fo="${NS.fo}" ` +
    `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:tableooo="http://openoffice.org/2009/table" office:version="1.2"><office:automatic-styles>` +
    `<style:style style:name="de-ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>` +
    `<style:style style:name="de-co1" style:family="table-column"><style:table-column-properties style:column-width="0.6667in"/></style:style>` +
    `<style:style style:name="de-ro1" style:family="table-row"><style:table-row-properties style:row-height="15pt" style:use-optimal-row-height="true"/></style:style>` +
    `<style:style style:name="de-ta2" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb" tableooo:tab-color="#f4b183"/></style:style>` +
    `</office:automatic-styles><office:body><office:spreadsheet>${TABLE("Sheet1", "de-ta1")}${TABLE("Sheet2", "de-ta2")}</office:spreadsheet></office:body></office:document-content>`,
  "settings.xml":
    `${DECL}<office:document-settings xmlns:office="${NS.office}" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" ` +
    `office:version="1.2"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views">` +
    `<config:config-item-map-entry><config:config-item config:name="ActiveTable" config:type="string">Sheet2</config:config-item></config:config-item-map-entry>` +
    `</config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`,
};

describe("from scratch: createBlankWorkbook() as .ods + addSheet + a tab colour", () => {
  it("writes mimetype first and STOREd, then exactly the listed parts", async () => {
    const { bytes, parts } = await save(blankPlus());
    expect(firstEntry(bytes)).toEqual({ name: "mimetype", stored: true });
    expect([...parts.keys()]).toEqual(["mimetype", "META-INF/manifest.xml", "meta.xml", "styles.xml", "content.xml", "settings.xml"]);
  });
  for (const name of Object.keys(GOLDEN)) {
    it(`${name} matches its golden byte for byte`, async () => {
      expect(text((await save(blankPlus())).parts, name)).toBe(GOLDEN[name]);
    });
  }
  it("exports the manifest and styles goldens", () => {
    expect(FROM_SCRATCH_ODS_MANIFEST_XML).toBe(GOLDEN["META-INF/manifest.xml"]);
    expect(FROM_SCRATCH_ODS_STYLES_XML).toBe(GOLDEN["styles.xml"]);
  });
  it("writes table:calculation-settings with the null date only for 1904 or odsNullDate", async () => {
    const settings = async (model: SpreadsheetModel): Promise<string | undefined> =>
      /<table:calculation-settings>.*?<\/table:calculation-settings>/.exec(text((await save(model)).parts, "content.xml"))?.[0];
    expect(await settings(blank())).toBeUndefined();
    expect(await settings({ ...blank(), dateSystem: "1904" })).toBe(`<table:calculation-settings><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>`);
    const custom = { ...blank(), odsNullDate: "1899-12-31" };
    expect(await settings(custom)).toBe(`<table:calculation-settings><table:null-date table:date-value="1899-12-31"/></table:calculation-settings>`);
    expect((await readSpreadsheetBytes((await save(custom)).bytes, "ods", "t")).odsNullDate).toBe("1899-12-31");
  });
  it("every from-scratch package re-opens to the same workbook (AC-3)", async () => {
    const varied = blankPlus();
    addSheet(varied, 1);
    renameSheet(varied, 2, `R&D "Q1" <x>`);
    setSheetHidden(varied, [0], true);
    setTabColor(varied, 2, "#00ff7f");
    varied.dateSystem = "1904";
    varied.sheets[1] = { ...varied.sheets[1]!, defaultColWidthPx: 100, defaultRowHeightPt: 20.25 };
    for (const model of [{ ...createBlankWorkbook(), format: "ods" as const }, blankPlus(), varied]) {
      await expectReopensAs((await save(model)).bytes, model);
    }
  });
});

// --- verbatim -------------------------------------------------------------------------------

describe("verbatim: nothing changed", () => {
  for (const build of [F.buildSpec121LinksOdsBytes, F.buildMiniBookOdsBytes]) {
    it(`${build.name}: no command, and commands all undone, re-emit every part equalBytes`, async () => {
      const { source, model } = await open(build);
      const check = async (): Promise<void> => {
        const { bytes, parts } = await save(model);
        expect(firstEntry(bytes)).toEqual({ name: "mimetype", stored: true });
        expect([...parts.keys()]).toEqual([...source.keys()]);
        for (const [name, data] of source) {
          expect(equalBytes(parts.get(name)!, data), name).toBe(true);
        }
      };
      await check();
      const snapshot = structuredClone(model);
      addSheet(model, 0);
      renameSheet(model, 1, "Renamed");
      duplicateSheet(model, 1);
      setTabColor(model, 1, "#123456");
      deleteSheet(model, 2);
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

// --- substitution on spec121-links.ods ----------------------------------------------------------

describe("substitution on spec121-links.ods", () => {
  it('renaming "Data" to "Facts": only the pass locations and Data\'s table:name change, in content.xml, settings.xml and Object 1', async () => {
    const { source, model } = await links();
    const envelope = model.odsContentEnvelope!;
    const [summary, data, lookup, archive] = ["Summary", "Data", "Lookup", "Archive"].map((name) => sourceXmlOf(model, name));
    renameSheet(model, indexOf(model, "Data"), "Facts");
    const { bytes, parts } = await save(model);
    const expected =
      replaceAll(envelope.prefix, [[`table:base-cell-address="Data.A2"`, `table:base-cell-address="$Facts.A2"`]]) +
      replaceAll(summary!, [
        ["of:=[$Data.B2]*2", "of:=[$Facts.B2]*2"],
        ["of:=SUM([$Data.A1:$Lookup.A1])", "of:=SUM([$Facts.A1:$Lookup.A1])"],
      ]) +
      replaceAll(data!, [
        [`table:name="Data"`, `table:name="Facts"`],
        [`table:end-cell-address="Data.F12"`, `table:end-cell-address="$Facts.F12"`],
        [`draw:notify-on-update-of-ranges="Data.A1:Data.B5"`, `draw:notify-on-update-of-ranges="$Facts.A1:.B5"`],
        [`table:base-cell-address="$Data.$B$2" table:cell-range-address="$Data.$B$2:.$B$5"`, `table:base-cell-address="$Facts.$B$2" table:cell-range-address="$Facts.$B$2:.$B$5"`],
        [`calcext:target-range-address="Data.B2:Data.B5"`, `calcext:target-range-address="$Facts.B2:.B5"`],
        [`calcext:base-cell-address="Data.B2"`, `calcext:base-cell-address="$Facts.B2"`],
      ]) +
      lookup! +
      archive! +
      replaceAll(envelope.suffix, [
        [`table:base-cell-address="$Data.$B$2" table:expression="of:=[$Data.$B$2]"`, `table:base-cell-address="$Facts.$B$2" table:expression="of:=[$Facts.$B$2]"`],
        [`table:target-range-address="Data.A1:Data.B5"`, `table:target-range-address="$Facts.A1:.B5"`],
      ]);
    expect(text(parts, "content.xml")).toBe(expected);
    expect(text(parts, "Object 1/content.xml")).toBe(
      replaceAll(text(source, "Object 1/content.xml"), [
        [`"Data.A1:Data.B5"`, `"$Facts.A1:.B5"`],
        [`"Data.A2:Data.A5"`, `"$Facts.A2:.A5"`],
        [`"Data.B2:Data.B5"`, `"$Facts.B2:.B5"`],
        [`"Data.B1"`, `"$Facts.B1"`],
      ]),
    );
    expect(text(parts, "settings.xml")).toBe(replaceAll(text(source, "settings.xml"), [[`config:name="Data"`, `config:name="Facts"`]]));
    for (const name of ["styles.xml", "meta.xml", "META-INF/manifest.xml", "Object 1/styles.xml", "ObjectReplacements/Object 1", "Thumbnails/thumbnail.png"]) {
      expect(equalBytes(parts.get(name)!, source.get(name)!), name).toBe(true);
    }
    await expectReopensAs(bytes, model);
  });

  it('renaming "Lookup" to a quoted name rewrites the validation and conditional-format conditions, the link and the named range', async () => {
    const { model } = await links();
    renameSheet(model, indexOf(model, "Lookup"), "Look Up");
    const { bytes, parts } = await save(model);
    const content = text(parts, "content.xml");
    for (const expected of [
      `table:condition="of:cell-content-is-in-list([$'Look Up'.$A$1:.$A$5])"`,
      `calcext:value="formula-is([.B2]&gt;[$'Look Up'.$B$1])"`,
      `xlink:href="#$'Look Up'.A1"`,
      `table:formula="of:=VLOOKUP([.A2];[$'Look Up'.$A$1:.$B$5];2;0)"`,
      `table:base-cell-address="$'Look Up'.$A$1" table:cell-range-address="$'Look Up'.$A$1:.$B$5"`,
    ]) {
      expect(content).toContain(expected);
    }
    expect(content).not.toMatch(/\$Lookup\./);
    await expectReopensAs(bytes, model);
  });

  it("hiding and recolouring: de-ta styles appended to the automatic styles, only table:style-name spliced, TabColor dropped", async () => {
    const { source, model } = await links();
    const envelope = model.odsContentEnvelope!;
    const lookup = sourceXmlOf(model, "Lookup");
    const data = sourceXmlOf(model, "Data");
    setSheetHidden(model, [indexOf(model, "Lookup")], true);
    setTabColor(model, indexOf(model, "Data"), "#00ff7f");
    const { bytes, parts } = await save(model);
    const content = text(parts, "content.xml");
    const styles =
      `<style:style style:name="de-ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb" tableooo:tab-color="#00ff7f"/></style:style>` +
      `<style:style style:name="de-ta2" style:family="table" style:master-page-name="Default"><style:table-properties table:display="false" style:writing-mode="lr-tb"/></style:style>`;
    expect(content.startsWith(envelope.prefix.replace("</office:automatic-styles>", `${styles}</office:automatic-styles>`))).toBe(true);
    expect(content).toContain(lookup.replace(`table:style-name="ta1"`, `table:style-name="de-ta2"`));
    expect(content).toContain(data.replace(`table:style-name="ta2"`, `table:style-name="de-ta1"`));
    expect(content).toContain(sourceXmlOf(model, "Summary"));
    const settings = text(parts, "settings.xml");
    expect(settings).not.toContain(`config:name="TabColor"`);
    expect(settings).toBe(text(source, "settings.xml").replace(/<config:config-item config:name="TabColor"[^>]*>[^<]*<\/config:config-item>/, ""));
    await expectReopensAs(bytes, model);
  });

  it("starts the generated-style counter past every de- name already in the package", async () => {
    const { model } = await links();
    editPart(model, "styles.xml", (xml) => xml.replace("</office:styles>", `<style:style style:name="de-ta7" style:family="table"/></office:styles>`));
    setSheetHidden(model, [indexOf(model, "Lookup")], true);
    const content = text((await save(model)).parts, "content.xml");
    expect(content).toContain(`<style:style style:name="de-ta8" style:family="table"`);
    expect(content).not.toContain(`style:name="de-ta1"`);
  });

  it('deleting "Data" removes its database range and the chart object only it showed, with their manifest entries', async () => {
    const { source, model } = await links();
    deleteSheet(model, indexOf(model, "Data"));
    const { bytes, parts } = await save(model);
    expect([...parts.keys()]).toEqual([...source.keys()].filter((name) => !/^(Object 1\/|ObjectReplacements\/Object 1$)/.test(name)));
    const manifest = text(parts, "META-INF/manifest.xml");
    expect(manifest).not.toContain("Object 1");
    expect(manifest).toBe(text(source, "META-INF/manifest.xml").replace(/<manifest:file-entry manifest:full-path="(Object 1\/[^"]*|ObjectReplacements\/Object 1)"[^>]*\/>/g, ""));
    const content = text(parts, "content.xml");
    expect(content).not.toContain(`table:name="Table1"`);
    expect(content).not.toContain(`table:name="Data"`);
    expect(content).toContain(`table:name="LookupRange"`);
    expect(content).toContain(sourceXmlOf(model, "Lookup"));
    const settings = text(parts, "settings.xml");
    expect([...settings.matchAll(/config-item-map-entry config:name="([^"]*)"/g)].map((match) => match[1])).toEqual(["Summary", "Lookup", "Archive"]);
    expect(settings).toContain(`<config:config-item config:name="ActiveTable" config:type="string">Summary</config:config-item>`);
    await expectReopensAs(bytes, model);
  });

  it("duplicating Data: the copy's ids suffixed, its object copied and remapped, its settings entry copied", async () => {
    const { source, model } = await links();
    duplicateSheet(model, indexOf(model, "Data"));
    const { bytes, parts } = await save(model);
    for (const name of ["Object 2/content.xml", "Object 2/styles.xml", "Object 2/meta.xml", "ObjectReplacements/Object 2"]) {
      expect(parts.has(name), name).toBe(true);
    }
    const copy = /<table:table table:name="Data \(2\)"[\s\S]*?<\/table:shapes>/.exec(text(parts, "content.xml"))![0];
    expect(copy).toContain(`xml:id="id1-c1" draw:id="id1-c1"`);
    expect(copy).toContain(`xlink:href="./Object 2"`);
    expect(copy).toContain(`xlink:href="./ObjectReplacements/Object 2"`);
    expect(copy).toContain(`draw:notify-on-update-of-ranges="$'Data (2)'.A1:.B5"`);
    expect(text(parts, "Object 2/content.xml")).toContain(`chart:values-cell-range-address="$'Data (2)'.B2:.B5"`);
    expect(equalBytes(parts.get("Object 1/content.xml")!, source.get("Object 1/content.xml")!)).toBe(true);
    const settings = text(parts, "settings.xml");
    expect([...settings.matchAll(/config-item-map-entry config:name="([^"]*)"/g)].map((match) => match[1])).toEqual(["Summary", "Data", "Data (2)", "Lookup", "Archive"]);
    expect(settings).toContain(`<config:config-item config:name="ActiveTable" config:type="string">Data (2)</config:config-item>`);
    await expectReopensAs(bytes, model);
  });
});

describe("substitution on mini-book.ods", () => {
  it("add and move: every untouched table splice-equal, in the new order; integrity, re-open", async () => {
    const { model } = await open(F.buildMiniBookOdsBytes);
    const [archive, inventory, totals] = ["Archive", "Inventory", "Totals"].map((name) => sourceXmlOf(model, name));
    addSheet(model, model.sheets.length - 1);
    moveSheet(model, indexOf(model, "Totals"), 0);
    const { bytes, parts } = await save(model);
    const content = text(parts, "content.xml");
    expect(content).toContain(totals! + archive! + inventory! + `<table:table table:name="Sheet1" table:style-name="de-ta1">`);
    expect(content.endsWith(model.odsContentEnvelope!.suffix)).toBe(true);
    expect(text(parts, "settings.xml")).toContain(`<config:config-item config:name="ActiveTable" config:type="string">Totals</config:config-item>`);
    await expectReopensAs(bytes, model);
  });
});

// --- every fixture × every command ------------------------------------------------------------

const ODS_BUILDERS = [
  F.buildMiniBookOdsBytes,
  F.buildSpecWideOdsBytes,
  F.buildSpec117StylesOdsBytes,
  F.buildSpec118FormatsOdsBytes,
  F.buildSpec119FormulasOdsBytes,
  F.buildSpec120CalcOdsBytes,
  F.buildSpec121LinksOdsBytes,
];
const visible = (model: SpreadsheetModel): number[] => model.sheets.flatMap((sheet, index) => (sheet.visibility === "visible" ? [index] : []));
const ensureTwoVisible = (model: SpreadsheetModel): void => {
  if (visible(model).length < 2) {
    addSheet(model, model.sheets.length - 1);
  }
};
const COMMANDS: [string, (model: SpreadsheetModel) => boolean][] = [
  ["add", (model) => (addSheet(model, model.activeSheetIndex), true)],
  ["rename", (model) => (renameSheet(model, 0, `R&D "Q1" <x>`), true)],
  ["duplicate", (model) => {
    const index = model.sheets.findIndex((sheet, at) => sheet.visibility === "visible" && cannotDuplicateReason(model, at) === null);
    return index >= 0 && (duplicateSheet(model, index), true);
  }],
  ["delete", (model) => {
    const index = model.sheets.findIndex((_, at) => !isLastVisibleSheet(model, at));
    return index >= 0 && model.sheets.length > 1 && (deleteSheet(model, index), true);
  }],
  ["move", (model) => {
    if (model.sheets.length < 2) {
      addSheet(model, 0);
    }
    const from = model.activeSheetIndex;
    moveSheet(model, from, from === model.sheets.length - 1 ? 0 : model.sheets.length - 1);
    return true;
  }],
  ["hide", (model) => (ensureTwoVisible(model), setSheetHidden(model, [model.activeSheetIndex], true), true)],
  ["unhide", (model) => {
    if (!model.sheets.some((sheet) => sheet.visibility === "hidden")) {
      ensureTwoVisible(model);
      setSheetHidden(model, [visible(model).at(-1)!], true);
    }
    setSheetHidden(model, model.sheets.flatMap((sheet, index) => (sheet.visibility === "hidden" ? [index] : [])), false);
    return true;
  }],
  ["tab colour", (model) => {
    setTabColor(model, 0, { theme: "accent1", lumMod: 0.75 });
    const coloured = model.sheets.findIndex((sheet, index) => index > 0 && sheet.tabColor !== undefined);
    if (coloured > 0) {
      setTabColor(model, coloured, undefined);
    }
    return true;
  }],
];

describe("package integrity: every ods fixture × every command, written and re-opened", () => {
  for (const build of ODS_BUILDERS) {
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
        const { bytes } = await save(model);
        expect(firstEntry(bytes)).toEqual({ name: "mimetype", stored: true });
        await expectReopensAs(bytes, model);
      });
    }
  }
});
