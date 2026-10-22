// SPEC-121's reference-web fixtures (spec121-links.xlsx / spec121-links.ods,
// built by fixtures/build.ts): what SPEC121_LINKS_XLSX and SPEC121_LINKS_ODS
// record is what the files hold and what the reader gives, every cached
// value is what its formula computes (a full recalculation changes nothing),
// each package is whole (every relationship target and manifest entry is
// present, every part has a content type) and an untouched download
// re-emits every part byte-identical — the ground the sheet-command, writer
// and e2e specs stand on.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE } from "../core/odfXml";
import { parseRelationships, parseXml, resolvePartPath } from "../core/xml";
import { unzip } from "../core/zip";
import { formatA1, getCell, parseA1 } from "./address";
import { spreadsheetDownload } from "./download";
import { buildSpec121LinksOdsBytes, buildSpec121LinksXlsxBytes, SPEC121_LINKS_ODS, SPEC121_LINKS_XLSX } from "./fixtures/build";
import { recalculate } from "./formula/recalc";
import type { Cell, CellValue, Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { tintTransforms } from "./xlsx/themeColor";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const THREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const REVISION_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const CONFIG_NS = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
const CALCEXT_NS = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const CHART_ODF_NS = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The namespaces a Data table's sourceXml needs to parse on its own. */
const ODS_ROOT_NAMESPACES = [
  `xmlns:table="${TABLE_NAMESPACE}"`,
  `xmlns:text="${TEXT_NAMESPACE}"`,
  `xmlns:draw="${DRAW_NAMESPACE}"`,
  `xmlns:xlink="${XLINK_NAMESPACE}"`,
  `xmlns:calcext="${CALCEXT_NS}"`,
  `xmlns:dc="http://purl.org/dc/elements/1.1/"`,
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"`,
  `xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"`,
  `xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"`,
].join(" ");

const decoder = new TextDecoder();
const textOf = (parts: Map<string, Uint8Array>, name: string): string => decoder.decode(parts.get(name));
const xmlOf = (parts: Map<string, Uint8Array>, name: string): Element => parseXml(textOf(parts, name)).documentElement;
const all = (root: Element, ns: string, local: string): Element[] => Array.from(root.getElementsByTagNameNS(ns, local));
const one = (root: Element, ns: string, local: string): Element => {
  const found = all(root, ns, local);
  expect(found, `${local} elements`).toHaveLength(1);
  return found[0]!;
};

function sheetNamed(model: SpreadsheetModel, name: string): Sheet {
  const sheet = model.sheets.find((candidate) => candidate.name === name);
  expect(sheet, name).toBeDefined();
  return sheet!;
}

function cellAt(sheet: Sheet, a1: string): Cell | undefined {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col);
}

/** Every stored cell of a sheet as A1 → value. */
function storedValues(sheet: Sheet): Record<string, CellValue> {
  return Object.fromEntries(sheet.rows.flatMap((row) => row.cells.map((cell) => [formatA1({ row: row.index, col: cell.col }), cell.value])));
}

/** The name and STORE flag of a ZIP's first local entry, read off its header. */
function firstEntry(bytes: Uint8Array): { name: string; stored: boolean } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return { name: decoder.decode(bytes.subarray(30, 30 + view.getUint16(26, true))), stored: view.getUint16(8, true) === 0 };
}

/** The part a `.rels` part belongs to ("" for the package's `_rels/.rels`). */
function relsSource(relsName: string): string {
  const match = /^(.*?)_rels\/(.*)\.rels$/.exec(relsName)!;
  return `${match[1]}${match[2]}`;
}

/** OPC integrity (SPEC-121 "Tests"): every internal relationship target
 * exists, every `.rels` belongs to a part, every part has a content type
 * (a Default by extension or an Override), no Override names a missing
 * part, `sheetId`s and table ids/names are unique. Returns the problems. */
function xlsxIntegrityProblems(parts: Map<string, Uint8Array>): string[] {
  const problems: string[] = [];
  const types = xmlOf(parts, "[Content_Types].xml");
  const defaults = new Set(all(types, CONTENT_TYPES_NS, "Default").map((el) => el.getAttribute("Extension")!.toLowerCase()));
  const overrides = all(types, CONTENT_TYPES_NS, "Override").map((el) => el.getAttribute("PartName")!);
  const lowerNames = new Set([...parts.keys()].map((name) => name.toLowerCase()));
  for (const partName of overrides) {
    if (!lowerNames.has(partName.slice(1).toLowerCase())) {
      problems.push(`an Override names the missing part ${partName}`);
    }
  }
  const overridden = new Set(overrides.map((name) => name.slice(1).toLowerCase()));
  for (const name of parts.keys()) {
    const extension = name.slice(name.lastIndexOf(".") + 1).toLowerCase();
    if (name !== "[Content_Types].xml" && !overridden.has(name.toLowerCase()) && !defaults.has(extension)) {
      problems.push(`${name} has no content type`);
    }
  }
  for (const name of [...parts.keys()].filter((candidate) => candidate.endsWith(".rels"))) {
    const source = relsSource(name);
    if (source !== "" && !parts.has(source)) {
      problems.push(`${name} belongs to the missing part ${source}`);
    }
    for (const rel of parseRelationships(textOf(parts, name)).values()) {
      const target = resolvePartPath(source, rel.target);
      if (rel.targetMode !== "External" && !parts.has(target)) {
        problems.push(`${name} ${rel.id} targets the missing part ${target}`);
      }
    }
  }
  const unique = (label: string, values: (string | null)[]): void => {
    if (new Set(values).size !== values.length) {
      problems.push(`duplicate ${label}: ${values.join(", ")}`);
    }
  };
  unique("sheetIds", all(xmlOf(parts, "xl/workbook.xml"), MAIN_NS, "sheet").map((el) => el.getAttribute("sheetId")));
  const tables = [...parts.keys()].filter((name) => /^xl\/tables\/[^/]+\.xml$/.test(name)).map((name) => xmlOf(parts, name));
  unique("table ids", tables.map((table) => table.getAttribute("id")));
  unique("table names", tables.map((table) => table.getAttribute("name")?.toUpperCase() ?? null));
  return problems;
}

/** ODF package integrity: mimetype first, STOREd and right; every manifest
 * entry exists (a directory entry by its files); every entry but mimetype
 * and the manifest itself is listed; every embedded-object link resolves. */
function odsIntegrityProblems(bytes: Uint8Array, parts: Map<string, Uint8Array>): string[] {
  const problems: string[] = [];
  if (JSON.stringify(firstEntry(bytes)) !== JSON.stringify({ name: "mimetype", stored: true })) {
    problems.push("mimetype is not the first, STOREd entry");
  }
  if (textOf(parts, "mimetype") !== "application/vnd.oasis.opendocument.spreadsheet") {
    problems.push("mimetype is not a spreadsheet's");
  }
  const listed = all(xmlOf(parts, "META-INF/manifest.xml"), MANIFEST_NS, "file-entry").map((el) => el.getAttributeNS(MANIFEST_NS, "full-path")!);
  for (const path of listed.filter((entry) => entry !== "/")) {
    const present = path.endsWith("/") ? [...parts.keys()].some((name) => name.startsWith(path)) : parts.has(path);
    if (!present) {
      problems.push(`the manifest lists the missing ${path}`);
    }
  }
  for (const name of parts.keys()) {
    if (name !== "mimetype" && name !== "META-INF/manifest.xml" && !listed.includes(name)) {
      problems.push(`${name} is not in the manifest`);
    }
  }
  const content = xmlOf(parts, "content.xml");
  for (const link of [...all(content, DRAW_NAMESPACE, "object"), ...all(content, DRAW_NAMESPACE, "image")]) {
    const href = link.getAttributeNS(XLINK_NAMESPACE, "href")!.replace(/^\.\//, "");
    if (!listed.includes(href) && !listed.includes(`${href}/`)) {
      problems.push(`${link.localName} links to ${href}, which the manifest does not list`);
    }
  }
  return problems;
}

async function read(bytes: Uint8Array, format: SpreadsheetFormat): Promise<SpreadsheetModel> {
  return readSpreadsheetBytes(bytes, format, "spec121-links");
}

/** Downloads an untouched model and compares part for part. */
async function expectByteIdenticalDownload(source: Uint8Array, model: SpreadsheetModel, order: readonly string[]): Promise<void> {
  const download = await spreadsheetDownload(model);
  const out = await unzip(download.bytes);
  const original = await unzip(source);
  expect([...original.keys()]).toEqual(order);
  expect([...out.keys()]).toEqual(order);
  for (const [name, bytes] of original) {
    expect(equalBytes(out.get(name)!, bytes), name).toBe(true);
  }
  const stored = JSON.parse(JSON.stringify(model)) as SpreadsheetModel;
  const again = await unzip((await spreadsheetDownload(stored)).bytes);
  for (const [name, bytes] of original) {
    expect(equalBytes(again.get(name)!, bytes), name).toBe(true);
  }
}

describe("spec121-links.xlsx (Excel-shaped)", () => {
  const F = SPEC121_LINKS_XLSX;
  const bytes = buildSpec121LinksXlsxBytes();
  const parts = bytes.then((source) => unzip(source));

  it("reads the five sheets in order — Archive hidden, Vault very hidden — each at its part, Summary active", async () => {
    const model = await read(await bytes, "xlsx");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(F.sheetNames);
    expect(model.sheets.map((sheet) => sheet.visibility)).toEqual(F.sheets.map((sheet) => sheet.visibility));
    expect(model.sheets.map((sheet) => sheet.partName)).toEqual(F.sheets.map((sheet) => sheet.partName));
    expect(model.sheets.every((sheet) => sheet.kind === "grid")).toBe(true);
    expect(model.activeSheetIndex).toBe(F.activeSheetIndex);
    expect(model.calc).toBeUndefined();
  });

  it("reads every stored value, and Summary's formulas as canonical text with their cached values", async () => {
    const model = await read(await bytes, "xlsx");
    for (const [name, cells] of Object.entries(F.values)) {
      expect(storedValues(sheetNamed(model, name)), name).toEqual(cells);
    }
    const summary = sheetNamed(model, "Summary");
    for (const [a1, fact] of Object.entries(F.formulas)) {
      expect(cellAt(summary, a1), a1).toMatchObject({ formula: fact.formula, value: fact.value });
      expect(cellAt(summary, a1)?.formulaUnparsed, a1).toBeUndefined();
    }
    const formulaCells = model.sheets.flatMap((sheet) => sheet.rows.flatMap((row) => row.cells.filter((cell) => cell.formula !== undefined)));
    expect(formulaCells).toHaveLength(Object.keys(F.formulas).length);
  });

  it("reads Data's tab colour as accent2 with Excel's tint as lumMod/lumOff, the tint text kept", async () => {
    const model = await read(await bytes, "xlsx");
    expect(sheetNamed(model, F.tabColor.sheet).tabColor).toEqual({
      theme: F.tabColor.slot,
      ...tintTransforms(Number(F.tabColor.tint)),
      xlsxTint: F.tabColor.tint,
    });
    expect(model.sheets.filter((sheet) => sheet.tabColor !== undefined).map((sheet) => sheet.name)).toEqual([F.tabColor.sheet]);
  });

  it("caches what the formulas compute: a full manual recalculation changes no value", async () => {
    const model = await read(await bytes, "xlsx");
    const rowsBefore = model.sheets.map((sheet) => sheet.rows);
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.stale).toEqual([]);
    expect(result.circular).toEqual([]);
    expect(result.evaluated).toBe(Object.keys(F.formulas).length);
    model.sheets.forEach((sheet, index) => expect(sheet.rows, sheet.name).toBe(rowsBefore[index]));
    expect(model.forceFullCalcOnLoad).toBeUndefined();
  });

  it("downloads untouched: every part byte-identical, in the recorded order", async () => {
    const source = await bytes;
    await expectByteIdenticalDownload(source, await read(source, "xlsx"), F.parts);
  });

  it("is a whole package: targets exist, every part is typed, no Override dangles, ids unique", async () => {
    const source = await parts;
    expect(xlsxIntegrityProblems(source)).toEqual([]);
    // …and the check is not vacuous.
    const broken = new Map(source);
    broken.delete(F.chart.style);
    broken.set("xl/media/image1.png", new Uint8Array([1]));
    expect(xlsxIntegrityProblems(broken)).toEqual([
      `an Override names the missing part /${F.chart.style}`,
      "xl/media/image1.png has no content type",
      `xl/charts/_rels/chart1.xml.rels rId1 targets the missing part ${F.chart.style}`,
    ]);
  });

  it("holds the workbook-level references: sheets, definedNames, calcPr, calcChain, the persons relationship", async () => {
    const source = await parts;
    const workbook = xmlOf(source, "xl/workbook.xml");
    expect(
      all(workbook, MAIN_NS, "sheet").map((el) => ({
        name: el.getAttribute("name"),
        sheetId: Number(el.getAttribute("sheetId")),
        state: el.getAttribute("state"),
        rId: el.getAttributeNS(REL, "id"),
      })),
    ).toEqual(F.sheets.map((sheet) => ({ name: sheet.name, sheetId: sheet.sheetId, state: sheet.visibility === "visible" ? null : sheet.visibility, rId: sheet.rId })));
    expect(one(workbook, MAIN_NS, "workbookView").hasAttribute("activeTab")).toBe(false);
    expect(
      all(workbook, MAIN_NS, "definedName").map((el) => {
        const localSheetId = el.getAttribute("localSheetId");
        return { name: el.getAttribute("name"), ...(localSheetId === null ? {} : { localSheetId: Number(localSheetId) }), text: el.textContent };
      }),
    ).toEqual(F.definedNames);
    expect(textOf(source, "xl/workbook.xml")).toContain(`</definedNames>${F.calcPr}</workbook>`);
    // A chain cell without `i` is on the previous cell's sheet.
    let sheetId = Number.NaN;
    const chain = all(xmlOf(source, "xl/calcChain.xml"), MAIN_NS, "c").map((el) => {
      sheetId = Number(el.getAttribute("i") ?? sheetId);
      return { r: el.getAttribute("r"), i: sheetId };
    });
    expect(chain).toEqual(F.calcChain);
    const workbookRels = parseRelationships(textOf(source, "xl/_rels/workbook.xml.rels"));
    expect(workbookRels.get(F.person.rId)).toEqual({ id: F.person.rId, type: F.person.type, target: "persons/person.xml" });
    expect([...workbookRels.values()].filter((rel) => rel.type.endsWith("/calcChain")).map((rel) => rel.target)).toEqual(["calcChain.xml"]);
  });

  it("lists the sheets and the named ranges in docProps/app.xml", async () => {
    const app = xmlOf(await parts, "docProps/app.xml");
    const pairs = all(one(app, "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "HeadingPairs"), VT_NS, "variant").map(
      (variant) => variant.firstElementChild!.textContent,
    );
    expect(pairs).toEqual(F.headingPairs.flatMap(([heading, count]) => [heading, String(count)]));
    const titles = one(app, "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "TitlesOfParts");
    expect(all(titles, VT_NS, "lpstr").map((el) => el.textContent)).toEqual(F.titlesOfParts);
  });

  it("writes Excel's worksheet shape: mc:Ignorable, xr:uid, tabSelected on Summary alone, Data in CT_Worksheet order", async () => {
    const source = await parts;
    for (const [index, sheet] of F.sheets.entries()) {
      const root = xmlOf(source, sheet.partName);
      expect(root.getAttribute("mc:Ignorable"), sheet.name).toBe("x14ac xr xr2 xr3");
      expect(root.getAttributeNS(REVISION_NS, "uid"), sheet.name).toBe(sheet.uid);
      expect(one(root, MAIN_NS, "sheetView").getAttribute("tabSelected"), sheet.name).toBe(index === F.activeSheetIndex ? "1" : null);
    }
    const data = xmlOf(source, "xl/worksheets/sheet2.xml");
    expect(Array.from(data.children).map((el) => el.localName)).toEqual([
      "sheetPr",
      "dimension",
      "sheetViews",
      "sheetFormatPr",
      "sheetData",
      "conditionalFormatting",
      "dataValidations",
      "hyperlinks",
      "pageMargins",
      "drawing",
      "legacyDrawing",
      "tableParts",
    ]);
    const tabColor = one(data, MAIN_NS, "tabColor");
    expect([tabColor.getAttribute("theme"), tabColor.getAttribute("tint")]).toEqual([String(F.tabColor.theme), F.tabColor.tint]);
  });

  it("holds Data's references to Lookup: the conditional format, the list validation and the internal hyperlink", async () => {
    const data = xmlOf(await parts, "xl/worksheets/sheet2.xml");
    const cf = one(data, MAIN_NS, "conditionalFormatting");
    expect(cf.getAttribute("sqref")).toBe(F.conditionalFormatting.sqref);
    expect(one(cf, MAIN_NS, "cfRule").getAttribute("dxfId")).toBe(String(F.conditionalFormatting.dxfId));
    expect(one(cf, MAIN_NS, "formula").textContent).toBe(F.conditionalFormatting.formula);
    const dv = one(data, MAIN_NS, "dataValidation");
    expect([dv.getAttribute("type"), dv.getAttribute("sqref"), dv.getAttributeNS(REVISION_NS, "uid")]).toEqual(["list", F.dataValidation.sqref, F.dataValidation.uid]);
    expect(one(dv, MAIN_NS, "formula1").textContent).toBe(F.dataValidation.formula1);
    const link = one(data, MAIN_NS, "hyperlink");
    expect([link.getAttribute("ref"), link.getAttribute("location"), link.getAttribute("display"), link.getAttributeNS(REL, "id")]).toEqual([
      F.hyperlink.ref,
      F.hyperlink.location,
      F.hyperlink.display,
      null,
    ]);
    const styles = xmlOf(await parts, "xl/styles.xml");
    expect(one(styles, MAIN_NS, "dxfs").getAttribute("count")).toBe("1");
  });

  it("relates Data to its drawing, VML, table, comments and threaded comment; the drawing to the chart; the chart to its style and colours", async () => {
    const source = await parts;
    const dataRels = parseRelationships(textOf(source, "xl/worksheets/_rels/sheet2.xml.rels"));
    expect(Object.fromEntries([...dataRels.values()].map((rel) => [rel.id, { type: rel.type, target: rel.target }]))).toEqual(F.dataRelationships);
    const data = xmlOf(source, "xl/worksheets/sheet2.xml");
    const referenced = [
      one(data, MAIN_NS, "drawing").getAttributeNS(REL, "id"),
      one(data, MAIN_NS, "legacyDrawing").getAttributeNS(REL, "id"),
      one(data, MAIN_NS, "tablePart").getAttributeNS(REL, "id"),
    ];
    expect(referenced.map((id) => dataRels.get(id!)?.type.split("/").pop())).toEqual(["drawing", "vmlDrawing", "table"]);
    expect(parseRelationships(textOf(source, "xl/drawings/_rels/drawing1.xml.rels")).get(F.drawing.chartRId)?.target).toBe("../charts/chart1.xml");
    expect([...parseRelationships(textOf(source, "xl/charts/_rels/chart1.xml.rels")).values()].map((rel) => resolvePartPath(F.chart.partName, rel.target)).toSorted()).toEqual(
      [F.chart.colors, F.chart.style].toSorted(),
    );
    expect(all(xmlOf(source, F.chart.partName), CHART_NS, "f").map((el) => el.textContent)).toEqual(F.chart.formulas);
  });

  it("pairs the threaded comment with its legacy comment, VML shape and the workbook's person", async () => {
    const source = await parts;
    const legacy = one(xmlOf(source, F.comment.partName), MAIN_NS, "comment");
    expect([legacy.getAttribute("ref"), legacy.getAttributeNS(REVISION_NS, "uid")]).toEqual([F.comment.ref, F.comment.uid]);
    expect(one(xmlOf(source, F.comment.partName), MAIN_NS, "author").textContent).toBe(F.comment.author);
    const thread = one(xmlOf(source, F.threadedComment.partName), THREADED_NS, "threadedComment");
    expect([thread.getAttribute("ref"), thread.getAttribute("id"), thread.getAttribute("personId")]).toEqual([
      F.threadedComment.ref,
      F.threadedComment.id,
      F.threadedComment.personId,
    ]);
    expect(thread.getAttribute("id")).toBe(legacy.getAttributeNS(REVISION_NS, "uid"));
    expect(one(xmlOf(source, F.person.partName), THREADED_NS, "person").getAttribute("id")).toBe(F.person.id);
    const vml = textOf(source, F.vml.partName);
    expect(vml).toContain(`<o:idmap v:ext="edit" data="${F.vml.idmap}"/>`);
    expect([...vml.matchAll(/\b(?:id|o:spid)="(_x0000_s\d+)"/g)].map((match) => match[1])).toEqual(F.vml.shapeIds.flatMap((id) => [id, id]));
    expect(vml).toContain(`<x:Row>1</x:Row>\r\n   <x:Column>1</x:Column>`);
  });

  it("declares Table1 over Data!A1:B5, its header row the table's column names", async () => {
    const source = await parts;
    const table = xmlOf(source, F.table.partName);
    expect({
      id: Number(table.getAttribute("id")),
      name: table.getAttribute("name"),
      displayName: table.getAttribute("displayName"),
      ref: table.getAttribute("ref"),
      uid: table.getAttributeNS(REVISION_NS, "uid"),
    }).toEqual({ id: F.table.id, name: F.table.name, displayName: F.table.displayName, ref: F.table.ref, uid: F.table.uid });
    expect(all(table, MAIN_NS, "tableColumn").map((el) => el.getAttribute("name"))).toEqual(F.table.columns);
    expect(one(table, MAIN_NS, "tableStyleInfo").getAttribute("name")).toBe("TableStyleMedium2");
    const model = await read(await bytes, "xlsx");
    const data = sheetNamed(model, "Data");
    expect(["A1", "B1"].map((a1) => cellAt(data, a1)?.value)).toEqual(F.table.columns);
  });
});

describe("spec121-links.ods (LibreOffice-shaped)", () => {
  const O = SPEC121_LINKS_ODS;
  const bytes = buildSpec121LinksOdsBytes();
  const parts = bytes.then((source) => unzip(source));

  it("reads the four tables in order — Archive hidden by its table style — Summary active, Data's tab colour", async () => {
    const model = await read(await bytes, "ods");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(O.tableNames);
    expect(model.sheets.map((sheet) => sheet.visibility)).toEqual(O.visibility);
    expect(model.activeSheetIndex).toBe(O.activeSheetIndex);
    expect(model.sheets.map((sheet) => sheet.tabColor)).toEqual(O.tableNames.map((name) => (name === O.tabColor.table ? O.tabColor.hex : undefined)));
    expect(model.sheets.map((sheet) => /table:style-name="([^"]+)"/.exec(sheet.sourceXml!)?.[1])).toEqual(O.tableStyles);
  });

  it("reads every stored value, and Summary's of:= formulas as canonical text with their cached values", async () => {
    const model = await read(await bytes, "ods");
    for (const [name, cells] of Object.entries(O.values)) {
      expect(storedValues(sheetNamed(model, name)), name).toEqual(cells);
    }
    const summary = sheetNamed(model, "Summary");
    for (const [a1, fact] of Object.entries(O.formulas)) {
      expect(cellAt(summary, a1), a1).toMatchObject({ formula: fact.formula, value: fact.value });
      expect(cellAt(summary, a1)?.formulaUnparsed, a1).toBeUndefined();
      expect(summary.sourceXml, a1).toContain(`table:formula="${fact.wire}"`);
    }
  });

  it("caches what the formulas compute: a full manual recalculation changes no value", async () => {
    const model = await read(await bytes, "ods");
    const rowsBefore = model.sheets.map((sheet) => sheet.rows);
    const result = recalculate(model, [], { full: true, reason: "manual" });
    expect(result.valueChanged).toEqual([]);
    expect(result.stale).toEqual([]);
    expect(result.circular).toEqual([]);
    expect(result.evaluated).toBe(Object.keys(O.formulas).length);
    model.sheets.forEach((sheet, index) => expect(sheet.rows, sheet.name).toBe(rowsBefore[index]));
  });

  it("downloads untouched: every part byte-identical in the recorded order, mimetype first and STOREd; the envelope splices back", async () => {
    const source = await bytes;
    const model = await read(source, "ods");
    await expectByteIdenticalDownload(source, model, O.parts);
    expect(firstEntry((await spreadsheetDownload(model)).bytes)).toEqual({ name: "mimetype", stored: true });
    const envelope = model.odsContentEnvelope!;
    expect(envelope.prefix + model.sheets.map((sheet) => sheet.sourceXml).join("") + envelope.suffix).toBe(textOf(await parts, "content.xml"));
  });

  it("is a whole package: the manifest and the entries agree, the chart object and its replacement resolve", async () => {
    const source = await parts;
    expect(odsIntegrityProblems(await bytes, source)).toEqual([]);
    const broken = new Map(source);
    broken.delete(O.chart.replacement);
    broken.set("Pictures/stray.png", new Uint8Array([1]));
    expect(odsIntegrityProblems(await bytes, broken)).toEqual([
      `the manifest lists the missing ${O.chart.replacement}`,
      "Pictures/stray.png is not in the manifest",
    ]);
    const manifest = xmlOf(source, "META-INF/manifest.xml");
    const objectEntry = all(manifest, MANIFEST_NS, "file-entry").find((el) => el.getAttributeNS(MANIFEST_NS, "full-path") === "Object 1/");
    expect(objectEntry?.getAttributeNS(MANIFEST_NS, "media-type")).toBe("application/vnd.oasis.opendocument.chart");
  });

  it("puts the validation in the prefix and the named range, named expression and database range in the suffix", async () => {
    const model = await read(await bytes, "ods");
    const { prefix, suffix } = model.odsContentEnvelope!;
    const content = xmlOf(await parts, "content.xml");
    const validation = one(content, TABLE_NAMESPACE, "content-validation");
    expect([validation.getAttributeNS(TABLE_NAMESPACE, "name"), validation.getAttributeNS(TABLE_NAMESPACE, "condition"), validation.getAttributeNS(TABLE_NAMESPACE, "base-cell-address")]).toEqual([
      O.contentValidation.name,
      O.contentValidation.condition,
      O.contentValidation.baseCellAddress,
    ]);
    expect(prefix).toContain("<table:content-validations>");
    const ranges = all(content, TABLE_NAMESPACE, "named-range").map((el) => ({
      name: el.getAttributeNS(TABLE_NAMESPACE, "name"),
      baseCellAddress: el.getAttributeNS(TABLE_NAMESPACE, "base-cell-address"),
      cellRangeAddress: el.getAttributeNS(TABLE_NAMESPACE, "cell-range-address"),
    }));
    const { table: localTable, ...local } = O.localNamedRange;
    expect(ranges).toEqual([local, O.namedRange]);
    const expression = one(content, TABLE_NAMESPACE, "named-expression");
    expect({
      name: expression.getAttributeNS(TABLE_NAMESPACE, "name"),
      baseCellAddress: expression.getAttributeNS(TABLE_NAMESPACE, "base-cell-address"),
      expression: expression.getAttributeNS(TABLE_NAMESPACE, "expression"),
    }).toEqual(O.namedExpression);
    const database = one(content, TABLE_NAMESPACE, "database-range");
    expect([database.getAttributeNS(TABLE_NAMESPACE, "name"), database.getAttributeNS(TABLE_NAMESPACE, "target-range-address")]).toEqual([
      O.databaseRange.name,
      O.databaseRange.targetRangeAddress,
    ]);
    expect(suffix).toContain(`table:name="${O.namedRange.name}"`);
    expect(suffix).toContain(`table:name="${O.namedExpression.name}"`);
    expect(suffix).toContain("<table:database-ranges>");
    expect(sheetNamed(model, localTable).sourceXml).toContain(`table:name="${local.name}"`);
  });

  it("holds Data's drawing layer and cell-level references: the chart frame, the annotation, the link, the validated cells, the conditional format", async () => {
    const model = await read(await bytes, "ods");
    const data = sheetNamed(model, "Data");
    const table = parseXml(`<root xmlns:office="${OFFICE_NAMESPACE}" ${ODS_ROOT_NAMESPACES}>${data.sourceXml}</root>`).documentElement;
    const frame = one(one(table, TABLE_NAMESPACE, "shapes"), DRAW_NAMESPACE, "frame");
    expect({
      xmlId: frame.getAttributeNS(XML_NS, "id"),
      drawId: frame.getAttributeNS(DRAW_NAMESPACE, "id"),
      name: frame.getAttributeNS(DRAW_NAMESPACE, "name"),
      endCellAddress: frame.getAttributeNS(TABLE_NAMESPACE, "end-cell-address"),
    }).toEqual(O.chart.frame);
    const object = one(frame, DRAW_NAMESPACE, "object");
    expect([object.getAttributeNS(DRAW_NAMESPACE, "notify-on-update-of-ranges"), object.getAttributeNS(XLINK_NAMESPACE, "href")]).toEqual([
      O.chart.notifyOnUpdateOfRanges,
      O.chart.href,
    ]);
    expect(one(frame, DRAW_NAMESPACE, "image").getAttributeNS(XLINK_NAMESPACE, "href")).toBe(O.chart.replacementHref);
    const annotation = one(table, OFFICE_NAMESPACE, "annotation");
    expect(annotation.textContent).toBe(`${O.annotation.creator}${O.annotation.date}${O.annotation.text}`);
    expect(cellAt(data, O.annotation.a1)?.value).toBe(O.values.Data.B2);
    const link = one(table, TEXT_NAMESPACE, "a");
    expect([link.getAttributeNS(XLINK_NAMESPACE, "href"), link.textContent]).toEqual([O.link.href, O.link.text]);
    expect(cellAt(data, O.link.a1)?.value).toBe(O.link.text);
    expect(all(table, TABLE_NAMESPACE, "table-cell").filter((el) => el.getAttributeNS(TABLE_NAMESPACE, "content-validation-name") === O.contentValidation.name)).toHaveLength(
      O.contentValidation.cells.length,
    );
    const format = one(table, CALCEXT_NS, "conditional-format");
    const condition = one(format, CALCEXT_NS, "condition");
    expect({
      table: "Data",
      targetRangeAddress: format.getAttributeNS(CALCEXT_NS, "target-range-address"),
      value: condition.getAttributeNS(CALCEXT_NS, "value"),
      baseCellAddress: condition.getAttributeNS(CALCEXT_NS, "base-cell-address"),
      applyStyleName: condition.getAttributeNS(CALCEXT_NS, "apply-style-name"),
    }).toEqual(O.conditionalFormat);
    expect(textOf(await parts, "styles.xml")).toContain(`style:name="${O.conditionalFormat.applyStyleName}"`);
    expect(model.sheets.filter((sheet) => /<table:shapes|<office:annotation/.test(sheet.sourceXml!)).map((sheet) => sheet.name)).toEqual(["Data"]);
  });

  it("names Data in the chart object's plot area, categories, series values and label", async () => {
    const chart = xmlOf(await parts, "Object 1/content.xml");
    expect(one(chart, CHART_ODF_NS, "plot-area").getAttributeNS(TABLE_NAMESPACE, "cell-range-address")).toBe(O.chart.plotAreaRange);
    expect(one(chart, CHART_ODF_NS, "categories").getAttributeNS(TABLE_NAMESPACE, "cell-range-address")).toBe(O.chart.categoriesRange);
    const series = one(chart, CHART_ODF_NS, "series");
    expect([series.getAttributeNS(CHART_ODF_NS, "values-cell-range-address"), series.getAttributeNS(CHART_ODF_NS, "label-cell-address")]).toEqual([
      O.chart.valuesRange,
      O.chart.labelCell,
    ]);
  });

  it("keeps ActiveTable and one settings.xml entry per table, Data's carrying TabColor", async () => {
    const settings = xmlOf(await parts, "settings.xml");
    const items = all(settings, CONFIG_NS, "config-item");
    expect(items.find((el) => el.getAttributeNS(CONFIG_NS, "name") === "ActiveTable")?.textContent).toBe(O.activeTable);
    const tables = all(settings, CONFIG_NS, "config-item-map-named").find((el) => el.getAttributeNS(CONFIG_NS, "name") === "Tables")!;
    const entries = Array.from(tables.children).map((entry) => ({
      name: entry.getAttributeNS(CONFIG_NS, "name"),
      tabColor: Array.from(entry.children).find((item) => item.getAttributeNS(CONFIG_NS, "name") === "TabColor")?.textContent,
    }));
    expect(entries).toEqual(O.settingsTables.map((name) => ({ name, tabColor: name === O.tabColor.table ? String(O.tabColor.settingsValue) : undefined })));
    expect(O.tabColor.settingsValue).toBe(Number.parseInt(O.tabColor.hex.slice(1), 16));
  });
});

