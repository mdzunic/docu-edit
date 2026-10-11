// SPEC-124 "Retained references", the .xlsx table: every row moved by a
// structural edit through structuralMapper — the adjusted part differing from
// its source only inside the expected attribute values and text nodes (the
// span-diff helper), and inside nothing else — plus the removals (an emptied
// range list takes its element, the containers that cannot stay empty, a
// legacy comment's VML shape), source-time sheet names (a renamed sheet, a
// duplicated sheet's override, another sheet's references left alone),
// copy-on-write identity, the returned set of changed sheets, and the
// worksheet tail path (the cells are never decoded or scanned).
import { describe, expect, it } from "vitest";
import { base64ToBytes, bytesToBase64, type OpaquePart } from "../../core/packageParts";
import { relsPathFor } from "../../core/xml";
import { formatRange } from "../address";
import { changedSpans } from "../fixtures/spanDiff";
import type { Sheet, SpreadsheetModel } from "../model";
import type { RetainedReferenceMapper } from "../retainedReferences";
import { structuralMapper, type StructuralEdit } from "../structure";
import { rewriteXlsxRetainedReferences } from "./retainedReferences";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM = "http://schemas.microsoft.com/office/excel/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DML = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const THREADED = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const THREADED_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

const SHEET1 = "xl/worksheets/sheet1.xml";
const SHEET2 = "xl/worksheets/sheet2.xml";
const SHEET3 = "xl/worksheets/sheet3.xml";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8", { ignoreBOM: true });
const toData = (text: string): string => bytesToBase64(encoder.encode(text));

// --- hand-built packages -------------------------------------------------------------------

type Rel = readonly [id: string, type: string, target: string];

function relsXml(rels: readonly Rel[]): string {
  const items = rels.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type.includes("/") ? type : `${R}/${type}`}" Target="${target}"/>`);
  return `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join("")}</Relationships>`;
}

/** Cells, dimension, view, columns: modelled — never adjusted here. */
const CELLS = `<sheetData><row r="5" spans="1:2"><c r="A5"><f>A1+B5</f><v>3</v></c><c r="B5" t="inlineStr"><is><t>Über €</t></is></c></row></sheetData>`;

function worksheet(tail = "", cells = CELLS): string {
  return (
    `${DECL}<worksheet xmlns="${MAIN}" xmlns:r="${R}"><dimension ref="A1:B5"/><sheetViews><sheetView workbookViewId="0">` +
    `<selection activeCell="A5" sqref="A5"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
    `<cols><col min="1" max="2" width="12" customWidth="1"/></cols>${cells}${tail}</worksheet>`
  );
}

interface SheetSpec {
  name: string;
  /** The name the source workbook gives it (default `name`). */
  sourceName?: string;
  xml?: string;
  rels?: readonly Rel[];
  /** A chartsheet (foreign) instead of a worksheet. */
  chartsheet?: true;
}

/** Sheets s1, s2, … over parts sheet1.xml, sheet2.xml, … in source order. */
function book(specs: readonly SheetSpec[], parts: Readonly<Record<string, string>> = {}, workbookTail = ""): SpreadsheetModel {
  const partNames = specs.map((spec, i) => `xl/${spec.chartsheet ? "chartsheets" : "worksheets"}/sheet${i + 1}.xml`);
  const sheetsXml = specs.map((spec, i) => `<sheet name="${spec.sourceName ?? spec.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
  const files: Record<string, string> = {
    "[Content_Types].xml": `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    "_rels/.rels": relsXml([["rId1", "officeDocument", "xl/workbook.xml"]]),
    "xl/workbook.xml": `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets>${sheetsXml}</sheets>${workbookTail}<calcPr calcId="191029"/></workbook>`,
    "xl/_rels/workbook.xml.rels": relsXml(specs.map((spec, i): Rel => [`rId${i + 1}`, spec.chartsheet ? "chartsheet" : "worksheet", partNames[i]!.slice(3)])),
  };
  specs.forEach((spec, i) => {
    files[partNames[i]!] = spec.xml ?? worksheet();
    if (spec.rels) {
      files[relsPathFor(partNames[i]!)] = relsXml(spec.rels);
    }
  });
  Object.assign(files, parts);
  return {
    dxfs: [],
    kind: "spreadsheet",
    id: "retained",
    title: "Retained references",
    format: "xlsx",
    readerVersion: 1,
    sheets: specs.map((spec, i) => sheetOf(`s${i + 1}`, spec.name, partNames[i]!, spec.chartsheet ? "foreign" : "grid")),
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: Object.entries(files).map(([name, text]) => ({ name, data: toData(text) })),
    styles: [{}],
    defaultStyleId: 0,
  };
}

function sheetOf(id: string, name: string, partName: string, kind: Sheet["kind"] = "grid", extra: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], id, name, partName, rows: [], columns: [], merges: [], visibility: "visible", kind, defaultColWidthPx: 64, defaultRowHeightPt: 15, ...extra };
}

function textOf(opaque: readonly OpaquePart[], name: string): string {
  const part = opaque.find((each) => each.name === name);
  if (!part) {
    throw new Error(`no part ${name}`);
  }
  return decoder.decode(base64ToBytes(part.data));
}

function partOf(model: SpreadsheetModel, name: string): OpaquePart | undefined {
  return model.opaque.find((each) => each.name === name);
}

/** One structural edit's mapper over the model's retained parts. */
function adjust(model: SpreadsheetModel, kind: StructuralEdit["kind"], at: number, count = 1, sheetId = "s1"): Set<string> {
  return rewriteXlsxRetainedReferences(model, structuralMapper(model, { kind, sheetId, at, count }));
}

/** The [before, after] values a part's adjustment changed — the span diff
 * throws when anything else (with `removed` taken out of the source) differs. */
function spans(before: readonly OpaquePart[], model: SpreadsheetModel, name: string, removed: readonly string[] = []): [string, string][] {
  return changedSpans(textOf(before, name), textOf(model.opaque, name), removed).map((span) => [span.before, span.after]);
}

/** Whether the part is the very object it was. */
function untouched(before: readonly OpaquePart[], model: SpreadsheetModel, name: string): boolean {
  return partOf(model, name) === before.find((each) => each.name === name);
}

// --- part builders -------------------------------------------------------------------------

const ext = (uri: string, body: string): string => `<ext uri="${uri}" xmlns:x14="${X14}">${body}</ext>`;
const X14_CF = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
const X14_DV = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";
const X14_SPARKLINES = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

function x14Cf(id: number, sqref: string, f?: string): string {
  const min = f === undefined ? `<x14:cfvo type="autoMin"/>` : `<x14:cfvo type="num"><xm:f>${f}</xm:f></x14:cfvo>`;
  return (
    `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="dataBar" id="{0000000${id}-0000-4000-8000-000000000000}">` +
    `<x14:dataBar minLength="0" maxLength="100">${min}<x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`
  );
}

function x14Dv(sqref: string, formula1: string, formula2?: string): string {
  const second = formula2 === undefined ? "" : `<x14:formula2><xm:f>${formula2}</xm:f></x14:formula2>`;
  return `<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>${formula1}</xm:f></x14:formula1>${second}<xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`;
}

const sparkline = (f: string, sqref: string): string => `<x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${sqref}</xm:sqref></x14:sparkline>`;

function sparklineGroup(items: string, dateAxis?: string): string {
  const axis = dateAxis === undefined ? "" : `<xm:f>${dateAxis}</xm:f>`;
  return `<x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries theme="4"/><x14:colorNegative theme="5"/>${axis}<x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`;
}

const corner = (tag: "from" | "to", col: number, row: number): string =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
const SHAPE = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/>`;
const drawingXml = (anchors: string): string => `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${DML}">${anchors}</xdr:wsDr>`;
/** An absolute anchor holding chart rId1 — nothing in it moves. */
const CHART_FRAME =
  `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="4572000" cy="2743200"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
  `<xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
  `<a:graphic><a:graphicData uri="${CHART}"><c:chart xmlns:c="${CHART}" xmlns:r="${R}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>`;

function chartXml(formulas: readonly string[]): string {
  const series = formulas.map(
    (f, i) =>
      `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val><c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache></c:numRef></c:val></c:ser>`,
  );
  return `${DECL}<c:chartSpace xmlns:c="${CHART}" xmlns:a="${DML}" xmlns:r="${R}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/>${series.join("")}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
}

/** A drawing relating one chart, for a sheet's rels. */
function chartParts(drawing: string, chart: string, formulas: readonly string[]): Record<string, string> {
  return {
    [drawing]: drawingXml(CHART_FRAME),
    [relsPathFor(drawing)]: relsXml([["rId1", "chart", `../charts/${chart.slice(chart.lastIndexOf("/") + 1)}`]]),
    [chart]: chartXml(formulas),
  };
}

const comment = (ref: string, text: string): string => `<comment ref="${ref}" authorId="0" shapeId="0"><text><r><t>${text}</t></r></text></comment>`;
const commentsXml = (items: string): string => `${DECL}<comments xmlns="${MAIN}"><authors><author>Ann</author></authors><commentList>${items}</commentList></comments>`;

function threaded(ref: string, id: string, parent?: string): string {
  const reply = parent === undefined ? "" : ` parentId="{${parent}}"`;
  return `<threadedComment ref="${ref}" dT="2026-01-02T03:04:05.00" personId="{P-1}" id="{${id}}"${reply}><text>${id}</text></threadedComment>`;
}
const threadedXml = (items: string): string => `${DECL}<ThreadedComments xmlns="${THREADED}" xmlns:x="${MAIN}">${items}</ThreadedComments>`;

/** A legacy note's shape, as Excel writes it (the anchor on its own line). */
function noteShape(id: number, anchor: string, row: number, col: number): string {
  return (
    `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style='position:absolute;\n  margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;\n  visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">` +
    `\n  <v:fill color2="#ffffe1"/>\n  <v:shadow on="t" color="black" obscured="t"/>\n  <v:path o:connecttype="none"/>\n  <v:textbox style='mso-direction-alt:auto'>\n   <div style='text-align:left'></div>\n  </v:textbox>` +
    `\n  <x:ClientData ObjectType="Note">\n   <x:MoveWithCells/>\n   <x:SizeWithCells/>\n   <x:Anchor>\n    ${anchor}</x:Anchor>\n   <x:AutoFill>False</x:AutoFill>` +
    `\n   <x:Row>${row}</x:Row>\n   <x:Column>${col}</x:Column>\n  </x:ClientData>\n </v:shape>`
  );
}

function vmlXml(shapes: string): string {
  return (
    `<xml xmlns:v="urn:schemas-microsoft-com:vml"\n xmlns:o="urn:schemas-microsoft-com:office:office"\n xmlns:x="urn:schemas-microsoft-com:office:excel">\n <o:shapelayout v:ext="edit">\n  <o:idmap v:ext="edit" data="1"/>\n </o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"\n  path="m,l,21600r21600,l21600,xe">\n  <v:stroke joinstyle="miter"/>\n  <v:path gradientshapeok="t" o:connecttype="rect"/>\n </v:shapetype>${shapes}</xml>`
  );
}

const NOTE_RELS: readonly Rel[] = [
  ["rId1", "comments", "../comments1.xml"],
  ["rId2", "vmlDrawing", "../drawings/vmlDrawing1.vml"],
  ["rId3", THREADED_REL, "../threadedComments/threadedComment1.xml"],
];
const COMMENTS = "xl/comments1.xml";
const VML = "xl/drawings/vmlDrawing1.vml";
const THREADED_PART = "xl/threadedComments/threadedComment1.xml";

const pivotCacheXml = (source: string): string =>
  `${DECL}<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${R}" r:id="rId1" refreshOnLoad="1"><cacheSource type="worksheet">${source}</cacheSource>` +
  `<cacheFields count="1"><cacheField name="a" numFmtId="0"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>`;

// --- the sheet part --------------------------------------------------------------------------

describe("the sheet part", () => {
  it("conditionalFormatting@sqref and cfRule/formula move; cells, merges, dimension, views and columns do not", () => {
    const cf =
      `<mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells><conditionalFormatting sqref="A1:A10 $C$5"><cfRule type="expression" dxfId="0" priority="1">` +
      `<formula>$A3&gt;Sheet2!$B$1</formula><formula>B5</formula></cfRule></conditionalFormatting>`;
    const rows = book([{ name: "Sheet1", xml: worksheet(cf) }, { name: "Sheet2" }]);
    const before = rows.opaque;
    expect(adjust(rows, "insertRows", 2, 2)).toEqual(new Set(["s1"]));
    expect(spans(before, rows, SHEET1)).toEqual([
      ["A1:A10 $C$5", "A1:A12 $C$7"],
      ["$A3&gt;Sheet2!$B$1", "$A5&gt;Sheet2!$B$1"],
      ["B5", "B7"],
    ]);

    const cols = book([{ name: "Sheet1", xml: worksheet(cf) }, { name: "Sheet2" }]);
    const colsBefore = cols.opaque;
    adjust(cols, "insertCols", 0, 1);
    expect(spans(colsBefore, cols, SHEET1)).toEqual([
      ["A1:A10 $C$5", "B1:B10 $D$5"],
      ["$A3&gt;Sheet2!$B$1", "$B3&gt;Sheet2!$B$1"],
      ["B5", "C5"],
    ]);
  });

  it("a range list keeps its separators, `$` markers, entry shapes and unparseable entries; a dropped entry takes the separator before it", () => {
    const model = book([
      { name: "Sheet1", xml: worksheet(`<conditionalFormatting sqref="A1  $B$3:B$9 XFE1 A3 C2:C3"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>`) },
    ]);
    const before = model.opaque;
    adjust(model, "deleteRows", 2, 1);
    expect(spans(before, model, SHEET1)).toEqual([["A1  $B$3:B$9 XFE1 A3 C2:C3", "A1  $B$3:B$8 XFE1 C2:C2"]]);
  });

  it("dataValidation@sqref, formula1 and formula2", () => {
    const model = book([
      {
        name: "Sheet1",
        xml: worksheet(
          `<dataValidations count="2"><dataValidation type="whole" operator="between" allowBlank="1" sqref="B2:B20"><formula1>$D$1</formula1>` +
            `<formula2>$D$30</formula2></dataValidation><dataValidation type="list" sqref="E5"><formula1>"yes,no"</formula1></dataValidation></dataValidations>`,
        ),
      },
    ]);
    const before = model.opaque;
    adjust(model, "deleteRows", 0, 2);
    expect(spans(before, model, SHEET1)).toEqual([
      ["B2:B20", "B1:B18"],
      ["$D$1", "#REF!"],
      ["$D$30", "$D$28"],
      ["E5", "E3"],
    ]);
  });

  it("SPEC-128: leaves the root-level autoFilter, its sortState and the worksheet sortState to the model; a customSheetView's autoFilter@ref, sortState@ref and sortCondition@ref still move, and an area gone leaves them", () => {
    // The root-level elements are modelled (the autofilter adjuster moves
    // them, the writer serialises them): moving them here too would move
    // them twice.
    const filter =
      `<autoFilter ref="A1:D30"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn>` +
      `<sortState ref="A2:D30"><sortCondition descending="1" ref="B2:B30"/></sortState></autoFilter>` +
      `<sortState ref="F2:G30"><sortCondition ref="G2:G30"/></sortState>`;
    for (const [kind, at, count] of [
      ["insertRows", 9, 3],
      ["insertCols", 1, 1],
      ["deleteCols", 0, 4],
    ] as const) {
      const model = book([{ name: "Sheet1", xml: worksheet(filter) }]);
      const before = model.opaque;
      expect(adjust(model, kind, at, count), kind).toEqual(new Set());
      expect(model.opaque, kind).toBe(before);
    }

    const view =
      `<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}"><selection sqref="A1"/>` +
      `<autoFilter ref="A1:D30"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn>` +
      `<sortState ref="A2:D30"><sortCondition descending="1" ref="B2:B30"/></sortState></autoFilter></customSheetView></customSheetViews>`;
    const rows = book([{ name: "Sheet1", xml: worksheet(view) }]);
    const rowsBefore = rows.opaque;
    adjust(rows, "insertRows", 9, 3);
    expect(spans(rowsBefore, rows, SHEET1)).toEqual([
      ["A1:D30", "A1:D33"],
      ["A2:D30", "A2:D33"],
      ["B2:B30", "B2:B33"],
    ]);

    const cols = book([{ name: "Sheet1", xml: worksheet(view) }]);
    const colsBefore = cols.opaque;
    adjust(cols, "insertCols", 1, 1);
    expect(spans(colsBefore, cols, SHEET1)).toEqual([
      ["A1:D30", "A1:E30"],
      ["A2:D30", "A2:E30"],
      ["B2:B30", "C2:C30"],
    ]);

    const gone = book([{ name: "Sheet1", xml: worksheet(view) }]);
    const goneBefore = gone.opaque;
    expect(adjust(gone, "deleteCols", 0, 4)).toEqual(new Set());
    expect(gone.opaque).toBe(goneBefore);
  });

  it("hyperlink@ref", () => {
    const model = book([
      {
        name: "Sheet1",
        xml: worksheet(`<hyperlinks><hyperlink ref="A5" r:id="rId1" display="site"/><hyperlink ref="B2:C3" location="Sheet1!A9" display="inside"/></hyperlinks>`),
      },
    ]);
    const before = model.opaque;
    adjust(model, "insertRows", 3, 1);
    expect(spans(before, model, SHEET1)).toEqual([["A5", "A6"]]);
  });

  it("rowBreaks/brk@id moves with rows, colBreaks/brk@id with columns; a break in a deleted band goes to its start", () => {
    const breaks =
      `<rowBreaks count="2" manualBreakCount="2"><brk id="10" max="16383" man="1"/><brk id="20" max="16383" man="1"/></rowBreaks>` +
      `<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>`;
    const cases: [StructuralEdit["kind"], number, number, [string, string][]][] = [
      ["insertRows", 15, 2, [["20", "22"]]],
      ["insertCols", 0, 1, [["3", "4"]]],
      [
        "deleteRows",
        9,
        2,
        [
          ["10", "9"],
          ["20", "18"],
        ],
      ],
    ];
    for (const [kind, at, count, expected] of cases) {
      const model = book([{ name: "Sheet1", xml: worksheet(breaks) }]);
      const before = model.opaque;
      adjust(model, kind, at, count);
      expect(spans(before, model, SHEET1), kind).toEqual(expected);
    }
  });

  it("protectedRange@sqref, inputCells@r, dataRef@ref (on its @sheet, else the host), cellWatch@r and ignoredError@sqref", () => {
    const model = book([
      {
        name: "Sheet1",
        xml: worksheet(
          `<protectedRanges><protectedRange sqref="B2:C5 E1" name="Inputs"/></protectedRanges>` +
            `<scenarios current="0" show="0"><scenario name="Best" count="1" user="me"><inputCells r="C3" val="10"/></scenario></scenarios>` +
            `<dataConsolidate><dataRefs count="3"><dataRef ref="A1:C9" sheet="Sheet1"/><dataRef ref="A1:C9" sheet="Sheet2"/><dataRef ref="B1:B4"/></dataRefs></dataConsolidate>` +
            `<cellWatches><cellWatch r="D4"/></cellWatches><ignoredErrors><ignoredError sqref="B1:B30" numberStoredAsText="1"/></ignoredErrors>`,
        ),
      },
      { name: "Sheet2" },
    ]);
    const before = model.opaque;
    adjust(model, "insertCols", 1, 2);
    expect(spans(before, model, SHEET1)).toEqual([
      ["B2:C5 E1", "D2:E5 G1"],
      ["C3", "E3"],
      ["A1:C9", "A1:E9"],
      ["B1:B4", "D1:D4"],
      ["D4", "F4"],
      ["B1:B30", "D1:D30"],
    ]);
  });

  it("the extLst: x14:conditionalFormatting's xm:sqref and xm:f, x14:dataValidation's xm:sqref and formula1/2, x14:sparkline's xm:f and xm:sqref, a group's date axis", () => {
    const model = book([
      {
        name: "Sheet1",
        xml: worksheet(
          `<extLst>${ext(X14_CF, `<x14:conditionalFormattings>${x14Cf(1, "A2:A20", "$H$3")}</x14:conditionalFormattings>`)}` +
            `${ext(X14_DV, `<x14:dataValidations count="1" xmlns:xm="${XM}">${x14Dv("C3:C9 E4", "Sheet2!$A$1:$A$5", "$H$4")}</x14:dataValidations>`)}` +
            `${ext(
              X14_SPARKLINES,
              `<x14:sparklineGroups xmlns:xm="${XM}">${sparklineGroup(sparkline("Sheet1!B2:E2", "F2") + sparkline("Sheet1!B3:E3", "F3"), "Sheet1!$B$12:$E$12")}</x14:sparklineGroups>`,
            )}</extLst>`,
        ),
      },
      { name: "Sheet2" },
    ]);
    const before = model.opaque;
    expect(adjust(model, "insertRows", 2, 1)).toEqual(new Set(["s1"]));
    expect(spans(before, model, SHEET1)).toEqual([
      ["$H$3", "$H$4"],
      ["A2:A20", "A2:A21"],
      ["$H$4", "$H$5"],
      ["C3:C9 E4", "C4:C10 E5"],
      ["Sheet1!$B$12:$E$12", "Sheet1!$B$13:$E$13"],
      ["Sheet1!B3:E3", "Sheet1!B4:E4"],
      ["F3", "F4"],
    ]);
  });
});

// --- workbook-level and related parts ----------------------------------------------------------

describe("the workbook and the parts a sheet relates", () => {
  const NAMES =
    `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$D$30</definedName>` +
    `<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$30</definedName><definedName name="Here" localSheetId="1">$A$5</definedName>` +
    `<definedName name="There" localSheetId="0">$A$5</definedName><definedName name="Total">SUM(Sheet1!$B$2:$B$30)</definedName>` +
    `<definedName name="Other">Sheet2!$A$1:$A$9</definedName><definedName name="Bare">$A$5</definedName></definedNames>`;

  it("definedName text, a localSheetId naming its host in the SOURCE workbook's sheet order; SPEC-128: never the model's _xlnm._FilterDatabase", () => {
    // Sheet1 has no autofilter, in its part or the model: its _xlnm._FilterDatabase
    // (an Advanced Filter's) is an ordinary name and moves like the rest.
    const model = book([{ name: "Sheet1" }, { name: "Sheet2" }], {}, NAMES);
    model.sheets = [model.sheets[1]!, model.sheets[0]!]; // moved: the model's order is no longer the source's
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 2, "s1")).toEqual(new Set());
    expect(spans(before, model, "xl/workbook.xml")).toEqual([
      ["Sheet1!$A$1:$D$30", "Sheet1!$A$3:$D$32"],
      ["Sheet1!$A$1:$D$30", "Sheet1!$A$3:$D$32"],
      ["$A$5", "$A$7"],
      ["SUM(Sheet1!$B$2:$B$30)", "SUM(Sheet1!$B$4:$B$32)"],
    ]);
    // With an autofilter in Sheet1's part (or the model) the name is the
    // model's: it stays, and the writer reconciles it with the autofilter.
    const filtered = book([{ name: "Sheet1", xml: worksheet(`<autoFilter ref="A1:D30"/>`) }, { name: "Sheet2" }], {}, NAMES);
    const modelled = book([{ name: "Sheet1" }, { name: "Sheet2" }], {}, NAMES);
    modelled.sheets[0] = { ...modelled.sheets[0]!, autoFilter: { range: { top: 0, left: 0, bottom: 29, right: 3 }, columns: [] } };
    for (const owned of [filtered, modelled]) {
      const ownedBefore = owned.opaque;
      adjust(owned, "insertRows", 0, 2, "s1");
      expect(textOf(owned.opaque, "xl/workbook.xml")).toContain(
        `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$D$30</definedName>`,
      );
      expect(spans(ownedBefore, owned, "xl/workbook.xml")).toEqual([
        ["Sheet1!$A$1:$D$30", "Sheet1!$A$3:$D$32"],
        ["$A$5", "$A$7"],
        ["SUM(Sheet1!$B$2:$B$30)", "SUM(Sheet1!$B$4:$B$32)"],
      ]);
    }
    // One without a localSheetId names no sheet to own it: it moves.
    const global = book([{ name: "Sheet1", xml: worksheet(`<autoFilter ref="A1:D30"/>`) }], {}, `<definedNames><definedName name="_xlnm._FilterDatabase" hidden="1">Sheet1!$A$1:$D$30</definedName></definedNames>`);
    const globalBefore = global.opaque;
    adjust(global, "insertRows", 0, 2, "s1");
    expect(spans(globalBefore, global, "xl/workbook.xml")).toEqual([["Sheet1!$A$1:$D$30", "Sheet1!$A$3:$D$32"]]);

    const other = book([{ name: "Sheet1" }, { name: "Sheet2" }], {}, NAMES);
    other.sheets = [other.sheets[1]!, other.sheets[0]!];
    const otherBefore = other.opaque;
    adjust(other, "insertRows", 0, 2, "s2");
    expect(spans(otherBefore, other, "xl/workbook.xml")).toEqual([
      ["$A$5", "$A$7"],
      ["Sheet2!$A$1:$A$9", "Sheet2!$A$3:$A$11"],
    ]);
  });

  it("a table part's table@ref, autoFilter@ref, sortState@ref and sortCondition@ref; an area gone leaves them", () => {
    const table =
      `${DECL}<table xmlns="${MAIN}" id="1" name="Table1" displayName="Table1" ref="A1:C10" totalsRowShown="0"><autoFilter ref="A1:C10">` +
      `<sortState ref="A2:C10"><sortCondition ref="B2:B10"/></sortState></autoFilter><tableColumns count="3"><tableColumn id="1" name="a"/>` +
      `<tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
    const spec: SheetSpec = { name: "Sheet1", xml: worksheet(`<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`), rels: [["rId1", "table", "../tables/table1.xml"]] };
    const model = book([spec], { "xl/tables/table1.xml": table });
    const before = model.opaque;
    expect(adjust(model, "insertRows", 4, 2)).toEqual(new Set(["s1"]));
    expect(untouched(before, model, SHEET1)).toBe(true);
    expect(spans(before, model, "xl/tables/table1.xml")).toEqual([
      ["A1:C10", "A1:C12"],
      ["A1:C10", "A1:C12"],
      ["A2:C10", "A2:C12"],
      ["B2:B10", "B2:B12"],
    ]);

    const gone = book([spec], { "xl/tables/table1.xml": table });
    const goneBefore = gone.opaque;
    expect(adjust(gone, "deleteCols", 0, 3)).toEqual(new Set());
    expect(gone.opaque).toBe(goneBefore);
  });

  it("every c:f of a chart, hosted by the sheet whose drawing relates it; a chart no drawing relates has no host", () => {
    const model = book(
      [{ name: "Sheet1" }, { name: "Sheet2", xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] }],
      {
        ...chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$B$1", "Sheet1!$A$2:$A$30", "Sheet1!$B$2:$B$30", "Sheet2!$C$2:$C$30"]),
        "xl/charts/chart2.xml": chartXml(["Sheet1!$A$1:$A$5", "$A$1:$A$5"]),
        "xl/charts/style1.xml": `${DECL}<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>`,
      },
    );
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s2"]));
    expect(spans(before, model, "xl/charts/chart1.xml")).toEqual([
      ["Sheet1!$B$1", "Sheet1!$B$2"],
      ["Sheet1!$A$2:$A$30", "Sheet1!$A$3:$A$31"],
      ["Sheet1!$B$2:$B$30", "Sheet1!$B$3:$B$31"],
    ]);
    expect(spans(before, model, "xl/charts/chart2.xml")).toEqual([["Sheet1!$A$1:$A$5", "Sheet1!$A$2:$A$6"]]);
    for (const name of [SHEET1, SHEET2, "xl/drawings/drawing1.xml", "xl/charts/style1.xml"]) {
      expect(untouched(before, model, name), name).toBe(true);
    }
  });

  it("a chartsheet (a foreign sheet) hosts its drawing's chart", () => {
    const chartsheet = `${DECL}<chartsheet xmlns="${MAIN}" xmlns:r="${R}"><sheetPr/><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="rId1"/></chartsheet>`;
    const model = book(
      [{ name: "Sheet1" }, { name: "Chart1", chartsheet: true, xml: chartsheet, rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] }],
      chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$B$2:$B$30"]),
    );
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s2"]));
    expect(spans(before, model, "xl/charts/chart1.xml")).toEqual([["Sheet1!$B$2:$B$30", "Sheet1!$B$3:$B$31"]]);
  });

  it("a pivot cache's worksheetSource@ref on the sheet its @sheet names; one naming a table, or another sheet, is left alone", () => {
    const parts = {
      "xl/pivotCache/pivotCacheDefinition1.xml": pivotCacheXml(`<worksheetSource ref="A1:C30" sheet="Sheet1"/>`),
      "xl/pivotCache/pivotCacheDefinition2.xml": pivotCacheXml(`<worksheetSource name="Table1"/>`),
      "xl/pivotCache/pivotCacheDefinition3.xml": pivotCacheXml(`<worksheetSource ref="A1:C30" sheet="Sheet2"/>`),
    };
    const model = book([{ name: "Sheet1" }, { name: "Sheet2" }], parts);
    const before = model.opaque;
    expect(adjust(model, "insertRows", 10, 5)).toEqual(new Set());
    expect(spans(before, model, "xl/pivotCache/pivotCacheDefinition1.xml")).toEqual([["A1:C30", "A1:C35"]]);
    expect(untouched(before, model, "xl/pivotCache/pivotCacheDefinition2.xml")).toBe(true);
    expect(untouched(before, model, "xl/pivotCache/pivotCacheDefinition3.xml")).toBe(true);
  });

  it("a pivot table's location@ref, on the sheet relating it", () => {
    const pivotTable =
      `${DECL}<pivotTableDefinition xmlns="${MAIN}" name="PivotTable1" cacheId="1" dataCaption="Values"><location ref="A3:B8" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>` +
      `<pivotFields count="1"><pivotField showAll="0"/></pivotFields></pivotTableDefinition>`;
    const specs: SheetSpec[] = [{ name: "Sheet1" }, { name: "Sheet2", rels: [["rId1", "pivotTable", "../pivotTables/pivotTable1.xml"]] }];
    const model = book(specs, { "xl/pivotTables/pivotTable1.xml": pivotTable });
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 2, "s2")).toEqual(new Set(["s2"]));
    expect(spans(before, model, "xl/pivotTables/pivotTable1.xml")).toEqual([["A3:B8", "A5:B10"]]);

    const elsewhere = book(specs, { "xl/pivotTables/pivotTable1.xml": pivotTable });
    const elsewhereBefore = elsewhere.opaque;
    expect(adjust(elsewhere, "insertRows", 0, 2, "s1")).toEqual(new Set());
    expect(elsewhere.opaque).toBe(elsewhereBefore);
  });
});

// --- drawings, comments, VML -------------------------------------------------------------------

describe("drawings, comments and VML", () => {
  const ANCHORS =
    `<xdr:twoCellAnchor editAs="oneCell">${corner("from", 1, 4)}${corner("to", 3, 9)}${SHAPE}</xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${corner("from", 0, 1)}${corner("to", 2, 3)}${SHAPE}</xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor editAs="absolute">${corner("from", 5, 5)}${corner("to", 7, 8)}${SHAPE}</xdr:twoCellAnchor>` +
    `<xdr:oneCellAnchor>${corner("from", 4, 6)}<xdr:ext cx="952500" cy="190500"/>${SHAPE}</xdr:oneCellAnchor>` +
    `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="190500"/>${SHAPE}</xdr:absoluteAnchor>`;

  function drawingBook(anchors: string): SpreadsheetModel {
    return book([{ name: "Sheet1", xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] }], {
      "xl/drawings/drawing1.xml": drawingXml(anchors),
    });
  }

  it("xdr:from/xdr:to xdr:row and xdr:col of two-cell anchors and a one-cell anchor's from; editAs=absolute and absolute anchors untouched", () => {
    const rows = drawingBook(ANCHORS);
    const rowsBefore = rows.opaque;
    expect(adjust(rows, "insertRows", 2, 3)).toEqual(new Set(["s1"]));
    expect(spans(rowsBefore, rows, "xl/drawings/drawing1.xml")).toEqual([
      ["4", "7"],
      ["9", "12"],
      ["3", "6"],
      ["6", "9"],
    ]);

    const cols = drawingBook(ANCHORS);
    const colsBefore = cols.opaque;
    adjust(cols, "insertCols", 0, 1);
    expect(spans(colsBefore, cols, "xl/drawings/drawing1.xml")).toEqual([
      ["1", "2"],
      ["3", "4"],
      ["0", "1"],
      ["2", "3"],
      ["4", "5"],
    ]);

    const absolute = drawingBook(`<xdr:twoCellAnchor editAs="absolute">${corner("from", 5, 5)}${corner("to", 7, 8)}${SHAPE}</xdr:twoCellAnchor>`);
    const absoluteBefore = absolute.opaque;
    expect(adjust(absolute, "insertRows", 0, 3)).toEqual(new Set());
    expect(absolute.opaque).toBe(absoluteBefore);
  });

  function notesBook(): SpreadsheetModel {
    return book([{ name: "Sheet1", xml: worksheet(`<legacyDrawing r:id="rId2"/>`), rels: NOTE_RELS }], {
      [COMMENTS]: commentsXml(comment("B3", "first") + comment("D10", "second")),
      [THREADED_PART]: threadedXml(threaded("B3", "T-1") + threaded("B3", "T-2", "T-1") + threaded("D10", "T-3")),
      [VML]: vmlXml(noteShape(1025, "2, 15, 1, 10, 4, 15, 5, 4", 2, 1) + noteShape(1026, "4, 15, 8, 10, 6, 15, 12, 4", 9, 3)),
    });
  }

  it("comment@ref, threadedComment@ref and the VML x:Row, x:Column and x:Anchor (numbers 1, 3, 5 and 7)", () => {
    const rows = notesBook();
    const rowsBefore = rows.opaque;
    expect(adjust(rows, "insertRows", 1, 2)).toEqual(new Set(["s1"]));
    expect(untouched(rowsBefore, rows, SHEET1)).toBe(true);
    expect(spans(rowsBefore, rows, COMMENTS)).toEqual([
      ["B3", "B5"],
      ["D10", "D12"],
    ]);
    expect(spans(rowsBefore, rows, THREADED_PART)).toEqual([
      ["B3", "B5"],
      ["B3", "B5"],
      ["D10", "D12"],
    ]);
    expect(spans(rowsBefore, rows, VML)).toEqual([
      ["\n    2, 15, 1, 10, 4, 15, 5, 4", "\n    2, 15, 3, 10, 4, 15, 7, 4"],
      ["2", "4"],
      ["\n    4, 15, 8, 10, 6, 15, 12, 4", "\n    4, 15, 10, 10, 6, 15, 14, 4"],
      ["9", "11"],
    ]);

    const cols = notesBook();
    const colsBefore = cols.opaque;
    adjust(cols, "insertCols", 2, 1);
    expect(spans(colsBefore, cols, COMMENTS)).toEqual([["D10", "E10"]]);
    expect(spans(colsBefore, cols, THREADED_PART)).toEqual([["D10", "E10"]]);
    expect(spans(colsBefore, cols, VML)).toEqual([
      ["\n    2, 15, 1, 10, 4, 15, 5, 4", "\n    3, 15, 1, 10, 5, 15, 5, 4"],
      ["\n    4, 15, 8, 10, 6, 15, 12, 4", "\n    5, 15, 8, 10, 7, 15, 12, 4"],
      ["3", "4"],
    ]);
  });
});

// --- removal ---------------------------------------------------------------------------------

describe("removal", () => {
  it("a conditional format whose range is gone is removed; one losing an entry keeps the rest", () => {
    const gone = `<conditionalFormatting sqref="A3:B4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>A3</formula></cfRule></conditionalFormatting>`;
    const kept = `<conditionalFormatting sqref="C3 C8"><cfRule type="expression" dxfId="1" priority="2"><formula>C8&gt;0</formula></cfRule></conditionalFormatting>`;
    const model = book([{ name: "Sheet1", xml: worksheet(gone + kept) }]);
    const before = model.opaque;
    adjust(model, "deleteRows", 2, 2);
    expect(spans(before, model, SHEET1, [gone])).toEqual([
      ["C3 C8", "C6"],
      ["C8&gt;0", "C6&gt;0"],
    ]);
  });

  it("dataValidation: emptied items removed, dataValidations@count kept right, the container removed with its last item", () => {
    const first = `<dataValidation type="list" sqref="A3"><formula1>$X$1:$X$3</formula1></dataValidation>`;
    const second = `<dataValidation type="whole" sqref="B4"><formula1>0</formula1></dataValidation>`;
    const third = `<dataValidation type="whole" sqref="B9"><formula1>1</formula1></dataValidation>`;
    const model = book([{ name: "Sheet1", xml: worksheet(`<dataValidations count="3">${first}${second}${third}</dataValidations>`) }]);
    const before = model.opaque;
    adjust(model, "deleteRows", 2, 2);
    expect(spans(before, model, SHEET1, [first, second])).toEqual([
      ["3", "1"],
      ["B9", "B7"],
    ]);

    const container = `<dataValidations count="2">${first}${second}</dataValidations>`;
    const emptied = book([{ name: "Sheet1", xml: worksheet(container) }]);
    const emptiedBefore = emptied.opaque;
    adjust(emptied, "deleteRows", 2, 2);
    expect(spans(emptiedBefore, emptied, SHEET1, [container])).toEqual([]);
  });

  it("hyperlink: removed with its cell, <hyperlinks> with its last", () => {
    const first = `<hyperlink ref="A3" r:id="rId1"/>`;
    const second = `<hyperlink ref="A9" location="'Sheet 2'!A1" display="x"/>`;
    const model = book([{ name: "Sheet1", xml: worksheet(`<hyperlinks>${first}${second}</hyperlinks>`) }]);
    const before = model.opaque;
    adjust(model, "deleteRows", 2, 2);
    expect(spans(before, model, SHEET1, [first])).toEqual([["A9", "A7"]]);

    const container = `<hyperlinks>${first}</hyperlinks>`;
    const emptied = book([{ name: "Sheet1", xml: worksheet(container) }]);
    const emptiedBefore = emptied.opaque;
    adjust(emptied, "deleteRows", 2, 1);
    expect(spans(emptiedBefore, emptied, SHEET1, [container])).toEqual([]);
  });

  it("protectedRange, cellWatch and ignoredError: removed when emptied, their containers with their last item", () => {
    const protectedRanges = `<protectedRanges><protectedRange sqref="A3" name="p1"/></protectedRanges>`;
    const watch = `<cellWatch r="B4"/>`;
    const ignored = `<ignoredError sqref="C3:C4" numberStoredAsText="1"/>`;
    const model = book([
      {
        name: "Sheet1",
        xml: worksheet(`${protectedRanges}<cellWatches>${watch}<cellWatch r="B9"/></cellWatches><ignoredErrors>${ignored}<ignoredError sqref="C3:C10" formula="1"/></ignoredErrors>`),
      },
    ]);
    const before = model.opaque;
    adjust(model, "deleteRows", 2, 2);
    expect(spans(before, model, SHEET1, [protectedRanges, watch, ignored])).toEqual([
      ["B9", "B7"],
      ["C3:C10", "C3:C8"],
    ]);
  });

  it("x14:conditionalFormatting and x14:dataValidation (its count kept right) removed; x14:sparkline removed, an emptied sparklineGroup with it", () => {
    const goneCf = x14Cf(1, "A3:A4");
    const goneDv = x14Dv("B3", "Sheet2!$A$1:$A$3");
    const goneGroup = sparklineGroup(sparkline("Sheet1!B3:E3", "F3"));
    const goneSparkline = sparkline("Sheet1!B4:E4", "F4");
    const model = book([
      {
        name: "Sheet1",
        xml: worksheet(
          `<extLst>${ext(X14_CF, `<x14:conditionalFormattings>${goneCf}${x14Cf(2, "A9")}</x14:conditionalFormattings>`)}` +
            `${ext(X14_DV, `<x14:dataValidations count="2" xmlns:xm="${XM}">${goneDv}${x14Dv("B10", "Sheet2!$A$1:$A$3")}</x14:dataValidations>`)}` +
            `${ext(X14_SPARKLINES, `<x14:sparklineGroups xmlns:xm="${XM}">${goneGroup}${sparklineGroup(goneSparkline + sparkline("Sheet1!B9:E9", "F9"))}</x14:sparklineGroups>`)}</extLst>`,
        ),
      },
      { name: "Sheet2" },
    ]);
    const before = model.opaque;
    adjust(model, "deleteRows", 2, 2);
    expect(spans(before, model, SHEET1, [goneCf, goneDv, goneGroup, goneSparkline])).toEqual([
      ["A9", "A7"],
      ["2", "1"],
      ["B10", "B8"],
      ["Sheet1!B9:E9", "Sheet1!B7:E7"],
      ["F9", "F7"],
    ]);
  });

  it("a comment whose cell is deleted goes, with its VML shape (matched by x:Row/x:Column); a threaded comment with its replies", () => {
    const shape = noteShape(1025, "2, 15, 1, 10, 4, 15, 5, 4", 2, 1);
    const note = comment("B3", "first");
    const thread = threaded("B3", "T-1");
    const reply = threaded("B3", "T-2", "T-1");
    const model = book([{ name: "Sheet1", xml: worksheet(`<legacyDrawing r:id="rId2"/>`), rels: NOTE_RELS }], {
      [COMMENTS]: commentsXml(note + comment("D10", "second")),
      [THREADED_PART]: threadedXml(thread + reply + threaded("D10", "T-3")),
      [VML]: vmlXml(shape + noteShape(1026, "4, 15, 8, 10, 6, 15, 12, 4", 9, 3)),
    });
    const before = model.opaque;
    expect(adjust(model, "deleteRows", 2, 1)).toEqual(new Set(["s1"]));
    expect(spans(before, model, COMMENTS, [note])).toEqual([["D10", "D9"]]);
    expect(spans(before, model, THREADED_PART, [thread, reply])).toEqual([["D10", "D9"]]);
    expect(spans(before, model, VML, [shape])).toEqual([
      ["\n    4, 15, 8, 10, 6, 15, 12, 4", "\n    4, 15, 7, 10, 6, 15, 11, 4"],
      ["9", "8"],
    ]);
  });
});

// --- sheet names ----------------------------------------------------------------------------

describe("sheet names in retained text", () => {
  it("asks the mapper in model terms: range and index with the model name, formula with source-time text and the host id; a null index leaves the text", () => {
    const calls: string[] = [];
    const mapper: RetainedReferenceMapper = {
      range: (sheetName, range) => {
        calls.push(`range ${sheetName} ${formatRange(range)}`);
        return range;
      },
      index: (sheetName, axis, index) => {
        calls.push(`index ${sheetName} ${axis} ${index}`);
        return null;
      },
      formula: (canonical, hostSheetId) => {
        calls.push(`formula ${hostSheetId || '""'} ${canonical}`);
        return canonical;
      },
    };
    const model = book(
      [
        {
          name: "Data",
          sourceName: "Sheet1",
          xml: worksheet(
            `<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1"><formula>$A1</formula></cfRule></conditionalFormatting>` +
              `<rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks>`,
          ),
        },
        { name: "Sheet2", xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] },
      ],
      {
        ...chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$B$2:$B$30"]),
        "xl/pivotCache/pivotCacheDefinition1.xml": pivotCacheXml(`<worksheetSource ref="A1:C30" sheet="Sheet1"/>`),
      },
      `<definedNames><definedName name="Start">Sheet1!$A$1</definedName></definedNames>`,
    );
    const before = model.opaque;
    expect(rewriteXlsxRetainedReferences(model, mapper)).toEqual(new Set());
    expect(model.opaque).toBe(before);
    expect(calls).toEqual([
      "range Data A1:A10",
      "formula s1 $A1",
      "index Data row 4",
      "formula s2 Sheet1!$B$2:$B$30",
      'formula "" Sheet1!$A$1',
      "range Data A1:C30",
    ]);
  });

  it("a renamed sheet: what names its source-time name still moves, the name written back exactly as found", () => {
    const model = book(
      [
        { name: "Data", sourceName: "Sheet1", xml: worksheet(`<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1"><formula>$A1</formula></cfRule></conditionalFormatting>`) },
        { name: "Sheet2", xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] },
        { name: "Totals", sourceName: "My Data" },
      ],
      {
        ...chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$B$2:$B$30", "'My Data'!$A$1:$A$5"]),
        "xl/pivotCache/pivotCacheDefinition1.xml": pivotCacheXml(`<worksheetSource ref="A1:C30" sheet="Sheet1"/>`),
      },
      `<definedNames><definedName name="Start">sheet1!$A$1</definedName></definedNames>`,
    );
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s1", "s2"]));
    expect(spans(before, model, SHEET1)).toEqual([
      ["A1:A10", "A2:A11"],
      ["$A1", "$A2"],
    ]);
    expect(spans(before, model, "xl/charts/chart1.xml")).toEqual([["Sheet1!$B$2:$B$30", "Sheet1!$B$3:$B$31"]]);
    expect(spans(before, model, "xl/pivotCache/pivotCacheDefinition1.xml")).toEqual([["A1:C30", "A2:C31"]]);
    expect(spans(before, model, "xl/workbook.xml")).toEqual([["sheet1!$A$1", "sheet1!$A$2"]]);

    const quoted = book(
      [{ name: "Sheet1" }, { name: "Sheet2", xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] }, { name: "Totals", sourceName: "My Data" }],
      chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$B$2:$B$30", "'My Data'!$A$1:$A$5"]),
    );
    const quotedBefore = quoted.opaque;
    adjust(quoted, "insertRows", 0, 1, "s3");
    expect(spans(quotedBefore, quoted, "xl/charts/chart1.xml")).toEqual([["'My Data'!$A$1:$A$5", "'My Data'!$A$2:$A$6"]]);
  });

  it("a reference to another sheet is left alone, on either side", () => {
    const own = `<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1"><formula>$A1&gt;Sheet2!B5</formula></cfRule></conditionalFormatting>`;
    const theirs = `<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1"><formula>$A1&gt;Sheet2!B5</formula></cfRule></conditionalFormatting>`;
    const model = book([
      { name: "Sheet1", xml: worksheet(own) },
      { name: "Sheet2", xml: worksheet(theirs) },
    ]);
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s1"]));
    expect(spans(before, model, SHEET1)).toEqual([
      ["A1:A10", "A2:A11"],
      ["$A1&gt;Sheet2!B5", "$A2&gt;Sheet2!B5"],
    ]);
    expect(untouched(before, model, SHEET2)).toBe(true);
  });

  it("a duplicated sheet's own parts read its origin's name as the copy", () => {
    const drawingSheet = (id: string): SheetSpec => ({ name: id, xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] });
    const build = (): SpreadsheetModel => {
      const model = book([drawingSheet("Sheet1")], {
        ...chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$A$1:$A$5"]),
        // The copy's closure, at next-free names (SPEC-121's duplicate).
        [SHEET2]: worksheet(`<drawing r:id="rId1"/>`),
        [relsPathFor(SHEET2)]: relsXml([["rId1", "drawing", "../drawings/drawing2.xml"]]),
        ...chartParts("xl/drawings/drawing2.xml", "xl/charts/chart2.xml", ["Sheet1!$A$1:$A$5"]),
      });
      model.sheets.push(sheetOf("copy", "Sheet1 (2)", SHEET2, "grid", { copiedFrom: "Sheet1", dirty: true }));
      return model;
    };
    const copy = build();
    const copyBefore = copy.opaque;
    expect(adjust(copy, "insertRows", 0, 1, "copy")).toEqual(new Set(["copy"]));
    expect(untouched(copyBefore, copy, "xl/charts/chart1.xml")).toBe(true);
    expect(spans(copyBefore, copy, "xl/charts/chart2.xml")).toEqual([["Sheet1!$A$1:$A$5", "Sheet1!$A$2:$A$6"]]);

    const original = build();
    const originalBefore = original.opaque;
    expect(adjust(original, "insertRows", 0, 1, "s1")).toEqual(new Set(["s1"]));
    expect(spans(originalBefore, original, "xl/charts/chart1.xml")).toEqual([["Sheet1!$A$1:$A$5", "Sheet1!$A$2:$A$6"]]);
    expect(untouched(originalBefore, original, "xl/charts/chart2.xml")).toBe(true);
  });
});

// --- copy-on-write and the result ----------------------------------------------------------------

describe("copy-on-write and the changed sheets", () => {
  const CF = `<conditionalFormatting sqref="A1:A10"><cfRule type="expression" priority="1"><formula>Sheet2!$A$1</formula></cfRule></conditionalFormatting>`;

  it("when nothing moves, opaque keeps its identity", () => {
    const model = book([{ name: "Sheet1", xml: worksheet(CF) }, { name: "Sheet2" }]);
    const before = model.opaque;
    expect(adjust(model, "insertRows", 100, 5)).toEqual(new Set());
    expect(adjust(model, "deleteCols", 3, 2)).toEqual(new Set());
    expect(model.opaque).toBe(before);
  });

  it("a change makes a new opaque array with new parts only where they changed; the old array and parts are untouched", () => {
    const model = book([{ name: "Sheet1", xml: worksheet(CF) }, { name: "Sheet2" }]);
    const before = model.opaque;
    const snapshot = before.map((part) => ({ ...part }));
    adjust(model, "insertRows", 0, 1);
    expect(model.opaque).not.toBe(before);
    expect(before).toEqual(snapshot);
    expect(model.opaque.map((part) => part.name)).toEqual(before.map((part) => part.name));
    model.opaque.forEach((part, i) => {
      if (part.name === SHEET1) {
        expect(part).not.toBe(before[i]);
        expect(part.data).toBe(toData(textOf(before, SHEET1).replace(`sqref="A1:A10"`, `sqref="A2:A11"`)));
      } else {
        expect(part, part.name).toBe(before[i]);
      }
    });
  });

  it("returns the sheets whose own parts changed: the edited sheet's, another's chart naming it — not a sheet only naming others", () => {
    const model = book(
      [
        { name: "Sheet1", xml: worksheet(CF) },
        { name: "Sheet2", xml: worksheet(`<drawing r:id="rId1"/>`), rels: [["rId1", "drawing", "../drawings/drawing1.xml"]] },
        { name: "Sheet3", xml: worksheet(CF) },
      ],
      chartParts("xl/drawings/drawing1.xml", "xl/charts/chart1.xml", ["Sheet1!$B$2:$B$30"]),
    );
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s1", "s2"]));
    expect(untouched(before, model, SHEET3)).toBe(true);
  });
});

// --- worksheet parts -----------------------------------------------------------------------------

describe("worksheet parts", () => {
  it("never decodes or scans the cells: a large part's tail is spliced onto its untouched bytes, whatever the byte alignment, its BOM kept", () => {
    const rows = Array.from({ length: 3000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>Zeile — Grüße ${i}</t></is></c><c r="B${i + 1}"><f>A${i + 1}</f></c></row>`);
    const links = Array.from({ length: 2500 }, (_, i) => `<hyperlink ref="C${i + 1}" display="Verknüpfung ${i}"/>`).join("");
    for (const pad of ["", "ä", "€"]) {
      // Malformed cells: a whole-part scan would throw and leave the part alone.
      const cells = `<sheetData><row r="1"><c r="Z1"><v>${pad}</c></row>${rows.join("")}</sheetData>`;
      const xml = `﻿${worksheet(`<hyperlinks>${links}</hyperlinks>`, cells)}`;
      const model = book([{ name: "Sheet1", xml }]);
      expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s1"]));
      const expected = xml.replace(/<hyperlink ref="C(\d+)"/g, (_, n: string) => `<hyperlink ref="C${Number(n) + 1}"`);
      expect(textOf(model.opaque, SHEET1) === expected, `pad ${JSON.stringify(pad)}`).toBe(true);
      expect(partOf(model, SHEET1)!.data === toData(expected)).toBe(true);
    }
  });

  it("a part with no literal <sheetData (a prefixed namespace) is adjusted whole", () => {
    const xml =
      `${DECL}<x:worksheet xmlns:x="${MAIN}"><x:sheetData><x:row r="2"><x:c r="A2"><x:v>1</x:v></x:c></x:row></x:sheetData>` +
      `<x:conditionalFormatting sqref="A1:A4"><x:cfRule type="expression" priority="1"><x:formula>A1</x:formula></x:cfRule></x:conditionalFormatting></x:worksheet>`;
    const model = book([{ name: "Sheet1", xml }]);
    const before = model.opaque;
    adjust(model, "insertRows", 0, 1);
    expect(spans(before, model, SHEET1)).toEqual([
      ["A1:A4", "A2:A5"],
      ["A1", "A2"],
    ]);
  });

  it("a part that is not well-formed is left as it is; the rest still moves", () => {
    const model = book([{ name: "Sheet1", xml: worksheet(`<legacyDrawing r:id="rId2"/>`), rels: NOTE_RELS }], {
      [COMMENTS]: commentsXml(comment("B3", "first")),
      [THREADED_PART]: threadedXml(threaded("B3", "T-1")),
      [VML]: `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape><x:ClientData><x:Row>2</x:Row></v:shape></xml>`,
    });
    const before = model.opaque;
    expect(adjust(model, "insertRows", 0, 1)).toEqual(new Set(["s1"]));
    expect(untouched(before, model, VML)).toBe(true);
    expect(spans(before, model, COMMENTS)).toEqual([["B3", "B4"]]);
  });
});
