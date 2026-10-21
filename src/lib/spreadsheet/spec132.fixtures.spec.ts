// SPEC-132's chart fixture (spec132-charts.xlsx, built by fixtures/build.ts):
// package-level facts only — the builder is deterministic, the committed e2e
// file is its bytes, the package is whole, it holds exactly the parts
// SPEC132_CHARTS lists, every XML part parses, and each chart part carries
// the references, anchors and producer shapes the readers and writers are
// proven on. Nothing here asserts how the app reads it.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { parseXml } from "../core/xml";
import { unzip } from "../core/zip";
import { buildSpec132ChartsXlsxBytes, SPEC132_CHARTS } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const committed = (name: string): Uint8Array => new Uint8Array(readFileSync(join(E2E_FIXTURES, name)));
const texts = (xml: string, tag: string): string[] => [...xml.matchAll(new RegExp(`<${tag}>([^<]*)</${tag}>`, "g"))].map((match) => match[1]!);
/** An `xdr:from`/`xdr:to` corner with zero offsets, as Excel writes it. */
const corner = (tag: string, [col, , row]: readonly number[]): string =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

describe("spec132-charts.xlsx", () => {
  const X = SPEC132_CHARTS.xlsx;
  const C = X.charts;

  it("is deterministic, the committed bytes, whole, with exactly the listed parts", async () => {
    const built = await buildSpec132ChartsXlsxBytes();
    expect(equalBytes(built, await buildSpec132ChartsXlsxBytes())).toBe(true);
    const bytes = committed(X.file);
    expect(equalBytes(bytes, built)).toBe(true);
    expect(await packageIntegrity(bytes)).toEqual([]);
    expect([...(await unzip(bytes)).keys()]).toEqual([...X.parts]);
  });

  it("every XML part parses", async () => {
    const parts = await unzip(committed(X.file));
    for (const [name, data] of parts) {
      expect(() => parseXml(decoder.decode(data)), name).not.toThrow();
    }
  });

  it("each chart part names its series' references, with caches", async () => {
    const parts = await unzip(committed(X.file));
    for (const chart of [C.column, C.pie, C.column3D, C.scatter, C.line, C.chartSheet]) {
      const xml = decoder.decode(parts.get(chart.part));
      expect(xml, chart.part).toContain(`<c:${chart.typeElement}>`);
      expect(texts(xml, "c:f"), chart.part).toEqual(chart.series.flatMap((series) => [series.name, series.categories, series.values]));
      expect(xml.match(/<c:f>/g)!.length, chart.part).toBe((xml.match(/<c:(strCache|numCache)>/g) ?? []).length);
      expect(xml).toContain(`<c:plotVisOnly val="1"/>`);
    }
    const column = decoder.decode(parts.get(C.column.part));
    expect(column).toContain(`<c:barDir val="col"/><c:grouping val="clustered"/>`);
    expect(column).toContain(`xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"`);
    expect(column).toContain(`<c:legendPos val="r"/>`);
    expect(column).toContain(`<mc:Choice Requires="c14"`);
    expect(column.match(/<c16:uniqueId /g)).toHaveLength(3);
    expect(column).toContain(`<c:pt idx="1"><c:v>30</c:v></c:pt>`); // D3's cached value
    const rels = decoder.decode(parts.get(C.column.rels));
    expect(rels).toContain(`Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"`);
    expect(rels).toContain(`Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"`);
    expect(decoder.decode(parts.get(C.column3D.part))).toMatch(/<c:view3D>.*<c:floor>.*<c:sideWall>.*<c:backWall>.*<c:bar3DChart>.*<c:shape val="box"\/>/);
    expect(decoder.decode(parts.get(C.scatter.part))).toContain(`<c:scatterStyle val="lineMarker"/>`);
    expect(decoder.decode(parts.get(C.line.part))).toMatch(/<c:grouping val="standard"\/>.*<c:marker val="1"\/>/);
    expect(decoder.decode(parts.get(C.pie.part))).toContain(`<c:pieChart><c:varyColors val="1"/>`);
    const chartEx = decoder.decode(parts.get(C.waterfall.part));
    expect(texts(chartEx, "cx:f")).toEqual([...C.waterfall.formulas]);
    expect(chartEx).toContain(`<cx:series layoutId="waterfall"`);
  });

  it("anchors each frame where SPEC132_CHARTS says, the waterfall wrapped cx1 with Excel's Fallback", async () => {
    const parts = await unzip(committed(X.file));
    for (const chart of [C.column, C.pie, C.column3D, C.waterfall, C.scatter, C.line]) {
      const drawing = decoder.decode(parts.get(chart.drawingPart));
      expect(drawing, chart.name).toContain(`<xdr:twoCellAnchor>${corner("from", chart.from)}${corner("to", chart.to)}<xdr:graphicFrame macro="">`);
      expect(drawing, chart.name).toContain(`<xdr:cNvPr id="${chart.cNvPrId}" name="${chart.name}">`);
    }
    const data = decoder.decode(parts.get(X.worksheets.Data.drawingPart));
    expect(data).toContain(`<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"><xdr:twoCellAnchor>`);
    expect(data).toContain(`<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">`);
    expect(data).toContain(`<mc:Fallback xmlns=""><xdr:twoCellAnchor>`);
    expect(data).toContain(C.waterfall.fallbackText);
    const drawingRels = decoder.decode(parts.get(X.worksheets.Data.drawingRels));
    expect(drawingRels).toContain(`Id="rId4" Type="${C.waterfall.relationshipType}" Target="../charts/chartEx1.xml"`);
    const types = decoder.decode(parts.get("[Content_Types].xml"));
    expect(types).toContain(`<Override PartName="/${C.waterfall.part}" ContentType="${C.waterfall.contentType}"/>`);
    expect(types).toContain(`<Override PartName="/${X.chartSheet.part}" ContentType="${X.chartSheet.contentType}"/>`);
  });

  it("holds the chart sheet: workbook entry, chartsheet → drawing4's absoluteAnchor → chart6", async () => {
    const parts = await unzip(committed(X.file));
    const workbook = decoder.decode(parts.get("xl/workbook.xml"));
    expect(workbook).toContain(`<sheet name="${X.chartSheet.name}" sheetId="4" r:id="rId4"/>`);
    expect(workbook).not.toContain("activeTab");
    expect(decoder.decode(parts.get("xl/_rels/workbook.xml.rels"))).toContain(`Id="rId4" Type="${X.chartSheet.relationshipType}" Target="chartsheets/sheet1.xml"`);
    expect(decoder.decode(parts.get(X.chartSheet.part))).toContain(`<sheetView zoomScale="118" workbookViewId="0" zoomToFit="1"/>`);
    const drawing = decoder.decode(parts.get(X.chartSheet.drawingPart));
    expect(drawing).toContain(`<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="8666667" cy="6293304"/>`);
    expect(decoder.decode(parts.get(X.chartSheet.drawingRels))).toContain(`Target="../charts/chart6.xml"`);
    for (const sheet of Object.values(X.worksheets)) {
      expect(decoder.decode(parts.get(sheet.part))).toMatch(/<pageMargins[^>]*\/><drawing r:id="rId1"\/><\/worksheet>$/);
    }
  });
});
