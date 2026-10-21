// SPEC-132 ods chart wire: golden equality of every created part and element,
// reference reading on LibreOffice-shaped objects (spec132-charts.ods's
// among them; a scatter's shared domain; an object saved without its cache),
// patches touching only the addresses and the local-table, and the frame's
// notify list.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { buildSpec132ChartsOdsBytes, SPEC132_CHARTS_ODS } from "../fixtures/build";
import type { SheetChartRefs } from "../model";
import {
  GOLDEN_ODS_CHART_MANIFEST_ENTRIES,
  GOLDEN_ODS_CHART_OBJECT_PART_NAMES,
  GOLDEN_ODS_CHART_SIZE_CM,
  GOLDEN_ODS_COLUMN_CONTENT,
  GOLDEN_ODS_COLUMN_OBJECT_ELEMENT,
  GOLDEN_ODS_COLUMN_REFS,
  GOLDEN_ODS_COLUMN_VISUAL,
  GOLDEN_ODS_OBJECT_ELEMENT_WITHOUT_RANGES,
  GOLDEN_ODS_PIE_CONTENT,
  GOLDEN_ODS_PIE_REFS,
  GOLDEN_ODS_PIE_VISUAL,
  GOLDEN_ODS_SCATTER_CONTENT,
  GOLDEN_ODS_SCATTER_NOTIFY_RANGES,
  GOLDEN_ODS_SCATTER_REFS,
  GOLDEN_ODS_SCATTER_VISUAL,
} from "./chartGolden";
import {
  chartObjectElementXml,
  newOdsChartObjectParts,
  notifyRangesOf,
  odsChartContentXml,
  odsChartManifestEntries,
  patchOdfChartContent,
  readOdfChartRefs,
  readOdsChart,
  withNotifyRanges,
} from "./chartXml";

const { width, height } = GOLDEN_ODS_CHART_SIZE_CM;

describe("goldens", () => {
  it("a new object's content.xml, its parts, its draw:object and its manifest entries equal the hand-authored goldens", () => {
    expect(odsChartContentXml(GOLDEN_ODS_COLUMN_VISUAL, width, height, undefined, GOLDEN_ODS_COLUMN_REFS)).toBe(GOLDEN_ODS_COLUMN_CONTENT);
    expect(odsChartContentXml(GOLDEN_ODS_SCATTER_VISUAL, width, height, undefined, GOLDEN_ODS_SCATTER_REFS)).toBe(GOLDEN_ODS_SCATTER_CONTENT);
    expect(odsChartContentXml(GOLDEN_ODS_PIE_VISUAL, width, height, undefined, GOLDEN_ODS_PIE_REFS)).toBe(GOLDEN_ODS_PIE_CONTENT);
    expect(chartObjectElementXml("Object 1", notifyRangesOf(GOLDEN_ODS_COLUMN_REFS))).toBe(GOLDEN_ODS_COLUMN_OBJECT_ELEMENT);
    expect(chartObjectElementXml("Object 1", undefined)).toBe(GOLDEN_ODS_OBJECT_ELEMENT_WITHOUT_RANGES);
    expect(notifyRangesOf(GOLDEN_ODS_SCATTER_REFS)).toBe(GOLDEN_ODS_SCATTER_NOTIFY_RANGES);
    expect(newOdsChartObjectParts("Object 1", GOLDEN_ODS_COLUMN_CONTENT).map((part) => part.name)).toEqual([...GOLDEN_ODS_CHART_OBJECT_PART_NAMES]);
    expect(odsChartManifestEntries("Object 1")).toBe(GOLDEN_ODS_CHART_MANIFEST_ENTRIES);
  });

  it("a created object reads back its references (plotVisibleOnly included) and visual; patching it with the same refs and visual is the identical string", () => {
    for (const [content, refs] of [
      [GOLDEN_ODS_COLUMN_CONTENT, GOLDEN_ODS_COLUMN_REFS],
      [GOLDEN_ODS_SCATTER_CONTENT, GOLDEN_ODS_SCATTER_REFS],
      [GOLDEN_ODS_PIE_CONTENT, GOLDEN_ODS_PIE_REFS],
    ] as const) {
      const read = readOdsChart(content);
      expect(read.uncached).toBeUndefined();
      expect(read.refs).toEqual(refs);
      expect(patchOdfChartContent(content, refs, read.visual)).toBe(content);
    }
  });
});

/** A LibreOffice-shaped object over one bar series. */
function lo(plotArea: string, categories: string, values: string, label: string, hidden = `<style:chart-properties chart:include-hidden-cells="false"/>`): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
    `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
    `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
    `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:automatic-styles>` +
    `<style:style style:name="ch2" style:family="chart">${hidden}</style:style></office:automatic-styles><office:body><office:chart>` +
    `<chart:chart svg:width="8cm" svg:height="5cm" chart:class="chart:bar"><chart:plot-area chart:style-name="ch2" table:cell-range-address="${plotArea}" chart:data-source-has-labels="both">` +
    `<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="${categories}"/></chart:axis>` +
    `<chart:axis chart:dimension="y" chart:name="primary-y"><chart:grid chart:class="major"/></chart:axis>` +
    `<chart:series chart:values-cell-range-address="${values}" chart:label-cell-address="${label}" chart:class="chart:bar"><chart:data-point chart:repeated="2"/></chart:series>` +
    `<chart:wall/></chart:plot-area><table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
    `<table:table-columns><table:table-column/></table:table-columns><table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell>` +
    `<table:table-cell office:value-type="string"><text:p>Amount</text:p></table:table-cell></table:table-row></table:table-header-rows><table:table-rows>` +
    `<table:table-row><table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>` +
    `<table:table-row><table:table-cell office:value-type="string"><text:p>B</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row>` +
    `</table:table-rows></table:table></chart:chart></office:chart></office:body></office:document-content>`
  );
}

describe("LibreOffice-shaped objects", () => {
  it("reads sheet addresses (absolute markers as written), quoted names and include-hidden-cells", () => {
    expect(readOdfChartRefs(lo("Data.A1:Data.B3", "Data.A2:Data.A3", "Data.B2:Data.B3", "Data.B1"))).toEqual({
      series: [{ name: "Data!B1", categories: "Data!A2:A3", values: "Data!B2:B3" }],
      plotArea: "Data!A1:B3",
      plotVisibleOnly: true,
    });
    expect(readOdfChartRefs(lo("$'Q1 Data'.$A$1:.$B$3", "$'Q1 Data'.$A$2:.$A$3", "$'Q1 Data'.$B$2:.$B$3", "$'Q1 Data'.$B$1", ""))).toEqual({
      series: [{ name: "'Q1 Data'!$B$1", categories: "'Q1 Data'!$A$2:$A$3", values: "'Q1 Data'!$B$2:$B$3" }],
      plotArea: "'Q1 Data'!$A$1:$B$3",
      plotVisibleOnly: false,
    });
  });

  it("an address into the local-table means no refs; a chart:surface chart is unsupported", () => {
    const literal = lo("local-table.$A$1:.$B$3", "local-table.$A$2:.$A$3", "local-table.$B$2:.$B$3", "local-table.$B$1");
    expect(readOdfChartRefs(literal)).toBeUndefined();
    expect(readOdsChart(literal).refs).toBeUndefined();
    const surface = lo("Data.A1:Data.B3", "Data.A2:Data.A3", "Data.B2:Data.B3", "Data.B1").replace('chart:class="chart:bar">', 'chart:class="chart:surface">');
    expect(readOdsChart(surface)).toEqual({ visual: expect.objectContaining({ unsupportedChartXml: surface }) });
  });

  it("patching rewrites only the addresses that changed, and replaces the local-table", () => {
    const source = lo("Data.A1:Data.B3", "Data.A2:Data.A3", "Data.B2:Data.B3", "Data.B1");
    const refs = readOdfChartRefs(source)!;
    expect(patchOdfChartContent(source, refs)).toBe(source);
    const moved = { ...refs, plotArea: "Data!A2:B4", series: [{ name: "Data!B2", categories: "Data!A3:A4", values: "Data!B3:B4" }] };
    expect(patchOdfChartContent(source, moved)).toBe(lo("$Data.A2:.B4", "$Data.A3:.A4", "$Data.B3:.B4", "$Data.B2"));
    const read = readOdsChart(source);
    const visual = { ...read.visual, series: [{ ...read.visual.series[0]!, values: [5, 6] }] };
    const patched = patchOdfChartContent(source, refs, visual);
    expect(readOdsChart(patched).visual.series[0]!.values).toEqual([5, 6]);
    expect(patched.slice(0, patched.indexOf("<table:table "))).toBe(source.slice(0, source.indexOf("<table:table ")));
  });
});

/** A LibreOffice-shaped scatter: x values as the FIRST series' `chart:domain` only (the second inherits it). */
function loScatter(domain: string, cache = true): string {
  const table =
    `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
    `<table:table-columns><table:table-column table:number-columns-repeated="3"/></table:table-columns><table:table-header-rows><table:table-row>` +
    `<table:table-cell><text:p/></table:table-cell><table:table-cell office:value-type="string"><text:p>X</text:p></table:table-cell>` +
    `<table:table-cell office:value-type="string"><text:p>Y1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Y2</text:p></table:table-cell>` +
    `</table:table-row></table:table-header-rows><table:table-rows>` +
    [
      [1, 2, 3],
      [2, 4, 6],
    ]
      .map(
        (row, i) =>
          `<table:table-row><table:table-cell office:value-type="string"><text:p>${i + 1}</text:p></table:table-cell>` +
          row.map((value) => `<table:table-cell office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`).join("") +
          `</table:table-row>`,
      )
      .join("") +
    `</table:table-rows></table:table>`;
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
    `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
    `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
    `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:body><office:chart>` +
    `<chart:chart svg:width="8cm" svg:height="5cm" chart:class="chart:scatter"><chart:plot-area table:cell-range-address="Data.A1:Data.C3" chart:data-source-has-labels="both">` +
    `<chart:axis chart:dimension="x" chart:name="primary-x"/><chart:axis chart:dimension="y" chart:name="primary-y"/>` +
    `<chart:series chart:values-cell-range-address="Data.B2:Data.B3" chart:label-cell-address="Data.B1" chart:class="chart:scatter">` +
    `<chart:domain table:cell-range-address="${domain}"/><chart:data-point chart:repeated="2"/></chart:series>` +
    `<chart:series chart:values-cell-range-address="Data.C2:Data.C3" chart:label-cell-address="Data.C1" chart:class="chart:scatter"><chart:data-point chart:repeated="2"/></chart:series>` +
    `</chart:plot-area>${cache ? table : ""}</chart:chart></office:chart></office:body></office:document-content>`
  );
}

describe("a scatter's shared domain", () => {
  it("every series reads the first series' chart:domain as its categories (LibreOffice writes it once)", () => {
    expect(readOdfChartRefs(loScatter("Data.A2:Data.A3"))).toEqual({
      series: [
        { name: "Data!B1", categories: "Data!A2:A3", values: "Data!B2:B3" },
        { name: "Data!C1", categories: "Data!A2:A3", values: "Data!C2:C3" },
      ],
      plotArea: "Data!A1:C3",
      plotVisibleOnly: false,
    });
  });

  it("moved x values rewrite the one domain and add none to the series inheriting it; a series given its own x values gains a domain", () => {
    const source = loScatter("Data.A2:Data.A3");
    const refs = readOdfChartRefs(source)!;
    const moved = { ...refs, series: refs.series.map((each) => ({ ...each, categories: "Data!A3:A4" })) };
    expect(patchOdfChartContent(source, moved)).toBe(loScatter("$Data.A3:.A4"));
    const own = { ...refs, series: [refs.series[0]!, { ...refs.series[1]!, categories: "Data!D2:D3" }] };
    const patched = patchOdfChartContent(source, own);
    expect(readOdfChartRefs(patched)!.series.map((each) => each.categories)).toEqual(["Data!A2:A3", "Data!D2:D3"]);
    expect(patched).toContain(`chart:class="chart:scatter"><chart:domain table:cell-range-address="$Data.D2:.D3"/><chart:data-point chart:repeated="2"/></chart:series>`);
  });
});

describe("an object saved without its cache", () => {
  it("reads its references and an empty visual, uncached — refreshed from the cells for display at open", () => {
    const read = readOdsChart(lo("Data.A1:Data.B3", "Data.A2:Data.A3", "Data.B2:Data.B3", "Data.B1").replace(/<table:table table:name="local-table">.*<\/table:table>/, ""));
    expect(read.uncached).toBe(true);
    expect(read.refs).toEqual({ series: [{ name: "Data!B1", categories: "Data!A2:A3", values: "Data!B2:B3" }], plotArea: "Data!A1:B3", plotVisibleOnly: true });
    expect(read.visual.chartType).toBe("columnClustered");
    expect(read.visual.unsupportedChartXml).toBeUndefined();
    expect(read.visual.series).toHaveLength(1);
    // A cached object is not.
    expect(readOdsChart(lo("Data.A1:Data.B3", "Data.A2:Data.A3", "Data.B2:Data.B3", "Data.B1")).uncached).toBeUndefined();
  });
});

describe("spec132-charts.ods's objects", () => {
  it("Object 1 reads as the cached clustered column over Data with its sheet references; Object 2 (chart:surface) as unsupported with none", async () => {
    const parts = await unzip(await buildSpec132ChartsOdsBytes());
    const { bar, surface } = SPEC132_CHARTS_ODS.charts;
    const read = readOdsChart(new TextDecoder().decode(parts.get(bar.part)));
    expect(read.uncached).toBeUndefined();
    expect(read.refs).toEqual(bar.refs);
    expect(read.visual).toMatchObject({ chartType: "columnClustered", categories: ["Q1", "Q2", "Q3", "Q4"], options: { legend: "right", title: bar.title, gridlines: true } });
    expect(read.visual.series.map((series) => [series.name, series.values])).toEqual([
      ["North", [10, 20, 30, 40]],
      ["South", [15, 25, 35, 45]],
      ["West", [12, 30, 18, 24]],
    ]);
    const surfaceXml = new TextDecoder().decode(parts.get(surface.part));
    expect(readOdsChart(surfaceXml)).toEqual({ visual: expect.objectContaining({ unsupportedChartXml: surfaceXml }) });
  });
});

/** A chart frame notified of `notify`. */
const frame = (notify: string): string =>
  `<draw:frame draw:name="Object 1"><draw:object draw:notify-on-update-of-ranges="${notify}" xlink:href="./Object 1" xlink:type="simple"><loext:p/></draw:object></draw:frame>`;

describe("the frame's notify list", () => {
  it("keeps the producer's bytes when the list names the same ranges, rewrites it when they differ, and drops it for none", () => {
    const calc = frame("Data.A1:Data.B5");
    expect(withNotifyRanges(calc, "$Data.A1:.B5")).toBe(calc);
    expect(withNotifyRanges(calc, "$Data.A2:.B6")).toBe(frame("$Data.A2:.B6"));
    expect(withNotifyRanges(calc, undefined)).toBe(`<draw:frame draw:name="Object 1"><draw:object xlink:href="./Object 1" xlink:type="simple"><loext:p/></draw:object></draw:frame>`);
    const image = `<draw:frame draw:name="Image 1"><draw:image xlink:href="Pictures/a.png"/></draw:frame>`;
    expect(withNotifyRanges(image, "$Data.A1:.B5")).toBe(image);
  });

  it("names the plot area's range when there is one, else every distinct series reference", () => {
    const { plotArea, series, plotVisibleOnly } = SPEC132_CHARTS_ODS.charts.bar.refs;
    const withoutPlotArea: SheetChartRefs = { series: series.map((each) => ({ ...each })), plotVisibleOnly };
    expect(notifyRangesOf({ ...withoutPlotArea, plotArea })).toBe("$Data.A1:.D5");
    expect(notifyRangesOf(withoutPlotArea)).toBe("$Data.B1 $Data.A2:.A5 $Data.B2:.B5 $Data.C1 $Data.C2:.C5 $Data.D1 $Data.D2:.D5");
    expect(notifyRangesOf(undefined)).toBeUndefined();
  });
});
