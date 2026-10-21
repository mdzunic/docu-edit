// SPEC-132 ods chart wire: golden equality of every created part and element,
// reference reading on a LibreOffice-shaped object, and the no-op patch.
import { describe, expect, it } from "vitest";
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
