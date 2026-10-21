// SPEC-132 golden wire forms for the ODS chart objects this product CREATES
// or regenerates (ods/chartXml.ts over core's odfChartContentXml) —
// hand-authored against ODF 1.2 before the serialisers were written, never
// screen-scraped from their output. An imported object nobody edited is never
// written from these: its parts pass through, or are patched in place (path
// 2), byte for byte where nothing changed.
//
// Settled here, not guessed in code:
// - the object is `content.xml` ONLY. ODF 1.2 Part 1 §3.1.3 makes styles
//   and meta sub-files optional (an `office:document-content` alone is a
//   complete chart document), LibreOffice's chart import skips a missing
//   `styles.xml` / `meta.xml` stream, and the text and presentation products
//   already create their chart objects this way (odf/chartWritePlan.ts,
//   presentation/odp/writer.ts) — so no styles.xml, meta.xml or
//   `ObjectReplacements/` picture is generated (LibreOffice paints the
//   replacement itself);
// - content.xml is core's SPEC-052/092 chart document (no automatic styles:
//   the colours ride `chart:color`) with the sheet references added (ODF 1.2
//   §19 attributes, §11 `chart:` elements): `chart:plot-area@table:cell-range-
//   address` (§19.599) when the chart has a plot-area range; the x axis's
//   `chart:categories` (§11.10) after its title and before its grids (the
//   schema's `chart:axis` content order); per series `chart:values-cell-range-
//   address` (§19.21) and `chart:label-cell-address` (§19.12), and a
//   `chart:domain` child (§11.14) holding a scatter series' x values on EVERY
//   series (LibreOffice writes it once and lets later series inherit — both
//   read the same). Addresses are ods/rangeAddress.ts's form: `$Sheet.$A$1:.$B$5`,
//   the sheet on the start only, quoted when a formula would quote it;
// - a pie or doughnut (no axes otherwise) gets an x axis carrying just its
//   categories, as LibreOffice writes one;
// - a series whose name is literal (no name reference), or a reference with
//   no ODS form (`#REF!`), writes the EMPTY address `""` — ODF's "no range"
//   (a cellRangeAddressList may be empty), which LibreOffice itself writes
//   for an empty series. Never a `local-table` address beside sheet ones:
//   LibreOffice switches a chart with an unresolvable range to internal data,
//   and this reader treats any `local-table` address as a literal chart. The
//   literal name still travels in the `local-table` header;
// - the `local-table` data cache (§9.1.1 `table:table`, ODF 1.2 §11.3.4's
//   local table) is still written — what a consumer without the sheet shows;
// - a chart that plots visible cells only (`plotVisibleOnly`, every chart
//   insert from selection makes) says so the way LibreOffice does: one
//   automatic chart style `ch1` (§16.2, `style:family="chart"`) whose
//   `style:chart-properties` carry `chart:include-hidden-cells="false"`
//   (§20.21 — its default is true, so a chart without it plots hidden cells
//   too), named by the plot area's `chart:style-name` (its first attribute),
//   the root declaring the style namespace before `office:version`; a chart
//   that includes hidden cells writes no style (the default says so);
// - the frame holds `<draw:object>` (§10.4.6) with `draw:notify-on-update-
//   of-ranges` (§19.130) first, then `xlink:href="./Object {n}"` and the
//   simple/embed/onLoad triple as LibreOffice writes it; the list is the
//   plot-area range, else every distinct series reference (name, categories,
//   values, series by series), space-separated; no attribute without one;
// - META-INF/manifest.xml gains the object directory first — `Object {n}/`
//   with `manifest:version="1.2"` (Part 3 §4.16.11: a sub-document's entry
//   carries its version) and the chart media type — then its content.xml.

import { chartAccentRef, NEW_CHART_HEIGHT_CM, NEW_CHART_WIDTH_CM, type ChartVisual } from "../../core/charts";
import type { SheetChartRefs } from "../model";

/** The column chart every reference golden is built from: Data!A1:C4 —
 * months in A2:A4, "North" (B) and "South" (C) with a blank C3, at
 * NEW_CHART_WIDTH_CM × NEW_CHART_HEIGHT_CM, coloured as an inserted chart is
 * (chartAccentRef: accent1 #4472c4, accent2 #ed7d31). */
export const GOLDEN_ODS_COLUMN_VISUAL: ChartVisual = {
  chartType: "columnClustered",
  categories: ["Jan", "Feb", "Mar"],
  series: [
    { name: "North", values: [10, 20, 30], color: chartAccentRef(0) },
    { name: "South", values: [15, null, 35], color: chartAccentRef(1) },
  ],
  options: { legend: "right", dataLabels: false, gridlines: true },
};

/** Its references: the plot area, the categories (shared) and two series. */
export const GOLDEN_ODS_COLUMN_REFS: SheetChartRefs = {
  series: [
    { name: "Data!$B$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" },
    { name: "Data!$C$1", categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" },
  ],
  plotArea: "Data!$A$1:$C$4",
  plotVisibleOnly: true,
};

export const GOLDEN_ODS_CHART_SIZE_CM = { width: NEW_CHART_WIDTH_CM, height: NEW_CHART_HEIGHT_CM } as const;

const GOLDEN_ROOT =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.2"><office:body><office:chart>`;

const GOLDEN_END = `</chart:chart></office:chart></office:body></office:document-content>`;

/** The root, automatic styles and body opening of a chart that plots
 * visible cells only: the style namespace declared, `ch1` excluding hidden
 * cells. */
const GOLDEN_VISIBLE_ONLY_ROOT =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.2">` +
  `<office:automatic-styles><style:style style:name="ch1" style:family="chart"><style:chart-properties chart:include-hidden-cells="false"/></style:style></office:automatic-styles>` +
  `<office:body><office:chart>`;

/** A string cell and a float cell of the `local-table`. */
const str = (text: string): string => `<table:table-cell office:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
const num = (value: number): string => `<table:table-cell office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`;

/** The column chart's `local-table`: the corner, the two names, one row per
 * month; the blank C3 an empty cell (a gap, not zero). */
export const GOLDEN_ODS_COLUMN_LOCAL_TABLE =
  `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
  `<table:table-columns><table:table-column/><table:table-column/></table:table-columns>` +
  `<table:table-header-rows><table:table-row><table:table-cell/>${str("North")}${str("South")}</table:table-row></table:table-header-rows>` +
  `<table:table-rows>` +
  `<table:table-row>${str("Jan")}${num(10)}${num(15)}</table:table-row>` +
  `<table:table-row>${str("Feb")}${num(20)}<table:table-cell/></table:table-row>` +
  `<table:table-row>${str("Mar")}${num(30)}${num(35)}</table:table-row>` +
  `</table:table-rows></table:table>`;

/** A new (or regenerated) reference column chart's `Object {n}/content.xml`:
 * visible cells only (`ch1` on the plot area), the plot-area range, the x
 * axis's categories, each series' values and label on the sheet, then the
 * cache. */
export const GOLDEN_ODS_COLUMN_CONTENT =
  GOLDEN_VISIBLE_ONLY_ROOT +
  `<chart:chart chart:class="chart:bar" chart:vertical="true" chart:stacked="false" svg:width="12cm" svg:height="7.5cm">` +
  `<chart:legend chart:legend-position="end"/>` +
  `<chart:plot-area chart:style-name="ch1" table:cell-range-address="$Data.$A$1:.$C$4">` +
  `<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="$Data.$A$2:.$A$4"/></chart:axis>` +
  `<chart:axis chart:dimension="y" chart:name="primary-y"><chart:grid chart:class="major"/></chart:axis>` +
  `<chart:series chart:class="chart:bar" chart:color="#4472c4" chart:values-cell-range-address="$Data.$B$2:.$B$4" chart:label-cell-address="$Data.$B$1"/>` +
  `<chart:series chart:class="chart:bar" chart:color="#ed7d31" chart:values-cell-range-address="$Data.$C$2:.$C$4" chart:label-cell-address="$Data.$C$1"/>` +
  `</chart:plot-area>` +
  GOLDEN_ODS_COLUMN_LOCAL_TABLE +
  GOLDEN_END;

/** The column chart's frame object at `Object 1`, notified of its plot area. */
export const GOLDEN_ODS_COLUMN_OBJECT_ELEMENT =
  `<draw:object draw:notify-on-update-of-ranges="$Data.$A$1:.$C$4" xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;

/** A chart with no references to notify of (a literal chart). */
export const GOLDEN_ODS_OBJECT_ELEMENT_WITHOUT_RANGES = `<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;

/** The scatter golden: sheet "Q1 Data" (quoted on the wire), x values in
 * A2:A4, "Y1" (B) and "Y2" (C, blank C3) — no plot-area range, title "Fit",
 * no legend, no gridlines. */
export const GOLDEN_ODS_SCATTER_VISUAL: ChartVisual = {
  chartType: "scatter",
  categories: ["0.5", "1", "2"],
  series: [
    { name: "Y1", values: [1, 4, 9], color: chartAccentRef(0) },
    { name: "Y2", values: [2, null, 5], color: chartAccentRef(1) },
  ],
  options: { legend: "none", dataLabels: false, gridlines: false, title: "Fit" },
};

export const GOLDEN_ODS_SCATTER_REFS: SheetChartRefs = {
  series: [
    { name: "'Q1 Data'!$B$1", categories: "'Q1 Data'!$A$2:$A$4", values: "'Q1 Data'!$B$2:$B$4" },
    { name: "'Q1 Data'!$C$1", categories: "'Q1 Data'!$A$2:$A$4", values: "'Q1 Data'!$C$2:$C$4" },
  ],
  plotVisibleOnly: true,
};

/** The scatter's content.xml: visible cells only, no categories on the x
 * axis — every series carries its x values as a `chart:domain`. */
export const GOLDEN_ODS_SCATTER_CONTENT =
  GOLDEN_VISIBLE_ONLY_ROOT +
  `<chart:chart chart:class="chart:scatter" svg:width="12cm" svg:height="7.5cm">` +
  `<chart:title><text:p>Fit</text:p></chart:title>` +
  `<chart:plot-area chart:style-name="ch1">` +
  `<chart:axis chart:dimension="x" chart:name="primary-x"></chart:axis>` +
  `<chart:axis chart:dimension="y" chart:name="primary-y"></chart:axis>` +
  `<chart:series chart:class="chart:scatter" chart:color="#4472c4" chart:values-cell-range-address="$'Q1 Data'.$B$2:.$B$4" chart:label-cell-address="$'Q1 Data'.$B$1">` +
  `<chart:domain table:cell-range-address="$'Q1 Data'.$A$2:.$A$4"/></chart:series>` +
  `<chart:series chart:class="chart:scatter" chart:color="#ed7d31" chart:values-cell-range-address="$'Q1 Data'.$C$2:.$C$4" chart:label-cell-address="$'Q1 Data'.$C$1">` +
  `<chart:domain table:cell-range-address="$'Q1 Data'.$A$2:.$A$4"/></chart:series>` +
  `</chart:plot-area>` +
  `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
  `<table:table-columns><table:table-column/><table:table-column/></table:table-columns>` +
  `<table:table-header-rows><table:table-row><table:table-cell/>${str("Y1")}${str("Y2")}</table:table-row></table:table-header-rows>` +
  `<table:table-rows>` +
  `<table:table-row>${str("0.5")}${num(1)}${num(2)}</table:table-row>` +
  `<table:table-row>${str("1")}${num(4)}<table:table-cell/></table:table-row>` +
  `<table:table-row>${str("2")}${num(9)}${num(5)}</table:table-row>` +
  `</table:table-rows></table:table>` +
  GOLDEN_END;

/** No plot area: the frame is notified of every distinct series reference —
 * Y1's name, the shared x values, Y1's values, Y2's name, Y2's values. */
export const GOLDEN_ODS_SCATTER_NOTIFY_RANGES = "$'Q1 Data'.$B$1 $'Q1 Data'.$A$2:.$A$4 $'Q1 Data'.$B$2:.$B$4 $'Q1 Data'.$C$1 $'Q1 Data'.$C$2:.$C$4";

/** The pie golden: one series "Share" with a LITERAL name (no name
 * reference) over Data!B2:B4, categories Data!A2:A4, legend at the bottom. */
export const GOLDEN_ODS_PIE_VISUAL: ChartVisual = {
  chartType: "pie",
  categories: ["Jan", "Feb", "Mar"],
  series: [{ name: "Share", values: [10, 20, 30], color: chartAccentRef(0) }],
  options: { legend: "bottom", dataLabels: true, gridlines: false },
};

export const GOLDEN_ODS_PIE_REFS: SheetChartRefs = {
  series: [{ categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" }],
  plotVisibleOnly: false,
};

/** The pie's content.xml (hidden cells included: no style): an x axis
 * carrying just the categories, the literal name's empty label address, its
 * text in the `local-table` header. */
export const GOLDEN_ODS_PIE_CONTENT =
  GOLDEN_ROOT +
  `<chart:chart chart:class="chart:circle" svg:width="12cm" svg:height="7.5cm">` +
  `<chart:legend chart:legend-position="bottom"/>` +
  `<chart:plot-area chart:data-labels="true">` +
  `<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="$Data.$A$2:.$A$4"/></chart:axis>` +
  `<chart:series chart:class="chart:circle" chart:color="#4472c4" chart:values-cell-range-address="$Data.$B$2:.$B$4" chart:label-cell-address=""/>` +
  `</chart:plot-area>` +
  `<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>` +
  `<table:table-columns><table:table-column/></table:table-columns>` +
  `<table:table-header-rows><table:table-row><table:table-cell/>${str("Share")}</table:table-row></table:table-header-rows>` +
  `<table:table-rows>` +
  `<table:table-row>${str("Jan")}${num(10)}</table:table-row>` +
  `<table:table-row>${str("Feb")}${num(20)}</table:table-row>` +
  `<table:table-row>${str("Mar")}${num(30)}</table:table-row>` +
  `</table:table-rows></table:table>` +
  GOLDEN_END;

/** The parts a created object at `Object 1` adds to the package. */
export const GOLDEN_ODS_CHART_OBJECT_PART_NAMES = ["Object 1/content.xml"] as const;

/** Its META-INF/manifest.xml entries: the directory (a sub-document, its
 * version and the chart media type), then its content.xml. */
export const GOLDEN_ODS_CHART_MANIFEST_ENTRIES =
  `<manifest:file-entry manifest:full-path="Object 1/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.chart"/>` +
  `<manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/>`;
