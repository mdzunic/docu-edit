// The golden wire forms of a chart created or regenerated on a sheet
// (SPEC-132 "New chart parts", "Writing" path 1): hand-authored against
// ECMA-376 Part 1 §21.2 (DrawingML — Charts) and §20.5.2.16 (SpreadsheetML
// Drawing `graphicFrame`) and written BEFORE the serialisers — core
// chartPartXml's `refs` mode and xlsx/chartXml.ts's chartGraphicFrameXml —
// which chartXml.spec.ts holds to these strings byte for byte. THIS FILE is
// the normative source for every detail the spec left to the goldens:
// - a regenerated part is SPEC-052's canonical `c:chartSpace` (the same
//   scaffold the pptx golden pins: `c:autoTitleDeleted` or a rich title, one
//   type element, its fixed axis ids, `c:catAx`/`c:valAx`) with the series
//   reading the sheet: each `c:f` is the reference's wire text
//   (toXlsxFormula, XML-escaped) — never SPEC-052's placeholder
//   `Sheet1!$A$2…` formulas — and each cache is the visual's;
// - a series name with a reference is `c:tx/c:strRef` (`c:f` + a one-point
//   `c:strCache`); without one it is the literal `c:tx/c:v` (CT_SerTx's
//   other choice);
// - categories with a reference are `c:cat/c:strRef` with a `c:strCache`
//   (a scatter's x values `c:xVal/c:numRef` with a `General` `c:numCache`);
//   WITHOUT one the series has no `c:cat`/`c:xVal` at all — both are
//   optional (CT_BarSer, CT_LineSer, CT_ScatterSer …) and Excel numbers the
//   categories 1, 2, … itself, exactly what SPEC-132's refresh computes for
//   a chart with no categories reference;
// - values with a reference are `c:val/c:numRef` (scatter `c:yVal`) with a
//   `General` `c:numCache`; without one a schema-valid literal
//   `c:val/c:numLit` (pinned by chartPartOoxml.spec.ts);
// - a gap (null) has no `c:pt`, its index still counted by `c:ptCount`;
// - `c:plotVisOnly` carries the chart's plotVisibleOnly and is followed by
//   `<c:dispBlanksAs val="gap"/>` (CT_DispBlanksAs): SPEC-132 plots blank
//   cells as gaps, the schema's default for the element is "zero", and
//   LibreOffice applies that default to any part not written by Office
//   2007 — so the part says "gap" in so many words;
// - no `c:externalData`, `c:style`, `mc:AlternateContent` or chartStyle/
//   colors siblings (SPEC-132: none is generated);
// - the anchor's `xdr:graphicFrame` is Excel's own scaffold: `macro=""`, an
//   empty `xdr:cNvGraphicFramePr`, the zero `xdr:xfrm` Excel writes for a
//   chart in a cell anchor (the anchor alone places it), and a `c:chart`
//   declaring `c` and `r` itself, so the frame is valid in any drawing part
//   whatever its root declares; a cNvPr id is k + 1 for "Chart {k}"
//   (drawingGolden.ts's rule).
import type { ChartVisual } from "../../core/charts";
import type { SheetChartRefs } from "../model";

const CHART_SPACE_OPEN =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
  ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;

/** `dLbls` as every non-scatter canonical chart writes it (no labels). */
const NO_DATA_LABELS =
  `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
  `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;

/** `catAx` + `valAx` with major gridlines (defaultChartOptions). */
const CATEGORY_AND_VALUE_AXES =
  `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>` +
  `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111111111"/></c:valAx>`;

// --- (a) a clustered column over Data!$A$1:$C$4 -------------------------------------------

/** The input of GOLDEN_COLUMN_CHART_PART: a clustered column as Insert ▸
 * Chart builds it from Data!A1:C4 (a name row, a category column, two
 * value columns) — accent colours, default options, a gap in South. */
export const GOLDEN_COLUMN_CHART_VISUAL: ChartVisual = {
  chartType: "columnClustered",
  categories: ["Jan", "Feb", "Mar"],
  series: [
    { name: "North", values: [10, 20, 30], color: { theme: "accent1" } },
    { name: "South", values: [4, null, 6], color: { theme: "accent2" } },
  ],
  options: { legend: "right", dataLabels: false, gridlines: true },
};

/** GOLDEN_COLUMN_CHART_VISUAL's references (canonical). */
export const GOLDEN_COLUMN_CHART_REFS: SheetChartRefs = {
  series: [
    { name: "Data!$B$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" },
    { name: "Data!$C$1", categories: "Data!$A$2:$A$4", values: "Data!$C$2:$C$4" },
  ],
  plotVisibleOnly: true,
};

/** `chartSpace` > `chart` > `plotArea` > `barChart` (`barDir` col,
 * `grouping` clustered) holding two `ser`, each in CT_BarSer order — `idx`,
 * `order`, `tx/strRef` (`f` + one-point `strCache`), `spPr` (a scheme
 * colour), `cat/strRef` (`f` + `strCache`), `val/numRef` (`f` + `General`
 * `numCache`; South's gap has no `pt`) — then `dLbls`, `gapWidth`, the axis
 * ids, the two axes, the right-hand legend, `plotVisOnly` 1 and
 * `dispBlanksAs` gap. */
export const GOLDEN_COLUMN_CHART_PART =
  CHART_SPACE_OPEN +
  `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
  `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
  `<c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>` +
  `<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser>` +
  `<c:ser><c:idx val="1"/><c:order val="1"/>` +
  `<c:tx><c:strRef><c:f>Data!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>South</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr>` +
  `<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Data!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser>` +
  NO_DATA_LABELS +
  `<c:gapWidth val="150"/><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>` +
  CATEGORY_AND_VALUE_AXES +
  `</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>` +
  `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;

// --- (b) a scatter over Data!$A$1:$B$4 --------------------------------------------------

/** The input of GOLDEN_SCATTER_CHART_PART: a one-series scatter (x in
 * column A, y in column B) titled after its series, plotting hidden cells
 * too (plotVisibleOnly false). */
export const GOLDEN_SCATTER_CHART_VISUAL: ChartVisual = {
  chartType: "scatter",
  categories: ["1", "2", "4"],
  series: [{ name: "Growth", values: [1.5, 3, 7.25], color: { theme: "accent1" } }],
  options: { title: "Growth", legend: "right", dataLabels: false, gridlines: true },
};

/** GOLDEN_SCATTER_CHART_VISUAL's references (canonical). */
export const GOLDEN_SCATTER_CHART_REFS: SheetChartRefs = {
  series: [{ name: "Data!$B$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" }],
  plotVisibleOnly: false,
};

/** `scatterChart` (`scatterStyle` lineMarker) holding one CT_ScatterSer —
 * `tx/strRef`, `spPr`, `xVal/numRef` and `yVal/numRef`, each with a `General` `numCache`, then
 * `smooth` 0 — under a rich title; two `valAx` (x at the bottom, y on the
 * left with gridlines); `plotVisOnly` 0. */
export const GOLDEN_SCATTER_CHART_PART =
  CHART_SPACE_OPEN +
  `<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Growth</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
  `<c:plotArea><c:layout/>` +
  `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
  `<c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Growth</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>` +
  `<c:xVal><c:numRef><c:f>Data!$A$2:$A$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:xVal>` +
  `<c:yVal><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>7.25</c:v></c:pt></c:numCache></c:numRef></c:yVal>` +
  `<c:smooth val="0"/></c:ser>` +
  `<c:axId val="111111111"/><c:axId val="222222222"/></c:scatterChart>` +
  `<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:valAx>` +
  `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111111111"/></c:valAx>` +
  `</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>` +
  `<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;

// --- (c) a line with no categories reference and a literal name ------------------------------

/** The input of GOLDEN_LINE_CHART_PART: two lines over a sheet whose name
 * needs quoting and escaping ('R&D'), with no category column — so the
 * categories are Excel's numbering — the first series named by a cell,
 * the second by a literal name. */
export const GOLDEN_LINE_CHART_VISUAL: ChartVisual = {
  chartType: "line",
  categories: ["1", "2", "3"],
  series: [
    { name: "Actual", values: [12, 15, 11], color: { theme: "accent1" } },
    { name: "Target", values: [14, 14, 14], color: { theme: "accent2" } },
  ],
  options: { legend: "bottom", dataLabels: false, gridlines: true },
};

/** GOLDEN_LINE_CHART_VISUAL's references (canonical): no categories, and
 * no name for the second series. */
export const GOLDEN_LINE_CHART_REFS: SheetChartRefs = {
  series: [{ name: "'R&D'!$B$1", values: "'R&D'!$B$2:$B$4" }, { values: "'R&D'!$C$2:$C$4" }],
  plotVisibleOnly: true,
};

/** `lineChart` holding two CT_LineSer — `tx` (the first a
 * `strRef` whose `c:f` escapes the `&` of 'R&D', the second the literal
 * `c:tx/c:v`), `spPr`, `marker` none, NO `cat`, `val/numRef` — then `dLbls`,
 * `marker` 0, the axes; the legend at the bottom. */
export const GOLDEN_LINE_CHART_PART =
  CHART_SPACE_OPEN +
  `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
  `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
  `<c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>'R&amp;D'!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Actual</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>` +
  `<c:marker><c:symbol val="none"/></c:marker>` +
  `<c:val><c:numRef><c:f>'R&amp;D'!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt><c:pt idx="2"><c:v>11</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser>` +
  `<c:ser><c:idx val="1"/><c:order val="1"/>` +
  `<c:tx><c:v>Target</c:v></c:tx>` +
  `<c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr>` +
  `<c:marker><c:symbol val="none"/></c:marker>` +
  `<c:val><c:numRef><c:f>'R&amp;D'!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` +
  `<c:pt idx="0"><c:v>14</c:v></c:pt><c:pt idx="1"><c:v>14</c:v></c:pt><c:pt idx="2"><c:v>14</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser>` +
  NO_DATA_LABELS +
  `<c:marker val="0"/><c:axId val="111111111"/><c:axId val="222222222"/></c:lineChart>` +
  CATEGORY_AND_VALUE_AXES +
  `</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>` +
  `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;

// --- (d) the anchor's graphic frame -----------------------------------------------------------

/** The input of GOLDEN_CHART_GRAPHIC_FRAME: chartGraphicFrameXml(3, "Chart
 * 2", "rId2") — the second chart on a sheet, its part related as rId2. */
export const GOLDEN_CHART_GRAPHIC_FRAME_INPUT = { cNvPrId: 3, name: "Chart 2", relId: "rId2" } as const;

/** §20.5.2.16 `graphicFrame` (`macro` empty): `nvGraphicFramePr` (`cNvPr`
 * + an empty `cNvGraphicFramePr`), `xfrm` zero, `a:graphic` >
 * `a:graphicData` with the chart URI holding the `c:chart r:id` reference
 * (CT_RelId), declaring its own `c` and `r`. What
 * the twoCell anchor holds before its `xdr:clientData`. */
export const GOLDEN_CHART_GRAPHIC_FRAME =
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 2"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
  `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>` +
  `</a:graphicData></a:graphic></xdr:graphicFrame>`;

/** ECMA-376 Part 2 §9.3: the drawing `.rels` entry relating a new chart
 * part (CHART_RELATIONSHIP; the target relative to `xl/drawings/`). */
export const GOLDEN_CHART_DRAWING_RELATIONSHIP = `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>`;

/** The `[Content_Types].xml` Override a new chart part takes (CHART_CONTENT_TYPE). */
export const GOLDEN_CHART_OVERRIDE = `<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
