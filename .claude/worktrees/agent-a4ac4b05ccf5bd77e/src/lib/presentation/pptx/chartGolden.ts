// SPEC-092 AC-29: hand-authored canonical `.pptx` chart bytes, validated
// against ECMA-376 §19.3.1.21 (`p:graphicFrame`) and §21.2 (`c:chartSpace`).
// The serialiser's output must equal these and the reader must accept them
// (plus PowerPoint-flavoured attribute-reordered variants) — the golden-first
// process SPEC-086 established, since the exact `p:nvGraphicFramePr` scaffold
// PowerPoint requires is not certain and must not be invented.
import type { ChartFrame } from "../model";

/** The clean model a golden round trip reproduces (id aside). */
export const GOLDEN_CHART_FRAME: ChartFrame = {
  id: "chart-golden",
  xCm: 2.54,
  yCm: 1.27,
  widthCm: 12,
  heightCm: 7.5,
  kind: "chart",
  chartType: "columnClustered",
  categories: ["A", "B"],
  series: [{ name: "S", values: [1, 2], color: "#4e79a7" }],
  options: { legend: "none", dataLabels: false, gridlines: false },
  sourcePart: "ppt/charts/chart1.xml",
};

/** The canonical `p:graphicFrame` container for `GOLDEN_CHART_FRAME` at
 * `p:cNvPr id="2"`, referencing the chart part under `rId2`. */
export const CHART_GRAPHIC_FRAME_GOLDEN =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart 2"/>` +
  `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
  `<p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="914400" y="457200"/><a:ext cx="4320000" cy="2700000"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>` +
  `</a:graphicData></a:graphic></p:graphicFrame>`;

/** The canonical `c:chartSpace` part body for `GOLDEN_CHART_FRAME` — a single
 * clustered-column series with no title, legend, labels or gridlines. No
 * `c:externalData` (this editor has no workbook writer, D5). */
export const CHART_PART_GOLDEN =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
  ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
  `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
  `<c:ser><c:idx val="0"/><c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:spPr><a:solidFill><a:srgbClr val="4E79A7"/></a:solidFill></c:spPr>` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>` +
  `<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/>` +
  `<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser>` +
  `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
  `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>` +
  `<c:gapWidth val="150"/><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>` +
  `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>` +
  `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>` +
  `</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
