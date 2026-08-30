// SPEC-092 AC-29: hand-authored canonical `.odp` chart bytes, validated
// against ODF 1.2 §10.4.6 (`draw:frame`/`draw:object`) and §16 (chart). The
// serialiser's output must equal this and the reader must accept it (plus
// Impress-flavoured attribute-reordered variants) — the golden-first process,
// since whether Impress needs `Object N/styles.xml`/`meta.xml` or an
// `ObjectReplacements` preview is not certain and must not be invented.
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
  sourcePart: "Object 1/content.xml",
};

/** The canonical slide-flavoured `draw:frame` + `draw:object` for
 * `GOLDEN_CHART_FRAME` — `presentation:class="chart"`, geometry in
 * `svg:x/y/width/height`, holding the `draw:object` reference alone. */
export const CHART_DRAW_FRAME_GOLDEN =
  `<draw:frame presentation:class="chart"` +
  ` svg:x="2.54cm" svg:y="1.27cm" svg:width="12cm" svg:height="7.5cm">` +
  `<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `</draw:frame>`;
