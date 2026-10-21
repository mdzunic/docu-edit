// xlsx sheet charts on the wire (SPEC-132 "The chart model", "Read-only
// charts", "Writing", "New chart parts"): THIS FILE is the normative source
// for how an `xl/charts/chartN.xml` part becomes a SheetChartContent's visual
// and references, how a part is patched in place (path 2: `c:f` texts and
// `c:strCache`/`c:numCache` only, every other byte kept), and the canonical
// wire forms of a chart created here (the part through core chartPartXml,
// the anchor's `xdr:graphicFrame`) — hand-authored goldens in chartGolden.ts.
//
// SPEC-132 SKELETON: the signatures and contracts below are final; the bodies
// are filled in by the xlsx chart wire change.
import type { ChartPartRefs } from "../../core/chartPartOoxml";
import type { ChartVisual } from "../../core/charts";
import type { ChartCell } from "../charts";
import { unsupportedChartVisual } from "../charts";
import type { SheetChartRefs } from "../model";

/** What reading one chart part gives. */
export interface XlsxChartRead {
  visual: ChartVisual;
  /** Absent: the part names no reference (a literal chart), or it is out of
   * the catalogue (unmodelled — `visual.unsupportedChartXml` set). */
  refs?: SheetChartRefs;
  readOnlyReason?: string;
  /** Some `c:f` has no cache beside it: the reader refreshes the chart for
   * display at open (refreshSheetChart forDisplay). */
  uncached?: true;
}

/** Reads a chart part: core parseChartXml → visual (null → the part text as
 * `unsupportedChartXml`, no refs); readChartPartRefs → refs;
 * chartReadOnlyReason. */
export function readXlsxChart(partXml: string): XlsxChartRead {
  return { visual: unsupportedChartVisual(partXml) };
}

/** Each `c:ser` of the part's one chart-type element, in order: `c:tx/
 * c:strRef/c:f` → name, `c:cat` or `c:xVal` (`c:strRef`/`c:numRef`/
 * `c:multiLvlStrRef`)`/c:f` → categories, `c:val` or `c:yVal` `c:numRef/c:f`
 * → values, each through fromXlsxFormula (a literal `c:tx/c:v` leaves name
 * absent); `c:plotVisOnly@val` (absent = true). Undefined when no series
 * names a reference. `uncached`: some `c:f` without a cache sibling. */
export function readChartPartRefs(partXml: string): { refs: SheetChartRefs; uncached: boolean } | undefined {
  void partXml;
  return undefined;
}

/** "Read-only charts": PIVOT_CHART_REASON for a `c:pivotSource`,
 * MULTI_LEVEL_CHART_REASON for a `c:multiLvlStrRef`, else undefined. */
export function chartReadOnlyReason(partXml: string): string | undefined {
  void partXml;
  return undefined;
}

/** Path 2 ("Writing"): the part patched in place with scanXml offsets —
 * series i is the i-th `c:ser` of the part's one type element; each of its
 * `c:f` texts becomes toXlsxFormula(ref); with `cellsOf`, each of its
 * `c:strCache`/`c:numCache` is rewritten from the cells: `c:ptCount` = the
 * cell count, a `c:pt` per non-blank cell (numCache: raw numbers only;
 * strCache: the formatted text), `c:formatCode` and `c:extLst` kept. A
 * reference `cellsOf` cannot resolve (null) keeps its cache. Without
 * `cellsOf` (a read-only chart's refsDirty) only the `c:f` texts change.
 * Every other byte is the source's; the identical string when nothing
 * changes. */
export function patchXlsxChartPart(partXml: string, refs: SheetChartRefs, cellsOf?: (canonical: string) => ChartCell[] | null): string {
  void refs;
  void cellsOf;
  return partXml;
}

/** The wire refs core chartPartXml writes for a regenerated or new part:
 * each canonical reference through toXlsxFormula. */
export function xlsxChartWireRefs(refs: SheetChartRefs): ChartPartRefs {
  return { series: refs.series.map(() => ({})) };
}

/** A regenerated or new chart part (path 1 / New chart parts):
 * chartPartXml(visual, xlsxChartWireRefs(refs)). */
export function xlsxChartPartXml(visual: ChartVisual, refs: SheetChartRefs | undefined): string {
  void visual;
  void refs;
  return "";
}

/** The canonical `xdr:graphicFrame` a created chart's anchor holds ("New
 * chart parts" — the golden's scaffold and Excel's zero `xdr:xfrm`). */
export function chartGraphicFrameXml(cNvPrId: number, name: string, relId: string): string {
  void cNvPrId;
  void name;
  void relId;
  return "";
}

/** `xl/charts/chart{n}.xml`, n = 1 + the highest across `taken`. */
export function nextChartPartName(taken: Iterable<string>): string {
  let max = 0;
  for (const name of taken) {
    const match = /^xl\/charts\/chart(\d+)\.xml$/.exec(name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `xl/charts/chart${max + 1}.xml`;
}
