// ods sheet charts on the wire (SPEC-132 "The chart model", "Writing", "New
// chart parts"): THIS FILE is the normative source for how an embedded
// chart object's `Object N/content.xml` becomes a SheetChartContent's visual
// and references, how an object is patched in place (path 2: the five
// address attributes and the `local-table` data cache, every other byte
// kept), the `draw:notify-on-update-of-ranges` list a chart frame carries,
// and the canonical parts of a chart object created here — hand-authored
// goldens in chartGolden.ts.
//
// SPEC-132 SKELETON: the signatures and contracts below are final; the bodies
// are filled in by the ods chart wire change.
import type { OdfChartRefs } from "../../core/chartPartOdf";
import type { ChartVisual } from "../../core/charts";
import type { DocumentTheme } from "../../core/theme";
import { unsupportedChartVisual } from "../charts";
import type { SheetChartRefs } from "../model";

/** What reading one chart object gives. */
export interface OdsChartRead {
  visual: ChartVisual;
  /** Absent: the object's addresses point into its own `local-table` (a
   * literal chart), or it is out of the catalogue. */
  refs?: SheetChartRefs;
  /** Some address has no data in the `local-table`: refreshed for display
   * at open. */
  uncached?: true;
}

/** Reads a chart object's content.xml: core parseOdfChartXml → visual (null
 * → the text as `unsupportedChartXml`, no refs); readOdfChartRefs → refs. */
export function readOdsChart(contentXml: string): OdsChartRead {
  return { visual: unsupportedChartVisual(contentXml) };
}

/** `chart:series@chart:values-cell-range-address` → values,
 * `@chart:label-cell-address` → name, `chart:domain@table:cell-range-address`
 * → categories (scatter x), else the x axis's `chart:categories@table:
 * cell-range-address` → every series' categories; `chart:plot-area@table:
 * cell-range-address` → plotArea; plotVisibleOnly = !include-hidden-cells
 * (the plot area's chart style; absent = ODF's default, true → false) —
 * each through odsRangeAddressToCanonical. Undefined when an address points
 * into the object's own `local-table` or none is present. */
export function readOdfChartRefs(contentXml: string): SheetChartRefs | undefined {
  void contentXml;
  return undefined;
}

/** Path 2 ("Writing"): the five address attributes rewritten from `refs`
 * through canonicalToOdsRangeAddress (refs only) and, with `visual`, the
 * `local-table` replaced by core odfChartTableXml(visual) — every other
 * byte the source's; the identical string when nothing changes. */
export function patchOdfChartContent(contentXml: string, refs: SheetChartRefs, visual?: ChartVisual): string {
  void refs;
  void visual;
  return contentXml;
}

/** The wire refs core odfChartContentXml writes for a regenerated or new
 * object: each canonical reference through canonicalToOdsRangeAddress. */
export function odfChartWireRefs(refs: SheetChartRefs, visual: ChartVisual): OdfChartRefs {
  void visual;
  return { series: refs.series.map(() => ({})) };
}

/** A regenerated or new object's content.xml: odfChartContentXml(visual, w,
 * h, theme, odfChartWireRefs(refs)). */
export function odsChartContentXml(visual: ChartVisual, widthCm: number, heightCm: number, theme: DocumentTheme | undefined, refs: SheetChartRefs | undefined): string {
  void visual;
  void widthCm;
  void heightCm;
  void theme;
  void refs;
  return "";
}

/** The `draw:notify-on-update-of-ranges` list a frame carries for `refs`
 * (ODS address form, space-separated): the plot area's range when present,
 * else every distinct series reference in order. Undefined for no refs. */
export function notifyRangesOf(refs: SheetChartRefs | undefined): string | undefined {
  void refs;
  return undefined;
}

/** A frame element's text with its `draw:object@draw:notify-on-update-of-
 * ranges` set to `ranges` (removed when undefined); the identical string
 * when it already says so. */
export function withNotifyRanges(frameXml: string, ranges: string | undefined): string {
  void ranges;
  return frameXml;
}

/** The `draw:object` a created chart's `draw:frame` holds ("New chart
 * parts", the golden's): `<draw:object draw:notify-on-update-of-ranges="…"
 * xlink:href="./Object {n}" xlink:type="simple" xlink:show="embed"
 * xlink:actuate="onLoad"/>` (no notify attribute for undefined ranges). The
 * ods drawing writer wraps it in the frame with SPEC-131's anchor
 * attributes. */
export function chartObjectElementXml(objectDir: string, notifyRanges: string | undefined): string {
  void objectDir;
  void notifyRanges;
  return "";
}

/** The parts of a chart object created here at `Object {n}/` ("New chart
 * parts", the golden's): content.xml plus the styles.xml/meta.xml the
 * golden settles on — name and text each. */
export function newOdsChartObjectParts(objectDir: string, contentXml: string): { name: string; text: string }[] {
  void objectDir;
  void contentXml;
  return [];
}

/** The manifest entries of a created object (`Object {n}/` with the chart
 * media type, then each of its parts). */
export function odsChartManifestEntries(objectDir: string): string {
  void objectDir;
  return "";
}

/** `Object {n}`, n = 1 + the highest across `taken` (part names). */
export function nextOdsObjectDir(taken: Iterable<string>): string {
  let max = 0;
  for (const name of taken) {
    const match = /^Object (\d+)\//.exec(name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `Object ${max + 1}`;
}
