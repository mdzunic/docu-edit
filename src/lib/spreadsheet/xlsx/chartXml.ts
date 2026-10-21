// xlsx sheet charts on the wire (SPEC-132 "The chart model", "Read-only
// charts", "Writing", "New chart parts"): THIS FILE is the normative source
// for how an `xl/charts/chartN.xml` part becomes a SheetChartContent's visual
// and references, how a part is patched in place (path 2: `c:f` texts and
// `c:strCache`/`c:numCache` only, every other byte kept), and the canonical
// wire forms of a chart created here (the part through core chartPartXml,
// the anchor's `xdr:graphicFrame`) — hand-authored goldens in chartGolden.ts.
//
// Reading:
// - the chart namespace is resolved from the part's own declarations (any
//   prefix, or none; transitional or strict URI), never assumed to be `c:`;
// - a part's series are the `c:ser` children of `c:plotArea`'s ONE chart-type
//   element (none or several — a combo — is not modelled: no references);
//   the i-th `c:ser` is series i, index-aligned with core parseChartXml's;
// - a series' references are exactly `c:tx/c:strRef/c:f` (name), `c:cat` or
//   `c:xVal` → `c:strRef`|`c:numRef`|`c:multiLvlStrRef` → `c:f`
//   (categories) and `c:val` or `c:yVal` → `c:numRef/c:f` (values) — direct
//   children only, so a `c15:` extension's `c15:f`, a filtered series' copy
//   or a data-label range is never mistaken for one. Each text is decoded
//   and read through fromXlsxFormula; an empty `c:f` is no reference, and
//   a literal (`c:tx/c:v`, `c:strLit`, `c:numLit`) leaves its slot absent;
// - `plotVisibleOnly` is `c:chart/c:plotVisOnly@val` ("0"/"false" → false;
//   absent, valueless or anything else → true, the schema's default);
// - `uncached`: a reference read above whose element has no cache child
//   (`c:strCache`/`c:numCache`/`c:multiLvlStrCache`);
// - a chart with references but no categories reference anywhere shows
//   Excel's numbering "1", "2", … for the longest series (what a refresh
//   computes); a multi-level chart — read-only, painted from its caches,
//   which parseChartXml cannot read — shows its cache's innermost level.
//
// Patching (path 2) touches only, per modelled reference: the `c:f` text,
// rewritten when fromXlsxFormula(its text) differs from the model's
// reference (so an unchanged one keeps its exact bytes, `$`s, spaces and
// entity spelling included); and, given `cellsOf`, the cache's `c:ptCount` +
// `c:pt` run, rewritten from the cells when it differs from them (numbers
// compared as numbers) — a `c:formatCode` before it and a `c:extLst` after
// it stay, a `c:pt`'s own formatCode is dropped only when the run changes,
// and a reference with no cache gains one right after its `c:f` (a
// `c:numCache` opening with `<c:formatCode>General</c:formatCode>`). New
// elements are written with the prefix the reference element is written
// with. A reference the model lacks (or the part lacks) is left alone: the
// patch never adds or removes a reference — a change of shape regenerates.
import { C_NAMESPACE, chartPartXml, parseChartXml, type ChartPartRefs } from "../../core/chartPartOoxml";
import type { ChartVisual } from "../../core/charts";
import { scanXml, XmlScanError } from "../../core/xmlScanner";
import { MULTI_LEVEL_CHART_REASON, PIVOT_CHART_REASON, unsupportedChartVisual, type ChartCell } from "../charts";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import type { SheetChartRefs, SheetChartSeriesRefs } from "../model";
import { applySplices, escapeAttributeText, escapeAttributeValue, escapeElementText, type Splice } from "../xmlSplice";
import { elementTree, nodeAttr, type XmlNode } from "./elementTree";

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

const STRICT_C_NAMESPACE = "http://purl.oclc.org/ooxml/drawingml/chart";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The `c:plotArea` children that are a chart-type element (core's list). */
const CHART_TYPE_ELEMENTS: ReadonlySet<string> = new Set([
  "barChart",
  "bar3DChart",
  "lineChart",
  "line3DChart",
  "areaChart",
  "area3DChart",
  "pieChart",
  "pie3DChart",
  "ofPieChart",
  "doughnutChart",
  "scatterChart",
  "radarChart",
  "stockChart",
  "bubbleChart",
  "surfaceChart",
  "surface3DChart",
]);
const CACHE_ELEMENTS: ReadonlySet<string> = new Set(["strCache", "numCache", "multiLvlStrCache"]);

function isChartNs(ns: string | undefined): boolean {
  return ns === C_NAMESPACE || ns === STRICT_C_NAMESPACE;
}

/** The first chart-namespace child of `node` with one of the local names. */
function chartChild(node: XmlNode | undefined, ...locals: string[]): XmlNode | undefined {
  return node?.children.find((child) => isChartNs(child.ns) && locals.includes(child.local));
}

function chartChildren(node: XmlNode | undefined, local: string): XmlNode[] {
  return node?.children.filter((child) => isChartNs(child.ns) && child.local === local) ?? [];
}

// --- the part's references on the wire ------------------------------------------------------

/** One reference as the part holds it. */
interface WireRef {
  /** `c:strRef` / `c:numRef` / `c:multiLvlStrRef`. */
  ref: XmlNode;
  f: XmlNode;
  cache: XmlNode | undefined;
  /** fromXlsxFormula of the `c:f` text. */
  canonical: string;
}

interface WireSeries {
  name?: WireRef;
  categories?: WireRef;
  values?: WireRef;
}

interface ChartWire {
  series: WireSeries[];
  plotVisibleOnly: boolean;
}

function wireRef(ref: XmlNode | undefined): WireRef | undefined {
  const f = chartChild(ref, "f");
  if (!ref || !f || f.text.trim() === "") {
    return undefined;
  }
  return { ref, f, cache: ref.children.find((child) => isChartNs(child.ns) && CACHE_ELEMENTS.has(child.local)), canonical: fromXlsxFormula(f.text) };
}

/** The part's series references and plotVisOnly; undefined when the part is
 * not well-formed or has no single chart-type element. */
function chartWire(partXml: string): ChartWire | undefined {
  const root = elementTree(partXml)?.[0];
  if (!root || root.local !== "chartSpace" || !isChartNs(root.ns)) {
    return undefined;
  }
  const chart = chartChild(root, "chart");
  const types = chartChild(chart, "plotArea")?.children.filter((child) => isChartNs(child.ns) && CHART_TYPE_ELEMENTS.has(child.local)) ?? [];
  if (types.length !== 1) {
    return undefined;
  }
  const series = chartChildren(types[0], "ser").map((ser): WireSeries => {
    const wire: WireSeries = {};
    const name = wireRef(chartChild(chartChild(ser, "tx"), "strRef"));
    const categories = wireRef(chartChild(chartChild(ser, "cat", "xVal"), "strRef", "numRef", "multiLvlStrRef"));
    const values = wireRef(chartChild(chartChild(ser, "val", "yVal"), "numRef"));
    if (name) {
      wire.name = name;
    }
    if (categories) {
      wire.categories = categories;
    }
    if (values) {
      wire.values = values;
    }
    return wire;
  });
  const plotVisOnly = chartChild(chart, "plotVisOnly");
  const val = plotVisOnly ? nodeAttr(plotVisOnly, "val")?.trim() : undefined;
  return { series, plotVisibleOnly: val !== "0" && val !== "false" };
}

function refsOfWire(wire: ChartWire): { refs: SheetChartRefs; uncached: boolean } | undefined {
  let named = false;
  let uncached = false;
  const series = wire.series.map((wired): SheetChartSeriesRefs => {
    const refs: SheetChartSeriesRefs = {};
    for (const slot of ["name", "categories", "values"] as const) {
      const ref = wired[slot];
      if (ref) {
        refs[slot] = ref.canonical;
        named = true;
        uncached ||= ref.cache === undefined;
      }
    }
    return refs;
  });
  return named ? { refs: { series, plotVisibleOnly: wire.plotVisibleOnly }, uncached } : undefined;
}

/** A multi-level category cache's innermost level (its first `c:lvl`). */
function innermostLevel(cache: XmlNode): string[] {
  const ptCount = chartChild(cache, "ptCount");
  const count = Number(ptCount ? nodeAttr(ptCount, "val") : undefined);
  const labels = Array.from<string>({ length: Number.isInteger(count) && count > 0 ? count : 0 }).fill("");
  for (const pt of chartChildren(chartChild(cache, "lvl"), "pt")) {
    const idx = Number(nodeAttr(pt, "idx"));
    if (Number.isInteger(idx) && idx >= 0 && idx < labels.length) {
      labels[idx] = chartChild(pt, "v")?.text ?? "";
    }
  }
  return labels;
}

// --- reading ---------------------------------------------------------------------------------

/** Reads a chart part: core parseChartXml → visual (null → the part text as
 * `unsupportedChartXml`, no refs); readChartPartRefs → refs;
 * chartReadOnlyReason. */
export function readXlsxChart(partXml: string): XlsxChartRead {
  const parsed = parseChartXml(partXml);
  if (!parsed) {
    return { visual: unsupportedChartVisual(partXml) };
  }
  const visual: ChartVisual = { chartType: parsed.chartType, categories: parsed.categories, series: parsed.series, options: parsed.options };
  const read: XlsxChartRead = { visual };
  const wire = chartWire(partXml);
  const found = wire && refsOfWire(wire);
  if (wire && found) {
    read.refs = found.refs;
    if (found.uncached) {
      read.uncached = true;
    }
    const categories = wire.series.find((series) => series.categories)?.categories;
    if (!categories) {
      // No categories reference: Excel numbers them, as a refresh would.
      const count = Math.max(0, ...visual.series.map((series) => series.values.length));
      visual.categories = Array.from({ length: count }, (_, i) => String(i + 1));
    } else if (visual.categories.length === 0 && categories.ref.local === "multiLvlStrRef" && categories.cache) {
      visual.categories = innermostLevel(categories.cache);
    }
  }
  const reason = chartReadOnlyReason(partXml);
  if (reason !== undefined) {
    read.readOnlyReason = reason;
  }
  return read;
}

/** Each `c:ser` of the part's one chart-type element, in order: `c:tx/
 * c:strRef/c:f` → name, `c:cat` or `c:xVal` (`c:strRef`/`c:numRef`/
 * `c:multiLvlStrRef`)`/c:f` → categories, `c:val` or `c:yVal` `c:numRef/c:f`
 * → values, each through fromXlsxFormula (a literal `c:tx/c:v` leaves name
 * absent); `c:plotVisOnly@val` (absent = true). Undefined when no series
 * names a reference. `uncached`: some `c:f` without a cache sibling. */
export function readChartPartRefs(partXml: string): { refs: SheetChartRefs; uncached: boolean } | undefined {
  const wire = chartWire(partXml);
  return wire && refsOfWire(wire);
}

/** "Read-only charts": PIVOT_CHART_REASON for a `c:pivotSource`,
 * MULTI_LEVEL_CHART_REASON for a `c:multiLvlStrRef`, else undefined. */
export function chartReadOnlyReason(partXml: string): string | undefined {
  let pivot = false;
  let multiLevel = false;
  try {
    scanXml(partXml, (event) => {
      if (event.type === "open" && isChartNs(event.ns)) {
        if (event.local === "pivotSource") {
          pivot = true;
          return "stop";
        }
        multiLevel ||= event.local === "multiLvlStrRef";
      }
      return undefined;
    });
  } catch (error) {
    if (!(error instanceof XmlScanError)) {
      throw error;
    }
  }
  if (pivot) {
    return PIVOT_CHART_REASON;
  }
  return multiLevel ? MULTI_LEVEL_CHART_REASON : undefined;
}

// --- patching (path 2) ------------------------------------------------------------------------

interface CachePoint {
  idx: number;
  v: string;
}

/** The prefix (colon included) an element is written with. */
function prefixOf(node: XmlNode): string {
  const colon = node.name.indexOf(":");
  return colon < 0 ? "" : node.name.slice(0, colon + 1);
}

/** The splice replacing an element's text content. */
function textSplice(xml: string, node: XmlNode, text: string): Splice {
  if (node.selfClosing) {
    return { start: node.start, end: node.end, text: `<${node.name}>${text}</${node.name}>` };
  }
  return { start: node.tagEnd, end: xml.lastIndexOf("</", node.end - 1), text };
}

/** A cache's points from the cells: one per non-blank cell, at its position
 * (a number cache: numbers only, written raw; a string cache: the text). */
function pointsOf(cells: readonly ChartCell[], numeric: boolean): CachePoint[] {
  const points: CachePoint[] = [];
  cells.forEach((cell, idx) => {
    if (cell.value === null) {
      return;
    }
    if (!numeric) {
      points.push({ idx, v: cell.text });
    } else if (typeof cell.value === "number") {
      points.push({ idx, v: String(cell.value) });
    }
  });
  return points;
}

function pointsXml(prefix: string, count: number, points: readonly CachePoint[]): string {
  return `<${prefix}ptCount val="${count}"/>` + points.map((pt) => `<${prefix}pt idx="${pt.idx}"><${prefix}v>${escapeElementText(pt.v)}</${prefix}v></${prefix}pt>`).join("");
}

/** Whether a cache's `c:ptCount` + `c:pt`s already say `count` + `points`. */
function cacheSays(cache: XmlNode, count: number, points: readonly CachePoint[], numeric: boolean): boolean {
  const ptCount = chartChild(cache, "ptCount");
  const pts = chartChildren(cache, "pt");
  if (!ptCount || Number(nodeAttr(ptCount, "val")) !== count || pts.length !== points.length) {
    return false;
  }
  return pts.every((pt, i) => {
    const want = points[i]!;
    const v = chartChild(pt, "v")?.text ?? "";
    return Number(nodeAttr(pt, "idx")) === want.idx && (v === want.v || (numeric && v.trim() !== "" && Number(v) === Number(want.v)));
  });
}

/** The splices rewriting an existing cache's `c:ptCount` + `c:pt` run. */
function cacheSplices(cache: XmlNode, prefix: string, count: number, points: readonly CachePoint[], numeric: boolean): Splice[] {
  if (cacheSays(cache, count, points, numeric)) {
    return [];
  }
  const run = cache.children.filter((child) => isChartNs(child.ns) && (child.local === "ptCount" || child.local === "pt"));
  const text = pointsXml(prefix, count, points);
  if (cache.selfClosing) {
    return [{ start: cache.start, end: cache.end, text: `<${cache.name}>${text}</${cache.name}>` }];
  }
  if (run.length === 0) {
    const at = chartChild(cache, "formatCode")?.end ?? cache.tagEnd;
    return [{ start: at, end: at, text }];
  }
  const start = run[0]!.start;
  const end = run.at(-1)!.end;
  if (cache.children.some((child) => !run.includes(child) && child.start > start && child.start < end)) {
    // Something sits among the points (never in a valid part): replace the
    // first, drop the rest, leave it be.
    return run.map((node, i) => ({ start: node.start, end: node.end, text: i === 0 ? text : "" }));
  }
  return [{ start, end, text }];
}

function patchReference(xml: string, wire: WireRef | undefined, canonical: string | undefined, cellsOf: ((canonical: string) => ChartCell[] | null) | undefined, splices: Splice[]): void {
  if (!wire || canonical === undefined) {
    return;
  }
  if (wire.canonical !== canonical) {
    splices.push(textSplice(xml, wire.f, escapeElementText(toXlsxFormula(canonical))));
  }
  if (!cellsOf || wire.ref.local === "multiLvlStrRef") {
    return;
  }
  const cells = cellsOf(canonical);
  if (!cells) {
    return;
  }
  const numeric = wire.ref.local === "numRef";
  const prefix = prefixOf(wire.ref);
  const points = pointsOf(cells, numeric);
  if (wire.cache) {
    splices.push(...cacheSplices(wire.cache, prefixOf(wire.cache), cells.length, points, numeric));
    return;
  }
  const cache = `${prefix}${numeric ? "numCache" : "strCache"}`;
  const formatCode = numeric ? `<${prefix}formatCode>General</${prefix}formatCode>` : "";
  splices.push({ start: wire.f.end, end: wire.f.end, text: `<${cache}>${formatCode}${pointsXml(prefix, cells.length, points)}</${cache}>` });
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
  const wire = chartWire(partXml);
  if (!wire) {
    return partXml;
  }
  const splices: Splice[] = [];
  wire.series.forEach((series, i) => {
    const want = refs.series[i];
    if (want) {
      patchReference(partXml, series.name, want.name, cellsOf, splices);
      patchReference(partXml, series.categories, want.categories, cellsOf, splices);
      patchReference(partXml, series.values, want.values, cellsOf, splices);
    }
  });
  return applySplices(partXml, splices);
}

// --- regenerated and new parts ----------------------------------------------------------------

/** The wire refs core chartPartXml writes for a regenerated or new part:
 * each canonical reference through toXlsxFormula. */
export function xlsxChartWireRefs(refs: SheetChartRefs): ChartPartRefs {
  return {
    series: refs.series.map((series) => {
      const wire: ChartPartRefs["series"][number] = {};
      for (const slot of ["name", "categories", "values"] as const) {
        const canonical = series[slot];
        if (canonical !== undefined) {
          wire[slot] = toXlsxFormula(canonical);
        }
      }
      return wire;
    }),
    plotVisibleOnly: refs.plotVisibleOnly,
  };
}

/** A regenerated or new chart part (path 1 / New chart parts):
 * chartPartXml(visual, xlsxChartWireRefs(refs)). */
export function xlsxChartPartXml(visual: ChartVisual, refs: SheetChartRefs | undefined): string {
  return chartPartXml(visual, refs && xlsxChartWireRefs(refs));
}

/** The canonical `xdr:graphicFrame` a created chart's anchor holds ("New
 * chart parts" — the golden's scaffold and Excel's zero `xdr:xfrm`). */
export function chartGraphicFrameXml(cNvPrId: number, name: string, relId: string): string {
  return (
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${cNvPrId}" name="${escapeAttributeText(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_NAMESPACE}"><c:chart xmlns:c="${C_NAMESPACE}" xmlns:r="${R_NAMESPACE}" r:id="${escapeAttributeValue(relId)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
  );
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
