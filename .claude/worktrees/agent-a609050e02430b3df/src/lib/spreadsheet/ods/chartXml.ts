// ods sheet charts on the wire (SPEC-132 "The chart model", "Writing", "New
// chart parts"): THIS FILE is the normative source for how an embedded
// chart object's `Object N/content.xml` becomes a SheetChartContent's visual
// and references, how an object is patched in place (path 2: the five
// address attributes and the `local-table` data cache, every other byte
// kept), the `draw:notify-on-update-of-ranges` list a chart frame carries,
// and the canonical parts of a chart object created here — hand-authored
// goldens in chartGolden.ts.
//
// - Where the addresses live: `chart:plot-area@table:cell-range-address`;
//   the first `chart:axis chart:dimension="x"`'s first `chart:categories@
//   table:cell-range-address`; per `chart:series` of the plot area, in
//   order, `@chart:values-cell-range-address`, `@chart:label-cell-address`
//   and its first `chart:domain@table:cell-range-address`.
// - A series' categories are its own domain; in a scatter chart a series
//   without one inherits the first series domain (LibreOffice writes a domain
//   shared by every series once, on the first, and its import reads it so);
//   otherwise the x axis's categories.
// - Each address becomes canonical through odsRangeAddressToCanonical
//   (absolute markers as written). An empty address is ODF's "no range" —
//   absent; one with no canonical form (LibreOffice's deleted-sheet
//   `$#REF!.A1`) reads "#REF!", resolving to nothing, so its series keeps its
//   cache. An address into the object's own `local-table` makes the whole
//   chart a literal one: no refs.
// - plotVisibleOnly = the plot area's chart style (office:automatic-styles)
//   says `chart:include-hidden-cells="false"`; absent is ODF's default, true.
//   A regenerated or new object that plots visible cells only writes that
//   style (`ch1`, chartGolden.ts) around core's document, which has none.
// - Writing a reference: canonicalToOdsRangeAddress; one with no ODS form
//   (`#REF!`, a name) and a series' literal name write the empty address,
//   never a `local-table` one beside sheet ones (LibreOffice detaches a chart
//   with an unresolvable range from its sheet; this reader would read the
//   chart as literal).
// - Patching (path 2) splices start tags at scanXml's offsets: an attribute
//   is rewritten only when its canonical reading differs from the ref (so an
//   unchanged address keeps its bytes, `Data.A1:Data.B5` included), added
//   only where a ref exists; a missing x axis categories element, a missing
//   x axis (before the plot area's first axis or series) and a scatter
//   series' missing domain (its first child) are inserted the same way.
import { CHART_NAMESPACE, CHART_OBJECT_CONTENT_TYPE, odfChartContentXml, odfChartTableXml, parseOdfChartXml, type OdfChartRefs } from "../../core/chartPartOdf";
import { DEFAULT_CHART_COLORS, type ChartSeries, type ChartVisual } from "../../core/charts";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TABLE_NAMESPACE } from "../../core/odfXml";
import { DEFAULT_THEME, type DocumentTheme } from "../../core/theme";
import { scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { unsupportedChartVisual } from "../charts";
import type { SheetChartRefs, SheetChartSeriesRefs } from "../model";
import { applySplices, attributeValue, escapeAttributeValue, setAttribute, type Splice } from "../xmlSplice";
import { canonicalToOdsRangeAddress, odsRangeAddressToCanonical, splitOdsRangeAddressList } from "./rangeAddress";

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

// --- locating the addresses ----------------------------------------------------------

/** One element as scanXml met it: its start tag's span and, for an element
 * with content, its end tag's. */
interface Tag {
  name: string;
  start: number;
  end: number;
  selfClosing: boolean;
  attrs: XmlScanAttribute[];
  closeStart: number;
  closeEnd: number;
}

interface ChartLayout {
  chartClass?: string;
  chart?: Tag;
  plotArea?: Tag;
  xAxis?: { tag: Tag; title?: Tag; categories?: Tag };
  /** The plot area's first `chart:axis`. */
  firstAxis?: Tag;
  /** Where the plot area's axes end: its first series (or stock marker,
   * wall, floor). */
  afterAxes?: number;
  series: { tag: Tag; domain?: Tag }[];
  localTable?: Tag;
  /** Whether chart:chart holds any `table:table`. */
  hasTable: boolean;
  /** Automatic style name → its `chart:include-hidden-cells`. */
  includeHiddenCells: Map<string, string>;
}

type Role = "root" | "styles" | "style" | "body" | "officeChart" | "chart" | "plotArea" | "xAxis" | "series" | "other";

const AFTER_AXES = new Set(["series", "stock-gain-marker", "stock-loss-marker", "stock-range-line", "wall", "floor"]);

function attr(attrs: readonly XmlScanAttribute[], ns: string, local: string): string | undefined {
  return attrs.find((each) => each.ns === ns && each.local === local)?.value;
}

/** One pass over a chart object's content.xml; null when it is not XML. */
function scanChartLayout(xml: string): ChartLayout | null {
  const layout: ChartLayout = { series: [], hasTable: false, includeHiddenCells: new Map() };
  const stack: { role: Role; tag: Tag; styleName?: string }[] = [];
  try {
    scanXml(xml, (event) => {
      if (event.type === "text") {
        return;
      }
      if (event.type === "close") {
        const top = stack.pop();
        if (top) {
          top.tag.closeStart = event.start;
          top.tag.closeEnd = event.end;
        }
        return;
      }
      const parent = stack.at(-1);
      const tag: Tag = { name: event.name, start: event.start, end: event.end, selfClosing: event.selfClosing, attrs: event.attrs, closeStart: event.end, closeEnd: event.end };
      const is = (ns: string, local: string): boolean => event.ns === ns && event.local === local;
      let role: Role = "other";
      let styleName: string | undefined;
      switch (parent?.role) {
        case undefined:
          role = is(OFFICE_NAMESPACE, "document-content") ? "root" : "other";
          break;
        case "root":
          role = is(OFFICE_NAMESPACE, "automatic-styles") ? "styles" : is(OFFICE_NAMESPACE, "body") ? "body" : "other";
          break;
        case "styles":
          if (is(STYLE_NAMESPACE, "style")) {
            role = "style";
            styleName = attr(event.attrs, STYLE_NAMESPACE, "name");
          }
          break;
        case "style": {
          const include = is(STYLE_NAMESPACE, "chart-properties") ? attr(event.attrs, CHART_NAMESPACE, "include-hidden-cells") : undefined;
          if (include !== undefined && parent.styleName !== undefined) {
            layout.includeHiddenCells.set(parent.styleName, include);
          }
          break;
        }
        case "body":
          role = is(OFFICE_NAMESPACE, "chart") ? "officeChart" : "other";
          break;
        case "officeChart":
          if (is(CHART_NAMESPACE, "chart") && !layout.chart) {
            role = "chart";
            layout.chart = tag;
            layout.chartClass = attr(event.attrs, CHART_NAMESPACE, "class");
          }
          break;
        case "chart":
          if (is(CHART_NAMESPACE, "plot-area") && !layout.plotArea) {
            role = "plotArea";
            layout.plotArea = tag;
          } else if (is(TABLE_NAMESPACE, "table")) {
            layout.hasTable = true;
            if (attr(event.attrs, TABLE_NAMESPACE, "name") === LOCAL_TABLE && !layout.localTable) {
              layout.localTable = tag;
            }
          }
          break;
        case "plotArea":
          if (is(CHART_NAMESPACE, "axis")) {
            layout.firstAxis ??= tag;
            if (attr(event.attrs, CHART_NAMESPACE, "dimension") === "x" && !layout.xAxis) {
              role = "xAxis";
              layout.xAxis = { tag };
            }
          } else if (event.ns === CHART_NAMESPACE && AFTER_AXES.has(event.local)) {
            layout.afterAxes ??= event.start;
            if (event.local === "series") {
              role = "series";
              layout.series.push({ tag });
            }
          }
          break;
        case "xAxis":
          if (is(CHART_NAMESPACE, "title")) {
            layout.xAxis!.title ??= tag;
          } else if (is(CHART_NAMESPACE, "categories")) {
            layout.xAxis!.categories ??= tag;
          }
          break;
        case "series":
          if (is(CHART_NAMESPACE, "domain")) {
            layout.series.at(-1)!.domain ??= tag;
          }
          break;
        default:
          break;
      }
      if (!event.selfClosing) {
        stack.push(styleName === undefined ? { role, tag } : { role, tag, styleName });
      }
    });
  } catch (err) {
    if (err instanceof XmlScanError) {
      return null;
    }
    throw err;
  }
  return layout;
}

const LOCAL_TABLE = "local-table";

/** Whether an address list names the object's own `local-table`. */
function namesLocalTable(list: string): boolean {
  const pieces = splitOdsRangeAddressList(list);
  for (let i = 0; i < pieces.length; i += 2) {
    const entry = pieces[i]!.replace(/^\$/, "");
    if (entry.startsWith(`${LOCAL_TABLE}.`) || entry.startsWith(`'${LOCAL_TABLE}'.`)) {
      return true;
    }
  }
  return false;
}

/** An address attribute's canonical reading (see the module header):
 * undefined when absent or empty, "#REF!" when it has no canonical form. */
function canonicalOf(address: string | undefined): string | undefined {
  if (address === undefined || address.trim() === "") {
    return undefined;
  }
  return odsRangeAddressToCanonical(address) ?? "#REF!";
}

const rangeOf = (tag: Tag | undefined): string | undefined => (tag ? attr(tag.attrs, TABLE_NAMESPACE, "cell-range-address") : undefined);
const valuesOf = (tag: Tag): string | undefined => attr(tag.attrs, CHART_NAMESPACE, "values-cell-range-address");
const labelOf = (tag: Tag): string | undefined => attr(tag.attrs, CHART_NAMESPACE, "label-cell-address");

/** The refs a located object carries (see readOdfChartRefs). */
function refsOf(layout: ChartLayout): SheetChartRefs | undefined {
  if (!layout.plotArea) {
    return undefined;
  }
  const addresses = [rangeOf(layout.plotArea), rangeOf(layout.xAxis?.categories)];
  for (const { tag, domain } of layout.series) {
    addresses.push(valuesOf(tag), labelOf(tag), rangeOf(domain));
  }
  if (addresses.some((address) => address !== undefined && namesLocalTable(address))) {
    return undefined;
  }
  const axisCategories = canonicalOf(rangeOf(layout.xAxis?.categories));
  const scatter = layout.chartClass === "chart:scatter";
  let firstDomain: string | undefined;
  const series = layout.series.map(({ tag, domain }): SheetChartSeriesRefs => {
    const own = canonicalOf(rangeOf(domain));
    firstDomain ??= own;
    const read: SheetChartSeriesRefs = {};
    const name = canonicalOf(labelOf(tag));
    const categories = own ?? (scatter ? firstDomain : undefined) ?? axisCategories;
    const values = canonicalOf(valuesOf(tag));
    if (name !== undefined) {
      read.name = name;
    }
    if (categories !== undefined) {
      read.categories = categories;
    }
    if (values !== undefined) {
      read.values = values;
    }
    return read;
  });
  const plotArea = canonicalOf(rangeOf(layout.plotArea));
  if (plotArea === undefined && series.every((each) => Object.keys(each).length === 0)) {
    return undefined;
  }
  const style = attr(layout.plotArea.attrs, CHART_NAMESPACE, "style-name");
  const plotVisibleOnly = style !== undefined && layout.includeHiddenCells.get(style) === "false";
  return plotArea === undefined ? { series, plotVisibleOnly } : { series, plotArea, plotVisibleOnly };
}

// --- reading -----------------------------------------------------------------------

/** Reads a chart object's content.xml: core parseOdfChartXml → visual (null
 * → the text as `unsupportedChartXml`, no refs); readOdfChartRefs → refs. */
export function readOdsChart(contentXml: string): OdsChartRead {
  const layout = scanChartLayout(contentXml);
  const refs = layout ? refsOf(layout) : undefined;
  let parsed = parseOdfChartXml(contentXml);
  let cacheMissing = false;
  if (!parsed && refs && layout?.chart && !layout.chart.selfClosing && !layout.hasTable) {
    // A sheet chart saved without its cache: read over an empty one, and
    // refresh it from the cells for display.
    const empty = odfChartTableXml({ ...unsupportedChartVisual(""), series: refs.series.map(() => ({ name: "", values: [], color: DEFAULT_CHART_COLORS[0]! })) });
    parsed = parseOdfChartXml(contentXml.slice(0, layout.chart.closeStart) + empty + contentXml.slice(layout.chart.closeStart));
    cacheMissing = true;
  }
  if (!parsed) {
    return { visual: unsupportedChartVisual(contentXml) };
  }
  const visual: ChartVisual = { chartType: parsed.chartType, categories: parsed.categories, series: parsed.series, options: parsed.options };
  if (!refs || !layout) {
    return { visual };
  }
  // The cache's columns must line up with the part's series: LibreOffice's
  // scatter cache gives its x values columns of their own.
  const aligned = parsed.series.length === refs.series.length;
  if (!aligned) {
    visual.series = refs.series.map((_, i): ChartSeries => ({ name: `Series ${i + 1}`, values: [], color: seriesColor(layout, i) }));
  }
  if (refs.series.every((each) => each.categories === undefined)) {
    const count = Math.max(0, ...visual.series.map((each) => each.values.length));
    visual.categories = Array.from({ length: count }, (_, i) => String(i + 1));
  }
  const uncached = cacheMissing || !layout.localTable || parsed.categories.length === 0 || !aligned;
  return uncached ? { visual, refs, uncached: true } : { visual, refs };
}

/** A series' colour as core's reader takes it: `chart:color`, else the
 * default cycle. */
function seriesColor(layout: ChartLayout, index: number): string {
  const color = layout.series[index] ? attr(layout.series[index].tag.attrs, CHART_NAMESPACE, "color") : undefined;
  return color ? color.toLowerCase() : DEFAULT_CHART_COLORS[index % DEFAULT_CHART_COLORS.length]!;
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
  const layout = scanChartLayout(contentXml);
  return layout ? refsOf(layout) : undefined;
}

// --- writing references ------------------------------------------------------------

/** A canonical reference's wire text; undefined for none or no ODS form. */
function wireOf(canonical: string | undefined): string | undefined {
  return canonical === undefined ? undefined : (canonicalToOdsRangeAddress(canonical) ?? undefined);
}

/** An attribute's qualified name as the tag writes it, else `prefix:local`. */
function qnameIn(tag: Tag, ns: string, local: string, prefix: string): string {
  return tag.attrs.find((each) => each.ns === ns && each.local === local)?.name ?? `${prefix}:${local}`;
}

/** Whether an address read as `current` (canonicalOf) already says `ref`. */
function sameAddress(current: string | undefined, ref: string): boolean {
  const wire = canonicalToOdsRangeAddress(ref);
  if (wire === null) {
    return current === undefined || canonicalToOdsRangeAddress(current) === null;
  }
  return current !== undefined && (current === ref || canonicalToOdsRangeAddress(current) === wire);
}

/** Path 2 ("Writing"): the five address attributes rewritten from `refs`
 * through canonicalToOdsRangeAddress (refs only) and, with `visual`, the
 * `local-table` replaced by core odfChartTableXml(visual) — every other
 * byte the source's; the identical string when nothing changes. */
export function patchOdfChartContent(contentXml: string, refs: SheetChartRefs, visual?: ChartVisual): string {
  const layout = scanChartLayout(contentXml);
  if (!layout?.chart || !layout.plotArea) {
    return contentXml;
  }
  const edits = new Map<Tag, { text: string; children: string }>();
  const splices: Splice[] = [];
  const editOf = (tag: Tag): { text: string; children: string } => {
    let edit = edits.get(tag);
    if (!edit) {
      edit = { text: contentXml.slice(tag.start, tag.end), children: "" };
      edits.set(tag, edit);
    }
    return edit;
  };
  /** Sets one address attribute when the ref says otherwise. */
  const setAddress = (tag: Tag, ns: string, local: string, prefix: string, ref: string | undefined): void => {
    if (ref === undefined || sameAddress(canonicalOf(attr(tag.attrs, ns, local)), ref)) {
      return;
    }
    const edit = editOf(tag);
    edit.text = setAttribute(edit.text, qnameIn(tag, ns, local, prefix), wireOf(ref) ?? "");
  };
  /** Inserts a child element at `at` (inside `parent`). */
  const insertChild = (parent: Tag, at: number, xml: string): void => {
    if (parent.selfClosing) {
      editOf(parent).children += xml;
    } else {
      splices.push({ start: at, end: at, text: xml });
    }
  };
  const rangeElement = (local: string, ref: string): string => `<chart:${local} table:cell-range-address="${escapeAttributeValue(wireOf(ref) ?? "")}"/>`;

  setAddress(layout.plotArea, TABLE_NAMESPACE, "cell-range-address", "table", refs.plotArea);

  // Categories: a series' own domain, a scatter series' inherited one, else
  // the x axis (from the first series whose categories ride it).
  const scatter = layout.chartClass === "chart:scatter";
  const count = Math.min(refs.series.length, layout.series.length);
  const axisRef = refs.series.slice(0, count).find((each, i) => each.categories !== undefined && !layout.series[i]!.domain && !scatter)?.categories;
  const axisNow = canonicalOf(rangeOf(layout.xAxis?.categories));
  let axisAfter = axisNow;
  if (axisRef !== undefined && !sameAddress(axisNow, axisRef)) {
    axisAfter = axisRef;
    const axis = layout.xAxis;
    if (axis?.categories) {
      setAddress(axis.categories, TABLE_NAMESPACE, "cell-range-address", "table", axisRef);
    } else if (axis) {
      insertChild(axis.tag, axis.title ? axis.title.closeEnd : axis.tag.end, rangeElement("categories", axisRef));
    } else {
      const at = layout.firstAxis?.start ?? layout.afterAxes ?? layout.plotArea.closeStart;
      insertChild(layout.plotArea, at, `<chart:axis chart:dimension="x" chart:name="primary-x">${rangeElement("categories", axisRef)}</chart:axis>`);
    }
  }
  let firstDomainAfter: string | undefined;
  for (let i = 0; i < count; i += 1) {
    const ref = refs.series[i]!;
    const { tag, domain } = layout.series[i]!;
    setAddress(tag, CHART_NAMESPACE, "values-cell-range-address", "chart", ref.values);
    setAddress(tag, CHART_NAMESPACE, "label-cell-address", "chart", ref.name);
    if (domain) {
      setAddress(domain, TABLE_NAMESPACE, "cell-range-address", "table", ref.categories);
      firstDomainAfter ??= ref.categories ?? canonicalOf(rangeOf(domain));
    } else if (scatter && ref.categories !== undefined && !sameAddress(firstDomainAfter ?? axisAfter, ref.categories)) {
      insertChild(tag, tag.end, rangeElement("domain", ref.categories));
      firstDomainAfter ??= ref.categories;
    }
  }

  if (visual) {
    const table = odfChartTableXml(visual);
    const local = layout.localTable;
    if (local) {
      const end = local.selfClosing ? local.end : local.closeEnd;
      if (contentXml.slice(local.start, end) !== table) {
        splices.push({ start: local.start, end, text: table });
      }
    } else if (!layout.chart.selfClosing) {
      splices.push({ start: layout.chart.closeStart, end: layout.chart.closeStart, text: table });
    }
  }

  for (const [tag, edit] of edits) {
    let text = edit.text;
    if (edit.children !== "") {
      text = `${text.replace(/\s*\/>$/, ">")}${edit.children}</${tag.name}>`;
    }
    if (text !== contentXml.slice(tag.start, tag.end)) {
      splices.push({ start: tag.start, end: tag.end, text });
    }
  }
  return applySplices(contentXml, splices);
}

/** The wire refs core odfChartContentXml writes for a regenerated or new
 * object: each canonical reference through canonicalToOdsRangeAddress. */
export function odfChartWireRefs(refs: SheetChartRefs, visual: ChartVisual): OdfChartRefs {
  const scatter = visual.chartType === "scatter";
  const wire: OdfChartRefs = {
    // Every written series gets a sheet address or the empty one — never a
    // local-table address beside sheet ones (see the module header).
    series: visual.series.map((_, i) => {
      const ref = refs.series[i] ?? {};
      const entry: OdfChartRefs["series"][number] = { values: wireOf(ref.values) ?? "", label: wireOf(ref.name) ?? "" };
      const domain = scatter ? wireOf(ref.categories) : undefined;
      if (domain !== undefined) {
        entry.domain = domain;
      }
      return entry;
    }),
  };
  const plotArea = wireOf(refs.plotArea);
  if (plotArea !== undefined) {
    wire.plotArea = plotArea;
  }
  const categories = scatter ? undefined : wireOf(refs.series.find((each) => each.categories !== undefined)?.categories);
  if (categories !== undefined) {
    wire.categories = categories;
  }
  return wire;
}

/** The automatic style a chart plotting visible cells only names on its
 * plot area (chartGolden.ts): LibreOffice's `chart:include-hidden-cells="false"`. */
const VISIBLE_ONLY_STYLE_NAME = "ch1";
const VISIBLE_ONLY_STYLES =
  `<office:automatic-styles><style:style style:name="${VISIBLE_ONLY_STYLE_NAME}" style:family="chart">` +
  `<style:chart-properties chart:include-hidden-cells="false"/></style:style></office:automatic-styles>`;

/** Core's chart document with the visible-cells-only style spliced in: the
 * style namespace declared on the root (before `office:version`), the
 * automatic styles before `office:body`, the plot area naming the style
 * first. The text unchanged when it has no plot area. */
function withVisibleOnlyStyle(xml: string): string {
  const rootStart = xml.indexOf("<office:document-content");
  const rootEnd = rootStart < 0 ? -1 : xml.indexOf(">", rootStart);
  const body = xml.indexOf("<office:body>", rootEnd);
  const plotArea = body < 0 ? -1 : xml.indexOf("<chart:plot-area", body);
  if (rootEnd < 0 || body < 0 || plotArea < 0) {
    return xml;
  }
  const root = xml.slice(rootStart, rootEnd + 1);
  const declared = attributeValue(root, "xmlns:style") !== undefined;
  const version = root.indexOf(" office:version=");
  const styledRoot = declared
    ? root
    : version < 0
      ? setAttribute(root, "xmlns:style", STYLE_NAMESPACE)
      : `${root.slice(0, version)} xmlns:style="${STYLE_NAMESPACE}"${root.slice(version)}`;
  const tagAt = plotArea + "<chart:plot-area".length;
  return (
    xml.slice(0, rootStart) +
    styledRoot +
    xml.slice(rootEnd + 1, body) +
    VISIBLE_ONLY_STYLES +
    xml.slice(body, tagAt) +
    ` chart:style-name="${VISIBLE_ONLY_STYLE_NAME}"` +
    xml.slice(tagAt)
  );
}

/** A regenerated or new object's content.xml: odfChartContentXml(visual, w,
 * h, theme, odfChartWireRefs(refs)), with the visible-cells-only style when
 * the references plot visible cells only (core writes no automatic styles;
 * without it the chart reads back, here and in LibreOffice, as including
 * hidden cells). */
export function odsChartContentXml(visual: ChartVisual, widthCm: number, heightCm: number, theme: DocumentTheme | undefined, refs: SheetChartRefs | undefined): string {
  const xml = odfChartContentXml(visual, widthCm, heightCm, theme ?? DEFAULT_THEME, refs && odfChartWireRefs(refs, visual));
  return refs?.plotVisibleOnly ? withVisibleOnlyStyle(xml) : xml;
}

// --- the frame ---------------------------------------------------------------------

/** The `draw:notify-on-update-of-ranges` list a frame carries for `refs`
 * (ODS address form, space-separated): the plot area's range when present,
 * else every distinct series reference in order. Undefined for no refs. */
export function notifyRangesOf(refs: SheetChartRefs | undefined): string | undefined {
  if (!refs) {
    return undefined;
  }
  const plotArea = wireOf(refs.plotArea);
  if (plotArea !== undefined) {
    return plotArea;
  }
  const ranges: string[] = [];
  for (const series of refs.series) {
    for (const ref of [series.name, series.categories, series.values]) {
      const wire = wireOf(ref);
      if (wire !== undefined && !ranges.includes(wire)) {
        ranges.push(wire);
      }
    }
  }
  return ranges.length > 0 ? ranges.join(" ") : undefined;
}

/** A frame element's text with its `draw:object@draw:notify-on-update-of-
 * ranges` set to `ranges` (removed when undefined); the identical string
 * when it already says so — read as references, so LibreOffice's
 * `Data.A1:Data.B5` already says `$Data.A1:.B5`. */
export function withNotifyRanges(frameXml: string, ranges: string | undefined): string {
  let object: { start: number; end: number; name: string } | undefined;
  try {
    scanXml(frameXml, (event) => {
      if (event.type === "open" && (event.ns === DRAW_NAMESPACE || event.ns === undefined) && event.local === "object" && (event.ns !== undefined || event.name.endsWith(":object"))) {
        object = { start: event.start, end: event.end, name: event.name };
        return "stop";
      }
      return undefined;
    });
  } catch (err) {
    if (err instanceof XmlScanError) {
      return frameXml;
    }
    throw err;
  }
  if (!object) {
    return frameXml;
  }
  const tag = frameXml.slice(object.start, object.end);
  const colon = object.name.indexOf(":");
  const qname = `${colon < 0 ? "draw" : object.name.slice(0, colon)}:notify-on-update-of-ranges`;
  const current = attributeValue(tag, qname);
  if (ranges !== undefined && current !== undefined) {
    const now = odsRangeAddressToCanonical(current);
    if (current === ranges || (now !== null && now === odsRangeAddressToCanonical(ranges))) {
      return frameXml;
    }
  }
  const next = setAttribute(tag, qname, ranges ?? null);
  return next === tag ? frameXml : frameXml.slice(0, object.start) + next + frameXml.slice(object.end);
}

// --- new chart parts ---------------------------------------------------------------

/** The `draw:object` a created chart's `draw:frame` holds ("New chart
 * parts", the golden's): `<draw:object draw:notify-on-update-of-ranges="…"
 * xlink:href="./Object {n}" xlink:type="simple" xlink:show="embed"
 * xlink:actuate="onLoad"/>` (no notify attribute for undefined ranges). The
 * ods drawing writer wraps it in the frame with SPEC-131's anchor
 * attributes. */
export function chartObjectElementXml(objectDir: string, notifyRanges: string | undefined): string {
  const notify = notifyRanges === undefined ? "" : ` draw:notify-on-update-of-ranges="${escapeAttributeValue(notifyRanges)}"`;
  return `<draw:object${notify} xlink:href="./${escapeAttributeValue(objectDir)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;
}

/** The parts of a chart object created here at `Object {n}/` ("New chart
 * parts", the golden's): content.xml plus the styles.xml/meta.xml the
 * golden settles on — name and text each. The golden settles on content.xml
 * alone (see chartGolden.ts). */
export function newOdsChartObjectParts(objectDir: string, contentXml: string): { name: string; text: string }[] {
  return [{ name: `${objectDir}/content.xml`, text: contentXml }];
}

/** The manifest entries of a created object (`Object {n}/` with the chart
 * media type, then each of its parts). */
export function odsChartManifestEntries(objectDir: string): string {
  return (
    manifestEntry(`${objectDir}/`, ` manifest:version="1.2" manifest:media-type="${CHART_OBJECT_CONTENT_TYPE}"`) +
    newOdsChartObjectParts(objectDir, "")
      .map((part) => manifestEntry(part.name, ` manifest:media-type="text/xml"`))
      .join("")
  );
}

function manifestEntry(path: string, extra: string): string {
  return `<manifest:file-entry manifest:full-path="${escapeAttributeValue(path)}"${extra}/>`;
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
