// ODS conditional formats (SPEC-129 "ODS wire mapping") — THIS FILE is the
// normative source for how a table's `calcext:conditional-formats` (and,
// in a table without one, the legacy `style:map` conditions of its cells'
// automatic styles) read into Sheet.conditionalFormats and the dxfs they
// apply, and how the model's rules write back.
//
// Reading (readOdsTableConditionalFormats — the reader after its scan, per
// table, and SPEC-129's reader upgrade over a dirty sheet's retained text):
// - `calcext:conditional-formats` is a child of `table:table`, found from
//   the end of the table text (LibreOffice writes it after the rows); a
//   match that is not the table's own child is not one. Its names are
//   matched under every prefix the content root binds to the calcext
//   namespace (`calcext` itself when unbound: a table text read alone). Each
//   `calcext:conditional-format` is one ConditionalFormat: its ranges from
//   `calcext:target-range-address` (a space-separated list, each entry
//   through parseOdsRangeAddress; the sheet part is the table's own), its
//   rules from its entries in document order, its `sourceXml` the element
//   with every byte before it back to the previous format's end (the start
//   tag's for the first). Priorities follow document order across the
//   table (renumberPriorities keeps every `sourceXml`).
// - Entries:
//   | `calcext:condition@calcext:value` | kind |
//   | `=f` `!=f` `>f` `>=f` `<f` `<=f`, `between(f1,f2)`, `not-between(f1,f2)` | cellIs |
//   | `formula-is(f)` | expression |
//   | `duplicate` / `unique` | duplicates |
//   | `top-elements(n)` `bottom-elements(n)` `top-percent(n)` `bottom-percent(n)` | top10 |
//   | `above-average` `below-average` `above-equal-average` `below-equal-average` | aboveAverage |
//   | `begins-with("t")` `ends-with("t")` `contains-text("t")` `not-contains-text("t")` | text |
//   | `is-error` / `is-no-error` | errors |
//   `calcext:date-is@calcext:date` (today … next-year) → timePeriod;
//   `calcext:color-scale` › 2–3 `calcext:color-scale-entry` (value, type,
//   colour) → colorScale; `calcext:data-bar` (positive, negative and axis
//   colours, `gradient` — default true —, `axis-position` automatic /
//   middle / none, `show-value`, `min-length` 0 and `max-length` 100 by
//   default; direction "context") › 2 `calcext:formatting-entry` → dataBar;
//   `calcext:icon-set@calcext:icon-set-type` (`show-value`) › one
//   `calcext:formatting-entry` per icon, `calcext:gte="false"` meaning `>`
//   (no attribute: `≥`), never reversed → iconSet. Entry types minimum,
//   maximum, auto-minimum, auto-maximum, number, percent, percentile,
//   formula ↔ min, max, autoMin, autoMax, num, percent, percentile,
//   formula.
// - Operands are OpenFormula as `table:formula` writes it minus `of:=`
//   (references bracketed): fromOdsFormula, then — for a condition, whose
//   operands are relative to `calcext:base-cell-address` — rebased onto the
//   format's origin (its first range's top-left) with shiftFormula. A
//   value-object formula is absolute and is not rebased.
// - `calcext:apply-style-name` (`calcext:style` on date-is, a legacy map's
//   `style:apply-style-name`) names a common cell style of styles.xml,
//   possibly encoded ("Accent_20_3" is "Accent 3"):
//   OdsStyleBook.conditionalStyle turns it into a Dxf (its own properties
//   and its ancestors' short of "Default") carrying the resolved
//   `sourceStyleName`, held ONCE per style name in the holder's `dxfs`
//   (appended copy-on-write, as internDxf appends — internDxf itself drops
//   source fields, so it cannot be the one to append them); a name no
//   common style answers to reads as an empty dxf under that name.
// - Anything else — an unknown value form or entry element, an icon set
//   this model does not name (or `NoIcons`, which LibreOffice has not), a
//   custom one, an operand the dialect cannot read — is a `preserved` rule
//   (where "ods") holding the entry element verbatim.
// - Legacy: a table with no `calcext:conditional-formats` reads the
//   `style:map`s of the automatic cell styles its cells use —
//   `cell-content()<op>v`, `cell-content-is-between(a,b)`,
//   `cell-content-is-not-between(a,b)` → cellIs, `is-true-formula(f)` →
//   expression, operands relative to `style:base-cell-address` — one format
//   per mapped style, in first-use order, its rules the style's maps in
//   order. Its ranges are the cells using the style (a cell's own
//   `table:style-name`, else the row's default cell style, else the
//   column's, as LibreOffice applies them), accumulated as row-run ×
//   column-run rectangles — a repeated run is one rectangle and nothing is
//   ever materialised; a run covering the same columns as one ending on
//   the row run just above extends it. Legacy formats have no `sourceXml`.
//
// Writing (OdsConditionalFormatWriter, one per save — ods/writer.ts hands
// it to serializeTable, which calls splice for every rebuilt or new table;
// a clean table is its source):
// - The table's `calcext:conditional-formats` (never `style:map`) is
//   re-emitted where the source had it (among the leading or trailing
//   children), a new one as the table's last child, none when no rule is
//   writable. Byte-identical when every format still has its `sourceXml`
//   and there are as many as the source element holds (commands drop the
//   source of what they touch, so those are its children as read);
//   otherwise its start tag and closing text are kept (a container under
//   another prefix of the namespace holds the `calcext:` formats written
//   into it — ods/writer.ts declares `xmlns:calcext` on a root lacking it)
//   and the formats are written in priority order — rules taken by
//   ascending priority, consecutive rules of one format grouped into one
//   element (ODS numbers priorities by document order, so a format whose
//   rules interleave another's is split)
//   — each group that is a whole format with its `sourceXml` from that
//   (through the sheet-name pass, as all retained text; only an ODS
//   element's, into a table whose source had a container — a model from
//   another format never writes its source here), else serialised:
//   `calcext:target-range-address` and `calcext:base-cell-address` (the
//   origin) in LibreOffice's form (`Sheet1.B2:Sheet1.B13`), operands
//   through toOdsFormula without `of:=`, relative to the origin. blanks →
//   `formula-is(LEN(TRIM(A1))=0)` (`>0` for no blanks; reads back as
//   expression), an aboveAverage with `stdDev` → `formula-is(AND(ISNUMBER
//   (A1),A1>AVERAGE(r)+k*STDEVP(r)))` over the format's ranges (below: `<`,
//   `−`), with A1 the origin. Not written (no wire form): `reverse` on an
//   icon set, `NoIcons`, a data bar's border colours and direction,
//   `stopIfTrue`, and a rule preserved from xlsx.
// - A dxf's style: its `sourceStyleName` while the target styles.xml
//   defines that common style; else a generated common style `de-cf<n>`
//   (`style:display-name` "Conditional <n>", parent "Default", the dxf's
//   properties through odsDxfProperties, a number format as a `de-N<n>`
//   data style written before it — equal dxfs share one) collected in
//   `stylesXml`, which the writer splices before `</office:styles>`;
//   "Default" when there is no styles.xml to add to. `n` comes from the
//   style registry's counter, past every `de-cf` name the package holds.
//
// Layering rule (boundary.spec.ts): framework-free, imports only
// spreadsheet siblings and src/lib/core.
import { DEFAULT_THEME, resolveColor, type DocumentTheme, type ModelColor } from "../../core/theme";
import { scanXml } from "../../core/xmlScanner";
import { colToLetters, formatA1, type CellRef, type GridRange } from "../address";
import {
  formatOrigin,
  ICON_SET_NAMES,
  iconSetSize,
  mintCfId,
  renumberPriorities,
  rulesByPriority,
  sheetFormats,
  type CfOperator,
  type CfRule,
  type CfRuleBody,
  type CfTimePeriod,
  type CfValue,
  type ConditionalFormat,
  type Dxf,
  type IconSetName,
} from "../conditionalFormats";
import { fromOdsFormula, toOdsFormula } from "../formula/dialect";
import { FormulaSyntaxError } from "../formula/parse";
import { sheetNameNeedsQuotes } from "../formula/print";
import { shiftFormula } from "../formula/references";
import { MAX_COLS, MAX_ROWS, type Sheet, type SpreadsheetModel } from "../model";
import { odsNumberStyleXml } from "../numberFormat/odsNumberStyles";
import { dxfKey } from "../styles";
import { escapeAttributeValue, startTagAt } from "../xmlSplice";
import { parseOdsRangeAddress, splitOdsRangeAddressList } from "./rangeAddress";
import { rootNamespaces } from "./sheetNamePass";
import { odsColor, odsDxfProperties, odsPropertyElements, odsStyleBookFromSource, type OdsStyleBook, type OdsStyleMap, type OdsStyleRegistry } from "./styles";

// --- the wire's vocabularies --------------------------------------------------------------------

const OPERATORS: readonly (readonly [string, CfOperator])[] = [
  ["<=", "lessThanOrEqual"],
  [">=", "greaterThanOrEqual"],
  ["!=", "notEqual"],
  ["<", "lessThan"],
  ["=", "equal"],
  [">", "greaterThan"],
];

const OPERATOR_WRITE: Readonly<Record<Exclude<CfOperator, "between" | "notBetween">, string>> = {
  equal: "=",
  notEqual: "!=",
  greaterThan: ">",
  greaterThanOrEqual: ">=",
  lessThan: "<",
  lessThanOrEqual: "<=",
};

const AVERAGES: Readonly<Record<string, { below: boolean; equalAverage: boolean }>> = {
  "above-average": { below: false, equalAverage: false },
  "below-average": { below: true, equalAverage: false },
  "above-equal-average": { below: false, equalAverage: true },
  "below-equal-average": { below: true, equalAverage: true },
};

const RANKS: Readonly<Record<string, { bottom: boolean; percent: boolean }>> = {
  "top-elements": { bottom: false, percent: false },
  "bottom-elements": { bottom: true, percent: false },
  "top-percent": { bottom: false, percent: true },
  "bottom-percent": { bottom: true, percent: true },
};

type TextOperator = Extract<CfRuleBody, { kind: "text" }>["operator"];

const TEXTS: Readonly<Record<string, TextOperator>> = {
  "begins-with": "beginsWith",
  "ends-with": "endsWith",
  "contains-text": "contains",
  "not-contains-text": "notContains",
};

const TEXT_WRITE: Readonly<Record<TextOperator, string>> = {
  beginsWith: "begins-with",
  endsWith: "ends-with",
  contains: "contains-text",
  notContains: "not-contains-text",
};

const PERIOD_WRITE: Readonly<Record<CfTimePeriod, string>> = {
  today: "today",
  yesterday: "yesterday",
  tomorrow: "tomorrow",
  last7Days: "last-7-days",
  thisWeek: "this-week",
  lastWeek: "last-week",
  nextWeek: "next-week",
  thisMonth: "this-month",
  lastMonth: "last-month",
  nextMonth: "next-month",
  thisYear: "this-year",
  lastYear: "last-year",
  nextYear: "next-year",
};

const PERIODS: ReadonlyMap<string, CfTimePeriod> = new Map(Object.entries(PERIOD_WRITE).map(([period, wire]) => [wire, period as CfTimePeriod]));

const VALUE_TYPE_WRITE: Readonly<Record<CfValue["type"], string>> = {
  min: "minimum",
  max: "maximum",
  autoMin: "auto-minimum",
  autoMax: "auto-maximum",
  num: "number",
  percent: "percent",
  percentile: "percentile",
  formula: "formula",
};

const VALUE_TYPES: ReadonlyMap<string, CfValue["type"]> = new Map(Object.entries(VALUE_TYPE_WRITE).map(([type, wire]) => [wire, type as CfValue["type"]]));

/** The value types that carry no value (LibreOffice writes "0"). */
const VALUELESS: ReadonlySet<CfValue["type"]> = new Set(["min", "max", "autoMin", "autoMax"]);

/** LibreOffice's default data-bar colour (ScDataBarFormatData's
 * COL_LIGHTBLUE), for a bar that names none. */
const DEFAULT_BAR_COLOR = "#0000ff";

// --- reading ------------------------------------------------------------------------------------

/** What reading a table's conditional formats resolves against. */
export interface OdsCfReadContext {
  /** The workbook's styles: the common styles rules apply and the
   * automatic cell styles' `style:map`s (the reader's OdsStyleBook). */
  styles: Pick<OdsStyleBook, "conditionalStyle" | "cellStyleMaps" | "mappedCellStyles">;
  /** content.xml's root prefix → namespace declarations (an operand's
   * dialect prefix). */
  namespaces: ReadonlyMap<string, string>;
}

/** The read context of a retained package: its styles.xml text (undefined
 * when the package has none) and its content envelope's prefix, which holds
 * the root's declarations and the automatic styles. */
export function odsCfReadContext(stylesXml: string | undefined, contentPrefix: string): OdsCfReadContext {
  return { styles: odsStyleBookFromSource(contentPrefix, stylesXml), namespaces: rootNamespaces(contentPrefix) };
}

/** One table's conditional formats, from its text (a sheet's `sourceXml`:
 * the whitespace before `<table:table>` and the element), the dxfs they
 * apply interned into `holder.dxfs` (see the module header). */
export function readOdsTableConditionalFormats(tableXml: string, context: OdsCfReadContext, holder: { dxfs: Dxf[] }): ConditionalFormat[] {
  const reader = new CfReader(context, holder);
  const calcext = calcextPrefixes(context.namespaces);
  const found = locateConditionalFormats(tableXml, calcext);
  const formats = found ? reader.calcext(tableXml.slice(found.start, found.end), calcext) : reader.legacy(tableXml);
  return renumberPriorities(formats, { keepSource: true });
}

/** An element of a small part, as scanned: qualified names, decoded
 * attribute values, offsets into the scanned text. */
interface WireElement {
  name: string;
  attrs: Map<string, string>;
  start: number;
  /** The end of its start tag. */
  tagEnd: number;
  end: number;
  children: WireElement[];
}

/** The tree of a well-formed fragment with one root element; element and
 * attribute names under any of `calcext` (the prefixes bound to the calcext
 * namespace) are named `calcext:…`. */
function wireTree(xml: string, calcext: ReadonlySet<string> = new Set(["calcext"])): WireElement | null {
  const qualified = (name: string): string => {
    const colon = name.indexOf(":");
    return colon > 0 && calcext.has(name.slice(0, colon)) ? `calcext:${name.slice(colon + 1)}` : name;
  };
  let root: WireElement | null = null;
  const stack: WireElement[] = [];
  scanXml(xml, (event) => {
    if (event.type === "open") {
      const element: WireElement = {
        name: qualified(event.name),
        attrs: new Map(event.attrs.map((attr) => [qualified(attr.name), attr.value])),
        start: event.start,
        tagEnd: event.end,
        end: event.end,
        children: [],
      };
      const parent = stack.at(-1);
      if (parent) {
        parent.children.push(element);
      } else {
        root ??= element;
      }
      if (!event.selfClosing) {
        stack.push(element);
      }
    } else if (event.type === "close") {
      const element = stack.pop();
      if (element) {
        element.end = event.end;
      }
    }
  });
  return root;
}

/** Whether `text` (a run of sibling content) is balanced — so an element
 * it follows shares their parent. */
function balanced(text: string): boolean {
  if (!text.includes("<")) {
    return true;
  }
  try {
    scanXml(`<w>${text}</w>`, () => undefined);
    return true;
  } catch {
    return false;
  }
}

const CONDITIONAL_FORMATS = ":conditional-formats";

/** LibreOffice's Calc extension namespace. */
const CALCEXT_NAMESPACE = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

/** The prefixes the content root binds to the calcext namespace — and
 * `calcext` itself unless bound to another (a table text read alone). */
function calcextPrefixes(namespaces: ReadonlyMap<string, string>): ReadonlySet<string> {
  const prefixes = new Set([...namespaces].flatMap(([prefix, uri]) => (uri === CALCEXT_NAMESPACE ? [prefix] : [])));
  if (!namespaces.has("calcext")) {
    prefixes.add("calcext");
  }
  return prefixes;
}

/** The table's own `calcext:conditional-formats` (under any prefix of
 * `calcext`): its offsets in the table text; null when the table has none.
 * Searched from the end; a match must be followed by balanced siblings up
 * to the table's close tag (so it is the table's child). */
function locateConditionalFormats(tableXml: string, calcext: ReadonlySet<string>): { start: number; end: number } | null {
  const tableClose = tableXml.lastIndexOf("</");
  let at = tableXml.lastIndexOf(CONDITIONAL_FORMATS, tableClose);
  while (at > 0) {
    let nameStart = at;
    while (nameStart > 0 && /[\w.-]/.test(tableXml[nameStart - 1]!)) {
      nameStart -= 1;
    }
    const prefix = tableXml.slice(nameStart, at);
    const follows = tableXml[at + CONDITIONAL_FORMATS.length] ?? "";
    if (calcext.has(prefix) && /[\s/>]/.test(follows)) {
      let start = -1;
      let end = -1;
      if (tableXml.startsWith("</", nameStart - 2)) {
        start = tableXml.lastIndexOf(`<${prefix}${CONDITIONAL_FORMATS}`, nameStart - 2);
        end = tableXml.indexOf(">", at) + 1;
      } else if (tableXml[nameStart - 1] === "<") {
        // A self-closing element (a paired one's close tag is met first).
        const tag = startTagAt(tableXml, nameStart - 1);
        if (tag.endsWith("/>")) {
          start = nameStart - 1;
          end = start + tag.length;
        }
      }
      if (start >= 0 && end > 0 && end <= tableClose && balanced(tableXml.slice(end, tableClose))) {
        return { start, end };
      }
    }
    at = tableXml.lastIndexOf(CONDITIONAL_FORMATS, nameStart - 2);
  }
  return null;
}

/** `inner`'s arguments split at the commas outside quotes, brackets and
 * parentheses. */
function splitArguments(inner: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let quote = "";
  let from = 0;
  for (let i = 0; i < inner.length; i += 1) {
    const ch = inner[i]!;
    if (quote !== "") {
      if (ch === quote) {
        quote = "";
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === "(" || ch === "[" || ch === "{") {
      depth += 1;
    } else if (ch === ")" || ch === "]" || ch === "}") {
      depth -= 1;
    } else if (ch === "," && depth === 0) {
      out.push(inner.slice(from, i));
      from = i + 1;
    }
  }
  out.push(inner.slice(from));
  return out;
}

/** `name(args)` → the arguments; null when `value` is not that call. */
function callArguments(value: string, name: string): string[] | null {
  return value.startsWith(`${name}(`) && value.endsWith(")") ? splitArguments(value.slice(name.length + 1, -1)) : null;
}

/** An OpenFormula string literal's text, or null. */
function stringLiteral(expression: string): string | null {
  const match = /^"((?:[^"]|"")*)"$/.exec(expression.trim());
  return match ? match[1]!.replaceAll('""', '"') : null;
}

/** A cell address (`Sheet1.B2`, `$Sheet1.$B$2`, `.B2`), or null. */
function cellAddress(text: string | undefined): CellRef | null {
  const ref = text === undefined ? null : parseOdsRangeAddress(text.trim());
  return ref && ref.start.row !== undefined && ref.start.col !== undefined ? { row: ref.start.row, col: ref.start.col } : null;
}

/** `calcext:target-range-address` → its ranges (unreadable entries
 * dropped; a whole column or row spans the grid). */
function targetRanges(list: string | undefined): GridRange[] {
  const ranges: GridRange[] = [];
  const pieces = list === undefined ? [] : splitOdsRangeAddressList(list);
  for (let i = 0; i < pieces.length; i += 2) {
    const ref = pieces[i] === "" ? null : parseOdsRangeAddress(pieces[i]!);
    if (!ref) {
      continue;
    }
    const end = ref.end ?? ref.start;
    const rows = [ref.start.row ?? 0, end.row ?? MAX_ROWS - 1];
    const cols = [ref.start.col ?? 0, end.col ?? MAX_COLS - 1];
    ranges.push({ top: Math.min(rows[0]!, rows[1]!), left: Math.min(cols[0]!, cols[1]!), bottom: Math.max(rows[0]!, rows[1]!), right: Math.max(cols[0]!, cols[1]!) });
  }
  return ranges;
}

/** A rule body and the style it applies (formatting kinds). */
interface ReadEntry {
  body: CfRuleBody;
  style?: string;
}

class CfReader {
  private priority = 1;

  constructor(
    private readonly context: OdsCfReadContext,
    private readonly holder: { dxfs: Dxf[] },
  ) {}

  /** The formats of a `calcext:conditional-formats` element's text, its
   * names under any of `prefixes` (see calcextPrefixes). */
  calcext(xml: string, prefixes: ReadonlySet<string>): ConditionalFormat[] {
    const root = wireTree(xml, prefixes);
    const prefix = "calcext";
    if (!root) {
      return [];
    }
    const formats: ConditionalFormat[] = [];
    let previousEnd = root.tagEnd;
    for (const child of root.children) {
      if (child.name !== `${prefix}:conditional-format`) {
        continue; // carried as the text before the next format
      }
      const ranges = targetRanges(child.attrs.get(`${prefix}:target-range-address`));
      const origin = formatOrigin({ ranges });
      const rules = child.children.map((entry) => this.rule(this.entry(entry, prefix, origin), xml.slice(entry.start, entry.end)));
      formats.push({ id: mintCfId(), ranges, rules, sourceXml: xml.slice(previousEnd, child.end) });
      previousEnd = child.end;
    }
    return formats;
  }

  /** A rule with its id, the next priority and its dxf; a preserved one
   * (`xml` verbatim) when the entry did not read. */
  private rule(entry: ReadEntry | null, xml: string): CfRule {
    const rule = { id: mintCfId(), priority: this.priority, ...(entry?.body ?? { kind: "preserved", xml, where: "ods" }) } as CfRule;
    this.priority += 1;
    if (entry?.style !== undefined) {
      rule.dxfId = this.dxfId(entry.style);
    }
    return rule;
  }

  private entry(element: WireElement, prefix: string, origin: CellRef): ReadEntry | null {
    const attr = (local: string): string | undefined => element.attrs.get(`${prefix}:${local}`);
    switch (element.name.slice(prefix.length + 1)) {
      case "condition": {
        const base = cellAddress(attr("base-cell-address")) ?? origin;
        const body = this.condition(attr("value") ?? "", (expression) => this.relative(expression, base, origin));
        if (!body) {
          return null;
        }
        const style = attr("apply-style-name");
        return style === undefined ? { body } : { body, style };
      }
      case "date-is": {
        const period = PERIODS.get(attr("date") ?? "");
        if (period === undefined) {
          return null;
        }
        const style = attr("style");
        return style === undefined ? { body: { kind: "timePeriod", period } } : { body: { kind: "timePeriod", period }, style };
      }
      case "color-scale": {
        const stops: Extract<CfRuleBody, { kind: "colorScale" }>["stops"] = [];
        for (const child of element.children) {
          const at = this.value(child, prefix);
          const color = odsColor(child.attrs.get(`${prefix}:color`));
          if (child.name !== `${prefix}:color-scale-entry` || !at || color === undefined) {
            return null;
          }
          stops.push({ at, color });
        }
        return stops.length === 2 || stops.length === 3 ? { body: { kind: "colorScale", stops } } : null;
      }
      case "data-bar":
        return this.dataBar(element, prefix, attr);
      case "icon-set":
        return this.iconSet(element, prefix, attr);
      default:
        return null;
    }
  }

  private dataBar(element: WireElement, prefix: string, attr: (local: string) => string | undefined): ReadEntry | null {
    const entries = element.children.map((child) => (child.name === `${prefix}:formatting-entry` ? this.value(child, prefix) : null));
    const [min, max] = entries;
    if (entries.length !== 2 || !min || !max) {
      return null;
    }
    const axis = attr("axis-position");
    const length = (local: string, fallback: number): number => {
      const value = Number(attr(local));
      return attr(local) !== undefined && Number.isFinite(value) ? value : fallback;
    };
    const bar: Extract<CfRuleBody, { kind: "dataBar" }>["bar"] = {
      min,
      max,
      color: odsColor(attr("positive-color")) ?? DEFAULT_BAR_COLOR,
      gradient: attr("gradient") !== "false",
      axis: axis === "middle" || axis === "none" ? axis : "automatic",
      direction: "context",
      minLengthPct: length("min-length", 0),
      maxLengthPct: length("max-length", 100),
      showValue: attr("show-value") !== "false",
    };
    const negative = odsColor(attr("negative-color"));
    if (negative !== undefined) {
      bar.negativeColor = negative;
    }
    const axisColor = odsColor(attr("axis-color"));
    if (axisColor !== undefined) {
      bar.axisColor = axisColor;
    }
    return { body: { kind: "dataBar", bar } };
  }

  private iconSet(element: WireElement, prefix: string, attr: (local: string) => string | undefined): ReadEntry | null {
    const set = attr("icon-set-type") as IconSetName | undefined;
    if (set === undefined || !ICON_SET_NAMES.includes(set) || set === "NoIcons" || attr("custom") === "true") {
      return null;
    }
    const thresholds: Extract<CfRuleBody, { kind: "iconSet" }>["icons"]["thresholds"] = [];
    for (const child of element.children) {
      const at = child.name === `${prefix}:formatting-entry` ? this.value(child, prefix) : null;
      if (!at) {
        return null;
      }
      thresholds.push({ at, gte: child.attrs.get(`${prefix}:gte`) !== "false" });
    }
    if (thresholds.length !== iconSetSize(set)) {
      return null;
    }
    return { body: { kind: "iconSet", icons: { set, thresholds, reverse: false, showValue: attr("show-value") !== "false" } } };
  }

  /** A `calcext:value`/`calcext:type` pair as a value object, or null. */
  private value(element: WireElement, prefix: string): CfValue | null {
    const type = VALUE_TYPES.get(element.attrs.get(`${prefix}:type`) ?? "");
    if (type === undefined) {
      return null;
    }
    if (VALUELESS.has(type)) {
      return { type };
    }
    const text = (element.attrs.get(`${prefix}:value`) ?? "").trim();
    if (type !== "formula" && text !== "" && Number.isFinite(Number(text))) {
      return { type, value: text };
    }
    const formula = this.formula(text);
    return formula === null ? null : { type, value: formula };
  }

  /** A `calcext:value` condition → its rule body; `operand` reads (and
   * rebases) one operand. */
  private condition(value: string, operand: (expression: string) => string | null): CfRuleBody | null {
    const both = (args: string[] | null): string[] | null => {
      if (args?.length !== 2) {
        return null;
      }
      const formulas = args.map(operand);
      return formulas.every((formula) => formula !== null) ? (formulas as string[]) : null;
    };
    for (const [name, operator] of [
      ["between", "between"],
      ["not-between", "notBetween"],
    ] as const) {
      const args = callArguments(value, name);
      if (args) {
        const formulas = both(args);
        return formulas ? { kind: "cellIs", operator, formulas } : null;
      }
    }
    for (const [wire, operator] of OPERATORS) {
      if (value.startsWith(wire)) {
        const formula = operand(value.slice(wire.length));
        return formula === null ? null : { kind: "cellIs", operator, formulas: [formula] };
      }
    }
    if (value === "duplicate" || value === "unique") {
      return { kind: "duplicates", unique: value === "unique" };
    }
    if (value === "is-error" || value === "is-no-error") {
      return { kind: "errors", not: value === "is-no-error" };
    }
    const average = AVERAGES[value];
    if (average) {
      return { kind: "aboveAverage", ...average };
    }
    const expression = callArguments(value, "formula-is");
    if (expression) {
      const formula = expression.length === 1 ? operand(expression[0]!) : null;
      return formula === null ? null : { kind: "expression", formula };
    }
    for (const [name, rank] of Object.entries(RANKS)) {
      const args = callArguments(value, name);
      if (args) {
        const n = args.length === 1 && args[0]!.trim() !== "" ? Number(args[0]) : Number.NaN;
        return Number.isFinite(n) && n > 0 ? { kind: "top10", ...rank, rank: n } : null;
      }
    }
    for (const [name, operator] of Object.entries(TEXTS)) {
      const args = callArguments(value, name);
      if (args) {
        const text = args.length === 1 ? this.text(args[0]!) : null;
        return text === null ? null : { kind: "text", operator, text };
      }
    }
    return null;
  }

  /** A text rule's operand: a string literal's text, or a number's. */
  private text(expression: string): string | null {
    const literal = stringLiteral(expression);
    if (literal !== null) {
      return literal;
    }
    const formula = this.formula(expression);
    return formula !== null && Number.isFinite(Number(formula)) ? formula : null;
  }

  /** An operand as canonical text, or null when the dialect cannot read it. */
  private formula(expression: string): string | null {
    const text = expression.trim();
    if (text === "") {
      return null;
    }
    try {
      return fromOdsFormula(text, this.context.namespaces);
    } catch (error) {
      if (error instanceof FormulaSyntaxError) {
        return null;
      }
      throw error;
    }
  }

  /** A relative operand written at `base`, rebased onto `origin`. */
  private relative(expression: string, base: CellRef, origin: CellRef): string | null {
    const formula = this.formula(expression);
    if (formula === null) {
      return null;
    }
    try {
      return shiftFormula(formula, origin.row - base.row, origin.col - base.col);
    } catch (error) {
      if (error instanceof FormulaSyntaxError) {
        return null;
      }
      throw error;
    }
  }

  /** The holder's dxf for a common style name (see the module header). */
  private dxfId(name: string): number {
    const resolved = this.context.styles.conditionalStyle(name);
    const dxf: Dxf = { ...resolved?.dxf, sourceStyleName: resolved?.name ?? name };
    const key = dxfKey(dxf);
    const found = this.holder.dxfs.findIndex((each) => each.sourceStyleName === dxf.sourceStyleName && dxfKey(each) === key);
    if (found >= 0) {
      return found;
    }
    this.holder.dxfs = [...this.holder.dxfs, dxf];
    return this.holder.dxfs.length - 1;
  }

  // --- legacy ---

  /** The formats a table's cells' `style:map`s make (see the module
   * header); [] when no automatic style the table names has one. */
  legacy(tableXml: string): ConditionalFormat[] {
    const styles = this.context.styles;
    // Only a table whose text names a mapped style is scanned.
    const names = styles.mappedCellStyles();
    if (!names.some((name) => tableXml.includes(`style-name="${escapeAttributeValue(name)}"`))) {
      return [];
    }
    const regions = mappedRegions(tableXml, (name) => styles.cellStyleMaps(name) !== undefined);
    const formats: ConditionalFormat[] = [];
    for (const [style, ranges] of regions) {
      const origin = formatOrigin({ ranges });
      const rules = styles.cellStyleMaps(style)!.map((map) => this.rule(this.legacyEntry(map, origin), ""));
      formats.push({ id: mintCfId(), ranges, rules: rules.filter((rule) => rule.kind !== "preserved") });
    }
    return formats.filter((format) => format.rules.length > 0);
  }

  private legacyEntry(map: OdsStyleMap, origin: CellRef): ReadEntry | null {
    const base = cellAddress(map.baseCellAddress) ?? origin;
    const operand = (expression: string): string | null => this.relative(expression, base, origin);
    const condition = map.condition.trim();
    let body: CfRuleBody | null = null;
    const content = /^cell-content\(\)\s*(<=|>=|!=|<>|<|>|=)(.*)$/s.exec(condition);
    if (content) {
      const operator = OPERATORS.find(([wire]) => wire === (content[1] === "<>" ? "!=" : content[1]))![1];
      const formula = operand(content[2]!);
      body = formula === null ? null : { kind: "cellIs", operator, formulas: [formula] };
    } else {
      for (const [name, operator] of [
        ["cell-content-is-between", "between"],
        ["cell-content-is-not-between", "notBetween"],
      ] as const) {
        const args = callArguments(condition, name);
        if (args?.length === 2) {
          const formulas = args.map(operand);
          body = formulas.every((formula) => formula !== null) ? { kind: "cellIs", operator, formulas: formulas as string[] } : null;
        }
      }
      const expression = callArguments(condition, "is-true-formula");
      if (expression?.length === 1) {
        const formula = operand(expression[0]!);
        body = formula === null ? null : { kind: "expression", formula };
      }
    }
    // A legacy condition this model cannot read is dropped (it has no
    // element of its own to carry); `rule` makes it a preserved one, which
    // legacy() filters out.
    return body ? { body, style: map.applyStyleName } : null;
  }
}

/** The cells of a table using each mapped automatic cell style, as
 * rectangles (see the module header), styles in first-use order. */
function mappedRegions(tableXml: string, mapped: (name: string) => boolean): Map<string, GridRange[]> {
  const regions = new Map<string, { ranges: GridRange[]; open: Map<string, GridRange> }>();
  let pending: { style: string; left: number; right: number }[] = [];
  const segment = (style: string, left: number, right: number): void => {
    const last = pending.at(-1);
    if (last && last.style === style && last.right + 1 === left) {
      last.right = right;
    } else {
      pending.push({ style, left, right });
    }
  };
  const commit = (top: number, bottom: number): void => {
    for (const { style, left, right } of pending) {
      let region = regions.get(style);
      if (!region) {
        region = { ranges: [], open: new Map() };
        regions.set(style, region);
      }
      const key = `${left}:${right}`;
      const above = region.open.get(key);
      if (above && above.bottom === top - 1) {
        above.bottom = bottom;
      } else {
        const range = { top, left, bottom, right };
        region.ranges.push(range);
        region.open.set(key, range);
      }
    }
    pending = [];
  };

  // Elements are matched by the qualified names the table's own prefix
  // gives (the text is a fragment: its prefixes are unbound).
  const qname = /<([^\s/>]+)/.exec(tableXml)?.[1] ?? "table:table";
  const prefix = qname.includes(":") ? qname.slice(0, qname.indexOf(":") + 1) : "";

  // The column runs whose default cell style is mapped.
  const columns: { min: number; max: number; style: string }[] = [];
  let nextColumn = 0;
  let nextRow = 0;
  let row: { top: number; repeat: number; col: number; default?: string } | null = null;
  type Frame = "table" | "group" | "row" | "skip";
  const stack: Frame[] = [];
  const local = (name: string): string | undefined => (name.startsWith(prefix) ? name.slice(prefix.length) : undefined);
  const attr = (attrs: readonly { name: string; value: string }[], name: string): string | undefined => attrs.find((each) => each.name === `${prefix}${name}`)?.value;
  const repeat = (attrs: readonly { name: string; value: string }[], name: string): number => {
    const value = Number(attr(attrs, name));
    return Number.isInteger(value) && value > 1 ? value : 1;
  };
  /** Columns left..right of the current row, unstyled or in `style`: the
   * style named, else the row's default, else each column's. */
  const cells = (current: NonNullable<typeof row>, left: number, right: number, own: string | undefined): void => {
    const style = own ?? current.default;
    if (style !== undefined) {
      if (mapped(style)) {
        segment(style, left, right);
      }
      return;
    }
    for (const column of columns) {
      if (column.max >= left && column.min <= right) {
        segment(column.style, Math.max(column.min, left), Math.min(column.max, right));
      }
    }
  };
  const endRow = (): void => {
    if (row && row.top < MAX_ROWS) {
      // A row whose cells stop short of the columns: the rest are unstyled cells.
      if (row.col < Math.min(nextColumn, MAX_COLS)) {
        cells(row, row.col, Math.min(nextColumn, MAX_COLS) - 1, undefined);
      }
      commit(row.top, Math.min(row.top + row.repeat, MAX_ROWS) - 1);
    }
    pending = [];
    row = null;
  };

  scanXml(tableXml, (event) => {
    if (event.type === "text") {
      return undefined;
    }
    if (event.type === "close") {
      if (stack.pop() === "row") {
        endRow();
      }
      return undefined;
    }
    const parent = stack.at(-1);
    let frame: Frame = "skip";
    if (parent === undefined) {
      frame = "table";
    } else if (parent === "table" || parent === "group") {
      const name = local(event.name);
      if (name === "table-column") {
        const count = repeat(event.attrs, "number-columns-repeated");
        const style = attr(event.attrs, "default-cell-style-name");
        if (style !== undefined && mapped(style) && nextColumn < MAX_COLS) {
          columns.push({ min: nextColumn, max: Math.min(nextColumn + count, MAX_COLS) - 1, style });
        }
        nextColumn += count;
      } else if (name === "table-row") {
        row = { top: nextRow, repeat: repeat(event.attrs, "number-rows-repeated"), col: 0 };
        const rowDefault = attr(event.attrs, "default-cell-style-name");
        if (rowDefault !== undefined) {
          row.default = rowDefault;
        }
        nextRow += row.repeat;
        frame = "row";
        if (event.selfClosing) {
          endRow();
        }
      } else if (name !== undefined && /^table-(?:rows|header-rows|row-group|columns|header-columns|column-group)$/.test(name)) {
        frame = "group";
      }
    } else if (parent === "row" && row) {
      const name = local(event.name);
      if (name === "table-cell" || name === "covered-table-cell") {
        const count = repeat(event.attrs, "number-columns-repeated");
        const left = row.col;
        row.col += count;
        if (left < MAX_COLS) {
          cells(row, left, Math.min(left + count, MAX_COLS) - 1, attr(event.attrs, "style-name"));
        }
      }
    }
    if (!event.selfClosing) {
      stack.push(frame);
    }
    return undefined;
  });
  return new Map([...regions].map(([style, region]) => [style, region.ranges]));
}

// --- writing ------------------------------------------------------------------------------------

/** The styles.xml a save adds generated conditional styles to. */
export interface OdsCfWriteTarget {
  /** Whether it defines a common table-cell style of this name. */
  definesStyle(name: string): boolean;
  /** Whether there is one to add to (a package may lack styles.xml). */
  appendable: boolean;
}

/** A source `calcext:conditional-formats` element in parts: its start tag,
 * how many formats it holds, the text after its last format and its close
 * tag. */
interface SourceElement {
  open: string;
  formats: number;
  tail: string;
  close: string;
}

function sourceElement(xml: string): SourceElement {
  const root = wireTree(xml)!;
  if (root.tagEnd === root.end && xml.slice(root.start, root.end).endsWith("/>")) {
    const open = xml.slice(root.start, root.tagEnd).replace(/\s*\/>$/, ">");
    return { open, formats: 0, tail: "", close: `</${root.name}>` };
  }
  const closeStart = xml.lastIndexOf("</");
  // Its formats under any prefix (a container of another prefix holds the
  // `calcext:` ones a previous save serialised into it).
  const formats = root.children.filter((child) => child.name.endsWith(":conditional-format"));
  const lastEnd = formats.at(-1)?.end ?? root.tagEnd;
  return { open: xml.slice(root.start, root.tagEnd), formats: formats.length, tail: xml.slice(lastEnd, closeStart), close: xml.slice(closeStart, root.end) };
}

/** The first top-level `*:conditional-formats` element of a run of sibling
 * content (a table's leading or trailing children), by offsets. */
function topLevelConditionalFormats(text: string): { start: number; end: number } | null {
  if (!text.includes(CONDITIONAL_FORMATS)) {
    return null;
  }
  let depth = 0;
  let found: { start: number; end: number } | null = null;
  let open = -1;
  scanXml(`<w>${text}</w>`, (event) => {
    if (event.type === "open") {
      if (depth === 1 && open < 0 && event.name.endsWith(CONDITIONAL_FORMATS)) {
        if (event.selfClosing) {
          found = { start: event.start - 3, end: event.end - 3 };
          return "stop";
        }
        open = event.start - 3;
      }
      if (!event.selfClosing) {
        depth += 1;
      }
    } else if (event.type === "close") {
      depth -= 1;
      if (depth === 1 && open >= 0) {
        found = { start: open, end: event.end - 3 };
        return "stop";
      }
    }
    return undefined;
  });
  return found;
}

/** Whether a format's `sourceXml` is an ODS format element (whitespace
 * first) — only then, and only into a table that had a container, is it
 * re-emitted: a model from another format carries no ODS source. */
function isOdsFormatSource(sourceXml: string | undefined): sourceXml is string {
  return sourceXml !== undefined && /^\s*<[\w.-]+:conditional-format[\s/>]/.test(sourceXml);
}

/** A sheet name as an ODS address writes it (quoted when it must be). */
function odsSheetName(name: string): string {
  return sheetNameNeedsQuotes(name) || name.includes(".") ? `'${name.replaceAll("'", "''")}'` : name;
}

/** A cell in LibreOffice's address form: `Sheet1.B2`. */
function odsCell(sheet: string, row: number, col: number): string {
  return `${sheet}.${colToLetters(col)}${row + 1}`;
}

/** Ranges as `calcext:target-range-address`: `Sheet1.B2:Sheet1.B13`
 * entries, space-separated. */
function odsTargetRanges(sheetName: string, ranges: readonly GridRange[]): string {
  const sheet = odsSheetName(sheetName);
  return ranges
    .map((range) => {
      const start = odsCell(sheet, range.top, range.left);
      return range.top === range.bottom && range.left === range.right ? start : `${start}:${odsCell(sheet, range.bottom, range.right)}`;
    })
    .join(" ");
}

/** A canonical formula as a condition operand: OpenFormula without `of:=`. */
function odsOperand(canonical: string): string {
  return toOdsFormula(canonical.replace(/^=/, "")).slice("of:=".length);
}

/** An absolute cell (`$B$2`). */
function absoluteCell(row: number, col: number): string {
  return `$${colToLetters(col)}$${row + 1}`;
}

/** An absolute range (`$B$2:$B$13`). */
function absoluteRange(range: GridRange): string {
  const start = absoluteCell(range.top, range.left);
  return range.top === range.bottom && range.left === range.right ? start : `${start}:${absoluteCell(range.bottom, range.right)}`;
}

/** Whether ODS can carry a rule (see the module header). */
function writable(rule: CfRule): boolean {
  if (rule.kind === "preserved") {
    return rule.where === "ods";
  }
  return rule.kind !== "iconSet" || rule.icons.set !== "NoIcons";
}

/** A value object's `calcext:value` and `calcext:type` attributes. */
function valueAttributes(at: CfValue): string {
  if (VALUELESS.has(at.type)) {
    return ` calcext:value="0" calcext:type="${VALUE_TYPE_WRITE[at.type]}"`;
  }
  const text = (at.value ?? "0").trim();
  if (at.type !== "formula" && text !== "" && Number.isFinite(Number(text))) {
    return ` calcext:value="${escapeAttributeValue(text)}" calcext:type="${VALUE_TYPE_WRITE[at.type]}"`;
  }
  // A formula where a number belongs is written as the formula it is.
  return ` calcext:value="${escapeAttributeValue(odsOperand(text))}" calcext:type="formula"`;
}

/** One save's conditional formats (see the module header). */
export class OdsConditionalFormatWriter {
  private readonly theme: DocumentTheme;
  private readonly byDxf = new Map<number, string>();
  private readonly byKey = new Map<string, string>();
  private readonly dataStyles = new Map<string, string>();
  private generated = "";

  constructor(
    private readonly model: Pick<SpreadsheetModel, "dxfs" | "theme">,
    private readonly registry: OdsStyleRegistry,
    private readonly target: OdsCfWriteTarget,
  ) {
    this.theme = model.theme ?? DEFAULT_THEME;
  }

  /** The generated common styles (and their data styles), for styles.xml's
   * `office:styles`; "" when none was needed. */
  get stylesXml(): string {
    return this.generated;
  }

  /** A rebuilt table's leading and trailing children (verbatim text, the
   * sheet-name pass applied) with its `calcext:conditional-formats` written
   * from the model: replaced where it was, removed (with the whitespace
   * before it) when nothing is writable, else appended as the last child.
   * `passRetained` is the pass for the model's `sourceXml`. */
  splice(sheet: Sheet, leading: string, trailing: string, passRetained: (text: string) => string = (text) => text): { leading: string; trailing: string } {
    const inLeading = topLevelConditionalFormats(leading);
    const found = inLeading ?? topLevelConditionalFormats(trailing);
    const host = inLeading ? leading : trailing;
    const element = this.element(sheet, found ? host.slice(found.start, found.end) : undefined, passRetained);
    if (!found) {
      return { leading, trailing: trailing + element };
    }
    let start = found.start;
    if (element === "") {
      while (start > 0 && /\s/.test(host[start - 1]!)) {
        start -= 1;
      }
    }
    const next = host.slice(0, start) + element + host.slice(found.end);
    return inLeading ? { leading: next, trailing } : { leading, trailing: next };
  }

  /** The table's element ("" for none). */
  private element(sheet: Sheet, source: string | undefined, pass: (text: string) => string): string {
    const formats = sheetFormats(sheet);
    const parts = source === undefined ? undefined : sourceElement(source);
    // Unchanged: a format a command touched has lost its source, a deleted
    // one is missing, a new one has none — so every format still carrying
    // its source, as many as the element holds, are its children as read.
    if (parts && parts.formats === formats.length && formats.every((format) => isOdsFormatSource(format.sourceXml))) {
      return source!;
    }
    const groups: { format: ConditionalFormat; rules: CfRule[] }[] = [];
    for (const { format, rule } of rulesByPriority(formats)) {
      if (!writable(rule) || format.ranges.length === 0) {
        continue;
      }
      const last = groups.at(-1);
      if (last?.format === format) {
        last.rules.push(rule);
      } else {
        groups.push({ format, rules: [rule] });
      }
    }
    if (groups.length === 0) {
      return "";
    }
    const body = groups
      .map(({ format, rules }) =>
        parts !== undefined && isOdsFormatSource(format.sourceXml) && rules.length === format.rules.length ? pass(format.sourceXml) : this.formatXml(sheet.name, format, rules),
      )
      .join("");
    return (parts?.open ?? "<calcext:conditional-formats>") + body + (parts?.tail ?? "") + (parts?.close ?? "</calcext:conditional-formats>");
  }

  private formatXml(sheetName: string, format: ConditionalFormat, rules: readonly CfRule[]): string {
    const origin = formatOrigin(format);
    const entries = rules.map((rule) => this.entryXml(sheetName, format, origin, rule)).join("");
    return `<calcext:conditional-format calcext:target-range-address="${escapeAttributeValue(odsTargetRanges(sheetName, format.ranges))}">${entries}</calcext:conditional-format>`;
  }

  private entryXml(sheetName: string, format: ConditionalFormat, origin: CellRef, rule: CfRule): string {
    switch (rule.kind) {
      case "preserved":
        return rule.xml;
      case "timePeriod":
        return `<calcext:date-is calcext:style="${escapeAttributeValue(this.styleName(rule.dxfId))}" calcext:date="${PERIOD_WRITE[rule.period]}"/>`;
      case "colorScale":
        return (
          `<calcext:color-scale>` +
          rule.stops.map((stop) => `<calcext:color-scale-entry${valueAttributes(stop.at)} calcext:color="${this.color(stop.color)}"/>`).join("") +
          `</calcext:color-scale>`
        );
      case "dataBar": {
        const bar = rule.bar;
        let attributes = ` calcext:positive-color="${this.color(bar.color)}"`;
        if (bar.negativeColor !== undefined) {
          attributes += ` calcext:negative-color="${this.color(bar.negativeColor)}"`;
        }
        if (bar.axisColor !== undefined) {
          attributes += ` calcext:axis-color="${this.color(bar.axisColor)}"`;
        }
        if (!bar.gradient) {
          attributes += ` calcext:gradient="false"`;
        }
        if (bar.axis !== "automatic") {
          attributes += ` calcext:axis-position="${bar.axis}"`;
        }
        if (!bar.showValue) {
          attributes += ` calcext:show-value="false"`;
        }
        if (bar.minLengthPct !== 0) {
          attributes += ` calcext:min-length="${bar.minLengthPct}"`;
        }
        if (bar.maxLengthPct !== 100) {
          attributes += ` calcext:max-length="${bar.maxLengthPct}"`;
        }
        return `<calcext:data-bar${attributes}><calcext:formatting-entry${valueAttributes(bar.min)}/><calcext:formatting-entry${valueAttributes(bar.max)}/></calcext:data-bar>`;
      }
      case "iconSet": {
        const icons = rule.icons;
        return (
          `<calcext:icon-set calcext:icon-set-type="${icons.set}"${icons.showValue ? "" : ` calcext:show-value="false"`}>` +
          icons.thresholds.map((threshold) => `<calcext:formatting-entry${valueAttributes(threshold.at)}${threshold.gte ? "" : ` calcext:gte="false"`}/>`).join("") +
          `</calcext:icon-set>`
        );
      }
      default:
        return (
          `<calcext:condition calcext:apply-style-name="${escapeAttributeValue(this.styleName(rule.dxfId))}" ` +
          `calcext:value="${escapeAttributeValue(conditionValue(rule, format, origin))}" calcext:base-cell-address="${escapeAttributeValue(odsCell(odsSheetName(sheetName), origin.row, origin.col))}"/>`
        );
    }
  }

  private color(color: ModelColor): string {
    return resolveColor(color, this.theme).toLowerCase();
  }

  /** The common style a rule's dxf is written as (see the module header). */
  private styleName(dxfId: number | undefined): string {
    if (dxfId === undefined) {
      return "Default";
    }
    let name = this.byDxf.get(dxfId);
    if (name === undefined) {
      const dxf = this.model.dxfs?.[dxfId];
      if (dxf?.sourceStyleName !== undefined && this.target.definesStyle(dxf.sourceStyleName)) {
        name = dxf.sourceStyleName;
      } else if (dxf === undefined || !this.target.appendable) {
        name = "Default";
      } else {
        name = this.generatedStyle(dxf);
      }
      this.byDxf.set(dxfId, name);
    }
    return name;
  }

  /** A `de-cf<n>` common style for `dxf` (equal dxfs share one). */
  private generatedStyle(dxf: Dxf): string {
    const key = dxfKey(dxf);
    let name = this.byKey.get(key);
    if (name !== undefined) {
      return name;
    }
    name = this.registry.freshName("cf");
    this.byKey.set(key, name);
    const code = dxf.numberFormat?.trim();
    let dataStyle: string | undefined;
    if (code !== undefined && code !== "" && !/^general$/i.test(code)) {
      dataStyle = this.dataStyles.get(code);
      if (dataStyle === undefined) {
        dataStyle = this.registry.freshName("N");
        this.dataStyles.set(code, dataStyle);
        this.generated += odsNumberStyleXml(code, dataStyle);
      }
    }
    const body = odsPropertyElements(odsDxfProperties(dxf, this.theme));
    const attributes =
      ` style:name="${name}" style:display-name="Conditional ${name.slice("de-cf".length)}" style:family="table-cell" style:parent-style-name="Default"` +
      (dataStyle === undefined ? "" : ` style:data-style-name="${dataStyle}"`);
    this.generated += body === "" ? `<style:style${attributes}/>` : `<style:style${attributes}>${body}</style:style>`;
    return name;
  }
}

/** A formatting rule's `calcext:value` (see the module header). */
function conditionValue(rule: CfRule, format: ConditionalFormat, origin: CellRef): string {
  const cell = formatA1(origin);
  switch (rule.kind) {
    case "cellIs": {
      const [a = "0", b = "0"] = rule.formulas.map(odsOperand);
      if (rule.operator === "between" || rule.operator === "notBetween") {
        return `${rule.operator === "between" ? "between" : "not-between"}(${a},${b})`;
      }
      return OPERATOR_WRITE[rule.operator] + a;
    }
    case "expression":
      return `formula-is(${odsOperand(rule.formula)})`;
    case "duplicates":
      return rule.unique ? "unique" : "duplicate";
    case "top10":
      return `${rule.bottom ? "bottom" : "top"}-${rule.percent ? "percent" : "elements"}(${rule.rank})`;
    case "aboveAverage": {
      if (rule.stdDev === undefined) {
        return `${rule.below ? "below" : "above"}-${rule.equalAverage ? "equal-" : ""}average`;
      }
      const ranges = format.ranges.map(absoluteRange).join(",");
      const bound = `AVERAGE(${ranges})${rule.below ? "-" : "+"}${rule.stdDev}*STDEVP(${ranges})`;
      const compare = `${rule.below ? "<" : ">"}${rule.equalAverage ? "=" : ""}`;
      return `formula-is(${odsOperand(`AND(ISNUMBER(${cell}),${cell}${compare}${bound})`)})`;
    }
    case "text":
      return `${TEXT_WRITE[rule.operator]}("${rule.text.replaceAll('"', '""')}")`;
    case "blanks":
      return `formula-is(${odsOperand(`LEN(TRIM(${cell}))${rule.not ? ">" : "="}0`)})`;
    case "errors":
      return rule.not ? "is-no-error" : "is-error";
    default:
      return "";
  }
}
