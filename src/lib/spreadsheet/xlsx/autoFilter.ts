// The xlsx wire mapping of SPEC-128's autofilter and sort record ("xlsx
// wire mapping", normative there; filter.ts and sort.ts are the model's
// sources): what xlsx/sheetXml.ts's reader makes of a worksheet's root-level
// `<autoFilter>` and `<sortState>`, what its writer serialises back, and the
// text of the `_xlnm._FilterDatabase` defined name xlsx/writer.ts keeps in
// step with every autofilter.
//
// | Wire | Model |
// | --- | --- |
// | `<autoFilter ref xr:uid>` | `AutoFilter.range`; the whole element → `sourceXml` |
// | `filterColumn@colId` / `@hiddenButton="1"` or `@showButton="0"` | `FilterColumn.colId` / `hideButton` |
// | `filters@blank="1"`, `filter@val`, `dateGroupItem@year/month/day/hour/minute/second/dateTimeGrouping` | `values` (`blanks`, `values`, `dateGroups`) |
// | `customFilters@and="1"`, `customFilter@operator` (default `equal`) `@val` | `custom` (`join` "and", else "or"; one or two `conditions`) |
// | `top10@top` (default 1) `@percent` `@val`; `@filterVal` | `top10`; `filterVal` written (the threshold), ignored on read |
// | `dynamicFilter@type`; `@val`, `@maxVal`, `@valIso` | `dynamic`; `val` written (the mean) for the average types, all three ignored on read |
// | `colorFilter`, `iconFilter` | `preserved`, the element verbatim, label "colour" / "icon" |
// | `autoFilter/sortState`, worksheet `<sortState ref caseSensitive>` | `AutoFilter.sort`, `Sheet.sortState` (`ref`, `caseSensitive`, `keys`, `sourceXml`) |
// | `sortCondition@ref` / `@descending="1"` | `SheetSortKey.col` (its ref's left column) / `descending` |
// | `sheetProtection@sort="0"` / `@autoFilter="0"` | `protection.allowSort` / `allowAutoFilter` (sheetXml.ts) |
// | `sheetPr@filterMode="1"` | written iff some row is `filterHidden` (sheetXml.ts) |
// | `row@hidden="1"` | `hidden` or `filterHidden` — the read rule tells them apart (reader.ts, filter.ts classifyHiddenRows) |
// | `definedName name="_xlnm._FilterDatabase" localSheetId hidden="1"` | derived from each sheet's autofilter (writer.ts) |
//
// Reading. Each element is parsed from its own source slice (sheetXml.ts
// cuts it out of the part by the scanner's offsets, in the sheet's one
// pass), wrapped in an element binding the slice's prefix to the namespace
// the part gave it — so `<x:autoFilter>` in a prefixed part reads like an
// unprefixed one; only SpreadsheetML children count, every other element
// (`extLst`, a revision namespace's) is skipped.
// - An element whose `ref` does not parse is not modelled; the writer keeps
//   such a source element verbatim while the model has none.
// - `columns` ascending by colId, the first `filterColumn` per colId kept; a
//   column whose children hold no criterion reads `{ colId }`. A column's
//   criterion is its first recognised child. `filter` without `val` and a
//   `dateGroupItem` without `year` or with an unknown `dateTimeGrouping` are
//   skipped; a `customFilter` operator outside the six reads `equal`, and
//   only the first two conditions are read (the schema allows no more); a
//   `customFilters` with none is no criterion; `top10@val` absent → 10.
// - An unknown `dynamicFilter@type` (ST_DynamicFilterType's `null`, or one a
//   later Excel adds) is kept as `preserved` with label "icon": like a
//   colour or icon filter it is not evaluated, so the rows keep the
//   filter state the file gave them and the element goes back verbatim; the
//   label only picks the popover's notice, and the model has no third one.
// - A sortState's keys are its `sortCondition`s whose `sortBy` is absent or
//   "value" and that name no `customList`; a `columnSort="1"` state (left to
//   right, never produced) has none. `sourceXml` is always the element.
//
// Writing (sheetXml.ts's serializeWorksheet, the SPEC-122 dirty-sheet
// splice). A `sourceXml` is re-emitted verbatim when it fits the part
// (fitsPart: one SpreadsheetML element of the expected name, every prefix in
// it bound by the part's root — an ODS database range or an element read from
// another part's namespaces never fits); otherwise the element is serialised
// from the model with the root's prefix:
//   <autoFilter ref="A1:D13"> {filterColumn}* [sortState] </autoFilter>
//   <filterColumn colId="n"[ hiddenButton="1"]> criterion </filterColumn>
//   <filters[ blank="1"]><filter val="…"/>…<dateGroupItem year=… [month=…]
//     [day=…] [hour=…] [minute=…] [second=…] dateTimeGrouping="…"/>…</filters>
//   <customFilters[ and="1"]><customFilter[ operator="…"] val="…"/>…</customFilters>
//     (operator omitted for equal)
//   <top10[ top="0"][ percent="1"] val="…"[ filterVal="…"]/> — filterVal the
//     threshold over the column's numeric data cells (filter.ts's top10
//     rule: the value-th largest/smallest, or rank max(1, floor(count ×
//     value / 100)) for percent), omitted when the column holds no number
//   <dynamicFilter type="…"[ val="…"]/> — val the mean of the column's
//     numeric data cells, for aboveAverage/belowAverage only
//   a preserved criterion's xml verbatim (dropped if it does not fit)
//   <sortState ref="A2:D13"[ caseSensitive="1"]><sortCondition[ descending="1"]
//     ref="C2:C13"/>…</sortState> — each key's ref its column over the
//     state's rows; a sort's own sourceXml verbatim while it fits.
// Elements with no children close themselves. Attribute text is escaped.
import { scanAttr, scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { colToLetters, formatRange, parseRange, type GridRange } from "../address";
import type { AutoFilter, DateGroupItem, DynamicFilterType, FilterColumn, FilterCondition, FilterCriteria, FilterOperator } from "../filter";
import { quoteSheetName } from "../formula/print";
import type { Row, Sheet } from "../model";
import type { SheetSortKey, SortState } from "../sort";
import { escapeAttributeValue } from "../xmlSplice";
import { isSpreadsheetMl } from "./styles";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** The reserved name Excel gives each autofilter's range. */
export const FILTER_DATABASE_NAME = "_xlnm._FilterDatabase";

const OPERATORS: ReadonlySet<string> = new Set<FilterOperator>(["equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"]);

const DYNAMIC_TYPES: ReadonlySet<string> = new Set<DynamicFilterType>([
  "aboveAverage",
  "belowAverage",
  "today",
  "yesterday",
  "tomorrow",
  "thisWeek",
  "lastWeek",
  "nextWeek",
  "thisMonth",
  "lastMonth",
  "nextMonth",
  "thisQuarter",
  "lastQuarter",
  "nextQuarter",
  "thisYear",
  "lastYear",
  "nextYear",
  "yearToDate",
  "Q1",
  "Q2",
  "Q3",
  "Q4",
  "M1",
  "M2",
  "M3",
  "M4",
  "M5",
  "M6",
  "M7",
  "M8",
  "M9",
  "M10",
  "M11",
  "M12",
]);

const GROUPINGS: ReadonlySet<string> = new Set<DateGroupItem["grouping"]>(["year", "month", "day", "hour", "minute", "second"]);

function isOn(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function isOff(value: string | undefined): boolean {
  return value === "0" || value === "false";
}

function parseNumber(text: string | undefined): number | undefined {
  if (text === undefined || text.trim() === "") {
    return undefined;
  }
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

function parseIndex(text: string | undefined): number | undefined {
  const value = parseNumber(text);
  return value !== undefined && Number.isInteger(value) && value >= 0 ? value : undefined;
}

function numberText(value: number): string {
  return Object.is(value, -0) ? "0" : String(value);
}

// --- reading ------------------------------------------------------------------------------

/** One element of a slice: offsets into the slice. */
interface SliceElement {
  local: string;
  /** In SpreadsheetML (transitional or strict). */
  main: boolean;
  attrs: readonly XmlScanAttribute[];
  start: number;
  end: number;
  children: SliceElement[];
}

/** The qualified name a slice's root element is written with. */
function rootNameOf(slice: string): string {
  return /^<([^\s/>]+)/.exec(slice)?.[1] ?? "";
}

/** The slice's root element as a small tree — the slice's prefix bound to
 * `ns` — or undefined when the slice is not one well-formed element. */
function sliceTree(slice: string, ns: string): SliceElement | undefined {
  const name = rootNameOf(slice);
  const colon = name.indexOf(":");
  const head = `<_ ${colon < 0 ? "xmlns" : `xmlns:${name.slice(0, colon)}`}="${escapeAttributeValue(ns)}">`;
  const offset = head.length;
  const stack: SliceElement[] = [];
  let root: SliceElement | undefined;
  let depth = 0;
  try {
    scanXml(`${head}${slice}</_>`, (event) => {
      if (event.type === "open") {
        depth += 1;
        if (depth >= 2) {
          const element: SliceElement = {
            local: event.local,
            main: isSpreadsheetMl(event.ns),
            attrs: event.attrs,
            start: event.start - offset,
            end: event.end - offset,
            children: [],
          };
          const parent = stack.at(-1);
          if (parent) {
            parent.children.push(element);
          } else if (root === undefined) {
            root = element;
          }
          if (!event.selfClosing) {
            stack.push(element);
          }
        }
        if (event.selfClosing) {
          depth -= 1;
        }
      } else if (event.type === "close") {
        if (depth >= 2) {
          stack.pop()!.end = event.end - offset;
        }
        depth -= 1;
      }
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return undefined;
    }
    throw error;
  }
  return root;
}

function rangeAttr(element: SliceElement): GridRange | null {
  return parseRange(scanAttr(element.attrs, "ref")?.trim() ?? "");
}

function readDateGroup(element: SliceElement): DateGroupItem | undefined {
  const grouping = scanAttr(element.attrs, "dateTimeGrouping");
  const year = parseIndex(scanAttr(element.attrs, "year"));
  if (grouping === undefined || !GROUPINGS.has(grouping) || year === undefined) {
    return undefined;
  }
  const item: DateGroupItem = { grouping: grouping as DateGroupItem["grouping"], year };
  for (const field of ["month", "day", "hour", "minute", "second"] as const) {
    const value = parseIndex(scanAttr(element.attrs, field));
    if (value !== undefined) {
      item[field] = value;
    }
  }
  return item;
}

/** A `filterColumn` child's criterion (see the module header); undefined
 * for a child that is none. */
function readCriteria(slice: string, element: SliceElement): FilterCriteria | undefined {
  switch (element.local) {
    case "filters": {
      const values: string[] = [];
      const dateGroups: DateGroupItem[] = [];
      for (const child of element.children) {
        if (!child.main) {
          continue;
        }
        if (child.local === "filter") {
          const value = scanAttr(child.attrs, "val");
          if (value !== undefined) {
            values.push(value);
          }
        } else if (child.local === "dateGroupItem") {
          const item = readDateGroup(child);
          if (item) {
            dateGroups.push(item);
          }
        }
      }
      return { kind: "values", values, blanks: isOn(scanAttr(element.attrs, "blank")), dateGroups };
    }
    case "customFilters": {
      const conditions: FilterCondition[] = [];
      for (const child of element.children) {
        if (child.main && child.local === "customFilter" && conditions.length < 2) {
          const operator = scanAttr(child.attrs, "operator");
          conditions.push({
            operator: operator !== undefined && OPERATORS.has(operator) ? (operator as FilterOperator) : "equal",
            value: scanAttr(child.attrs, "val") ?? "",
          });
        }
      }
      if (conditions.length === 0) {
        return undefined;
      }
      return {
        kind: "custom",
        join: isOn(scanAttr(element.attrs, "and")) ? "and" : "or",
        conditions: conditions.length === 1 ? [conditions[0]!] : [conditions[0]!, conditions[1]!],
      };
    }
    case "top10":
      return {
        kind: "top10",
        top: !isOff(scanAttr(element.attrs, "top")),
        percent: isOn(scanAttr(element.attrs, "percent")),
        value: parseNumber(scanAttr(element.attrs, "val")) ?? 10,
      };
    case "dynamicFilter": {
      const type = scanAttr(element.attrs, "type");
      if (type !== undefined && DYNAMIC_TYPES.has(type)) {
        return { kind: "dynamic", type: type as DynamicFilterType };
      }
      return { kind: "preserved", xml: slice.slice(element.start, element.end), label: "icon" };
    }
    case "colorFilter":
      return { kind: "preserved", xml: slice.slice(element.start, element.end), label: "colour" };
    case "iconFilter":
      return { kind: "preserved", xml: slice.slice(element.start, element.end), label: "icon" };
    default:
      return undefined;
  }
}

function readFilterColumn(slice: string, element: SliceElement): FilterColumn | undefined {
  const colId = parseIndex(scanAttr(element.attrs, "colId"));
  if (colId === undefined) {
    return undefined;
  }
  const column: FilterColumn = { colId };
  for (const child of element.children) {
    const criteria = child.main ? readCriteria(slice, child) : undefined;
    if (criteria !== undefined) {
      column.criteria = criteria;
      break;
    }
  }
  if (isOn(scanAttr(element.attrs, "hiddenButton")) || isOff(scanAttr(element.attrs, "showButton"))) {
    column.hideButton = true;
  }
  return column;
}

function readSortStateElement(slice: string, element: SliceElement): SortState | undefined {
  const ref = rangeAttr(element);
  if (ref === null) {
    return undefined;
  }
  const keys: SheetSortKey[] = [];
  if (!isOn(scanAttr(element.attrs, "columnSort"))) {
    for (const child of element.children) {
      if (!child.main || child.local !== "sortCondition") {
        continue;
      }
      const sortBy = scanAttr(child.attrs, "sortBy");
      const at = rangeAttr(child);
      if ((sortBy === undefined || sortBy === "value") && scanAttr(child.attrs, "customList") === undefined && at !== null) {
        keys.push({ col: at.left, descending: isOn(scanAttr(child.attrs, "descending")) });
      }
    }
  }
  const state: SortState = { ref, keys };
  if (isOn(scanAttr(element.attrs, "caseSensitive"))) {
    state.caseSensitive = true;
  }
  state.sourceXml = slice.slice(element.start, element.end);
  return state;
}

/** A worksheet's `<autoFilter>` element (its source text; `ns`, the
 * namespace its prefix is bound to in the part) → the model; undefined when
 * it is not modelled (no parseable `ref`). */
export function readAutoFilterXml(xml: string, ns: string = MAIN_NS): AutoFilter | undefined {
  const root = sliceTree(xml, ns);
  if (root === undefined || !root.main || root.local !== "autoFilter") {
    return undefined;
  }
  const range = rangeAttr(root);
  if (range === null) {
    return undefined;
  }
  const columns: FilterColumn[] = [];
  const seen = new Set<number>();
  let sort: SortState | undefined;
  for (const child of root.children) {
    if (!child.main) {
      continue;
    }
    if (child.local === "filterColumn") {
      const column = readFilterColumn(xml, child);
      if (column !== undefined && !seen.has(column.colId)) {
        seen.add(column.colId);
        columns.push(column);
      }
    } else if (child.local === "sortState" && sort === undefined) {
      sort = readSortStateElement(xml, child);
    }
  }
  const filter: AutoFilter = { range, columns: columns.toSorted((a, b) => a.colId - b.colId) };
  if (sort !== undefined) {
    filter.sort = sort;
  }
  filter.sourceXml = xml;
  return filter;
}

/** A worksheet's root-level `<sortState>` element → the model; undefined
 * when it is not modelled (no parseable `ref`). */
export function readSortStateXml(xml: string, ns: string = MAIN_NS): SortState | undefined {
  const root = sliceTree(xml, ns);
  return root !== undefined && root.main && root.local === "sortState" ? readSortStateElement(xml, root) : undefined;
}

/** The autofilter and worksheet sort record a whole worksheet part holds —
 * its root-level elements, read as the worksheet reader reads them (for a
 * reader upgrade re-reading a dirty sheet's retained source). The cells are
 * cut out unscanned. */
export function readSheetFilterElements(worksheetXml: string): { autoFilter?: AutoFilter; sortState?: SortState } {
  const xml = worksheetXml.replace(/<([\w.-]+:)?sheetData\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/\1sheetData\s*>)/, "");
  const found: { autoFilter?: AutoFilter; sortState?: SortState } = {};
  let depth = 0;
  let capture: { local: "autoFilter" | "sortState"; start: number; ns: string } | null = null;
  const take = (local: "autoFilter" | "sortState", slice: string, ns: string): void => {
    if (local === "autoFilter") {
      const filter = readAutoFilterXml(slice, ns);
      if (filter !== undefined) {
        found.autoFilter = filter;
      }
    } else {
      const state = readSortStateXml(slice, ns);
      if (state !== undefined) {
        found.sortState = state;
      }
    }
  };
  scanXml(xml, (event) => {
    if (event.type === "open") {
      depth += 1;
      if (depth === 2 && isSpreadsheetMl(event.ns) && (event.local === "autoFilter" || event.local === "sortState") && !(event.local in found)) {
        if (event.selfClosing) {
          take(event.local, xml.slice(event.start, event.end), event.ns!);
        } else {
          capture = { local: event.local, start: event.start, ns: event.ns! };
        }
      }
      if (event.selfClosing) {
        depth -= 1;
      }
    } else if (event.type === "close") {
      if (depth === 2 && capture !== null) {
        const open: { local: "autoFilter" | "sortState"; start: number; ns: string } = capture;
        take(open.local, xml.slice(open.start, event.end), open.ns);
        capture = null;
      }
      depth -= 1;
    }
  });
  return found;
}

// --- writing ------------------------------------------------------------------------------

/** Whether a retained element may be re-emitted verbatim in the part: see
 * fitsPart. `local` undefined accepts any SpreadsheetML element. */
export type FitsPart = (xml: string, local?: string) => boolean;

/** Whether `xml` can go verbatim into a part whose root start tag is
 * `rootTag`: exactly one SpreadsheetML element (named `local`, when given)
 * whose every element and attribute prefix that root (or the element
 * itself) binds. */
export function fitsPart(xml: string, rootTag: string, local?: string): boolean {
  const rootName = rootNameOf(rootTag);
  if (rootName === "" || !xml.startsWith("<")) {
    return false;
  }
  const open = rootTag.endsWith("/>") ? `${rootTag.slice(0, -2).trimEnd()}>` : rootTag;
  let depth = 0;
  let elements = 0;
  let fits = true;
  try {
    scanXml(`${open}${xml}</${rootName}>`, (event) => {
      if (event.type === "text") {
        if (depth === 1 && event.text.trim() !== "") {
          fits = false;
        }
        return fits ? undefined : "stop";
      }
      if (event.type === "close") {
        depth -= 1;
        return undefined;
      }
      depth += 1;
      if (depth >= 2) {
        if (depth === 2) {
          elements += 1;
          if (elements > 1 || !isSpreadsheetMl(event.ns) || (local !== undefined && event.local !== local)) {
            fits = false;
          }
        }
        if ((event.name.includes(":") && event.ns === undefined) || event.attrs.some((attr) => attr.name.includes(":") && attr.ns === undefined)) {
          fits = false;
        }
      }
      if (event.selfClosing) {
        depth -= 1;
      }
      return fits ? undefined : "stop";
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return false;
    }
    throw error;
  }
  return fits && elements === 1;
}

/** The numbers among the cells of `col` on rows top..bottom. */
function numericCells(sheet: Pick<Sheet, "rows">, col: number, top: number, bottom: number): number[] {
  const numbers: number[] = [];
  for (const row of sheet.rows) {
    if (row.index < top) {
      continue;
    }
    if (row.index > bottom) {
      break;
    }
    const value = cellIn(row, col)?.value;
    if (typeof value === "number") {
      numbers.push(value);
    }
  }
  return numbers;
}

function cellIn(row: Row, col: number): Row["cells"][number] | undefined {
  const cells = row.cells;
  let lo = 0;
  let hi = cells.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const at = cells[mid]!.col;
    if (at === col) {
      return cells[mid];
    }
    if (at < col) {
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return undefined;
}

/** The top10 threshold over a column's numeric data cells (filter.ts's
 * rule); undefined when the column holds no number. */
export function top10Threshold(numbers: readonly number[], top: boolean, percent: boolean, value: number): number | undefined {
  if (numbers.length === 0) {
    return undefined;
  }
  const ordered = numbers.toSorted((a, b) => (top ? b - a : a - b));
  const rank = percent ? Math.max(1, Math.floor((ordered.length * value) / 100)) : Math.max(1, Math.floor(value));
  return ordered[Math.min(rank, ordered.length) - 1];
}

function attrXml(name: string, value: string): string {
  return ` ${name}="${escapeAttributeValue(value)}"`;
}

function elementXml(p: string, local: string, attrs: string, children: string): string {
  return children === "" ? `<${p}${local}${attrs}/>` : `<${p}${local}${attrs}>${children}</${p}${local}>`;
}

function dateGroupXml(item: DateGroupItem, p: string): string {
  let attrs = attrXml("year", String(item.year));
  for (const field of ["month", "day", "hour", "minute", "second"] as const) {
    const value = item[field];
    if (value !== undefined) {
      attrs += attr(field, String(value));
    }
  }
  return `<${p}dateGroupItem${attrs}${attrXml("dateTimeGrouping", item.grouping)}/>`;
}

function criteriaXml(criteria: FilterCriteria, col: number, filter: AutoFilter, sheet: Pick<Sheet, "rows">, p: string, fits: FitsPart): string {
  const data = (): number[] => numericCells(sheet, col, filter.range.top + 1, filter.range.bottom);
  switch (criteria.kind) {
    case "values":
      return elementXml(
        p,
        "filters",
        criteria.blanks ? ` blank="1"` : "",
        criteria.values.map((value) => `<${p}filter${attrXml("val", value)}/>`).join("") + criteria.dateGroups.map((item) => dateGroupXml(item, p)).join(""),
      );
    case "custom":
      return elementXml(
        p,
        "customFilters",
        criteria.join === "and" ? ` and="1"` : "",
        criteria.conditions.map((condition) => `<${p}customFilter${condition.operator === "equal" ? "" : attrXml("operator", condition.operator)}${attrXml("val", condition.value)}/>`).join(""),
      );
    case "top10": {
      const threshold = top10Threshold(data(), criteria.top, criteria.percent, criteria.value);
      return (
        `<${p}top10${criteria.top ? "" : ` top="0"`}${criteria.percent ? ` percent="1"` : ""}${attrXml("val", numberText(criteria.value))}` +
        `${threshold === undefined ? "" : attrXml("filterVal", numberText(threshold))}/>`
      );
    }
    case "dynamic": {
      let val = "";
      if (criteria.type === "aboveAverage" || criteria.type === "belowAverage") {
        const numbers = data();
        if (numbers.length > 0) {
          val = attrXml("val", numberText(numbers.reduce((sum, n) => sum + n, 0) / numbers.length));
        }
      }
      return `<${p}dynamicFilter${attrXml("type", criteria.type)}${val}/>`;
    }
    case "preserved":
      return fits(criteria.xml) ? criteria.xml : "";
  }
}

function filterColumnXml(column: FilterColumn, filter: AutoFilter, sheet: Pick<Sheet, "rows">, p: string, fits: FitsPart): string {
  const body = column.criteria === undefined ? "" : criteriaXml(column.criteria, filter.range.left + column.colId, filter, sheet, p, fits);
  return elementXml(p, "filterColumn", `${attrXml("colId", String(column.colId))}${column.hideButton ? ` hiddenButton="1"` : ""}`, body);
}

/** A `<sortState>` for the model's record (see the module header): its
 * sourceXml while it fits the part, else serialised. */
export function sortStateXml(state: SortState, p: string, fits: FitsPart): string {
  if (state.sourceXml !== undefined && fits(state.sourceXml, "sortState")) {
    return state.sourceXml;
  }
  const conditions = state.keys
    .map((key) => {
      const ref = formatRange({ top: state.ref.top, left: key.col, bottom: state.ref.bottom, right: key.col });
      return `<${p}sortCondition${key.descending ? ` descending="1"` : ""}${attrXml("ref", ref)}/>`;
    })
    .join("");
  return elementXml(p, "sortState", `${attrXml("ref", formatRange(state.ref))}${state.caseSensitive ? ` caseSensitive="1"` : ""}`, conditions);
}

/** An `<autoFilter>` for the model's autofilter (see the module header):
 * its sourceXml while it fits the part, else serialised over `sheet`'s
 * cells (top10's filterVal, the averages' val). */
export function autoFilterXml(filter: AutoFilter, sheet: Pick<Sheet, "rows">, p: string, fits: FitsPart): string {
  if (filter.sourceXml !== undefined && fits(filter.sourceXml, "autoFilter")) {
    return filter.sourceXml;
  }
  const columns = filter.columns.map((column) => filterColumnXml(column, filter, sheet, p, fits)).join("");
  const sort = filter.sort === undefined ? "" : sortStateXml(filter.sort, p, fits);
  return elementXml(p, "autoFilter", attrXml("ref", formatRange(filter.range)), columns + sort);
}

// --- _xlnm._FilterDatabase ---------------------------------------------------------------------

function absoluteCell(row: number, col: number): string {
  return `$${colToLetters(col)}$${row + 1}`;
}

/** The `_xlnm._FilterDatabase` text for an autofilter: the sheet's name as a
 * reference quotes it, then the range with absolute rows and columns —
 * `'My Sheet'!$A$1:$D$20`. */
export function filterDatabaseText(sheetName: string, range: GridRange): string {
  return `${quoteSheetName(sheetName)}!${absoluteCell(range.top, range.left)}:${absoluteCell(range.bottom, range.right)}`;
}

/** Whether a `_xlnm._FilterDatabase` text (decoded) already names `range` on
 * `sheetName` — however it is written (quoted or not, `$` or not, one cell or
 * two): such a name is left as written. */
export function filterDatabaseNames(text: string, sheetName: string, range: GridRange): boolean {
  const bang = text.lastIndexOf("!");
  if (bang <= 0) {
    return false;
  }
  let sheet = text.slice(0, bang).trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replaceAll("''", "'");
  }
  if (sheet.toUpperCase() !== sheetName.toUpperCase()) {
    return false;
  }
  const at = parseRange(text.slice(bang + 1).trim());
  return at !== null && at.top === range.top && at.left === range.left && at.bottom === range.bottom && at.right === range.right;
}
