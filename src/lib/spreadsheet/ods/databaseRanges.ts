// The ODS wire mapping of a sheet's autofilter (SPEC-128 "ODS wire mapping")
// — THIS FILE is the normative source for which `table:database-range` is a
// sheet's autofilter, how its `table:filter` and `table:sort` map onto the
// model, and how the writer regenerates `table:database-ranges`.
//
// - Where. `table:database-ranges` sits in the content envelope's suffix
//   (after the last table, after `table:named-expressions`). Elements are
//   matched by their qualified names as written (`table:`), as
//   sheetNamePass.ts does; each `table:database-range` is scanned on its own.
// - The claim (odsAutoFilterClaims). A `table:database-range` with
//   `table:display-filter-buttons="true"`, a `table:target-range-address` of
//   ONE cell range on ONE sheet (`Sheet1.A1:Sheet1.D20`, quoted names, `$`
//   ignored — rangeAddress.ts's grammar; whole columns or rows, 3-D and lists
//   are no autofilter) naming a source table, and well-formed XML is that
//   table's autofilter — the anonymous `__Anonymous_Sheet_DB__<n>` first (the
//   number is informative: LibreOffice assigns an anonymous range to the
//   sheet its target lies on), else the first named one. The claim depends
//   on the source text and the source tables' names only, so the reader, the
//   writer and SPEC-124's retained-reference pass (which leaves a claimed
//   element alone — the model's autofilter adjuster moves it) agree on it
//   whatever the model has become.
// - Read (readOdsAutoFilters). `range` = the target; `sourceXml` = the whole
//   element verbatim; `odsName` = its `table:name`. The `table:filter`'s one
//   child — a `table:filter-and`, or a lone `table:filter-or` /
//   `table:filter-condition` — is split per `table:field-number` (offset from
//   `range.left` = colId, inside the range):
//
//   | `table:operator` | Model |
//   | --- | --- |
//   | `=` with `table:filter-set-item`s; a `table:filter-or` of `=` / `empty` on one field, every `=` text without `*` `?` `~`; a text `=` `""` | `values` (an item or value `""`, or `empty`, → `blanks`; `dateGroups` []) |
//   | `=` `!=` `<` `>` `<=` `>=` | `custom` equal, notEqual, lessThan, greaterThan, lessThanOrEqual, greaterThanOrEqual — the value verbatim (`table:data-type="number"`: invariant number text; text: wildcard text, `*` `?` `~` as LibreOffice's wildcard mode reads them) |
//   | two conditions on one field inside the `table:filter-and` / a `table:filter-or` of two on one field | one `custom`, join "and" / "or" |
//   | `begins-with` / `ends-with` / `contains` x | `equal` `"x*"` / `"*x"` / `"*x*"` (`*` `?` `~` in x escaped with `~`) |
//   | `does-not-begin-with` / `does-not-end-with` / `does-not-contain` x | `notEqual` with the same patterns |
//   | `empty` / `!empty` | `values` blanks only / `custom notEqual " "` |
//   | `top values` / `bottom values` / `top percent` / `bottom percent` n | `top10` (n > 0) |
//   | `match`, `!match`, a colour condition (`loext:` or another data type), any other attribute, a field outside the range, an OR across fields, more than two conditions on a field, nested `table:filter-and`, a `table:filter` with a copy-output or criteria-range source or `table:display-duplicates="false"` | the whole `table:filter` verbatim → `odsUnmappedFilter`, no per-column criteria |
//
//   `table:sort` (`table:case-sensitive="true"` → caseSensitive) with
//   `table:sort-by table:field-number table:order` → `autoFilter.sort`: `ref`
//   the range's data rows (the range less its header row), key col =
//   `range.left + field-number`, descending when `table:order="descending"`;
//   a sort-by whose `table:data-type` is a user list (not automatic, number,
//   alphanumeric or text) or whose field lies outside the range is no key;
//   the element verbatim → `sort.sourceXml`. Rows are not classified: the
//   table states `table:visibility="filter"` itself (tableXml.ts).
// - Write, substitution path (odsSuffixWithAutoFilters, whenever a sheet is
//   dirty or `structureDirty`). Inside the source's `table:database-ranges`,
//   each claimed element is replaced in place by its owner's autofilter —
//   the owner being the ORIGINAL sheet whose source name the claim names —
//   or removed when the owner (or its autofilter) is gone; every element the
//   model does not own is re-emitted through `retained` (the writer's
//   sheet-reference pass; ranges on deleted sheets dropped); an autofilter
//   with no claim (a new one, a copy's — whose `odsName` duplicateSheet
//   dropped) is appended at the end, LibreOffice's place for sheet-local
//   ranges; a container emptied that way is removed, and a missing one is
//   created after `table:named-expressions`, before `table:data-pilot-tables`
//   (`table:consolidation`, `table:dde-links`, else `</office:spreadsheet>`).
//   The owner's element:
//   - while `autoFilter.sourceXml` is the claimed element (unchanged since
//     read), that text verbatim through the owner's sheet pass, but for its
//     `table:target-range-address`, re-serialised when the sheet's name is
//     no longer the one it names, and an anonymous `table:name` renumbered
//     to the sheet's index when the sheet moved;
//   - otherwise the claimed element patched: `table:name` (anonymous →
//     `__Anonymous_Sheet_DB__<sheet index>`, a named one keeps its name),
//     `table:target-range-address`, `table:display-filter-buttons="true"`,
//     the `table:filter` and `table:sort` regenerated after any
//     `table:database-source-*` child; every other attribute and child kept
//     through the sheet pass.
//   A sheet with no claim gets `<table:database-range table:name=
//   "__Anonymous_Sheet_DB__<sheet index>" table:target-range-address=…
//   table:display-filter-buttons="true">`. Generated addresses are written
//   as LibreOffice does, the sheet on both ends (`Sheet1.A1:Sheet1.D20`;
//   the sheet text from canonicalToOdsRangeAddress).
// - Write, from scratch (odsDatabaseRangesXml): every sheet's autofilter
//   generated, the container after the tables.
// - The generated `table:filter`: `odsUnmappedFilter` verbatim when present;
//   else one element per criterion in ascending colId — inside a
//   `table:filter-and` when there are two or more, alone otherwise: `values`
//   as `<table:filter-condition table:field-number="f" table:value="<first>"
//   table:operator="=">` holding one `table:filter-set-item` per value, then
//   `table:value=""` for blanks; `custom` conditions — `notEqual " "` as
//   `!empty`, a pure `"x*"` / `"*x"` / `"*x*"` pattern as begins-with /
//   ends-with / contains (does-not-… for notEqual), else the operator with
//   the value verbatim and `table:data-type="number"` when it is an
//   invariant number — two of them as two conditions in the and, or one
//   `table:filter-or`; `top10` as `top values` / `bottom values` / `top
//   percent` / `bottom percent` with `table:data-type="number"`. With no ODS
//   form, left out: a `preserved` (xlsx colour or icon) criterion, a
//   `dynamic` one, the `dateGroups` of a `values` (a groups-only one
//   entirely) and an empty `values` (nothing checked). `table:sort`:
//   `sort.sourceXml` verbatim while it is a `table:sort`, else
//   `<table:sort [table:case-sensitive="true"]>` with one `<table:sort-by
//   table:field-number table:data-type="automatic" table:order>` per key
//   inside the range; none for no keys.
// - Known non-identity: a `custom` or of two plain-text equals reads back as
//   the `values` of the two (ODS has one encoding for both).
import { scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { colToLetters, type GridRange } from "../address";
import type { AutoFilter, FilterColumn, FilterCondition, FilterCriteria, FilterOperator } from "../filter";
import { formatReference, sheetNameNeedsQuotes } from "../formula/print";
import { isOriginal, MAX_COLS, MAX_ROWS, type Sheet } from "../model";
import type { SheetSortKey, SortState } from "../sort";
import { attributeValue, escapeAttributeValue, setAttribute, startTagAt } from "../xmlSplice";
import { tableNameOf } from "./package";
import { canonicalToOdsRangeAddress, parseOdsRangeAddress } from "./rangeAddress";

/** LibreOffice's name prefix for a sheet's anonymous database range. */
export const ODS_ANONYMOUS_SHEET_DB = "__Anonymous_Sheet_DB__";

/** The anonymous database range name of the sheet at `index`. */
export function anonymousSheetDbName(index: number): string {
  return `${ODS_ANONYMOUS_SHEET_DB}${index}`;
}

const isAnonymous = (name: string | undefined): boolean => name !== undefined && name.startsWith(ODS_ANONYMOUS_SHEET_DB);

// --- locating the elements ---------------------------------------------------------------

/** One `table:database-range` of the suffix. */
export interface OdsDatabaseRange {
  /** [start, end) in the scanned text. */
  start: number;
  end: number;
  xml: string;
  name?: string;
  /** The target's sheet as written (unquoted) and its area, when the target
   * is one cell range on one sheet. */
  sheet?: string;
  range?: GridRange;
  filterButtons: boolean;
}

/** The `table:database-ranges` element of a suffix and its ranges. */
export interface OdsDatabaseRanges {
  container?: {
    start: number;
    end: number;
    /** Where its content starts and its close tag starts (both `end` when
     * self-closing). */
    contentStart: number;
    contentEnd: number;
    selfClosing: boolean;
  };
  ranges: OdsDatabaseRange[];
}

const CONTAINER = /<table:database-ranges(?=[\s/>])[^>]*?(?:\/>|>[\s\S]*?<\/table:database-ranges>)/;
const RANGE = /<table:database-range(?=[\s/>])[^>]*?(?:\/>|>[\s\S]*?<\/table:database-range>)/g;

/** A target address → its sheet and area; undefined unless it is one cell
 * range (or cell) on one sheet. */
function targetOf(address: string | undefined): { sheet: string; range: GridRange } | undefined {
  const ref = address === undefined ? null : parseOdsRangeAddress(address.trim());
  if (!ref || ref.sheet === undefined || ref.sheetEnd !== undefined) {
    return undefined;
  }
  const end = ref.end ?? ref.start;
  const { row: r1, col: c1 } = ref.start;
  const { row: r2, col: c2 } = end;
  if (r1 === undefined || c1 === undefined || r2 === undefined || c2 === undefined) {
    return undefined;
  }
  return { sheet: ref.sheet, range: { top: Math.min(r1, r2), left: Math.min(c1, c2), bottom: Math.max(r1, r2), right: Math.max(c1, c2) } };
}

/** The suffix's `table:database-ranges` element and every range in it. */
export function odsDatabaseRanges(suffix: string): OdsDatabaseRanges {
  const match = CONTAINER.exec(suffix);
  if (!match) {
    return { ranges: [] };
  }
  const start = match.index;
  const end = start + match[0].length;
  const tag = startTagAt(suffix, start);
  const selfClosing = tag.endsWith("/>");
  const contentStart = selfClosing ? end : start + tag.length;
  const contentEnd = selfClosing ? end : end - "</table:database-ranges>".length;
  const ranges: OdsDatabaseRange[] = [];
  const content = suffix.slice(contentStart, contentEnd);
  for (const found of content.matchAll(RANGE)) {
    const xml = found[0];
    const rangeTag = startTagAt(xml, 0);
    const range: OdsDatabaseRange = {
      start: contentStart + found.index,
      end: contentStart + found.index + xml.length,
      xml,
      filterButtons: attributeValue(rangeTag, "table:display-filter-buttons") === "true",
    };
    const name = attributeValue(rangeTag, "table:name");
    if (name !== undefined) {
      range.name = name;
    }
    const target = targetOf(attributeValue(rangeTag, "table:target-range-address"));
    if (target) {
      range.sheet = target.sheet;
      range.range = target.range;
    }
    ranges.push(range);
  }
  return { container: { start, end, contentStart, contentEnd, selfClosing }, ranges };
}

// --- a small element tree ----------------------------------------------------------------

/** One element of a scanned fragment (qualified names as written). */
interface Node {
  name: string;
  attrs: readonly XmlScanAttribute[];
  children: Node[];
  start: number;
  end: number;
}

/** The fragment's root element, or null when it does not scan. */
function parseTree(xml: string): Node | null {
  const stack: Node[] = [];
  let root: Node | null = null;
  try {
    scanXml(xml, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "close") {
        const node = stack.pop();
        if (node) {
          node.end = event.end;
        }
        return undefined;
      }
      const node: Node = { name: event.name, attrs: event.attrs, children: [], start: event.start, end: event.end };
      const parent = stack.at(-1);
      if (parent) {
        parent.children.push(node);
      } else if (root === null) {
        root = node;
      }
      if (!event.selfClosing) {
        stack.push(node);
      }
      return undefined;
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return null;
    }
    throw error;
  }
  return root;
}

const attr = (node: Node, qname: string): string | undefined => node.attrs.find((each) => each.name === qname)?.value;

// --- the claim -----------------------------------------------------------------------------

/** A table's autofilter as the source holds it. */
export interface OdsAutoFilterClaim {
  /** The claimed table's index among `tableNames`. */
  tableIndex: number;
  element: OdsDatabaseRange & { range: GridRange };
  tree: Node;
}

/** Each source table's claimed autofilter element (see the module header),
 * in document order. `tableNames` are the source tables' names, in order. */
export function odsAutoFilterClaims(found: OdsDatabaseRanges, tableNames: readonly string[]): OdsAutoFilterClaim[] {
  const indexByName = new Map<string, number>();
  tableNames.forEach((name, index) => {
    const key = name.toUpperCase();
    if (!indexByName.has(key)) {
      indexByName.set(key, index);
    }
  });
  const claims = new Map<number, OdsAutoFilterClaim>();
  for (const element of found.ranges) {
    if (!element.filterButtons || element.sheet === undefined || element.range === undefined) {
      continue;
    }
    const tableIndex = indexByName.get(element.sheet.toUpperCase());
    if (tableIndex === undefined) {
      continue;
    }
    const held = claims.get(tableIndex);
    if (held && (isAnonymous(held.element.name) || !isAnonymous(element.name))) {
      continue;
    }
    const tree = parseTree(element.xml);
    if (tree) {
      claims.set(tableIndex, { tableIndex, element: { ...element, range: element.range }, tree });
    }
  }
  return [...claims.values()].toSorted((a, b) => a.element.start - b.element.start);
}

/** The [start, end) spans of the suffix's claimed elements — the database
 * ranges the model owns, which SPEC-124's retained-reference pass leaves
 * alone. */
export function odsOwnedDatabaseRangeSpans(suffix: string, tableNames: readonly string[]): { start: number; end: number }[] {
  if (!suffix.includes("table:display-filter-buttons")) {
    return [];
  }
  return odsAutoFilterClaims(odsDatabaseRanges(suffix), tableNames).map((claim) => ({ start: claim.element.start, end: claim.element.end }));
}

// --- reading ---------------------------------------------------------------------------------

const FILTER = "table:filter";
const FILTER_AND = "table:filter-and";
const FILTER_OR = "table:filter-or";
const CONDITION = "table:filter-condition";
const SET_ITEM = "table:filter-set-item";
const SORT = "table:sort";
const SORT_BY = "table:sort-by";

const COMPARISONS: Readonly<Record<string, FilterOperator>> = {
  "=": "equal",
  "!=": "notEqual",
  "<": "lessThan",
  ">": "greaterThan",
  "<=": "lessThanOrEqual",
  ">=": "greaterThanOrEqual",
};
const PATTERNS: Readonly<Record<string, { operator: FilterOperator; wrap: (x: string) => string }>> = {
  "begins-with": { operator: "equal", wrap: (x) => `${x}*` },
  "ends-with": { operator: "equal", wrap: (x) => `*${x}` },
  contains: { operator: "equal", wrap: (x) => `*${x}*` },
  "does-not-begin-with": { operator: "notEqual", wrap: (x) => `${x}*` },
  "does-not-end-with": { operator: "notEqual", wrap: (x) => `*${x}` },
  "does-not-contain": { operator: "notEqual", wrap: (x) => `*${x}*` },
};
const TOP10: Readonly<Record<string, { top: boolean; percent: boolean }>> = {
  "top values": { top: true, percent: false },
  "bottom values": { top: false, percent: false },
  "top percent": { top: true, percent: true },
  "bottom percent": { top: false, percent: true },
};

/** The attributes a mapped condition may carry. */
const CONDITION_ATTRIBUTES: ReadonlySet<string> = new Set(["table:field-number", "table:value", "table:operator", "table:data-type", "table:case-sensitive"]);

/** `*`, `?` and `~` escaped with `~` (a literal inside a wildcard pattern). */
function escapeWildcards(text: string): string {
  return text.replace(/[~*?]/g, "~$&");
}

/** One condition, read: its field and what it maps to; null when it does
 * not map (the whole filter is then kept verbatim). */
type ReadCondition =
  | { field: number; kind: "values"; values: string[]; blanks: boolean }
  /** `plain`: a text equality without wildcards (a `table:filter-or` of
   * them is a values list). */
  | { field: number; kind: "custom"; condition: FilterCondition; plain?: true }
  | { field: number; kind: "top10"; criteria: FilterCriteria };

function readCondition(node: Node, width: number): ReadCondition | null {
  if (node.name !== CONDITION || node.attrs.some((each) => !CONDITION_ATTRIBUTES.has(each.name))) {
    return null;
  }
  const fieldText = attr(node, "table:field-number") ?? "";
  const field = Number(fieldText);
  if (!/^\d+$/.test(fieldText) || field >= width) {
    return null;
  }
  const dataType = attr(node, "table:data-type");
  if (dataType !== undefined && dataType !== "number" && dataType !== "text") {
    return null;
  }
  const operator = attr(node, "table:operator") ?? "=";
  const value = attr(node, "table:value") ?? "";
  if (node.children.length > 0) {
    if (operator !== "=" || node.children.some((child) => child.name !== SET_ITEM || child.attrs.some((each) => each.name !== "table:value") || child.children.length > 0)) {
      return null;
    }
    const items = node.children.map((child) => attr(child, "table:value") ?? "");
    return { field, kind: "values", values: items.filter((item) => item !== ""), blanks: items.includes("") };
  }
  const text = dataType !== "number";
  if (operator === "=" && text && value === "") {
    return { field, kind: "values", values: [], blanks: true };
  }
  const comparison = COMPARISONS[operator];
  if (comparison !== undefined) {
    const read: ReadCondition = { field, kind: "custom", condition: { operator: comparison, value } };
    if (operator === "=" && text && !/[~*?]/.test(value)) {
      read.plain = true;
    }
    return read;
  }
  const pattern = PATTERNS[operator];
  if (pattern !== undefined) {
    return { field, kind: "custom", condition: { operator: pattern.operator, value: pattern.wrap(escapeWildcards(value)) } };
  }
  if (operator === "empty") {
    return { field, kind: "values", values: [], blanks: true };
  }
  if (operator === "!empty") {
    return { field, kind: "custom", condition: { operator: "notEqual", value: " " } };
  }
  const top10 = TOP10[operator];
  const n = Number(value);
  if (top10 !== undefined && value.trim() !== "" && Number.isFinite(n) && n > 0) {
    return { field, kind: "top10", criteria: { kind: "top10", top: top10.top, percent: top10.percent, value: n } };
  }
  return null;
}

/** A field's conditions → its criterion, or null. */
function fieldCriteria(items: readonly ({ or: ReadCondition[] } | ReadCondition)[]): FilterCriteria | null {
  if (items.length === 1) {
    const item = items[0]!;
    if ("or" in item) {
      return orCriteria(item.or);
    }
    switch (item.kind) {
      case "values":
        return { kind: "values", values: item.values, blanks: item.blanks, dateGroups: [] };
      case "custom":
        return { kind: "custom", join: "and", conditions: [item.condition] };
      case "top10":
        return item.criteria;
    }
  }
  if (items.length === 2 && items.every((item) => !("or" in item) && item.kind === "custom")) {
    const [a, b] = items as [ReadCondition & { kind: "custom" }, ReadCondition & { kind: "custom" }];
    return { kind: "custom", join: "and", conditions: [a.condition, b.condition] };
  }
  return null;
}

/** A `table:filter-or` on one field: `values` when every condition is a
 * plain equality or empty, else a custom or of two. */
function orCriteria(conditions: readonly ReadCondition[]): FilterCriteria | null {
  if (conditions.length === 1) {
    return fieldCriteria(conditions);
  }
  if (conditions.every((each) => each.kind === "values" || (each.kind === "custom" && each.plain))) {
    const values: string[] = [];
    let blanks = false;
    for (const each of conditions) {
      if (each.kind === "values") {
        values.push(...each.values);
        blanks ||= each.blanks;
      } else if (each.kind === "custom") {
        values.push(each.condition.value);
      }
    }
    return { kind: "values", values, blanks, dateGroups: [] };
  }
  if (conditions.length === 2 && conditions.every((each) => each.kind === "custom")) {
    const [a, b] = conditions as [ReadCondition & { kind: "custom" }, ReadCondition & { kind: "custom" }];
    return { kind: "custom", join: "or", conditions: [a.condition, b.condition] };
  }
  return null;
}

/** The `table:filter` attributes a mapped filter may carry. */
function filterMaps(filter: Node): boolean {
  return filter.attrs.every((each) => each.name === "table:display-duplicates" && each.value === "true");
}

/** A `table:filter` → per-column criteria, or null when it does not
 * decompose (see the module header). */
function readFilter(filter: Node, width: number): FilterColumn[] | null {
  if (!filterMaps(filter) || filter.children.length > 1) {
    return null;
  }
  const top = filter.children[0];
  if (!top) {
    return [];
  }
  const fields = new Map<number, ({ or: ReadCondition[] } | ReadCondition)[]>();
  const add = (field: number, item: { or: ReadCondition[] } | ReadCondition): void => {
    fields.set(field, [...(fields.get(field) ?? []), item]);
  };
  /** A `table:filter-or` whose conditions are all on one field. */
  const readOr = (node: Node): boolean => {
    if (node.attrs.length > 0 || node.children.length === 0) {
      return false;
    }
    const conditions = node.children.map((child) => readCondition(child, width));
    if (conditions.some((each) => each === null) || new Set(conditions.map((each) => each!.field)).size !== 1) {
      return false;
    }
    add(conditions[0]!.field, { or: conditions as ReadCondition[] });
    return true;
  };
  if (top.name === FILTER_AND) {
    if (top.attrs.length > 0) {
      return null;
    }
    for (const child of top.children) {
      if (child.name === FILTER_OR) {
        if (!readOr(child)) {
          return null;
        }
        continue;
      }
      const condition = readCondition(child, width);
      if (!condition) {
        return null;
      }
      add(condition.field, condition);
    }
  } else if (top.name === FILTER_OR) {
    if (!readOr(top)) {
      return null;
    }
  } else {
    const condition = readCondition(top, width);
    if (!condition) {
      return null;
    }
    add(condition.field, condition);
  }
  const columns: FilterColumn[] = [];
  for (const [colId, items] of [...fields].toSorted((a, b) => a[0] - b[0])) {
    const criteria = fieldCriteria(items);
    if (!criteria) {
      return null;
    }
    columns.push({ colId, criteria });
  }
  return columns;
}

/** The data types a sort key may have (anything else names a user list). */
const KEY_DATA_TYPES: ReadonlySet<string> = new Set(["automatic", "number", "alphanumeric", "text"]);

/** The range's data rows (the range less its header row). */
function dataRows(range: GridRange): GridRange {
  return { top: Math.min(range.top + 1, MAX_ROWS - 1), left: range.left, bottom: Math.max(range.bottom, Math.min(range.top + 1, MAX_ROWS - 1)), right: range.right };
}

function readSort(sort: Node, xml: string, range: GridRange): SortState {
  const keys: SheetSortKey[] = [];
  for (const by of sort.children) {
    const fieldText = attr(by, "table:field-number") ?? "";
    const col = range.left + Number(fieldText);
    if (by.name !== SORT_BY || !/^\d+$/.test(fieldText) || col > range.right || !KEY_DATA_TYPES.has(attr(by, "table:data-type") ?? "automatic")) {
      continue;
    }
    keys.push({ col, descending: attr(by, "table:order") === "descending" });
  }
  const state: SortState = { ref: dataRows(range), keys, sourceXml: xml.slice(sort.start, sort.end) };
  if (attr(sort, "table:case-sensitive") === "true") {
    state.caseSensitive = true;
  }
  return state;
}

/** A claimed element → the sheet's autofilter (see the module header). */
export function readOdsAutoFilter(claim: OdsAutoFilterClaim): AutoFilter {
  const { element, tree } = claim;
  const range = element.range;
  const autoFilter: AutoFilter = { range: { ...range }, columns: [], sourceXml: element.xml };
  if (element.name !== undefined) {
    autoFilter.odsName = element.name;
  }
  const filter = tree.children.find((child) => child.name === FILTER);
  if (filter) {
    const columns = readFilter(filter, range.right - range.left + 1);
    if (columns) {
      autoFilter.columns = columns;
    } else {
      autoFilter.odsUnmappedFilter = element.xml.slice(filter.start, filter.end);
    }
  }
  const sort = tree.children.find((child) => child.name === SORT);
  if (sort) {
    autoFilter.sort = readSort(sort, element.xml, range);
  }
  return autoFilter;
}

/** The source tables' autofilters by table index (`tableNames` in order). */
export function readOdsAutoFilters(suffix: string, tableNames: readonly string[]): Map<number, AutoFilter> {
  const filters = new Map<number, AutoFilter>();
  if (!suffix.includes("table:display-filter-buttons")) {
    return filters;
  }
  for (const claim of odsAutoFilterClaims(odsDatabaseRanges(suffix), tableNames)) {
    filters.set(claim.tableIndex, readOdsAutoFilter(claim));
  }
  return filters;
}

// --- writing ---------------------------------------------------------------------------------

const OPERATOR_TEXT: Readonly<Record<FilterOperator, string>> = {
  equal: "=",
  notEqual: "!=",
  lessThan: "<",
  greaterThan: ">",
  lessThanOrEqual: "<=",
  greaterThanOrEqual: ">=",
};

const INVARIANT_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

const isStar = (at: { ch: string; wild: boolean } | undefined): boolean => at !== undefined && at.wild && at.ch === "*";

/** A wildcard pattern as a literal with a `*` at its start, end or both —
 * undefined for any other pattern (a `?`, an inner `*`, no wildcard). */
function purePattern(value: string): { literal: string; start: boolean; end: boolean } | undefined {
  const chars: { ch: string; wild: boolean }[] = [];
  for (let i = 0; i < value.length; i += 1) {
    const ch = value[i]!;
    if (ch === "~" && i + 1 < value.length) {
      i += 1;
      chars.push({ ch: value[i]!, wild: false });
    } else {
      chars.push({ ch, wild: ch === "*" || ch === "?" });
    }
  }
  const start = isStar(chars[0]);
  const end = chars.length > 1 && isStar(chars.at(-1));
  const inner = chars.slice(start ? 1 : 0, end ? chars.length - 1 : chars.length);
  if ((!start && !end) || inner.length === 0 || inner.some((each) => each.wild)) {
    return undefined;
  }
  return { literal: inner.map((each) => each.ch).join(""), start, end };
}

/** A pure pattern's operator, by where its `*`s are, for equal and notEqual. */
const PATTERN_OPERATORS = {
  contains: { equal: "contains", notEqual: "does-not-contain" },
  begins: { equal: "begins-with", notEqual: "does-not-begin-with" },
  ends: { equal: "ends-with", notEqual: "does-not-end-with" },
} as const;

function conditionXml(field: number, condition: FilterCondition): string {
  const head = `<table:filter-condition table:field-number="${field}"`;
  if (condition.operator === "notEqual" && condition.value === " ") {
    return `${head} table:value="" table:operator="!empty"/>`;
  }
  if (condition.operator === "equal" || condition.operator === "notEqual") {
    const pattern = purePattern(condition.value);
    if (pattern) {
      const where = pattern.start && pattern.end ? "contains" : pattern.end ? "begins" : "ends";
      return `${head} table:value="${escapeAttributeValue(pattern.literal)}" table:operator="${PATTERN_OPERATORS[where][condition.operator]}"/>`;
    }
  }
  const numeric = INVARIANT_NUMBER.test(condition.value) && Number.isFinite(Number(condition.value));
  return `${head} table:value="${escapeAttributeValue(condition.value)}" table:operator="${escapeAttributeValue(OPERATOR_TEXT[condition.operator])}"${numeric ? ` table:data-type="number"` : ""}/>`;
}

/** One criterion's elements inside the filter (none when it has no ODS form). */
function criteriaXml(field: number, criteria: FilterCriteria): string[] {
  switch (criteria.kind) {
    case "values": {
      const items = [...criteria.values, ...(criteria.blanks ? [""] : [])];
      if (items.length === 0) {
        return [];
      }
      const set = items.map((item) => `<table:filter-set-item table:value="${escapeAttributeValue(item)}"/>`).join("");
      return [`<table:filter-condition table:field-number="${field}" table:value="${escapeAttributeValue(items[0]!)}" table:operator="=">${set}</table:filter-condition>`];
    }
    case "custom": {
      const conditions = criteria.conditions.map((condition) => conditionXml(field, condition));
      return criteria.join === "or" && conditions.length > 1 ? [`<table:filter-or>${conditions.join("")}</table:filter-or>`] : conditions;
    }
    case "top10": {
      const operator = `${criteria.top ? "top" : "bottom"} ${criteria.percent ? "percent" : "values"}`;
      return [`<table:filter-condition table:field-number="${field}" table:value="${criteria.value}" table:operator="${operator}" table:data-type="number"/>`];
    }
    case "dynamic":
    case "preserved":
      return [];
  }
}

/** A sheet's retained ODS pieces as the writer passes them. */
type Pass = (text: string) => string;
const identity: Pass = (text) => text;

/** The autofilter's `table:filter` (see the module header); "" for none. */
export function odsFilterXml(autoFilter: AutoFilter, pass: Pass = identity): string {
  if (autoFilter.odsUnmappedFilter !== undefined) {
    return pass(autoFilter.odsUnmappedFilter);
  }
  const width = autoFilter.range.right - autoFilter.range.left + 1;
  const parts = autoFilter.columns
    .filter((column) => column.criteria !== undefined && column.colId >= 0 && column.colId < width)
    .toSorted((a, b) => a.colId - b.colId)
    .flatMap((column) => criteriaXml(column.colId, column.criteria!));
  if (parts.length === 0) {
    return "";
  }
  return `<table:filter>${parts.length === 1 ? parts[0] : `<table:filter-and>${parts.join("")}</table:filter-and>`}</table:filter>`;
}

/** The autofilter's `table:sort`; "" for none. */
export function odsSortXml(autoFilter: AutoFilter, pass: Pass = identity): string {
  const sort = autoFilter.sort;
  if (!sort) {
    return "";
  }
  if (sort.sourceXml !== undefined && /^<table:sort(?=[\s/>])/.test(sort.sourceXml)) {
    return pass(sort.sourceXml);
  }
  const { left, right } = autoFilter.range;
  const keys = sort.keys.filter((key) => key.col >= left && key.col <= right);
  if (keys.length === 0) {
    return "";
  }
  const by = keys
    .map((key) => `<table:sort-by table:field-number="${key.col - left}" table:data-type="automatic" table:order="${key.descending ? "descending" : "ascending"}"/>`)
    .join("");
  return `<table:sort${sort.caseSensitive ? ` table:case-sensitive="true"` : ""}>${by}</table:sort>`;
}

/** A sheet name as an ODS address writes it, without the `$`. */
function odsSheetText(name: string): string {
  const single = canonicalToOdsRangeAddress(formatReference({ sheet: name, start: { row: 0, col: 0, rowAbs: false, colAbs: false } }));
  if (single !== null && single.startsWith("$") && single.endsWith(".A1")) {
    return single.slice(1, -3);
  }
  return sheetNameNeedsQuotes(name) || name.includes(".") ? `'${name.replaceAll("'", "''")}'` : name;
}

/** `Sheet1.A1:Sheet1.D20`, as LibreOffice writes a database range's target. */
export function odsTargetRangeAddress(sheetName: string, range: GridRange): string {
  const sheet = odsSheetText(sheetName);
  const cell = (row: number, col: number): string => `${sheet}.${colToLetters(Math.min(col, MAX_COLS - 1))}${Math.min(row, MAX_ROWS - 1) + 1}`;
  return `${cell(range.top, range.left)}:${cell(range.bottom, range.right)}`;
}

/** A new `table:database-range` for a sheet's autofilter. */
export function odsDatabaseRangeXml(autoFilter: AutoFilter, sheetName: string, name: string, pass: Pass = identity): string {
  const body = odsFilterXml(autoFilter, pass) + odsSortXml(autoFilter, pass);
  const head =
    `<table:database-range table:name="${escapeAttributeValue(name)}" table:target-range-address="${escapeAttributeValue(odsTargetRangeAddress(sheetName, autoFilter.range))}" ` +
    `table:display-filter-buttons="true"`;
  return body === "" ? `${head}/>` : `${head}>${body}</table:database-range>`;
}

/** A claimed element rewritten for its owner's changed autofilter (see the
 * module header): its other attributes and children kept through `pass`. */
function patchedElement(claim: OdsAutoFilterClaim, autoFilter: AutoFilter, sheetName: string, index: number, pass: Pass): string {
  const { element, tree } = claim;
  const xml = element.xml;
  let tag = pass(startTagAt(xml, 0));
  if (isAnonymous(element.name) || element.name === undefined) {
    tag = setAttribute(tag, "table:name", anonymousSheetDbName(index));
  }
  tag = setAttribute(tag, "table:target-range-address", odsTargetRangeAddress(sheetName, autoFilter.range));
  tag = setAttribute(tag, "table:display-filter-buttons", "true");
  const kept = tree.children.filter((child) => child.name !== FILTER && child.name !== SORT);
  const sources = kept.filter((child) => child.name.startsWith("table:database-source"));
  const rest = kept.filter((child) => !child.name.startsWith("table:database-source"));
  const slice = (nodes: readonly Node[]): string => nodes.map((node) => pass(xml.slice(node.start, node.end))).join("");
  const body = slice(sources) + odsFilterXml(autoFilter, pass) + odsSortXml(autoFilter, pass) + slice(rest);
  const open = tag.replace(/\s*\/>$/, ">");
  const qname = /^<([^\s/>]+)/.exec(tag)?.[1] ?? "table:database-range";
  return body === "" ? open.replace(/>$/, "/>") : `${open}${body}</${qname}>`;
}

/** The owner's element for a claim (see the module header). */
function ownedElement(claim: OdsAutoFilterClaim, autoFilter: AutoFilter, sheetName: string, index: number, pass: Pass): string {
  if (autoFilter.sourceXml !== claim.element.xml) {
    return patchedElement(claim, autoFilter, sheetName, index, pass);
  }
  const text = pass(claim.element.xml);
  const original = startTagAt(text, 0);
  let tag = original;
  if (claim.element.sheet !== sheetName) {
    tag = setAttribute(tag, "table:target-range-address", odsTargetRangeAddress(sheetName, autoFilter.range));
  }
  if (isAnonymous(claim.element.name)) {
    tag = setAttribute(tag, "table:name", anonymousSheetDbName(index));
  }
  return tag === original ? text : tag + text.slice(original.length);
}

export interface OdsAutoFilterWriteOptions {
  /** The source tables' names in source order (SheetRemap.orderBefore). */
  sourceNames: readonly string[];
  /** The writer's pass over retained text the model does not own (the
   * sheet-reference pass, ranges on deleted sheets removed). */
  retained: Pass;
  /** The pass over one sheet's own retained pieces (its copy's remap). */
  sheetPass: (sheet: Sheet, text: string) => string;
}

/** Where a new `table:database-ranges` goes in a suffix. */
function insertionPoint(suffix: string): number {
  const at = /<table:(?:data-pilot-tables|consolidation|dde-links)(?=[\s/>])|<\/office:spreadsheet>/.exec(suffix);
  return at ? at.index : 0;
}

/** The substitution path's suffix (see the module header) and the text it
 * generated (for the root's namespace declarations). */
export function odsSuffixWithAutoFilters(suffix: string, sheets: readonly Sheet[], options: OdsAutoFilterWriteOptions): { suffix: string; generated: string } {
  const found = odsDatabaseRanges(suffix);
  const claims = found.ranges.some((range) => range.filterButtons) ? odsAutoFilterClaims(found, options.sourceNames) : [];
  const sourceIndex = new Map<string, number>();
  options.sourceNames.forEach((name, index) => {
    const key = name.toUpperCase();
    if (!sourceIndex.has(key)) {
      sourceIndex.set(key, index);
    }
  });
  // Each claim's owner: the original holding that source table.
  const owners = new Map<number, { sheet: Sheet; index: number }>();
  sheets.forEach((sheet, index) => {
    const name = isOriginal(sheet) && sheet.sourceXml !== undefined ? tableNameOf(sheet.sourceXml) : undefined;
    const at = name === undefined ? undefined : sourceIndex.get(name.toUpperCase());
    if (at !== undefined && !owners.has(at)) {
      owners.set(at, { sheet, index });
    }
  });
  let generated = "";
  const replaced = new Map<OdsAutoFilterClaim, string>();
  const placed = new Set<Sheet>();
  for (const claim of claims) {
    const owner = owners.get(claim.tableIndex);
    const autoFilter = owner?.sheet.autoFilter;
    if (!owner || !autoFilter) {
      replaced.set(claim, "");
      continue;
    }
    const text = ownedElement(claim, autoFilter, owner.sheet.name, owner.index, (piece) => options.sheetPass(owner.sheet, piece));
    if (text !== claim.element.xml) {
      generated += text;
    }
    replaced.set(claim, text);
    placed.add(owner.sheet);
  }
  let fresh = "";
  sheets.forEach((sheet, index) => {
    if (sheet.autoFilter && !placed.has(sheet)) {
      fresh += odsDatabaseRangeXml(sheet.autoFilter, sheet.name, anonymousSheetDbName(index), (piece) => options.sheetPass(sheet, piece));
    }
  });
  generated += fresh;
  const container = found.container;
  if (!container) {
    if (fresh === "") {
      return { suffix: options.retained(suffix), generated };
    }
    const at = insertionPoint(suffix);
    return {
      suffix: `${options.retained(suffix.slice(0, at))}<table:database-ranges>${fresh}</table:database-ranges>${options.retained(suffix.slice(at))}`,
      generated,
    };
  }
  // The container's content: claimed elements replaced, the rest passed.
  let content = "";
  let at = container.contentStart;
  for (const claim of claims) {
    content += options.retained(suffix.slice(at, claim.element.start)) + replaced.get(claim)!;
    at = claim.element.end;
  }
  content += options.retained(suffix.slice(at, container.contentEnd)) + fresh;
  const tag = startTagAt(suffix, container.start);
  let element: string;
  if (found.ranges.length > 0 && !/<table:database-range(?=[\s/>])/.test(content)) {
    element = ""; // emptied
  } else if (container.selfClosing) {
    element = content === "" ? tag : `${tag.replace(/\s*\/>$/, ">")}${content}</table:database-ranges>`;
  } else {
    element = `${tag}${content}</table:database-ranges>`;
  }
  return {
    suffix: options.retained(suffix.slice(0, container.start)) + element + options.retained(suffix.slice(container.end)),
    generated,
  };
}

/** From scratch: every sheet's autofilter in one `table:database-ranges`
 * ("" when no sheet has one). */
export function odsDatabaseRangesXml(sheets: readonly Sheet[]): string {
  const ranges = sheets.map((sheet, index) => (sheet.autoFilter ? odsDatabaseRangeXml(sheet.autoFilter, sheet.name, anonymousSheetDbName(index)) : "")).join("");
  return ranges === "" ? "" : `<table:database-ranges>${ranges}</table:database-ranges>`;
}
