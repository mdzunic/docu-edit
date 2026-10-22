// The xlsx wire form of Excel tables (SPEC-133 "Tables on the wire";
// tables.ts is the model's normative source, tableGolden.ts the new part's).
//
// Reading: the worksheet's `<tableParts>` → its `…/relationships/table` →
// `xl/tables/tableN.xml` → SheetTable: `displayName` (fallback `name`) →
// `name`; `ref`; `headerRowCount` "0" → no header (absent → header);
// `totalsRowCount` ≥ 1 → totals; `tableStyleInfo@name`/`show*` ("1"/"true");
// `tableColumn` `name` (`_xHHHH_` decoded), `id` → `xlsxId`,
// `totalsRowFunction` (none → absent), `totalsRowLabel`, and the children
// `calculatedColumnFormula`/`totalsRowFormula` through fromXlsxFormula +
// structuredFromXlsx (one that does not parse stays out of the model and in
// the part); the table's `autoFilter` verbatim as `autoFilter.sourceXml`
// (present = filter buttons); `tableType` queryTable/xml → `locked`;
// `table@id` → `xlsxId`; `partName`. A part whose root or `ref` does not read
// is not modelled — it stays in the package with its relationship.
//
// Writing (writer.ts through planTableWrites, a save only):
// - A clean table whose part still says what the model says → its part
//   verbatim; a `dirty` one (or one that no longer agrees) → regenerated
//   (regenerateTablePart): the source root start tag with `id`, `name`,
//   `displayName`, `ref`, `headerRowCount`, `totalsRowCount`, `totalsRowShown`
//   rewritten; `autoFilter` = the part's own (kept current by SPEC-124's
//   retained `sortState`/`sortCondition` rows) or the model's `sourceXml`,
//   with `ref` patched to header + data and each `filterColumn@colId` moved
//   to its column's new position (by `xlsxId`; a removed column's goes) —
//   omitted without a header row or filter buttons, `<autoFilter ref/>` when
//   there is no source; `sortState` and `extLst` verbatim; each `tableColumn`
//   keeps its source start tag (`dataDxfId`, `xr3:uid`, … ) and unmodelled
//   children by `xlsxId`, a formula child its source start tag (`array="1"`).
// - A table without a part (new, a copy's, one whose part another table
//   claimed) → the golden form at the next free `xl/tables/table{n}.xml` with
//   a workbook-unique `id` (max + 1), the sheet relationship and the
//   Override. A table moved from another sheet keeps its part and gains this
//   sheet's relationship.
// - Each written grid sheet whose `tables` are known (or which is dirty —
//   absent then means none): a table relationship the sheet no longer needs
//   (deleted, moved away, or a part SPEC-121's closure copy duplicated for a
//   copy that none of the copy's tables claims — sheets.ts gives each copied
//   table its closure-copied part) is removed — the part, its `.rels`
//   and Override then fall to the writer's reachability rule — and
//   `<tableParts>` is rewritten (in CT_Worksheet order, before `extLst`)
//   whenever its relationship list changed. A part the reader could not
//   model is always kept, listed after the model's tables.
import { relsPathFor } from "../../core/xml";
import { formatRange, parseRange, type GridRange } from "../address";
import { fromXlsxFormula, toXlsxFormula } from "../formula/dialect";
import { FormulaSyntaxError, parseFormulaAt } from "../formula/parse";
import type { SheetTable, SheetTableColumn, TotalsFunction } from "../model";
import { structuredFromXlsx, structuredToXlsx } from "../tables";
import { applySplices, escapeAttributeValue, escapeElementText, setAttribute, type Splice } from "../xmlSplice";
import { worksheetTail } from "./drawingXml";
import { elementTree, nodeAttr, type XmlNode } from "./elementTree";
import { nextFreePartName, nextRelationshipId, partTextOf, readRels, relationshipKind, relationshipXml, relativeTarget, spliceRels, targetPart, type PartLookup, type RelEntry } from "./package";
import { decodeXmlEscapes, encodeXlsxAttribute } from "./sharedStrings";

export const TABLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const REL_NAMESPACES = new Set([REL_NS, STRICT_REL_NS]);
const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

const TOTALS_FUNCTIONS: ReadonlySet<string> = new Set<TotalsFunction>(["sum", "min", "max", "average", "count", "countNums", "stdDev", "var", "custom"]);
const isTrue = (value: string | undefined): boolean => value === "1" || value === "true";
const child = (node: XmlNode, local: string): XmlNode | undefined => node.children.find((each) => each.local === local);

// --- reading -----------------------------------------------------------------------------

/** A formula child's text as the model stores it; undefined when it does
 * not parse (the part keeps it). */
function readTableFormula(text: string): string | undefined {
  try {
    const canonical = structuredFromXlsx(fromXlsxFormula(text.trim()));
    return parseFormulaAt(canonical, { row: 0, col: 0 }) instanceof FormulaSyntaxError ? undefined : canonical;
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return undefined;
    }
    throw error;
  }
}

/** The root `table` element of a part. */
function tableRoot(xml: string): XmlNode | undefined {
  return elementTree(xml)?.find((node) => node.local === "table");
}

/** A table part as the model reads it (no `id`, `partName`); null when it
 * is not modelled. */
export function readTablePart(xml: string): Omit<SheetTable, "id"> | null {
  const root = tableRoot(xml);
  const ref = root === undefined ? null : parseRange(nodeAttr(root, "ref")?.trim() ?? "");
  if (root === undefined || ref === null) {
    return null;
  }
  const columns: SheetTableColumn[] = (child(root, "tableColumns")?.children ?? [])
    .filter((node) => node.local === "tableColumn")
    .map((node) => {
      const column: SheetTableColumn = { name: decodeXmlEscapes(nodeAttr(node, "name") ?? "") };
      const id = Number(nodeAttr(node, "id"));
      if (Number.isInteger(id) && id > 0) {
        column.xlsxId = id;
      }
      const fn = nodeAttr(node, "totalsRowFunction");
      if (fn !== undefined && TOTALS_FUNCTIONS.has(fn)) {
        column.totalsFunction = fn as TotalsFunction;
      }
      const label = nodeAttr(node, "totalsRowLabel");
      if (label !== undefined) {
        column.totalsLabel = decodeXmlEscapes(label);
      }
      const calculated = child(node, "calculatedColumnFormula");
      const calculatedFormula = calculated === undefined ? undefined : readTableFormula(calculated.text);
      if (calculatedFormula !== undefined) {
        column.calculatedFormula = calculatedFormula;
      }
      const totals = child(node, "totalsRowFormula");
      const totalsFormula = totals === undefined ? undefined : readTableFormula(totals.text);
      if (totalsFormula !== undefined) {
        column.totalsFormula = totalsFormula;
      }
      return column;
    });
  const styleInfo = child(root, "tableStyleInfo");
  const styleName = styleInfo === undefined ? undefined : nodeAttr(styleInfo, "name");
  const flag = (name: string): boolean => styleInfo !== undefined && isTrue(nodeAttr(styleInfo, name));
  const table: Omit<SheetTable, "id"> = {
    name: nodeAttr(root, "displayName") ?? nodeAttr(root, "name") ?? "",
    ref,
    headerRow: nodeAttr(root, "headerRowCount")?.trim() !== "0",
    totalsRow: Number(nodeAttr(root, "totalsRowCount") ?? 0) > 0,
    columns,
    style: {
      ...(styleName === undefined ? {} : { name: styleName }),
      showFirstColumn: flag("showFirstColumn"),
      showLastColumn: flag("showLastColumn"),
      showRowStripes: flag("showRowStripes"),
      showColumnStripes: flag("showColumnStripes"),
    },
  };
  const filter = child(root, "autoFilter");
  if (filter !== undefined) {
    table.autoFilter = { sourceXml: xml.slice(filter.start, filter.end) };
  }
  const type = nodeAttr(root, "tableType");
  if (type === "queryTable" || type === "xml") {
    table.locked = true;
  }
  const id = Number(nodeAttr(root, "id"));
  if (Number.isInteger(id) && id > 0) {
    table.xlsxId = id;
  }
  return table;
}

/** A `.rels` part's relationships; none for an absent or unreadable one. */
function relsIn(text: string | undefined): RelEntry[] {
  if (text === undefined || text.trim() === "") {
    return [];
  }
  try {
    return readRels(text);
  } catch {
    return [];
  }
}

/** The relationship ids a worksheet's `<tableParts>` lists, in order;
 * undefined when it has none. */
export function worksheetTablePartIds(xml: string): string[] | undefined {
  const tail = worksheetTail(xml);
  const element = tail.children.find((each) => each.main && each.local === "tableParts");
  if (!element) {
    return undefined;
  }
  const node = elementTree(xml.slice(element.start, element.end), [tail.rootTag])?.[0];
  return (node?.children ?? [])
    .filter((each) => each.local === "tablePart")
    .flatMap((each) => each.attrs.find((attr) => attr.local === "id" && attr.ns !== undefined && REL_NAMESPACES.has(attr.ns))?.value ?? []);
}

/** A worksheet's tables (see the module header); `mint` names each. */
export function readXlsxSheetTables(parts: PartLookup, sheetPartName: string, sheetXml: string, mint: () => string): SheetTable[] {
  if (!sheetXml.includes("tablePart")) {
    return [];
  }
  const ids = worksheetTablePartIds(sheetXml) ?? [];
  const rels = new Map(relsIn(partTextOf(parts, relsPathFor(sheetPartName))).map((rel) => [rel.id, rel] as const));
  const tables: SheetTable[] = [];
  for (const rId of ids) {
    const rel = rels.get(rId);
    const partName = rel && relationshipKind(rel.type) === "table" ? targetPart(sheetPartName, rel) : undefined;
    const xml = partName === undefined ? undefined : partTextOf(parts, partName);
    const read = xml === undefined ? null : readTablePart(xml);
    if (read) {
      tables.push({ id: mint(), ...read, partName: partName! });
    }
  }
  return tables;
}

// --- writing: the part ---------------------------------------------------------------------

const formulaText = (canonical: string): string => escapeElementText(toXlsxFormula(structuredToXlsx(canonical)));

/** The rows the filter covers: header + data. */
function filterRange(table: SheetTable): GridRange {
  return { ...table.ref, bottom: table.ref.bottom - (table.totalsRow ? 1 : 0) };
}

const onOff = (value: boolean): string => (value ? "1" : "0");

function styleInfoAttrs(table: SheetTable): string {
  const { style } = table;
  return (
    (style.name === undefined ? "" : ` name="${escapeAttributeValue(style.name)}"`) +
    ` showFirstColumn="${onOff(style.showFirstColumn)}" showLastColumn="${onOff(style.showLastColumn)}"` +
    ` showRowStripes="${onOff(style.showRowStripes)}" showColumnStripes="${onOff(style.showColumnStripes)}"`
  );
}

/** A column's ids: its own when unique and positive, else 1 + the max. */
function columnIds(table: SheetTable): number[] {
  const used = new Set<number>();
  let max = Math.max(0, ...table.columns.map((column) => column.xlsxId ?? 0));
  return table.columns.map((column) => {
    let id = column.xlsxId;
    if (id === undefined || id <= 0 || used.has(id)) {
      max += 1;
      id = max;
    }
    used.add(id);
    return id;
  });
}

/** The modelled attributes of a new `tableColumn`. */
function columnAttrs(column: SheetTableColumn, id: number): string {
  return (
    ` id="${id}" name="${encodeXlsxAttribute(column.name)}"` +
    (column.totalsFunction === undefined ? "" : ` totalsRowFunction="${column.totalsFunction}"`) +
    (column.totalsLabel === undefined ? "" : ` totalsRowLabel="${encodeXlsxAttribute(column.totalsLabel)}"`)
  );
}

/** A new `tableColumn`. */
function newColumnXml(column: SheetTableColumn, id: number, p: string): string {
  const children =
    (column.calculatedFormula === undefined ? "" : `<${p}calculatedColumnFormula>${formulaText(column.calculatedFormula)}</${p}calculatedColumnFormula>`) +
    (column.totalsFunction === "custom" && column.totalsFormula !== undefined ? `<${p}totalsRowFormula>${formulaText(column.totalsFormula)}</${p}totalsRowFormula>` : "");
  return children === "" ? `<${p}tableColumn${columnAttrs(column, id)}/>` : `<${p}tableColumn${columnAttrs(column, id)}>${children}</${p}tableColumn>`;
}

/** A new part for `table` (the golden form, tableGolden.ts). */
export function newTablePartXml(table: SheetTable, id: number): string {
  const ids = columnIds(table);
  const root =
    `<table xmlns="${MAIN_NS}" id="${id}" name="${escapeAttributeValue(table.name)}" displayName="${escapeAttributeValue(table.name)}" ref="${formatRange(table.ref)}"` +
    (table.headerRow ? "" : ` headerRowCount="0"`) +
    (table.totalsRow ? ` totalsRowCount="1"` : "") +
    `>`;
  const filter = table.headerRow && table.autoFilter !== undefined ? newFilterXml(table) : "";
  return (
    XML_DECLARATION +
    root +
    filter +
    `<tableColumns count="${table.columns.length}">${table.columns.map((column, index) => newColumnXml(column, ids[index]!, "")).join("")}</tableColumns>` +
    `<tableStyleInfo${styleInfoAttrs(table)}/></table>`
  );
}

/** A new part's filter: the model's source when it uses no prefix (a copy's
 * criteria), else a plain one. */
function newFilterXml(table: SheetTable): string {
  const source = table.autoFilter?.sourceXml;
  if (source !== undefined && !/<\/?[\w.-]+:|\s[\w.-]+:[\w.-]+\s*=/.test(source)) {
    const patched = patchFilter(source, table, undefined);
    if (patched !== undefined) {
      return patched;
    }
  }
  return `<autoFilter ref="${formatRange(filterRange(table))}"/>`;
}

/** A filter element's text with `ref` patched and its filterColumns moved
 * to their columns' new positions (`sourceIds[i]`: the source column i's
 * id; absent: positions kept). */
function patchFilter(text: string, table: SheetTable, sourceIds: readonly (number | undefined)[] | undefined): string | undefined {
  const node = elementTree(text, [`<table xmlns="${MAIN_NS}">`])?.[0] ?? elementTree(text)?.[0];
  if (node === undefined || node.local !== "autoFilter") {
    return undefined;
  }
  const splices: Splice[] = [{ start: 0, end: node.tagEnd, text: setAttribute(text.slice(0, node.tagEnd), "ref", formatRange(filterRange(table))) }];
  if (sourceIds !== undefined) {
    const position = new Map(table.columns.flatMap((column, index) => (column.xlsxId === undefined ? [] : [[column.xlsxId, index] as const])));
    for (const filterColumn of node.children.filter((each) => each.local === "filterColumn")) {
      const colId = Number(nodeAttr(filterColumn, "colId"));
      const id = Number.isInteger(colId) ? sourceIds[colId] : undefined;
      const next = id === undefined ? undefined : position.get(id);
      if (next === undefined) {
        splices.push({ start: filterColumn.start, end: filterColumn.end, text: "" });
      } else if (next !== colId) {
        const tag = text.slice(filterColumn.start, filterColumn.tagEnd);
        splices.push({ start: filterColumn.start, end: filterColumn.tagEnd, text: setAttribute(tag, "colId", String(next)) });
      }
    }
  }
  return applySplices(text, splices);
}

/** A source `tableColumn` brought in line with the model column. */
function patchedColumnXml(xml: string, source: XmlNode, column: SheetTableColumn, id: number, p: string): string {
  let tag = xml.slice(source.start, source.tagEnd);
  tag = setAttribute(tag, "id", String(id));
  tag = setAttribute(tag, "name", column.name, encodeXlsxAttribute);
  tag = setAttribute(tag, "totalsRowFunction", column.totalsFunction ?? null);
  tag = setAttribute(tag, "totalsRowLabel", column.totalsLabel ?? null, encodeXlsxAttribute);
  const formulaChild = (local: string, canonical: string | undefined): string => {
    const own = child(source, local);
    if (canonical === undefined) {
      // Kept while it is one the reader could not model.
      return own !== undefined && readTableFormula(own.text) === undefined ? xml.slice(own.start, own.end) : "";
    }
    if (own !== undefined && readTableFormula(own.text) === canonical) {
      return xml.slice(own.start, own.end);
    }
    const open = own === undefined ? `<${p}${local}>` : xml.slice(own.start, own.tagEnd).replace(/\s*\/>$/, ">");
    const name = /^<([^\s/>]+)/.exec(open)![1];
    return `${open}${formulaText(canonical)}</${name}>`;
  };
  const children =
    formulaChild("calculatedColumnFormula", column.calculatedFormula) +
    formulaChild("totalsRowFormula", column.totalsFunction === "custom" ? column.totalsFormula : undefined) +
    source.children
      .filter((each) => each.local !== "calculatedColumnFormula" && each.local !== "totalsRowFormula")
      .map((each) => xml.slice(each.start, each.end))
      .join("");
  const name = /^<([^\s/>]+)/.exec(tag)![1];
  return children === "" ? tag.replace(/\s*\/?>$/, "/>") : `${tag.replace(/\s*\/>$/, ">")}${children}</${name}>`;
}

/** A retained table part regenerated for `table` (see the module header). */
export function regenerateTablePart(xml: string, table: SheetTable, id: number): string {
  const root = tableRoot(xml);
  if (root === undefined) {
    return newTablePartXml(table, id);
  }
  const p = root.name.includes(":") ? `${root.name.slice(0, root.name.indexOf(":"))}:` : "";
  let tag = xml.slice(root.start, root.tagEnd);
  tag = setAttribute(tag, "id", String(id));
  tag = setAttribute(tag, "name", table.name);
  tag = setAttribute(tag, "displayName", table.name);
  tag = setAttribute(tag, "ref", formatRange(table.ref));
  tag = setAttribute(tag, "headerRowCount", table.headerRow ? (/headerRowCount\s*=\s*["']1["']/.test(tag) ? "1" : null) : "0");
  tag = setAttribute(tag, "totalsRowCount", table.totalsRow ? "1" : null);
  if (table.totalsRow) {
    tag = setAttribute(tag, "totalsRowShown", null);
  }
  const splices: Splice[] = [{ start: root.start, end: root.tagEnd, text: tag }];
  const sourceColumns = child(root, "tableColumns");
  const sourceColumnNodes = (sourceColumns?.children ?? []).filter((node) => node.local === "tableColumn");
  const sourceIds = sourceColumnNodes.map((node) => {
    const value = Number(nodeAttr(node, "id"));
    return Number.isInteger(value) ? value : undefined;
  });
  // autoFilter
  const filter = child(root, "autoFilter");
  let filterText = "";
  if (table.headerRow && table.autoFilter !== undefined) {
    const base = filter !== undefined ? xml.slice(filter.start, filter.end) : table.autoFilter.sourceXml;
    filterText = (base === undefined ? undefined : patchFilter(base, table, filter !== undefined ? sourceIds : undefined)) ?? `<${p}autoFilter ref="${formatRange(filterRange(table))}"/>`;
  }
  if (filter !== undefined) {
    splices.push({ start: filter.start, end: filter.end, text: filterText });
  } else if (filterText !== "") {
    const before = root.children.find((node) => node.local !== "autoFilter");
    const at = before?.start ?? root.end - `</${root.name}>`.length;
    splices.push({ start: at, end: at, text: filterText });
  }
  // tableColumns
  const ids = columnIds(table);
  const claimed = new Set<XmlNode>();
  const columnsXml = table.columns
    .map((column, index) => {
      const source = sourceColumnNodes.find((node) => !claimed.has(node) && column.xlsxId !== undefined && Number(nodeAttr(node, "id")) === column.xlsxId);
      if (source === undefined) {
        return newColumnXml(column, ids[index]!, p);
      }
      claimed.add(source);
      return patchedColumnXml(xml, source, column, ids[index]!, p);
    })
    .join("");
  if (sourceColumns !== undefined) {
    const open = setAttribute(xml.slice(sourceColumns.start, sourceColumns.tagEnd).replace(/\s*\/>$/, ">"), "count", String(table.columns.length));
    splices.push({ start: sourceColumns.start, end: sourceColumns.end, text: `${open}${columnsXml}</${sourceColumns.name}>` });
  }
  // tableStyleInfo
  const styleInfo = child(root, "tableStyleInfo");
  if (styleInfo !== undefined) {
    let styleTag = xml.slice(styleInfo.start, styleInfo.tagEnd);
    styleTag = setAttribute(styleTag, "name", table.style.name ?? null);
    for (const [name, value] of [
      ["showFirstColumn", table.style.showFirstColumn],
      ["showLastColumn", table.style.showLastColumn],
      ["showRowStripes", table.style.showRowStripes],
      ["showColumnStripes", table.style.showColumnStripes],
    ] as const) {
      if (isTrue(nodeAttr(styleInfo, name)) !== value) {
        styleTag = setAttribute(styleTag, name, value ? "1" : "0");
      }
    }
    splices.push({ start: styleInfo.start, end: styleInfo.tagEnd, text: styleTag });
  } else if (sourceColumns !== undefined) {
    splices.push({ start: sourceColumns.end, end: sourceColumns.end, text: `<${p}tableStyleInfo${styleInfoAttrs(table)}/>` });
  }
  return applySplices(xml, splices);
}

// --- writing: the worksheet ----------------------------------------------------------------

/** A worksheet part with `<tableParts>` listing `rIds` — replaced, inserted
 * at its CT_Worksheet place (before `extLst`) or, for none, removed. */
export function spliceWorksheetTableParts(xml: string, rIds: readonly string[]): string {
  const tail = worksheetTail(xml);
  const existing = tail.children.find((each) => each.main && each.local === "tableParts");
  if (rIds.length === 0) {
    return existing ? xml.slice(0, existing.start) + xml.slice(existing.end) : xml;
  }
  let relPrefix: string | undefined;
  for (const match of tail.rootTag.matchAll(/\sxmlns:([\w.-]+)\s*=\s*["']([^"']*)["']/g)) {
    if (REL_NAMESPACES.has(match[2]!)) {
      relPrefix = `${match[1]}:`;
    }
  }
  const p = tail.prefix;
  const declare = relPrefix === undefined ? ` xmlns:r="${REL_NS}"` : "";
  const element = `<${p}tableParts count="${rIds.length}"${declare}>${rIds.map((rId) => `<${p}tablePart ${relPrefix ?? "r:"}id="${escapeAttributeValue(rId)}"/>`).join("")}</${p}tableParts>`;
  if (existing) {
    return xml.slice(0, existing.start) + element + xml.slice(existing.end);
  }
  const at = tail.children.find((each) => each.main && each.local === "extLst")?.start ?? tail.rootClose;
  return xml.slice(0, at) + element + xml.slice(at);
}

// --- writing: the save's plan ----------------------------------------------------------------

/** One sheet as the save lays it out. */
export interface TableSheetPlan {
  sheet: { id: string; kind: string; dirty?: true; tables?: SheetTable[] };
  partName: string;
  /** The sheet part comes from the retained package. */
  retained: boolean;
}

/** What a save's tables change in the package. */
export interface TableWritePlan {
  /** Table parts written (regenerated and new). */
  texts: Map<string, string>;
  /** Sheet `.rels` texts after the save (retained or new). */
  rels: Map<string, string>;
  /** Each processed sheet part's relationship targets after the save. */
  edges: Map<string, string[]>;
  /** Worksheet parts whose `<tableParts>` is rewritten, with its ids. */
  tableParts: Map<string, string[]>;
  /** New table parts (their Overrides). */
  added: string[];
}

function relationshipBase(rels: readonly RelEntry[]): string {
  return rels.some((rel) => rel.type.startsWith(`${STRICT_REL_NS}/`)) ? STRICT_REL_NS : REL_NS;
}

/** The save's table writes (see the module header). `relsText(sheetPart)`
 * is the sheet's `.rels` as the save has it so far (the drawing layer's
 * changes included); `taken` the package's part names, which new parts join. */
export function planTableWrites(parts: PartLookup, plans: readonly TableSheetPlan[], relsText: (sheetPart: string) => string | undefined, taken: Set<string>): TableWritePlan {
  const out: TableWritePlan = { texts: new Map(), rels: new Map(), edges: new Map(), tableParts: new Map(), added: [] };
  const readCache = new Map<string, Omit<SheetTable, "id"> | null>();
  const read = (partName: string): Omit<SheetTable, "id"> | null => {
    if (!readCache.has(partName)) {
      const xml = partTextOf(parts, partName);
      readCache.set(partName, xml === undefined ? null : readTablePart(xml));
    }
    return readCache.get(partName)!;
  };
  const processed = plans.filter((plan) => plan.sheet.kind === "grid" && (plan.sheet.tables !== undefined || plan.sheet.dirty === true));
  const relsOf = new Map(plans.map((plan) => [plan.partName, relsIn(relsText(plan.partName))] as const));
  const tableRels = (plan: TableSheetPlan): RelEntry[] => (relsOf.get(plan.partName) ?? []).filter((rel) => relationshipKind(rel.type) === "table" && !rel.external);
  // Parts another (unprocessed) sheet still relates.
  const foreign = new Set(plans.filter((plan) => !processed.includes(plan)).flatMap((plan) => tableRels(plan).flatMap((rel) => targetPart(plan.partName, rel) ?? [])));
  // Claims: a sheet's own related parts first, then a part moved in.
  const claim = new Map<SheetTable, string>();
  const claimedParts = new Set<string>(foreign);
  for (const plan of processed) {
    const own = new Set(tableRels(plan).flatMap((rel) => targetPart(plan.partName, rel) ?? []));
    for (const table of plan.sheet.tables ?? []) {
      if (table.partName !== undefined && own.has(table.partName) && !claimedParts.has(table.partName)) {
        claim.set(table, table.partName);
        claimedParts.add(table.partName);
      }
    }
  }
  for (const plan of processed) {
    for (const table of plan.sheet.tables ?? []) {
      if (!claim.has(table) && table.partName !== undefined && parts.has(table.partName) && !claimedParts.has(table.partName) && read(table.partName) !== null) {
        claim.set(table, table.partName);
        claimedParts.add(table.partName);
      }
    }
  }
  // Table ids in use: every part no model table claims (counting a part
  // about to go only skips a number), then each claimed table's own.
  const usedIds = new Set<number>();
  const claimedNames = new Set(claim.values());
  for (const name of parts.keys()) {
    if (/^xl\/tables\/[^/]+\.xml$/.test(name) && !claimedNames.has(name)) {
      const id = partTableId(partTextOf(parts, name) ?? "");
      if (id !== undefined) {
        usedIds.add(id);
      }
    }
  }
  const idFor = new Map<SheetTable, number>();
  for (const [table, partName] of claim) {
    const id = table.xlsxId ?? partTableId(partTextOf(parts, partName) ?? "");
    if (id !== undefined && !usedIds.has(id)) {
      usedIds.add(id);
      idFor.set(table, id);
    }
  }
  const mintId = (table: SheetTable): number => {
    let id = table.xlsxId;
    if (id === undefined || id <= 0 || usedIds.has(id)) {
      id = Math.max(0, ...usedIds) + 1;
    }
    usedIds.add(id);
    return id;
  };
  for (const plan of processed) {
    const sheetPart = plan.partName;
    const relsPath = relsPathFor(sheetPart);
    const rels = relsOf.get(sheetPart) ?? [];
    const current = relsText(sheetPart);
    const remove = new Set<string>();
    const keptUnmodelled: string[] = [];
    const relIdOf = new Map<string, string>();
    for (const rel of tableRels(plan)) {
      const target = targetPart(sheetPart, rel);
      if (target === undefined) {
        continue;
      }
      const mine = (plan.sheet.tables ?? []).some((table) => claim.get(table) === target);
      if (mine && !relIdOf.has(target)) {
        relIdOf.set(target, rel.id);
      } else if (!parts.has(target) || read(target) === null) {
        keptUnmodelled.push(rel.id);
      } else {
        remove.add(rel.id);
      }
    }
    const ids = new Set(rels.map((rel) => rel.id));
    const added: string[] = [];
    const listed: string[] = [];
    const base = relationshipBase(rels);
    for (const table of plan.sheet.tables ?? []) {
      let partName = claim.get(table);
      if (partName === undefined) {
        partName = nextFreePartName(taken, "xl/tables/table1.xml");
        taken.add(partName);
        out.added.push(partName);
        out.texts.set(partName, newTablePartXml(table, mintId(table)));
      } else if (table.dirty || !sameTable(read(partName), table)) {
        out.texts.set(partName, regenerateTablePart(partTextOf(parts, partName) ?? "", table, idFor.get(table) ?? mintId(table)));
      }
      let rId = relIdOf.get(partName);
      if (rId === undefined) {
        rId = nextRelationshipId(ids);
        ids.add(rId);
        added.push(relationshipXml(rId, `${base}/table`, relativeTarget(sheetPart, partName)));
        relIdOf.set(partName, rId);
      }
      listed.push(rId);
    }
    listed.push(...keptUnmodelled);
    if (remove.size > 0 || added.length > 0) {
      const next = current === undefined ? `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${added.join("")}</Relationships>` : spliceRels(current, remove, added);
      out.rels.set(relsPath, next);
      out.edges.set(sheetPart, readRels(next).flatMap((rel) => targetPart(sheetPart, rel) ?? []));
    }
    const sheetXml = plan.retained ? partTextOf(parts, sheetPart) : undefined;
    const before = sheetXml === undefined ? [] : (worksheetTablePartIds(sheetXml) ?? []);
    if (before.length !== listed.length || before.some((id, index) => id !== listed[index])) {
      out.tableParts.set(sheetPart, listed);
    }
  }
  return out;
}

/** A table part's `id`. */
function partTableId(xml: string): number | undefined {
  const match = /<(?:[\w.-]+:)?table\b[^>]*?\sid\s*=\s*["'](\d+)["']/.exec(xml);
  return match ? Number(match[1]) : undefined;
}

/** What sameTable compares of a table. */
function tableShape(each: Omit<SheetTable, "id">): string {
  return JSON.stringify([
    each.name,
    formatRange(each.ref),
    each.headerRow,
    each.totalsRow,
    each.columns.map((column) => [column.name, column.xlsxId, column.totalsFunction, column.totalsFormula, column.totalsLabel, column.calculatedFormula]),
    [each.style.name, each.style.showFirstColumn, each.style.showLastColumn, each.style.showRowStripes, each.style.showColumnStripes],
    each.autoFilter !== undefined,
    each.locked === true,
  ]);
}

/** Whether a part still says what the model's table says. */
function sameTable(source: Omit<SheetTable, "id"> | null, table: SheetTable): boolean {
  return source !== null && tableShape(source) === tableShape(table);
}
