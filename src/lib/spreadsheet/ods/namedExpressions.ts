// ODS names on the wire (SPEC-133 "Names on the wire", "Tables in ODS") —
// THIS FILE is the normative source for how the model's defined names are
// read from and written to content.xml, and for the database range an
// Excel table becomes (ODS has no tables).
//
// Read (readOdsNames):
// - `table:named-expressions` directly in `office:spreadsheet` (the content
//   envelope's suffix) is the workbook's scope; the one a `table:table`
//   holds after its rows (its sourceXml) is that sheet's. Workbook names
//   come first in source order, then each table's, in sheet order.
// - `table:named-range`: `table:cell-range-address` through
//   odsRangeAddressToCanonical; `table:named-expression`: `table:expression`
//   through fromOdsFormula (its dialect prefix resolved against the root's
//   namespaces) and SPEC-119's parse check. Both are rebased from
//   `table:base-cell-address` to A1 — rebaseRelative(text, −base.row,
//   −base.col): relative parts move, wrapping modulo the grid; absolute
//   parts stay — so `formula` is relative to A1 (no base cell: A1).
// - `sourceAttrs`: the element's other attributes, verbatim with their
//   leading whitespace, in source order (` table:range-usable-as="…"`).
// - Text neither reads (`$#REF!.A1`, an unknown dialect, an unreadable base
//   cell) is kept `formulaUnparsed`: `formula` is the attribute value as
//   written and `sourceAttrs` every attribute but `table:name`, verbatim —
//   the element is written back as read.
// - A `table:database-range` never becomes a table (SPEC-128 reads the
//   autofilter ones).
//
// Write (the writer calls these; the generated text never goes through the
// sheet-name pass):
// - A scope's names: every model name of that scope but the reserved
//   `_xlnm.*` ones (ODS print ranges and filters are SPEC-136's and
//   SPEC-128's own mapping — they stay in the model only), `_xlfn.`/`_xlpm.`
//   placeholders, xlsx orphans and unparsed text not read from an ODS name.
// - A reference-only formula — one sheet-qualified cell or cell range,
//   3-D included, after expandStructuredReferences (what one cell-range
//   address writes) — is a `table:named-range` (`table:cell-range-address`
//   from canonicalToOdsRangeAddress); anything else a
//   `table:named-expression` (`table:expression` = toOdsFormula of the
//   formula, its structured references expanded to A1, hosted at A1 of the
//   base cell's sheet). The base cell is `$<scope sheet, else the first
//   sheet>.$A$1`, so the A1-relative formula needs no rebase. Attributes in
//   LibreOffice's order: `table:name`, `table:base-cell-address`, the address
//   or expression, then the `table:` attributes of `sourceAttrs`
//   (`table:range-usable-as` only on a named range).
// - When: whenever content.xml is rebuilt — from scratch, and on the
//   substitution path (which `namesDirty` alone now takes). A scope's
//   retained element (already through the writer's sheet-name pass) is kept
//   verbatim while reading it gives exactly the model's names of that scope
//   (name, formula, unparsed flag and sourceAttrs, in order); otherwise it
//   is replaced in place by the generated one (removed when the scope has
//   none to write) or, where there was none, inserted — the workbook's at
//   the start of the suffix (right after the last table, before
//   `table:database-ranges`), a sheet's as the last child of its table. The
//   same rule covers a clean table (otherwise verbatim), so a name an edit
//   moved, renamed or deleted always reaches the file.
//
// Tables: each SheetTable is written as `<table:database-range
// table:name="{name}" table:target-range-address="{Sheet.A1:Sheet.D10}"
// table:contains-header="{headerRow}" table:display-filter-buttons="{autoFilter
// present}"/>` at the end of the suffix's `table:database-ranges` (created
// after `table:named-expressions`, before `table:data-pilot-tables`, when
// absent). A table keeps its name, which its formulas used: any other
// database range of the same name (case-insensitive — SPEC-128's or a
// retained one) is renamed `<name>_<n>`, the smallest n ≥ 2 no database
// range or table holds. The look baked into the cells and the structured
// references written as ranges are the table writer's (tableXml.ts).
import { readOdsFormula, toOdsFormula } from "../formula/dialect";
import { formatReference } from "../formula/print";
import { tokenize } from "../formula/tokens";
import type { DefinedName, Sheet, SheetTable, SpreadsheetModel } from "../model";
import { isOrphanName, isPlaceholderName, isReservedName, rebaseRelative } from "../names";
import { expandStructuredReferences } from "../tables";
import { attributeValue, escapeAttributeText, escapeAttributeValue, setAttribute, startTagAt } from "../xmlSplice";
import { odsDatabaseRanges, odsTargetRangeAddress } from "./databaseRanges";
import { tableStartTag } from "./package";
import { canonicalToOdsRangeAddress, odsRangeAddressToCanonical, parseOdsRangeAddress } from "./rangeAddress";

/** An element's [start, end) in a text. */
export interface OdsElementSpan {
  start: number;
  end: number;
}

const OPEN = "<table:named-expressions";
const CLOSE = "</table:named-expressions>";
const A1 = { row: 0, col: 0 };
const RAW_ATTRIBUTE = /(\s+)([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g;

/** The `table:named-expressions` element starting at `at`. */
function elementAt(xml: string, at: number): OdsElementSpan | undefined {
  if (at < 0 || !/[\s/>]/.test(xml[at + OPEN.length] ?? "")) {
    return undefined;
  }
  const tag = startTagAt(xml, at);
  if (tag.endsWith("/>")) {
    return { start: at, end: at + tag.length };
  }
  const close = xml.indexOf(CLOSE, at + tag.length);
  return close < 0 ? undefined : { start: at, end: close + CLOSE.length };
}

/** The suffix's `table:named-expressions` — the workbook's (the suffix
 * holds no table, so its first one is `office:spreadsheet`'s). */
export function odsWorkbookNamesSpan(suffix: string): OdsElementSpan | undefined {
  return elementAt(suffix, suffix.indexOf(OPEN));
}

/** A table text's own `table:named-expressions`: the last one, after every
 * row (one inside a cell is not the table's). */
export function odsTableNamesSpan(table: string): OdsElementSpan | undefined {
  const at = table.lastIndexOf(OPEN);
  return at > Math.max(table.lastIndexOf("</table:table-row>"), table.lastIndexOf("<table:table-row")) ? elementAt(table, at) : undefined;
}

// --- reading ---------------------------------------------------------------------------------

/** A start tag's attributes but `modelled`, verbatim with their leading
 * whitespace. */
function otherAttributes(tag: string, modelled: readonly string[]): string {
  let out = "";
  for (const match of tag.matchAll(RAW_ATTRIBUTE)) {
    if (!modelled.includes(match[2]!)) {
      out += match[0];
    }
  }
  return out;
}

/** A base cell's position; null when it is not one readable cell. */
function baseCell(text: string): { row: number; col: number } | null {
  const ref = parseOdsRangeAddress(text.trim());
  return ref && !ref.end && ref.start.row !== undefined && ref.start.col !== undefined ? { row: ref.start.row, col: ref.start.col } : null;
}

/** One `table:named-range` / `table:named-expression` start tag → its name
 * (see the module header); undefined without a `table:name`. */
function readNameTag(tag: string, range: boolean, scopeSheetId: string | undefined, namespaces: ReadonlyMap<string, string>): DefinedName | undefined {
  const name = attributeValue(tag, "table:name");
  if (name === undefined || name === "") {
    return undefined;
  }
  const attribute = range ? "table:cell-range-address" : "table:expression";
  const wire = attributeValue(tag, attribute) ?? "";
  const baseText = attributeValue(tag, "table:base-cell-address");
  const base = baseText === undefined ? A1 : baseCell(baseText);
  let canonical: string | null = null;
  if (base) {
    if (range) {
      canonical = odsRangeAddressToCanonical(wire);
    } else {
      const read = readOdsFormula(wire, namespaces, A1);
      canonical = read.formulaUnparsed ? null : read.formula;
    }
  }
  const out: DefinedName = canonical === null || !base ? { name, formula: wire, formulaUnparsed: true } : { name, formula: rebaseRelative(canonical, -base.row, -base.col) };
  if (scopeSheetId !== undefined) {
    out.scopeSheetId = scopeSheetId;
  }
  const rest = otherAttributes(tag, out.formulaUnparsed ? ["table:name"] : ["table:name", attribute, "table:base-cell-address"]);
  if (rest !== "") {
    out.sourceAttrs = rest;
  }
  return out;
}

/** The names of one `table:named-expressions` element of `xml`. */
function readNamesIn(xml: string, span: OdsElementSpan, scopeSheetId: string | undefined, namespaces: ReadonlyMap<string, string>): DefinedName[] {
  const names: DefinedName[] = [];
  const pattern = /<table:named-(range|expression)(?=[\s/>])/g;
  pattern.lastIndex = span.start;
  for (let match = pattern.exec(xml); match !== null && match.index < span.end; match = pattern.exec(xml)) {
    const name = readNameTag(startTagAt(xml, match.index), match[1] === "range", scopeSheetId, namespaces);
    if (name) {
      names.push(name);
    }
  }
  return names;
}

/** The workbook-scope names a suffix defines, in source order. */
export function readOdsWorkbookNames(suffix: string, namespaces: ReadonlyMap<string, string> = new Map()): DefinedName[] {
  const span = odsWorkbookNamesSpan(suffix);
  return span ? readNamesIn(suffix, span, undefined, namespaces) : [];
}

/** The names a table text defines, scoped to `scopeSheetId`, in source order. */
export function readOdsTableNames(table: string, scopeSheetId: string, namespaces: ReadonlyMap<string, string> = new Map()): DefinedName[] {
  const span = odsTableNamesSpan(table);
  return span ? readNamesIn(table, span, scopeSheetId, namespaces) : [];
}

/** Every name of a read workbook: the suffix's, then each sheet's own
 * (scoped to it), in sheet order. */
export function readOdsNames(
  envelope: { suffix: string } | undefined,
  sheets: readonly Pick<Sheet, "id" | "sourceXml">[],
  namespaces: ReadonlyMap<string, string> = new Map(),
): DefinedName[] {
  const names = readOdsWorkbookNames(envelope?.suffix ?? "", namespaces);
  for (const sheet of sheets) {
    if (sheet.sourceXml !== undefined) {
      names.push(...readOdsTableNames(sheet.sourceXml, sheet.id, namespaces));
    }
  }
  return names;
}

// --- writing ---------------------------------------------------------------------------------

type NameModel = Pick<SpreadsheetModel, "sheets" | "names">;

/** The element an unparsed ODS name was read from, by its kept attributes. */
function unparsedKind(name: DefinedName): "range" | "expression" | undefined {
  const attrs = name.sourceAttrs ?? "";
  if (/(?:^|\s)table:cell-range-address\s*=/.test(attrs)) {
    return "range";
  }
  return /(?:^|\s)table:expression\s*=/.test(attrs) ? "expression" : undefined;
}

/** Whether a name is written to an .ods (see the module header). */
export function isOdsWritableName(name: DefinedName): boolean {
  if (isReservedName(name) || isPlaceholderName(name) || isOrphanName(name)) {
    return false;
  }
  return !name.formulaUnparsed || unparsedKind(name) !== undefined;
}

/** The names of a scope (undefined: the workbook's) an .ods holds, in model order. */
export function odsScopeNames(model: Pick<SpreadsheetModel, "names">, scopeSheetId: string | undefined): DefinedName[] {
  return (model.names ?? []).filter((name) => name.scopeSheetId === scopeSheetId && isOdsWritableName(name));
}

/** A reference-only canonical formula's cell-range address — one
 * sheet-qualified cell or cell range — else null. */
export function odsNamedRangeAddress(formula: string): string | null {
  const tokens = tokenize(formula).filter((token) => token.kind !== "end");
  const ref = tokens.length === 1 && tokens[0]!.kind === "ref" ? tokens[0]!.ref : undefined;
  if (!ref || ref.sheet === undefined || ref.start.row === undefined || ref.start.col === undefined || (ref.end && (ref.end.row === undefined || ref.end.col === undefined))) {
    return null;
  }
  return canonicalToOdsRangeAddress(formula);
}

/** The `table:` attributes of `sourceAttrs` a written name keeps. */
function keptAttributes(sourceAttrs: string | undefined, range: boolean): string {
  if (sourceAttrs === undefined) {
    return "";
  }
  const regenerated = ["table:name", "table:base-cell-address", "table:cell-range-address", "table:expression", ...(range ? [] : ["table:range-usable-as"])];
  let out = "";
  for (const match of sourceAttrs.matchAll(RAW_ATTRIBUTE)) {
    if (match[2]!.startsWith("table:") && !regenerated.includes(match[2]!)) {
      out += match[0];
    }
  }
  return out;
}

/** One name's element (see the module header). */
export function odsNameXml(model: NameModel, name: DefinedName): string {
  const head = (kind: string): string => `<table:named-${kind} table:name="${escapeAttributeValue(name.name)}"`;
  if (name.formulaUnparsed) {
    return `${head(unparsedKind(name) ?? "expression")}${name.sourceAttrs ?? ""}/>`;
  }
  const sheet = (name.scopeSheetId === undefined ? undefined : model.sheets.find((each) => each.id === name.scopeSheetId)) ?? model.sheets[0];
  const base = canonicalToOdsRangeAddress(formatReference({ sheet: sheet?.name ?? "Sheet1", start: { row: 0, col: 0, rowAbs: true, colAbs: true } })) ?? "$Sheet1.$A$1";
  const formula = sheet ? expandStructuredReferences(model, name.formula, { sheetId: sheet.id, row: 0, col: 0 }) : name.formula;
  const address = odsNamedRangeAddress(formula);
  const rest = keptAttributes(name.sourceAttrs, address !== null);
  const baseAttr = ` table:base-cell-address="${escapeAttributeValue(base)}"`;
  if (address !== null) {
    return `${head("range")}${baseAttr} table:cell-range-address="${escapeAttributeValue(address)}"${rest}/>`;
  }
  let expression: string;
  try {
    expression = toOdsFormula(formula);
  } catch {
    expression = `of:=${formula}`;
  }
  return `${head("expression")}${baseAttr} table:expression="${escapeAttributeText(expression)}"${rest}/>`;
}

/** A `table:named-expressions` element holding `names` ("" for none). */
export function odsNamedExpressionsXml(model: NameModel, names: readonly DefinedName[]): string {
  return names.length === 0 ? "" : `<table:named-expressions>${names.map((name) => odsNameXml(model, name)).join("")}</table:named-expressions>`;
}

function sameNames(read: readonly DefinedName[], model: readonly DefinedName[]): boolean {
  return (
    read.length === model.length &&
    read.every((name, i) => {
      const other = model[i]!;
      return name.name === other.name && name.formula === other.formula && !name.formulaUnparsed === !other.formulaUnparsed && (name.sourceAttrs ?? "") === (other.sourceAttrs ?? "");
    })
  );
}

/** `text` with a scope's element reconciled with the model (see the module
 * header): kept, replaced, removed, or inserted at `insertAt`. A model with
 * no names to speak of — a record a pre-SPEC-133 reader stored, written
 * through the substitution layout while its reader upgrade re-reads it —
 * leaves the element as written (as xlsx/writer.ts's workbookNames does). */
function reconcile(
  text: string,
  span: OdsElementSpan | undefined,
  model: NameModel,
  scopeSheetId: string | undefined,
  namespaces: ReadonlyMap<string, string>,
  insertAt: number,
): { text: string; generated: string } {
  if (model.names === undefined) {
    return { text, generated: "" };
  }
  const names = odsScopeNames(model, scopeSheetId);
  if (span ? sameNames(readNamesIn(text, span, scopeSheetId, namespaces), names) : names.length === 0) {
    return { text, generated: "" };
  }
  const xml = odsNamedExpressionsXml(model, names);
  const [start, end] = span ? [span.start, span.end] : [insertAt, insertAt];
  return { text: text.slice(0, start) + xml + text.slice(end), generated: xml };
}

/** The substitution path's suffix with the workbook's names (see the module
 * header), and the text generated. */
export function odsSuffixWithNames(suffix: string, model: NameModel, namespaces: ReadonlyMap<string, string>): { suffix: string; generated: string } {
  const { text, generated } = reconcile(suffix, odsWorkbookNamesSpan(suffix), model, undefined, namespaces, 0);
  return { suffix: text, generated };
}

/** A table's trailing children (the table writer's, after its rows) with
 * the sheet's names: the retained element reconciled, a new one last. */
export function odsTrailingWithNames(trailing: string, model: NameModel, sheet: Pick<Sheet, "id">, namespaces: ReadonlyMap<string, string>): string {
  const at = trailing.indexOf(OPEN);
  return reconcile(trailing, at < 0 ? undefined : elementAt(trailing, at), model, sheet.id, namespaces, trailing.length).text;
}

/** A clean table's whole text with the sheet's names reconciled (a new
 * element as the table's last child), and the text generated. */
export function odsTableTextWithNames(table: string, model: NameModel, sheet: Pick<Sheet, "id">, namespaces: ReadonlyMap<string, string>): { text: string; generated: string } {
  const span = odsTableNamesSpan(table);
  if (!span && odsScopeNames(model, sheet.id).length === 0) {
    return { text: table, generated: "" };
  }
  const { start, tag } = tableStartTag(table);
  if (tag.endsWith("/>")) {
    const qname = /^<([^\s/>]+)/.exec(tag)?.[1] ?? "table:table";
    const open = `${tag.replace(/\s*\/>$/, ">")}</${qname}>`;
    return odsTableTextWithNames(table.slice(0, start) + open + table.slice(start + tag.length), model, sheet, namespaces);
  }
  const close = table.lastIndexOf("</");
  return reconcile(table, span, model, sheet.id, namespaces, close < 0 ? table.length : close);
}

// --- tables ------------------------------------------------------------------------------------

/** A table's `table:database-range` (see the module header). */
export function odsTableDatabaseRangeXml(sheetName: string, table: SheetTable): string {
  return (
    `<table:database-range table:name="${escapeAttributeValue(table.name)}" table:target-range-address="${escapeAttributeValue(odsTargetRangeAddress(sheetName, table.ref))}" ` +
    `table:contains-header="${table.headerRow}" table:display-filter-buttons="${table.autoFilter !== undefined}"/>`
  );
}

/** `text` (a suffix, or a from-scratch `table:database-ranges` element or
 * "") with every sheet's tables written as database ranges (see the module
 * header), and the text generated. */
export function odsWithTableDatabaseRanges(text: string, sheets: readonly Pick<Sheet, "name" | "tables">[]): { text: string; generated: string } {
  const tables = sheets.flatMap((sheet) => (sheet.tables ?? []).map((table) => ({ sheet, table })));
  if (tables.length === 0) {
    return { text, generated: "" };
  }
  const xml = tables.map(({ sheet, table }) => odsTableDatabaseRangeXml(sheet.name, table)).join("");
  const found = odsDatabaseRanges(text);
  const container = found.container;
  if (!container) {
    const at = /<table:(?:data-pilot-tables|consolidation|dde-links)(?=[\s/>])|<\/office:spreadsheet>/.exec(text)?.index ?? text.length;
    const element = `<table:database-ranges>${xml}</table:database-ranges>`;
    return { text: text.slice(0, at) + element + text.slice(at), generated: element };
  }
  const tableNames = new Set(tables.map(({ table }) => table.name.toUpperCase()));
  const used = new Set([...tableNames, ...found.ranges.flatMap((range) => (range.name === undefined ? [] : [range.name.toUpperCase()]))]);
  let out = "";
  let at = container.contentStart;
  for (const range of found.ranges) {
    if (range.name === undefined || !tableNames.has(range.name.toUpperCase())) {
      continue;
    }
    let n = 2;
    while (used.has(`${range.name}_${n}`.toUpperCase())) {
      n += 1;
    }
    const renamed = `${range.name}_${n}`;
    used.add(renamed.toUpperCase());
    const tag = startTagAt(text, range.start);
    out += text.slice(at, range.start) + setAttribute(tag, "table:name", renamed);
    at = range.start + tag.length;
  }
  out += text.slice(at, container.contentEnd) + xml;
  const head = container.selfClosing ? startTagAt(text, container.start).replace(/\s*\/>$/, ">") : text.slice(container.start, container.contentStart);
  const element = `${head}${out}</table:database-ranges>`;
  return { text: text.slice(0, container.start) + element + text.slice(container.end), generated: xml };
}
