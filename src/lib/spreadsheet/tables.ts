// Excel tables (SPEC-133 "Structured references", "Canonical and wire
// forms") — THIS FILE is the normative source for what a structured
// reference means and how it is spelled. SPEC-119 keeps a table reference as
// one `structured` token (`Sales[@Amount]`, `[Qty]`); this module parses its
// text, resolves it against `Sheet.tables`, and makes every rewrite of it —
// all token-level over `tokenize`, touching `structured` tokens (and a bare
// table name's `name` token) only, so whitespace and every other token
// survive byte-for-byte and a formula nothing touched comes back identical.
//
// | Form | Resolves to |
// | `T`, `T[]`, `T[#Data]` | the data body, all columns |
// | `T[#All]` | `ref` |
// | `T[#Headers]` / `T[#Totals]` | the header / totals row (#REF! when absent) |
// | `T[Col]`, `T[[Col]]` | Col's data cells |
// | `T[[#Headers],[Col]]`, `T[[#Totals],[Col]]`, `T[[#All],[Col]]` | that part of Col |
// | `T[[#Headers],[#Data],[Col]]`, `T[[#Data],[#Totals],[Col]]` | the adjacent parts of Col |
// | `T[[A]:[C]]` (either order) | the data cells of columns A..C |
// | `T[@Col]`, `T[@]`, `T[@[A]:[C]]` | the host row's cells (a host outside the data body: #VALUE!) |
// A missing table, column, header or totals row is #REF!.
//
// Spelling: `Cell.formula` always names the table and writes this-row as
// `@` (`Sales[@Amount]`); the xlsx file writes `Sales[[#This Row],[Amount]]`
// (structuredToXlsx / structuredFromXlsx, around SPEC-119's dialect); ODS
// has no tables, so its formulas get the A1 equivalent
// (expandStructuredReferences: this-row column-absolute and row-relative,
// `$C7`; the rest absolute, `$C$2:$C$10`). A column name holding any of
// , : . [ ] # ' " { } $ ^ & * + = - > < /, a tab or a line break prints in
// the inner-bracket form `T[[a.b]]`; `[`, `]`, `#` and `'` are escaped with
// `'`.
//
// The data body is `ref` minus the header row (if any) and the totals row
// (if any); it is never empty.
import type { GridRange } from "./address";
import { colToLetters } from "./address";
import { formatSheetPrefix } from "./formula/print";
import { joinTokens, sheetPrefixOf, tokenize, type Token } from "./formula/tokens";
import type { Area, RangeValue } from "./formula/values";
import { REF_ERROR, VALUE_ERROR } from "./formula/values";
import type { CellError, Sheet, SheetTable, SpreadsheetModel } from "./model";

/** A structured reference's special items, as this module spells them. */
export type StructuredItem = "#All" | "#Data" | "#Headers" | "#Totals";

/** A parsed structured reference. */
export interface StructuredSpec {
  /** The table name as written; absent for `[@Qty]`-style references. */
  table?: string;
  /** The special items named, in source order (none = the data body). */
  items: StructuredItem[];
  /** `@` (canonical) or `[#This Row]` (xlsx): the host row. */
  thisRow: boolean;
  /** The column, or the column range `[A]:[C]` (unescaped names). */
  columns?: { first: string; last?: string };
}

const ITEMS: Readonly<Record<string, StructuredItem | "#This Row">> = {
  "#ALL": "#All",
  "#DATA": "#Data",
  "#HEADERS": "#Headers",
  "#TOTALS": "#Totals",
  "#THIS ROW": "#This Row",
};

// --- the grammar -------------------------------------------------------------------------

/** The bracketed part at `i` (`text[i]` is `[`): its unescaped content and
 * the index past its `]`; null when it does not close. */
function bracketAt(text: string, i: number): { content: string; raw: string; end: number } | null {
  let content = "";
  for (let j = i + 1; j < text.length; j += 1) {
    const ch = text[j]!;
    if (ch === "'") {
      j += 1;
      if (j >= text.length) {
        return null;
      }
      content += text[j]!;
    } else if (ch === "]") {
      return { content, raw: text.slice(i + 1, j), end: j + 1 };
    } else if (ch === "[") {
      return null;
    } else {
      content += ch;
    }
  }
  return null;
}

/** A simple (unbracketed) column name's unescaped text. */
function unescapeSimple(text: string): string | null {
  let out = "";
  for (let j = 0; j < text.length; j += 1) {
    const ch = text[j]!;
    if (ch === "'") {
      j += 1;
      if (j >= text.length) {
        return null;
      }
      out += text[j]!;
    } else if (ch === "[" || ch === "]") {
      return null;
    } else {
      out += ch;
    }
  }
  return out;
}

const skipSpaces = (text: string, i: number): number => {
  let j = i;
  while (text[j] === " ") {
    j += 1;
  }
  return j;
};

/** A comma-separated list of bracketed parts (`[#Headers],[A]:[C]`):
 * special items and at most one column (range); null when malformed. */
function parseBracketList(inner: string): { items: (StructuredItem | "#This Row")[]; columns?: { first: string; last?: string } } | null {
  const items: (StructuredItem | "#This Row")[] = [];
  let columns: { first: string; last?: string } | undefined;
  let i = skipSpaces(inner, 0);
  while (i < inner.length) {
    if (inner[i] !== "[") {
      return null;
    }
    const part = bracketAt(inner, i);
    if (!part) {
      return null;
    }
    i = skipSpaces(inner, part.end);
    const item = part.content.startsWith("#") ? ITEMS[part.content.trim().toUpperCase()] : undefined;
    if (part.content.startsWith("#") && part.raw.startsWith("#")) {
      if (!item) {
        return null;
      }
      items.push(item);
    } else {
      if (columns) {
        return null;
      }
      columns = { first: part.content };
      if (inner[i] === ":") {
        i = skipSpaces(inner, i + 1);
        const last = inner[i] === "[" ? bracketAt(inner, i) : null;
        if (!last) {
          return null;
        }
        columns.last = last.content;
        i = skipSpaces(inner, last.end);
      }
    }
    if (i >= inner.length) {
      break;
    }
    if (inner[i] !== ",") {
      return null;
    }
    i = skipSpaces(inner, i + 1);
  }
  return columns ? { items, columns } : { items };
}

/** A structured token's text parsed (see the module header); null when it
 * is not a structured reference this module reads. */
export function parseStructuredText(text: string): StructuredSpec | null {
  const open = text.indexOf("[");
  if (open < 0 || !text.endsWith("]")) {
    return null;
  }
  const table = open > 0 ? text.slice(0, open) : undefined;
  const inner = text.slice(open + 1, -1);
  const base: StructuredSpec = table === undefined ? { items: [], thisRow: false } : { table, items: [], thisRow: false };
  const trimmed = inner.trim();
  if (trimmed === "") {
    return base;
  }
  if (trimmed.startsWith("@")) {
    const rest = trimmed.slice(1).trim();
    if (rest === "") {
      return { ...base, thisRow: true };
    }
    if (rest.startsWith("[")) {
      const list = parseBracketList(rest);
      return list && list.items.length === 0 && list.columns ? { ...base, thisRow: true, columns: list.columns } : null;
    }
    const name = unescapeSimple(rest);
    return name === null ? null : { ...base, thisRow: true, columns: { first: name } };
  }
  let list: ReturnType<typeof parseBracketList>;
  if (trimmed.startsWith("[")) {
    list = parseBracketList(trimmed);
  } else if (trimmed.startsWith("#")) {
    const item = ITEMS[trimmed.toUpperCase()];
    list = item ? { items: [item] } : null;
  } else {
    const name = unescapeSimple(inner);
    list = name === null ? null : { items: [], columns: { first: name } };
  }
  if (!list) {
    return null;
  }
  const thisRow = list.items.includes("#This Row");
  const items = list.items.filter((item): item is StructuredItem => item !== "#This Row");
  if (thisRow && items.length > 0) {
    return null;
  }
  return { ...base, items, thisRow, ...(list.columns ? { columns: list.columns } : {}) };
}

// --- printing ----------------------------------------------------------------------------

const BRACKETED = /[,:.[\]#'"{}$^&*+=\-></\t\r\n]/;

/** Whether a column name prints in the inner-bracket form `T[[a.b]]`. */
export function columnNeedsBrackets(name: string): boolean {
  return BRACKETED.test(name);
}

/** A column name with `[`, `]`, `#` and `'` escaped by `'`. */
export function escapeColumnName(name: string): string {
  return name.replace(/[[\]#']/g, (ch) => `'${ch}`);
}

const bracket = (name: string): string => `[${escapeColumnName(name)}]`;

function columnsText(columns: { first: string; last?: string }): string {
  return columns.last === undefined ? bracket(columns.first) : `${bracket(columns.first)}:${bracket(columns.last)}`;
}

/** A spec in the canonical spelling (`@` for this-row). */
export function formatStructured(spec: StructuredSpec): string {
  const table = spec.table ?? "";
  const { columns, items } = spec;
  if (spec.thisRow) {
    if (!columns) {
      return `${table}[@]`;
    }
    if (columns.last === undefined && !columnNeedsBrackets(columns.first)) {
      return `${table}[@${escapeColumnName(columns.first)}]`;
    }
    return `${table}[@${columnsText(columns)}]`;
  }
  if (!columns) {
    if (items.length === 0) {
      return `${table}[]`;
    }
    return items.length === 1 ? `${table}[${items[0]}]` : `${table}[${items.map((item) => `[${item}]`).join(",")}]`;
  }
  if (items.length === 0) {
    if (columns.last === undefined && !columnNeedsBrackets(columns.first)) {
      return `${table}[${escapeColumnName(columns.first)}]`;
    }
    return `${table}[${columnsText(columns)}]`;
  }
  return `${table}[${[...items.map((item) => `[${item}]`), columnsText(columns)].join(",")}]`;
}

/** A spec as an xlsx file spells it: this-row as `[#This Row]`. */
export function formatStructuredXlsx(spec: StructuredSpec): string {
  if (!spec.thisRow) {
    return formatStructured(spec);
  }
  const table = spec.table ?? "";
  return spec.columns ? `${table}[[#This Row],${columnsText(spec.columns)}]` : `${table}[#This Row]`;
}

// --- token-level rewrites ------------------------------------------------------------------

/** Rewrites structured tokens (and, with `names`, `name` tokens) through
 * `visit`; the identical string when nothing changed. */
function rewriteTokens(text: string, visit: (token: Token) => string | undefined): string {
  const tokens = tokenize(text);
  let changed = false;
  for (const token of tokens) {
    if (token.kind !== "structured" && token.kind !== "name") {
      continue;
    }
    const next = visit(token);
    if (next !== undefined && next !== token.text) {
      token.text = next;
      changed = true;
    }
  }
  return changed ? joinTokens(tokens) : text;
}

/** Rewrites every structured token whose text parses through `visit`. */
export function rewriteStructured(text: string, visit: (spec: StructuredSpec, token: Token) => string | undefined): string {
  if (!text.includes("[")) {
    return text;
  }
  return rewriteTokens(text, (token) => {
    if (token.kind !== "structured" || token.unterminated) {
      return undefined;
    }
    const spec = parseStructuredText(token.text);
    return spec ? visit(spec, token) : undefined;
  });
}

/** canonical → the xlsx file's form: `@` becomes `[#This Row]`. */
export function structuredToXlsx(canonical: string): string {
  return rewriteStructured(canonical, (spec) => (spec.thisRow ? formatStructuredXlsx(spec) : undefined));
}

/** The file's `[#This Row]` (after fromXlsxFormula) → canonical `@`. */
export function structuredFromXlsx(text: string): string {
  return rewriteStructured(text, (spec, token) => (spec.thisRow && !/\[\s*@/.test(token.text) ? formatStructured(spec) : undefined));
}

/** A formula committed inside (or growing) `tableName`'s table: every
 * structured reference that omits the table name is qualified with it, and
 * `[#This Row]` becomes `@`. */
export function qualifyStructuredReferences(text: string, tableName: string): string {
  return rewriteStructured(text, (spec, token) => {
    if (spec.table !== undefined) {
      return spec.thisRow && !/\[\s*@/.test(token.text) ? formatStructured(spec) : undefined;
    }
    return spec.thisRow ? formatStructured({ ...spec, table: tableName }) : tableName + token.text;
  });
}

const same = (a: string | undefined, b: string): boolean => a !== undefined && a.toUpperCase() === b.toUpperCase();

/** Every reference to table `from` names `to` instead — structured tokens
 * and a bare table-name token (`=SUM(Sales)`). */
export function renameTableInFormula(text: string, from: string, to: string): string {
  return rewriteTokens(text, (token) => {
    if (token.kind === "name") {
      return sheetPrefixOf(token.text) === "" && same(token.text, from) ? to : undefined;
    }
    const open = token.text.indexOf("[");
    return open > 0 && same(token.text.slice(0, open), from) ? to + token.text.slice(open) : undefined;
  });
}

/** Every reference to column `from` of table `table` names `to` instead. */
export function renameColumnInFormula(text: string, table: string, from: string, to: string): string {
  return rewriteStructured(text, (spec) => {
    if (!same(spec.table, table) || !spec.columns) {
      return undefined;
    }
    const first = same(spec.columns.first, from) ? to : spec.columns.first;
    const last = spec.columns.last !== undefined && same(spec.columns.last, from) ? to : spec.columns.last;
    if (first === spec.columns.first && last === spec.columns.last) {
      return undefined;
    }
    return formatStructured({ ...spec, columns: last === undefined ? { first } : { first, last } });
  });
}

/** Every reference to table `table` — or, with `column`, to that column of
 * it — becomes `#REF!` (the table or column was deleted). */
export function removeTableReferences(text: string, table: string, column?: string): string {
  return rewriteTokens(text, (token) => {
    if (token.kind === "name") {
      return column === undefined && sheetPrefixOf(token.text) === "" && same(token.text, table) ? "#REF!" : undefined;
    }
    if (token.unterminated) {
      return undefined;
    }
    const spec = parseStructuredText(token.text);
    if (!spec || !same(spec.table, table)) {
      return undefined;
    }
    if (column === undefined) {
      return "#REF!";
    }
    return spec.columns && (same(spec.columns.first, column) || same(spec.columns.last, column)) ? "#REF!" : undefined;
  });
}

/** The table names a formula's structured references (and bare names)
 * mention, upper-cased. */
export function mentionedTables(text: string): Set<string> {
  const out = new Set<string>();
  if (!text.includes("[")) {
    return out;
  }
  for (const token of tokenize(text)) {
    if (token.kind === "structured") {
      const open = token.text.indexOf("[");
      if (open > 0) {
        out.add(token.text.slice(0, open).toUpperCase());
      }
    }
  }
  return out;
}

// --- the model -----------------------------------------------------------------------------

/** A sheet's tables (none when absent). */
export function sheetTables(sheet: Pick<Sheet, "tables">): readonly SheetTable[] {
  return sheet.tables ?? [];
}

const inRange = (range: GridRange, row: number, col: number): boolean =>
  row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;

/** The table whose `ref` holds a cell. */
export function tableAt(sheet: Pick<Sheet, "tables">, row: number, col: number): SheetTable | undefined {
  return sheet.tables?.find((table) => inRange(table.ref, row, col));
}

/** The table named `name` (case-insensitively) and its sheet. */
export function findTable(model: Pick<SpreadsheetModel, "sheets">, name: string): { sheet: Sheet; table: SheetTable } | undefined {
  const key = name.toUpperCase();
  for (const sheet of model.sheets) {
    const table = sheet.tables?.find((each) => each.name.toUpperCase() === key);
    if (table) {
      return { sheet, table };
    }
  }
  return undefined;
}

/** Every table of the workbook with its sheet. */
export function allTables(model: Pick<SpreadsheetModel, "sheets">): { sheet: Sheet; table: SheetTable }[] {
  return model.sheets.flatMap((sheet) => sheetTables(sheet).map((table) => ({ sheet, table })));
}

/** The table's data body: `ref` minus its header and totals rows. */
export function dataBody(table: Pick<SheetTable, "ref" | "headerRow" | "totalsRow">): GridRange {
  return {
    top: table.ref.top + (table.headerRow ? 1 : 0),
    left: table.ref.left,
    bottom: table.ref.bottom - (table.totalsRow ? 1 : 0),
    right: table.ref.right,
  };
}

/** The header row's index, when the table has one. */
export function headerRowIndex(table: Pick<SheetTable, "ref" | "headerRow">): number | undefined {
  return table.headerRow ? table.ref.top : undefined;
}

/** The totals row's index, when the table has one. */
export function totalsRowIndex(table: Pick<SheetTable, "ref" | "totalsRow">): number | undefined {
  return table.totalsRow ? table.ref.bottom : undefined;
}

/** A column's index within the table (case-insensitive), -1 when none. */
export function columnIndex(table: Pick<SheetTable, "columns">, name: string): number {
  const key = name.toUpperCase();
  return table.columns.findIndex((column) => column.name.toUpperCase() === key);
}

/** The rows a spec's items cover, or #REF! (a missing header or totals
 * row, parts that are not adjacent). */
function itemRows(table: SheetTable, items: readonly StructuredItem[]): { top: number; bottom: number } | CellError {
  const body = dataBody(table);
  if (items.length === 0) {
    return { top: body.top, bottom: body.bottom };
  }
  if (items.includes("#All")) {
    return { top: table.ref.top, bottom: table.ref.bottom };
  }
  const parts: { top: number; bottom: number }[] = [];
  for (const item of ["#Headers", "#Data", "#Totals"] as const) {
    if (!items.includes(item)) {
      continue;
    }
    if (item === "#Headers") {
      if (!table.headerRow) {
        return REF_ERROR;
      }
      parts.push({ top: table.ref.top, bottom: table.ref.top });
    } else if (item === "#Totals") {
      if (!table.totalsRow) {
        return REF_ERROR;
      }
      parts.push({ top: table.ref.bottom, bottom: table.ref.bottom });
    } else {
      parts.push({ top: body.top, bottom: body.bottom });
    }
  }
  for (let i = 1; i < parts.length; i += 1) {
    if (parts[i]!.top !== parts[i - 1]!.bottom + 1) {
      return REF_ERROR;
    }
  }
  return { top: parts[0]!.top, bottom: parts.at(-1)!.bottom };
}

/** The area a parsed spec covers in `table` on `sheetId`, or its error. */
export function structuredArea(
  table: SheetTable,
  sheetId: string,
  spec: StructuredSpec,
  host: { sheetId: string; row: number; col: number },
): Area | CellError {
  let left = table.ref.left;
  let right = table.ref.right;
  if (spec.columns) {
    const first = columnIndex(table, spec.columns.first);
    const last = spec.columns.last === undefined ? first : columnIndex(table, spec.columns.last);
    if (first < 0 || last < 0) {
      return REF_ERROR;
    }
    left = table.ref.left + Math.min(first, last);
    right = table.ref.left + Math.max(first, last);
  }
  if (spec.thisRow) {
    const body = dataBody(table);
    if (host.sheetId !== sheetId || !inRange(body, host.row, host.col)) {
      return VALUE_ERROR;
    }
    return { sheetId, top: host.row, left, bottom: host.row, right };
  }
  const rows = itemRows(table, spec.items);
  if ("error" in rows) {
    return rows;
  }
  return { sheetId, top: rows.top, left, bottom: rows.bottom, right };
}

/** A structured node at `host` (see the module header): its range, or
 * #REF! / #VALUE!. A node without a table name means the host's table. */
export function resolveStructuredReference(
  model: Pick<SpreadsheetModel, "sheets">,
  node: { table?: string; text: string },
  host: { sheetId: string; row: number; col: number },
): RangeValue | CellError {
  let found: { sheet: Sheet; table: SheetTable } | undefined;
  if (node.table !== undefined || !node.text.includes("[")) {
    // A bare `T` names its table by its whole text.
    found = findTable(model, node.table ?? node.text);
  } else {
    const sheet = model.sheets.find((each) => each.id === host.sheetId);
    const table = sheet ? tableAt(sheet, host.row, host.col) : undefined;
    found = sheet && table ? { sheet, table } : undefined;
  }
  if (!found) {
    return REF_ERROR;
  }
  const spec = node.text.includes("[") ? parseStructuredText(node.text) : { table: node.text, items: [], thisRow: false };
  if (!spec) {
    return REF_ERROR;
  }
  const area = structuredArea(found.table, found.sheet.id, spec, host);
  return "error" in area ? area : { kind: "range", areas: [area] };
}

// --- ODS: structured references as A1 ------------------------------------------------------

/** One area as canonical A1 text on `host`'s sheet: this-row parts
 * column-absolute and row-relative (`$C7`), the rest absolute. */
function areaText(model: Pick<SpreadsheetModel, "sheets">, area: Area, hostSheetId: string, thisRow: boolean): string {
  const prefix = area.sheetId === hostSheetId ? "" : formatSheetPrefix(model.sheets.find((sheet) => sheet.id === area.sheetId)?.name ?? "");
  const row = (index: number): string => (thisRow ? `${index + 1}` : `$${index + 1}`);
  const start = `$${colToLetters(area.left)}${row(area.top)}`;
  const end = `$${colToLetters(area.right)}${row(area.bottom)}`;
  return prefix + (start === end ? start : `${start}:${end}`);
}

/** Every structured reference (and bare table name) of a canonical formula
 * hosted at `host` replaced by its A1 equivalent — what an ODS file holds
 * (see the module header); an unresolvable one becomes its error. With
 * `onlyTable`, references to other tables are left as they are (Convert to
 * range). */
export function expandStructuredReferences(
  model: Pick<SpreadsheetModel, "sheets" | "names">,
  text: string,
  host: { sheetId: string; row: number; col: number },
  onlyTable?: string,
): string {
  return rewriteTokens(text, (token) => {
    if (token.kind === "name") {
      if (sheetPrefixOf(token.text) !== "" || model.names.some((name) => same(name.name, token.text)) || (onlyTable !== undefined && !same(token.text, onlyTable))) {
        return undefined;
      }
      const found = findTable(model, token.text);
      return found ? areaText(model, { sheetId: found.sheet.id, ...dataBody(found.table) }, host.sheetId, false) : undefined;
    }
    if (token.unterminated) {
      return undefined;
    }
    const open = token.text.indexOf("[");
    if (onlyTable !== undefined) {
      const hostSheet = open > 0 ? undefined : model.sheets.find((sheet) => sheet.id === host.sheetId);
      const named = open > 0 ? token.text.slice(0, open) : hostSheet && tableAt(hostSheet, host.row, host.col)?.name;
      if (!same(named, onlyTable)) {
        return undefined;
      }
    }
    const resolved = resolveStructuredReference(model, open > 0 ? { table: token.text.slice(0, open), text: token.text } : { text: token.text }, host);
    if ("error" in resolved) {
      return resolved.error;
    }
    const spec = parseStructuredText(token.text);
    return areaText(model, resolved.areas[0]!, host.sheetId, spec?.thisRow === true);
  });
}
