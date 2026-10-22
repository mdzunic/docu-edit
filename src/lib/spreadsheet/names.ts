// Defined names (SPEC-133 "Name validity", "Resolution", "Reserved names")
// — THIS FILE is the normative source for which text can name a range, how
// a name in a formula finds its DefinedName, how the reserved `_xlnm.*`
// names are stored, and the name commands. The evaluator's `resolveName`
// hook (formula/referenceHooks.ts) evaluates what resolveDefinedName finds.
//
// Validity — the first failing rule wins, in this order:
// | NameProblem | Rule |
// | empty | zero characters |
// | tooLong | over 255 characters |
// | looksLikeReference | an A1 cell inside the grid (`B12`, `$xfd$1`), or R1C1 (`R`, `C`, `R2`, `C3`, `R1C1`, any case) |
// | invalidCharacters | not `^[\p{L}_\\][\p{L}\p{N}_.\\?]*$` |
// | reserved | starts `_xlnm.` |
// | duplicate | the same name (case-insensitive) in the same scope |
// | tableConflict | equals a table name — or, for a table, any name of any scope or another table |
// `XFE1` (outside the grid) and `Sales2024` are valid.
//
// Resolution: a qualified name (`Sheet2!X`) looks in that sheet's scope,
// then the workbook's; an unqualified one in the host sheet's scope, then
// the workbook's. A reserved name also answers to its short form
// (`Print_Area`) within its sheet. `_xlfn.`/`_xlpm.` placeholders, names
// whose `localSheetId` named no sheet (kept in `sourceAttrs`) and unparsed
// names never resolve to a value.
//
// Reserved names: every `_xlnm.*` name is a sheet-scoped DefinedName spelled
// as in the file. They are never shown in the Name Box or the Name Manager
// and never created or edited there; their owners (SPEC-128's
// `_FilterDatabase`, SPEC-136's `Print_Area`) read and write them through
// getReservedName / setReservedName.
import { MAX_COLS, MAX_ROWS, type DefinedName, type SpreadsheetModel } from "./model";
import { lettersToCol } from "./address";
import { rewriteReferences, type RefPart, type ReferenceNode } from "./formula/references";
import { parseSheetPrefix, sheetPrefixOf, tokenize } from "./formula/tokens";
import { formatReference } from "./formula/print";
import { allTables, findTable, mentionedTables } from "./tables";

export type NameProblem = "empty" | "tooLong" | "invalidCharacters" | "looksLikeReference" | "reserved" | "duplicate" | "tableConflict";

export const NAME_PROBLEM_MESSAGES: Readonly<Record<NameProblem, string>> = {
  empty: "Enter a name",
  tooLong: "Names can be at most 255 characters",
  invalidCharacters: "Names must start with a letter, underscore or backslash and can't contain spaces or symbols",
  looksLikeReference: "A name can't look like a cell reference, such as A1 or R1C1",
  reserved: "Names starting with _xlnm. are reserved",
  duplicate: "That name already exists in this scope",
  tableConflict: "That name is already used by a table or a range name",
};

export const MAX_NAME_LENGTH = 255;
export const RESERVED_PREFIX = "_xlnm.";
/** Excel's future-function and LAMBDA-parameter placeholders. */
const PLACEHOLDER_PREFIXES = ["_xlfn.", "_xlpm."];

const NAME_PATTERN = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u;
const A1_CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const R1C1_SHAPED = /^(?:r\d*c\d*|r\d*|c\d*)$/i;

/** Whether text is an A1 cell inside the grid or R1C1-shaped. */
export function looksLikeReference(text: string): boolean {
  if (R1C1_SHAPED.test(text)) {
    return true;
  }
  const match = A1_CELL.exec(text);
  if (!match) {
    return false;
  }
  const col = lettersToCol(match[1]!.toUpperCase());
  const row = Number(match[2]);
  return col !== null && col < MAX_COLS && row >= 1 && row <= MAX_ROWS;
}

const sameText = (a: string, b: string): boolean => a.toUpperCase() === b.toUpperCase();

/** The rules every name and table name share (see the module header). */
function shapeProblem(name: string): NameProblem | null {
  if (name.length === 0) {
    return "empty";
  }
  if (name.length > MAX_NAME_LENGTH) {
    return "tooLong";
  }
  if (looksLikeReference(name)) {
    return "looksLikeReference";
  }
  if (!NAME_PATTERN.test(name)) {
    return "invalidCharacters";
  }
  if (name.toLowerCase().startsWith(RESERVED_PREFIX)) {
    return "reserved";
  }
  return null;
}

/** Why `name` cannot be defined in `scopeSheetId`'s scope (absent: the
 * workbook's), or null. `except` is the name being edited, which the
 * uniqueness rule ignores. */
export function validateName(model: Pick<SpreadsheetModel, "names" | "sheets">, name: string, scopeSheetId?: string, except?: DefinedName): NameProblem | null {
  const shape = shapeProblem(name);
  if (shape) {
    return shape;
  }
  if (model.names.some((each) => each !== except && each.scopeSheetId === scopeSheetId && sameText(each.name, name))) {
    return "duplicate";
  }
  if (findTable(model, name)) {
    return "tableConflict";
  }
  return null;
}

/** Why `name` cannot name a table, or null: the name rules plus
 * workbook-wide uniqueness against every name of every scope and every
 * other table. `exceptTableId` is the table being renamed. */
export function validateTableName(model: Pick<SpreadsheetModel, "names" | "sheets">, name: string, exceptTableId?: string): NameProblem | null {
  const shape = shapeProblem(name);
  if (shape) {
    return shape;
  }
  if (model.names.some((each) => sameText(each.name, name)) || allTables(model).some(({ table }) => table.id !== exceptTableId && sameText(table.name, name))) {
    return "tableConflict";
  }
  return null;
}

/** A new table's name: `Table{n}`, smallest n ≥ 1 — or, for a copy of
 * table `copyOf`, `<copyOf>_<n>`, smallest n ≥ 2 — passing the validator.
 * `taken`: names already handed out in the same command. */
export function uniqueTableName(model: Pick<SpreadsheetModel, "names" | "sheets">, copyOf?: string, taken: readonly string[] = []): string {
  const free = (candidate: string): boolean => validateTableName(model, candidate) === null && !taken.some((name) => sameText(name, candidate));
  for (let n = copyOf === undefined ? 1 : 2; ; n += 1) {
    const candidate = copyOf === undefined ? `Table${n}` : `${copyOf}_${n}`;
    if (free(candidate)) {
      return candidate;
    }
  }
}

// --- kinds of names --------------------------------------------------------------------------

/** An `_xlnm.*` name. */
export function isReservedName(name: Pick<DefinedName, "name">): boolean {
  return name.name.toLowerCase().startsWith(RESERVED_PREFIX);
}

/** An `_xlfn.`/`_xlpm.` placeholder, kept verbatim and never resolved. */
export function isPlaceholderName(name: Pick<DefinedName, "name">): boolean {
  const lower = name.name.toLowerCase();
  return PLACEHOLDER_PREFIXES.some((prefix) => lower.startsWith(prefix));
}

/** A name whose `localSheetId` named no sheet: kept hidden with the
 * attribute in `sourceAttrs`, never resolved, written back as read. */
export function isOrphanName(name: Pick<DefinedName, "sourceAttrs">): boolean {
  return name.sourceAttrs !== undefined && /(?:^|\s)localSheetId\s*=/.test(name.sourceAttrs);
}

/** Whether a name may resolve at all. */
function resolvable(name: DefinedName): boolean {
  return !isPlaceholderName(name) && !isOrphanName(name);
}

/** What the Name Box and the Name Manager show: not hidden, not reserved,
 * not a placeholder or an orphan. */
export function isVisibleName(name: DefinedName): boolean {
  return name.hidden !== true && !isReservedName(name) && resolvable(name);
}

/** Every visible name, in model order. */
export function visibleNames(model: Pick<SpreadsheetModel, "names">): DefinedName[] {
  return model.names.filter(isVisibleName);
}

// --- resolution ----------------------------------------------------------------------------

/** The DefinedName a formula's `name` (`qualifier`: the sheet a qualified
 * `Sheet2!X` names) means on the sheet `hostSheetId` (see the module
 * header); undefined when none. */
export function resolveDefinedName(
  model: Pick<SpreadsheetModel, "names" | "sheets">,
  name: string,
  qualifier: string | undefined,
  hostSheetId: string,
): DefinedName | undefined {
  const scopeId = qualifier === undefined ? hostSheetId : model.sheets.find((sheet) => sameText(sheet.name, qualifier))?.id;
  if (qualifier !== undefined && scopeId === undefined) {
    return undefined;
  }
  const key = name.toUpperCase();
  const reservedKey = (RESERVED_PREFIX + name).toUpperCase();
  let workbook: DefinedName | undefined;
  for (const each of model.names) {
    if (!resolvable(each)) {
      continue;
    }
    const upper = each.name.toUpperCase();
    if (each.scopeSheetId === undefined) {
      if (upper === key) {
        workbook ??= each;
      }
    } else if (each.scopeSheetId === scopeId && (upper === key || upper === reservedKey)) {
      return each;
    }
  }
  return workbook;
}

/** SPEC-132's chart hook (resolveChartRange's `resolveName`): the formula a
 * name means on a sheet; null for none, or one kept unparsed. */
export function nameFormulaResolver(model: Pick<SpreadsheetModel, "names" | "sheets">): (name: string, sheetId: string) => string | null {
  return (name, sheetId) => {
    const found = resolveDefinedName(model, name, undefined, sheetId);
    return found && !found.formulaUnparsed ? found.formula : null;
  };
}

/** A reserved name of a sheet by its short form (`Print_Area`). */
export function getReservedName(model: Pick<SpreadsheetModel, "names">, sheetId: string, short: string): DefinedName | undefined {
  const key = (RESERVED_PREFIX + short).toUpperCase();
  return model.names.find((name) => name.scopeSheetId === sheetId && name.name.toUpperCase() === key);
}

/** Sets (or, with null, removes) a sheet's reserved name: an existing one
 * keeps its place, spelling and other fields; a new one is appended — but a
 * new `_FilterDatabase` goes first, hidden, where Excel and SPEC-128 write
 * it. Marks `namesDirty` when anything changed. */
export function setReservedName(model: SpreadsheetModel, sheetId: string, short: string, formula: string | null): void {
  const existing = getReservedName(model, sheetId, short);
  if (formula === null) {
    if (existing) {
      model.names = model.names.filter((name) => name !== existing);
      model.namesDirty = true;
    }
    return;
  }
  if (existing) {
    if (existing.formula !== formula || existing.formulaUnparsed) {
      const { formulaUnparsed: _unparsed, ...rest } = existing;
      model.names = model.names.map((name) => (name === existing ? { ...rest, formula } : name));
      model.namesDirty = true;
    }
    return;
  }
  const added: DefinedName = { name: RESERVED_PREFIX + short, scopeSheetId: sheetId, formula };
  if (short.toLowerCase() === "_filterdatabase") {
    added.hidden = true;
    model.names = [added, ...model.names];
  } else {
    model.names = [...model.names, added];
  }
  model.namesDirty = true;
}

// --- text ----------------------------------------------------------------------------------

/** A part moved by (dRow, dCol) wrapping modulo the grid. */
function wrapPart(part: RefPart, dRow: number, dCol: number): RefPart {
  const moved: RefPart = { rowAbs: part.rowAbs, colAbs: part.colAbs };
  if (part.row !== undefined) {
    moved.row = part.rowAbs ? part.row : (((part.row + dRow) % MAX_ROWS) + MAX_ROWS) % MAX_ROWS;
  }
  if (part.col !== undefined) {
    moved.col = part.colAbs ? part.col : (((part.col + dCol) % MAX_COLS) + MAX_COLS) % MAX_COLS;
  }
  return moved;
}

/** `text` with every relative part moved by (dRow, dCol), wrapping modulo
 * the grid (Excel's rule for names) — how an ODS name relative to its base
 * cell becomes relative to A1. */
export function rebaseRelative(text: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) {
    return text;
  }
  return rewriteReferences(text, (ref) => {
    const start = wrapPart(ref.start, dRow, dCol);
    const moved: ReferenceNode = { ...ref, start };
    if (ref.end) {
      const end = wrapPart(ref.end, dRow, dCol);
      // A range whose wrap crossed its corners keeps them in order.
      if (start.row !== undefined && end.row !== undefined && start.row > end.row) {
        [start.row, end.row] = [end.row, start.row];
        [start.rowAbs, end.rowAbs] = [end.rowAbs, start.rowAbs];
      }
      if (start.col !== undefined && end.col !== undefined && start.col > end.col) {
        [start.col, end.col] = [end.col, start.col];
        [start.colAbs, end.colAbs] = [end.colAbs, start.colAbs];
      }
      moved.end = end;
    }
    return moved;
  });
}

export interface FormulaMentions {
  /** Upper-cased names of `name` tokens (sheet prefix dropped). */
  names: ReadonlySet<string>;
  /** Upper-cased table names of structured tokens. */
  tables: ReadonlySet<string>;
}

const MENTIONS_CACHE_SIZE = 20_000;
const mentionsCache = new Map<string, FormulaMentions>();
const NO_MENTIONS: FormulaMentions = { names: new Set(), tables: new Set() };

/** The names and tables a canonical formula mentions — a token scan,
 * cached per formula text. */
export function formulaMentions(text: string): FormulaMentions {
  let found = mentionsCache.get(text);
  if (found) {
    return found;
  }
  const names = new Set<string>();
  for (const token of tokenize(text)) {
    if (token.kind === "name") {
      names.add(token.text.slice(sheetPrefixOf(token.text).length).toUpperCase());
    }
  }
  const tables = mentionedTables(text);
  found = names.size === 0 && tables.size === 0 ? NO_MENTIONS : { names, tables };
  if (mentionsCache.size >= MENTIONS_CACHE_SIZE) {
    mentionsCache.clear();
  }
  mentionsCache.set(text, found);
  return found;
}

/** `text` hosted on `hostSheetId` with every `name` token that resolves to
 * `target` renamed `to` (a qualified token keeps its prefix). */
export function renameNameInFormula(
  model: Pick<SpreadsheetModel, "names" | "sheets">,
  text: string,
  hostSheetId: string,
  target: DefinedName,
  to: string,
): string {
  if (!formulaMentions(text).names.has(target.name.toUpperCase())) {
    return text;
  }
  const tokens = tokenize(text);
  let changed = false;
  for (const token of tokens) {
    if (token.kind !== "name") {
      continue;
    }
    const prefix = sheetPrefixOf(token.text);
    const qualifier = prefix === "" ? undefined : parseSheetPrefix(prefix).sheet;
    if (resolveDefinedName(model, token.text.slice(prefix.length), qualifier, hostSheetId) === target) {
      token.text = prefix + to;
      changed = true;
    }
  }
  return changed ? tokens.map((token) => token.lead + token.text).join("") : text;
}

// --- the name commands -----------------------------------------------------------------------

/** Every formula of the workbook a name change can reach, rewritten
 * copy-on-write through `rewrite(text, hostSheetId)`: cell formulas, every
 * name's formula (host = its scope sheet, else the first sheet) and every
 * table's calculated and totals formulas. */
export function rewriteWorkbookFormulas(model: SpreadsheetModel, rewrite: (text: string, hostSheetId: string) => string): void {
  const firstSheetId = model.sheets[0]?.id ?? "";
  let sheetsChanged = false;
  const sheets = model.sheets.map((sheet) => {
    let rowsChanged = false;
    const rows = sheet.rows.map((row) => {
      let cellsChanged = false;
      const cells = row.cells.map((cell) => {
        if (cell.formula === undefined || cell.formulaUnparsed) {
          return cell;
        }
        const formula = rewrite(cell.formula, sheet.id);
        if (formula === cell.formula) {
          return cell;
        }
        cellsChanged = true;
        return { ...cell, formula };
      });
      if (!cellsChanged) {
        return row;
      }
      rowsChanged = true;
      return { ...row, cells };
    });
    let tablesChanged = false;
    const tables = sheet.tables?.map((table) => {
      let columnsChanged = false;
      const columns = table.columns.map((column) => {
        const calculatedFormula = column.calculatedFormula === undefined ? undefined : rewrite(column.calculatedFormula, sheet.id);
        const totalsFormula = column.totalsFormula === undefined ? undefined : rewrite(column.totalsFormula, sheet.id);
        if (calculatedFormula === column.calculatedFormula && totalsFormula === column.totalsFormula) {
          return column;
        }
        columnsChanged = true;
        return {
          ...column,
          ...(calculatedFormula === undefined ? {} : { calculatedFormula }),
          ...(totalsFormula === undefined ? {} : { totalsFormula }),
        };
      });
      if (!columnsChanged) {
        return table;
      }
      tablesChanged = true;
      return { ...table, columns, dirty: true as const };
    });
    if (!rowsChanged && !tablesChanged) {
      return sheet;
    }
    sheetsChanged = true;
    return { ...sheet, ...(rowsChanged ? { rows } : {}), ...(tablesChanged ? { tables } : {}), dirty: true as const };
  });
  if (sheetsChanged) {
    model.sheets = sheets;
  }
  let namesChanged = false;
  const names = model.names.map((name) => {
    if (name.formulaUnparsed) {
      return name;
    }
    const formula = rewrite(name.formula, name.scopeSheetId ?? firstSheetId);
    if (formula === name.formula) {
      return name;
    }
    namesChanged = true;
    return { ...name, formula };
  });
  if (namesChanged) {
    model.names = names;
    model.namesDirty = true;
  }
}

/** Appends a new name (already validated). */
export function defineName(model: SpreadsheetModel, name: DefinedName): void {
  model.names = [...model.names, name];
  model.namesDirty = true;
}

/** Replaces `target` with `next` in place. When the text changed, every
 * formula that used the old name is rewritten to the new one first. */
export function updateName(model: SpreadsheetModel, target: DefinedName, next: DefinedName): void {
  if (target.name !== next.name) {
    rewriteWorkbookFormulas(model, (text, host) => renameNameInFormula(model, text, host, target, next.name));
  }
  // The rename may have replaced the target itself (its own formula).
  model.names = model.names.map((name) => (name === target || (name.name === target.name && name.scopeSheetId === target.scopeSheetId) ? next : name));
  model.namesDirty = true;
}

/** Removes a name; formulas using it show #NAME? on recalculation. */
export function deleteName(model: SpreadsheetModel, target: DefinedName): void {
  const names = model.names.filter((name) => name !== target);
  if (names.length !== model.names.length) {
    model.names = names;
    model.namesDirty = true;
  }
}

/** The canonical, absolute, sheet-qualified text of a selection's ranges
 * (a union for several areas) — what the Name Box defines a name as. */
export function absoluteRangesFormula(sheetName: string, ranges: readonly { top: number; left: number; bottom: number; right: number }[]): string {
  const texts = ranges.map((range) =>
    formatReference({
      sheet: sheetName,
      start: { row: range.top, col: range.left, rowAbs: true, colAbs: true },
      ...(range.top === range.bottom && range.left === range.right ? {} : { end: { row: range.bottom, col: range.right, rowAbs: true, colAbs: true } }),
    }),
  );
  return texts.length === 1 ? texts[0]! : `(${texts.join(",")})`;
}
