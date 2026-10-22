// SPEC-133's names surfaces, framework-free (useNames holds the state): the
// Name Box's interpretation of a typed entry ("The Name Box"), its dropdown's
// options, and the Name Manager's rows.
//
// - interpretNameBox, in order (SPEC-134 replaced the first two steps with
//   Go To's resolver, so the Name Box and the Go To dialog accept exactly the
//   same text): (1) lib/spreadsheet/find.ts's resolveGoToTarget succeeds — a
//   reference, a visible name (evaluated at the active cell) or a table → a
//   jump; (2) it fails for a reason other than "not a reference or a name" —
//   a name that is no reference (NOT_A_RANGE_MESSAGE), a hidden sheet
//   (HIDDEN_SHEET_MESSAGE) — → that refusal; (3) otherwise a new
//   workbook-scope name, validated (validateName; a problem is refused with
//   its NAME_PROBLEM_MESSAGES text). Valid names never parse as references,
//   so (1) and (3) never compete.
// - nameJump / tableJump: where a dropdown pick selects — a name's
//   evaluation at the active cell (the first area's sheet, its areas only),
//   a table's data body; a hidden sheet is refused with HIDDEN_SHEET_MESSAGE,
//   a non-grid one as not a range.
// - nameBoxOptions: the visible names of the workbook's and the active
//   sheet's scope (a sheet-scoped name hides a workbook one of the same
//   text) and every table, merged and sorted alphabetically,
//   case-insensitively.
// - nameManagerRows: one row per visible name of every scope, sorted by name
//   (case-insensitive) then scope: its value — generalText of its evaluation
//   at the active cell (its scope sheet's for a sheet-scoped name); a
//   single-cell reference shows that cell's value, a multi-cell or several
//   areas, like an array, ARRAY_VALUE_TEXT ("{…}"), an error its code — its
//   refers-to ("=" + the formula, relative parts shown relative to the active
//   cell, as Excel shows them) and its scope ("Workbook" or the sheet's name).
import type { GridRange } from "../../lib/spreadsheet/address";
import { GO_TO_HIDDEN_SHEET_MESSAGE, GO_TO_INVALID_MESSAGE, GO_TO_NOT_A_RANGE_MESSAGE, resolveGoToTarget } from "../../lib/spreadsheet/find";
import { storedValueContext } from "../../lib/spreadsheet/formula/evaluate";
import { parseFormula } from "../../lib/spreadsheet/formula/parse";
import { printFormula } from "../../lib/spreadsheet/formula/print";
import { evaluateDefinedName } from "../../lib/spreadsheet/formula/referenceHooks";
import { isArray, isEmpty, isRange, type EvalValue } from "../../lib/spreadsheet/formula/values";
import { generalText } from "../../lib/spreadsheet/general";
import type { DefinedName, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { NAME_PROBLEM_MESSAGES, rebaseRelative, validateName, visibleNames } from "../../lib/spreadsheet/names";
import { allTables, dataBody, findTable } from "../../lib/spreadsheet/tables";

/** SPEC-133's strings, which Go To shares (find.ts holds them). */
export const NOT_A_RANGE_MESSAGE = GO_TO_NOT_A_RANGE_MESSAGE;
export const HIDDEN_SHEET_MESSAGE = GO_TO_HIDDEN_SHEET_MESSAGE;
export const FORMULA_PROBLEM_MESSAGE = "There's a problem with this formula";
export const ARRAY_VALUE_TEXT = "{…}";
export const WORKBOOK_SCOPE_LABEL = "Workbook";

/** Where an evaluation happens: the active cell. */
export interface NameHost {
  sheetId: string;
  row: number;
  col: number;
}

export type NameBoxOutcome = { kind: "jump"; sheetId: string; ranges: GridRange[] } | { kind: "define"; name: string } | { kind: "refuse"; message: string };

export interface NameBoxOption {
  kind: "name" | "table";
  name: string;
  /** Unique among the options: `name:<scope>:<NAME>` / `table:<id>`. */
  key: string;
}

export interface NameManagerRow {
  key: string;
  name: string;
  value: string;
  refersTo: string;
  scope: string;
  scopeSheetId?: string;
}

const upper = (text: string): string => text.toUpperCase();
const byName = (a: { name: string }, b: { name: string }): number => a.name.localeCompare(b.name, "en", { sensitivity: "base" });

/** A name's identity in the lists (unique: names are unique per scope). */
export function nameKey(name: Pick<DefinedName, "name" | "scopeSheetId">): string {
  return `name:${name.scopeSheetId ?? ""}:${upper(name.name)}`;
}

/** A jump to `ranges` on `sheetId`, refused when the sheet cannot show. */
function jumpTo(model: SpreadsheetModel, sheetId: string, ranges: GridRange[]): NameBoxOutcome {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet || sheet.kind !== "grid" || ranges.length === 0) {
    return { kind: "refuse", message: NOT_A_RANGE_MESSAGE };
  }
  return sheet.visibility === "visible" ? { kind: "jump", sheetId, ranges } : { kind: "refuse", message: HIDDEN_SHEET_MESSAGE };
}

/** A name's evaluation at `host` (its scope sheet's for a local name). */
function evaluateName(model: SpreadsheetModel, name: DefinedName, host: NameHost): EvalValue {
  return evaluateDefinedName(storedValueContext(model), name, { ...host, sheetId: name.scopeSheetId ?? host.sheetId });
}

/** Where a visible name selects (see the module header). */
export function nameJump(model: SpreadsheetModel, name: DefinedName, host: NameHost): NameBoxOutcome {
  const value = evaluateName(model, name, host);
  if (!isRange(value) || value.areas.length === 0) {
    return { kind: "refuse", message: NOT_A_RANGE_MESSAGE };
  }
  const sheetId = value.areas[0]!.sheetId;
  return jumpTo(
    model,
    sheetId,
    value.areas.filter((area) => area.sheetId === sheetId).map(({ top, left, bottom, right }) => ({ top, left, bottom, right })),
  );
}

/** Where a table selects: its data body. */
export function tableJump(model: SpreadsheetModel, name: string): NameBoxOutcome | null {
  const found = findTable(model, name);
  return found ? jumpTo(model, found.sheet.id, [dataBody(found.table)]) : null;
}

/** A typed Name Box entry's meaning (see the module header). */
export function interpretNameBox(model: SpreadsheetModel, text: string, host: NameHost): NameBoxOutcome {
  const entry = text.trim();
  const target = resolveGoToTarget(model, entry, host.sheetId, host);
  if (!("error" in target)) {
    return { kind: "jump", sheetId: target.sheetId, ranges: target.ranges };
  }
  if (target.error !== GO_TO_INVALID_MESSAGE) {
    return { kind: "refuse", message: target.error };
  }
  const problem = validateName(model, entry);
  return problem ? { kind: "refuse", message: NAME_PROBLEM_MESSAGES[problem] } : { kind: "define", name: entry };
}

/** The Name Box dropdown's options (see the module header). */
export function nameBoxOptions(model: SpreadsheetModel, sheetId: string): NameBoxOption[] {
  const names = visibleNames(model).filter((name) => name.scopeSheetId === sheetId || name.scopeSheetId === undefined);
  const local = new Set(names.filter((name) => name.scopeSheetId !== undefined).map((name) => upper(name.name)));
  const options: NameBoxOption[] = names
    .filter((name) => name.scopeSheetId !== undefined || !local.has(upper(name.name)))
    .map((name) => ({ kind: "name", name: name.name, key: nameKey(name) }));
  for (const { table } of allTables(model)) {
    options.push({ kind: "table", name: table.name, key: `table:${table.id}` });
  }
  return options.toSorted(byName);
}

/** A name's evaluation as the Name Manager's Value column shows it. */
function valueText(model: SpreadsheetModel, value: EvalValue): string {
  if (isRange(value)) {
    const area = value.areas.length === 1 ? value.areas[0]! : null;
    if (!area || area.top !== area.bottom || area.left !== area.right) {
      return ARRAY_VALUE_TEXT;
    }
    const cell = storedValueContext(model).valueAt(area.sheetId, area.top, area.left);
    return isEmpty(cell) ? "" : generalText(cell);
  }
  if (isArray(value)) {
    return ARRAY_VALUE_TEXT;
  }
  return isEmpty(value) ? "" : generalText(value);
}

/** A name's formula as shown and edited: relative parts relative to the active cell. */
export function shownFormula(name: Pick<DefinedName, "formula" | "formulaUnparsed">, host: NameHost): string {
  return `=${name.formulaUnparsed ? name.formula : rebaseRelative(name.formula, host.row, host.col)}`;
}

/** The canonical formula a Refers to text stores (relative to A1), or null
 * when it does not parse. */
export function storedFormula(text: string, host: NameHost): string | null {
  const body = (text.trim().startsWith("=") ? text.trim().slice(1) : text).trim();
  if (body === "") {
    return null;
  }
  try {
    return rebaseRelative(printFormula(parseFormula(body)), -host.row, -host.col);
  } catch {
    return null;
  }
}

/** The Name Manager's rows (see the module header). */
export function nameManagerRows(model: SpreadsheetModel, host: NameHost): NameManagerRow[] {
  const sheetName = (id: string | undefined): string => (id === undefined ? WORKBOOK_SCOPE_LABEL : (model.sheets.find((sheet) => sheet.id === id)?.name ?? ""));
  return visibleNames(model)
    .map((name) => ({
      key: nameKey(name),
      name: name.name,
      value: valueText(model, evaluateName(model, name, host)),
      refersTo: shownFormula(name, host),
      scope: sheetName(name.scopeSheetId),
      ...(name.scopeSheetId === undefined ? {} : { scopeSheetId: name.scopeSheetId }),
    }))
    .toSorted((a, b) => byName(a, b) || (a.scopeSheetId === undefined ? -1 : b.scopeSheetId === undefined ? 1 : a.scope.localeCompare(b.scope)));
}
