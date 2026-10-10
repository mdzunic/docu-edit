// Sheet commands (SPEC-121 "Sheet commands") — THIS FILE is the normative
// source for creating a workbook and for the workbook-level sheet commands:
// insert, rename, duplicate, delete, move, hide/unhide and tab colour, the
// sheet-name rules, and the cross-format copy (`stripRetention`,
// `unconvertibleContent`). Framework-free: the view's useSheetCommands drives
// these through the undo stack (record → command → commit).
//
// Every command keeps SPEC-116's invariants and SPEC-117's: a formula rewrite
// builds new Cell/Row objects and a new `rows` array for the sheets it
// touches (and marks those sheets dirty); a changed sheet is replaced by a
// new object, never mutated; `model.sheets`, `opaque` and
// `odsContentEnvelope` are replaced, not spliced in place, when they change.
// Every command except setTabColor on an unchanged colour sets
// `structureDirty`. Indices are into `model.sheets` (hidden sheets
// included); the UI maps visible tab positions onto them. The active sheet
// is always visible, and each command leaves `model.activeSheetIndex` on
// the sheet its result names.
//
// - Insert: `Sheet<N>`, N one more than the largest number in any
//   `Sheet<number>` name (case-insensitive, hidden sheets included; 1 when
//   none) — Excel's numbering, not the first gap.
// - Duplicate: `<base> (<n>)` (see duplicateName); the copy shares the
//   source's rows by reference, clones its small fields, and gets its own
//   retained source at once (xlsx/duplicate.ts, ods/duplicate.ts).
// - Rename and delete rewrite every formula naming the sheet (SPEC-119's
//   renameSheetInFormula; formula/sheetRefs.ts's remapSheetReferences turns
//   a deleted sheet's references into #REF!). The retained XML keeps
//   source-time names; the writers remap it on save.
// - `alsoRecalc` lists the formula cells whose RESULT may change though
//   their text did not: 3-D references whose span gained, lost or reordered
//   a sheet, and references to a name that did not exist before.
import { relsPathFor } from "../core/xml";
import { DEFAULT_THEME, resolveColor, sameModelColor, type ModelColor } from "../core/theme";
import { cellKey, type CellKey } from "./formula/dependencies";
import { formulaForTarget } from "./formula/dialect";
import { renameSheetInFormula } from "./formula/references";
import { remapSheetReferences } from "./formula/sheetRefs";
import { parseSheetPrefix, sheetPrefixOf, tokenize } from "./formula/tokens";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "./geometry";
import {
  DEFAULT_SPREADSHEET_TITLE,
  type Cell,
  type Row,
  type Sheet,
  type SpreadsheetFormat,
  type SpreadsheetModel,
} from "./model";
import { copyOdsSheetSource, removeCopyObjects } from "./ods/duplicate";
import { SPREADSHEET_READER_VERSION } from "./open";
import { sourceNameOf } from "./sourceNames";
import { UNCONVERTIBLE_CONTENT, type ConversionSource } from "./unconvertible";
import { copyXlsxSheetParts, hasOnlySafeRelationships, removeOwnedParts, renameTablesInFormula } from "./xlsx/duplicate";
import { LazyParts } from "./xlsx/package";

export { UNCONVERTIBLE_CONTENT, type UnconvertibleRow } from "./unconvertible";

export interface SheetCommandResult {
  /** The sheet to make active. */
  activeIndex: number;
  /** Formula cells whose RESULT may change though their text did not: 3-D
   *  references spanning a sheet added/duplicated/moved/deleted, and cells
   *  naming a sheet that did not exist before (add/duplicate/rename). The
   *  pipeline adds them to the recalculation keys. */
  alsoRecalc: CellKey[];
}

// --- messages ------------------------------------------------------------------------

/** SPEC-121 "Sheet names": the message each failing rule shows. */
export const SHEET_NAME_MESSAGES = {
  blank: "A sheet name can't be blank",
  tooLong: "A sheet name can be at most 31 characters",
  characters: "A sheet name can't contain any of these characters: [ ] : * ? / \\",
  apostrophe: "A sheet name can't begin or end with an apostrophe",
  reserved: '"History" is reserved — choose another name',
} as const;

/** "There's already a sheet called "<name>"". */
export function duplicateNameMessage(name: string): string {
  return `There's already a sheet called "${name}"`;
}

export const LAST_VISIBLE_SHEET_MESSAGE = "A workbook must keep at least one visible sheet";
export const FOREIGN_SHEET_COPY_MESSAGE = "This sheet type can't be copied yet";
export const UNSAFE_SHEET_COPY_MESSAGE = "This sheet has content that can't be copied yet";

export const MAX_SHEET_NAME_LENGTH = 31;

// --- the blank workbook -----------------------------------------------------------------

/** A new, empty grid sheet — born dirty, with no source. */
export function emptySheet(name: string): Sheet {
  return {
    id: crypto.randomUUID(),
    name,
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    merges: [],
    defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
    defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
    dirty: true,
  };
}

/** A new workbook (SPEC-121 "The blank workbook"): one visible grid sheet
 * "Sheet1", `.xlsx` from birth, the 1900 date system, the single default
 * style (Calibri through the theme's minor font, 11 pt, dk1) and the Office
 * theme. No `properties`: the from-scratch core.xml carries the title and
 * `modified` only. */
export function createBlankWorkbook(): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title: DEFAULT_SPREADSHEET_TITLE,
    format: "xlsx",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets: [emptySheet("Sheet1")],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{ font: { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } }, namedStyle: "Normal" }],
    defaultStyleId: 0,
    theme: structuredClone(DEFAULT_THEME),
  };
}

// --- names -------------------------------------------------------------------------------

/** Why `name` cannot name a sheet (SPEC-121 "Sheet names", first failing rule
 * wins), or null when it can. `ownIndex` is the sheet being renamed, which
 * uniqueness ignores (so a case-only change of its own name is valid). */
export function validateSheetName(model: SpreadsheetModel, name: string, ownIndex?: number): string | null {
  if (name.trim() === "") {
    return SHEET_NAME_MESSAGES.blank;
  }
  if (name.length > MAX_SHEET_NAME_LENGTH) {
    return SHEET_NAME_MESSAGES.tooLong;
  }
  if (/[[\]:*?/\\]/.test(name)) {
    return SHEET_NAME_MESSAGES.characters;
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    return SHEET_NAME_MESSAGES.apostrophe;
  }
  const upper = name.toUpperCase();
  if (upper === "HISTORY") {
    return SHEET_NAME_MESSAGES.reserved;
  }
  if (model.sheets.some((sheet, index) => index !== ownIndex && sheet.name.toUpperCase() === upper)) {
    return duplicateNameMessage(name);
  }
  return null;
}

/** The name Insert gives a new sheet: `Sheet<N>` (see the module header). */
export function nextSheetName(model: SpreadsheetModel): string {
  let max = 0;
  for (const sheet of model.sheets) {
    const match = /^sheet(\d+)$/i.exec(sheet.name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `Sheet${max + 1}`;
}

/** The name Duplicate gives a copy of `source`: the source name less a
 * trailing ` (<digits>)` is the base; the copy is `<base> (<n>)` for the
 * smallest free n ≥ 2, the base truncated to fit 31 characters. */
export function duplicateName(model: SpreadsheetModel, source: string): string {
  const base = source.replace(/ \(\d+\)$/, "");
  const taken = new Set(model.sheets.map((sheet) => sheet.name.toUpperCase()));
  for (let n = 2; ; n += 1) {
    const suffix = ` (${n})`;
    const candidate = `${base.slice(0, Math.max(0, MAX_SHEET_NAME_LENGTH - suffix.length))}${suffix}`;
    if (!taken.has(candidate.toUpperCase())) {
      return candidate;
    }
  }
}

// --- formula plumbing ----------------------------------------------------------------------

/** A sheet with every parsed formula passed through `rewrite`: the same
 * object when nothing changed, else a new sheet (marked dirty) with new
 * rows and cells for what changed and every other row kept by identity. */
function rewriteFormulas(sheet: Sheet, rewrite: (formula: string) => string): Sheet {
  let rows: Row[] | null = null;
  sheet.rows.forEach((row, index) => {
    let cells: Cell[] | null = null;
    row.cells.forEach((cell, at) => {
      if (cell.formula === undefined || cell.formulaUnparsed) {
        return;
      }
      const next = rewrite(cell.formula);
      if (next !== cell.formula) {
        cells ??= row.cells.slice();
        cells[at] = { ...cell, formula: next };
      }
    });
    if (cells) {
      rows ??= sheet.rows.slice();
      rows[index] = { ...row, cells };
    }
  });
  return rows ? { ...sheet, rows, dirty: true } : sheet;
}

/** Every formula cell (parsed formulas only) of the model, with its key. */
function* formulaCells(model: SpreadsheetModel): Generator<{ key: CellKey; formula: string }> {
  for (const sheet of model.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.formula !== undefined && !cell.formulaUnparsed) {
          yield { key: cellKey(sheet.id, row.index, cell.col), formula: cell.formula };
        }
      }
    }
  }
}

/** The sheet names a formula's references and qualified names mention
 * (upper-cased), and its 3-D spans. */
function sheetMentions(formula: string): { names: Set<string>; spans: [string, string][] } {
  const names = new Set<string>();
  const spans: [string, string][] = [];
  if (!formula.includes("!")) {
    return { names, spans };
  }
  for (const token of tokenize(formula)) {
    if (token.kind === "ref" && token.ref?.sheet !== undefined) {
      names.add(token.ref.sheet.toUpperCase());
      if (token.ref.sheetEnd !== undefined) {
        names.add(token.ref.sheetEnd.toUpperCase());
        spans.push([token.ref.sheet.toUpperCase(), token.ref.sheetEnd.toUpperCase()]);
      }
    } else if (token.kind === "name" || token.kind === "error" || token.kind === "function") {
      const prefix = sheetPrefixOf(token.text);
      if (prefix !== "") {
        const { sheet, sheetEnd } = parseSheetPrefix(prefix);
        names.add(sheet.toUpperCase());
        if (sheetEnd !== undefined) {
          names.add(sheetEnd.toUpperCase());
        }
      }
    }
  }
  return { names, spans };
}

/** The ids a 3-D span covers in a sheet order (empty when an end is
 * missing). */
function spanIds(order: readonly { id: string; name: string }[], from: string, to: string): string {
  const a = order.findIndex((sheet) => sheet.name.toUpperCase() === from);
  const b = order.findIndex((sheet) => sheet.name.toUpperCase() === to);
  if (a < 0 || b < 0) {
    return "";
  }
  return order
    .slice(Math.min(a, b), Math.max(a, b) + 1)
    .map((sheet) => sheet.id)
    .join("\n");
}

/** `alsoRecalc` for a command: formula cells naming one of `newNames`, and
 * cells whose 3-D spans cover a different set of sheets in `after` than in
 * `before`. */
function cellsToRecalc(
  model: SpreadsheetModel,
  before: readonly { id: string; name: string }[],
  newNames: readonly string[],
): CellKey[] {
  const fresh = new Set(newNames.map((name) => name.toUpperCase()));
  const after = model.sheets.map((sheet) => ({ id: sheet.id, name: sheet.name }));
  const orderChanged = before.length !== after.length || before.some((sheet, index) => sheet.id !== after[index]?.id);
  const keys: CellKey[] = [];
  for (const { key, formula } of formulaCells(model)) {
    const { names, spans } = sheetMentions(formula);
    if ([...names].some((name) => fresh.has(name))) {
      keys.push(key);
    } else if (orderChanged && spans.some(([from, to]) => spanIds(before, from, to) !== spanIds(after, from, to))) {
      keys.push(key);
    }
  }
  return keys;
}

function orderOf(model: SpreadsheetModel): { id: string; name: string }[] {
  return model.sheets.map((sheet) => ({ id: sheet.id, name: sheet.name }));
}

function sheetAt(model: SpreadsheetModel, index: number): Sheet {
  const sheet = model.sheets[index];
  if (!sheet) {
    throw new RangeError(`No sheet at index ${index}`);
  }
  return sheet;
}

function visibleCount(sheets: readonly Sheet[]): number {
  return sheets.filter((sheet) => sheet.visibility === "visible").length;
}

/** The visible sheet to activate when the one at `index` goes (deleted or
 * hidden): the next visible to the right, else to the left — as an index
 * into `sheets` once the sheet at `index` is `removed` (or not). */
function neighbourAfter(sheets: readonly Sheet[], index: number, removed: boolean): number {
  for (let i = index + 1; i < sheets.length; i += 1) {
    if (sheets[i]!.visibility === "visible") {
      return removed ? i - 1 : i;
    }
  }
  for (let i = index - 1; i >= 0; i -= 1) {
    if (sheets[i]!.visibility === "visible") {
      return i;
    }
  }
  return 0;
}

// --- the commands ------------------------------------------------------------------------

/** Insert: an empty grid sheet right after `afterIndex`, active. */
export function addSheet(model: SpreadsheetModel, afterIndex: number): SheetCommandResult {
  const before = orderOf(model);
  const sheet = emptySheet(nextSheetName(model));
  const at = Math.min(Math.max(afterIndex + 1, 0), model.sheets.length);
  const sheets = model.sheets.slice();
  sheets.splice(at, 0, sheet);
  model.sheets = sheets;
  model.structureDirty = true;
  model.activeSheetIndex = at;
  return { activeIndex: at, alsoRecalc: cellsToRecalc(model, before, [sheet.name]) };
}

/** Rename: throws with the "Sheet names" message for an invalid name; the
 * unchanged name changes nothing. */
export function renameSheet(model: SpreadsheetModel, index: number, name: string): SheetCommandResult {
  const sheet = sheetAt(model, index);
  if (sheet.name === name) {
    // Unchanged — even a name read from a file that breaks the rules.
    return { activeIndex: model.activeSheetIndex, alsoRecalc: [] };
  }
  const error = validateSheetName(model, name, index);
  if (error !== null) {
    throw new Error(error);
  }
  const before = orderOf(model);
  model.sheets = model.sheets.map((each, at) => {
    const rewritten = rewriteFormulas(each, (formula) => renameSheetInFormula(formula, sheet.name, name));
    return at === index ? { ...rewritten, name } : rewritten;
  });
  model.structureDirty = true;
  return { activeIndex: model.activeSheetIndex, alsoRecalc: cellsToRecalc(model, before, [name]) };
}

/** Why the sheet at `index` cannot be duplicated, or null. */
export function cannotDuplicateReason(model: SpreadsheetModel, index: number): string | null {
  const sheet = sheetAt(model, index);
  if (sheet.kind === "foreign") {
    return FOREIGN_SHEET_COPY_MESSAGE;
  }
  if (model.format === "xlsx" && sheet.partName !== undefined && !hasOnlySafeRelationships(new LazyParts(model.opaque), sheet.partName)) {
    return UNSAFE_SHEET_COPY_MESSAGE;
  }
  return null;
}

/** Duplicate: a copy right after the source, active (see the module
 * header). Throws for a sheet cannotDuplicateReason refuses. */
export function duplicateSheet(model: SpreadsheetModel, index: number): SheetCommandResult {
  const source = sheetAt(model, index);
  const reason = cannotDuplicateReason(model, index);
  if (reason !== null) {
    throw new Error(reason);
  }
  const before = orderOf(model);
  const name = duplicateName(model, source.name);
  const { id: _id, partName, sourceXml, copiedFrom: _copiedFrom, dirty: _dirty, ...fields } = source;
  let copy: Sheet = structuredCloneExceptRows({ ...fields, id: crypto.randomUUID(), name, rows: source.rows });
  if (partName !== undefined || sourceXml !== undefined) {
    copy.copiedFrom = source.copiedFrom ?? sourceNameOf(model, source);
    if (source.dirty) {
      copy.dirty = true;
    }
    if (partName !== undefined && model.format === "xlsx") {
      const parts = copyXlsxSheetParts(model.opaque, partName, copy.id);
      model.opaque = parts.opaque;
      copy.partName = parts.partName;
      copy = rewriteFormulas(copy, (formula) => renameTablesInFormula(formula, parts.tableRenames));
    } else if (sourceXml !== undefined) {
      const copied = copyOdsSheetSource(model, sourceXml);
      model.opaque = copied.opaque;
      copy.sourceXml = copied.sourceXml;
    } else if (partName !== undefined) {
      copy.partName = partName;
    }
  } else {
    copy.dirty = true;
  }
  copy = rewriteFormulas(copy, (formula) => renameSheetInFormula(formula, source.name, name));
  const sheets = model.sheets.slice();
  sheets.splice(index + 1, 0, copy);
  model.sheets = sheets;
  model.structureDirty = true;
  model.activeSheetIndex = index + 1;
  return { activeIndex: index + 1, alsoRecalc: cellsToRecalc(model, before, [name]) };
}

/** A sheet's small fields cloned, its rows kept by reference. */
function structuredCloneExceptRows(sheet: Sheet): Sheet {
  const { rows, ...rest } = sheet;
  return { ...structuredClone(rest), rows };
}

/** Whether deleting the sheet should ask first (SPEC-121 "Deleting a
 * sheet"): it holds cells, is foreign, has an xlsx part with relationships
 * of its own, or an ods table with shapes or comments. */
export function sheetHasContent(model: SpreadsheetModel, index: number): boolean {
  const sheet = sheetAt(model, index);
  if (sheet.rows.length > 0 || sheet.kind === "foreign") {
    return true;
  }
  if (sheet.partName !== undefined && new LazyParts(model.opaque).has(relsPathFor(sheet.partName))) {
    return true;
  }
  return sheet.sourceXml !== undefined && (sheet.sourceXml.includes("<table:shapes") || sheet.sourceXml.includes("<office:annotation"));
}

/** Whether the sheet at `index` is the workbook's last visible one. */
export function isLastVisibleSheet(model: SpreadsheetModel, index: number): boolean {
  return sheetAt(model, index).visibility === "visible" && visibleCount(model.sheets) <= 1;
}

/** Delete: every reference to the sheet becomes #REF!; deleting the active
 * sheet activates the next visible sheet to the right (else left). Throws
 * for the last visible sheet. */
export function deleteSheet(model: SpreadsheetModel, index: number): SheetCommandResult {
  const sheet = sheetAt(model, index);
  if (isLastVisibleSheet(model, index)) {
    throw new Error(LAST_VISIBLE_SHEET_MESSAGE);
  }
  const before = orderOf(model);
  const remap = { renamed: {}, deleted: [sheet.name], orderBefore: model.sheets.map((each) => each.name) };
  const active = model.activeSheetIndex;
  const activeIndex = active === index ? neighbourAfter(model.sheets, index, true) : active > index ? active - 1 : active;
  const sheets = model.sheets.filter((_, at) => at !== index).map((each) => rewriteFormulas(each, (formula) => remapSheetReferences(formula, remap)));
  if (sheet.copiedFrom !== undefined) {
    if (sheet.partName !== undefined && model.format === "xlsx") {
      model.opaque = removeOwnedParts(model.opaque, sheet.partName);
    } else if (sheet.sourceXml !== undefined) {
      const remaining = [model.odsContentEnvelope?.prefix ?? "", model.odsContentEnvelope?.suffix ?? "", ...sheets.map((each) => each.sourceXml ?? "")];
      const opaque = removeCopyObjects(model.opaque, sheet.sourceXml, remaining);
      if (opaque !== model.opaque) {
        model.opaque = [...opaque];
      }
    }
  }
  model.sheets = sheets;
  model.structureDirty = true;
  model.activeSheetIndex = activeIndex;
  return { activeIndex, alsoRecalc: cellsToRecalc(model, before, []) };
}

/** Move: the sheet at `from` to index `to` of the resulting order; it
 * stays active. */
export function moveSheet(model: SpreadsheetModel, from: number, to: number): SheetCommandResult {
  const sheet = sheetAt(model, from);
  const target = Math.min(Math.max(to, 0), model.sheets.length - 1);
  if (target === from) {
    return { activeIndex: from, alsoRecalc: [] };
  }
  const before = orderOf(model);
  const sheets = model.sheets.filter((_, at) => at !== from);
  sheets.splice(target, 0, sheet);
  model.sheets = sheets;
  model.structureDirty = true;
  model.activeSheetIndex = target;
  return { activeIndex: target, alsoRecalc: cellsToRecalc(model, before, []) };
}

/** Hide or unhide sheets. Hiding the active sheet activates its neighbour
 * as delete does; unhiding activates the last of `indices`. Throws when
 * hiding would leave no visible sheet. */
export function setSheetHidden(model: SpreadsheetModel, indices: number[], hidden: boolean): SheetCommandResult {
  const targets = new Set(indices.filter((index) => model.sheets[index] !== undefined));
  const visibility: Sheet["visibility"] = hidden ? "hidden" : "visible";
  const sheets = model.sheets.map((sheet, at) => (targets.has(at) && sheet.visibility !== visibility ? { ...sheet, visibility } : sheet));
  if (sheets.every((sheet, at) => sheet === model.sheets[at])) {
    return { activeIndex: model.activeSheetIndex, alsoRecalc: [] };
  }
  if (visibleCount(sheets) === 0) {
    throw new Error(LAST_VISIBLE_SHEET_MESSAGE);
  }
  let activeIndex = model.activeSheetIndex;
  if (hidden) {
    if (sheets[activeIndex]?.visibility !== "visible") {
      activeIndex = neighbourAfter(sheets, activeIndex, false);
    }
  } else {
    activeIndex = Math.max(...targets);
  }
  model.sheets = sheets;
  model.structureDirty = true;
  model.activeSheetIndex = activeIndex;
  return { activeIndex, alsoRecalc: [] };
}

/** A tab colour as the tab bar paints it: hex against the workbook's theme,
 * "" for none. */
export function tabColorHex(model: SpreadsheetModel, color: ModelColor | undefined): string {
  return color === undefined ? "" : resolveColor(color, model.theme ?? DEFAULT_THEME);
}

/** Tab colour: sets or clears `tabColor` (theme references kept as
 * references); an unchanged colour changes nothing. */
export function setTabColor(model: SpreadsheetModel, index: number, color: ModelColor | undefined): void {
  const sheet = sheetAt(model, index);
  if (color === undefined ? sheet.tabColor === undefined : sheet.tabColor !== undefined && sameModelColor(sheet.tabColor, color)) {
    return;
  }
  const { tabColor: _old, ...rest } = sheet;
  const next: Sheet = color === undefined ? rest : { ...rest, tabColor: structuredClone(color) };
  model.sheets = model.sheets.map((each, at) => (at === index ? next : each));
  model.structureDirty = true;
}

// --- crossing formats -----------------------------------------------------------------------

/** The model a cross-format copy writes (SPEC-121 "Cross-format copy"),
 * given the model with its `format` already set to the target: no retained
 * package, no source on any sheet, foreign sheets gone, styles without their
 * source indices, every formula cell through formulaForTarget. Pure — the
 * open model is untouched; rows are shared except where a cell must
 * change. */
export function stripRetention(model: SpreadsheetModel): SpreadsheetModel {
  const { opaque: _opaque, odsContentEnvelope: _envelope, structureDirty: _structure, packageVariant: _variant, ...rest } = model;
  const kept = model.sheets.filter((sheet) => sheet.kind !== "foreign");
  const sheets: Sheet[] = kept.map((sheet) => {
    const { partName: _part, sourceXml: _source, copiedFrom: _copied, dirty: _dirty, ...fields } = sheet;
    return { ...fields, rows: crossRows(sheet.rows, model.format) };
  });
  if (sheets.length === 0) {
    const { dirty: _dirty, ...fallback } = emptySheet("Sheet1");
    sheets.push(fallback);
  }
  if (!sheets.some((sheet) => sheet.visibility === "visible")) {
    sheets[0] = { ...sheets[0]!, visibility: "visible" };
  }
  const active = model.sheets[model.activeSheetIndex];
  let activeSheetIndex = active ? sheets.findIndex((sheet) => sheet.id === active.id) : -1;
  if (activeSheetIndex < 0 || sheets[activeSheetIndex]!.visibility !== "visible") {
    activeSheetIndex = sheets.findIndex((sheet) => sheet.visibility === "visible");
  }
  return {
    ...rest,
    opaque: [],
    sheets,
    activeSheetIndex,
    styles: model.styles.map((style) => {
      const { sourceXf: _xf, sourceStyleName: _name, ...styleRest } = style;
      return styleRest;
    }),
  };
}

/** Whether a cell carries content only its source format can carry
 * (SPEC-122): preserved attributes, children or text — an ods
 * `table:content-validation-name` means nothing in an xlsx `<c>`, nor an
 * xlsx `vm` in a `table:table-cell`. (An xlsx `sstIndex` crosses harmlessly:
 * only the xlsx writer reads it, and a copy never reaches the xlsx writer
 * from an xlsx source.) */
function hasFormatBoundContent(cell: Cell): boolean {
  return cell.preservedAttrs !== undefined || cell.preservedXml !== undefined || cell.preservedText !== undefined;
}

/** A sheet's rows as they cross to `target` — the same array when no
 * formula cell must change: formula cells through formulaForTarget, and
 * format-bound preserved content (SPEC-122) dropped. */
function crossRows(rows: Row[], target: SpreadsheetFormat): Row[] {
  let out: Row[] | null = null;
  rows.forEach((row, index) => {
    let cells: Cell[] | null = null;
    row.cells.forEach((cell, at) => {
      let crossed = cell;
      if (cell.formula !== undefined || cell.dataTableXml !== undefined || cell.cm !== undefined) {
        crossed = formulaForTarget(cell, target).cell;
      }
      if (hasFormatBoundContent(crossed)) {
        const { preservedAttrs: _attrs, preservedXml: _xml, preservedText: _text, ...rest } = crossed;
        crossed = rest;
      }
      if (crossed !== cell) {
        cells ??= row.cells.slice();
        cells[at] = crossed;
      }
    });
    if (cells) {
      out ??= rows.slice();
      out[index] = { ...row, cells };
    }
  });
  return out ?? rows;
}

/** The labels of UNCONVERTIBLE_CONTENT whose marker the source holds, in
 * table order — what a copy to `target` leaves out. Empty for the model's
 * own format. */
export function unconvertibleContent(model: SpreadsheetModel, target: SpreadsheetFormat): string[] {
  if (target === model.format) {
    return [];
  }
  const source: ConversionSource = { model, parts: new LazyParts(model.opaque) };
  const labels: string[] = [];
  for (const row of UNCONVERTIBLE_CONTENT) {
    const marker = model.format === "xlsx" ? row.xlsx : row.ods;
    const found = marker?.(source) ?? false;
    if (found === true) {
      labels.push(row.label);
    } else if (Array.isArray(found)) {
      labels.push(...found);
    }
  }
  return labels;
}
