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
//   retained source at once (xlsx/duplicate.ts, ods/duplicate.ts). SPEC-128:
//   the clones include the autofilter (its criteria, the rows' filterHidden
//   riding in the shared rows) and the sort records — the autofilter without
//   its `odsName` (the database range is the source sheet's; the writer
//   names the copy's) and, in an .ods, without its `sourceXml`, which names
//   the source sheet in its target address.
// - Rename and delete rewrite every formula naming the sheet (SPEC-119's
//   renameSheetInFormula; formula/sheetRefs.ts's remapSheetReferences turns
//   a deleted sheet's references into #REF!). The retained XML keeps
//   source-time names; the writers remap it on save. SPEC-129: every
//   conditional-format rule formula on every sheet (cfFormulas.ts's table —
//   value objects and a preserved xlsx rule's texts included) goes through
//   remapSheetReferences the same way; a changed format is replaced without
//   its sourceXml (a dirty sheet then writes it from the model). SPEC-130:
//   so does every data-validation rule's formula1/formula2 on every sheet
//   (mapValidationsFormulas — a changed rule without its sourceXml, its x14
//   placement re-derived for its sheet's name once the command is done).
//   Unlike a cell-formula rewrite neither leaves the sheet dirty: its
//   retained rule text goes through the writers' sheet-reference pass, as a
//   hyperlink's does.
// - SPEC-129, duplicate: the copy's conditional formats get fresh format and
//   rule ids (x14Ids kept — a GUID pairs only within one worksheet), their
//   formulas renamed as the copy's cells' are, and in an .ods they lose
//   their sourceXml, which names the source sheet in its target address.
// - SPEC-130, duplicate: the copy's data-validation rules get fresh ids and
//   no `odsName` (the content-validation is the source sheet's; the writer
//   names the copy's), their formulas renamed as the copy's cells' are and
//   their x14 placement derived for the copy's name; in an .ods they lose
//   their sourceXml, whose base-cell address names the source sheet.
// - SPEC-131, duplicate: the copy keeps its origin's drawings (their `id`s
//   are unique per sheet, so the copy keeps them too), written into the
//   drawing part SPEC-121's closure copy names for the copy (xlsx: the
//   copy's `drawingPart`); in an .ods each drawing's source XML takes the
//   same id suffixes and object retargets as the copy's table text.
// - SPEC-133: rename and delete apply their SheetRemap through
//   remapSheetReferences to every name's formula (`namesDirty`) and every
//   table's calculated and totals formulas (the table `dirty`); delete also
//   drops the sheet's local names and its tables, their references becoming
//   #REF! everywhere. Duplicate copies the origin's tables on the model —
//   fresh ids and workbook-unique xlsxIds (max + 1), `uniqueTableName`
//   names (`<name>_<n>`), `dirty` — with the copy's cell, calculated and
//   totals formulas renamed to them (renameTableInFormula, instead of the
//   xlsx closure copy's renames: the model holds the tables) and the copy
//   `dirty`, so its `<tableParts>` is rebuilt. A copied table's `partName`
//   is the part SPEC-121's closure copy made of its origin's (none without
//   one), so the writer regenerates that part — keeping its unmodelled XML
//   (`dataDxfId`, `xr3:uid`, …) — rather than dropping it for a golden one.
//   It also clones the origin's local names with the copy's scope, their
//   formulas renamed as the copy's cells' are, each placed right after its
//   origin (SPEC-121's order) — but a `_xlnm._FilterDatabase`, which goes
//   first, where SPEC-128 writes a sheet's new one. A cross-format copy
//   keeps names and tables, each table without its partName (the package is
//   gone) and `dirty`.
// - `alsoRecalc` lists the formula cells whose RESULT may change though
//   their text did not: 3-D references whose span gained, lost or reordered
//   a sheet, and references to a name that did not exist before.
import { ensureMediaEncoded, type MediaPart } from "../core/media";
import { relsPathFor } from "../core/xml";
import { nextMediaId } from "./drawings";
import { DEFAULT_THEME, resolveColor, sameModelColor, type ModelColor } from "../core/theme";
import { mapFormatsFormulas, withoutFormatSource } from "./cfFormulas";
import { mapNameFormulas, mapTableFormulas, removeTableReferencesEverywhere } from "./nameTableAdjusters";
import { RESERVED_PREFIX, uniqueTableName } from "./names";
import { allTables, renameTableInFormula } from "./tables";
import { CHART_SHEET_DUPLICATE_TITLE, chartContentOf, isModelledChart, remapChartReferences, syncChartNotifyRanges } from "./charts";
import { mintCfId, renumberPriorities, sheetFormats, withX14Placement, type CfRule, type ConditionalFormat } from "./conditionalFormats";
import type { AutoFilter } from "./filter";
import { cellKey, type CellKey } from "./formula/dependencies";
import { formulaForTarget } from "./formula/dialect";
import { renameSheetInFormula } from "./formula/references";
import { remapSheetReferences, type SheetRemap } from "./formula/sheetRefs";
import { parseSheetPrefix, sheetPrefixOf, tokenize } from "./formula/tokens";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "./geometry";
import {
  DEFAULT_SPREADSHEET_TITLE,
  isCsvFormat,
  type Cell,
  type CellThread,
  type DefinedName,
  type DrawingAnchor,
  type Row,
  type Sheet,
  type SheetChartRefs,
  type SheetDrawing,
  type SheetTable,
  type SpreadsheetFormat,
  type SpreadsheetModel,
} from "./model";
import { crossAnnotations, remapPlaceSheets } from "./annotations";
import { copyOdsSheetSource, removeCopyObjects } from "./ods/duplicate";
import { objectsReferenced } from "./ods/package";
import { SPREADSHEET_READER_VERSION } from "./open";
import { sourceNameOf } from "./sourceNames";
import { UNCONVERTIBLE_CONTENT, type ConversionSource } from "./unconvertible";
import {
  mapValidationsFormulas,
  mintValidationId,
  sheetValidations,
  withoutValidationSource,
  withX14Placement as placeValidation,
  type DataValidation,
} from "./validation";
import { copyXlsxSheetParts, hasOnlySafeRelationships, mintedId, removeOwnedParts, renameTablesInFormula } from "./xlsx/duplicate";
import { LazyParts, relationshipKind, relationshipsOf, targetPart } from "./xlsx/package";

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
    conditionalFormats: [],
    validations: [],
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
    dxfs: [],
    names: [],
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

/** A sheet with every conditional-format rule formula (cfFormulas.ts) and
 * every data-validation rule formula passed through `rewrite`: the same
 * object when nothing changed, else a new sheet whose changed formats and
 * rules lost their sourceXml — its dirty flag untouched (see the header).
 * `sheetName`: the sheet's name once the command is done (x14 placement). */
function rewriteRuleFormulas(format: SpreadsheetFormat, sheet: Sheet, rewrite: (formula: string) => string, sheetName = sheet.name): Sheet {
  const formats = sheetFormats(sheet);
  const nextFormats = mapFormatsFormulas(formats, rewrite, { format, sheetName });
  const validations = sheetValidations(sheet);
  const nextValidations = mapValidationsFormulas(validations, rewrite, format, sheetName);
  if (nextFormats === formats && nextValidations === validations) {
    return sheet;
  }
  const next: Sheet = { ...sheet };
  if (nextFormats !== formats) {
    next.conditionalFormats = nextFormats;
  }
  if (nextValidations !== validations) {
    next.validations = nextValidations;
  }
  return next;
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
  const remap = { renamed: { [sheet.name]: name }, deleted: [], orderBefore: model.sheets.map((each) => each.name) };
  model.sheets = model.sheets.map((each, at) => {
    const rewritten = rewriteFormulas(each, (formula) => renameSheetInFormula(formula, sheet.name, name));
    const remapped = rewriteRuleFormulas(model.format, rewritten, (formula) => remapSheetReferences(formula, remap), at === index ? name : each.name);
    return at === index ? { ...remapped, name } : remapped;
  });
  remapChartReferences(model, remap);
  remapNamesAndTables(model, remap);
  model.structureDirty = true;
  return { activeIndex: model.activeSheetIndex, alsoRecalc: cellsToRecalc(model, before, [name]) };
}

/** SPEC-133: a rename's or delete's remap over every name's formula and
 * every table's calculated and totals formulas. */
function remapNamesAndTables(model: SpreadsheetModel, remap: SheetRemap): void {
  mapNameFormulas(model, (defined) => remapSheetReferences(defined.formula, remap));
  mapTableFormulas(model, (text) => remapSheetReferences(text, remap));
  // SPEC-135: every hyperlink's place location, beside Cell.formula.
  remapPlaceSheets(model, remap);
}

/** Why the sheet at `index` cannot be duplicated, or null. */
export function cannotDuplicateReason(model: SpreadsheetModel, index: number): string | null {
  const sheet = sheetAt(model, index);
  if (sheet.kind === "chart") {
    return CHART_SHEET_DUPLICATE_TITLE;
  }
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
  if (copy.autoFilter) {
    copy.autoFilter = copiedAutoFilter(copy.autoFilter, model.format);
  }
  copy.conditionalFormats = copiedFormats(sheetFormats(copy), model.format);
  copy.validations = copiedValidations(sheetValidations(copy), model.format, name);
  if (copy.threads?.length) {
    // SPEC-135: a thread id is unique in the workbook — the copy's threads
    // get fresh ones, and its comment parts are rebuilt from the model.
    copy.threads = copiedThreads(copy.threads, copy.id);
    copy.annotationsDirty = true;
  }
  let partRenames: ReadonlyMap<string, string> | undefined;
  if (partName !== undefined || sourceXml !== undefined) {
    copy.copiedFrom = source.copiedFrom ?? sourceNameOf(model, source);
    if (source.dirty) {
      copy.dirty = true;
    }
    if (partName !== undefined && model.format === "xlsx") {
      const parts = copyXlsxSheetParts(model.opaque, partName, copy.id);
      model.opaque = parts.opaque;
      copy.partName = parts.partName;
      partRenames = parts.partRenames;
      copy = withCopiedDrawingPart(copy, new LazyParts(model.opaque));
      copy = withCopiedChartParts(copy, (part) => parts.partRenames.get(part));
      if (!copy.tables?.length) {
        // SPEC-133: tables the model holds are renamed on the model below.
        copy = rewriteFormulas(copy, (formula) => renameTablesInFormula(formula, parts.tableRenames));
      }
    } else if (sourceXml !== undefined) {
      const copied = copyOdsSheetSource(model, sourceXml, (copy.drawings ?? []).map((drawing) => drawing.sourceXml ?? ""));
      model.opaque = copied.opaque;
      copy.sourceXml = copied.sourceXml;
      if (copy.drawings) {
        copy.drawings = copy.drawings.map((drawing, at) => (drawing.sourceXml === undefined ? drawing : { ...drawing, sourceXml: copied.drawingTexts[at]! }));
        copy = withCopiedChartParts(copy, (_part, drawing) => {
          const object = drawing.sourceXml === undefined ? undefined : objectsReferenced(drawing.sourceXml)[0];
          return object === undefined ? undefined : `${object}/content.xml`;
        });
      }
    } else if (partName !== undefined) {
      copy.partName = partName;
    }
  } else {
    copy.dirty = true;
  }
  copy = rewriteFormulas(copy, (formula) => renameSheetInFormula(formula, source.name, name));
  copy = rewriteRuleFormulas(model.format, copy, (formula) => renameSheetInFormula(formula, source.name, name));
  copy = withRenamedChartRefs(copy, (formula) => renameSheetInFormula(formula, source.name, name));
  const renameTables = copiedTables(model, copy, partRenames);
  if (copy.tables?.length) {
    copy = rewriteFormulas(copy, renameTables.rename);
    copy.tables = renameTables.tables;
    copy.dirty = true;
  }
  if (model.names.some((defined) => defined.scopeSheetId === source.id)) {
    const rewrite = (formula: string): string => renameTables.rename(renameSheetInFormula(formula, source.name, name));
    const filterDatabase = `${RESERVED_PREFIX}_FilterDatabase`.toUpperCase();
    const first: DefinedName[] = [];
    const rest: DefinedName[] = [];
    for (const defined of model.names) {
      rest.push(defined);
      if (defined.scopeSheetId !== source.id) {
        continue;
      }
      const cloned = defined.formulaUnparsed ? { ...defined, scopeSheetId: copy.id } : { ...defined, scopeSheetId: copy.id, formula: rewrite(defined.formula) };
      (defined.name.toUpperCase() === filterDatabase ? first : rest).push(cloned);
    }
    model.names = [...first, ...rest];
    model.namesDirty = true;
  }
  const sheets = model.sheets.slice();
  sheets.splice(index + 1, 0, copy);
  model.sheets = sheets;
  syncChartNotifyRanges(model);
  model.structureDirty = true;
  model.activeSheetIndex = index + 1;
  return { activeIndex: index + 1, alsoRecalc: cellsToRecalc(model, before, [name]) };
}

/** A copy whose drawings live in the drawing part its closure copy made
 * (SPEC-131): `drawingPart` names the part the copy's own `<drawing>`
 * relationship targets — none when it has none. */
function withCopiedDrawingPart(copy: Sheet, parts: LazyParts): Sheet {
  if (copy.drawingPart === undefined || copy.partName === undefined) {
    return copy;
  }
  const rel = relationshipsOf(parts, copy.partName).find((each) => relationshipKind(each.type) === "drawing" && !each.external);
  const drawingPart = rel ? targetPart(copy.partName, rel) : undefined;
  const next = { ...copy };
  if (drawingPart === undefined) {
    delete next.drawingPart;
  } else {
    next.drawingPart = drawingPart;
  }
  return next;
}

/** SPEC-132: a copy whose chart drawings name the chart parts its closure
 * copy made — `copied(sourcePart, drawing)` answers the copy's own part
 * (xlsx: the closure copy's rename; ods: the object the drawing's
 * retargeted frame shows). The chart write paths key on `sourcePart`, so a
 * copy left naming its origin's part would patch or regenerate the
 * ORIGIN's chart and leave its own stale. */
function withCopiedChartParts(copy: Sheet, copied: (sourcePart: string, drawing: SheetDrawing) => string | undefined): Sheet {
  if (!copy.drawings?.some((drawing) => chartContentOf(drawing)?.sourcePart !== undefined)) {
    return copy;
  }
  const drawings = copy.drawings.map((drawing) => {
    const content = chartContentOf(drawing);
    const own = content?.sourcePart === undefined ? undefined : copied(content.sourcePart, drawing);
    return content && own !== undefined && own !== content.sourcePart ? { ...drawing, content: { ...content, sourcePart: own } } : drawing;
  });
  return { ...copy, drawings };
}

/** SPEC-132: a copy's chart references renamed as its cells' formulas are
 * (Excel's copy plots the copy's own data), `refsDirty` where one changed. */
function withRenamedChartRefs(copy: Sheet, rename: (formula: string) => string): Sheet {
  if (!copy.drawings?.some((drawing) => chartContentOf(drawing)?.refs !== undefined)) {
    return copy;
  }
  const drawings = copy.drawings.map((drawing) => {
    const content = chartContentOf(drawing);
    if (!content?.refs) {
      return drawing;
    }
    const refs = renamedChartRefs(content.refs, rename);
    return refs === content.refs ? drawing : { ...drawing, content: { ...content, refs, refsDirty: true as const } };
  });
  return { ...copy, drawings };
}

/** Chart references through `rename` — the same object when none changes. */
function renamedChartRefs(refs: SheetChartRefs, rename: (formula: string) => string): SheetChartRefs {
  let changed = false;
  const each = (text: string | undefined): string | undefined => {
    if (text === undefined) {
      return undefined;
    }
    const next = rename(text);
    changed ||= next !== text;
    return next;
  };
  const series = refs.series.map((ref) => {
    const next: SheetChartRefs["series"][number] = {};
    const name = each(ref.name);
    const categories = each(ref.categories);
    const values = each(ref.values);
    if (name !== undefined) {
      next.name = name;
    }
    if (categories !== undefined) {
      next.categories = categories;
    }
    if (values !== undefined) {
      next.values = values;
    }
    return next;
  });
  const plotArea = each(refs.plotArea);
  if (!changed) {
    return refs;
  }
  return plotArea === undefined ? { series, plotVisibleOnly: refs.plotVisibleOnly } : { series, plotArea, plotVisibleOnly: refs.plotVisibleOnly };
}

/** SPEC-133: a copy's tables (see the module header) and the rewrite of a
 * formula's references to the origin's tables into the copies'. */
function copiedTables(model: SpreadsheetModel, copy: Sheet, partRenames: ReadonlyMap<string, string> | undefined): { tables: SheetTable[]; rename: (formula: string) => string } {
  const origin = copy.tables ?? [];
  if (origin.length === 0) {
    return { tables: [], rename: (formula) => formula };
  }
  let xlsxId = Math.max(0, ...allTables(model).map(({ table }) => table.xlsxId ?? 0));
  const taken: string[] = [];
  const renames: [string, string][] = [];
  const tables = origin.map((table): SheetTable => {
    const name = uniqueTableName(model, table.name, taken);
    taken.push(name);
    renames.push([table.name, name]);
    xlsxId += 1;
    const { partName, ...rest } = table;
    const copiedPart = partName === undefined ? undefined : partRenames?.get(partName);
    return { ...rest, id: crypto.randomUUID(), name, xlsxId, dirty: true, ...(copiedPart === undefined ? {} : { partName: copiedPart }) };
  });
  const rename = (formula: string): string => renames.reduce((text, [from, to]) => renameTableInFormula(text, from, to), formula);
  const renamed = tables.map((table) => ({
    ...table,
    columns: table.columns.map((column) => ({
      ...column,
      ...(column.calculatedFormula === undefined ? {} : { calculatedFormula: rename(column.calculatedFormula) }),
      ...(column.totalsFormula === undefined ? {} : { totalsFormula: rename(column.totalsFormula) }),
    })),
  }));
  return { tables: renamed, rename };
}

/** A copy's conditional formats (see the module header). */
function copiedFormats(formats: readonly ConditionalFormat[], format: SpreadsheetFormat): ConditionalFormat[] {
  return formats.map((each) => {
    const copy: ConditionalFormat = { ...each, id: mintCfId(), rules: each.rules.map((rule) => ({ ...rule, id: mintCfId() })) };
    return format === "ods" ? withoutFormatSource(copy) : copy;
  });
}

/** A copy's data-validation rules (see the module header) — before its
 * formulas are renamed, which re-derives a renamed rule's placement. A rule
 * still carrying its source keeps the location that source has. */
function copiedValidations(rules: readonly DataValidation[], format: SpreadsheetFormat, copyName: string): DataValidation[] {
  return rules.map((each) => {
    const { odsName: _name, ...rest } = each;
    const copy: DataValidation = { ...rest, id: mintValidationId() };
    const own = format === "ods" ? withoutValidationSource(copy) : copy;
    return own.sourceXml === undefined ? placeValidation(own, format, copyName) : own;
  });
}

/** SPEC-135: a copy's threads, without their source, every thread and
 * message id re-minted as xlsx/duplicate.ts re-mints the copied threaded
 * part's (mintedId of the copy's sheet id and the original id), so the model
 * and the copy's parts — its placeholders' `tc={id}` authors included —
 * name the same threads. */
function copiedThreads(threads: readonly CellThread[], copySheetId: string): CellThread[] {
  return threads.map((thread) => {
    const messages = thread.messages.map(({ source: _source, ...message }) => ({ ...message, messageId: mintedId(copySheetId, message.messageId) }));
    return { ...thread, threadId: messages[0]?.messageId ?? mintedId(copySheetId, thread.threadId), messages };
  });
}

/** A copy's autofilter (see the module header). */
function copiedAutoFilter(filter: AutoFilter, format: SpreadsheetFormat): AutoFilter {
  const { odsName: _name, sourceXml, ...rest } = filter;
  return format === "ods" || sourceXml === undefined ? rest : { ...rest, sourceXml };
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
  const remapped = (formula: string): string => remapSheetReferences(formula, remap);
  const sheets = model.sheets.filter((_, at) => at !== index).map((each) => rewriteRuleFormulas(model.format, rewriteFormulas(each, remapped), remapped));
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
  remapChartReferences(model, remap);
  const names = model.names.filter((defined) => defined.scopeSheetId !== sheet.id);
  if (names.length !== model.names.length) {
    model.names = names;
    model.namesDirty = true;
  }
  remapNamesAndTables(model, remap);
  removeTableReferencesEverywhere(
    model,
    (sheet.tables ?? []).map((table) => table.name),
  );
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
 * package, no source on any sheet, no CSV dialect (SPEC-126), foreign sheets
 * gone, styles without their source indices, every formula cell through
 * formulaForTarget, conditional formats per crossFormats and differential
 * styles without their sourceXml/sourceStyleName (SPEC-129), data-validation
 * rules per crossValidations and no `odsOrphanValidations` — the source's
 * unreferenced content-validations are its own (SPEC-130), drawings per
 * crossDrawings with `media` holding only what they show, re-pathed, and no
 * `drawingPart`/`drawingsDirty` (SPEC-131), every table without its
 * `partName` and `dirty`, and no name kept unparsed (SPEC-133). Pure — the open
 * model is untouched; rows are shared except where a cell must change. */
export function stripRetention(model: SpreadsheetModel): SpreadsheetModel {
  const {
    opaque: _opaque,
    odsContentEnvelope: _envelope,
    structureDirty: _structure,
    packageVariant: _variant,
    csv: _csv,
    odsOrphanValidations: _orphans,
    persons: _persons,
    ...rest
  } = model;
  const kept = model.sheets.filter((sheet) => sheet.kind !== "foreign");
  const media = new CrossedMedia(model);
  const sheets: Sheet[] = kept.map((sheet) => {
    const {
      partName: _part,
      sourceXml: _source,
      copiedFrom: _copied,
      dirty: _dirty,
      drawings,
      drawingPart: _drawingPart,
      drawingsDirty: _drawingsDirty,
      notes: _notes,
      threads: _threads,
      hyperlinks: _hyperlinks,
      annotationsDirty: _annotationsDirty,
      pageSetupDirty: _pageSetupDirty,
      ...fields
    } = sheet;
    const crossed: Sheet = {
      ...fields,
      rows: crossRows(sheet.rows, model.format),
      conditionalFormats: crossFormats(sheetFormats(sheet), model.format, sheet.name),
      validations: crossValidations(sheetValidations(sheet), model.format, sheet.name),
      // SPEC-135: notes, threads and links cross field for field, without their sources.
      ...crossAnnotations(sheet, model.format),
    };
    if (fields.tables) {
      // SPEC-133: the table parts stay behind with the package.
      crossed.tables = fields.tables.map(({ partName: _tablePart, ...table }) => ({ ...table, dirty: true as const }));
    }
    const crossedDrawings = crossDrawings(drawings ?? [], media);
    if (sheet.kind === "chart" && model.format !== "xlsx") {
      // SPEC-132: a chart sheet becomes a grid sheet holding its chart at the top-left.
      crossed.kind = "grid";
      crossed.drawings = crossedDrawings.map((drawing) => ({ ...drawing, anchor: topLeftAnchor(drawing.anchor) }));
      if (crossed.drawings.length === 0) {
        delete crossed.drawings;
      }
      return crossed;
    }
    if (crossedDrawings.length > 0) {
      crossed.drawings = crossedDrawings;
    }
    return crossed;
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
  const { media: _media, ...crossedRest } = rest;
  // SPEC-135: the thread authors cross to an .xlsx, without their source.
  const persons = model.format === "xlsx" ? (model.persons ?? []).map(({ source: _personSource, ...person }) => person) : [];
  return {
    ...crossedRest,
    ...(media.parts.length > 0 ? { media: media.parts } : {}),
    ...(persons.length > 0 ? { persons } : {}),
    opaque: [],
    sheets,
    activeSheetIndex,
    // SPEC-133: a name kept unparsed is text in the source's dialect, which
    // the target cannot read (formulaForTarget drops such a cell formula too).
    names: (model.names ?? []).filter((name) => !name.formulaUnparsed),
    styles: model.styles.map((style) => {
      const { sourceXf: _xf, sourceStyleName: _name, ...styleRest } = style;
      return styleRest;
    }),
    dxfs: model.dxfs.map((dxf) => {
      const { sourceXml: _xml, sourceStyleName: _name, ...dxfRest } = dxf;
      return dxfRest;
    }),
  };
}

/** SPEC-131: the images crossing drawings show, re-pathed into the target's
 * media folder (nextMediaId's naming — `xl/media/image{n}` or
 * `Pictures/image{n}`), each part once however many drawings show it. */
class CrossedMedia {
  readonly parts: MediaPart[] = [];
  private readonly renamed = new Map<string, string>();
  private readonly source: ReadonlyMap<string, MediaPart>;

  constructor(private readonly model: SpreadsheetModel) {
    this.source = new Map((model.media ?? []).map((part) => [part.mediaId, part] as const));
  }

  /** The target path of a source image, or undefined when the model holds
   * no such part. */
  crossed(mediaId: string): string | undefined {
    const known = this.renamed.get(mediaId);
    if (known !== undefined) {
      return known;
    }
    const part = this.source.get(mediaId);
    if (!part) {
      return undefined;
    }
    ensureMediaEncoded(part);
    const next = nextMediaId({ format: this.model.format, opaque: [], media: this.parts }, part.contentType);
    this.parts.push({ mediaId: next, contentType: part.contentType, data: part.data });
    this.renamed.set(mediaId, next);
    return next;
  }
}

/** A sheet's drawings as they cross (SPEC-131 "Cross-format"): image and
 * shape drawings through the target's canonical writers — anchor and
 * behaviour kept, `sourceXml` dropped, an image re-pathed (CrossedMedia);
 * SPEC-132: a catalogue chart through the target's regenerate path — its
 * visual and canonical references kept (the target's writer converts them:
 * toXlsxFormula / canonicalToOdsRangeAddress), its `sourcePart` and write
 * flags dropped, so the target mints a new part; `foreign` (a chartex chart
 * among them), `group`, out-of-catalogue and read-only `chart` drawings, and
 * an image whose bytes the model does not hold, dropped (unconvertible.ts's
 * drawings row names each one). */
function crossDrawings(drawings: readonly SheetDrawing[], media: CrossedMedia): SheetDrawing[] {
  return drawings.flatMap((drawing): SheetDrawing[] => {
    const { sourceXml: _source, ...rest } = drawing;
    if (drawing.content.kind === "shape") {
      return [rest];
    }
    if (drawing.content.kind === "chart") {
      const { visual, refs, readOnlyReason } = drawing.content;
      if (!isModelledChart(drawing.content) || readOnlyReason !== undefined) {
        return [];
      }
      return [{ ...rest, content: refs ? { kind: "chart", visual, refs } : { kind: "chart", visual } }];
    }
    if (drawing.content.kind !== "image") {
      return [];
    }
    const mediaId = media.crossed(drawing.content.mediaId);
    return mediaId === undefined ? [] : [{ ...rest, content: { ...drawing.content, mediaId } }];
  });
}

/** SPEC-132: where a chart sheet's chart lands on the grid sheet an .ods
 * copy makes of it — the top-left cell, at the chart's own size. */
function topLeftAnchor(anchor: DrawingAnchor): DrawingAnchor {
  const origin = { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 };
  return { kind: "oneCell", from: origin, extEmu: anchor.extEmu ?? { cx: CHART_SHEET_FALLBACK_EMU.cx, cy: CHART_SHEET_FALLBACK_EMU.cy } };
}

/** A chart sheet's chart with no extent of its own: 12 × 7.5 cm (SPEC-052's new chart). */
const CHART_SHEET_FALLBACK_EMU = { cx: 4320000, cy: 2700000 };

/** A sheet's conditional formats as they cross to `target` (SPEC-129): each
 * without its sourceXml; each rule without its x14 GUID and placed for the
 * target (withX14Placement); a `preserved` rule — another format's XML —
 * dropped, a format left with no rule removed and the priorities renumbered;
 * to an .ods, an icon set's `reverse` dropped (it has no wire form there).
 * unconvertible.ts's "some conditional formatting" row names what is lost. */
function crossFormats(formats: readonly ConditionalFormat[], target: SpreadsheetFormat, sheetName: string): ConditionalFormat[] {
  let dropped = false;
  const crossed = formats.flatMap((format) => {
    const rules = format.rules.flatMap((rule): CfRule[] => {
      if (rule.kind === "preserved") {
        dropped = true;
        return [];
      }
      let next: CfRule = rule;
      if (next.x14Id !== undefined) {
        const { x14Id: _id, ...rest } = next;
        next = rest as CfRule;
      }
      if (target === "ods" && next.kind === "iconSet" && next.icons.reverse) {
        next = { ...next, icons: { ...next.icons, reverse: false } };
      }
      return [withX14Placement(next, target, sheetName)];
    });
    return rules.length === 0 ? [] : [{ ...withoutFormatSource(format), rules }];
  });
  return dropped ? renumberPriorities(crossed) : crossed;
}

/** A sheet's data-validation rules as they cross to `target` (SPEC-130
 * "Round trip"): every type crosses through the model — each rule without
 * its sourceXml and odsName (the writer serialises and names it afresh),
 * placed for the target (validation.ts's withX14Placement: an x14 rule
 * becomes an ordinary ODS one, a cross-sheet one goes to an .xlsx copy's
 * x14 ext), and — to anything but .ods — without `odsListOrder`, the one
 * property that does not cross (an ODS sort-ascending list becomes an
 * ordinary xlsx list). */
function crossValidations(rules: readonly DataValidation[], target: SpreadsheetFormat, sheetName: string): DataValidation[] {
  return rules.map((rule) => {
    const { sourceXml: _source, odsName: _name, odsListOrder, ...rest } = rule;
    const crossed: DataValidation = target === "ods" && odsListOrder !== undefined ? { ...rest, odsListOrder } : rest;
    return placeValidation(crossed, target, sheetName);
  });
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
 * own format, and for a csv/tsv source (SPEC-126: a delimited-text file
 * carries no package content to leave out). */
export function unconvertibleContent(model: SpreadsheetModel, target: SpreadsheetFormat): string[] {
  if (target === model.format || isCsvFormat(model.format)) {
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
