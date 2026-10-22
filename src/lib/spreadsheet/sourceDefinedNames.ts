// Defined names of the retained source file (SPEC-130 "List sources") —
// provisional until SPEC-133 models names. SPEC-120's evaluator finds every
// name uncomputable until SPEC-133 adds its `resolveName` hook, so a list
// rule whose Source is a name (`Colours`) resolves it here, to the text the
// source file defines for it, and evaluates that. Nothing else looks a name
// up here; SPEC-133's names model replaces this module.
//
// - xlsx: the retained workbook part's `definedName` elements, their text
//   through fromXlsxFormula. One scoped to a sheet (`localSheetId`) answers
//   only there, before a workbook-scoped one. Relative references are
//   relative to A1.
// - ods: content.xml's `table:named-range` (`table:cell-range-address`) and
//   `table:named-expression` (`table:expression`, fromOdsFormula) — a
//   table's own `table:named-expressions` (in its `sourceXml`) before the
//   spreadsheet's (the content envelope's suffix). Relative parts are
//   relative to `table:base-cell-address`.
//
// The text's sheet names stay the source file's, so it goes through
// SPEC-121's sourceSheetRemap. SPEC-133 retired SPEC-124's retained rows for
// name text: the model's names follow structural edits and moves, so an entry
// the model holds (same name and scope) answers with the model's formula,
// rebased to the entry's origin; the text answers only for the rest. Names
// match case-insensitively. A name the source does not define (a workbook
// made here, a name the source never had) resolves to nothing.
import type { CellRef } from "./address";
import { fromOdsFormula, fromXlsxFormula } from "./formula/dialect";
import { FormulaSyntaxError } from "./formula/parse";
import { remapSheetReferences } from "./formula/sheetRefs";
import type { Sheet, SpreadsheetModel } from "./model";
import { rebaseRelative } from "./names";
import { odsRangeAddressToCanonical, parseOdsRangeAddress } from "./ods/rangeAddress";
import { resolveRetainedSheetName, sourcePartsOf, sourceSheetRemap } from "./sourceNames";
import { partTextOf, readSourceWorkbook, workbookPathOf, type PartLookup } from "./xlsx/package";
import { attributeValue, decodeXmlText, startTagAt } from "./xmlSplice";

export interface SourceDefinedName {
  /** Canonical, in the model's sheet names. */
  formula: string;
  /** The cell its relative references are relative to. */
  origin: CellRef;
}

interface NameEntry {
  name: string;
  /** xlsx: the source name of the sheet `localSheetId` scopes it to. */
  scope?: string;
  /** Canonical, in the source's sheet names. */
  formula: string;
  origin: CellRef;
}

const A1: CellRef = { row: 0, col: 0 };

/** Canonical text of source formula text, or null when it does not read. */
function canonicalOf(read: () => string | null): string | null {
  try {
    return read();
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return null;
    }
    throw error;
  }
}

/** The retained workbook part's `definedName`s. */
function xlsxEntries(parts: PartLookup): NameEntry[] {
  const path = workbookPathOf(parts);
  const xml = path === undefined ? undefined : partTextOf(parts, path);
  if (xml === undefined || !xml.includes("definedName")) {
    return [];
  }
  const sheets = readSourceWorkbook(parts)?.sheets ?? [];
  const entries: NameEntry[] = [];
  for (const match of xml.matchAll(/<((?:[\w.-]+:)?definedName)(?=[\s/>])/g)) {
    const tag = startTagAt(xml, match.index);
    const name = attributeValue(tag, "name");
    const close = tag.endsWith("/>") ? -1 : xml.indexOf(`</${match[1]}`, match.index + tag.length);
    if (name === undefined || close < 0) {
      continue;
    }
    const formula = canonicalOf(() => fromXlsxFormula(decodeXmlText(xml.slice(match.index + tag.length, close))));
    const local = attributeValue(tag, "localSheetId")?.trim();
    const scope = local === undefined ? undefined : /^\d+$/.test(local) ? sheets[Number(local)]?.name : undefined;
    if (formula === null || (local !== undefined && scope === undefined)) {
      continue;
    }
    entries.push(scope === undefined ? { name, formula, origin: A1 } : { name, scope, formula, origin: A1 });
  }
  return entries;
}

/** The `table:named-range`s and `table:named-expression`s in an ODS text
 * from `from` on. */
function odsEntries(xml: string, from: number): NameEntry[] {
  const entries: NameEntry[] = [];
  const pattern = /<table:named-(range|expression)(?=[\s/>])/g;
  pattern.lastIndex = from;
  for (const match of xml.matchAll(pattern)) {
    const tag = startTagAt(xml, match.index);
    const name = attributeValue(tag, "table:name");
    const formula =
      match[1] === "range"
        ? odsRangeAddressToCanonical(attributeValue(tag, "table:cell-range-address") ?? "")
        : canonicalOf(() => fromOdsFormula(attributeValue(tag, "table:expression") ?? ""));
    if (name === undefined || formula === null) {
      continue;
    }
    const base = parseOdsRangeAddress(attributeValue(tag, "table:base-cell-address") ?? "")?.start;
    entries.push({ name, formula, origin: { row: base?.row ?? 0, col: base?.col ?? 0 } });
  }
  return entries;
}

/** A table's own names: its `table:named-expressions` follow its rows. */
function odsTableEntries(sourceXml: string | undefined): NameEntry[] {
  const at = sourceXml?.lastIndexOf("<table:named-expressions") ?? -1;
  return at < 0 ? [] : odsEntries(sourceXml!, at);
}

const XLSX_NAMES = new WeakMap<SpreadsheetModel["opaque"], NameEntry[]>();
const ODS_NAMES = new WeakMap<NonNullable<SpreadsheetModel["odsContentEnvelope"]>, NameEntry[]>();
const ODS_TABLE_NAMES = new WeakMap<Sheet, NameEntry[]>();

function cached<K extends object>(cache: WeakMap<K, NameEntry[]>, key: K, read: () => NameEntry[]): NameEntry[] {
  let entries = cache.get(key);
  if (!entries) {
    entries = read();
    cache.set(key, entries);
  }
  return entries;
}

/** What the source file defines `name` as, used on the sheet `hostSheetId`
 * (`sheetQualifier`: the sheet a qualified name — `Lists!Colours` — names,
 * in model terms); undefined when it defines no such name. */
export function sourceDefinedName(
  model: SpreadsheetModel,
  name: string,
  hostSheetId: string,
  sheetQualifier?: string,
): SourceDefinedName | undefined {
  const scopeSheet =
    sheetQualifier === undefined
      ? model.sheets.find((sheet) => sheet.id === hostSheetId)
      : model.sheets.find((sheet) => sheet.name.toUpperCase() === sheetQualifier.toUpperCase());
  if (!scopeSheet) {
    return undefined;
  }
  const key = name.toUpperCase();
  const named = (entry: NameEntry): boolean => entry.name.toUpperCase() === key;
  const parts = sourcePartsOf(model);
  let found: NameEntry | undefined;
  let own: NameEntry | undefined;
  if (model.format === "xlsx") {
    const entries = cached(XLSX_NAMES, model.opaque, () => xlsxEntries(parts)).filter(named);
    own = entries.find((entry) => entry.scope !== undefined && resolveRetainedSheetName(model, parts, entry.scope, scopeSheet.id)?.id === scopeSheet.id);
    found = own ?? entries.find((entry) => entry.scope === undefined);
  } else if (model.format === "ods") {
    const envelope = model.odsContentEnvelope;
    own = cached(ODS_TABLE_NAMES, scopeSheet, () => odsTableEntries(scopeSheet.sourceXml)).find(named);
    found = own ?? (envelope ? cached(ODS_NAMES, envelope, () => odsEntries(envelope.suffix, 0)).find(named) : undefined);
  }
  if (!found) {
    return undefined;
  }
  // SPEC-133: the model owns names — its copy follows every edit (SPEC-124's
  // retained rows for name text are retired), so a name the model holds in
  // the entry's scope answers with the model's formula, at the entry's origin.
  const scopeSheetId = own ? scopeSheet.id : undefined;
  const modelled = (model.names ?? []).find((each) => !each.formulaUnparsed && each.scopeSheetId === scopeSheetId && each.name.toUpperCase() === key);
  if (modelled) {
    return { formula: rebaseRelative(modelled.formula, found.origin.row, found.origin.col), origin: found.origin };
  }
  return { formula: remapSheetReferences(found.formula, sourceSheetRemap(model, parts)), origin: found.origin };
}
