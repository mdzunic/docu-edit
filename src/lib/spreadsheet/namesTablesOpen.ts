// SPEC-133's reader-upgrade step (13 → 14, open.ts "Version history"):
// defined names and tables for a record edited before this version. Loaded
// on demand by open.ts's READER_UPGRADES.
//
// - upgradeWorkbookNames (ReaderUpgrade.upgradeWorkbook, once per edited
//   record): the names of the retained workbook part (xlsx — `localSheetId`
//   through the SOURCE sheet order to the model sheet holding that part,
//   names of deleted sheets dropped) or of the retained content envelope and
//   each sheet's retained `sourceXml` (ods — the sheet whose retained
//   `table:name` matches), every formula through SPEC-121's
//   sourceSheetRemap. A copy (SPEC-121's `copiedFrom`) gets its origin's
//   sheet-scoped names through its own remap (copyRemap) — xlsx: SPEC-121's
//   writer duplicated them at write time; ods: they are in the copy's
//   retained table text. A record that already holds names keeps them.
// - upgradeSheetTables (upgradeDirtySheet): a dirty xlsx sheet's tables
//   from its retained part's `<tableParts>` and the table parts in `opaque`
//   (SPEC-124's retained rewrite kept their refs current), their formulas
//   remapped the same way. ods has no tables.
import type { DefinedName, Sheet, SpreadsheetModel } from "./model";
import { remapSheetReferences, type SheetRemap } from "./formula/sheetRefs";
import { readOdsTableNames, readOdsWorkbookNames } from "./ods/namedExpressions";
import { copyRemap, sourcePartsOf, sourceSheetRemap } from "./sourceNames";
import { readSourceDefinedNames } from "./xlsx/definedNames";
import { partTextOf } from "./xlsx/package";
import { readXlsxSheetTables } from "./xlsx/tableXml";

const TABLE_NAME = /<table:table\s[^>]*?table:name="([^"]*)"/;

/** A parsed name's formula through `remap`. */
function remapped(name: DefinedName, remap: SheetRemap): DefinedName {
  return name.formulaUnparsed ? name : { ...name, formula: remapSheetReferences(name.formula, remap) };
}

/** The record's names, read from its retained source (see the header). */
export function upgradeWorkbookNames(model: SpreadsheetModel, sourceParts: Map<string, Uint8Array>): void {
  if (model.names.length > 0) {
    return;
  }
  const parts = sourcePartsOf({ ...model, opaque: model.opaque });
  const remap = sourceSheetRemap(model, parts);
  let names: DefinedName[] = [];
  if (model.format === "xlsx") {
    names = readSourceDefinedNames(sourceParts.size > 0 ? sourceParts : parts, model.sheets, remap);
  } else if (model.format === "ods" && model.odsContentEnvelope) {
    names = readOdsWorkbookNames(model.odsContentEnvelope.suffix).map((name) => remapped(name, remap));
    for (const sheet of model.sheets) {
      if (sheet.sourceXml !== undefined && TABLE_NAME.test(sheet.sourceXml)) {
        // A copy's retained text is its origin's: its names, through the copy's own remap.
        const by = copyRemap(remap, sheet);
        names.push(...readOdsTableNames(sheet.sourceXml, sheet.id).map((name) => remapped(name, by)));
      }
    }
  }
  if (names.length > 0) {
    model.names = names;
  }
}

/** A dirty xlsx sheet's tables from its retained source (see the header). */
export function upgradeSheetTables(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>, record: SpreadsheetModel): void {
  if (model.format !== "xlsx" || sheet.partName === undefined || sheet.tables !== undefined) {
    return;
  }
  const xml = partTextOf(sourceParts, sheet.partName);
  if (xml === undefined) {
    return;
  }
  const remap = sourceSheetRemap(record, sourcePartsOf(record));
  const tables = readXlsxSheetTables(sourceParts, sheet.partName, xml, () => crypto.randomUUID()).map((table) => ({
    ...table,
    columns: table.columns.map((column) => ({
      ...column,
      ...(column.calculatedFormula === undefined ? {} : { calculatedFormula: remapSheetReferences(column.calculatedFormula, remap) }),
      ...(column.totalsFormula === undefined ? {} : { totalsFormula: remapSheetReferences(column.totalsFormula, remap) }),
    })),
  }));
  if (tables.length > 0) {
    sheet.tables = tables;
  }
}
