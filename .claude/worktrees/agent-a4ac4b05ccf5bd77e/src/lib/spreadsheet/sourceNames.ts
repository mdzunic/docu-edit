// Source-time sheet names (SPEC-121 "Sheet references in formulas",
// retained side). Sheet names inside retained text — `opaque`,
// `odsContentEnvelope`, `sourceXml` — stay as the source file wrote them:
// commands rewrite the model's formulas, and SPEC-124's command-time rewrites
// adjust rows and columns in retained text but never rename. So each writer
// derives ONE remap from the source workbook and the model:
//
// - sourceSheetRemap: an original whose source name (xlsx: the source
//   `<sheet name>` for its `partName`; ods: its sourceXml's `table:name`)
//   differs from its model name is renamed; a source sheet no original holds
//   any more is deleted; `orderBefore` is the source order.
// - copyRemap: a copy's retained text is a copy of its origin's source-time
//   text, so it is remapped with the same map plus `<copiedFrom> → <copy
//   name>` (its references to its origin go to itself).
// - resolveRetainedSheetName: the same mapping for code that interprets a
//   name found in retained text (SPEC-124's RetainedReferenceMapper, a
//   reader-upgrade step), `hostSheetId` selecting a copy's override.
import type { SheetRemap } from "./formula/sheetRefs";
import { isOriginal, type Sheet, type SpreadsheetModel } from "./model";
import { sourceTables, tableNameOf } from "./ods/package";
import { LazyParts, partTextOf, readSourceWorkbook, type PartLookup } from "./xlsx/package";

const XLSX_SOURCE_NAMES = new WeakMap<PartLookup, Map<string, string>>();

/** xlsx: partName → source sheet name, from the retained workbook part. */
function xlsxSourceNames(parts: PartLookup): Map<string, string> {
  let names = XLSX_SOURCE_NAMES.get(parts);
  if (!names) {
    names = new Map((readSourceWorkbook(parts)?.sheets ?? []).map((sheet) => [sheet.partName, sheet.name] as const));
    XLSX_SOURCE_NAMES.set(parts, names);
  }
  return names;
}

/** The name the source file gave a sheet's retained text: its own for an
 * original, its origin's (`copiedFrom`) for a copy; undefined for a sheet
 * with no source. */
export function sourceNameOf(model: SpreadsheetModel, sheet: Sheet, parts?: PartLookup): string | undefined {
  if (sheet.copiedFrom !== undefined) {
    return sheet.copiedFrom;
  }
  if (sheet.sourceXml !== undefined) {
    return tableNameOf(sheet.sourceXml);
  }
  if (sheet.partName !== undefined && model.format === "xlsx") {
    return xlsxSourceNames(parts ?? sourcePartsOf(model)).get(sheet.partName);
  }
  return undefined;
}

/** The source workbook's sheets in source order: name and (xlsx) part. */
export function sourceSheetOrder(model: SpreadsheetModel, parts: PartLookup): { name: string; partName?: string }[] {
  if (model.format === "xlsx") {
    return (readSourceWorkbook(parts)?.sheets ?? []).map((sheet) => ({ name: sheet.name, partName: sheet.partName }));
  }
  const envelope = model.odsContentEnvelope;
  const content = envelope ? partTextOf(parts, "content.xml") : undefined;
  if (!envelope || content === undefined) {
    return [];
  }
  return sourceTables(content, envelope.prefix.length, envelope.suffix.length).map((table) => ({ name: table.name }));
}

/** The remap of the workbook's retained text (see the module header). */
export function sourceSheetRemap(model: SpreadsheetModel, parts: PartLookup): SheetRemap {
  const order = sourceSheetOrder(model, parts);
  const renamed: Record<string, string> = {};
  const heldParts = new Set<string>();
  const heldNames = new Set<string>();
  for (const sheet of model.sheets) {
    if (!isOriginal(sheet)) {
      continue;
    }
    const source = sourceNameOf(model, sheet, parts);
    if (source === undefined) {
      continue;
    }
    heldNames.add(source.toUpperCase());
    if (sheet.partName !== undefined) {
      heldParts.add(sheet.partName);
    }
    if (source !== sheet.name) {
      renamed[source] = sheet.name;
    }
  }
  const deleted = order
    .filter((sheet) => (model.format === "xlsx" ? !heldParts.has(sheet.partName ?? "") : !heldNames.has(sheet.name.toUpperCase())))
    .map((sheet) => sheet.name);
  return { renamed, deleted, orderBefore: order.map((sheet) => sheet.name) };
}

/** The remap of one sheet's own retained text: `base`, plus its origin's
 * name → its own for a copy. */
export function copyRemap(base: SheetRemap, sheet: Sheet): SheetRemap {
  const origin = sheet.copiedFrom;
  if (origin === undefined) {
    return base;
  }
  const key = origin.toUpperCase();
  const renamed: Record<string, string> = {};
  for (const [from, to] of Object.entries(base.renamed)) {
    if (from.toUpperCase() !== key) {
      renamed[from] = to;
    }
  }
  renamed[origin] = sheet.name;
  return { renamed, deleted: base.deleted.filter((name) => name.toUpperCase() !== key), orderBefore: base.orderBefore };
}

/** The model sheet a sheet name found in retained text now means (see the
 * module header); undefined for a deleted or unknown sheet. */
export function resolveRetainedSheetName(
  model: SpreadsheetModel,
  parts: PartLookup,
  name: string,
  hostSheetId?: string,
): Sheet | undefined {
  const key = name.toUpperCase();
  const host = hostSheetId === undefined ? undefined : model.sheets.find((sheet) => sheet.id === hostSheetId);
  if (host?.copiedFrom !== undefined && host.copiedFrom.toUpperCase() === key) {
    return host;
  }
  return model.sheets.find((sheet) => isOriginal(sheet) && sourceNameOf(model, sheet, parts)?.toUpperCase() === key);
}

const LAZY = new WeakMap<SpreadsheetModel["opaque"], PartLookup>();

/** A lookup over the model's opaque parts, kept per `opaque` array. */
export function sourcePartsOf(model: SpreadsheetModel): PartLookup {
  let lookup = LAZY.get(model.opaque);
  if (!lookup) {
    lookup = new LazyParts(model.opaque);
    LAZY.set(model.opaque, lookup);
  }
  return lookup;
}
