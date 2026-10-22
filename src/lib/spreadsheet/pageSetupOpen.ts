// SPEC-136's reader step (open.ts, version 15 → 16): an edited record's
// page setups and the reserved names ODS keeps in table structure, read from
// what the record retained. Loaded lazily, as the other steps are.
//
// - A dirty xlsx sheet reads its page elements from its retained worksheet
//   part (SPEC-122 kept them verbatim; SPEC-124's `retained` adjuster kept
//   `brk@id` current until this spec retired that row). Its print area and
//   titles are defined names, which SPEC-133 already models.
// - A dirty ods sheet reads its master page from `styles.xml` in `opaque`
//   and its header rows and row/column breaks from its retained table text —
//   whose positions may predate a structural edit made before this spec
//   landed, since no adjuster rewrote them.
// - Every ods sheet of an edited record (upgradeWorkbookPrintNames) gets the
//   `_xlnm.Print_Area` its retained start tag's `table:print-ranges` means
//   (kept current by SPEC-124's `retained` adjuster) and the
//   `_xlnm.Print_Titles` its retained header rows/columns mean — unless the
//   record already names one: the record's names win.
import type { Sheet, SpreadsheetModel } from "./model";
import { getReservedName } from "./names";
import { normalizePageSetup } from "./pageSetup";
import { readOdsPageStyles, readOdsTablePageSetup } from "./ods/pageSetup";
import { readXlsxPageSetup } from "./xlsx/pageSetup";
import { partTextOf } from "./xlsx/package";

const decoder = new TextDecoder();

function stylesText(sourceParts: Map<string, Uint8Array>): string | undefined {
  const bytes = sourceParts.get("styles.xml");
  return bytes === undefined ? undefined : decoder.decode(bytes);
}

/** One dirty sheet's page setup from its retained source (see the header). */
export function upgradeSheetPageSetup(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>, record: SpreadsheetModel): void {
  if (sheet.kind !== "grid") {
    return;
  }
  let setup = sheet.pageSetup;
  if (model.format === "xlsx" && sheet.partName !== undefined) {
    const xml = partTextOf(sourceParts, sheet.partName);
    setup = xml === undefined ? undefined : readXlsxPageSetup(xml);
  } else if (model.format === "ods" && sheet.sourceXml !== undefined) {
    const prefix = record.odsContentEnvelope?.prefix ?? model.odsContentEnvelope?.prefix ?? "";
    setup = readOdsTablePageSetup(sheet, prefix, readOdsPageStyles(stylesText(sourceParts))).setup;
  }
  const normalized = normalizePageSetup(setup);
  if (normalized === undefined) {
    delete sheet.pageSetup;
  } else {
    sheet.pageSetup = normalized;
  }
}

/** Every ods sheet's print area and titles from its retained table text
 * (see the header). */
export function upgradeWorkbookPrintNames(model: SpreadsheetModel, sourceParts: Map<string, Uint8Array>): void {
  if (model.format !== "ods") {
    return;
  }
  const styles = readOdsPageStyles(stylesText(sourceParts));
  const prefix = model.odsContentEnvelope?.prefix ?? "";
  for (const sheet of model.sheets) {
    if (sheet.kind !== "grid" || sheet.sourceXml === undefined) {
      continue;
    }
    const { names } = readOdsTablePageSetup(sheet, prefix, styles);
    const missing = names.filter((name) => getReservedName(model, sheet.id, name.name.slice("_xlnm.".length)) === undefined);
    if (missing.length > 0) {
      model.names = [...(model.names ?? []), ...missing];
    }
  }
}
