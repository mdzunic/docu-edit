// SPEC-135's reader-upgrade step (14 → 15, open.ts "Version history"):
// notes, threads and hyperlinks for a sheet edited before this version.
// Loaded on demand by open.ts's READER_UPGRADES; a clean sheet is re-read
// whole and gains them from the readers.
//
// - xlsx: the sheet part's `.rels` in `opaque` lead to its comments, VML and
//   threaded parts, parsed as xlsx/annotations.ts reads them; `model.persons`
//   is read once from the workbook rels when the record has none; the
//   hyperlinks come from the retained worksheet XML, which already carries
//   SPEC-124's textual `@ref` adjustments.
// - ods: every `office:annotation` moves out of `Cell.preservedXml` into
//   `notes`; a link of the retained `sourceXml` is attached only where the
//   current cell at its address still holds the source cell's value (a value
//   or structural edit since read makes the address unreliable, and SPEC-122
//   had already dropped such a link from any save).
// Either way a `place` location is source-time text, so SPEC-121's
// sourceSheetRemap (a copy's own remap included) renames it. Nothing is
// marked dirty: the annotations are as the file holds them.
import { remapSheetReferences } from "./formula/sheetRefs";
import type { CellHyperlink, Sheet, SpreadsheetModel } from "./model";
import { copyRemap, sourcePartsOf, sourceSheetRemap } from "./sourceNames";
import { partTextOf } from "./xlsx/package";

/** ReaderUpgrade.upgradeDirtySheet for SPEC-135 (see the module header). */
export async function upgradeSheetAnnotations(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>, record: SpreadsheetModel): Promise<void> {
  if (sheet.notes !== undefined || sheet.threads !== undefined || sheet.hyperlinks !== undefined) {
    return;
  }
  const remap = copyRemap(sourceSheetRemap(record, sourcePartsOf(record)), sheet);
  const remapLocation = (location: string): string => remapSheetReferences(location, remap);
  if (model.format === "xlsx") {
    const xml = sheet.partName === undefined ? undefined : partTextOf(sourceParts, sheet.partName);
    if (xml === undefined) {
      return;
    }
    const { readXlsxPersons, readXlsxSheetAnnotations } = await import("./xlsx/annotations");
    if (model.persons === undefined) {
      const persons = readXlsxPersons(sourceParts);
      if (persons !== undefined && persons.length > 0) {
        model.persons = persons;
      }
    }
    const read = readXlsxSheetAnnotations(sourceParts, sheet.partName!, xml, { persons: model.persons });
    if (read.notes !== undefined && read.notes.length > 0) {
      sheet.notes = read.notes;
    }
    if (read.threads !== undefined && read.threads.length > 0) {
      sheet.threads = read.threads;
    }
    if (read.hyperlinks !== undefined && read.hyperlinks.length > 0) {
      sheet.hyperlinks = read.hyperlinks.map((link) => remappedPlace(link, remapLocation));
    }
    return;
  }
  const prefix = model.odsContentEnvelope?.prefix;
  if (prefix === undefined) {
    return;
  }
  const { upgradeOdsSheetAnnotations } = await import("./ods/annotations");
  const read = upgradeOdsSheetAnnotations({ sheet, prefix, dates: record, remapLocation });
  sheet.rows = read.rows;
  if (read.notes.length > 0) {
    sheet.notes = read.notes;
  }
  if (read.hyperlinks.length > 0) {
    sheet.hyperlinks = read.hyperlinks;
  }
}

/** `link` with its `place` location through `remapLocation`. */
function remappedPlace(link: CellHyperlink, remapLocation: (location: string) => string): CellHyperlink {
  return link.target.kind === "place" ? { ...link, target: { kind: "place", location: remapLocation(link.target.location) } } : link;
}
