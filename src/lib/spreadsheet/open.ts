// Ties workbook bytes to the workbook model (SPEC-116): assemble a parsed
// package into a SpreadsheetModel, and bring an already-stored workbook up to
// the current reader version. The app-level funnel (openDocument.ts) owns
// choosing between products; this module owns everything from "it's a
// workbook" down. The readers load on demand (each format's codec is its own
// chunk), so the store can import this module to check reader versions
// without paying for either reader.
//
// Reader versions. A spec that teaches the readers to extract something new
// bumps SPREADSHEET_READER_VERSION and appends its ReaderUpgrade. On load, a
// record behind the current version is upgraded: if no sheet is dirty
// (SPEC-121 adds `Sheet.dirty`) it is simply re-read from `opaque`, keeping
// `id`, `title`, the library's starred/trash state (which live on the store
// record, not the model) and each sheet's `id` (matched by position);
// otherwise clean sheets are re-read and each dirty sheet runs every pending
// step in order. The store writes the upgraded record back once. (The same
// move SPEC-079 made with upgradePresentationModel, made general up front
// because this queue grows the readers fourteen times.)
import { decodedPartBytes } from "../core/packageParts";
import type { Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import type { ParsedSpreadsheet } from "./parsed";
import { buildPreview } from "./preview";

export const SPREADSHEET_READER_VERSION = 1;

export interface ReaderUpgrade {
  from: number; // upgrades a record at `from` to `from + 1`
  /** Brings ONE dirty sheet up to date by reading the newly modelled element
   *  from the sheet's retained source and merging it into the model. */
  upgradeDirtySheet(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>): void;
}

export const READER_UPGRADES: ReaderUpgrade[] = [];

/** The title a workbook gets when nothing names it. */
export const DEFAULT_SPREADSHEET_TITLE = "Untitled workbook";

/** Assembles a parsed package into a fresh SpreadsheetModel — the shared
 * tail of every workbook open path, minting a new id the way
 * assembleDocumentModel/assemblePresentation do. Optional fields are
 * omitted when empty, matching the other models' convention. */
export function assembleSpreadsheet(title: string, format: SpreadsheetFormat, parsed: ParsedSpreadsheet): SpreadsheetModel {
  const model: SpreadsheetModel = {
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title,
    format,
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets: parsed.sheets,
    activeSheetIndex: parsed.activeSheetIndex,
    dateSystem: parsed.dateSystem,
    opaque: parsed.opaque,
  };
  if (parsed.packageVariant !== undefined) {
    model.packageVariant = parsed.packageVariant;
  }
  if (parsed.odsNullDate !== undefined) {
    model.odsNullDate = parsed.odsNullDate;
  }
  if (parsed.odsContentEnvelope !== undefined) {
    model.odsContentEnvelope = parsed.odsContentEnvelope;
  }
  if (Object.keys(parsed.properties).length > 0) {
    model.properties = parsed.properties;
  }
  model.preview = buildPreview(model);
  return model;
}

async function parse(
  bytes: Uint8Array,
  format: SpreadsheetFormat,
  parts: Map<string, Uint8Array> | undefined,
): Promise<ParsedSpreadsheet> {
  if (format === "xlsx") {
    const { parseXlsx } = await import("./xlsx/reader");
    return parseXlsx(bytes, parts);
  }
  const { parseOds } = await import("./ods/reader");
  return parseOds(bytes, parts);
}

/** Reads and parses workbook bytes into a new SpreadsheetModel, with no
 * `File`/FileReader involved (the Node-friendly split readDocxBytes and
 * readPresentationBytes make). `parts`, when the caller already inflated the
 * package, is used instead of `bytes`. */
export async function readSpreadsheetBytes(
  bytes: Uint8Array,
  format: SpreadsheetFormat,
  title: string,
  parts?: Map<string, Uint8Array>,
): Promise<SpreadsheetModel> {
  return assembleSpreadsheet(title, format, await parse(bytes, format, parts));
}

/** Whether a sheet carries edits its source does not (SPEC-121 adds the
 * flag; until then no sheet is ever dirty). */
function isDirty(sheet: Sheet): boolean {
  return (sheet as Sheet & { dirty?: boolean }).dirty === true;
}

/** Brings a stored workbook up to SPREADSHEET_READER_VERSION (see the module
 * header). A current record comes back as the same object; an older one is
 * re-read from its retained package. */
export async function upgradeSpreadsheetModel(model: SpreadsheetModel): Promise<SpreadsheetModel> {
  if (model.readerVersion >= SPREADSHEET_READER_VERSION) {
    return model;
  }
  const sourceParts = decodedPartBytes(model.opaque);
  const fresh = await readSpreadsheetBytes(new Uint8Array(0), model.format, model.title, sourceParts);
  fresh.id = model.id;
  const pending = READER_UPGRADES.filter((step) => step.from >= model.readerVersion).toSorted((a, b) => a.from - b.from);
  fresh.sheets = fresh.sheets.map((sheet, index) => {
    const stored = model.sheets[index];
    if (!stored) {
      return sheet;
    }
    if (isDirty(stored)) {
      for (const step of pending) {
        step.upgradeDirtySheet(fresh, stored, sourceParts);
      }
      return stored;
    }
    return { ...sheet, id: stored.id };
  });
  fresh.preview = buildPreview(fresh);
  return fresh;
}
