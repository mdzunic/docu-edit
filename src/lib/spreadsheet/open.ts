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
//
// Version history:
// 1 — SPEC-116: values, rich-text runs, sheets, the date system.
// 2 — SPEC-117: cell styles (xl/styles.xml, the ODS automatic and common
//     styles), the theme, column/row sizes, merges, frozen panes, hidden rows
//     and columns, view settings, style-only cells and ODS covered cells.)
// 3 — SPEC-118: number formats (xlsx numFmts and built-in ids, ODS data
//     styles) and the implied formats of unformatted dates, times and
//     percentages.
// 4 — SPEC-119: formulas (xlsx `<f>` — normal, shared, array, data table,
//     dynamic array — and `c@cm`; ODS `table:formula` and matrix spans) and
//     `sheetView@showFormulas`.
// 5 — SPEC-120: the calculation settings (`calc`): xlsx `calcPr`
//     fullCalcOnLoad, calcMode and iterate; ODS `table:calculation-settings`
//     and settings.xml's AutoCalculate. A re-read keeps the stored model's
//     `forceFullCalcOnLoad` (set by recalculation, never cleared).
import { decodedPartBytes } from "../core/packageParts";
import type { Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import type { ParsedSpreadsheet } from "./parsed";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { buildPreview } from "./preview";
import { defaultStyles } from "./styles";

export const SPREADSHEET_READER_VERSION = 5;

/** The locale a preview built here formats in: the reader knows no
 * preference — the store rebuilds the card in the user's locale (SPEC-118). */
const READ_PREVIEW_LOCALE = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];

export interface ReaderUpgrade {
  from: number; // upgrades a record at `from` to `from + 1`
  /** Brings ONE dirty sheet up to date by reading the newly modelled element
   *  from the sheet's retained source and merging it into the model. */
  upgradeDirtySheet(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>): void;
}

export const READER_UPGRADES: ReaderUpgrade[] = [
  {
    // SPEC-117 (1 → 2). A record behind this version is re-read whole from
    // `opaque`: no sheet can be dirty before SPEC-121 lands, and every read
    // path upgrades. The step exists to keep READER_UPGRADES contiguous, and
    // throws so the impossible case is loud instead of silently merging half
    // a style table into an edited sheet.
    from: 1,
    upgradeDirtySheet() {
      throw new Error("unreachable: no sheet is dirty before SPEC-121");
    },
  },
  {
    // SPEC-118 (2 → 3): number formats. The same choice and reason as
    // SPEC-117's step — every read path upgrades by re-reading `opaque`, and
    // no sheet can be dirty before SPEC-121; this keeps the registry
    // contiguous.
    from: 2,
    upgradeDirtySheet() {
      throw new Error("unreachable: no sheet is dirty before SPEC-121");
    },
  },
  {
    // SPEC-119 (3 → 4): formulas. A record from version 3 re-reads its
    // (necessarily clean) sheets from the retained source and gains its
    // formulas: a sheet can only become dirty under a build that has
    // SPEC-121, and such a build has upgraded the record past this step on
    // load. The step keeps the registry contiguous.
    from: 3,
    upgradeDirtySheet() {
      throw new Error("SPEC-119 upgrade: a sheet cannot be dirty before SPEC-121");
    },
  },
  {
    // SPEC-120 (4 → 5): calculation settings, which are the workbook's, not a
    // sheet's — a re-read gains them from the retained source. The step keeps
    // the registry contiguous and refuses the impossible dirty sheet, as
    // SPEC-119's does.
    from: 4,
    upgradeDirtySheet() {
      throw new Error("SPEC-120 upgrade: a sheet cannot be dirty before SPEC-121");
    },
  },
];

export { DEFAULT_SPREADSHEET_TITLE } from "./model";

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
    styles: parsed.styles.length > 0 ? parsed.styles : defaultStyles(),
    defaultStyleId: 0,
  };
  if (parsed.theme !== undefined) {
    model.theme = parsed.theme;
  }
  if (parsed.indexedColors !== undefined) {
    model.indexedColors = parsed.indexedColors;
  }
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
  if (parsed.calc !== undefined && Object.keys(parsed.calc).length > 0) {
    model.calc = parsed.calc;
  }
  model.preview = buildPreview(model, READ_PREVIEW_LOCALE);
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
  if (model.forceFullCalcOnLoad) {
    fresh.forceFullCalcOnLoad = true;
  }
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
  fresh.preview = buildPreview(fresh, READ_PREVIEW_LOCALE);
  return fresh;
}
