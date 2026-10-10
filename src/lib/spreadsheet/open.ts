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
// record behind the current version is upgraded: if nothing was edited (no
// `structureDirty`, no dirty sheet — SPEC-121) it is simply re-read from
// `opaque`, keeping `id`, `title`, the library's starred/trash state (which
// live on the store record, not the model) and each sheet's `id` (matched by
// position); otherwise (SPEC-121 "Reader upgrades after edits") the model's
// sheet list is kept and each clean sheet — a copy included — is re-read
// from its own retained source, while each dirty sheet runs every pending
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
// 6 — SPEC-122: the content the writers re-emit for what the model does not
//     model — each cell's `sstIndex`, `preservedAttrs`, `preservedXml` and
//     `preservedText` — plus `Sheet.protection` and the workbook's
//     `structureProtected`. Its step is the first that runs on a dirty sheet:
//     it re-reads the sheet's retained source (an xlsx part with the source
//     package's shared strings and styles; an ods table text inside the
//     retained content envelope) and copies those fields onto each model
//     cell whose value and formula still equal the source's at that address,
//     and the protection flag onto the sheet; `structureProtected` is the
//     workbook part's, which every re-read gains. The xlsx reader also
//     stores, from version 6, a default-width column span or a row without
//     cells that carries only an unmodelled attribute the writer re-emits (an
//     `outlineLevel`, `collapsed` …); the step adds those to a dirty xlsx
//     sheet too — its rows and columns are version 5's, since no command
//     before SPEC-122 adds or removes one, so they are exactly the source's
//     spans that overlap no stored span and its rows without cells where none
//     is stored.
//
// Since SPEC-122 the writers rebuild a dirty sheet from the model, so the
// layout a record re-reads its clean sheets from numbers its styles afresh:
// each dirty sheet's style ids are moved into the re-read style table
// (interned — an equal style is shared), never left pointing into the
// stored one.
import { decodedPartBytes } from "../core/packageParts";
import { isCellError, type Cell, type CellStyle, type CellValue, type Row, type Sheet, type SpreadsheetFormat, type SpreadsheetModel } from "./model";
import type { ParsedSpreadsheet } from "./parsed";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { buildPreview } from "./preview";
import { cellStyle, defaultStyles, internStyle, styleKey } from "./styles";

export const SPREADSHEET_READER_VERSION = 6;

/** The locale a preview built here formats in: the reader knows no
 * preference — the store rebuilds the card in the user's locale (SPEC-118). */
const READ_PREVIEW_LOCALE = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];

export interface ReaderUpgrade {
  from: number; // upgrades a record at `from` to `from + 1`
  /** Brings ONE dirty sheet up to date by reading the newly modelled element
   *  from the sheet's retained source and merging it into the model (the
   *  sheet object is updated in place; its rows stay copy-on-write). */
  upgradeDirtySheet(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>): void | Promise<void>;
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
  {
    // SPEC-122 (5 → 6): preserved cell content and sheet protection (see the
    // version history). A record at version 5 can hold dirty sheets — SPEC-121
    // shipped editing — so this step does real work.
    from: 5,
    async upgradeDirtySheet(model, sheet, sourceParts) {
      const source = await retainedSheetRead(model, sheet, sourceParts);
      if (!source) {
        return;
      }
      if (source.protection) {
        sheet.protection = { sheet: true };
      }
      sheet.rows = withRetainedContent(sheet.rows, source.rows);
      if (model.format === "xlsx") {
        withRetainedOutline(sheet, source);
      }
    },
  },
];

// --- SPEC-122's step ------------------------------------------------------------------------

/** One read of a source package per upgrade, shared by its dirty sheets. */
const sourceReads = new WeakMap<Map<string, Uint8Array>, Promise<ParsedSpreadsheet>>();

/** The sheet as the current reader reads it from its own retained source:
 * an xlsx sheet's part (an original's, or for a copy whose part the source
 * workbook does not list, its origin's — the copy's part is a copy of it),
 * read with the source package's shared strings and styles; an ods sheet's
 * table text inside the retained content envelope. Undefined when there is
 * none to read. */
async function retainedSheetRead(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>): Promise<Sheet | undefined> {
  if (model.format === "xlsx") {
    if (sheet.partName === undefined) {
      return undefined;
    }
    let read = sourceReads.get(sourceParts);
    if (!read) {
      read = parse(new Uint8Array(0), "xlsx", sourceParts);
      sourceReads.set(sourceParts, read);
    }
    const parsed = await read;
    return (
      parsed.sheets.find((each) => each.partName === sheet.partName) ??
      (sheet.copiedFrom === undefined ? undefined : parsed.sheets.find((each) => each.name.toUpperCase() === sheet.copiedFrom!.toUpperCase()))
    );
  }
  const envelope = model.odsContentEnvelope;
  if (sheet.sourceXml === undefined || !envelope) {
    return undefined;
  }
  const parts = new Map(sourceParts);
  parts.set("content.xml", new TextEncoder().encode(envelope.prefix + sheet.sourceXml + envelope.suffix));
  return (await parse(new Uint8Array(0), "ods", parts)).sheets[0];
}

function sameValue(a: CellValue, b: CellValue): boolean {
  if (isCellError(a) || isCellError(b)) {
    return isCellError(a) && isCellError(b) && a.error === b.error;
  }
  return Object.is(a, b) || a === b;
}

const RETAINED_FIELDS = ["sstIndex", "preservedAttrs", "preservedXml", "preservedText"] as const;

/** The model's rows with the source's preserved cell content copied onto
 * every cell whose value and formula equal the source cell's (a shared
 * string's index only while its runs are unchanged too) — the same array
 * when nothing is copied; untouched rows and cells keep their identity. */
function withRetainedContent(rows: Row[], sourceRows: readonly Row[]): Row[] {
  const sourceByIndex = new Map(sourceRows.map((row) => [row.index, row] as const));
  let out: Row[] | null = null;
  rows.forEach((row, at) => {
    const sourceRow = sourceByIndex.get(row.index);
    if (!sourceRow) {
      return;
    }
    const sourceCells = new Map(sourceRow.cells.map((cell) => [cell.col, cell] as const));
    let cells: Cell[] | null = null;
    row.cells.forEach((cell, index) => {
      const source = sourceCells.get(cell.col);
      if (!source || !sameValue(cell.value, source.value) || cell.formula !== source.formula) {
        return;
      }
      let next: Cell = cell;
      for (const field of RETAINED_FIELDS) {
        const value = source[field];
        if (value === undefined || next[field] === value) {
          continue;
        }
        if (field === "sstIndex" && JSON.stringify(cell.runs ?? null) !== JSON.stringify(source.runs ?? null)) {
          continue;
        }
        next = next === cell ? { ...cell } : next;
        (next as unknown as Record<string, unknown>)[field] = value;
      }
      if (next !== cell) {
        cells ??= row.cells.slice();
        cells[index] = next;
      }
    });
    if (cells) {
      out ??= rows.slice();
      out[at] = { ...row, cells };
    }
  });
  return out ?? rows;
}

/** Adds the source's column spans that overlap no stored span and its rows
 * without cells where no row is stored (see the version history), in place;
 * arrays are replaced only when something is added. */
function withRetainedOutline(sheet: Sheet, source: Sheet): void {
  const stored = new Set(sheet.rows.map((row) => row.index));
  const rows = source.rows.filter((row) => row.cells.length === 0 && !stored.has(row.index));
  if (rows.length > 0) {
    sheet.rows = [...sheet.rows, ...rows].toSorted((a, b) => a.index - b.index);
  }
  const spans = source.columns.filter((span) => !sheet.columns.some((have) => have.min <= span.max && span.min <= have.max));
  if (spans.length > 0) {
    sheet.columns = [...sheet.columns, ...spans].toSorted((a, b) => a.min - b.min);
  }
}

/** Moves a dirty sheet's style ids from the stored table `from` into the
 * re-read workbook's table (see the module header), in place: an id whose
 * style the re-read table holds at the same index keeps it; any other is
 * interned. Arrays are replaced only when an id changes. */
function adoptStyles(sheet: Sheet, from: readonly CellStyle[], into: SpreadsheetModel): void {
  const ids = new Map<number, number>();
  const mapId = (id: number): number => {
    let next = ids.get(id);
    if (next === undefined) {
      const style = cellStyle({ styles: from as CellStyle[] }, id);
      const same = into.styles[id];
      next =
        same !== undefined && styleKey(same) === styleKey(style) && same.sourceXf === style.sourceXf && same.sourceStyleName === style.sourceStyleName
          ? id
          : internStyle(into, style);
      ids.set(id, next);
    }
    return next;
  };
  let rows: Row[] | null = null;
  sheet.rows.forEach((row, at) => {
    let cells: Cell[] | null = null;
    row.cells.forEach((cell, index) => {
      if (cell.styleId !== undefined && mapId(cell.styleId) !== cell.styleId) {
        cells ??= row.cells.slice();
        cells[index] = { ...cell, styleId: mapId(cell.styleId) };
      }
    });
    const rowStyle = row.styleId === undefined ? undefined : mapId(row.styleId);
    if (cells || rowStyle !== row.styleId) {
      rows ??= sheet.rows.slice();
      rows[at] = { ...row, cells: cells ?? row.cells, ...(rowStyle === undefined ? {} : { styleId: rowStyle }) };
    }
  });
  if (rows) {
    sheet.rows = rows;
  }
  if (sheet.columns.some((span) => span.styleId !== undefined && mapId(span.styleId) !== span.styleId)) {
    sheet.columns = sheet.columns.map((span) => (span.styleId === undefined ? span : { ...span, styleId: mapId(span.styleId) }));
  }
}

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
  if (parsed.structureProtected) {
    model.structureProtected = true;
  }
  if (parsed.namedStyles !== undefined && parsed.namedStyles.length > 0) {
    model.namedStyles = parsed.namedStyles;
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

/** Brings a stored workbook up to SPREADSHEET_READER_VERSION (see the module
 * header). A current record comes back as the same object; an older one is
 * re-read from its retained package — wholesale when nothing was edited,
 * sheet by sheet when something was (upgradeEdited). */
export async function upgradeSpreadsheetModel(model: SpreadsheetModel): Promise<SpreadsheetModel> {
  if (model.readerVersion >= SPREADSHEET_READER_VERSION) {
    return model;
  }
  const sourceParts = decodedPartBytes(model.opaque);
  const pending = READER_UPGRADES.filter((step) => step.from >= model.readerVersion).toSorted((a, b) => a.from - b.from);
  if (model.structureDirty || model.sheets.some((sheet) => sheet.dirty)) {
    return upgradeEdited(model, sourceParts, pending);
  }
  const fresh = await readSpreadsheetBytes(new Uint8Array(0), model.format, model.title, sourceParts);
  fresh.id = model.id;
  if (model.forceFullCalcOnLoad) {
    fresh.forceFullCalcOnLoad = true;
  }
  fresh.sheets = fresh.sheets.map((sheet, index) => {
    const stored = model.sheets[index];
    return stored ? { ...sheet, id: stored.id } : sheet;
  });
  fresh.preview = buildPreview(fresh, READ_PREVIEW_LOCALE);
  return fresh;
}

/** SPEC-121 "Reader upgrades after edits": a record with `structureDirty` or
 * a dirty sheet is never re-read wholesale — the source's sheet list is no
 * longer the workbook's. The model's sheet list, order, names, visibility,
 * tab colours and ids are kept; each clean sheet (copies included) is
 * re-read from its own retained source, replacing only its cell content and
 * per-sheet properties. The read goes through the writer's substitution
 * layout of the model — every sheet at its own part (or its own table text)
 * with sourceSheetRemap applied, a copy's own remap included — so the
 * formulas read name the sheets as the model does, while the record keeps
 * its retained package untouched. Dirty sheets run the pending steps (a step
 * never runs on a sheet with no source). */
async function upgradeEdited(model: SpreadsheetModel, sourceParts: Map<string, Uint8Array>, pending: readonly ReaderUpgrade[]): Promise<SpreadsheetModel> {
  const layout =
    model.format === "xlsx"
      ? (await import("./xlsx/writer")).substitutedXlsxParts(model)
      : (await import("./ods/writer")).substitutedOdsParts(model);
  const fresh = await readSpreadsheetBytes(new Uint8Array(0), model.format, model.title, layout);
  fresh.id = model.id;
  fresh.opaque = model.opaque;
  if (model.odsContentEnvelope) {
    fresh.odsContentEnvelope = model.odsContentEnvelope;
  } else {
    delete fresh.odsContentEnvelope;
  }
  fresh.activeSheetIndex = model.activeSheetIndex;
  fresh.structureDirty = true;
  if (!model.structureDirty) {
    delete fresh.structureDirty;
  }
  if (model.forceFullCalcOnLoad) {
    fresh.forceFullCalcOnLoad = true;
  }
  const sheets: Sheet[] = [];
  for (const [index, stored] of model.sheets.entries()) {
    sheets.push(await upgradedSheet(stored, index));
  }
  fresh.sheets = sheets;
  fresh.preview = buildPreview(fresh, READ_PREVIEW_LOCALE);
  return fresh;

  async function upgradedSheet(stored: Sheet, index: number): Promise<Sheet> {
    if (stored.dirty) {
      if (stored.partName !== undefined || stored.sourceXml !== undefined) {
        for (const step of pending) {
          await step.upgradeDirtySheet(fresh, stored, sourceParts);
        }
      }
      adoptStyles(stored, model.styles, fresh);
      return stored;
    }
    const read = fresh.sheets[index];
    if (!read) {
      return stored;
    }
    const { id: _id, name: _name, visibility: _visibility, kind: _kind, tabColor: _tab, partName: _part, sourceXml: _source, copiedFrom: _copied, ...content } = read;
    const sheet: Sheet = { ...content, id: stored.id, name: stored.name, visibility: stored.visibility, kind: stored.kind };
    if (stored.tabColor !== undefined) {
      sheet.tabColor = stored.tabColor;
    }
    if (stored.partName !== undefined) {
      sheet.partName = stored.partName;
    }
    if (stored.sourceXml !== undefined) {
      sheet.sourceXml = stored.sourceXml;
    }
    if (stored.copiedFrom !== undefined) {
      sheet.copiedFrom = stored.copiedFrom;
    }
    return sheet;
  }
}
