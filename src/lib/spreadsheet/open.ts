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
// 7 — SPEC-123: named cell styles — the workbook's `namedStyles` (xlsx
//     `cellStyleXfs` + `cellStyles`, ods `office:styles` common table-cell
//     styles) and each style's `namedStyle` (xlsx `cellXfs/xf@xfId`, ods
//     `style:parent-style-name`) — and what a protected sheet still permits
//     (xlsx `sheetProtection@formatCells|formatColumns|formatRows="0"`). Every
//     re-read gains them; the step, run for each dirty sheet, fills the
//     workbook-level fields from the retained styles part when the model
//     lacks them — idempotently, so any number of dirty sheets may run it —
//     and refreshes the dirty sheet's protection from its own retained
//     source. A dirty sheet's styles were stored without `namedStyle`, so
//     adopting them treats a stored style that is the re-read one's source
//     entry, but for the `namedStyle` it could not know, as that style.
// 8 — SPEC-124: the structural edits a protected sheet still permits (xlsx
//     `sheetProtection@insertRows|insertColumns|deleteRows|deleteColumns="0"`).
//     Every re-read gains them; the step refreshes a dirty sheet's
//     protection from its own retained source (SPEC-123's move).
// 9 — SPEC-128: the autofilter and the sort records — xlsx `<autoFilter>`
//     (with its `<sortState>`) and the worksheet `<sortState>`, ods the
//     sheet's anonymous `table:database-range` — the filter-hidden rows
//     (`Row.filterHidden`: an xlsx hidden row inside the filter's data rows
//     that a criterion rejects, ods `table:visibility="filter"`, which
//     version 8 read as `hidden`) and the protection's sort and autofilter
//     permits (`sheetProtection@sort|autoFilter="0"`). Every re-read gains
//     them. The step, run for each dirty sheet, takes `autoFilter` and
//     `sortState` from the sheet's retained source — whose references
//     SPEC-124's retained adjusters kept moving with every structural edit
//     and move, so they fit the edited sheet — refreshes its protection, and
//     reclassifies its hidden rows by the read rule (filter.ts's
//     classifyHiddenRows over the sheet's current values, in the workbook's
//     format context and the reader's locale). Both formats use the read
//     rule, not the ods wire's per-row `table:visibility`: a retained table
//     keeps its source-time row numbers, which an edit since may have
//     shifted, so only the criteria can say which of the edited sheet's
//     hidden rows are the filter's. The rule is reading's, so it never shows
//     or hides a row. (The stored sheet's style ids still index the stored
//     style table while the step runs — adoptStyles moves them after — and
//     a row's display text is read through `model.styles`, the re-read
//     table, which holds every source entry at the same index.)
// 10 — SPEC-129: conditional formats (`Sheet.conditionalFormats` — xlsx
//     `<conditionalFormatting>` with the worksheet's x14 ext, ods
//     `calcext:conditional-formats` or, failing those, legacy `style:map`s)
//     and the differential-style table (`model.dxfs` — xlsx `<dxfs>`, ods the
//     named styles rules apply). Every re-read gains them. The step, run for
//     each dirty sheet, fills `dxfs` from the retained styles part when the
//     model lacks it, and takes the sheet's conditional formats from its own
//     retained source — whose references SPEC-124's and SPEC-125's retained
//     adjusters kept moving with every structural edit and move, so their
//     positions fit the edited sheet — each rule's dxf interned into the
//     model's table (internDxf finds the source entry it came from) and its
//     formulas' source-time sheet names remapped onto the record's names.
// 11 — SPEC-130: data validations (`Sheet.validations` — xlsx
//     `<dataValidations>` with the worksheet's x14 ext `{CCE6A557…}` and
//     `disablePrompts`, ods `table:content-validations` with the cells'
//     `table:content-validation-name`; the workbook's
//     `odsOrphanValidations`). Every re-read gains them. The step, run for
//     each dirty sheet, parses only the sheet's validations: xlsx from its
//     retained part (base block and x14 ext), whose references SPEC-124's and
//     SPEC-125's retained adjusters kept moving with every structural edit
//     and move; ods from the retained prefix's definitions — likewise moved —
//     over the names the sheet's retained table text gives its cells. Those
//     names reached a version-10 model as each cell's preserved
//     `table:content-validation-name` (SPEC-122), which moved with its cell:
//     when the dirty sheet's cells carry them, they place the rules (the
//     retained table text keeps source-time positions), and the attribute is
//     then dropped from every cell — the writer derives it from the model
//     now, and a cell that held nothing else is no longer stored. Formulas'
//     source-time sheet names are remapped onto the record's names.
// 12 — SPEC-131: the drawing layer (`Sheet.drawings`, `Sheet.drawingPart`,
//     the images in `model.media` — xlsx: the sheet's `<drawing r:id>` part,
//     its `xdr:wsDr` children; ods: `table:shapes` and the `draw:*` elements
//     inside cells, which from this version no longer ride in
//     `Cell.preservedXml`). Every re-read gains them. The step, run for each
//     dirty sheet: xlsx — reads the `<drawing r:id>` from the sheet's
//     retained source and parses the drawing part from `opaque` (SPEC-124's
//     retained adjusters kept its anchors moving with every structural edit
//     until this version, so they fit the edited sheet), its images joining
//     the model's media; ods — lifts the drawing elements out of its cells'
//     `preservedXml` (each anchored to that cell's CURRENT position: the
//     cells moved with every edit since) and reads `table:shapes` from the
//     retained `sourceXml`. A sheet that is clean but `drawingsDirty` (a
//     later upgrade of a record made since) is re-read for its cells and
//     keeps its stored drawings; the stored `media` is kept, the re-read's
//     parts joining it.
//
// SPEC-126: a csv/tsv workbook never runs a ReaderUpgrade — its source is
// one delimited-text part, read by csv/parse.ts, which adds no element to
// the xlsx/ods readers and bumps no version. An entirely unedited one behind
// the current version is re-read through readCsvBytes with its recorded
// dialect (keeping `id`, `title` — and so the store's starred and trash
// state — and the sheet's id); any other is only re-stamped.
//
// Since SPEC-122 the writers rebuild a dirty sheet from the model, so the
// layout a record re-reads its clean sheets from numbers its styles afresh:
// each dirty sheet's style ids are moved into the re-read style table
// (interned — an equal style is shared), never left pointing into the
// stored one.
import type { MediaPart } from "../core/media";
import { decodedPartBytes } from "../core/packageParts";
import { isCellError, isCsvFormat, type Cell, type CellStyle, type CellValue, type Row, type Sheet, type SpreadsheetFormat, type SpreadsheetModel } from "./model";
import type { ParsedSpreadsheet } from "./parsed";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import { buildPreview } from "./preview";
import type { CfRule, CfValue, Dxf } from "./conditionalFormats";
import type { DataValidation } from "./validation";
import { cellStyle, defaultStyles, internDxf, internStyle, styleKey } from "./styles";

export const SPREADSHEET_READER_VERSION = 12;

/** The locale a preview built here formats in: the reader knows no
 * preference — the store rebuilds the card in the user's locale (SPEC-118). */
const READ_PREVIEW_LOCALE = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];

export interface ReaderUpgrade {
  from: number; // upgrades a record at `from` to `from + 1`
  /** Brings ONE dirty sheet up to date by reading the newly modelled element
   *  from the sheet's retained source and merging it into the model (the
   *  sheet object is updated in place; its rows stay copy-on-write).
   *  `record` is the stored model being upgraded — its sheet list, names and
   *  retained package — for a step that maps source-time names (absent: the
   *  model's own). */
  upgradeDirtySheet(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>, record?: SpreadsheetModel): void | Promise<void>;
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
  {
    // SPEC-123 (6 → 7): named cell styles and the protection's format
    // permits (see the version history).
    from: 6,
    async upgradeDirtySheet(model, sheet, sourceParts) {
      if (model.namedStyles === undefined) {
        const named = (await sourcePackageRead(model.format, sourceParts)).namedStyles;
        if (named !== undefined && named.length > 0) {
          model.namedStyles = named;
        }
      }
      const source = await retainedSheetRead(model, sheet, sourceParts);
      if (source?.protection) {
        sheet.protection = { ...source.protection };
      }
    },
  },
  {
    // SPEC-124 (7 → 8): the protection's structural permits.
    from: 7,
    async upgradeDirtySheet(model, sheet, sourceParts) {
      const source = await retainedSheetRead(model, sheet, sourceParts);
      if (source?.protection) {
        sheet.protection = { ...source.protection };
      }
    },
  },
  {
    // SPEC-128 (8 → 9): the autofilter, the sort records, filter-hidden rows
    // and the protection's sort and autofilter permits (see the version
    // history).
    from: 8,
    async upgradeDirtySheet(model, sheet, sourceParts) {
      const source = await retainedSheetRead(model, sheet, sourceParts);
      if (!source) {
        return;
      }
      if (source.protection) {
        sheet.protection = { ...source.protection };
      }
      if (source.autoFilter !== undefined && sheet.autoFilter === undefined) {
        sheet.autoFilter = source.autoFilter;
      }
      if (source.sortState !== undefined && sheet.sortState === undefined) {
        sheet.sortState = source.sortState;
      }
      // On demand, as the readers are: the store imports this module.
      const { classifyHiddenRows } = await import("./filter");
      sheet.rows = classifyHiddenRows(model, sheet, workbookFormatContext(model, READ_PREVIEW_LOCALE));
    },
  },
  {
    // SPEC-129 (9 → 10): conditional formats and the differential styles
    // (see the version history).
    from: 9,
    async upgradeDirtySheet(model, sheet, sourceParts, record) {
      if (model.dxfs === undefined) {
        model.dxfs = (await sourcePackageRead(model.format, sourceParts)).dxfs;
      }
      const source = await retainedSheetSource(model, sheet, sourceParts);
      if (!source || source.sheet.conditionalFormats.length === 0) {
        sheet.conditionalFormats = [];
        return;
      }
      const { remapSheetReferences, isEmptyRemap } = await import("./formula/sheetRefs");
      const { copyRemap, sourceSheetRemap } = await import("./sourceNames");
      const { LazyParts } = await import("./xlsx/package");
      const stored = record ?? model;
      const remap = copyRemap(sourceSheetRemap(stored, new LazyParts(stored.opaque)), sheet);
      const rename = isEmptyRemap(remap) ? (text: string) => text : (text: string) => remapSheetReferences(text, remap);
      sheet.conditionalFormats = source.sheet.conditionalFormats.map((format) => ({
        ...format,
        rules: format.rules.map((rule) => upgradedRule(rule, source.dxfs, model, rename)),
      }));
    },
  },
  {
    // SPEC-130 (10 → 11): data validations (see the version history).
    from: 10,
    async upgradeDirtySheet(model, sheet, sourceParts, record) {
      const stored = record ?? model;
      const read = await retainedValidations(model, sheet, sourceParts);
      if (model.format === "ods") {
        // The names live in the model now: a stale attribute would be written twice.
        sheet.rows = withoutValidationNames(sheet.rows);
      }
      if (read.promptsDisabled) {
        sheet.validationPromptsDisabled = true;
      }
      if (read.validations.length === 0) {
        sheet.validations = [];
        return;
      }
      const { remapSheetReferences, isEmptyRemap } = await import("./formula/sheetRefs");
      const { copyRemap, sourceSheetRemap } = await import("./sourceNames");
      const { LazyParts } = await import("./xlsx/package");
      const { mapValidationFormulas } = await import("./validation");
      const remap = copyRemap(sourceSheetRemap(stored, new LazyParts(stored.opaque)), sheet);
      sheet.validations = isEmptyRemap(remap) ? read.validations : read.validations.map((dv) => mapValidationFormulas(dv, (text) => remapSheetReferences(text, remap)));
    },
  },
  {
    // SPEC-131 (11 → 12): the drawing layer (see the version history).
    from: 11,
    async upgradeDirtySheet(model, sheet, sourceParts) {
      const media = new Map((model.media ?? []).map((part) => [part.mediaId, part] as const));
      if (model.format === "xlsx") {
        if (sheet.partName === undefined) {
          return;
        }
        const { readXlsxSheetDrawings } = await import("./xlsx/drawingXml");
        const read = readXlsxSheetDrawings(sourceParts, sheet.partName, sheet, media);
        if (read.drawings.length > 0) {
          sheet.drawings = read.drawings;
        }
        if (read.drawingPart !== undefined) {
          sheet.drawingPart = read.drawingPart;
        }
      } else {
        await upgradeOdsDrawings(model, sheet, sourceParts, media);
      }
      if (media.size > 0) {
        model.media = [...media.values()];
      }
    },
  },
];

// --- SPEC-131's step -------------------------------------------------------------------------

/** A dirty ods sheet's drawings (see the version history): lifted out of
 * its cells' preservedXml and read from its retained table:shapes. */
async function upgradeOdsDrawings(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>, media: Map<string, MediaPart>): Promise<void> {
  void model;
  void sheet;
  void sourceParts;
  void media;
}

// --- SPEC-130's step -------------------------------------------------------------------------

/** A dirty sheet's validations parsed from its retained source alone (see
 * the version history): none when it has no source to read. */
async function retainedValidations(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>): Promise<{ validations: DataValidation[]; promptsDisabled: boolean }> {
  if (model.format === "xlsx") {
    const bytes = sheet.partName === undefined ? undefined : sourceParts.get(sheet.partName);
    if (!bytes) {
      return { validations: [], promptsDisabled: false };
    }
    const { readWorksheetValidations } = await import("./xlsx/dataValidations");
    return readWorksheetValidations(new TextDecoder().decode(bytes));
  }
  const prefix = model.odsContentEnvelope?.prefix;
  if (prefix === undefined || sheet.sourceXml === undefined) {
    return { validations: [], promptsDisabled: false };
  }
  const { rangesFromCellNames, readTableValidations } = await import("./ods/contentValidations");
  const named = [...preservedValidationNames(sheet.rows)];
  const ranges = named.length > 0 ? rangesFromCellNames(named) : undefined;
  return { validations: readTableValidations(prefix, sheet.sourceXml, sheet.name, ranges), promptsDisabled: false };
}

const PRESERVED_VALIDATION_NAME = /\s[\w.-]*:?content-validation-name\s*=\s*(?:"([^"]*)"|'([^']*)')/;

/** Every stored cell's preserved `table:content-validation-name`. */
function* preservedValidationNames(rows: readonly Row[]): Generator<{ row: number; col: number; name: string }> {
  for (const row of rows) {
    for (const cell of row.cells) {
      const match = cell.preservedAttrs === undefined ? null : PRESERVED_VALIDATION_NAME.exec(cell.preservedAttrs);
      if (match) {
        yield { row: row.index, col: cell.col, name: unescapeXml(match[1] ?? match[2] ?? "") };
      }
    }
  }
}

function unescapeXml(text: string): string {
  return text.replace(/&(lt|gt|quot|apos|amp);/g, (_, entity: string) => ({ lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" })[entity]!);
}

/** The rows without any preserved `table:content-validation-name` (a cell
 * left holding nothing is dropped, a row left empty too) — the same array
 * when no cell carries one. */
function withoutValidationNames(rows: Row[]): Row[] {
  let out: Row[] | null = null;
  rows.forEach((row, at) => {
    if (!row.cells.some((cell) => cell.preservedAttrs !== undefined && PRESERVED_VALIDATION_NAME.test(cell.preservedAttrs))) {
      return;
    }
    const cells: Cell[] = [];
    for (const cell of row.cells) {
      if (cell.preservedAttrs === undefined || !PRESERVED_VALIDATION_NAME.test(cell.preservedAttrs)) {
        cells.push(cell);
        continue;
      }
      const attrs = cell.preservedAttrs.replace(new RegExp(PRESERVED_VALIDATION_NAME.source, "g"), "");
      const next: Cell = { ...cell, preservedAttrs: attrs };
      if (attrs === "") {
        delete next.preservedAttrs;
      }
      const { col: _col, value, ...rest } = next;
      if (value !== null || Object.keys(rest).length > 0) {
        cells.push(next);
      }
    }
    out ??= rows.slice();
    out[at] = { ...row, cells };
  });
  if (!out) {
    return rows;
  }
  return (out as Row[]).filter((row) => row.cells.length > 0 || row.heightPt !== undefined || row.hidden || row.filterHidden || row.styleId !== undefined || row.sourceAttrs !== undefined || row.sourceStyleName !== undefined || row.customHeight);
}

// --- SPEC-129's step -------------------------------------------------------------------------

/** A value object's formula, or undefined for a plain number (a `formula`
 * value is always one; num/percent/percentile hold a number or a formula). */
function valueFormula(value: CfValue): string | undefined {
  if (value.value === undefined) {
    return undefined;
  }
  return value.type === "formula" || !/^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?\s*$/i.test(value.value) ? value.value : undefined;
}

function renamedValue(value: CfValue, rename: (text: string) => string): CfValue {
  const formula = valueFormula(value);
  return formula === undefined ? value : { ...value, value: rename(formula) };
}

/** A rule read from a retained source, made the model's: its dxf interned
 * into the model's table (from the read's own), its formulas' sheet names
 * remapped. */
function upgradedRule(rule: CfRule, readDxfs: readonly Dxf[], model: SpreadsheetModel, rename: (text: string) => string): CfRule {
  let next: CfRule = rule;
  if (rule.dxfId !== undefined) {
    const dxf = readDxfs[rule.dxfId];
    next = dxf === undefined ? next : { ...next, dxfId: internDxf(model, dxf) };
  }
  switch (next.kind) {
    case "cellIs":
      return { ...next, formulas: next.formulas.map(rename) };
    case "expression":
      return { ...next, formula: rename(next.formula) };
    case "colorScale":
      return { ...next, stops: next.stops.map((stop) => ({ ...stop, at: renamedValue(stop.at, rename) })) };
    case "dataBar":
      return { ...next, bar: { ...next.bar, min: renamedValue(next.bar.min, rename), max: renamedValue(next.bar.max, rename) } };
    case "iconSet":
      return { ...next, icons: { ...next.icons, thresholds: next.icons.thresholds.map((threshold) => ({ ...threshold, at: renamedValue(threshold.at, rename) })) } };
    default:
      return next;
  }
}

// --- SPEC-122's step ------------------------------------------------------------------------

/** One read of a source package per upgrade, shared by its dirty sheets. */
const sourceReads = new WeakMap<Map<string, Uint8Array>, Promise<ParsedSpreadsheet>>();

/** The whole retained source package, read once per upgrade. */
function sourcePackageRead(format: SpreadsheetFormat, sourceParts: Map<string, Uint8Array>): Promise<ParsedSpreadsheet> {
  let read = sourceReads.get(sourceParts);
  if (!read) {
    read = parse(new Uint8Array(0), format, sourceParts);
    sourceReads.set(sourceParts, read);
  }
  return read;
}

/** The sheet read from its own retained source (retainedSheetRead) with the
 * dxf table its rules' ids index into — the read's own. */
async function retainedSheetSource(
  model: SpreadsheetModel,
  sheet: Sheet,
  sourceParts: Map<string, Uint8Array>,
): Promise<{ sheet: Sheet; dxfs: Dxf[] } | undefined> {
  if (model.format === "xlsx") {
    const read = await retainedSheetRead(model, sheet, sourceParts);
    return read ? { sheet: read, dxfs: (await sourcePackageRead("xlsx", sourceParts)).dxfs } : undefined;
  }
  const envelope = model.odsContentEnvelope;
  if (sheet.sourceXml === undefined || !envelope) {
    return undefined;
  }
  const parts = new Map(sourceParts);
  parts.set("content.xml", new TextEncoder().encode(envelope.prefix + sheet.sourceXml + envelope.suffix));
  const parsed = await parse(new Uint8Array(0), "ods", parts);
  return parsed.sheets[0] ? { sheet: parsed.sheets[0], dxfs: parsed.dxfs } : undefined;
}

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
    const parsed = await sourcePackageRead("xlsx", sourceParts);
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
 * interned. Arrays are replaced only when an id changes. `beforeNamedStyles`:
 * the record predates SPEC-123, so a stored style lacking only the
 * `namedStyle` its re-read source entry has is that entry. */
function adoptStyles(sheet: Sheet, from: readonly CellStyle[], into: SpreadsheetModel, beforeNamedStyles: boolean): void {
  const ids = new Map<number, number>();
  const sameEntry = (same: CellStyle, style: CellStyle): boolean => {
    if (same.sourceXf !== style.sourceXf || same.sourceStyleName !== style.sourceStyleName) {
      return false;
    }
    if (styleKey(same) === styleKey(style)) {
      return true;
    }
    const { namedStyle: _named, ...unnamed } = same;
    return beforeNamedStyles && style.namedStyle === undefined && (same.sourceXf !== undefined || same.sourceStyleName !== undefined) && styleKey(unnamed) === styleKey(style);
  };
  const mapId = (id: number): number => {
    let next = ids.get(id);
    if (next === undefined) {
      const style = cellStyle({ styles: from as CellStyle[] }, id);
      const same = into.styles[id];
      next = same !== undefined && sameEntry(same, style) ? id : internStyle(into, style);
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
    dxfs: parsed.dxfs,
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
  if (parsed.odsOrphanValidations !== undefined && parsed.odsOrphanValidations.length > 0) {
    model.odsOrphanValidations = parsed.odsOrphanValidations;
  }
  if (parsed.media !== undefined && parsed.media.length > 0) {
    model.media = parsed.media;
  }
  model.preview = buildPreview(model, READ_PREVIEW_LOCALE);
  return model;
}

async function parse(
  bytes: Uint8Array,
  format: SpreadsheetFormat,
  parts: Map<string, Uint8Array> | undefined,
): Promise<ParsedSpreadsheet> {
  switch (format) {
    case "xlsx": {
      const { parseXlsx } = await import("./xlsx/reader");
      return parseXlsx(bytes, parts);
    }
    case "ods": {
      const { parseOds } = await import("./ods/reader");
      return parseOds(bytes, parts);
    }
    case "csv":
    case "tsv":
      // A delimited-text file is no package: csv/parse.ts's readCsvBytes
      // reads it (upgradeCsvModel below).
      throw new Error(`a .${format} workbook is read by readCsvBytes, not as a package`);
  }
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
  if (isCsvFormat(model.format)) {
    return upgradeCsvModel(model, model.format);
  }
  const sourceParts = decodedPartBytes(model.opaque);
  const pending = READER_UPGRADES.filter((step) => step.from >= model.readerVersion).toSorted((a, b) => a.from - b.from);
  if (model.structureDirty || model.sheets.some((sheet) => sheet.dirty || sheet.drawingsDirty)) {
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

/** SPEC-126: a csv/tsv workbook behind the reader version (see the module
 * header) — re-read with its recorded dialect when nothing was edited (one
 * sheet, the one read from the file, clean, no structure change), else only
 * re-stamped. */
async function upgradeCsvModel(model: SpreadsheetModel, format: "csv" | "tsv"): Promise<SpreadsheetModel> {
  const dialect = model.csv;
  const sheet = model.sheets[0];
  const unedited = dialect !== undefined && model.sheets.length === 1 && sheet?.id === dialect.sheetId && !sheet.dirty && !model.structureDirty;
  if (!unedited) {
    // Re-stamped, never re-read: SPEC-129's and SPEC-130's fields start empty, as a read
    // of a delimited-text file leaves them.
    const sheets = model.sheets.some((each) => each.conditionalFormats === undefined || each.validations === undefined)
      ? model.sheets.map((each) => (each.conditionalFormats === undefined || each.validations === undefined ? { ...each, conditionalFormats: each.conditionalFormats ?? [], validations: each.validations ?? [] } : each))
      : model.sheets;
    return { ...model, sheets, dxfs: model.dxfs ?? [], readerVersion: SPREADSHEET_READER_VERSION };
  }
  const { csvSourceBytes, readCsvBytes } = await import("./csv/parse");
  const fresh = readCsvBytes(csvSourceBytes(model), format, model.title, {
    userLocaleTag: DEFAULT_SPREADSHEET_LOCALE,
    delimiter: dialect.delimiter,
    encoding: dialect.encoding,
    localeTag: dialect.localeTag,
  });
  fresh.id = model.id;
  fresh.sheets = fresh.sheets.map((each) => ({ ...each, id: sheet.id, name: sheet.name }));
  fresh.csv = { ...fresh.csv!, sheetId: sheet.id };
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
  // SPEC-131: the stored images (an inserted one has no retained part) win
  // over the re-read's, which add the ones a re-read sheet shows.
  const media = new Map([...(fresh.media ?? []), ...(model.media ?? [])].map((part) => [part.mediaId, part] as const));
  if (media.size > 0) {
    fresh.media = [...media.values()];
  } else {
    delete fresh.media;
  }
  if (model.format === "ods" && model.readerVersion < 11 && model.odsContentEnvelope) {
    // SPEC-130: the substitution layout wrote no orphan definitions (the
    // record had none modelled); they are the retained prefix's definitions
    // no upgraded sheet's rule names.
    const { readOdsContentValidations } = await import("./ods/contentValidations");
    const regions = sheets.map((sheet) => new Map((sheet.validations ?? []).flatMap((rule) => (rule.odsName === undefined ? [] : [[rule.odsName, rule.ranges] as const]))));
    const { orphans } = readOdsContentValidations(model.odsContentEnvelope.prefix, regions);
    if (orphans.length > 0) {
      fresh.odsOrphanValidations = orphans;
    } else {
      delete fresh.odsOrphanValidations;
    }
  }
  fresh.preview = buildPreview(fresh, READ_PREVIEW_LOCALE);
  return fresh;

  async function upgradedSheet(stored: Sheet, index: number): Promise<Sheet> {
    if (stored.dirty) {
      if (stored.partName !== undefined || stored.sourceXml !== undefined) {
        for (const step of pending) {
          await step.upgradeDirtySheet(fresh, stored, sourceParts, model);
        }
      }
      adoptStyles(stored, model.styles, fresh, model.readerVersion < 7);
      // A sheet born after the read (no source) had no rules to read.
      stored.validations ??= [];
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
    if (stored.drawingsDirty) {
      // SPEC-131: the edited drawing layer is the model's, not the source's.
      delete sheet.drawings;
      delete sheet.drawingPart;
      Object.assign(sheet, pick(stored, ["drawings", "drawingPart", "drawingsDirty"]));
    }
    return sheet;
  }
}

/** The fields of `from` named in `keys` that it has. */
function pick<T extends object, K extends keyof T>(from: T, keys: readonly K[]): Partial<Pick<T, K>> {
  const out: Partial<Pick<T, K>> = {};
  for (const key of keys) {
    if (from[key] !== undefined) {
      out[key] = from[key];
    }
  }
  return out;
}
