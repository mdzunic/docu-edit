// Sheet page setup (SPEC-136 "The page-setup model") — THIS FILE is the
// normative home of the model's defaults, the xlsx paper catalogue, the
// page-setup commands, the print area / print titles accessors and the
// `pageBreaks` reference adjuster.
//
// - Defaults. An absent field means: portrait; the paper of effectivePaper;
//   100%; DEFAULT_SHEET_MARGINS (Excel's "Normal"); no centring, gridlines or
//   headings; down, then over; automatic numbering; no header or footer; no
//   manual breaks. normalizePageSetup drops every field equal to its default
//   (a setup left with none is no setup at all), so a sheet whose file says
//   nothing about printing has no `pageSetup`, and a command that resets a
//   field writes the same model a file without it reads as.
// - Paper. `paper` holds portrait dimensions plus the xlsx `paperSize` code
//   it came from (XLSX_PAPER_SIZES maps codes 1, 3, 5, 8, 9, 11 onto
//   core/pageSetup's PAPER_SIZES and 7, 12, 13, 14, 66, 70 by dimension; an
//   unknown code is kept and paginates as Letter). Absent `paper`: the
//   user's default page size preference (the view passes it in); an xlsx
//   `<pageSetup>` without `paperSize` reads as Letter, the schema default.
// - Commands replace `Sheet.pageSetup`, never mutate it, and mark the sheet
//   through markPageSetupEdited: an .xlsx sheet `pageSetupDirty` (the writer
//   splices its page elements, xlsx/pageSetup.ts); an .ods layout edit
//   `pageSetupDirty` too (the writer splices styles.xml and the table's start
//   tag, ods/pageSetup.ts), an .ods break or print-titles edit `dirty` (both
//   live in the table's row and column structure).
// - Print area and print titles are SPEC-133's sheet-scoped reserved names
//   `_xlnm.Print_Area` and `_xlnm.Print_Titles` — printAreaOf/printTitlesOf/
//   setPrintArea/setPrintTitles are the one accessor every surface uses.
//   They have no adjuster here: SPEC-133's `names` adjusters move them.
// - `pageBreaks` (REFERENCE_ADJUSTERS, SPEC-124): a break at or after an
//   inserted band shifts by `count`; a break inside a deleted band moves to
//   `at`; the result deduplicated. No MoveAdjuster: a break belongs to a row
//   or column boundary, not to cells, so a cut-paste leaves it (as Excel).
import { PAPER_SIZES, type PageSize } from "../core/pageSetup";
import type { CellRef, GridRange } from "./address";
import { formatReference } from "./formula/print";
import { tokenize } from "./formula/tokens";
import { MAX_COLS, MAX_ROWS, type Sheet, type SheetHeaderFooter, type SheetMargins, type SheetPageSetup, type SpreadsheetModel } from "./model";
import { getReservedName, setReservedName } from "./names";
import { adjustIndex, editAxis, type StructuralEdit } from "./structure";

/** Excel's "Normal" margins: top/bottom 0.75 in, left/right 0.7 in,
 * header/footer 0.3 in. */
export const DEFAULT_SHEET_MARGINS: Readonly<SheetMargins> = {
  topCm: 1.905,
  rightCm: 1.778,
  bottomCm: 1.905,
  leftCm: 1.778,
  headerCm: 0.762,
  footerCm: 0.762,
};

export const MIN_PRINT_SCALE = 10;
export const MAX_PRINT_SCALE = 400;
export const PRINT_SCALE_MESSAGE = "Enter a scale between 10% and 400%";

export interface XlsxPaperSize {
  code: number;
  name: string;
  widthCm: number;
  heightCm: number;
}

/** The xlsx `paperSize` codes this product names, portrait. */
export const XLSX_PAPER_SIZES: readonly XlsxPaperSize[] = [
  { code: 1, name: "Letter", ...PAPER_SIZES.Letter },
  { code: 3, name: "Tabloid", ...PAPER_SIZES.Tabloid },
  { code: 5, name: "Legal", ...PAPER_SIZES.Legal },
  { code: 7, name: "Executive", widthCm: 18.415, heightCm: 26.67 },
  { code: 8, name: "A3", ...PAPER_SIZES.A3 },
  { code: 9, name: "A4", ...PAPER_SIZES.A4 },
  { code: 11, name: "A5", ...PAPER_SIZES.A5 },
  { code: 12, name: "B4", widthCm: 25, heightCm: 35.3 },
  { code: 13, name: "B5", widthCm: 17.6, heightCm: 25 },
  { code: 14, name: "Folio", widthCm: 21.59, heightCm: 33.02 },
  { code: 66, name: "A2", widthCm: 42, heightCm: 59.4 },
  { code: 70, name: "A6", widthCm: 10.5, heightCm: 14.8 },
];

export const CUSTOM_PAPER_NAME = "Custom";
export const LETTER_PAPER_CODE = 1;

type Paper = NonNullable<SheetPageSetup["paper"]>;

const PAPER_TOLERANCE_CM = 0.05;

function sameSize(a: PageSize, b: PageSize): boolean {
  return Math.abs(a.widthCm - b.widthCm) < PAPER_TOLERANCE_CM && Math.abs(a.heightCm - b.heightCm) < PAPER_TOLERANCE_CM;
}

/** The catalogue entry of an xlsx code. */
export function xlsxPaperSize(code: number): XlsxPaperSize | undefined {
  return XLSX_PAPER_SIZES.find((entry) => entry.code === code);
}

/** The paper an xlsx `paperSize` code means: its catalogue size, or — an
 * unknown code, kept — Letter's. */
export function paperForXlsxCode(code: number): Paper {
  const entry = xlsxPaperSize(code) ?? xlsxPaperSize(LETTER_PAPER_CODE)!;
  return { widthCm: entry.widthCm, heightCm: entry.heightCm, xlsxPaperSize: code };
}

/** The catalogue entry a paper is (by its code, else by its portrait size). */
export function paperEntryOf(paper: PageSize & { xlsxPaperSize?: number }): XlsxPaperSize | undefined {
  if (paper.xlsxPaperSize !== undefined) {
    const entry = xlsxPaperSize(paper.xlsxPaperSize);
    if (entry !== undefined) {
      return entry;
    }
  }
  return XLSX_PAPER_SIZES.find((entry) => sameSize(entry, paper));
}

/** The xlsx code a paper writes as: its own, else its size's; undefined for
 * a custom size (written as `paperWidth`/`paperHeight`). */
export function xlsxCodeOfPaper(paper: Paper): number | undefined {
  return paper.xlsxPaperSize ?? paperEntryOf(paper)?.code;
}

/** The dialog's name for a paper ("Custom" for none of the catalogue's). */
export function paperNameOf(paper: Paper | undefined, fallback: PageSize): string {
  return paperEntryOf(paper ?? fallback)?.name ?? CUSTOM_PAPER_NAME;
}

/** The paper a catalogue name selects. */
export function paperForName(name: string): Paper | null {
  const entry = XLSX_PAPER_SIZES.find((candidate) => candidate.name === name);
  return entry ? { widthCm: entry.widthCm, heightCm: entry.heightCm, xlsxPaperSize: entry.code } : null;
}

/** What a sheet prints on when its setup names no paper. */
export interface PrintPreferences {
  /** SPEC-066's default page size ("A4", "Letter", …). */
  defaultPageSizeId?: string;
}

/** The portrait paper a sheet prints on (see the module header). */
export function effectivePaper(sheet: Pick<Sheet, "pageSetup">, preferences: PrintPreferences = {}): PageSize {
  const paper = sheet.pageSetup?.paper;
  if (paper !== undefined) {
    return { widthCm: paper.widthCm, heightCm: paper.heightCm };
  }
  const preferred = PAPER_SIZES[(preferences.defaultPageSizeId ?? "A4") as keyof typeof PAPER_SIZES] ?? PAPER_SIZES.A4;
  return { ...preferred };
}

/** The paper as it lies: landscape swaps the portrait dimensions. */
export function orientedPaper(sheet: Pick<Sheet, "pageSetup">, preferences: PrintPreferences = {}): PageSize {
  const paper = effectivePaper(sheet, preferences);
  return sheet.pageSetup?.orientation === "landscape" ? { widthCm: paper.heightCm, heightCm: paper.widthCm } : paper;
}

export function effectiveMargins(setup: SheetPageSetup | undefined): SheetMargins {
  return setup?.margins ?? { ...DEFAULT_SHEET_MARGINS };
}

// --- normalising -------------------------------------------------------------------------------

const MARGIN_KEYS: readonly (keyof SheetMargins)[] = ["topCm", "rightCm", "bottomCm", "leftCm", "headerCm", "footerCm"];

function sameMargins(a: SheetMargins, b: SheetMargins): boolean {
  return MARGIN_KEYS.every((key) => Math.abs(a[key] - b[key]) < 1e-9);
}

function sortedUnique(values: readonly number[]): number[] {
  return [...new Set(values.filter((value) => Number.isInteger(value) && value > 0))].sort((a, b) => a - b);
}

function normalizeHeaderFooter(hf: SheetHeaderFooter): SheetHeaderFooter | undefined {
  const pair = (value: SheetHeaderFooter["odd"] | undefined): SheetHeaderFooter["odd"] | undefined => {
    if (value === undefined) {
      return undefined;
    }
    const out: SheetHeaderFooter["odd"] = {};
    for (const key of ["header", "footer"] as const) {
      const sections = value[key];
      if (sections === undefined) {
        continue;
      }
      const kept = Object.fromEntries(Object.entries(sections).filter(([, text]) => text !== undefined && text !== ""));
      if (Object.keys(kept).length > 0) {
        out[key] = kept;
      }
    }
    return out;
  };
  const next: SheetHeaderFooter = { odd: pair(hf.odd) ?? {} };
  const even = pair(hf.even);
  const first = pair(hf.first);
  if (even !== undefined) {
    next.even = even;
  }
  if (first !== undefined) {
    next.first = first;
  }
  if (hf.differentOddEven) {
    next.differentOddEven = true;
  }
  if (hf.differentFirst) {
    next.differentFirst = true;
  }
  if (hf.scaleWithDoc === false) {
    next.scaleWithDoc = false;
  }
  if (hf.alignWithMargins === false) {
    next.alignWithMargins = false;
  }
  const empty =
    Object.keys(next.odd).length === 0 &&
    next.even === undefined &&
    next.first === undefined &&
    !next.differentOddEven &&
    !next.differentFirst &&
    next.scaleWithDoc === undefined &&
    next.alignWithMargins === undefined;
  return empty ? undefined : next;
}

/** A setup with every default-valued field dropped (see the module header);
 * undefined when none is left. Never mutates its argument. */
export function normalizePageSetup(setup: SheetPageSetup | undefined): SheetPageSetup | undefined {
  if (setup === undefined) {
    return undefined;
  }
  const next: SheetPageSetup = {};
  if (setup.orientation === "landscape") {
    next.orientation = "landscape";
  }
  if (setup.paper !== undefined) {
    next.paper = { ...setup.paper };
  }
  if (setup.scale !== undefined && setup.scale !== 100) {
    next.scale = setup.scale;
  }
  if (setup.fitTo !== undefined) {
    next.fitTo = { ...setup.fitTo };
  }
  if (setup.odsScaleToPages !== undefined) {
    next.odsScaleToPages = setup.odsScaleToPages;
  }
  if (setup.margins !== undefined && !sameMargins(setup.margins, DEFAULT_SHEET_MARGINS)) {
    next.margins = { ...setup.margins };
  }
  for (const flag of ["centerHorizontally", "centerVertically", "printGridLines", "printHeadings"] as const) {
    if (setup[flag]) {
      next[flag] = true;
    }
  }
  if (setup.pageOrder === "overThenDown") {
    next.pageOrder = "overThenDown";
  }
  if (setup.firstPageNumber !== undefined) {
    next.firstPageNumber = setup.firstPageNumber;
  }
  const headerFooter = setup.headerFooter === undefined ? undefined : normalizeHeaderFooter(setup.headerFooter);
  if (headerFooter !== undefined) {
    next.headerFooter = headerFooter;
  }
  const rowBreaks = sortedUnique(setup.rowBreaks ?? []);
  if (rowBreaks.length > 0) {
    next.rowBreaks = rowBreaks;
  }
  const colBreaks = sortedUnique(setup.colBreaks ?? []);
  if (colBreaks.length > 0) {
    next.colBreaks = colBreaks;
  }
  return Object.keys(next).length === 0 ? undefined : next;
}

/** Plain-JSON equality of two setups (the model never holds anything else). */
export function samePageSetup(a: SheetPageSetup | undefined, b: SheetPageSetup | undefined): boolean {
  return JSON.stringify(normalizePageSetup(a) ?? null) === JSON.stringify(normalizePageSetup(b) ?? null);
}

// --- marking and installing -------------------------------------------------------------------

/** What a page-setup edit touched. */
export type PageSetupEdit = "layout" | "breaks" | "printArea" | "printTitles";

/** Installs `sheet` (by id) marked for what the edit touched (see the
 * module header); returns the installed sheet. A print-area edit marks
 * nothing here: the name it sets is `namesDirty`. */
export function markPageSetupEdited(model: SpreadsheetModel, sheet: Sheet, what: PageSetupEdit): Sheet {
  let marked = sheet;
  if (model.format === "xlsx" && (what === "layout" || what === "breaks")) {
    marked = sheet.pageSetupDirty ? sheet : { ...sheet, pageSetupDirty: true };
  } else if (model.format === "ods") {
    if (what === "layout") {
      marked = sheet.pageSetupDirty ? sheet : { ...sheet, pageSetupDirty: true };
    } else if (what === "breaks" || what === "printTitles") {
      marked = sheet.dirty ? sheet : { ...sheet, dirty: true };
    }
  }
  model.sheets = model.sheets.map((each) => (each.id === sheet.id ? marked : each));
  return marked;
}

/** A sheet with its setup replaced (removed when undefined). */
export function withPageSetup(sheet: Sheet, setup: SheetPageSetup | undefined): Sheet {
  const { pageSetup: _old, ...rest } = sheet;
  return setup === undefined ? rest : { ...rest, pageSetup: setup };
}

function currentSheet(model: SpreadsheetModel, sheet: Sheet | string): Sheet | undefined {
  const id = typeof sheet === "string" ? sheet : sheet.id;
  return model.sheets.find((each) => each.id === id);
}

/** A change to a setup: a field set to `undefined` is reset to its default. */
export type PageSetupPatch = { [K in keyof SheetPageSetup]?: SheetPageSetup[K] | undefined };

/** Applies `patch` to a sheet's setup; marks what it touched. False when
 * nothing changed. */
export function setPageSetup(model: SpreadsheetModel, sheet: Sheet | string, patch: PageSetupPatch): boolean {
  const target = currentSheet(model, sheet);
  if (target === undefined) {
    return false;
  }
  const merged: SheetPageSetup = { ...target.pageSetup };
  for (const [key, value] of Object.entries(patch) as [keyof SheetPageSetup, unknown][]) {
    if (value === undefined) {
      delete merged[key];
    } else {
      (merged as Record<string, unknown>)[key] = structuredClone(value);
    }
  }
  const next = normalizePageSetup(merged);
  if (samePageSetup(next, target.pageSetup)) {
    return false;
  }
  const breaksChanged =
    JSON.stringify([next?.rowBreaks, next?.colBreaks]) !== JSON.stringify([target.pageSetup?.rowBreaks, target.pageSetup?.colBreaks]);
  const layoutChanged = !samePageSetup({ ...next, rowBreaks: [], colBreaks: [] }, { ...target.pageSetup, rowBreaks: [], colBreaks: [] });
  const installed = withPageSetup(target, next);
  model.sheets = model.sheets.map((each) => (each.id === target.id ? installed : each));
  if (layoutChanged) {
    markPageSetupEdited(model, installed, "layout");
  }
  if (breaksChanged) {
    markPageSetupEdited(model, currentSheet(model, target.id)!, "breaks");
  }
  return true;
}

// --- manual breaks ---------------------------------------------------------------------------

/** Excel's Insert ▸ Page break needs a cell other than A1. */
export function canInsertPageBreak(at: CellRef): boolean {
  return at.row > 0 || at.col > 0;
}

/** The manual breaks that sit at the active cell: a row break above its
 * row, a column break left of its column. */
export function pageBreaksAt(sheet: Pick<Sheet, "pageSetup">, at: CellRef): { row: boolean; col: boolean } {
  return {
    row: sheet.pageSetup?.rowBreaks?.includes(at.row) === true,
    col: sheet.pageSetup?.colBreaks?.includes(at.col) === true,
  };
}

export function hasManualBreaks(sheet: Pick<Sheet, "pageSetup">): boolean {
  return (sheet.pageSetup?.rowBreaks?.length ?? 0) > 0 || (sheet.pageSetup?.colBreaks?.length ?? 0) > 0;
}

/** Insert ▸ Page break (Excel's rule by active cell): in row 1, a column
 * break left of the active column; in column A, a row break above the
 * active row; elsewhere both. Nothing at A1. Manual breaks are ignored
 * while the sheet fits to pages, so — as Excel does — inserting one there
 * switches it to "Adjust to" the percent the fit printed at (`fitScale`,
 * print.ts's choice; 100 when not given). */
export function insertPageBreak(model: SpreadsheetModel, sheet: Sheet | string, at: CellRef, fitScale?: number): boolean {
  const target = currentSheet(model, sheet);
  if (target === undefined || !canInsertPageBreak(at)) {
    return false;
  }
  const setup = target.pageSetup ?? {};
  const patch: PageSetupPatch = {};
  if (setup.fitTo !== undefined || setup.odsScaleToPages !== undefined) {
    patch.fitTo = undefined;
    patch.odsScaleToPages = undefined;
    patch.scale = fitScale ?? 100;
  }
  if (at.row > 0) {
    patch.rowBreaks = [...(setup.rowBreaks ?? []), at.row];
  }
  if (at.col > 0) {
    patch.colBreaks = [...(setup.colBreaks ?? []), at.col];
  }
  return setPageSetup(model, target, patch);
}

/** Remove page break: the manual breaks at the active cell (pageBreaksAt). */
export function removePageBreak(model: SpreadsheetModel, sheet: Sheet | string, at: CellRef): boolean {
  const target = currentSheet(model, sheet);
  if (target === undefined) {
    return false;
  }
  const setup = target.pageSetup ?? {};
  return setPageSetup(model, target, {
    rowBreaks: (setup.rowBreaks ?? []).filter((row) => row !== at.row),
    colBreaks: (setup.colBreaks ?? []).filter((col) => col !== at.col),
  });
}

/** Reset all page breaks: every manual break goes. */
export function resetPageBreaks(model: SpreadsheetModel, sheet: Sheet | string): boolean {
  return setPageSetup(model, sheet, { rowBreaks: undefined, colBreaks: undefined });
}

// --- print area and print titles ---------------------------------------------------------------

export const PRINT_AREA = "Print_Area";
export const PRINT_TITLES = "Print_Titles";

export interface PrintTitles {
  rows?: [number, number];
  cols?: [number, number];
}

/** Every reference of a name's formula as a range (whole rows and columns
 * reaching the grid's edges); unparseable text gives none. */
export function referenceRanges(formula: string): GridRange[] {
  const ranges: GridRange[] = [];
  for (const token of tokenize(formula)) {
    if (token.kind !== "ref" || token.ref === undefined) {
      continue;
    }
    const { start } = token.ref;
    const end = token.ref.end ?? start;
    const top = Math.min(start.row ?? 0, end.row ?? 0);
    const bottom = start.row === undefined || end.row === undefined ? MAX_ROWS - 1 : Math.max(start.row, end.row);
    const left = Math.min(start.col ?? 0, end.col ?? 0);
    const right = start.col === undefined || end.col === undefined ? MAX_COLS - 1 : Math.max(start.col, end.col);
    ranges.push({ top: start.row === undefined || end.row === undefined ? 0 : top, left: start.col === undefined || end.col === undefined ? 0 : left, bottom, right });
  }
  return ranges;
}

/** The sheet's print area: its `_xlnm.Print_Area` name's ranges, null when
 * it has none (or one that names no range). */
export function printAreaOf(model: Pick<SpreadsheetModel, "names">, sheet: Pick<Sheet, "id">): GridRange[] | null {
  const name = getReservedName(model, sheet.id, PRINT_AREA);
  if (name === undefined) {
    return null;
  }
  const ranges = referenceRanges(name.formula);
  return ranges.length === 0 ? null : ranges;
}

/** The sheet's print titles: the rows and columns of its
 * `_xlnm.Print_Titles` name (the first whole-row and first whole-column
 * reference), null when it has none. */
export function printTitlesOf(model: Pick<SpreadsheetModel, "names">, sheet: Pick<Sheet, "id">): PrintTitles | null {
  const name = getReservedName(model, sheet.id, PRINT_TITLES);
  if (name === undefined) {
    return null;
  }
  const titles: PrintTitles = {};
  for (const range of referenceRanges(name.formula)) {
    if (range.left === 0 && range.right === MAX_COLS - 1 && titles.rows === undefined) {
      titles.rows = [range.top, range.bottom];
    } else if (range.top === 0 && range.bottom === MAX_ROWS - 1 && titles.cols === undefined) {
      titles.cols = [range.left, range.right];
    }
  }
  return titles.rows === undefined && titles.cols === undefined ? null : titles;
}

/** The canonical text of a print area: absolute, sheet-qualified, the
 * ranges comma-joined (Excel's own form, no parentheses). */
export function printAreaFormula(sheetName: string, ranges: readonly GridRange[]): string {
  return ranges
    .map((range) =>
      formatReference({
        sheet: sheetName,
        start: { row: range.top, col: range.left, rowAbs: true, colAbs: true },
        ...(range.top === range.bottom && range.left === range.right ? {} : { end: { row: range.bottom, col: range.right, rowAbs: true, colAbs: true } }),
      }),
    )
    .join(",");
}

/** The canonical text of print titles: `Sheet1!$1:$2,Sheet1!$A:$A`. */
export function printTitlesFormula(sheetName: string, titles: PrintTitles): string {
  const parts: string[] = [];
  if (titles.rows !== undefined) {
    parts.push(formatReference({ sheet: sheetName, start: { row: titles.rows[0], rowAbs: true, colAbs: false }, end: { row: titles.rows[1], rowAbs: true, colAbs: false } }));
  }
  if (titles.cols !== undefined) {
    parts.push(formatReference({ sheet: sheetName, start: { col: titles.cols[0], rowAbs: false, colAbs: true }, end: { col: titles.cols[1], rowAbs: false, colAbs: true } }));
  }
  return parts.join(",");
}

/** File ▸ Print area ▸ Set print area (the selection's ranges) / Clear
 * print area (null). False when nothing changed. */
export function setPrintArea(model: SpreadsheetModel, sheet: Sheet | string, ranges: readonly GridRange[] | null): boolean {
  const target = currentSheet(model, sheet);
  if (target === undefined) {
    return false;
  }
  const names = model.names;
  setReservedName(model, target.id, PRINT_AREA, ranges === null || ranges.length === 0 ? null : printAreaFormula(target.name, ranges));
  if (model.names === names) {
    return false;
  }
  markPageSetupEdited(model, target, "printArea");
  return true;
}

/** Sets (null: clears) the sheet's print titles. False when nothing changed. */
export function setPrintTitles(model: SpreadsheetModel, sheet: Sheet | string, titles: PrintTitles | null): boolean {
  const target = currentSheet(model, sheet);
  if (target === undefined) {
    return false;
  }
  const names = model.names;
  const empty = titles === null || (titles.rows === undefined && titles.cols === undefined);
  setReservedName(model, target.id, PRINT_TITLES, empty ? null : printTitlesFormula(target.name, titles));
  if (model.names === names) {
    return false;
  }
  markPageSetupEdited(model, target, "printTitles");
  return true;
}

// --- the pageBreaks adjuster (SPEC-124 REFERENCE_ADJUSTERS) ---------------------------------------

/** Breaks along the edit's axis moved per the module header. */
export function adjustBreaks(breaks: readonly number[], edit: StructuralEdit): number[] {
  const moved = breaks.flatMap((index) => {
    const next = adjustIndex(index, edit);
    if (next !== null) {
      return [next];
    }
    return edit.kind === "deleteRows" || edit.kind === "deleteCols" ? [edit.at] : [];
  });
  return sortedUnique(moved);
}

/** SPEC-124's "pageBreaks" adjuster: the edited sheet's manual breaks. */
export function adjustPageBreaksForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  const sheet = model.sheets.find((each) => each.id === edit.sheetId);
  const setup = sheet?.pageSetup;
  if (sheet === undefined || setup === undefined) {
    return;
  }
  const key = editAxis(edit) === "row" ? "rowBreaks" : "colBreaks";
  const breaks = setup[key];
  if (breaks === undefined || breaks.length === 0) {
    return;
  }
  const moved = adjustBreaks(breaks, edit);
  if (JSON.stringify(moved) === JSON.stringify(breaks)) {
    return;
  }
  const next = normalizePageSetup({ ...setup, [key]: moved });
  const installed = withPageSetup(sheet, next);
  model.sheets = model.sheets.map((each) => (each.id === sheet.id ? installed : each));
  markPageSetupEdited(model, installed, "breaks");
}
