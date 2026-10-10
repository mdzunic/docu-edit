// What both workbook readers hand back (SPEC-116), and the bookkeeping they
// share: the cell ceiling, the out-of-grid parse error and the verbatim
// opaque copy of the source package. open.ts assembles a ParsedSpreadsheet
// into a SpreadsheetModel.
import type { DocumentProperties } from "../core/documentProperties";
import { bytesToBase64Async, type OpaquePart } from "../core/packageParts";
import type { DocumentTheme } from "../core/theme";
import { colToLetters } from "./address";
import { SpreadsheetOpenError } from "./errors";
import { MAX_WORKBOOK_CELLS, type CalcSettings, type CellStyle, type PackageVariant, type Sheet } from "./model";
import type { NamedCellStyle } from "./styles";

export interface ParsedSpreadsheet {
  sheets: Sheet[];
  /** SPEC-117: the cell styles, index 0 the default — never empty. */
  styles: CellStyle[];
  theme?: DocumentTheme;
  indexedColors?: string[];
  activeSheetIndex: number;
  dateSystem: "1900" | "1904";
  odsNullDate?: string;
  odsContentEnvelope?: { prefix: string; suffix: string };
  packageVariant?: PackageVariant;
  opaque: OpaquePart[];
  properties: DocumentProperties;
  /** SPEC-120: the file's calculation settings (absent: none set). */
  calc?: CalcSettings;
  /** SPEC-122: the workbook's structure is protected. */
  structureProtected?: true;
  /** SPEC-123: the named cell styles (absent: the file defines none). */
  namedStyles?: NamedCellStyle[];
}

/** Counts stored cells across every sheet of one read against
 * MAX_WORKBOOK_CELLS. `take` answers false — instead of throwing — once a
 * take would pass the ceiling, so a scan callback can return "stop" and the
 * reader refuses the workbook without materialising the remainder. */
export class CellBudget {
  private count = 0;

  take(cells: number): boolean {
    if (this.count + cells > MAX_WORKBOOK_CELLS) {
      return false;
    }
    this.count += cells;
    return true;
  }

  get used(): number {
    return this.count;
  }
}

/** The parse error for a cell or row past the 1,048,576 × 16,384 grid —
 * naming its address, never a silent drop. */
export function outOfGridError(sheetName: string, address: string): SpreadsheetOpenError {
  return new SpreadsheetOpenError(
    `This workbook can't be opened: "${sheetName}" has content at ${address}, outside the 1,048,576 × 16,384 sheet grid`,
  );
}

/** A malformed-content parse error, naming where. */
export function parseError(where: string, problem: string): SpreadsheetOpenError {
  return new SpreadsheetOpenError(`This workbook can't be opened: ${problem} (${where})`);
}

/** "B3" for (row 2, col 1), for messages — any column, even past XFD. */
export function addressText(row: number, col: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

/** The whole source package, verbatim and in source order, as the model's
 * opaque parts (base64 keeps the model plain JSON). */
export async function encodeOpaque(parts: Map<string, Uint8Array>): Promise<OpaquePart[]> {
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    opaque.push({ name, data: await bytesToBase64Async(data) });
  }
  return opaque;
}

const utf8 = new TextDecoder("utf-8");

/** A part's text (UTF-8, BOM dropped by the decoder). */
export function partText(bytes: Uint8Array): string {
  return utf8.decode(bytes);
}

/** The index of the sheet a file names as active, clamped into range, and
 * moved to the first visible sheet when it names a hidden one. */
export function resolveActiveSheet(requested: number, sheets: readonly Sheet[]): number {
  if (sheets.length === 0) {
    return 0;
  }
  const clamped = Number.isInteger(requested) ? Math.min(Math.max(requested, 0), sheets.length - 1) : 0;
  if (sheets[clamped]!.visibility === "visible") {
    return clamped;
  }
  const firstVisible = sheets.findIndex((sheet) => sheet.visibility === "visible");
  return firstVisible >= 0 ? firstVisible : clamped;
}

/** Restores the sparse invariants (rows ascending by index, cells ascending
 * by column, both unique — a later duplicate wins) for the rare file that
 * writes them out of order. Rows whose cells all merge away are dropped. */
export function normaliseRows(rows: { index: number; cells: { col: number }[] }[]): void {
  rows.sort((a, b) => a.index - b.index);
  for (let i = rows.length - 1; i > 0; i -= 1) {
    if (rows[i]!.index === rows[i - 1]!.index) {
      rows[i - 1]!.cells.push(...rows[i]!.cells);
      rows.splice(i, 1);
    }
  }
  for (const row of rows) {
    normaliseCells(row.cells);
  }
}

/** Sorts a row's cells by column, keeping the last of any duplicates. */
export function normaliseCells(cells: { col: number }[]): void {
  const byCol = new Map<number, { col: number }>();
  for (const cell of cells) {
    byCol.set(cell.col, cell);
  }
  const sorted = [...byCol.values()].toSorted((a, b) => a.col - b.col);
  cells.length = 0;
  cells.push(...sorted);
}
