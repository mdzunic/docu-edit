// In-memory workbooks for the calculation specs (SPEC-120): a model built
// from a map of A1 addresses to values and formulas, so a spec states its
// inputs the way the conformance table does ("A1 the text "5"") instead of
// assembling rows and cells by hand. DOM-free, like build.ts beside it.
import { parseA1, parseRange } from "../address";
import type { CalcSettings, Cell, CellValue, Row, Sheet, SpreadsheetModel } from "../model";

/** A cell: a value, "=…" for a formula with no cached value, or the fields
 * of a formula cell spelled out. */
export type CellSpec =
  | CellValue
  | {
      formula?: string;
      value?: CellValue;
      /** "B1:C3" — a legacy array anchor's range. */
      array?: string;
      dynamicArray?: true;
      formulaUnparsed?: true;
      dataTableXml?: string;
    };

export interface CalcWorkbookOptions {
  format?: "xlsx" | "ods";
  dateSystem?: "1900" | "1904";
  calc?: CalcSettings;
  /** Sheets that are foreign (chartsheets…), by name. */
  foreign?: string[];
}

function cellOf(col: number, spec: CellSpec): Cell {
  if (typeof spec === "string" && spec.startsWith("=")) {
    return { col, value: null, formula: spec.slice(1) };
  }
  if (spec === null || typeof spec !== "object" || "error" in spec) {
    return { col, value: spec as CellValue };
  }
  const cell: Cell = { col, value: spec.value ?? null };
  if (spec.formula !== undefined) {
    cell.formula = spec.formula.startsWith("=") ? spec.formula.slice(1) : spec.formula;
  }
  if (spec.array !== undefined) {
    cell.arrayRange = parseRange(spec.array)!;
  }
  if (spec.dynamicArray) {
    cell.dynamicArray = true;
  }
  if (spec.formulaUnparsed) {
    cell.formulaUnparsed = true;
  }
  if (spec.dataTableXml !== undefined) {
    cell.dataTableXml = spec.dataTableXml;
  }
  return cell;
}

/** A sheet from A1 address → cell spec. */
export function calcSheet(name: string, cells: Record<string, CellSpec>, kind: Sheet["kind"] = "grid"): Sheet {
  const rows = new Map<number, Cell[]>();
  for (const [address, spec] of Object.entries(cells)) {
    const at = parseA1(address);
    if (!at) {
      throw new Error(`Not an A1 address: ${address}`);
    }
    const row = rows.get(at.row) ?? [];
    row.push(cellOf(at.col, spec));
    rows.set(at.row, row);
  }
  const sorted: Row[] = [...rows.entries()]
    .toSorted(([a], [b]) => a - b)
    .map(([index, row]) => ({ index, cells: row.toSorted((a, b) => a.col - b.col) }));
  return {
    id: `sheet-${name}`,
    name,
    visibility: "visible",
    kind,
    rows: sorted,
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    conditionalFormats: [],
    validations: [],
  };
}

/** A workbook from sheet name → (A1 address → cell spec), in order. Sheet
 * ids are `sheet-<name>`. */
export function calcWorkbook(sheets: Record<string, Record<string, CellSpec>>, options: CalcWorkbookOptions = {}): SpreadsheetModel {
  const model: SpreadsheetModel = {
    kind: "spreadsheet",
    id: "calc-workbook",
    title: "Calc",
    format: options.format ?? "xlsx",
    readerVersion: 0,
    sheets: Object.entries(sheets).map(([name, cells]) => calcSheet(name, cells, options.foreign?.includes(name) ? "foreign" : "grid")),
    activeSheetIndex: 0,
    dateSystem: options.dateSystem ?? "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
    dxfs: [],
  };
  if (options.calc) {
    model.calc = options.calc;
  }
  return model;
}
