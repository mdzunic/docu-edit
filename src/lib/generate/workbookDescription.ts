// The JSON workbook-description format (SPEC-140): a plain-JSON object
// mapping 1:1 onto `WorkbookBuilder` — each key is a builder call,
// interpreted through the same workbookPlan.ts the fluent methods drive, so
// a description and the program that would build it produce the same model
// and fail identically, with the same path. The published JSON Schema
// (schema/workbook-description.json, beside SPEC-045's
// schema/document-description.json) describes this shape for external
// tooling; this file is its runtime validator. Every object is validated
// closed: a key this spec does not build ("names", "charts",
// "validations"…) is an unknown property, never silently ignored.
//
// | Key                                       | Builder call            |
// | `title`, `format`, `dateSystem`, `locale` | constructor options     |
// | `styles` (top level)                      | defineStyle per entry   |
// | `sheets[i]` + name/tabColor/hidden/…      | addSheet                |
// | `cells` (A1 → CellInput)                  | setCell, in key order   |
// | `data[]` (`at`, `rows`, `parseText?`)     | setRows                 |
// | `styles[]` (`range` + style/numberFormat) | style / numberFormat    |
// | `columns[]`, `rows[]`                     | columnWidth, rowHeight  |
// | `merges`, `freeze`                        | merge, freeze           |
//
// Within a sheet the keys are applied in the table's order — cells, data,
// styles, columns, rows, merges, freeze — whatever order the object lists
// them in. `{}` describes a one-sheet empty workbook.
import type { SpreadsheetModel } from "../spreadsheet/model";
import type { SpreadsheetLocaleTag } from "../spreadsheet/numberFormat/locale";
import { expectArray, expectPlainObject, rejectUnknownKeys } from "./validate";
import { OPTION_KEYS, SHEET_OPTION_KEYS, WorkbookPlan, type CellInput, type WorkbookOutputFormat } from "./workbookPlan";
import type { CellStyleSpec } from "./workbookStyles";

export const DESCRIPTION_KEYS = [...OPTION_KEYS, "styles", "sheets"] as const;
export const SHEET_KEYS = ["name", ...SHEET_OPTION_KEYS, "cells", "data", "styles", "columns", "rows", "merges", "freeze"] as const;
export const DATA_KEYS = ["at", "rows", "parseText"] as const;
export const STYLE_ENTRY_KEYS = ["range", "style", "numberFormat"] as const;
export const COLUMN_ENTRY_KEYS = ["columns", "width"] as const;
export const ROW_ENTRY_KEYS = ["rows", "height"] as const;

/** One sheet of a description (`addSheet` and every per-sheet call). */
export interface WorkbookSheetDescription {
  name: string;
  tabColor?: string;
  hidden?: boolean;
  showGridLines?: boolean;
  /** A1 → the cell's input, applied in key order. */
  cells?: Record<string, CellInput>;
  data?: { at: string; rows: CellInput[][]; parseText?: boolean }[];
  /** At least one of `style` and `numberFormat`; `style` applies first. */
  styles?: { range: string; style?: CellStyleSpec | string; numberFormat?: string }[];
  columns?: { columns: string; width: { chars: number } | { px: number } }[];
  rows?: { rows: string | number; height: { pt: number } | { px: number } }[];
  merges?: string[];
  freeze?: { rows?: number; cols?: number };
}

/** A JSON workbook description: every field optional — `{}` describes a
 * workbook with one empty "Sheet1". */
export interface WorkbookDescription {
  title?: string;
  format?: WorkbookOutputFormat;
  dateSystem?: "1900" | "1904";
  locale?: SpreadsheetLocaleTag;
  /** Named styles, defined before any sheet is built. */
  styles?: Record<string, CellStyleSpec>;
  sheets?: WorkbookSheetDescription[];
}

/** The entries of an optional array key, each validated as a closed object. */
function entries(sheet: Record<string, unknown>, key: string, keys: readonly string[], path: string): Record<string, unknown>[] {
  if (sheet[key] === undefined) {
    return [];
  }
  return expectArray(sheet[key], `${path}.${key}`).map((raw, k) => {
    const entryPath = `${path}.${key}[${k}]`;
    const entry = expectPlainObject(raw, entryPath);
    rejectUnknownKeys(entry, keys, entryPath);
    return entry;
  });
}

/** Validates and interprets a description into the builder's state — the
 * shared body of workbookFromDescription and WorkbookBuilder.fromDescription.
 * Throws DocumentGenerationError on the first invalid field, named by its
 * path from the description's root ("sheets[1].cells.C4.fomula: unknown
 * property"). */
export function planFromDescription(description: unknown): WorkbookPlan {
  const obj = expectPlainObject(description, "description");
  rejectUnknownKeys(obj, DESCRIPTION_KEYS, "");
  const options = Object.fromEntries(OPTION_KEYS.filter((key) => obj[key] !== undefined).map((key) => [key, obj[key]]));
  const plan = new WorkbookPlan(options);
  if (obj.styles !== undefined) {
    for (const [name, spec] of Object.entries(expectPlainObject(obj.styles, "styles"))) {
      plan.defineStyle(name, spec);
    }
  }
  if (obj.sheets === undefined) {
    return plan;
  }
  expectArray(obj.sheets, "sheets").forEach((raw, i) => {
    const path = `sheets[${i}]`;
    const sheet = expectPlainObject(raw, path);
    rejectUnknownKeys(sheet, SHEET_KEYS, path);
    plan.addSheet(sheet.name, Object.fromEntries(SHEET_OPTION_KEYS.filter((key) => sheet[key] !== undefined).map((key) => [key, sheet[key]])));
    if (sheet.cells !== undefined) {
      for (const [ref, input] of Object.entries(expectPlainObject(sheet.cells, `${path}.cells`))) {
        plan.setCell(ref, input);
      }
    }
    for (const entry of entries(sheet, "data", DATA_KEYS, path)) {
      plan.setRows(entry.at, entry.rows, entry.parseText === undefined ? undefined : { parseText: entry.parseText });
    }
    for (const entry of entries(sheet, "styles", STYLE_ENTRY_KEYS, path)) {
      plan.style(entry.range, entry.style, entry.numberFormat);
    }
    for (const entry of entries(sheet, "columns", COLUMN_ENTRY_KEYS, path)) {
      plan.columnWidth(entry.columns, entry.width);
    }
    for (const entry of entries(sheet, "rows", ROW_ENTRY_KEYS, path)) {
      plan.rowHeight(entry.rows, entry.height);
    }
    if (sheet.merges !== undefined) {
      for (const range of expectArray(sheet.merges, `${path}.merges`)) {
        plan.merge(range);
      }
    }
    if (sheet.freeze !== undefined) {
      plan.freeze(sheet.freeze);
    }
  });
  return plan;
}

/** Validates and interprets a JSON workbook description into the same
 * SpreadsheetModel the fluent builder produces for the same content, every
 * formula's cached value computed — never writing a byte. */
export function workbookFromDescription(description: unknown): SpreadsheetModel {
  const plan = planFromDescription(description);
  return plan.materialise(plan.format);
}
