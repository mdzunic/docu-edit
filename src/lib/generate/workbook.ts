// The fluent workbook builder (SPEC-140): the third product of the
// generation API, beside builder.ts's `DocumentBuilder` and presentation.ts's
// `PresentationBuilder`, whose output methods it mirrors. It composes the
// spreadsheet engine — SPEC-121's createBlankWorkbook and sheet-name rules,
// SPEC-123's applyStylePatch, SPEC-124's mergeCells and band commands,
// SPEC-126's CSV inference, SPEC-120's recalculation — and delegates every
// byte to SPEC-121's writers (`writeXlsx`/`writeOds`, from scratch): no XML,
// ZIP or format code of its own; a construct the writers cannot emit is
// fixed in the writer.
//
// Every method validates its arguments on entry (workbookPlan.ts) and
// returns `this`, so a DocumentGenerationError naming the offending path
// ("sheets[0].cells.B5.formula: unknown function SUMM") always precedes any
// byte. Cells accumulate in private state and the model is materialised
// once per output, at toWorkbook()/toBytes(), with every formula's cached
// value computed — the values Excel would show.
//
// Imports lib/spreadsheet, lib/core and generate's product-neutral modules
// only (workbookBoundary.spec.ts): never the text or deck builders.
import { spreadsheetFileName } from "../spreadsheet/download";
import { spreadsheetMimeType } from "../spreadsheet/fileTypes";
import type { SpreadsheetModel } from "../spreadsheet/model";
import { writeOds } from "../spreadsheet/ods/writer";
import { createBlankWorkbook } from "../spreadsheet/sheets";
import { writeXlsx } from "../spreadsheet/xlsx/writer";
import { DocumentGenerationError } from "./errors";
import { expectEnum, expectPlainObject, rejectUnknownKeys } from "./validate";
import { planFromDescription } from "./workbookDescription";
import {
  WORKBOOK_OUTPUT_FORMATS,
  WorkbookPlan,
  type AddSheetOptions,
  type CellInput,
  type ColumnWidthSpec,
  type FreezeSpec,
  type RowHeightSpec,
  type SetRowsOptions,
  type WorkbookBuilderOptions,
  type WorkbookOutputFormat,
} from "./workbookPlan";
import type { CellStyleSpec } from "./workbookStyles";

export type {
  AddSheetOptions,
  CellInput,
  CellInputObject,
  ColumnWidthSpec,
  FreezeSpec,
  RowHeightSpec,
  SetRowsOptions,
  WorkbookBuilderOptions,
  WorkbookOutputFormat,
} from "./workbookPlan";
export type { BorderSpec, CellStyleSpec } from "./workbookStyles";

export interface WorkbookGenerateOptions {
  /** The package's stamped modified time, passed straight to the writers'
   * own `modified`; defaults to now. Fix it for byte-identical output. */
  modified?: Date;
}

/** A fluent front end over a workbook: `addSheet` appends a sheet and makes
 * it current, `sheet` makes an existing one current, and every cell, range,
 * size, merge and freeze method acts on the current sheet (throwing "no
 * sheet — call addSheet first" when there is none). Every method returns
 * `this`; nothing is serialised until an output method is called, and the
 * builder stays reusable — each output builds a fresh model. */
export class WorkbookBuilder {
  #plan: WorkbookPlan;

  constructor(options: WorkbookBuilderOptions = {}) {
    this.#plan = new WorkbookPlan(options);
  }

  /** Interprets a JSON workbook description (workbookDescription.ts) into a
   * builder ready for more calls or for output — the description and this
   * class produce identical models, and identical errors, for identical
   * content. */
  static fromDescription(description: unknown): WorkbookBuilder {
    const builder = new WorkbookBuilder();
    builder.#plan = planFromDescription(description);
    return builder;
  }

  /** Registers a named style for `style`, a cell's `style` and later
   * presets to use by name. */
  defineStyle(name: string, spec: CellStyleSpec): this {
    this.#plan.defineStyle(name, spec);
    return this;
  }

  /** Appends a sheet and makes it current. The name follows SPEC-121's
   * rules (1–31 characters, none of : \ / ? * [ ], no leading or trailing
   * apostrophe, not "History", unique ignoring case). */
  addSheet(name: string, options?: AddSheetOptions): this {
    this.#plan.addSheet(name, options);
    return this;
  }

  /** Makes an existing sheet current (its name in any case). */
  sheet(name: string): this {
    this.#plan.selectSheet(name);
    return this;
  }

  /** Writes one cell ("B3"). */
  setCell(ref: string, input: CellInput): this {
    this.#plan.setCell(ref, input);
    return this;
  }

  /** Writes `values` left to right from `ref`. */
  setRow(ref: string, values: CellInput[]): this {
    this.#plan.setRows(ref, values, undefined, true);
    return this;
  }

  /** Writes a 2-D block with its top-left at `ref`. `parseText` routes every
   * string through SPEC-126's CSV inference in the builder's locale — never
   * into a formula. */
  setRows(ref: string, rows: CellInput[][], options?: SetRowsOptions): this {
    this.#plan.setRows(ref, rows, options);
    return this;
  }

  /** A formula in SPEC-119 canonical text (en-US syntax; a leading "=" is
   * accepted). Its cached value is computed at output. */
  setFormula(ref: string, formula: string): this {
    this.#plan.setCell(ref, { formula });
    return this;
  }

  /** Applies a style spec or a defined style's name to a range ("A1:E1",
   * "B:D", "3:5"). */
  style(range: string, style: CellStyleSpec | string): this {
    this.#plan.style(range, style, undefined);
    return this;
  }

  /** Applies a number format code (SPEC-118, en-US code syntax) to a range. */
  numberFormat(range: string, code: string): this {
    this.#plan.style(range, undefined, code);
    return this;
  }

  /** Sets the width of a column or columns ("B" or "B:D"): `{ chars }`
   * (0–255, SPEC-117's character width) or `{ px }`. */
  columnWidth(columns: string, width: ColumnWidthSpec): this {
    this.#plan.columnWidth(columns, width);
    return this;
  }

  /** Sets the height of a row or rows, 1-based (3, "3" or "3:5"): `{ pt }`
   * (0–409) or `{ px }`. */
  rowHeight(rows: number | string, height: RowHeightSpec): this {
    this.#plan.rowHeight(rows, height);
    return this;
  }

  /** Merges a range ("A1:D1"); refused at output when it would discard a
   * value outside its top-left cell, and on entry when it overlaps another
   * merge. */
  merge(range: string): this {
    this.#plan.merge(range);
    return this;
  }

  /** Freezes the first `rows` rows and `cols` columns; `{ rows: 0, cols: 0 }`
   * unfreezes. */
  freeze(spec: FreezeSpec): this {
    this.#plan.freeze(spec);
    return this;
  }

  /** The materialised workbook, in the builder's own format — a fresh,
   * plain-JSON model on every call. */
  toWorkbook(): SpreadsheetModel {
    return this.#plan.materialise(this.#plan.format);
  }

  #format(format: unknown): WorkbookOutputFormat {
    return format === undefined ? this.#plan.format : expectEnum(format, WORKBOOK_OUTPUT_FORMATS, "format");
  }

  /** The raw `.xlsx`/`.ods` package bytes, in `format` or the builder's own. */
  async toBytes(format?: WorkbookOutputFormat, options: WorkbookGenerateOptions = {}): Promise<Uint8Array> {
    const target = this.#format(format);
    const writeOptions = generateOptions(options);
    const model = this.#plan.materialise(target);
    return target === "ods" ? writeOds(model, writeOptions) : writeXlsx(model, writeOptions);
  }

  /** A `.xlsx`/`.ods` Blob (browser, or Node 18+). */
  async toBlob(format?: WorkbookOutputFormat, options: WorkbookGenerateOptions = {}): Promise<Blob> {
    const target = this.#format(format);
    const bytes = await this.toBytes(target, options);
    return new Blob([bytes as BlobPart], { type: spreadsheetMimeType(target) });
  }

  /** A `.xlsx`/`.ods` Buffer for Node callers. Throws where `Buffer` is not
   * global (a browser) — use `toBlob` or `toBytes` there. */
  async toBuffer(format?: WorkbookOutputFormat, options: WorkbookGenerateOptions = {}): Promise<Buffer> {
    if (typeof Buffer === "undefined") {
      throw new DocumentGenerationError("Buffer is not available in this environment — use toBlob or toBytes instead");
    }
    return Buffer.from(await this.toBytes(format, options));
  }

  /** The download file name for `format` (or the builder's own), by
   * spreadsheetDownload's rule: the title plus the extension, "Untitled
   * spreadsheet" when the title is blank. */
  fileName(format?: WorkbookOutputFormat): string {
    const named = createBlankWorkbook();
    named.format = this.#format(format);
    if (this.#plan.title !== undefined) {
      named.title = this.#plan.title;
    }
    return spreadsheetFileName(named);
  }
}

function generateOptions(options: unknown): { modified?: Date } {
  const obj = expectPlainObject(options, "options");
  rejectUnknownKeys(obj, ["modified"], "options");
  if (obj.modified === undefined) {
    return {};
  }
  if (!(obj.modified instanceof Date) || Number.isNaN(obj.modified.getTime())) {
    throw new DocumentGenerationError("options.modified: expected a valid Date");
  }
  return { modified: obj.modified };
}
