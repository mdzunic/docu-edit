// Mail-merge engine (SPEC-046): CSV/JSON data sources, template validation,
// per-record resolution and merge output (zip / single document). Pure,
// framework-free like the rest of lib/ — the editor composable drives it and
// owns idle-yielding (scheduleIdleWork) between records.
import {
  isTable,
  type Block,
  type DocumentModel,
  type HeaderFooterKind,
  type NoteContent,
  type Paragraph,
  type Run,
  type RunFormatting,
  type Table,
  type TableRow,
  isMergeChipField,
  type MergeFieldSpec,
} from "./documentModel";
import { zip, type ZipEntryInput } from "./core/zipWriter";
import { docxFileName, writeDocx } from "./docx/writer";
import { odtFileName, writeOdt } from "./odf/writer";

// --- data sources ------------------------------------------------------

export type MergeRecord = Record<string, unknown>;

export interface ParsedDataSource {
  kind: "csv" | "json";
  /** Top-level column/property names — CSV headers as-is, or the union of
   * every record's own top-level keys for JSON. */
  columns: string[];
  /** The subset of `columns` whose value is an array in at least one
   * record — what "Table -> Repeat row for each..." offers (JSON only). */
  arrayColumns: string[];
  records: MergeRecord[];
}

/** Typed `null` seed for a `.vue` script's attached-source ref, which espree
 * bars from writing the `ParsedDataSource | null` annotation directly (same
 * convention as commands.ts's noSelection). */
export function noParsedDataSource(): ParsedDataSource | null {
  return null;
}

/** Typed pass-through for a `ParsedDataSource | null` passed into a `.vue`
 * component as a runtime `Object` prop (same espree constraint as
 * documentModel.ts's asDocumentModel) — null stays null. */
export function asParsedDataSourceOrNull(value: unknown): ParsedDataSource | null {
  return (value as ParsedDataSource | null) ?? null;
}

export class MergeDataSourceError extends Error {}

/** Reads a picked File's text content — the same FileReader-based approach
 * openDocument.ts's readFileAsArrayBuffer uses for bytes, since jsdom's
 * File/Blob polyfill (this project's unit-test environment) has no
 * `.text()`, only the FileReader API every browser and jsdom both
 * implement. Used by the merge dialog to read an attached CSV/JSON file. */
export function readFileAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(reader.result as string));
    reader.addEventListener("error", () => reject(reader.error ?? new Error("Failed to read file")));
    reader.readAsText(file);
  });
}

/** A full CSV text into rows of raw string cells — a hand-rolled parser (no
 * dependency added, per CLAUDE.md's protected package.json) handling quoted
 * fields (including embedded commas/newlines/escaped quotes) and preserving
 * a row whose only cell is empty rather than dropping it. */
export function parseCsvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = "";
  let inQuotes = false;
  let sawContentThisRow = false;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i += 1;
        continue;
      }
      cur += ch;
      i += 1;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      sawContentThisRow = true;
      i += 1;
      continue;
    }
    if (ch === ",") {
      row.push(cur);
      cur = "";
      sawContentThisRow = true;
      i += 1;
      continue;
    }
    if (ch === "\r") {
      i += 1;
      continue;
    }
    if (ch === "\n") {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = "";
      sawContentThisRow = false;
      i += 1;
      continue;
    }
    cur += ch;
    sawContentThisRow = true;
    i += 1;
  }
  if (cur.length > 0 || row.length > 0 || sawContentThisRow) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

/** CSV -> ParsedDataSource (SPEC-046 "Data sources and field resolution"):
 * the first row is the header, one record per row after it, every value a
 * string. A CSV source has no arrays. */
export function parseCsvSource(text: string): ParsedDataSource {
  const rows = parseCsvRows(text);
  if (rows.length === 0) {
    return { kind: "csv", columns: [], arrayColumns: [], records: [] };
  }
  const header = rows[0];
  const records: MergeRecord[] = rows.slice(1).map((row) => {
    const record: MergeRecord = {};
    header.forEach((col, index) => {
      record[col] = row[index] ?? "";
    });
    return record;
  });
  return { kind: "csv", columns: [...header], arrayColumns: [], records };
}

/** JSON -> ParsedDataSource: a top-level array of objects, one per record. */
export function parseJsonSource(text: string): ParsedDataSource {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch (error) {
    throw new MergeDataSourceError(`Invalid JSON: ${(error as Error).message}`);
  }
  if (!Array.isArray(data)) {
    throw new MergeDataSourceError("A JSON data source must be a top-level array of objects.");
  }
  const records: MergeRecord[] = data.map((item) =>
    item !== null && typeof item === "object" && !Array.isArray(item) ? (item as MergeRecord) : {},
  );
  const columns = new Set<string>();
  const arrayColumns = new Set<string>();
  for (const record of records) {
    for (const key of Object.keys(record)) {
      columns.add(key);
      if (Array.isArray(record[key])) {
        arrayColumns.add(key);
      }
    }
  }
  return { kind: "json", columns: [...columns], arrayColumns: [...arrayColumns], records };
}

// --- field resolution ---------------------------------------------------

/** Resolution order for a name (SPEC-046 "Data sources and field
 * resolution"): an exact column/property of that name first, then dotted
 * traversal. Exact-first matters for CSV, where a column literally named
 * "customer.name" is a name, not a path. */
function resolveExactOrPath(record: MergeRecord, name: string): { found: boolean; value: unknown } {
  if (Object.prototype.hasOwnProperty.call(record, name)) {
    return { found: true, value: record[name] };
  }
  if (name.includes(".")) {
    let current: unknown = record;
    for (const part of name.split(".")) {
      if (current !== null && typeof current === "object" && !Array.isArray(current) && Object.prototype.hasOwnProperty.call(current, part)) {
        current = (current as MergeRecord)[part];
      } else {
        return { found: false, value: undefined };
      }
    }
    return { found: true, value: current };
  }
  return { found: false, value: undefined };
}

interface Scope {
  record: MergeRecord;
  /** The current array element, when resolving inside an expanded
   * repeating row — a name resolves against this first, falling back to
   * `record` (SPEC-046 "Repeating table rows"). */
  element?: unknown;
  sourceKind: "csv" | "json";
}

function resolveInScope(scope: Scope, name: string): { found: boolean; value: unknown } {
  if (scope.element !== undefined && scope.element !== null && typeof scope.element === "object" && !Array.isArray(scope.element)) {
    const fromElement = resolveExactOrPath(scope.element as MergeRecord, name);
    if (fromElement.found) {
      return fromElement;
    }
  }
  return resolveExactOrPath(scope.record, name);
}

/** JSON truthiness (SPEC-046 "Conditional blocks"): absent/null/false/0/""
 * /[]/{} are false — [] and {} deliberately, so `{{#if line_items}}`
 * guarding a repeating row does not fire on an empty array. */
function isTruthyJsonValue(value: unknown): boolean {
  if (value === undefined || value === null || value === false || value === 0 || value === "") {
    return false;
  }
  if (Array.isArray(value)) {
    return value.length > 0;
  }
  if (typeof value === "object") {
    return Object.keys(value as object).length > 0;
  }
  return true;
}

/** CSV truthiness: every value is a string, so "false"/"no"/"0" (trimmed,
 * case-insensitive) are false in addition to absence/emptiness. */
function isTruthyCsvValue(value: unknown): boolean {
  if (typeof value !== "string") {
    return isTruthyJsonValue(value);
  }
  const trimmed = value.trim().toLowerCase();
  return trimmed !== "" && trimmed !== "false" && trimmed !== "no" && trimmed !== "0";
}

function evaluateCondition(scope: Scope, name: string): boolean {
  const { found, value } = resolveInScope(scope, name);
  if (!found) {
    // "A field a condition tests and the record does not have is false, and
    // not a merge-report entry" — absence is an answer, not a fault.
    return false;
  }
  return scope.sourceKind === "csv" ? isTruthyCsvValue(value) : isTruthyJsonValue(value);
}

// --- named field list (SPEC-046 "The named field list") -----------------

export interface NamedField {
  name: string;
  array: boolean;
}

/** Typed pass-through for a `NamedField[]` passed into a `.vue` component as
 * a runtime `Array` prop (same espree constraint as asDocumentModel). */
export function asNamedFieldList(value: unknown): NamedField[] {
  return value as NamedField[];
}

/** Every merge/if/each chip name already used anywhere in the document
 * (body, headers, footers, notes) — part of the named field list even
 * without an attached source, so a template can be written before its data
 * exists (AC-15). */
export function collectUsedFieldNames(model: DocumentModel): { names: string[]; arrayNames: string[] } {
  const names: string[] = [];
  const arrayNames: string[] = [];
  const seen = new Set<string>();
  const seenArray = new Set<string>();
  const visitRun = (run: Run): void => {
    const field = run.field;
    if (!field || !isMergeChipField(field)) {
      return;
    }
    if (field.kind === "merge" || field.kind === "if") {
      if (!seen.has(field.name)) {
        seen.add(field.name);
        names.push(field.name);
      }
    } else if (field.kind === "each") {
      if (!seen.has(field.name)) {
        seen.add(field.name);
        names.push(field.name);
      }
      if (!seenArray.has(field.name)) {
        seenArray.add(field.name);
        arrayNames.push(field.name);
      }
    }
  };
  const visitParagraph = (paragraph: Paragraph): void => paragraph.runs.forEach(visitRun);
  const visitBlocks = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        visitParagraph(block);
      } else if (isTable(block)) {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            cell.blocks.forEach(visitParagraph);
          }
        }
      }
    }
  };
  model.sections.forEach((section) => {
    visitBlocks(section.blocks);
    (["default", "first", "even"] as HeaderFooterKind[]).forEach((kind) => {
      visitBlocks(section.headers?.[kind]?.blocks ?? []);
      visitBlocks(section.footers?.[kind]?.blocks ?? []);
    });
  });
  (model.notes ?? []).forEach((note) => visitBlocks(note.blocks));
  return { names, arrayNames };
}

/** The insertion list "Insert merge field" offers (SPEC-046 "The named
 * field list"): source columns/properties first, then names already used
 * by a chip in this document, deduplicated — in that order, so a template
 * can be written before its data exists. */
export function namedFieldList(source: ParsedDataSource | null, model: DocumentModel): NamedField[] {
  const used = collectUsedFieldNames(model);
  const arraySet = new Set([...(source?.arrayColumns ?? []), ...used.arrayNames]);
  const seen = new Set<string>();
  const list: NamedField[] = [];
  const add = (name: string, array: boolean): void => {
    if (seen.has(name)) {
      return;
    }
    seen.add(name);
    list.push({ name, array });
  };
  (source?.columns ?? []).forEach((name) => add(name, source?.arrayColumns.includes(name) ?? false));
  used.names.forEach((name) => add(name, arraySet.has(name)));
  return list;
}

/** The array-only subset "Table -> Repeat row for each..." offers. */
export function namedArrayFieldList(source: ParsedDataSource | null, model: DocumentModel): NamedField[] {
  return namedFieldList(source, model).filter((field) => field.array);
}

// --- field mapping report (AC-24) ---------------------------------------

export interface FieldMappingEntry {
  name: string;
  mapped: boolean;
}

export interface FieldMappingReport {
  /** One entry per source column/property — SPEC-046's "column-to-field
   * mapping including unmapped fields". */
  columns: FieldMappingEntry[];
  /** Chip names used in the document that match no source column exactly. */
  unmappedChips: string[];
}

export function buildFieldMappingReport(source: ParsedDataSource | null, model: DocumentModel): FieldMappingReport {
  const used = collectUsedFieldNames(model);
  const usedSet = new Set(used.names);
  const columns = (source?.columns ?? []).map((name) => ({ name, mapped: usedSet.has(name) }));
  const columnSet = new Set(source?.columns ?? []);
  const unmappedChips = used.names.filter((name) => !columnSet.has(name));
  return { columns, unmappedChips };
}

// --- template validation (AC-35..38) ------------------------------------

export type TemplateErrorKind = "unpaired-if" | "unpaired-endif" | "crossed" | "nested-each" | "outside-body";

export interface TemplateError {
  kind: TemplateErrorKind;
  name?: string;
  location: string;
  message: string;
}

export class MergeTemplateError extends Error {
  constructor(public readonly errors: TemplateError[]) {
    super(`Template has ${errors.length} error(s) that block the merge run: ${errors.map((e) => e.message).join("; ")}`);
  }
}

function chipFieldOf(run: Run): MergeFieldSpec | null {
  return run.field && isMergeChipField(run.field) ? run.field : null;
}

/** Validates one flat paragraph-list scope (the body, or one table cell —
 * a chip's pair can never cross that boundary, so each scope is validated
 * independently with its own stack). Any `if` left open at the end of this
 * scope, or an `endif` with nothing open, is a "crossed" error when this is
 * a nested (cell) scope — its partner is necessarily in a different scope —
 * and an ordinary unpaired error at the top level. */
function validateScope(blocks: Block[], location: string, nested: boolean, errors: TemplateError[]): void {
  const stack: Array<{ name: string; blockIndex: number }> = [];
  blocks.forEach((block, blockIndex) => {
    if (block.type !== "paragraph") {
      return;
    }
    block.runs.forEach((run) => {
      const field = chipFieldOf(run);
      if (!field) {
        return;
      }
      if (field.kind === "if") {
        stack.push({ name: field.name, blockIndex });
      } else if (field.kind === "endif") {
        if (stack.length === 0) {
          errors.push({
            kind: nested ? "crossed" : "unpaired-endif",
            location: `${location}, paragraph ${blockIndex + 1}`,
            message: `{{/if}} with no open {{#if}} in ${location}, paragraph ${blockIndex + 1}.`,
          });
        } else {
          stack.pop();
        }
      }
    });
  });
  for (const open of stack) {
    errors.push({
      kind: nested ? "crossed" : "unpaired-if",
      name: open.name,
      location: `${location}, paragraph ${open.blockIndex + 1}`,
      message: `{{#if ${open.name}}} in ${location}, paragraph ${open.blockIndex + 1} is never closed in the same scope.`,
    });
  }
}

function validateEachChipsInTable(table: Table, location: string, errors: TemplateError[]): void {
  table.rows.forEach((row, rowIndex) => {
    let count = 0;
    row.cells.forEach((cell) => {
      cell.blocks.forEach((paragraph) => {
        paragraph.runs.forEach((run) => {
          const field = chipFieldOf(run);
          if (field?.kind === "each") {
            count += 1;
          }
        });
      });
    });
    if (count > 1) {
      errors.push({
        kind: "nested-each",
        location: `${location}, row ${rowIndex + 1}`,
        message: `Row ${rowIndex + 1} of ${location} carries ${count} {{#each}} chips — a row can only repeat over one array.`,
      });
    }
  });
}

function scanBodyBlocks(blocks: Block[], location: string, errors: TemplateError[]): void {
  validateScope(blocks, location, false, errors);
  blocks.forEach((block, blockIndex) => {
    if (!isTable(block)) {
      return;
    }
    const tableLocation = `${location}, table ${blockIndex + 1}`;
    validateEachChipsInTable(block, tableLocation, errors);
    block.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, colIndex) => {
        validateScope(cell.blocks, `${tableLocation}, row ${rowIndex + 1}, column ${colIndex + 1}`, true, errors);
      });
    });
  });
}

function scanForChipsOutsideBody(blocks: Block[], location: string, errors: TemplateError[]): void {
  const visit = (paragraph: Paragraph, where: string): void => {
    paragraph.runs.forEach((run) => {
      const field = chipFieldOf(run);
      if (field && field.kind !== "merge") {
        errors.push({
          kind: "outside-body",
          name: field.kind === "endif" ? undefined : field.name,
          location: where,
          message: `{{#if}}/{{/if}}/{{#each}} chips are only allowed in the document body — found one in ${where}.`,
        });
      }
    });
  };
  blocks.forEach((block) => {
    if (block.type === "paragraph") {
      visit(block, location);
    } else if (isTable(block)) {
      block.rows.forEach((row) =>
        row.cells.forEach((cell) => cell.blocks.forEach((paragraph) => visit(paragraph, location))),
      );
    }
  });
}

/** Template errors (SPEC-046): unpaired, crossed or illegally nested chips,
 * and any conditional/repeat chip outside the document body — listed with
 * their locations, blocking the run before any document is produced. Data
 * errors (missing fields, a non-array `{{#each}}` target) are never
 * reported here; they never fail a run (AC-39). */
export function validateTemplate(model: DocumentModel): TemplateError[] {
  const errors: TemplateError[] = [];
  model.sections.forEach((section, sectionIndex) => {
    scanBodyBlocks(section.blocks, `body (section ${sectionIndex + 1})`, errors);
    (["default", "first", "even"] as HeaderFooterKind[]).forEach((kind) => {
      const header = section.headers?.[kind];
      if (header) {
        scanForChipsOutsideBody(header.blocks, `header (section ${sectionIndex + 1}, ${kind})`, errors);
      }
      const footer = section.footers?.[kind];
      if (footer) {
        scanForChipsOutsideBody(footer.blocks, `footer (section ${sectionIndex + 1}, ${kind})`, errors);
      }
    });
  });
  (model.notes ?? []).forEach((note: NoteContent) => {
    scanForChipsOutsideBody(note.blocks, `${note.kind} ${note.noteId}`, errors);
  });
  return errors;
}

// --- resolution -----------------------------------------------------------

export interface MergeReportAccumulator {
  missingFields: Array<{ recordIndex: number; field: string }>;
  arrayFieldErrors: Array<{ recordIndex: number; field: string }>;
}

function newReportAccumulator(): MergeReportAccumulator {
  return { missingFields: [], arrayFieldErrors: [] };
}

/** A missing field's clearly-marked gap (AC-44): the field name in
 * brackets, highlighted, so it reads as an obvious placeholder rather than
 * silently-empty text. */
function gapRun(name: string, formatting: RunFormatting): Run {
  return { text: `[${name}]`, formatting: { ...formatting, highlight: "yellow" } };
}

function isChipAloneInParagraph(paragraph: Paragraph, runIndex: number): boolean {
  return paragraph.runs.every((run, index) => {
    if (index === runIndex) {
      return true;
    }
    if (run.field && isMergeChipField(run.field)) {
      return false;
    }
    return run.text.trim() === "";
  });
}

function findChipRunIndex(paragraph: Paragraph, kinds: MergeFieldSpec["kind"][]): number {
  return paragraph.runs.findIndex((run) => {
    const field = chipFieldOf(run);
    return field !== null && kinds.includes(field.kind);
  });
}

/** Finds the `{{/if}}` matching the `{{#if}}` at (ifBlock, ifRun) within
 * blocks[searchFrom..end), honouring nesting (an inner if/endif pair does
 * not close the outer one). The template was already validated, so a match
 * always exists within this scope. */
function findMatchingEndif(
  blocks: Block[],
  ifBlock: number,
  ifRun: number,
  end: number,
): { block: number; run: number } {
  let depth = 0;
  const firstParagraph = blocks[ifBlock] as Paragraph;
  for (let r = ifRun + 1; r < firstParagraph.runs.length; r++) {
    const field = chipFieldOf(firstParagraph.runs[r]);
    if (field?.kind === "if") {
      depth += 1;
    } else if (field?.kind === "endif") {
      if (depth === 0) {
        return { block: ifBlock, run: r };
      }
      depth -= 1;
    }
  }
  for (let b = ifBlock + 1; b < end; b++) {
    const block = blocks[b];
    if (block.type !== "paragraph") {
      continue;
    }
    for (let r = 0; r < block.runs.length; r++) {
      const field = chipFieldOf(block.runs[r]);
      if (field?.kind === "if") {
        depth += 1;
      } else if (field?.kind === "endif") {
        if (depth === 0) {
          return { block: b, run: r };
        }
        depth -= 1;
      }
    }
  }
  throw new Error("Unreachable: findMatchingEndif called on an unvalidated template");
}

/** Removes exactly one run from a paragraph — used for an inline-form
 * true-condition chip, which removes only the chip itself. */
function withRunRemoved(paragraph: Paragraph, runIndex: number): Paragraph {
  return { ...paragraph, runs: paragraph.runs.filter((_, index) => index !== runIndex) };
}

/** Recursively resolves conditional (`if`/`endif`) regions in blocks[start,
 * end) per "Conditional blocks": block form (the chip alone in its
 * paragraph) always drops the two delimiter paragraphs and additionally
 * drops everything between them when false; inline form drops only the two
 * chip runs when true, or the whole range (paragraph-joined) when false —
 * mirroring the editor's own range deletion. Non-paragraph blocks (tables,
 * images) pass through a surviving region untouched at this stage; they are
 * resolved (row expansion, chip substitution) by the caller afterwards.
 *
 * Resolves exactly the leftmost if/endif pair in the range, splices the
 * result back in, and recurses over the (shorter) spliced range rather than
 * finalizing the touched paragraphs immediately. A merged paragraph — the
 * opening or closing fragment of an inline conditional, or the joined
 * before/after of a false one — can itself carry further if-chips (a
 * sibling conditional sharing the paragraph, or one nested inside the
 * fragment); recursing over the splice point re-scans it instead of letting
 * those chips survive unresolved into the output. */
function resolveConditionalRegion(blocks: Block[], start: number, end: number, scope: Scope): Block[] {
  for (let i = start; i < end; i++) {
    const block = blocks[i];
    if (block.type !== "paragraph") {
      continue;
    }
    const ifRunIndex = findChipRunIndex(block, ["if"]);
    if (ifRunIndex === -1) {
      continue;
    }
    const field = chipFieldOf(block.runs[ifRunIndex]) as Extract<MergeFieldSpec, { kind: "if" }>;
    const endif = findMatchingEndif(blocks, i, ifRunIndex, end);
    const endifParagraph = blocks[endif.block] as Paragraph;
    const blockForm = isChipAloneInParagraph(block, ifRunIndex) && isChipAloneInParagraph(endifParagraph, endif.run);
    const truthy = evaluateCondition(scope, field.name);

    let replacement: Block[];
    if (blockForm) {
      replacement = truthy ? blocks.slice(i + 1, endif.block) : [];
    } else if (truthy) {
      replacement =
        endif.block === i
          ? [withRunRemoved(withRunRemoved(block, ifRunIndex), endif.run - 1)]
          : [withRunRemoved(block, ifRunIndex), ...blocks.slice(i + 1, endif.block), withRunRemoved(endifParagraph, endif.run)];
    } else if (endif.block === i) {
      const before = block.runs.slice(0, ifRunIndex);
      const after = block.runs.slice(endif.run + 1);
      replacement = [{ ...block, runs: [...before, ...after] }];
    } else {
      const before = block.runs.slice(0, ifRunIndex);
      const after = endifParagraph.runs.slice(endif.run + 1);
      replacement = [{ ...block, runs: [...before, ...after] }];
    }

    const removedCount = endif.block - i + 1;
    const spliced = [...blocks.slice(0, i), ...replacement, ...blocks.slice(endif.block + 1)];
    const newEnd = end - removedCount + replacement.length;
    return [...blocks.slice(start, i), ...resolveConditionalRegion(spliced, i, newEnd, scope)];
  }
  return blocks.slice(start, end);
}

function findEachChip(row: TableRow): { cellIndex: number; blockIndex: number; runIndex: number; name: string } | null {
  for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
    const cell = row.cells[cellIndex];
    for (let blockIndex = 0; blockIndex < cell.blocks.length; blockIndex++) {
      const paragraph = cell.blocks[blockIndex];
      for (let runIndex = 0; runIndex < paragraph.runs.length; runIndex++) {
        const field = chipFieldOf(paragraph.runs[runIndex]);
        if (field?.kind === "each") {
          return { cellIndex, blockIndex, runIndex, name: field.name };
        }
      }
    }
  }
  return null;
}

function resolveMergeChipsInParagraph(
  paragraph: Paragraph,
  scope: Scope,
  recordIndex: number,
  report: MergeReportAccumulator,
): Paragraph {
  let changed = false;
  const runs = paragraph.runs.map((run) => {
    const field = chipFieldOf(run);
    if (!field || field.kind !== "merge") {
      return run;
    }
    changed = true;
    const { found, value } = resolveInScope(scope, field.name);
    if (!found || value === null || value === undefined) {
      report.missingFields.push({ recordIndex, field: field.name });
      return gapRun(field.name, run.formatting);
    }
    const rest: Run = { text: String(value), formatting: run.formatting };
    if (run.styleId !== undefined) {
      rest.styleId = run.styleId;
    }
    return rest;
  });
  return changed ? { ...paragraph, runs } : paragraph;
}

function resolveCellBlocks(blocks: Paragraph[], scope: Scope, recordIndex: number, report: MergeReportAccumulator): Paragraph[] {
  const afterConditionals = resolveConditionalRegion(blocks, 0, blocks.length, scope) as Paragraph[];
  return afterConditionals.map((paragraph) => resolveMergeChipsInParagraph(paragraph, scope, recordIndex, report));
}

function resolveTable(table: Table, scope: Scope, recordIndex: number, report: MergeReportAccumulator): Table | null {
  const newRows: TableRow[] = [];
  table.rows.forEach((row) => {
    if (row.isHeader) {
      newRows.push(row);
      return;
    }
    const each = findEachChip(row);
    if (!each) {
      newRows.push({
        ...row,
        cells: row.cells.map((cell) => ({ ...cell, blocks: resolveCellBlocks(cell.blocks, scope, recordIndex, report) })),
      });
      return;
    }
    const stripEach = (r: TableRow): TableRow => ({
      ...r,
      cells: r.cells.map((cell, cellIndex) => {
        if (cellIndex !== each.cellIndex) {
          return cell;
        }
        return {
          ...cell,
          blocks: cell.blocks.map((paragraph, blockIndex) =>
            blockIndex === each.blockIndex ? withRunRemoved(paragraph, each.runIndex) : paragraph,
          ),
        };
      }),
    });
    const { found, value } = resolveInScope(scope, each.name);
    if (!found || !Array.isArray(value)) {
      report.arrayFieldErrors.push({ recordIndex, field: each.name });
      const stripped = stripEach(row);
      newRows.push({
        ...stripped,
        cells: stripped.cells.map((cell) => ({ ...cell, blocks: resolveCellBlocks(cell.blocks, { ...scope, element: {} }, recordIndex, report) })),
      });
      return;
    }
    for (const element of value) {
      const elementScope: Scope = { record: scope.record, element, sourceKind: scope.sourceKind };
      const stripped = stripEach(row);
      newRows.push({
        ...stripped,
        cells: stripped.cells.map((cell) => ({ ...cell, blocks: resolveCellBlocks(cell.blocks, elementScope, recordIndex, report) })),
      });
    }
  });
  const hasDataRow = newRows.some((row) => !row.isHeader);
  if (!hasDataRow) {
    return null;
  }
  return { ...table, rows: newRows };
}

function resolveSectionBlocks(blocks: Block[], scope: Scope, recordIndex: number, report: MergeReportAccumulator): Block[] {
  const afterConditionals = resolveConditionalRegion(blocks, 0, blocks.length, scope);
  const result: Block[] = [];
  afterConditionals.forEach((block) => {
    if (isTable(block)) {
      const resolved = resolveTable(block, scope, recordIndex, report);
      if (resolved) {
        result.push(resolved);
      }
      return;
    }
    if (block.type === "paragraph") {
      result.push(resolveMergeChipsInParagraph(block, scope, recordIndex, report));
      return;
    }
    result.push(block);
  });
  return result;
}

/** Resolves one record against the (already-validated) template, producing
 * a full document with no merge chips left in its body (SPEC-046 "Merge
 * chips in the model": "no merged output document contains a merge
 * chip"). Headers/footers are left as the template's own — see
 * buildSingleDocument's header comment for why single-document delivery in
 * particular cannot resolve them per-record. */
export function resolveRecordToModel(
  template: DocumentModel,
  record: MergeRecord,
  sourceKind: "csv" | "json",
  recordIndex: number,
  report: MergeReportAccumulator = newReportAccumulator(),
): DocumentModel {
  const model = structuredClone(template);
  const scope: Scope = { record, sourceKind };
  model.sections.forEach((section) => {
    section.blocks = resolveSectionBlocks(section.blocks, scope, recordIndex, report);
  });
  return model;
}

// --- output (AC-40..43) --------------------------------------------------

/** `<document title>-merge.zip`, the title run through the same trim/
 * default-title fallback docxFileName uses. */
export function mergeZipFileName(title: string): string {
  const base = docxFileName(title).replace(/\.docx$/, "");
  return `${base}-merge.zip`;
}

/** Keeps `[A-Za-z0-9._-]`, collapses every other run of characters to a
 * single "-", trims leading/trailing "-", truncates to 60 characters
 * (SPEC-046 "Merge output"). A "/" would create a directory inside the
 * archive; Windows rejects `<>:"|?*`. */
export function slugify(value: string): string {
  const collapsed = value.replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return collapsed.slice(0, 60);
}

/** `NNN-<slug>.<ext>` (a name field chosen) or `NNN.<ext>` (none) — `NNN`
 * the 1-based record number, zero-padded to the record count's width. The
 * index prefix makes entry names unique by construction. */
export function zipEntryName(recordIndex: number, recordCount: number, slug: string | null, ext: string): string {
  const width = String(recordCount).length;
  const n = String(recordIndex + 1).padStart(width, "0");
  return slug && slug.length > 0 ? `${n}-${slug}.${ext}` : `${n}.${ext}`;
}

function nameFieldSlug(record: MergeRecord, nameField: string | undefined): string | null {
  if (!nameField) {
    return null;
  }
  const { found, value } = resolveExactOrPath(record, nameField);
  if (!found || value === null || value === undefined) {
    return null;
  }
  return slugify(String(value));
}

export interface MergeRunReport {
  recordCount: number;
  columnMapping: FieldMappingEntry[];
  unmappedChips: string[];
  missingFields: Array<{ recordIndex: number; field: string }>;
  arrayFieldErrors: Array<{ recordIndex: number; field: string }>;
}

export interface MergeRunResult {
  delivery: "zip" | "single";
  fileName: string;
  bytes: Uint8Array;
  report: MergeRunReport;
}

/** Typed `null` seed for a `.vue` script's post-run report ref (same
 * convention as noParsedDataSource above). */
export function noMergeRunReport(): MergeRunReport | null {
  return null;
}

export interface MergeRunOptions {
  format: "docx" | "odt";
  delivery: "zip" | "single";
  /** A scalar column/property used to build the zip entry slug; omit for
   * plain `NNN.<ext>` names. Ignored for single-document delivery. */
  nameField?: string;
}

/** Typed pass-through for a merge output format coming from a `.vue`
 * template's plain string (same espree constraint as
 * documentModel.ts's asPageNumberFormat). */
export function asMergeFormat(value: string): "docx" | "odt" {
  return value as "docx" | "odt";
}

/** Typed pass-through for a merge delivery mode coming from a `.vue`
 * template's plain string. */
export function asMergeDelivery(value: string): "zip" | "single" {
  return value as "zip" | "single";
}

async function writeResolved(model: DocumentModel, format: "docx" | "odt"): Promise<Uint8Array> {
  const { bytes } = format === "odt" ? await writeOdt(model) : await writeDocx(model);
  return bytes;
}

async function buildZip(
  template: DocumentModel,
  source: ParsedDataSource,
  options: MergeRunOptions,
  report: MergeReportAccumulator,
): Promise<Uint8Array> {
  const ext = options.format;
  const entries: ZipEntryInput[] = [];
  for (let i = 0; i < source.records.length; i++) {
    const record = source.records[i];
    const resolved = resolveRecordToModel(template, record, source.kind, i, report);
    const bytes = await writeResolved(resolved, options.format);
    const slug = nameFieldSlug(record, options.nameField);
    entries.push({ name: zipEntryName(i, source.records.length, slug, ext), data: bytes });
  }
  return zip(entries);
}

/** Concatenates every record's resolved body in source order, a page break
 * between consecutive records, keeping the template's section setup,
 * headers and footers (SPEC-046 "Merge output" — single document). Headers/
 * footers stay the template's own, unresolved: they are shared across every
 * page of the concatenated file, page-number fields must number
 * continuously across it (not per record), and there is no single "current
 * record" a shared header could resolve against without a section break per
 * record, which is SPEC-028's feature and out of this spec's scope. */
function buildSingleDocument(template: DocumentModel, resolvedModels: DocumentModel[]): DocumentModel {
  const model = structuredClone(template);
  if (resolvedModels.length === 0) {
    model.sections[0].blocks = [];
    return model;
  }
  const firstSection = model.sections[0];
  const blocks: Block[] = [];
  resolvedModels.forEach((resolved, index) => {
    if (index > 0) {
      blocks.push({ type: "break", kind: "page" } satisfies Block);
    }
    blocks.push(...(resolved.sections[0]?.blocks ?? []));
  });
  firstSection.blocks = blocks;
  model.sections = [firstSection];
  return model;
}

export function singleDocumentFileName(title: string, format: "docx" | "odt"): string {
  return format === "odt" ? odtFileName(title) : docxFileName(title);
}

/** Runs a mail merge end to end: validates the template (throws
 * MergeTemplateError, blocking before any document is produced, if it
 * finds unpaired/crossed/nested chips or a conditional/repeat chip outside
 * the body), then resolves every record and packages the result. Data
 * errors (missing fields, a non-array `{{#each}}` target) never throw —
 * they land in the returned report instead (AC-39). For a very large
 * source, prefer driving `resolveRecordToModel` directly in an
 * idle-yielding loop (see useMerge.ts) rather than this all-at-once
 * convenience wrapper. */
export async function runMerge(
  template: DocumentModel,
  source: ParsedDataSource,
  options: MergeRunOptions,
): Promise<MergeRunResult> {
  const templateErrors = validateTemplate(template);
  if (templateErrors.length > 0) {
    throw new MergeTemplateError(templateErrors);
  }
  const mapping = buildFieldMappingReport(source, template);
  const accumulator = newReportAccumulator();
  if (options.delivery === "zip") {
    const bytes = await buildZip(template, source, options, accumulator);
    return {
      delivery: "zip",
      fileName: mergeZipFileName(template.title),
      bytes,
      report: {
        recordCount: source.records.length,
        columnMapping: mapping.columns,
        unmappedChips: mapping.unmappedChips,
        ...accumulator,
      },
    };
  }
  const resolvedModels = source.records.map((record, index) => resolveRecordToModel(template, record, source.kind, index, accumulator));
  const singleModel = buildSingleDocument(template, resolvedModels);
  const bytes = await writeResolved(singleModel, options.format);
  return {
    delivery: "single",
    fileName: singleDocumentFileName(template.title, options.format),
    bytes,
    report: {
      recordCount: source.records.length,
      columnMapping: mapping.columns,
      unmappedChips: mapping.unmappedChips,
      ...accumulator,
    },
  };
}

/** The merge preview (AC-23): the first record's resolved section, for a
 * read-only render in the merge dialog. Throws MergeTemplateError the same
 * way runMerge does if the template itself is invalid. */
export function previewFirstRecord(template: DocumentModel, source: ParsedDataSource): DocumentModel {
  const templateErrors = validateTemplate(template);
  if (templateErrors.length > 0) {
    throw new MergeTemplateError(templateErrors);
  }
  if (source.records.length === 0) {
    return structuredClone(template);
  }
  return resolveRecordToModel(template, source.records[0], source.kind, 0);
}
