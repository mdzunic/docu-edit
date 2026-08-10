// Review-pane listing, navigation and accept/reject commands for tracked
// changes (SPEC-039). A command module outside commands.ts, like
// comments.ts/links.ts — see commands.ts's header comment for why that
// split exists.
import {
  isImage,
  isTable,
  type Block,
  type CellPropertyChange,
  type DocumentModel,
  type Image,
  type Paragraph,
  type ParagraphFormatting,
  type ParagraphTrackedChanges,
  type RowPropertyChange,
  type RunFormatting,
  type RunTrackedChanges,
  type Section,
  type Table,
  type TableCell,
  type TablePropertyChange,
  type TableRow,
  type TrackedChangeInfo,
} from "./documentModel";
import { blockText, compareCaret, normalizeRuns, type CaretPosition, type CellPath } from "./textEditing";

// --- review-pane listing ------------------------------------------------

export type ChangeKind = "insert" | "delete" | "format" | "paragraphMark" | "row" | "cell" | "tableProperty" | "image";

/** One tracked change as the review pane lists it (SPEC-039 AC "lists all
 * changes in document order"): every run/paragraph/row/image carrying the
 * same `changeId` collapses into a single entry, since they are all the
 * product of one edit event (see trackChangesDiff.ts's MarkPool). */
export interface ChangeEntry {
  changeId: string;
  kind: ChangeKind;
  author: string;
  date: string;
  /** e.g. "Inserted", "Deleted", "Formatted: Bold". */
  description: string;
  /** A short excerpt of the affected text, for the review-pane row. */
  preview: string;
  /** Where Next/Previous change navigation moves the caret to (AC-8). */
  at: CaretPosition;
}

const FORMAT_LABELS: Array<[keyof RunFormatting, string]> = [
  ["bold", "Bold"],
  ["italic", "Italic"],
  ["underline", "Underline"],
  ["strikethrough", "Strikethrough"],
  ["fontFamily", "Font"],
  ["fontSizePt", "Font size"],
  ["color", "Colour"],
  ["highlight", "Highlight"],
  ["vertAlign", "Vertical position"],
];

function describeRunFormatChange(previous: RunFormatting, current: RunFormatting): string {
  const labels: string[] = [];
  for (const [key, label] of FORMAT_LABELS) {
    if (previous[key] !== current[key]) {
      labels.push(label);
    }
  }
  return labels.length > 0 ? `Formatted: ${labels.join(", ")}` : "Formatted";
}

function describeParagraphFormatChange(previous: ParagraphFormatting | undefined, current: ParagraphFormatting | undefined): string {
  if (previous?.alignment !== current?.alignment) {
    return "Formatted: Paragraph alignment";
  }
  return "Formatted: Paragraph";
}

/** A cell property change's description (SPEC-039 AC-5/AC-15): a span
 * change is a merge/split, everything else a plain formatting edit. */
function describeCellPropertyChange(previous: CellPropertyChange["previous"], current: TableCell): string {
  if ((previous.colSpan ?? 1) !== (current.colSpan ?? 1) || (previous.rowSpan ?? 1) !== (current.rowSpan ?? 1)) {
    return (current.colSpan ?? 1) * (current.rowSpan ?? 1) > (previous.colSpan ?? 1) * (previous.rowSpan ?? 1)
      ? "Formatted: Cell merge"
      : "Formatted: Cell split";
  }
  if (previous.shadingFill !== current.shadingFill) {
    return "Formatted: Cell shading";
  }
  if (previous.paddingCm !== current.paddingCm) {
    return "Formatted: Cell padding";
  }
  if (previous.verticalAlign !== current.verticalAlign) {
    return "Formatted: Cell vertical alignment";
  }
  return "Formatted: Cell borders";
}

function describeRowPropertyChange(previous: RowPropertyChange["previous"], current: TableRow): string {
  return previous.minHeightCm !== current.minHeightCm ? "Formatted: Row height" : "Formatted: Row";
}

function describeTablePropertyChange(previous: TablePropertyChange["previous"], current: Table): string {
  if (previous.columnWidthsCm !== undefined) {
    return "Formatted: Column widths";
  }
  if (previous.alignment !== current.alignment) {
    return "Formatted: Table alignment";
  }
  if (previous.widthMode !== current.widthMode || previous.widthPercent !== current.widthPercent) {
    return "Formatted: Table width";
  }
  return "Formatted: Table";
}

interface ChangeAccumulator {
  changeId: string;
  kind: ChangeKind;
  author: string;
  date: string;
  description: string;
  previewParts: string[];
  at: CaretPosition;
  order: number;
}

type NoteFn = (
  changeId: string,
  kind: ChangeKind,
  info: TrackedChangeInfo,
  at: CaretPosition,
  previewPart: string,
  description: string,
) => void;

function walkParagraphForChanges(
  paragraph: Paragraph,
  sectionIndex: number,
  blockIndex: number,
  cell: CellPath | undefined,
  note: NoteFn,
): void {
  let offset = 0;
  for (const run of paragraph.runs) {
    const at: CaretPosition = { section: sectionIndex, block: blockIndex, offset, cell };
    const changes = run.changes;
    if (changes?.insert) {
      note(changes.insert.changeId, "insert", changes.insert, at, run.text, "Inserted");
    }
    if (changes?.delete) {
      note(changes.delete.changeId, "delete", changes.delete, at, run.text, "Deleted");
    }
    if (changes?.formatChange) {
      note(
        changes.formatChange.changeId,
        "format",
        changes.formatChange,
        at,
        run.text,
        describeRunFormatChange(changes.formatChange.previous, run.formatting),
      );
    }
    offset += run.text.length;
  }
  const paraChanges = paragraph.changes;
  const endAt: CaretPosition = { section: sectionIndex, block: blockIndex, offset, cell };
  if (paraChanges?.markInsert) {
    note(paraChanges.markInsert.changeId, "paragraphMark", paraChanges.markInsert, endAt, "", "Inserted paragraph break");
  }
  if (paraChanges?.markDelete) {
    note(paraChanges.markDelete.changeId, "paragraphMark", paraChanges.markDelete, endAt, "", "Deleted paragraph break");
  }
  if (paraChanges?.formatChange) {
    note(
      paraChanges.formatChange.changeId,
      "format",
      paraChanges.formatChange,
      { ...endAt, offset: 0 },
      "",
      describeParagraphFormatChange(paraChanges.formatChange.previous, paragraph.formatting),
    );
  }
}

function walkBlocksForChanges(blocks: Block[], sectionIndex: number, note: NoteFn): void {
  blocks.forEach((block, blockIndex) => {
    if (block.type === "paragraph") {
      walkParagraphForChanges(block, sectionIndex, blockIndex, undefined, note);
      return;
    }
    if (isImage(block)) {
      const at: CaretPosition = { section: sectionIndex, block: blockIndex, offset: 0 };
      if (block.changes?.insert) {
        note(block.changes.insert.changeId, "image", block.changes.insert, at, "[image]", "Inserted image");
      }
      if (block.changes?.delete) {
        note(block.changes.delete.changeId, "image", block.changes.delete, at, "[image]", "Deleted image");
      }
      return;
    }
    if (isTable(block)) {
      const tableAt: CaretPosition = { section: sectionIndex, block: blockIndex, offset: 0 };
      if (block.changes?.propertyChange) {
        note(
          block.changes.propertyChange.changeId,
          "tableProperty",
          block.changes.propertyChange,
          tableAt,
          "[table]",
          describeTablePropertyChange(block.changes.propertyChange.previous, block),
        );
      }
      block.rows.forEach((row, rowIndex) => {
        const at: CaretPosition = { section: sectionIndex, block: blockIndex, offset: 0, cell: { row: rowIndex, col: 0, block: 0 } };
        if (row.changes?.insert) {
          note(row.changes.insert.changeId, "row", row.changes.insert, at, "[table row]", "Inserted table row");
        }
        if (row.changes?.delete) {
          note(row.changes.delete.changeId, "row", row.changes.delete, at, "[table row]", "Deleted table row");
        }
        if (row.changes?.propertyChange) {
          note(
            row.changes.propertyChange.changeId,
            "tableProperty",
            row.changes.propertyChange,
            at,
            "[table row]",
            describeRowPropertyChange(row.changes.propertyChange.previous, row),
          );
        }
        row.cells.forEach((tableCell, colIndex) => {
          const cellAt: CaretPosition = { section: sectionIndex, block: blockIndex, offset: 0, cell: { row: rowIndex, col: colIndex, block: 0 } };
          if (tableCell.changes?.insert) {
            note(tableCell.changes.insert.changeId, "cell", tableCell.changes.insert, cellAt, "[table cell]", "Inserted table cell");
          }
          if (tableCell.changes?.delete) {
            note(tableCell.changes.delete.changeId, "cell", tableCell.changes.delete, cellAt, "[table cell]", "Deleted table cell");
          }
          if (tableCell.changes?.propertyChange) {
            note(
              tableCell.changes.propertyChange.changeId,
              "tableProperty",
              tableCell.changes.propertyChange,
              cellAt,
              "[table cell]",
              describeCellPropertyChange(tableCell.changes.propertyChange.previous, tableCell),
            );
          }
          tableCell.blocks.forEach((paragraph, pIndex) => {
            walkParagraphForChanges(paragraph, sectionIndex, blockIndex, { row: rowIndex, col: colIndex, block: pIndex }, note);
          });
        });
      });
    }
  });
}

/** Every tracked change in the document, in document order (SPEC-039 AC "A
 * review pane lists all changes in document order"). A change that spans
 * several runs (e.g. a whole typed word) is one entry, keyed by its shared
 * `changeId`. */
export function collectChanges(model: DocumentModel): ChangeEntry[] {
  const byId = new Map<string, ChangeAccumulator>();
  let order = 0;
  const note: NoteFn = (changeId, kind, info, at, previewPart, description) => {
    let acc = byId.get(changeId);
    if (!acc) {
      acc = { changeId, kind, author: info.author, date: info.date, description, previewParts: [], at, order: order++ };
      byId.set(changeId, acc);
    }
    if (previewPart.length > 0) {
      acc.previewParts.push(previewPart);
    }
  };
  model.sections.forEach((section, sectionIndex) => walkBlocksForChanges(section.blocks, sectionIndex, note));
  return [...byId.values()]
    .sort((a, b) => a.order - b.order)
    .map((acc) => ({
      changeId: acc.changeId,
      kind: acc.kind,
      author: acc.author,
      date: acc.date,
      description: acc.description,
      preview: acc.previewParts.join("").slice(0, 120),
      at: acc.at,
    }));
}

/** The change entry the given caret sits inside/at-or-after, and its
 * neighbours — Next/Previous change navigation (AC-8) reads this to move
 * the caret without needing the review pane open. */
export function nextChangeFrom(model: DocumentModel, caret: CaretPosition): ChangeEntry | null {
  return collectChanges(model).find((entry) => compareCaret(entry.at, caret) > 0) ?? null;
}

/** The change entry immediately before the given caret (AC-8's Previous). */
export function previousChangeFrom(model: DocumentModel, caret: CaretPosition): ChangeEntry | null {
  let found: ChangeEntry | null = null;
  for (const entry of collectChanges(model)) {
    if (compareCaret(entry.at, caret) < 0) {
      found = entry;
    }
  }
  return found;
}

// --- accept / reject commands -------------------------------------------

type Verdict = "accept" | "reject";

function cleanChanges<T extends object>(changes: T): T | undefined {
  return Object.values(changes).some((v) => v !== undefined) ? changes : undefined;
}

function runHasChange(changes: RunTrackedChanges | undefined, changeId: string): boolean {
  return (
    changes?.insert?.changeId === changeId ||
    changes?.delete?.changeId === changeId ||
    changes?.formatChange?.changeId === changeId
  );
}

function transformRun(run: Paragraph["runs"][number], changeId: string, verdict: Verdict): Paragraph["runs"][number] | null {
  const changes = run.changes;
  if (!changes) {
    return run;
  }
  if (changes.insert?.changeId === changeId) {
    if (verdict === "reject") {
      return null;
    }
    return { ...run, changes: cleanChanges({ ...changes, insert: undefined }) };
  }
  if (changes.delete?.changeId === changeId) {
    if (verdict === "accept") {
      return null;
    }
    return { ...run, changes: cleanChanges({ ...changes, delete: undefined }) };
  }
  if (changes.formatChange?.changeId === changeId) {
    const next = { ...run, changes: cleanChanges({ ...changes, formatChange: undefined }) };
    return verdict === "reject" ? { ...next, formatting: changes.formatChange.previous } : next;
  }
  return run;
}

function transformRunsForChange(runs: Paragraph["runs"], changeId: string, verdict: Verdict): Paragraph["runs"] {
  if (!runs.some((run) => runHasChange(run.changes, changeId))) {
    return runs;
  }
  const transformed: Paragraph["runs"] = [];
  for (const run of runs) {
    const result = transformRun(run, changeId, verdict);
    if (result) {
      transformed.push(result);
    }
  }
  return normalizeRuns(transformed);
}

function withParagraphChanges(paragraph: Paragraph, changes: ParagraphTrackedChanges): Paragraph {
  return { ...paragraph, changes: cleanChanges(changes) };
}

function applyParagraphFormatChangeVerdict(paragraph: Paragraph, verdict: Verdict): Paragraph {
  const formatChange = paragraph.changes!.formatChange!;
  const next = withParagraphChanges(paragraph, { ...paragraph.changes, formatChange: undefined });
  return verdict === "reject" ? { ...next, formatting: formatChange.previous } : next;
}

/** Transforms a homogeneous run of Paragraph blocks (a section's top-level
 * blocks restricted to paragraphs, or a table cell's blocks) for one
 * change's accept/reject verdict, including the structural split/merge
 * cases (a `markInsert`/`markDelete` verdict can join two paragraphs back
 * into one or leave a merge undone). */
function transformParagraphRun(paragraphs: Paragraph[], changeId: string, verdict: Verdict): Paragraph[] {
  const result: Paragraph[] = [];
  let i = 0;
  while (i < paragraphs.length) {
    let p = paragraphs[i];
    const runs = transformRunsForChange(p.runs, changeId, verdict);
    if (runs !== p.runs) {
      p = { ...p, runs };
    }
    if (p.changes?.formatChange?.changeId === changeId) {
      p = applyParagraphFormatChangeVerdict(p, verdict);
    }
    const markInsert = p.changes?.markInsert;
    const markDelete = p.changes?.markDelete;
    if (markInsert?.changeId === changeId || markDelete?.changeId === changeId) {
      const commit = markInsert?.changeId === changeId ? verdict === "accept" : verdict === "reject";
      if (commit) {
        result.push(withParagraphChanges(p, { ...p.changes, markInsert: undefined, markDelete: undefined }));
        i += 1;
      } else {
        const next = paragraphs[i + 1];
        if (next) {
          result.push({ ...next, runs: [...p.runs, ...next.runs] });
          i += 2;
        } else {
          result.push(p);
          i += 1;
        }
      }
      continue;
    }
    result.push(p);
    i += 1;
  }
  return result;
}

function applyRowPropertyChangeVerdict(row: TableRow, verdict: Verdict): TableRow {
  const propertyChange = row.changes!.propertyChange!;
  const next = { ...row, changes: cleanChanges({ ...row.changes, propertyChange: undefined }) };
  return verdict === "reject" ? { ...next, minHeightCm: propertyChange.previous.minHeightCm } : next;
}

function applyCellPropertyChangeVerdict(cell: TableCell, verdict: Verdict): TableCell {
  const propertyChange = cell.changes!.propertyChange!;
  const next = { ...cell, changes: cleanChanges({ ...cell.changes, propertyChange: undefined }) };
  if (verdict !== "reject") {
    return next;
  }
  return {
    ...next,
    colSpan: propertyChange.previous.colSpan,
    rowSpan: propertyChange.previous.rowSpan,
    borders: propertyChange.previous.borders,
    shadingFill: propertyChange.previous.shadingFill,
    paddingCm: propertyChange.previous.paddingCm,
    verticalAlign: propertyChange.previous.verticalAlign,
  };
}

function applyTablePropertyChangeVerdict(table: Table, verdict: Verdict): Table {
  const propertyChange = table.changes!.propertyChange!;
  const next = { ...table, changes: cleanChanges({ ...table.changes, propertyChange: undefined }) };
  if (verdict !== "reject") {
    return next;
  }
  return {
    ...next,
    columnWidthsCm: propertyChange.previous.columnWidthsCm ?? next.columnWidthsCm,
    alignment: propertyChange.previous.alignment,
    widthMode: propertyChange.previous.widthMode,
    widthPercent: propertyChange.previous.widthPercent,
  };
}

/** The grid column index every row's cell with this changeId shares (a
 * column insert/delete's cells are all marked with the one changeId — see
 * trackChangesDiff.ts's diffRowCells), or -1 if no cell carries it. */
function findCellChangeColumn(table: Table, changeId: string): number {
  for (const row of table.rows) {
    const index = row.cells.findIndex((cell) => cell.changes?.insert?.changeId === changeId || cell.changes?.delete?.changeId === changeId);
    if (index >= 0) {
      return index;
    }
  }
  return -1;
}

/** A tracked column insert/delete's verdict (SPEC-039 AC-15): every row
 * shares the one changeId across its cell at this column (see
 * findCellChangeColumn), so committing it drops the whole column — cells
 * and its `columnWidthsCm` entry — uniformly rather than per-row. */
function transformCellInsertDeleteColumn(table: Table, column: number, changeId: string, verdict: Verdict): Table {
  const isInsert = table.rows.some((row) => row.cells[column]?.changes?.insert?.changeId === changeId);
  const drop = isInsert ? verdict === "reject" : verdict === "accept";
  if (drop) {
    const rows = table.rows.map((row) => ({ ...row, cells: row.cells.filter((_, index) => index !== column) }));
    const columnWidthsCm = table.columnWidthsCm.filter((_, index) => index !== column);
    return { ...table, rows, columnWidthsCm };
  }
  const rows = table.rows.map((row) => {
    const cell = row.cells[column];
    if (!cell) {
      return row;
    }
    const cells = [...row.cells];
    cells[column] = { ...cell, changes: cleanChanges({ ...cell.changes, insert: undefined, delete: undefined }) };
    return { ...row, cells };
  });
  return { ...table, rows };
}

function transformTable(table: Table, changeId: string, verdict: Verdict): Table {
  if (table.changes?.propertyChange?.changeId === changeId) {
    return applyTablePropertyChangeVerdict(table, verdict);
  }
  const cellColumn = findCellChangeColumn(table, changeId);
  if (cellColumn >= 0) {
    return transformCellInsertDeleteColumn(table, cellColumn, changeId, verdict);
  }
  const rows: TableRow[] = [];
  for (const row of table.rows) {
    if (row.changes?.insert?.changeId === changeId) {
      if (verdict === "reject") {
        continue;
      }
      rows.push({ ...row, changes: cleanChanges({ ...row.changes, insert: undefined }) });
      continue;
    }
    if (row.changes?.delete?.changeId === changeId) {
      if (verdict === "accept") {
        continue;
      }
      rows.push({ ...row, changes: cleanChanges({ ...row.changes, delete: undefined }) });
      continue;
    }
    if (row.changes?.propertyChange?.changeId === changeId) {
      rows.push(applyRowPropertyChangeVerdict(row, verdict));
      continue;
    }
    const cells = row.cells.map((cell) => {
      if (cell.changes?.propertyChange?.changeId === changeId) {
        return applyCellPropertyChangeVerdict(cell, verdict);
      }
      return { ...cell, blocks: transformParagraphRun(cell.blocks, changeId, verdict) };
    });
    rows.push({ ...row, cells });
  }
  return { ...table, rows };
}

function transformImage(image: Image, changeId: string, verdict: Verdict): Image | null {
  if (image.changes?.insert?.changeId === changeId) {
    return verdict === "reject" ? null : { ...image, changes: cleanChanges({ ...image.changes, insert: undefined }) };
  }
  if (image.changes?.delete?.changeId === changeId) {
    return verdict === "accept" ? null : { ...image, changes: cleanChanges({ ...image.changes, delete: undefined }) };
  }
  return image;
}

function transformBlockArray(blocks: Block[], changeId: string, verdict: Verdict): Block[] {
  const result: Block[] = [];
  let i = 0;
  while (i < blocks.length) {
    const block = blocks[i];
    if (block.type === "paragraph") {
      const next = blocks[i + 1];
      if ((block.changes?.markInsert || block.changes?.markDelete) && next?.type === "paragraph") {
        // A markInsert/markDelete verdict can merge this paragraph into the
        // next one — run the pair through the shared paragraph-array
        // transform, which always consumes both source blocks it was given
        // (whether or not it actually merges them).
        result.push(...transformParagraphRun([block, next], changeId, verdict));
        i += 2;
        continue;
      }
      result.push(...transformParagraphRun([block], changeId, verdict));
      i += 1;
      continue;
    }
    if (isImage(block)) {
      const transformed = transformImage(block, changeId, verdict);
      if (transformed) {
        result.push(transformed);
      }
      i += 1;
      continue;
    }
    if (isTable(block)) {
      result.push(transformTable(block, changeId, verdict));
      i += 1;
      continue;
    }
    result.push(block);
    i += 1;
  }
  return result;
}

function transformSection(section: Section, changeId: string, verdict: Verdict): Section {
  return { ...section, blocks: transformBlockArray(section.blocks, changeId, verdict) };
}

/** Applies one change's accept/reject verdict across the whole document
 * (SPEC-039 AC-7's per-change Accept/Reject). Undoable like any other
 * command — the caller runs it through the normal applyEdit/undo-history
 * path (AC-10). */
export function applyChangeVerdict(model: DocumentModel, changeId: string, verdict: Verdict): DocumentModel {
  return { ...model, sections: model.sections.map((section) => transformSection(section, changeId, verdict)) };
}

export function acceptChange(model: DocumentModel, changeId: string): DocumentModel {
  return applyChangeVerdict(model, changeId, "accept");
}

export function rejectChange(model: DocumentModel, changeId: string): DocumentModel {
  return applyChangeVerdict(model, changeId, "reject");
}

/** Accepts every tracked change in the document (AC-7's "Accept all"). */
export function acceptAllChanges(model: DocumentModel): DocumentModel {
  let result = model;
  for (const entry of collectChanges(model)) {
    result = acceptChange(result, entry.changeId);
  }
  return result;
}

/** Rejects every tracked change in the document (AC-7's "Reject all"). */
export function rejectAllChanges(model: DocumentModel): DocumentModel {
  let result = model;
  for (const entry of collectChanges(model)) {
    result = rejectChange(result, entry.changeId);
  }
  return result;
}

// --- exclusion from word count / find / spell check (AC-17) -------------

/** A block's visible text with tracked deletions excluded (SPEC-039 AC
 * "Tracked deletions are excluded from word count, find/replace results
 * and spell check"). Mirrors `textEditing.ts`'s `blockText` exactly except
 * for this one difference, so offsets still line up for anything that
 * doesn't care about deleted runs. */
export function visibleBlockText(block: Block): string {
  if (block.type !== "paragraph") {
    return blockText(block);
  }
  let text = "";
  for (const run of block.runs) {
    if (run.changes?.delete) {
      continue;
    }
    text += run.text;
  }
  return text;
}

/** A paragraph's tracked-deletion spans, as half-open `[start, end)` ranges
 * into its concatenated (full, deletions included) run text — the offset
 * space `blockText`/`FindMatch`/spell-check tokens use. Find/replace and
 * spell check keep matching against the full text (offsets a replace's
 * `splitRuns` relies on must stay valid) and instead skip any match that
 * overlaps one of these ranges. */
export function deletedRanges(paragraph: Paragraph): Array<{ start: number; end: number }> {
  const ranges: Array<{ start: number; end: number }> = [];
  let offset = 0;
  for (const run of paragraph.runs) {
    const start = offset;
    offset += run.text.length;
    if (run.changes?.delete) {
      const last = ranges[ranges.length - 1];
      if (last && last.end === start) {
        last.end = offset;
      } else {
        ranges.push({ start, end: offset });
      }
    }
  }
  return ranges;
}

/** Whether `[start, end)` overlaps any tracked-deletion span of `paragraph`
 * (SPEC-039 AC-17). */
export function overlapsDeletion(paragraph: Paragraph, start: number, end: number): boolean {
  return deletedRanges(paragraph).some((range) => start < range.end && end > range.start);
}
