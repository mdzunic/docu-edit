// The supported ceiling (SPEC-047 "The supported ceiling"): beyond these
// limits a document opens read-only rather than the editor pretending it can
// hold arbitrarily large content. Counted after parse, before pagination —
// page count is not known until the (expensive) layout pass has already run,
// so these are the quantities available at the door.
import { isImage, isShape, isTable, type Block, type DocumentModel } from "./documentModel";

export interface CeilingLimit {
  /** The name used in the warning message, e.g. "images". */
  name: string;
  limit: number;
}

/** Order matters: the first limit exceeded is the one named (spec: "The
 * first limit exceeded is the one named"), so blocks (the broadest measure)
 * is checked first, then the more specific quantities. */
const LIMITS: readonly CeilingLimit[] = [
  { name: "blocks", limit: 25_000 },
  { name: "images", limit: 1_000 },
  { name: "tables", limit: 500 },
  { name: "footnotes and endnotes", limit: 1_000 },
];

export interface CeilingCounts {
  blocks: number;
  images: number;
  tables: number;
  notes: number;
}

/** Counts one block, recursing into a table's cells (SPEC-022: cells hold
 * only paragraphs, never nested tables/images) and a shape's text frame
 * (SPEC-051: frames may hold lists and tables) — a container counts as one
 * block itself, plus one per nested block. */
function countBlock(block: Block, counts: CeilingCounts): void {
  counts.blocks += 1;
  if (isImage(block)) {
    counts.images += 1;
    return;
  }
  if (isTable(block)) {
    counts.tables += 1;
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const paragraph of cell.blocks) {
          countBlock(paragraph, counts);
        }
      }
    }
    return;
  }
  if (isShape(block) && block.text) {
    for (const nested of block.text.blocks) {
      countBlock(nested, counts);
    }
  }
}

/** The document's raw complexity counts — every body block (recursing into
 * table cells), every note (footnote or endnote, each counted once
 * regardless of its own content), matching "The supported ceiling"'s
 * quantities. Headers/footers are not counted: they repeat per page rather
 * than growing with document length, so they are not what makes a document
 * too large to edit. */
export function ceilingCounts(model: DocumentModel): CeilingCounts {
  const counts: CeilingCounts = { blocks: 0, images: 0, tables: 0, notes: 0 };
  for (const section of model.sections) {
    for (const block of section.blocks) {
      countBlock(block, counts);
    }
  }
  // Note CONTENT counts toward the block/image/table quantities too — 1,000
  // notes of one line each and 1,000 notes of a hundred paragraphs each are
  // very different documents, and the block ceiling is the complexity
  // measure that tells them apart.
  for (const note of model.notes ?? []) {
    for (const block of note.blocks) {
      countBlock(block, counts);
    }
  }
  counts.notes = model.notes?.length ?? 0;
  return counts;
}

export interface CeilingResult {
  exceeded: boolean;
  /** The first limit exceeded and the value reached — set only when
   * `exceeded` is true. */
  name?: string;
  limit?: number;
  value?: number;
}

/** Typed "no limit exceeded" seed for refs in `.vue` scripts (espree bars
 * the annotation — same pattern as domCaret.noElement). */
export function withinCeiling(): CeilingResult {
  return { exceeded: false };
}

const VALUE_BY_NAME: Record<string, keyof CeilingCounts> = {
  blocks: "blocks",
  images: "images",
  tables: "tables",
  "footnotes and endnotes": "notes",
};

/** Checks a document's counts against "The supported ceiling", returning
 * the first limit exceeded (checked in the table's own order) or
 * `{exceeded: false}` when the document is within every limit. */
export function checkSupportedCeiling(model: DocumentModel): CeilingResult {
  const counts = ceilingCounts(model);
  for (const { name, limit } of LIMITS) {
    const value = counts[VALUE_BY_NAME[name]];
    if (value > limit) {
      return { exceeded: true, name, limit, value };
    }
  }
  return { exceeded: false };
}

/** The read-only warning text (spec example: "This document is too large to
 * edit (1,240 images; the limit is 1,000). It's open for reading only."),
 * naming the first limit exceeded and the value reached. */
export function ceilingWarningMessage(result: CeilingResult): string {
  if (!result.exceeded || result.name === undefined || result.value === undefined || result.limit === undefined) {
    return "";
  }
  const value = result.value.toLocaleString("en-US");
  const limit = result.limit.toLocaleString("en-US");
  return `This document is too large to edit (${value} ${result.name}; the limit is ${limit}). It's open for reading only.`;
}
